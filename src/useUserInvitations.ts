import { useCallback, useEffect, useRef, useState, type SetStateAction } from 'react'
import { collection, doc, onSnapshot, runTransaction, serverTimestamp, writeBatch } from 'firebase/firestore'
import { getDownloadURL, ref, uploadString } from 'firebase/storage'
import { firebaseDb, firebaseStorage } from './firebase'

type InvitationRecord = { id: number; coverImage?: string; shareToken?: string; shareMessage?: string; title?: string; type?: string; date?: string; dateInput?: string; place?: string; guests?: number; guestNames?: string[]; guestGroups?: { id: string; name: string; count: number; rsvp?: string; guestsComing?: number }[]; schedule?: unknown[]; tables?: { name: string; guests: string[] }[] }
type GuestResponse = { rsvp?: 'pending' | 'attending' | 'declined'; guestsComing?: number }

const firestoreSafe = (value: unknown): unknown => {
  if (Array.isArray(value)) return value.map(firestoreSafe)
  if (value && typeof value === 'object') {
    return Object.fromEntries(Object.entries(value).filter(([, entry]) => entry !== undefined).map(([key, entry]) => [key, firestoreSafe(entry)]))
  }
  return value
}

const localInvitationsKey = (ownerId: string) => `invitecyprus-user-${ownerId}-invitations`
const groupsForShare = (invite: InvitationRecord) => {
  if (invite.guestGroups) return invite.guestGroups
  const groups = (invite.guestNames ?? []).map((name, index) => ({ id: `legacy-${invite.id}-${index}`, name, count: 1 }))
  const remainder = Math.max(0, (invite.guests ?? 0) - groups.length)
  if (remainder > 0) groups.push({ id: `legacy-${invite.id}-additional`, name: 'Additional guests', count: remainder })
  return groups
}

const publicInvitationData = <T extends InvitationRecord>(invite: T, ownerId: string) => ({
  ownerUid: ownerId,
  invitationId: invite.id,
  title: invite.title ?? '',
  type: invite.type ?? 'Celebration',
  date: invite.date ?? '',
  dateInput: invite.dateInput ?? '',
  seatingVisibleAt: invite.dateInput ? new Date(`${invite.dateInput}T00:00:00`) : new Date('9999-12-31T00:00:00'),
  place: invite.place ?? '',
  coverImage: invite.coverImage?.startsWith('http') ? invite.coverImage : '',
  guestNames: invite.guestNames ?? [],
  guestGroups: groupsForShare(invite),
  guestGroupCounts: Object.fromEntries(groupsForShare(invite).map((group) => [group.id, group.count])),
  schedule: invite.schedule ?? [],
  shareMessage: invite.shareMessage ?? '',
})

const readLocalInvitations = <T,>(ownerId: string): T[] => {
  try {
    const saved = localStorage.getItem(localInvitationsKey(ownerId))
    return saved ? JSON.parse(saved) as T[] : []
  } catch {
    return []
  }
}

const compactImageData = async (dataUrl: string) => {
  if (dataUrl.length <= 550_000) return dataUrl
  const image = new Image()
  await new Promise<void>((resolve, reject) => {
    image.onload = () => resolve()
    image.onerror = () => reject(new Error('Could not migrate an invitation cover photo.'))
    image.src = dataUrl
  })
  const canvas = document.createElement('canvas')
  const scale = Math.min(1, 1000 / image.width, 800 / image.height)
  canvas.width = Math.max(1, Math.round(image.width * scale))
  canvas.height = Math.max(1, Math.round(image.height * scale))
  const context = canvas.getContext('2d')
  if (!context) throw new Error('Could not prepare an invitation cover photo for sync.')
  context.drawImage(image, 0, 0, canvas.width, canvas.height)
  for (const quality of [0.78, 0.62, 0.48, 0.34]) {
    const compact = canvas.toDataURL('image/jpeg', quality)
    if (compact.length <= 550_000) return compact
  }
  throw new Error('A cover photo is too large to sync. Remove it or choose a smaller image.')
}

const prepareInvitation = async <T extends InvitationRecord>(invite: T, ownerId: string): Promise<T> => {
  if (!invite.coverImage?.startsWith('data:')) return invite
  if (firebaseStorage) {
    const imageRef = ref(firebaseStorage, `users/${ownerId}/covers/${invite.id}.jpg`)
    await uploadString(imageRef, invite.coverImage, 'data_url', { contentType: 'image/jpeg' })
    return { ...invite, coverImage: await getDownloadURL(imageRef) }
  }
  return { ...invite, coverImage: await compactImageData(invite.coverImage) }
}

export function useUserInvitations<T extends InvitationRecord>(ownerId: string | null) {
  const [invites, setInvitesState] = useState<T[]>([])
  const [loadedOwner, setLoadedOwner] = useState<string | null>(null)
  const [error, setError] = useState('')
  const latestInvites = useRef<T[]>([])
  const responseCache = useRef(new Map<string, Map<string, GuestResponse>>())
  const activeOwner = useRef(ownerId)
  const writeQueue = useRef(Promise.resolve())
  const syncFailure = useRef<unknown>(null)
  activeOwner.current = ownerId

  useEffect(() => {
    let active = true
    let legacyMigrationStarted = false
    latestInvites.current = []
    responseCache.current.clear()
    setInvitesState([])
    setLoadedOwner(null)
    setError('')

    if (!ownerId) {
      setLoadedOwner(null)
      return () => { active = false }
    }

    if (ownerId === 'local-preview') {
      const local = readLocalInvitations<T>(ownerId)
      let invites = local
      if (!invites.length) {
        try { invites = JSON.parse(localStorage.getItem('invitecyprus-demo-invitations') ?? '[]') as T[] }
        catch { invites = [] }
      }
      latestInvites.current = invites
      setInvitesState(invites)
      setLoadedOwner(ownerId)
      return () => { active = false }
    }

    const database = firebaseDb
    if (!database) {
      setError('Cloud invitations are not configured. Add the Firebase project settings and deploy Firestore rules.')
      setLoadedOwner(ownerId)
      return () => { active = false }
    }

    const ownerCollection = collection(database, 'users', ownerId, 'invitations')
    const unsubscribe = onSnapshot(ownerCollection, (snapshot) => {
      if (snapshot.empty && snapshot.metadata.fromCache) return
      if (snapshot.empty && !legacyMigrationStarted) {
        legacyMigrationStarted = true
        const legacy = readLocalInvitations<T>(ownerId)
        if (legacy.length) {
          void (async () => {
            const prepared = await Promise.all(legacy.map((invite) => prepareInvitation(invite, ownerId)))
            await Promise.all(prepared.map((invite) => runTransaction(database, async (transaction) => {
              const invitationRef = doc(database, 'users', ownerId, 'invitations', String(invite.id))
              const existing = await transaction.get(invitationRef)
              if (!existing.exists()) transaction.set(invitationRef, firestoreSafe(invite) as object)
            })))
            localStorage.removeItem(localInvitationsKey(ownerId))
          })().catch((migrationError: unknown) => {
            if (!active) return
            setError(migrationError instanceof Error ? migrationError.message : 'Could not move your saved invitations to the cloud.')
            setLoadedOwner(ownerId)
          })
          return
        }
      }
      const next = snapshot.docs.map((invitation) => {
        const invite = invitation.data() as T
        const responses = invite.shareToken ? responseCache.current.get(invite.shareToken) : undefined
        if (!responses || !invite.guestGroups) return invite
        return { ...invite, guestGroups: invite.guestGroups.map((group) => ({ ...group, ...responses.get(group.id) })) }
      }).sort((left, right) => right.id - left.id)
      latestInvites.current = next
      setInvitesState(next)
      setLoadedOwner(ownerId)
      setError('')
    }, (snapshotError) => {
      if (!active) return
      setError(snapshotError.message || 'Could not load your cloud invitations.')
      setLoadedOwner(ownerId)
    })

    return () => {
      active = false
      unsubscribe()
    }
  }, [ownerId])

  const sharedSignature = invites.flatMap((invite) => invite.shareToken ? [`${invite.id}:${invite.shareToken}`] : []).sort().join('|')
  useEffect(() => {
    const database = firebaseDb
    if (!ownerId || ownerId === 'local-preview' || !database) return
    const tokens = invites.flatMap((invite) => invite.shareToken ? [invite.shareToken] : [])
    const unsubscribes = tokens.map((token) => onSnapshot(collection(database, 'publicInvitations', token, 'responses'), (snapshot) => {
      const responses = new Map<string, GuestResponse>()
      snapshot.docs.forEach((response) => responses.set(response.id, response.data() as GuestResponse))
      responseCache.current.set(token, responses)
      setInvitesState((current) => {
        const next = current.map((invite) => invite.shareToken === token && invite.guestGroups
          ? { ...invite, guestGroups: invite.guestGroups.map((group) => ({ ...group, ...responses.get(group.id) })) }
          : invite)
        latestInvites.current = next
        return next
      })
    }, (snapshotError) => {
      if (activeOwner.current === ownerId) setError(snapshotError.message || 'Could not load guest replies.')
    }))
    return () => unsubscribes.forEach((unsubscribe) => unsubscribe())
  }, [ownerId, sharedSignature])

  const setInvites = useCallback((update: SetStateAction<T[]>) => {
    const previous = latestInvites.current
    const next = typeof update === 'function' ? (update as (current: T[]) => T[])(previous) : update
    latestInvites.current = next
    setInvitesState(next)

    const ownerId = activeOwner.current
    if (!ownerId) return
    if (ownerId === 'local-preview') {
      try { localStorage.setItem(localInvitationsKey(ownerId), JSON.stringify(next)) }
      catch { setError('This browser could not save the preview invitations.') }
      return
    }
    const database = firebaseDb
    if (!database) return

    const previousById = new Map(previous.map((invite) => [invite.id, invite]))
    const nextById = new Map(next.map((invite) => [invite.id, invite]))
    const changed = next.filter((invite) => JSON.stringify(previousById.get(invite.id)) !== JSON.stringify(invite))
    const removed = previous.filter((invite) => !nextById.has(invite.id))
    if (!changed.length && !removed.length) return

    syncFailure.current = null
    writeQueue.current = writeQueue.current.then(async () => {
      const preparedChanges = await Promise.all(changed.map((invite) => prepareInvitation(invite, ownerId)))
      const operations = [
        ...preparedChanges.map((invite) => ({ type: 'set' as const, invite })),
        ...removed.map((invite) => ({ type: 'delete' as const, invite })),
      ]
      let batch = writeBatch(database)
      let writeCount = 0
      const queueWrite = async (write: (target: ReturnType<typeof writeBatch>) => void) => {
        if (writeCount >= 450) {
          await batch.commit()
          batch = writeBatch(database)
          writeCount = 0
        }
        write(batch)
        writeCount += 1
      }
      for (const operation of operations) {
          const invitationRef = doc(database, 'users', ownerId, 'invitations', String(operation.invite.id))
          if (operation.type === 'set') {
            const safeInvite = firestoreSafe(operation.invite) as object
            await queueWrite((target) => target.set(invitationRef, safeInvite))
            if (operation.invite.shareToken) {
              const token = operation.invite.shareToken
              await queueWrite((target) => target.set(doc(database, 'publicInvitations', token), publicInvitationData(operation.invite, ownerId)))
              const seatingRef = doc(database, 'publicInvitations', token, 'seating', 'plan')
              if (operation.invite.tables?.length) await queueWrite((target) => target.set(seatingRef, { tables: operation.invite.tables }))
              else await queueWrite((target) => target.delete(seatingRef))

              const beforeGroups = previousById.get(operation.invite.id)?.guestGroups ?? []
              const afterGroups = operation.invite.guestGroups ?? []
              for (const group of afterGroups) {
                const before = beforeGroups.find((item) => item.id === group.id)
                if (before && (before.rsvp !== group.rsvp || before.guestsComing !== group.guestsComing)) {
                  await queueWrite((target) => target.set(doc(database, 'publicInvitations', token, 'responses', group.id), {
                    rsvp: group.rsvp ?? 'pending',
                    guestsComing: group.rsvp === 'declined' ? 0 : Math.max(1, Number(group.guestsComing ?? group.count) || 1),
                    respondedAt: serverTimestamp(),
                  }))
                }
              }
            }
          } else {
            await queueWrite((target) => target.delete(invitationRef))
            if (operation.invite.shareToken) {
              await queueWrite((target) => target.delete(doc(database, 'publicInvitations', operation.invite.shareToken!)))
              await queueWrite((target) => target.delete(doc(database, 'publicInvitations', operation.invite.shareToken!, 'seating', 'plan')))
            }
          }
      }
      if (writeCount) await batch.commit()
      if (activeOwner.current === ownerId) setError('')
    }).catch((writeError: unknown) => {
      syncFailure.current = writeError
      if (activeOwner.current === ownerId) setError(writeError instanceof Error ? writeError.message : 'Could not sync your invitation changes.')
    })
  }, [])

  const waitForSync = useCallback(async () => {
    await writeQueue.current
    if (syncFailure.current) throw syncFailure.current
  }, [])

  return { invites, setInvites, waitForSync, loading: ownerId !== null && loadedOwner !== ownerId, error }
}
