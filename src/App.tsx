import { useEffect, useState, type FormEvent } from 'react'
import { flushSync } from 'react-dom'
import { createUserWithEmailAndPassword, FacebookAuthProvider, onAuthStateChanged, sendEmailVerification, sendPasswordResetEmail, signInWithEmailAndPassword, signInWithPopup, signOut, updateProfile, GoogleAuthProvider, type User } from 'firebase/auth'
import { ArrowLeft, ArrowRight, CalendarDays, Check, ChevronRight, CircleUserRound, Clock3, Crop, CreditCard, Download, ExternalLink, Eye, ImagePlus, LockKeyhole, LogOut, MapPin, Plus, Save, Shuffle, Trash2, Users, X } from 'lucide-react'
import { firebaseAuth, firebaseConfigured } from './firebase'

type ScheduleItem = { id: number; title: string; time: string; place: string; mapsUrl?: string }
type GuestGroup = { id: string; name: string; count: number; rsvp?: 'pending' | 'attending' | 'declined'; invitationSent?: boolean; sentVia?: string }
type Invite = { id: number; title: string; type: string; date: string; dateInput?: string; timeInput?: string; place: string; mapsUrl?: string; coverImage?: string; coverImagePosition?: { x: number; y: number; zoom: number }; guests: number; guestNames: string[]; guestGroups?: GuestGroup[]; tables: { name: string; guests: string[] }[]; schedule: ScheduleItem[] }
type Screen = 'login' | 'home' | 'create' | 'manage' | 'invitation-dashboard' | 'guest-management' | 'seating-management'
type SavedUi = { screen?: Screen; step?: number; kind?: string; title?: string; date?: string; time?: string; place?: string; mapsUrl?: string; guestGroups?: GuestGroup[]; guestText?: string; guestCount?: string; scheduleItems?: ScheduleItem[]; editingId?: number | null; manageId?: number | null; manageMode?: 'guests' | 'seating' | null; newGuest?: string; newGuestCount?: string; newTable?: string }
const readSavedUi = (key = 'invitecyprus-demo-ui'): SavedUi => {
  try { return JSON.parse(localStorage.getItem(key) ?? '{}') as SavedUi } catch { return {} }
}
const makeGuestGroup = (name = '', count = 1): GuestGroup => ({ id: `${Date.now()}-${Math.random().toString(36).slice(2)}`, name, count })
const guestGroupsForInvite = (invite: Invite): GuestGroup[] => {
  if (invite.guestGroups) return invite.guestGroups
  const groups = invite.guestNames.map((name, index) => ({ id: `legacy-${invite.id}-${index}`, name, count: 1 }))
  const remaining = invite.guests - groups.length
  if (remaining > 0) groups.push({ id: `legacy-${invite.id}-additional`, name: 'Additional guests', count: remaining })
  return groups
}
const peopleSeatedAtTable = (invite: Invite, table: { name: string; guests: string[] }) => table.guests.reduce((total, name) => total + (guestGroupsForInvite(invite).find((group) => group.name === name)?.count ?? 1), 0)
const isInviteCompleted = (invite: Invite) => {
  if (!invite.dateInput) return false
  const [year, month, day] = invite.dateInput.split('-').map(Number)
  if (!year || !month || !day) return false
  const eventDate = new Date(year, month - 1, day)
  const today = new Date()
  today.setHours(0, 0, 0, 0)
  return eventDate < today
}
const occasions = ['Wedding', 'Birthday', 'Baptism', 'Company event', 'Dinner party', 'Something else']
const invitationChannels = ['Facebook Messenger', 'SMS', 'Instagram', 'Viber', 'WhatsApp', 'In person', 'Email', 'Phone call', 'Other']
const scheduleSuggestions: Record<string, string[]> = {
  Wedding: ["Groom’s changing", "Bride’s changing", 'The ceremony', 'The celebration'],
  Birthday: ['Guest arrival', 'Games & activities', 'Cake cutting', 'The party'],
  Baptism: ['Church ceremony', 'Family photos', 'Reception'],
  'Company event': ['Welcome & networking', 'Presentations', 'Dinner'],
  'Dinner party': ['Guest arrival', 'Dinner is served', 'Dessert'],
  'Something else': ['Doors open', 'Main event', 'Celebration'],
}
const randomCoverPhotos = [
  'https://images.unsplash.com/photo-1519741497674-611481863552?auto=format&fit=crop&w=900&q=82',
  'https://images.unsplash.com/photo-1530103862676-de8c9debad1d?auto=format&fit=crop&w=900&q=82',
  'https://images.unsplash.com/photo-1511795409834-ef04bbd61622?auto=format&fit=crop&w=900&q=82',
]
const coverTransform = (position?: { x: number; y: number; zoom: number }) => {
  const x = position?.x ?? 50
  const y = position?.y ?? 50
  const zoom = position?.zoom ?? 100
  const panLimit = (zoom - 100) / 2
  return `translate(${((50 - x) / 50) * panLimit}%, ${((50 - y) / 50) * panLimit}%) scale(${zoom / 100})`
}

const xmlEscape = (value: string) => value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;')
const makeXlsx = (rows: string[][]) => {
  const encoder = new TextEncoder()
  const files = [
    ['[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'],
    ['_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'],
    ['xl/workbook.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Guest list" sheetId="1" r:id="rId1"/></sheets></workbook>'],
    ['xl/_rels/workbook.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'],
    ['xl/worksheets/sheet1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${rows.map((row, rowIndex) => `<row r="${rowIndex + 1}">${row.map((cell, colIndex) => `<c r="${String.fromCharCode(65 + colIndex)}${rowIndex + 1}" t="inlineStr"><is><t xml:space="preserve">${xmlEscape(cell)}</t></is></c>`).join('')}</row>`).join('')}</sheetData></worksheet>`],
  ] as const
  const crc32 = (bytes: Uint8Array) => {
    let crc = -1
    for (const byte of bytes) { crc ^= byte; for (let bit = 0; bit < 8; bit++) crc = (crc >>> 1) ^ (crc & 1 ? 0xedb88320 : 0) }
    return (crc ^ -1) >>> 0
  }
  const u16 = (value: number) => new Uint8Array([value & 255, value >>> 8 & 255])
  const u32 = (value: number) => new Uint8Array([value & 255, value >>> 8 & 255, value >>> 16 & 255, value >>> 24 & 255])
  const join = (parts: Uint8Array[]) => { const result = new Uint8Array(parts.reduce((sum, part) => sum + part.length, 0)); let offset = 0; for (const part of parts) { result.set(part, offset); offset += part.length } return result }
  const local: Uint8Array[] = [], central: Uint8Array[] = []
  let offset = 0
  for (const [name, contents] of files) {
    const filename = encoder.encode(name), data = encoder.encode(contents), crc = crc32(data)
    const header = join([u32(0x04034b50), u16(20), u16(0), u16(0), u16(0), u16(0), u32(crc), u32(data.length), u32(data.length), u16(filename.length), u16(0), filename, data])
    local.push(header)
    central.push(join([u32(0x02014b50), u16(20), u16(20), u16(0), u16(0), u16(0), u16(0), u32(crc), u32(data.length), u32(data.length), u16(filename.length), u16(0), u16(0), u16(0), u16(0), u32(0), u32(offset), filename]))
    offset += header.length
  }
  const centralBytes = join(central)
  return new Blob([join([...local, centralBytes, u32(0x06054b50), u16(0), u16(0), u16(files.length), u16(files.length), u32(centralBytes.length), u32(offset), u16(0)])], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
}

const isGoogleMapsUrl = (value: string) => {
  try { const url = new URL(value); return url.protocol === 'https:' && (url.hostname === 'maps.app.goo.gl' || url.hostname.endsWith('google.com') && url.pathname.includes('/maps') || url.hostname === 'goo.gl' && url.pathname.startsWith('/maps')) } catch { return false }
}

const googleMapsHref = (place: string, mapsUrl = '') => isGoogleMapsUrl(mapsUrl) ? mapsUrl : isGoogleMapsUrl(place) ? place : `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(place || mapsUrl)}`

const normalizeTime = (value: string) => {
  const trimmed = value.trim()
  if (/^(?:[01]\d|2[0-3]):[0-5]\d$/.test(trimmed)) return trimmed
  const digits = trimmed.replace(/\D/g, '')
  if (digits.length === 3) return `0${digits[0]}:${digits.slice(1)}`
  if (digits.length === 4) {
    const hours = Number(digits.slice(0, 2))
    const minutes = Number(digits.slice(2))
    if (hours < 24 && minutes < 60) return `${digits.slice(0, 2)}:${digits.slice(2)}`
  }
  return ''
}

const sortScheduleItemsChronologically = (items: ScheduleItem[]) => items
  .map((item, index) => ({ item, index, time: normalizeTime(item.time) }))
  .sort((a, b) => {
    if (!a.time && !b.time) return a.index - b.index
    if (!a.time) return 1
    if (!b.time) return -1
    return a.time.localeCompare(b.time) || a.index - b.index
  })
  .map(({ item }) => item)

const compressCoverImage = (file: File) => new Promise<string>((resolve, reject) => {
  const reader = new FileReader()
  reader.onerror = () => reject(new Error('Could not read that image.'))
  reader.onload = () => {
    const image = new Image()
    image.onerror = () => reject(new Error('That image could not be opened.'))
    image.onload = () => {
      const scale = Math.min(1, 1400 / image.width, 1100 / image.height)
      const canvas = document.createElement('canvas')
      canvas.width = Math.max(1, Math.round(image.width * scale))
      canvas.height = Math.max(1, Math.round(image.height * scale))
      const context = canvas.getContext('2d')
      if (!context) { reject(new Error('Could not prepare that image.')); return }
      context.fillStyle = '#fff'
      context.fillRect(0, 0, canvas.width, canvas.height)
      context.drawImage(image, 0, 0, canvas.width, canvas.height)
      resolve(canvas.toDataURL('image/jpeg', 0.8))
    }
    image.src = String(reader.result)
  }
  reader.readAsDataURL(file)
})

function App() {
  const savedUi = readSavedUi()
  const [access, setAccess] = useState<'checking' | 'locked' | 'open' | 'setup'>('checking')
  const [accessPassword, setAccessPassword] = useState('')
  const [accessError, setAccessError] = useState('')
  const [screen, setScreen] = useState<Screen>('login')
  const [authUser, setAuthUser] = useState<User | null>(null)
  const [localPreviewMode, setLocalPreviewMode] = useState(false)
  const [authReady, setAuthReady] = useState(!firebaseConfigured)
  const [authMode, setAuthMode] = useState<'signIn' | 'signUp'>('signIn')
  const [authName, setAuthName] = useState('')
  const [authEmail, setAuthEmail] = useState('')
  const [authPassword, setAuthPassword] = useState('')
  const [authPasswordConfirm, setAuthPasswordConfirm] = useState('')
  const [authBusy, setAuthBusy] = useState(false)
  const [authError, setAuthError] = useState('')
  const [authMessage, setAuthMessage] = useState('')
  const [inviteOwner, setInviteOwner] = useState<string | null>(null)
  const [profileMenuOpen, setProfileMenuOpen] = useState(false)
  const [profileName, setProfileName] = useState('')
  const [profileEmail, setProfileEmail] = useState('')
  const [profileDraftName, setProfileDraftName] = useState('')
  const [profileDraftEmail, setProfileDraftEmail] = useState('')
  const [profileSaved, setProfileSaved] = useState(false)
  const transitionTo = (nextScreen: Screen) => {
    if (nextScreen === screen) return
    setProfileMenuOpen(false)
    if (!document.startViewTransition || window.matchMedia('(prefers-reduced-motion: reduce)').matches) {
      setScreen(nextScreen)
      return
    }
    document.startViewTransition(() => flushSync(() => setScreen(nextScreen)))
  }
  const saveProfileSettings = async () => {
    const name = profileDraftName.trim()
    if (!name) {
      setFieldErrors((errors) => ({ ...errors, profileDisplayName: 'Enter your name.' }))
      return
    }
    if (!authUser && localPreviewMode) {
      setProfileName(name)
      setProfileSaved(true)
      setNotice('Profile settings saved for this local preview.')
      window.setTimeout(() => setNotice(''), 4000)
      return
    }
    if (!authUser) return
    try {
      await updateProfile(authUser, { displayName: name })
      setProfileName(name)
      setProfileSaved(true)
      setNotice('Profile settings saved.')
      window.setTimeout(() => setNotice(''), 4000)
    } catch {
      setNotice('Could not save your profile. Please try again.')
    }
  }
  const [invites, setInvites] = useState<Invite[]>([])
  const [step, setStep] = useState(savedUi.step ?? 1)
  const [kind, setKind] = useState(savedUi.kind ?? 'Wedding')
  const [title, setTitle] = useState(savedUi.title ?? '')
  const [date, setDate] = useState(savedUi.date ?? '')
  const [time, setTime] = useState(savedUi.time ?? '')
  const [place, setPlace] = useState(savedUi.place ?? '')
  const [mapsUrl, setMapsUrl] = useState(savedUi.mapsUrl ?? '')
  const [fieldErrors, setFieldErrors] = useState<Record<string, string>>({})
  const [guestGroups, setGuestGroups] = useState<GuestGroup[]>(() => {
    if (savedUi.guestGroups) return savedUi.guestGroups
    const groups = (savedUi.guestText ?? '').split(/[\n,;]/).map((name) => name.trim()).filter(Boolean).map((name) => makeGuestGroup(name))
    const oldTotal = Number.parseInt(savedUi.guestCount ?? '', 10) || 0
    const namedTotal = groups.length
    if (oldTotal > namedTotal) groups.push(makeGuestGroup('Additional guests', oldTotal - namedTotal))
    return groups
  })
  const [scheduleItems, setScheduleItems] = useState<ScheduleItem[]>(savedUi.scheduleItems ?? [])
  const [editingId, setEditingId] = useState<number | null>(savedUi.editingId ?? null)
  const [manageId, setManageId] = useState<number | null>(savedUi.manageId ?? null)
  const [manageMode, setManageMode] = useState<'guests' | 'seating' | null>(savedUi.manageMode ?? null)
  const [newGuest, setNewGuest] = useState(savedUi.newGuest ?? '')
  const [newGuestCount, setNewGuestCount] = useState(savedUi.newGuestCount ?? '1')
  const [newTable, setNewTable] = useState(savedUi.newTable ?? '')
  const [notice, setNotice] = useState('')
  const [coverEditorId, setCoverEditorId] = useState<number | null>(null)
  const [deleteConfirmId, setDeleteConfirmId] = useState<number | null>(null)
  const [previewInviteId, setPreviewInviteId] = useState<number | null>(null)

  const clearFieldError = (key: string) => setFieldErrors((errors) => {
    if (!errors[key]) return errors
    const next = { ...errors }
    delete next[key]
    return next
  })

  const validateInvitationDetails = () => {
    const errors: Record<string, string> = {}
    if (!title.trim()) errors.invitationTitle = 'Enter an invitation name.'
    if (!date) errors.eventDate = 'Choose the event date.'
    if (scheduleItems.length === 0) errors.scheduleSelection = 'Choose at least one event-day moment.'
    for (const item of scheduleItems) {
      if (!item.title.trim()) errors[`schedule-title-${item.id}`] = 'Enter a name for this moment.'
      if (!normalizeTime(item.time)) errors[`schedule-time-${item.id}`] = 'Enter a valid time (HH:MM).'
    }
    setFieldErrors(errors)
    return Object.keys(errors).length === 0
  }

  const validateInvitees = () => {
    const errors: Record<string, string> = {}
    if (guestGroups.length === 0) errors.guestRows = 'Add at least one invitee.'
    for (const group of guestGroups) {
      if (!group.name.trim()) errors[`invitee-${group.id}`] = 'Enter a person or group name.'
    }
    setFieldErrors(errors)
    return Object.keys(errors).length === 0
  }

  useEffect(() => {
    fetch('/__invitecyprus/access', { credentials: 'same-origin' })
      .then(async (response) => ({ response, data: await response.json() as { configured?: boolean; authorized?: boolean } }))
      .then(({ response, data }) => setAccess(!data.configured ? 'setup' : response.ok && data.authorized ? 'open' : 'locked'))
      .catch(() => setAccess('setup'))
  }, [])

  useEffect(() => {
    if (!firebaseAuth) return
    return onAuthStateChanged(firebaseAuth, (user) => {
      const passwordAccountNeedsVerification = user?.providerData.some((provider) => provider.providerId === 'password') && !user.emailVerified
      if (passwordAccountNeedsVerification) {
        setAuthUser(null)
        setAuthReady(true)
        setScreen('login')
        return
      }
      setAuthUser(user)
      setLocalPreviewMode(false)
      setAuthReady(true)
      setProfileName(user?.displayName || user?.email?.split('@')[0] || '')
      setProfileDraftName(user?.displayName || user?.email?.split('@')[0] || '')
      setProfileEmail(user?.email ?? '')
      setProfileDraftEmail(user?.email ?? '')
      setProfileMenuOpen(false)
      setAuthPassword('')
      if (!user) {
        setScreen('login')
        setInvites([])
        setInviteOwner(null)
        return
      }
      const uiKey = `invitecyprus-user-${user.uid}-ui`
      const userUi = readSavedUi(uiKey)
      try {
        const savedInvites = localStorage.getItem(`invitecyprus-user-${user.uid}-invitations`)
        setInvites(savedInvites ? JSON.parse(savedInvites) as Invite[] : [])
      } catch { setInvites([]) }
      setInviteOwner(user.uid)
      setStep(userUi.step ?? 1)
      setKind(userUi.kind ?? 'Wedding')
      setTitle(userUi.title ?? '')
      setDate(userUi.date ?? '')
      setTime(userUi.time ?? '')
      setPlace(userUi.place ?? '')
      setMapsUrl(userUi.mapsUrl ?? '')
      setGuestGroups(userUi.guestGroups ?? [makeGuestGroup()])
      setScheduleItems(userUi.scheduleItems ?? [])
      setEditingId(userUi.editingId ?? null)
      setManageId(userUi.manageId ?? null)
      setManageMode(userUi.manageMode ?? null)
      setNewGuest(userUi.newGuest ?? '')
      setNewGuestCount(userUi.newGuestCount ?? '1')
      setNewTable(userUi.newTable ?? '')
      setScreen(userUi.screen && userUi.screen !== 'login' ? userUi.screen : 'home')
    })
  }, [])

  const unlockPreview = async (event: FormEvent) => {
    event.preventDefault()
    setAccessError('')
    try {
      const response = await fetch('/__invitecyprus/access', { method: 'POST', headers: { 'Content-Type': 'application/json' }, credentials: 'same-origin', body: JSON.stringify({ password: accessPassword }) })
      if (!response.ok) { setAccessError(response.status === 401 ? 'That password doesn’t match. Try again.' : 'The password gate isn’t configured yet.'); return }
      setAccess('open')
      setAccessPassword('')
    } catch { setAccessError('Could not check the password. Make sure the app is running locally.') }
  }

  useEffect(() => {
    const ownerId = authUser?.uid ?? (localPreviewMode ? 'local-preview' : null)
    if (!ownerId) return
    try {
      localStorage.setItem(ownerId === 'local-preview' ? 'invitecyprus-user-local-preview-ui' : `invitecyprus-user-${ownerId}-ui`, JSON.stringify({ screen, step, kind, title, date, time, place, mapsUrl, guestGroups, scheduleItems, editingId, manageId, manageMode, newGuest, newGuestCount, newTable } satisfies SavedUi))
    } catch { /* Storage may be disabled; keep the app usable for this session. */ }
  }, [authUser, localPreviewMode, screen, step, kind, title, date, time, place, mapsUrl, guestGroups, scheduleItems, editingId, manageId, manageMode, newGuest, newGuestCount, newTable])

  useEffect(() => {
    const ownerId = authUser?.uid ?? (localPreviewMode ? 'local-preview' : null)
    if (!ownerId || inviteOwner !== ownerId) return
    try { localStorage.setItem(`invitecyprus-user-${ownerId}-invitations`, JSON.stringify(invites)) } catch { /* Storage may be full or disabled. */ }
  }, [authUser, localPreviewMode, inviteOwner, invites])

  const authErrorMessage = (error: unknown) => {
    const code = typeof error === 'object' && error !== null && 'code' in error ? String(error.code) : ''
    if (code === 'auth/invalid-credential' || code === 'auth/wrong-password' || code === 'auth/user-not-found') return 'That email and password combination was not found.'
    if (code === 'auth/email-already-in-use') return 'An account already exists for this email. Try logging in instead.'
    if (code === 'auth/weak-password') return 'Choose a password with at least 8 characters.'
    if (code === 'auth/invalid-email') return 'Enter a valid email address.'
    if (code === 'auth/popup-closed-by-user') return 'The sign-in window was closed before finishing.'
    if (code === 'auth/account-exists-with-different-credential') return 'An account with this email already uses a different sign-in method. Log in with that method first.'
    if (code === 'auth/operation-not-allowed') return 'This sign-in method is not enabled in the Firebase project yet.'
    if (code === 'auth/unauthorized-domain') return 'This website address is not authorized in Firebase Authentication settings.'
    return 'We could not complete that request. Check your connection and try again.'
  }

  const submitAuthForm = async (event: FormEvent) => {
    event.preventDefault()
    setAuthError('')
    setAuthMessage('')
    if (!firebaseAuth) { setAuthError('Add the Firebase settings to enable accounts.'); return }
    if (authMode === 'signUp') {
      if (!authName.trim()) { setAuthError('Enter your name to create an account.'); return }
      if (authPassword.length < 8) { setAuthError('Choose a password with at least 8 characters.'); return }
      if (authPassword !== authPasswordConfirm) { setAuthError('Those passwords do not match.'); return }
    }
    setAuthBusy(true)
    try {
      if (authMode === 'signUp') {
        const credential = await createUserWithEmailAndPassword(firebaseAuth, authEmail.trim(), authPassword)
        await updateProfile(credential.user, { displayName: authName.trim() })
        await sendEmailVerification(credential.user)
        await signOut(firebaseAuth)
        setAuthMode('signIn')
        setAuthPassword('')
        setAuthPasswordConfirm('')
        setAuthMessage('Your account is ready. Check your email to verify your address, then log in.')
      } else {
        const credential = await signInWithEmailAndPassword(firebaseAuth, authEmail.trim(), authPassword)
        if (!credential.user.emailVerified) {
          await sendEmailVerification(credential.user)
          await signOut(firebaseAuth)
          setAuthError('Please verify your email before logging in. We sent you another verification link.')
        }
      }
    } catch (error) { setAuthError(authErrorMessage(error)) }
    finally { setAuthBusy(false) }
  }

  const signInWithProvider = async (provider: 'google' | 'facebook') => {
    setAuthError('')
    setAuthMessage('')
    if (!firebaseAuth) { setAuthError('Add the Firebase settings to enable accounts.'); return }
    setAuthBusy(true)
    try {
      await signInWithPopup(firebaseAuth, provider === 'google' ? new GoogleAuthProvider() : new FacebookAuthProvider())
    } catch (error) { setAuthError(authErrorMessage(error)) }
    finally { setAuthBusy(false) }
  }

  const requestPasswordReset = async () => {
    setAuthError('')
    setAuthMessage('')
    if (!firebaseAuth) { setAuthError('Add the Firebase settings to enable password reset.'); return }
    if (!authEmail.trim()) { setAuthError('Enter your email address first, then choose “Forgot password?”.'); return }
    setAuthBusy(true)
    try {
      await sendPasswordResetEmail(firebaseAuth, authEmail.trim())
      setAuthMessage('If an account uses that email, a password reset link is on its way.')
    } catch (error) { setAuthError(authErrorMessage(error)) }
    finally { setAuthBusy(false) }
  }

  const logOut = async () => {
    if (!firebaseAuth || localPreviewMode) {
      setLocalPreviewMode(false)
      setInviteOwner(null)
      setInvites([])
      setScreen('login')
      return
    }
    try { await signOut(firebaseAuth) } catch { setNotice('Could not log out. Please try again.') }
  }

  const enterLocalPreview = () => {
    const ownerId = 'local-preview'
    const userUi = readSavedUi(`invitecyprus-user-${ownerId}-ui`)
    setLocalPreviewMode(true)
    setProfileName('Preview host')
    setProfileDraftName('Preview host')
    setProfileEmail('local-preview@invitecyprus.test')
    setProfileDraftEmail('local-preview@invitecyprus.test')
    setProfileMenuOpen(false)
    setInvites(() => {
      try {
        const saved = localStorage.getItem(`invitecyprus-user-${ownerId}-invitations`) ?? localStorage.getItem('invitecyprus-demo-invitations')
        return saved ? JSON.parse(saved) as Invite[] : []
      } catch { return [] }
    })
    setInviteOwner(ownerId)
    setStep(userUi.step ?? savedUi.step ?? 1)
    setKind(userUi.kind ?? savedUi.kind ?? 'Wedding')
    setTitle(userUi.title ?? savedUi.title ?? '')
    setDate(userUi.date ?? savedUi.date ?? '')
    setTime(userUi.time ?? savedUi.time ?? '')
    setPlace(userUi.place ?? savedUi.place ?? '')
    setMapsUrl(userUi.mapsUrl ?? savedUi.mapsUrl ?? '')
    setGuestGroups(userUi.guestGroups ?? savedUi.guestGroups ?? [makeGuestGroup()])
    setScheduleItems(userUi.scheduleItems ?? savedUi.scheduleItems ?? [])
    setEditingId(userUi.editingId ?? savedUi.editingId ?? null)
    setManageId(userUi.manageId ?? savedUi.manageId ?? null)
    setManageMode(userUi.manageMode ?? savedUi.manageMode ?? null)
    setNewGuest(userUi.newGuest ?? savedUi.newGuest ?? '')
    setNewGuestCount(userUi.newGuestCount ?? savedUi.newGuestCount ?? '1')
    setNewTable(userUi.newTable ?? savedUi.newTable ?? '')
    setScreen(userUi.screen && userUi.screen !== 'login' ? userUi.screen : 'home')
  }

  const create = () => {
    if (!editingId && !validateInvitees()) return
    if (!validateInvitationDetails()) {
      setStep(2)
      return
    }
    const normalizedGroups = guestGroups.filter((group) => group.name.trim()).map((group) => ({ ...group, name: group.name.trim(), count: Math.max(1, Number.parseInt(String(group.count), 10) || 1) }))
    const names = normalizedGroups.map((group) => group.name)
    const formattedDate = date ? new Date(`${date}T12:00:00`).toLocaleDateString('en-GB', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' }) : 'Date to be decided'
    const totalGuests = normalizedGroups.reduce((total, group) => total + group.count, 0)
    const normalizedSchedule = scheduleItems.map((item) => ({ ...item, time: normalizeTime(item.time) }))
    const scheduleLocation = normalizedSchedule.find((item) => item.place.trim())?.place.trim()
    setInvites((all) => editingId ? all.map((invite) => invite.id === editingId ? { ...invite, title: title || `${kind} invitation`, type: kind, date: date ? formattedDate : invite.date, dateInput: date || invite.dateInput, timeInput: '', place: place || invite.place, mapsUrl, guests: totalGuests, guestNames: names, guestGroups: normalizedGroups, tables: invite.tables.map((table) => ({ ...table, guests: table.guests.filter((guest) => names.includes(guest)) })), schedule: normalizedSchedule } : invite) : [{ id: Date.now(), title: title || `${kind} invitation`, type: kind, date: formattedDate, dateInput: date, timeInput: '', place: scheduleLocation || 'See event-day schedule for location details', mapsUrl: '', guests: totalGuests, guestNames: names, guestGroups: normalizedGroups, tables: [], schedule: normalizedSchedule }, ...all])
    transitionTo(editingId ? 'invitation-dashboard' : 'manage')
    setStep(1); setKind('Wedding'); setTitle(''); setDate(''); setTime(''); setPlace(''); setMapsUrl(''); setGuestGroups([makeGuestGroup()]); setScheduleItems([]); setEditingId(null)
    setNotice('Invitation saved. You can keep editing it whenever you’re ready.')
    window.setTimeout(() => setNotice(''), 5000)
  }

  const openInvitationDashboard = (inviteId: number) => {
    setManageId(inviteId)
    setManageMode(null)
    transitionTo('invitation-dashboard')
  }

  const openManageMode = (inviteId: number, mode: 'guests' | 'seating') => {
    setManageId(inviteId)
    setManageMode(null)
    transitionTo(mode === 'guests' ? 'guest-management' : 'seating-management')
  }

  const back = () => {
    if (editingId !== null) { transitionTo('invitation-dashboard'); return }
    if (screen === 'create' && step > 1) setStep(step - 1)
    else transitionTo('home')
  }

  const editInvite = (invite: Invite) => {
    setEditingId(invite.id); setKind(invite.type); setTitle(invite.title); setPlace(invite.place === 'Place to be decided' ? '' : invite.place); setMapsUrl(invite.mapsUrl ?? ''); setGuestGroups(guestGroupsForInvite(invite)); setScheduleItems(sortScheduleItemsChronologically(invite.schedule ?? [])); setDate(invite.dateInput ?? ''); setTime(''); setStep(2); transitionTo('create')
  }

  const addScheduleItem = (itemTitle: string) => {
    if (!itemTitle.trim() || scheduleItems.some((item) => item.title.toLowerCase() === itemTitle.toLowerCase())) return
    setScheduleItems((items) => [...items, { id: Date.now() + Math.random(), title: itemTitle, time: '', place: '' }])
    clearFieldError('scheduleSelection')
  }

  const removeScheduleItem = (id: number) => {
    const remaining = scheduleItems.filter((item) => item.id !== id)
    setScheduleItems(remaining)
    setFieldErrors((errors) => {
      const next = { ...errors }
      delete next[`schedule-title-${id}`]
      delete next[`schedule-time-${id}`]
      if (remaining.length === 0 && errors.scheduleSelection) next.scheduleSelection = 'Choose at least one event-day moment.'
      return next
    })
  }

  const updateScheduleItem = (id: number, update: Partial<ScheduleItem>) => {
    setScheduleItems((items) => items.map((item) => item.id === id ? { ...item, ...update } : item))
    if (update.title !== undefined) clearFieldError(`schedule-title-${id}`)
    if (update.time !== undefined) clearFieldError(`schedule-time-${id}`)
  }
  const commitScheduleTime = (id: number, value: string) => setScheduleItems((items) => sortScheduleItemsChronologically(items.map((item) => item.id === id ? { ...item, time: normalizeTime(value) } : item)))

  const continueToInvitees = () => {
    if (validateInvitationDetails()) {
      setFieldErrors({})
      setStep(3)
    }
  }

  const updateGuestGroup = (id: string, update: Partial<GuestGroup>) => {
    setGuestGroups((groups) => groups.map((group) => group.id === id ? { ...group, ...update } : group))
    if (update.name !== undefined) clearFieldError(`invitee-${id}`)
  }
  const addGuestGroup = () => {
    setGuestGroups((groups) => [...groups, makeGuestGroup()])
    clearFieldError('guestRows')
  }
  const removeGuestGroup = (id: string) => {
    setGuestGroups((groups) => groups.filter((group) => group.id !== id))
    clearFieldError(`invitee-${id}`)
  }

  const startNewInvitation = () => {
    setEditingId(null); setKind('Wedding'); setTitle(''); setDate(''); setTime(''); setPlace(''); setMapsUrl(''); setGuestGroups([makeGuestGroup()]); setScheduleItems([]); setStep(1); transitionTo('create')
  }

  const addGuest = () => {
    const name = newGuest.trim()
    if (!name) {
      setFieldErrors((errors) => ({ ...errors, manageNewGuest: 'Enter a name for this invitee.' }))
      document.getElementById('manage-new-guest-name')?.focus()
      return
    }
    if (manageId === null) return
    const count = Math.max(1, Number.parseInt(newGuestCount, 10) || 1)
    setInvites((all) => all.map((invite) => {
      if (invite.id !== manageId || guestGroupsForInvite(invite).some((group) => group.name.toLowerCase() === name.toLowerCase())) return invite
      const group = makeGuestGroup(name, count)
      return { ...invite, guests: invite.guests + count, guestNames: [...invite.guestNames, name], guestGroups: [...guestGroupsForInvite(invite), group] }
    }))
    setNewGuest('')
    clearFieldError('manageNewGuest')
    setNewGuestCount('1')
  }

  const removeGuest = (name: string) => {
    if (manageId === null) return
    setInvites((all) => all.map((invite) => {
      if (invite.id !== manageId) return invite
      const groups = guestGroupsForInvite(invite)
      const removedCount = groups.filter((group) => group.name === name).reduce((total, group) => total + group.count, 0)
      const remainingGroups = groups.filter((group) => group.name !== name)
      return { ...invite, guests: Math.max(0, invite.guests - removedCount), guestNames: invite.guestNames.filter((guest) => guest !== name), guestGroups: remainingGroups, tables: invite.tables.map((table) => ({ ...table, guests: table.guests.filter((guest) => guest !== name) })) }
    }))
  }

  const updateManagedGuestCount = (guestId: string, countValue: string) => {
    if (manageId === null) return
    const count = Math.max(1, Number.parseInt(countValue, 10) || 1)
    setInvites((all) => all.map((invite) => {
      if (invite.id !== manageId) return invite
      const guestGroups = guestGroupsForInvite(invite).map((group) => group.id === guestId ? { ...group, count } : group)
      return { ...invite, guestGroups, guests: guestGroups.reduce((total, group) => total + group.count, 0) }
    }))
  }

  const updateManagedGuestGroup = (guestId: string, update: Partial<GuestGroup>) => {
    if (manageId === null) return
    setInvites((all) => all.map((invite) => {
      if (invite.id !== manageId) return invite
      const oldGroups = guestGroupsForInvite(invite)
      const oldName = oldGroups.find((group) => group.id === guestId)?.name
      const guestGroups = oldGroups.map((group) => group.id === guestId ? { ...group, ...update } : group)
      const guestNames = guestGroups.map((group) => group.name)
      const tables = update.name !== undefined && oldName ? invite.tables.map((table) => ({ ...table, guests: table.guests.map((guest) => guest === oldName ? update.name! : guest) })) : invite.tables
      return { ...invite, guestGroups, guestNames, tables, guests: guestGroups.reduce((total, group) => total + Math.max(1, Number(group.count) || 1), 0) }
    }))
  }

  const addTable = () => {
    if (!newTable.trim()) {
      setFieldErrors((errors) => ({ ...errors, manageNewTable: 'Enter a name for this table.' }))
      document.getElementById('manage-new-table-name')?.focus()
      return
    }
    if (manageId === null) return
    setInvites((all) => all.map((invite) => invite.id === manageId ? { ...invite, tables: [...invite.tables, { name: newTable.trim(), guests: [] }] } : invite))
    setNewTable('')
    clearFieldError('manageNewTable')
  }

  const assignGuest = (tableIndex: number, guestName: string) => {
    if (manageId === null) return
    setInvites((all) => all.map((invite) => invite.id === manageId ? { ...invite, tables: invite.tables.map((table, index) => ({ ...table, guests: index === tableIndex ? [...new Set([...table.guests, guestName])] : table.guests.filter((guest) => guest !== guestName) })) } : invite))
  }

  const unassignGuest = (tableIndex: number, guestName: string) => {
    if (manageId === null) return
    setInvites((all) => all.map((invite) => invite.id === manageId ? { ...invite, tables: invite.tables.map((table, index) => ({ ...table, guests: index === tableIndex ? table.guests.filter((guest) => guest !== guestName) : table.guests })) } : invite))
  }

  const updateCoverPhoto = async (inviteId: number, file?: File) => {
    if (!file) return
    if (!file.type.startsWith('image/')) { setNotice('Choose an image file to use as the cover.'); return }
    if (file.size > 12 * 1024 * 1024) { setNotice('Choose an image smaller than 12 MB.'); return }
    try {
      const coverImage = await compressCoverImage(file)
      setInvites((all) => all.map((invite) => invite.id === inviteId ? { ...invite, coverImage, coverImagePosition: { x: 50, y: 50, zoom: 100 } } : invite))
      setNotice('Cover photo updated.')
    } catch (error) {
      setNotice(error instanceof Error ? error.message : 'Could not update the cover photo.')
    }
    window.setTimeout(() => setNotice(''), 5000)
  }

  const removeCoverPhoto = (inviteId: number) => {
    setInvites((all) => all.map((invite) => invite.id === inviteId ? { ...invite, coverImage: undefined, coverImagePosition: undefined } : invite))
    setCoverEditorId(null)
  }

  const deleteInvitation = (inviteId: number) => {
    setInvites((all) => all.filter((invite) => invite.id !== inviteId))
    if (manageId === inviteId) setManageMode(null)
    if (coverEditorId === inviteId) setCoverEditorId(null)
    setDeleteConfirmId(null)
    setNotice('Invitation deleted.')
    window.setTimeout(() => setNotice(''), 5000)
  }

  const useRandomCoverPhoto = (inviteId: number) => {
    setInvites((all) => all.map((invite) => {
      if (invite.id !== inviteId) return invite
      const choices = randomCoverPhotos.filter((photo) => photo !== invite.coverImage)
      return { ...invite, coverImage: choices[Math.floor(Math.random() * choices.length)] ?? randomCoverPhotos[0], coverImagePosition: { x: 50, y: 50, zoom: 100 } }
    }))
  }

  const updateCoverPosition = (inviteId: number, field: 'x' | 'y' | 'zoom', value: number) => {
    setInvites((all) => all.map((invite) => invite.id === inviteId ? { ...invite, coverImagePosition: { x: invite.coverImagePosition?.x ?? 50, y: invite.coverImagePosition?.y ?? 50, zoom: invite.coverImagePosition?.zoom ?? 100, [field]: value } } : invite))
  }

  const downloadGuestExcel = (invite: Invite) => {
    const groups = guestGroupsForInvite(invite)
    const rows = [['Invitee', 'Guests included', 'RSVP status', 'Invitation sent', 'Sent via', 'Assigned table', 'Invitation'], ...groups.map((group) => [group.name, String(group.count), group.rsvp ?? 'Awaiting reply', group.invitationSent ? 'Yes' : 'No', group.sentVia ?? '', invite.tables.find((table) => table.guests.includes(group.name))?.name ?? '', invite.title])]
    const blob = makeXlsx(rows)
    const url = URL.createObjectURL(blob)
    const link = document.createElement('a')
    link.href = url
    link.download = `${invite.title.replace(/[^a-z0-9]+/gi, '-').replace(/^-|-$/g, '').toLowerCase() || 'invitation'}-guests.xlsx`
    link.click()
    window.setTimeout(() => URL.revokeObjectURL(url), 1000)
  }

  const downloadInvitationImage = async (invite: Invite) => {
    const canvas = document.createElement('canvas')
    canvas.width = 1200
    canvas.height = 1500
    const context = canvas.getContext('2d')
    if (!context) { setNotice('Could not create the invitation image.'); return }
    context.fillStyle = '#f8f6ef'
    context.fillRect(0, 0, canvas.width, canvas.height)
    if (invite.coverImage) {
      try {
        const image = new Image()
        if (!invite.coverImage.startsWith('data:')) image.crossOrigin = 'anonymous'
        await new Promise<void>((resolve, reject) => { image.onload = () => resolve(); image.onerror = () => reject(new Error('Could not load the cover photo.')); image.src = invite.coverImage! })
        const ratio = Math.max(1200 / image.width, 620 / image.height)
        const width = image.width * ratio, height = image.height * ratio
        context.save()
        context.beginPath(); context.rect(0, 0, 1200, 620); context.clip()
        context.drawImage(image, (1200 - width) / 2, (620 - height) / 2, width, height)
        context.restore()
        context.fillStyle = '#283b2b38'; context.fillRect(0, 0, 1200, 620)
      } catch { setNotice('The cover photo could not be added. Downloading the invitation without it.'); window.setTimeout(() => setNotice(''), 4500) }
    } else {
      const gradient = context.createLinearGradient(0, 0, 1200, 620)
      gradient.addColorStop(0, '#e9eee5'); gradient.addColorStop(1, '#f1e5d9')
      context.fillStyle = gradient; context.fillRect(0, 0, 1200, 620)
      context.fillStyle = '#718369'; context.globalAlpha = .18
      for (const [x, y, radius] of [[160,140,100],[1020,430,180],[920,90,75]] as number[][]) { context.beginPath(); context.arc(x, y, radius, 0, Math.PI * 2); context.fill() }
      context.globalAlpha = 1
    }
    context.fillStyle = '#f8f6ef'; context.fillRect(0, 620, 1200, 880)
    context.textAlign = 'center'
    context.fillStyle = '#7b8b71'; context.font = '600 24px Arial'; context.fillText(invite.type.toUpperCase(), 600, 740)
    context.fillStyle = '#35453a'; context.font = '60px Georgia'
    const title = invite.title
    const words = title.split(/\s+/); let line = ''; let y = 850
    for (const word of words) { const next = line ? `${line} ${word}` : word; if (context.measureText(next).width > 970 && line) { context.fillText(line, 600, y); y += 76; line = word } else line = next }
    if (line) context.fillText(line, 600, y)
    context.strokeStyle = '#c7d0be'; context.lineWidth = 2; context.beginPath(); context.moveTo(460, y + 48); context.lineTo(740, y + 48); context.stroke()
    context.fillStyle = '#687667'; context.font = '30px Arial'; context.fillText(invite.date, 600, y + 115)
    context.font = '25px Arial'; context.fillText(invite.place, 600, y + 163, 980)
    let scheduleY = y + 245
    context.fillStyle = '#405b46'; context.font = '600 23px Arial'
    for (const item of invite.schedule.slice(0, 5)) { context.fillText(`${item.time || '—'}  ·  ${item.title}`, 600, scheduleY); scheduleY += 48 }
    context.textAlign = 'center'; context.fillStyle = '#93998e'; context.font = '20px Arial'; context.fillText('INVITECYPRUS · CELEBRATE TOGETHER', 600, 1440)
    canvas.toBlob((blob) => {
      if (!blob) { setNotice('Could not create the invitation image.'); return }
      const url = URL.createObjectURL(blob)
      const link = document.createElement('a')
      link.href = url; link.download = `${invite.title.replace(/[^a-z0-9]+/gi, '-').replace(/^-|-$/g, '').toLowerCase() || 'invitation'}.png`; link.click()
      window.setTimeout(() => URL.revokeObjectURL(url), 1000)
    }, 'image/png')
  }

  const managedInvite = invites.find((invite) => invite.id === manageId)
  const coverBeingEdited = invites.find((invite) => invite.id === coverEditorId)
  const invitationToDelete = invites.find((invite) => invite.id === deleteConfirmId)
  const invitationToPreview = invites.find((invite) => invite.id === previewInviteId)

  if (access !== 'open') return <main className="access-screen"><div className="access-card"><div className="access-brand"><span className="simple-mark"><i/><i/><i/><i/></span>invitecyprus</div><span className="login-icon"><LockKeyhole size={19}/></span><p className="simple-overline">PRIVATE PREVIEW</p><h1>{access === 'checking' ? 'Checking access…' : access === 'setup' ? 'Set up preview access' : 'Enter the password'}</h1>{access === 'setup' ? <p className="access-copy">Create a <code>.env.local</code> file in the project folder and add <code>INVITECYPRUS_ACCESS_PASSWORD=your-password</code>. Restart the dev server to apply it.</p> : access === 'checking' ? <p className="access-copy">One moment while we check this preview.</p> : <form onSubmit={unlockPreview}><p className="access-copy">Enter the preview password to continue.</p><label htmlFor="preview-password">Password</label><input id="preview-password" type="password" autoFocus autoComplete="current-password" value={accessPassword} onChange={(e) => setAccessPassword(e.target.value)} placeholder="Enter preview password" required/><button className="simple-primary full-button" type="submit">Open invitecyprus <ArrowRight size={15}/></button>{accessError && <p className="access-error">{accessError}</p>}</form>}</div></main>

  return <div className="simple-app">
    <header className="simple-header"><button className="simple-brand" onClick={() => transitionTo(screen === 'login' ? 'login' : 'home')}><span className="simple-mark"><i/><i/><i/><i/></span>invitecyprus</button>{screen !== 'login' && <div className="account-menu-wrap" onBlur={(event) => {if (!event.currentTarget.contains(event.relatedTarget as Node | null)) setProfileMenuOpen(false)}}><button type="button" className="account-chip" aria-haspopup="dialog" aria-expanded={profileMenuOpen} onClick={() => {setProfileDraftName(profileName);setProfileDraftEmail(profileEmail);setProfileSaved(false);setProfileMenuOpen((open) => !open)}}><span className="account-initial">{profileName.trim().charAt(0).toUpperCase() || 'U'}</span><span>{profileName}</span><ChevronRight size={14}/></button>{profileMenuOpen && <section className="profile-menu" role="dialog" aria-label="Profile settings"><div className="profile-menu-heading"><CircleUserRound size={17}/><span><strong>Profile settings</strong><small>Your account details</small></span></div><div className="profile-settings-fields"><label>Display name<input required value={profileDraftName} onChange={(event) => { setProfileDraftName(event.target.value); clearFieldError('profileDisplayName') }} placeholder="Your name"/>{fieldErrors.profileDisplayName && <small className="field-error">{fieldErrors.profileDisplayName}</small>}</label><label>Email address<input type="email" value={profileDraftEmail} readOnly placeholder="you@example.com"/></label><button className="profile-save-button" onClick={() => void saveProfileSettings()}><Save size={13}/>{profileSaved ? 'Saved' : 'Save changes'}</button></div><div className="profile-payment-section"><button className="profile-payment-disabled" disabled><CreditCard size={16}/><span><strong>Payment options</strong><small>Currently unavailable</small></span></button><p>Invitecyprus is free for a limited time. Payment options will be available later.</p></div><button className="profile-menu-logout" onClick={() => void logOut()}><LogOut size={14}/> Log out</button></section>}</div>}</header>

    {screen === 'login' && <main className="login-layout"><section className="login-welcome"><div className="welcome-art"><div className="art-photo"></div><div className="art-note"><span>✳</span><strong>A little invite.<br/>A lovely big moment.</strong></div><div className="art-circle"></div><div className="art-caption">INVITECYPRUS · CELEBRATE TOGETHER</div></div><div className="login-welcome-copy"><p className="simple-overline">FOR ALL THE MOMENTS THAT MATTER</p><h1>Bring your people<br/>a little closer.</h1><p>Beautiful invitations and easy RSVPs, all in one place.</p></div></section><section className="login-panel"><div className="login-form"><span className="login-icon"><CircleUserRound size={20}/></span><p className="simple-overline">YOUR INVITECYPRUS ACCOUNT</p><h2>{authMode === 'signUp' ? 'Create your account' : 'Welcome back'}</h2><p className="login-sub">{authMode === 'signUp' ? 'A few details, then you can start your invitation.' : 'Log in to create and manage your invitations.'}</p>{!firebaseConfigured && <><p className="auth-setup-note">Firebase sign-in isn’t set up yet. Continue with a local preview while you configure it.</p><button className="simple-primary full-button local-preview-button" type="button" onClick={enterLocalPreview}>Continue in local preview <ArrowRight size={16}/></button></>}<form onSubmit={(event) => void submitAuthForm(event)}>{authMode === 'signUp' && <><label htmlFor="auth-name">Your name</label><input id="auth-name" autoComplete="name" value={authName} onChange={(event) => setAuthName(event.target.value)} placeholder="e.g. Emma Wilson" required/></>}<label htmlFor="auth-email">Email address</label><input id="auth-email" type="email" autoComplete="email" value={authEmail} onChange={(event) => setAuthEmail(event.target.value)} placeholder="you@example.com" required/><label htmlFor="auth-password">Password</label><input id="auth-password" type="password" autoComplete={authMode === 'signUp' ? 'new-password' : 'current-password'} minLength={authMode === 'signUp' ? 8 : undefined} value={authPassword} onChange={(event) => setAuthPassword(event.target.value)} placeholder={authMode === 'signUp' ? 'At least 8 characters' : 'Enter your password'} required/>{authMode === 'signUp' && <><label htmlFor="auth-password-confirm">Confirm password</label><input id="auth-password-confirm" type="password" autoComplete="new-password" value={authPasswordConfirm} onChange={(event) => setAuthPasswordConfirm(event.target.value)} placeholder="Enter your password again" required/></>}{authMode === 'signIn' && <button className="forgot-link" type="button" onClick={() => void requestPasswordReset()} disabled={authBusy}>Forgot password?</button>}<button className="simple-primary full-button" type="submit" disabled={!firebaseConfigured || !authReady || authBusy}>{authBusy ? 'Please wait…' : authMode === 'signUp' ? 'Create account' : 'Log in'} <ArrowRight size={16}/></button></form>{authError && <p className="auth-feedback auth-feedback-error" role="alert">{authError}</p>}{authMessage && <p className="auth-feedback auth-feedback-success" role="status">{authMessage}</p>}<div className="login-divider"><span></span>or continue with<span></span></div><button className="google-button" type="button" onClick={() => void signInWithProvider('google')} disabled={!firebaseConfigured || !authReady || authBusy}><span>G</span> Continue with Google</button><button className="facebook-button" type="button" onClick={() => void signInWithProvider('facebook')} disabled={!firebaseConfigured || !authReady || authBusy}><span>f</span> Continue with Facebook</button><p className="signup-line">{authMode === 'signUp' ? 'Already have an account?' : 'New to invitecyprus?'} <button type="button" onClick={() => {setAuthMode(authMode === 'signUp' ? 'signIn' : 'signUp');setAuthError('');setAuthMessage('')}}>{authMode === 'signUp' ? 'Log in' : 'Create an account'}</button></p><p className="privacy-line"><LockKeyhole size={12}/> Your guests won’t need an account to RSVP.</p></div></section></main>}

    {screen === 'home' && <main className="simple-main home-screen"><div className="simple-heading"><p className="simple-overline">YOUR CELEBRATIONS, IN ONE PLACE</p><h1>What would you like to do?</h1><p>It only takes a few minutes to bring everyone together.</p></div><div className="choice-grid"><button className="choice-card choice-create" onClick={startNewInvitation}><span className="choice-icon"><Plus size={20}/></span><span className="choice-label">START SOMETHING NEW</span><strong>What do you want<br/>to invite people to?</strong><span className="choice-description">A wedding, birthday, baptism, work event, or anything worth celebrating.</span><span className="choice-action">Create an invitation <ArrowRight size={16}/></span><span className="choice-flower">✳</span></button><button className="choice-card choice-manage" onClick={() => transitionTo('manage')}><span className="choice-icon"><Users size={19}/></span><span className="choice-label">PICK UP WHERE YOU LEFT OFF</span><strong>Manage invitations</strong><span className="choice-description">Check guest replies, update the details, or send your invitation link.</span><span className="choice-action">View my invitations <ArrowRight size={16}/></span><span className="manage-preview"><span>O&J</span><span>✳</span><span>+{invites.length}</span></span></button></div><p className="home-reassurance"><LockKeyhole size={13}/> Guests can open and RSVP to your invitation without logging in.</p></main>}

    {screen === 'create' && <div className={editingId ? 'simple-modal-backdrop edit-invitation-backdrop' : ''} onMouseDown={(event) => {if (editingId && event.target === event.currentTarget) {setEditingId(null); setFieldErrors({}); transitionTo('invitation-dashboard')}}}><main className={`simple-main create-screen ${editingId ? 'edit-invitation-modal' : ''}`} role={editingId ? 'dialog' : undefined} aria-modal={editingId ? true : undefined} aria-label={editingId ? 'Edit invitation' : undefined}>{editingId && <button className="people-close" aria-label="Close edit invitation" onClick={() => {setEditingId(null); setFieldErrors({}); transitionTo('invitation-dashboard')}}><X size={17}/></button>}{!editingId && <button className="back-link" onClick={back}><ArrowLeft size={14}/>{step === 1 ? 'Back to choices' : 'Previous step'}</button>}<div className="wizard-wrap">{!editingId && <div className="wizard-top"><p className="simple-overline">LET’S GET THIS CELEBRATION STARTED</p><div className="wizard-progress"><span className={step >= 1 ? 'current' : ''}></span><span className={step >= 2 ? 'current' : ''}></span><span className={step >= 3 ? 'current' : ''}></span></div><small>STEP {step} OF 3</small></div>}
      {step === 1 && <section key={step} className="wizard-step"><h1>What are we celebrating?</h1><p>Choose the kind of invitation you’d like to create.</p><div className="occasion-grid">{occasions.map((occasion) => <button key={occasion} className={`occasion-option ${kind === occasion ? 'picked' : ''}`} onClick={() => {if (kind !== occasion) setScheduleItems([]); setKind(occasion)}}><span className="occasion-symbol">{occasion === 'Wedding' ? '♡' : occasion === 'Birthday' ? '✳' : occasion === 'Baptism' ? '⌁' : occasion === 'Company event' ? '▧' : occasion === 'Dinner party' ? '◌' : '✦'}</span>{occasion}{kind === occasion && <Check size={15}/>}</button>)}</div><button className="simple-primary wizard-next" onClick={() => setStep(2)}>Next <ArrowRight size={16}/></button></section>}
      {step === 2 && <section key={step} className="wizard-step"><h1>{editingId ? 'Edit invitation details.' : 'Add the important details.'}</h1><p>You can change these details later.</p><label htmlFor="invite-title">Give your invitation a name</label><input id="invite-title" required value={title} onChange={e => { setTitle(e.target.value); clearFieldError('invitationTitle') }} placeholder={kind === 'Wedding' ? 'e.g. Olivia & James’ wedding' : `e.g. My ${kind.toLowerCase()}`}/>{fieldErrors.invitationTitle && <small className="field-error">{fieldErrors.invitationTitle}</small>}<div className="form-pair form-pair-single"><div><label htmlFor="invite-date">Date</label><input id="invite-date" className="date-time-input date-input" type="date" lang="en-GB" required value={date} onChange={e => { setDate(e.target.value); setScheduleItems((items) => sortScheduleItemsChronologically(items)); clearFieldError('eventDate') }}/>{fieldErrors.eventDate && <small className="field-error">{fieldErrors.eventDate}</small>}</div></div><div className="schedule-builder"><div className="schedule-builder-heading"><div><h2>Event-day schedule</h2><p>Choose at least one moment for your {kind.toLowerCase()}.</p></div><Clock3 size={19}/></div><div className="schedule-suggestions">{(scheduleSuggestions[kind] ?? scheduleSuggestions['Something else']).map((suggestion) => {const added = scheduleItems.some((item) => item.title.toLowerCase() === suggestion.toLowerCase()); return <button key={suggestion} className={added ? 'suggestion-added' : ''} disabled={added} onClick={() => addScheduleItem(suggestion)}>{added ? <Check size={13}/> : <Plus size={13}/>} {suggestion}</button>})}<button onClick={() => addScheduleItem('New moment')}><Plus size={13}/> Custom moment</button></div>{fieldErrors.scheduleSelection && <small className="field-error schedule-selection-error">{fieldErrors.scheduleSelection}</small>}{scheduleItems.length > 0 && <div className="schedule-editor">{scheduleItems.map((item) => <article className="schedule-edit-item" key={item.id}><div className="schedule-edit-title"><span className="schedule-dot"></span><div className="schedule-title-field"><input id={`schedule-title-${item.id}`} aria-label="Schedule item title" required value={item.title} onChange={(event) => updateScheduleItem(item.id, { title: event.target.value })}/>{fieldErrors[`schedule-title-${item.id}`] && <small className="field-error">{fieldErrors[`schedule-title-${item.id}`]}</small>}</div><button type="button" aria-label={`Remove ${item.title}`} onClick={() => removeScheduleItem(item.id)}><Trash2 size={16}/></button></div><div className="schedule-edit-fields"><div className="schedule-time-field"><label><Clock3 size={14}/><input aria-label="Schedule time (24-hour HH:MM)" id={`schedule-time-${item.id}`} required className="schedule-time-input" type="text" inputMode="numeric" autoComplete="off" maxLength={5} pattern="(?:[01][0-9]|2[0-3]):[0-5][0-9]" placeholder="HH:MM" value={item.time} onChange={(event) => updateScheduleItem(item.id, { time: event.target.value })} onBlur={() => commitScheduleTime(item.id, item.time)}/></label>{fieldErrors[`schedule-time-${item.id}`] && <small className="field-error">{fieldErrors[`schedule-time-${item.id}`]}</small>}</div><label className="schedule-place-field"><MapPin size={14}/><input aria-label="Schedule place or Google Maps link" value={item.place} onChange={(event) => updateScheduleItem(item.id, { place: event.target.value })} placeholder="Place or paste a Google Maps link"/></label>{(item.place || item.mapsUrl) && <a href={googleMapsHref(item.place, item.mapsUrl)} target="_blank" rel="noreferrer"><ExternalLink size={13}/> Map</a>}</div></article>)}</div>}</div><button className="simple-primary wizard-next" onClick={editingId ? create : continueToInvitees}>{editingId ? <><Save size={15}/> Save changes</> : <>Next <ArrowRight size={16}/></>}</button></section>}
      {step === 3 && <section key={step} className="wizard-step"><h1>Who should we invite?</h1><p>Add a person or group and choose how many guests they include.</p><div className="guest-group-heading"><span>Invitee</span><span>Guests included</span></div><div className="guest-group-editor">{guestGroups.map((group, index) => <div className="guest-group-row" key={group.id}><span className="guest-group-index">{index + 1}</span><div className="guest-name-field"><input aria-label={`Invitee name ${index + 1}`} value={group.name} onChange={(event) => updateGuestGroup(group.id, { name: event.target.value })} placeholder="e.g. Thomas or Thomas family" required/>{fieldErrors[`invitee-${group.id}`] && <small className="field-error">{fieldErrors[`invitee-${group.id}`]}</small>}</div><input aria-label={`Number of guests for invitee ${index + 1}`} type="number" min="1" required step="1" inputMode="numeric" defaultValue={group.count} onBlur={(event) => { const count = Math.max(1, Number.parseInt(event.currentTarget.value, 10) || 1); event.currentTarget.value = String(count); updateGuestGroup(group.id, { count }) }}/><button type="button" aria-label={`Remove ${group.name || `invitee ${index + 1}`}`} onClick={() => removeGuestGroup(group.id)}><X size={16}/></button></div>)}<button type="button" className="guest-group-add" onClick={addGuestGroup}><Plus size={15}/> Add another invitee</button>{fieldErrors.guestRows && <small className="field-error">{fieldErrors.guestRows}</small>}</div><div className="guest-group-summary"><Users size={15}/><strong>{guestGroups.filter((group) => group.name.trim()).length}</strong> invitees <span>·</span> <strong>{guestGroups.filter((group) => group.name.trim()).reduce((total, group) => total + Math.max(1, Number(group.count) || 1), 0)}</strong> guests included</div><div className="private-note"><Users size={14}/> Each invitee starts with 1 guest. Increase the number for a partner or family. You can add or remove invitees later.</div><button className="simple-primary wizard-next" onClick={create}>{editingId ? 'Save changes' : 'Create invitation'} <ArrowRight size={16}/></button></section>}
      {!editingId && <div className="wizard-foot"><span><LockKeyhole size={12}/> Your invitation starts as a draft.</span><span><i>1</i> Choose <i>2</i> Details <i>3</i> Guests</span></div>}</div></main></div>}

    {screen === 'manage' && <main className="simple-main manage-screen"><button className="back-link" onClick={() => transitionTo('home')}><ArrowLeft size={14}/> Back to choices</button><div className="manage-heading"><div><p className="simple-overline">YOUR CELEBRATIONS</p><h1>Manage invitations</h1><p>Keep all your event details and replies together.</p></div><button className="simple-primary" onClick={startNewInvitation}><Plus size={16}/> New invitation</button></div>{invites.length ? <div className="managed-list">{invites.map((invite, i) => <article className="managed-card" key={invite.id} tabIndex={0} aria-label={`Open dashboard for ${invite.title}`} onClick={() => openInvitationDashboard(invite.id)} onKeyDown={(event) => {if (event.target === event.currentTarget && (event.key === 'Enter' || event.key === ' ')) {event.preventDefault(); openInvitationDashboard(invite.id)}}}><div className={`managed-cover cover-${i % 3} ${invite.coverImage ? 'has-cover-image' : ''}`}>{invite.coverImage && <img className="managed-cover-photo" src={invite.coverImage} alt={`Cover for ${invite.title}`} style={{ objectPosition: '50% 50%', transform: coverTransform(invite.coverImagePosition) }}/>}<span className="managed-cover-title">{invite.type === 'Wedding' ? 'O&J' : invite.type === 'Birthday' ? '✳' : invite.type.substring(0,2).toUpperCase()}</span>{invite.coverImage && <button className="cover-remove-control" aria-label={`Remove cover photo for ${invite.title}`} onClick={(event) => {event.stopPropagation(); removeCoverPhoto(invite.id)}}><X size={13}/></button>}<input id={`cover-upload-${invite.id}`} className="cover-file-input" type="file" onClick={(event) => event.stopPropagation()} accept="image/*" onChange={(event) => { void updateCoverPhoto(invite.id, event.target.files?.[0]); event.currentTarget.value = '' }}/><div className="cover-photo-actions"><label className="cover-upload-control" htmlFor={`cover-upload-${invite.id}`} onClick={(event) => event.stopPropagation()}><ImagePlus size={14}/>{invite.coverImage ? 'Replace photo' : 'Add photo'}</label><button className="cover-crop-control" onClick={(event) => {event.stopPropagation(); setCoverEditorId(invite.id)}}><Crop size={13}/>{invite.coverImage ? 'Photo options' : 'Choose photo'}</button></div></div><div className="managed-content"><span className={`draft-label ${isInviteCompleted(invite) ? 'completed-label' : ''}`}><i/> {isInviteCompleted(invite) ? 'COMPLETED' : 'DRAFT'}</span><h2>{invite.title}</h2><p><CalendarDays size={14}/>{invite.date}</p><p><MapPin size={14}/>{invite.place}</p><span className="managed-open-hint">Open invitation dashboard <ArrowRight size={14}/></span></div><button className="managed-delete-button" aria-label={`Delete ${invite.title}`} title="Delete invitation" onClick={(event) => {event.stopPropagation(); setDeleteConfirmId(invite.id)}}><Trash2 size={15}/></button></article>)}</div> : <div className="empty-invites"><span><ImagePlus size={22}/></span><h2>Your invitations will live here.</h2><p>Create your first invitation and we’ll keep everything organised for you.</p><button className="simple-primary" onClick={startNewInvitation}><Plus size={16}/> Create invitation</button></div>}<div className="manage-tip"><span className="tip-star">✳</span><p><strong>A little heads-up</strong><br/>Your guests can RSVP from their invitation link. You can add and manage guests at any time.</p></div></main>}

    {(screen === 'invitation-dashboard' || (screen === 'create' && editingId !== null)) && managedInvite && <main className="simple-main invitation-dashboard"><button className="back-link" onClick={() => transitionTo('manage')}><ArrowLeft size={14}/> All invitations</button><section className="invitation-dashboard-hero"><div className="dashboard-invitation-heading"><p className="simple-overline">INVITATION DASHBOARD</p><span className={`draft-label ${isInviteCompleted(managedInvite) ? 'completed-label' : ''}`}><i/> {isInviteCompleted(managedInvite) ? 'COMPLETED' : 'DRAFT'}</span><h1>{managedInvite.title}</h1><p>{managedInvite.type}</p></div><div className="dashboard-detail-grid"><div><CalendarDays size={16}/><span><small>DATE</small><strong>{managedInvite.date}</strong></span></div><div><MapPin size={16}/><span><small>LOCATION</small><strong>{managedInvite.place}</strong></span>{(managedInvite.place || managedInvite.mapsUrl) && managedInvite.place !== 'See event-day schedule for location details' && <a href={googleMapsHref(managedInvite.place, managedInvite.mapsUrl)} target="_blank" rel="noreferrer">Map <ExternalLink size={12}/></a>}</div>{managedInvite.schedule.length > 0 && <div className="dashboard-detail-schedule"><Clock3 size={16}/><span><small>EVENT-DAY SCHEDULE</small>{managedInvite.schedule.map((item) => <strong key={item.id}>{item.time || 'Time TBD'} · {item.title}{item.place ? ` — ${item.place}` : ''}</strong>)}</span></div>}</div><div className="dashboard-hero-actions"><button className="simple-secondary" onClick={() => setPreviewInviteId(managedInvite.id)}><Eye size={15}/> Preview invitation</button><button className="simple-primary" onClick={() => editInvite(managedInvite)}><Save size={15}/> Edit invitation</button></div></section><div className="dashboard-stats"><article><Users size={17}/><span><strong>{guestGroupsForInvite(managedInvite).length}</strong><small>Invitees</small></span></article><article><Users size={17}/><span><strong>{managedInvite.guests}</strong><small>Guests included</small></span></article><article><Clock3 size={17}/><span><strong>{managedInvite.schedule.length}</strong><small>Schedule moments</small></span></article><article><CircleUserRound size={17}/><span><strong>{managedInvite.tables.length}</strong><small>Seating tables</small></span></article></div><section className="dashboard-maintenance"><div className="dashboard-maintenance-heading"><div><p className="simple-overline">INVITATION MAINTENANCE</p><h2>Manage your celebration</h2></div><p>Everything you need to keep the event organized.</p></div><div className="dashboard-action-grid"><button className="dashboard-tool-primary" onClick={() => openManageMode(managedInvite.id, 'guests')}><span><Users size={20}/></span><strong>Guest list <em>VIEW ALL</em></strong><small>Update RSVP replies, guest counts, and delivery details.</small><ArrowRight size={15}/></button><button className="dashboard-tool-primary" onClick={() => openManageMode(managedInvite.id, 'seating')}><span><CircleUserRound size={20}/></span><strong>Seating arrangements <em>VIEW ALL</em></strong><small>See every table and assign guests to seats.</small><ArrowRight size={15}/></button><button onClick={() => downloadGuestExcel(managedInvite)}><span><Download size={18}/></span><strong>Download Excel file</strong><small>Export invitees, guest counts, and table assignments.</small><ArrowRight size={15}/></button><button onClick={() => setPreviewInviteId(managedInvite.id)}><span><Eye size={18}/></span><strong>Preview invitation</strong><small>See how the invitation details look together.</small><ArrowRight size={15}/></button><button onClick={() => void downloadInvitationImage(managedInvite)}><span><ImagePlus size={18}/></span><strong>Create invitation image</strong><small>Download a shareable PNG using your event details and cover photo.</small><ArrowRight size={15}/></button></div><p className="ai-image-note">AI artwork generation needs an image provider connection. The PNG export is ready to use with your current invitation design.</p></section><section className="dashboard-cover-maintenance"><div className={`dashboard-cover-preview ${managedInvite.coverImage ? 'has-cover-image' : ''}`}>{managedInvite.coverImage && <img src={managedInvite.coverImage} alt="Invitation cover" style={{ objectPosition: '50% 50%', transform: coverTransform(managedInvite.coverImagePosition) }}/>}<span>{managedInvite.coverImage ? '' : <ImagePlus size={22}/>}</span></div><div><p className="simple-overline">INVITATION DESIGN</p><h2>Cover photo</h2><p>Choose or adjust the photo guests will see on your invitation.</p></div><button className="simple-secondary" onClick={() => setCoverEditorId(managedInvite.id)}><Crop size={14}/>{managedInvite.coverImage ? 'Edit photo' : 'Add photo'}</button></section></main>}

    {screen === 'guest-management' && managedInvite && <main className="simple-main dedicated-management-page"><button className="back-link" onClick={() => transitionTo('invitation-dashboard')}><ArrowLeft size={14}/> Invitation dashboard</button><div className="dedicated-page-heading"><div><p className="simple-overline">INVITATION MAINTENANCE</p><h1>Guest list</h1><p>{managedInvite.title} · Track attendance and how each invitation was sent.</p></div><button className="simple-secondary" onClick={() => downloadGuestExcel(managedInvite)}><Download size={14}/> Download Excel</button></div><div className="guest-rsvp-summary"><article><strong>{guestGroupsForInvite(managedInvite).length}</strong><span>Invitees</span></article><article><strong>{guestGroupsForInvite(managedInvite).filter((group) => group.rsvp === 'attending').reduce((sum, group) => sum + group.count, 0)}</strong><span>Attending</span></article><article><strong>{guestGroupsForInvite(managedInvite).filter((group) => group.rsvp === 'declined').reduce((sum, group) => sum + group.count, 0)}</strong><span>Declined</span></article><article><strong>{guestGroupsForInvite(managedInvite).filter((group) => !group.rsvp || group.rsvp === 'pending').reduce((sum, group) => sum + group.count, 0)}</strong><span>Awaiting reply</span></article></div><section className="guest-management-panel"><div className="guest-add-toolbar"><div><h2>Invitees</h2><p>Add people now, or update their RSVP as replies arrive.</p></div><div className="guest-add-controls"><input id="manage-new-guest-name" aria-label="Invitee name" required value={newGuest} onChange={(event) => {setNewGuest(event.target.value);clearFieldError('manageNewGuest')}} onKeyDown={(event) => {if (event.key === 'Enter') {event.preventDefault();addGuest()}}} placeholder="Name or family name"/><label><span>Guests</span><input aria-label="Number of guests" type="number" min="1" step="1" value={newGuestCount} onChange={(event) => setNewGuestCount(event.target.value)}/></label><button className="simple-primary" onClick={addGuest}><Plus size={14}/> Add invitee</button></div>{fieldErrors.manageNewGuest && <small className="field-error guest-add-error">{fieldErrors.manageNewGuest}</small>}</div><div className="guest-list-table-head"><span>Invitee</span><span>Guests</span><span>RSVP</span><span>Invitation</span><span>Sent via</span><span></span></div>{guestGroupsForInvite(managedInvite).length ? <div className="guest-management-list">{guestGroupsForInvite(managedInvite).map((group, index) => <article className="guest-management-row" key={group.id}><div className="guest-managed-name"><span className="guest-group-index">{index + 1}</span><input aria-label={`Invitee name ${index + 1}`} value={group.name} onChange={(event) => updateManagedGuestGroup(group.id, { name: event.target.value })}/></div><label className="guest-managed-count"><span>Guests</span><input aria-label={`Guests included for ${group.name}`} type="number" min="1" step="1" inputMode="numeric" defaultValue={group.count} onBlur={(event) => {const count = Math.max(1, Number.parseInt(event.currentTarget.value, 10) || 1); event.currentTarget.value = String(count); updateManagedGuestCount(group.id, String(count))}}/></label><label className={`guest-rsvp-select rsvp-${group.rsvp ?? 'pending'}`}><span>RSVP</span><select aria-label={`RSVP status for ${group.name}`} value={group.rsvp ?? 'pending'} onChange={(event) => updateManagedGuestGroup(group.id, { rsvp: event.target.value as GuestGroup['rsvp'] })}><option value="pending">Awaiting reply</option><option value="attending">Attending</option><option value="declined">Declined</option></select></label><label className="guest-sent-toggle"><span>Sent</span><input aria-label={`Invitation sent to ${group.name}`} type="checkbox" checked={group.invitationSent ?? false} onChange={(event) => updateManagedGuestGroup(group.id, { invitationSent: event.target.checked })}/></label><label className="guest-channel-select"><span>Sent via</span><select aria-label={`Invitation method for ${group.name}`} value={group.sentVia ?? ''} onChange={(event) => updateManagedGuestGroup(group.id, { sentVia: event.target.value })}><option value="">Choose method</option>{invitationChannels.map((channel) => <option key={channel} value={channel}>{channel}</option>)}</select></label><button className="guest-managed-remove" aria-label={`Remove ${group.name}`} onClick={() => removeGuest(group.name)}><X size={16}/></button></article>)}</div> : <div className="dedicated-empty-state"><Users size={22}/><strong>No invitees yet</strong><span>Add a person or group above to start your guest list.</span></div>}<p className="guest-management-note">Invitation sent and RSVP details are for your own tracking. Guests can reply without creating an account.</p></section></main>}

    {screen === 'seating-management' && managedInvite && <main className="simple-main dedicated-management-page"><button className="back-link" onClick={() => transitionTo('invitation-dashboard')}><ArrowLeft size={14}/> Invitation dashboard</button><div className="dedicated-page-heading"><div><p className="simple-overline">INVITATION MAINTENANCE</p><h1>Seating arrangements</h1><p>{managedInvite.title} · {managedInvite.tables.length} {managedInvite.tables.length === 1 ? 'table' : 'tables'} created.</p></div><button className="simple-secondary" onClick={() => transitionTo('guest-management')}><Users size={14}/> View all guests</button></div><section className="seating-management-panel"><div className="seating-add-toolbar"><div><h2>Your tables</h2><p>Create tables, then add guest groups to each one.</p></div><div className="seating-add-controls"><input id="manage-new-table-name" aria-label="Table name" required value={newTable} onChange={(event) => {setNewTable(event.target.value);clearFieldError('manageNewTable')}} onKeyDown={(event) => {if (event.key === 'Enter') {event.preventDefault();addTable()}}} placeholder="Table name, e.g. Table 1"/><button className="simple-primary" onClick={addTable}><Plus size={14}/> Add table</button></div>{fieldErrors.manageNewTable && <small className="field-error seating-add-error">{fieldErrors.manageNewTable}</small>}</div>{managedInvite.tables.length ? <div className="seating-page-grid">{managedInvite.tables.map((table, index) => <article className="seating-table seating-page-table" key={`${table.name}-${index}`}><div className="seating-table-head"><span className="table-icon">◉</span><strong>{table.name}</strong><span>{peopleSeatedAtTable(managedInvite, table)} people seated</span></div>{table.guests.map((guest) => <div className="seat-person" key={guest}><span>{guest} · {guestGroupsForInvite(managedInvite).find((group) => group.name === guest)?.count ?? 1}</span><button aria-label={`Unassign ${guest}`} onClick={() => unassignGuest(index, guest)}><X size={13}/></button></div>)}<select aria-label={`Add guest to ${table.name}`} defaultValue="" onChange={(event) => {if (event.target.value) assignGuest(index, event.target.value); event.target.value = ''}}><option value="">+ Add a guest to this table</option>{managedInvite.guestNames.filter((guest) => !managedInvite.tables.some((other, otherIndex) => otherIndex !== index && other.guests.includes(guest)) && !table.guests.includes(guest)).map((guest) => <option key={guest} value={guest}>{guest} · {guestGroupsForInvite(managedInvite).find((group) => group.name === guest)?.count ?? 1}</option>)}</select></article>)}</div> : <div className="dedicated-empty-state"><CircleUserRound size={22}/><strong>No tables yet</strong><span>Add a table above to start assigning seats.</span></div>}<div className="people-hint"><Users size={14}/> Add guests to the guest list before assigning seats.</div></section></main>}

    {invitationToPreview && <div className="simple-modal-backdrop invitation-preview-backdrop" onMouseDown={(event) => {if (event.target === event.currentTarget) setPreviewInviteId(null)}}><section className="invitation-preview-modal" role="dialog" aria-modal="true" aria-label={`Preview invitation for ${invitationToPreview.title}`}><button className="people-close" aria-label="Close invitation preview" onClick={() => setPreviewInviteId(null)}><X size={17}/></button><div className="invitation-preview-cover">{invitationToPreview.coverImage && <img src={invitationToPreview.coverImage} alt="Invitation cover" style={{ objectPosition: '50% 50%', transform: coverTransform(invitationToPreview.coverImagePosition) }}/>}<span>{invitationToPreview.type}</span></div><div className="invitation-preview-copy"><p className="simple-overline">YOU’RE INVITED</p><h2>{invitationToPreview.title}</h2><p className="invitation-preview-date"><CalendarDays size={15}/>{invitationToPreview.date}</p><p className="invitation-preview-place"><MapPin size={15}/>{invitationToPreview.place}</p>{invitationToPreview.schedule.length > 0 && <div className="invitation-preview-schedule"><h3>Event-day schedule</h3>{invitationToPreview.schedule.map((item) => <div key={item.id}><time>{item.time || 'Time TBD'}</time><span><strong>{item.title}</strong>{item.place && <small>{item.place}</small>}</span></div>)}</div>}<button className="simple-primary" onClick={() => void downloadInvitationImage(invitationToPreview)}><Download size={14}/> Download invitation image</button></div></section></div>}

    {invitationToDelete && <div className="simple-modal-backdrop delete-confirm-backdrop" onMouseDown={(event) => {if (event.target === event.currentTarget) setDeleteConfirmId(null)}}><section className="delete-confirm-modal" role="alertdialog" aria-modal="true" aria-labelledby="delete-invitation-title"><button className="people-close" aria-label="Close confirmation" onClick={() => setDeleteConfirmId(null)}><X size={17}/></button><span className="delete-confirm-icon"><Trash2 size={19}/></span><p className="simple-overline">DELETE INVITATION</p><h2 id="delete-invitation-title">Delete this invitation?</h2><p className="delete-confirm-copy"><strong>{invitationToDelete.title}</strong> and its guest list, schedule, and seating arrangements will be removed.</p><div className="delete-confirm-actions"><button className="delete-cancel-button" onClick={() => setDeleteConfirmId(null)}>Keep invitation</button><button className="delete-confirm-button" onClick={() => deleteInvitation(invitationToDelete.id)}><Trash2 size={14}/> Delete invitation</button></div></section></div>}
    {coverBeingEdited && <div className="simple-modal-backdrop cover-editor-backdrop" onMouseDown={(event) => {if (event.target === event.currentTarget) setCoverEditorId(null)}}><section className="cover-editor-modal" role="dialog" aria-modal="true" aria-label={`Adjust cover photo for ${coverBeingEdited.title}`}><button className="people-close" aria-label="Close photo editor" onClick={() => setCoverEditorId(null)}><X size={17}/></button><p className="simple-overline">COVER PHOTO</p><h2>{coverBeingEdited.coverImage ? 'Adjust the crop' : 'Choose a cover photo'}</h2><p className="cover-editor-sub">Choose a photo, then adjust how it fits your invitation.</p><div className="cover-crop-preview">{coverBeingEdited.coverImage ? <img src={coverBeingEdited.coverImage} alt="Cover crop preview" style={{ objectPosition: '50% 50%', transform: coverTransform(coverBeingEdited.coverImagePosition) }}/> : <div className="cover-empty-preview"><ImagePlus size={25}/><span>Your cover photo preview</span></div>}</div><div className="cover-editor-actions"><button className="cover-random-control" onClick={() => useRandomCoverPhoto(coverBeingEdited.id)}><Shuffle size={14}/> Random photo</button><label className="cover-modal-upload" htmlFor={`cover-upload-${coverBeingEdited.id}`}><ImagePlus size={14}/>{coverBeingEdited.coverImage ? 'Upload a different photo' : 'Upload a photo'}</label></div>{coverBeingEdited.coverImage && <><label className="cover-range"><span>Zoom <strong>{coverBeingEdited.coverImagePosition?.zoom ?? 100}%</strong></span><input type="range" min="100" max="220" step="1" value={coverBeingEdited.coverImagePosition?.zoom ?? 100} onChange={(event) => updateCoverPosition(coverBeingEdited.id, 'zoom', Number(event.target.value))}/></label><label className="cover-range"><span>Horizontal position</span><input type="range" min="0" max="100" step="1" value={coverBeingEdited.coverImagePosition?.x ?? 50} onChange={(event) => updateCoverPosition(coverBeingEdited.id, 'x', Number(event.target.value))}/></label><label className="cover-range"><span>Vertical position</span><input type="range" min="0" max="100" step="1" value={coverBeingEdited.coverImagePosition?.y ?? 50} onChange={(event) => updateCoverPosition(coverBeingEdited.id, 'y', Number(event.target.value))}/></label></>}<button className="simple-primary cover-editor-done" onClick={() => setCoverEditorId(null)}><Check size={15}/> Done</button></section></div>}
    {screen === 'manage' && managedInvite && manageMode && <div className="simple-modal-backdrop" onMouseDown={(event) => {if (event.target === event.currentTarget) setManageMode(null)}}><section className="people-modal" role="dialog" aria-modal="true" aria-label={manageMode === 'guests' ? 'Manage guest list' : 'Seating arrangements'}><button className="people-close" aria-label="Close" onClick={() => setManageMode(null)}><X size={17}/></button><p className="simple-overline">{managedInvite.title}</p><h2>{manageMode === 'guests' ? 'Who’s on your guest list?' : 'Seating arrangements'}</h2><p className="people-modal-sub">{manageMode === 'guests' ? 'Add people now, and remove anyone whenever you need.' : 'Add tables, then choose where each guest will sit.'}</p>{manageMode === 'guests' ? <><div className="add-person-row guest-add-row"><div className="manage-required-field guest-add-name-field"><input id="manage-new-guest-name" aria-label="Invitee name" required value={newGuest} onChange={e => {setNewGuest(e.target.value);clearFieldError('manageNewGuest')}} onKeyDown={e => {if (e.key === 'Enter') {e.preventDefault();addGuest()}}} placeholder="Name or family name"/>{fieldErrors.manageNewGuest && <small className="field-error">{fieldErrors.manageNewGuest}</small>}</div><label className="guest-add-count"><span>Guests</span><input aria-label="Number of guests included" type="number" min="1" required step="1" inputMode="numeric" value={newGuestCount} onChange={e => setNewGuestCount(e.target.value)}/></label><button className="simple-primary" onClick={addGuest}><Plus size={15}/> Add</button></div><div className="people-list">{guestGroupsForInvite(managedInvite).length ? guestGroupsForInvite(managedInvite).map((group) => <div className="person-row guest-group-person-row" key={group.id}><span className="person-avatar">{group.name.split(' ').map((part) => part[0]).join('').slice(0, 2)}</span><div className="person-group-copy"><strong>{group.name}</strong><small>{group.count} {group.count === 1 ? 'guest' : 'guests'} included</small></div><label className="managed-guest-count"><span>Guests</span><input aria-label={`Number of guests for ${group.name}`} type="number" min="1" required step="1" inputMode="numeric" defaultValue={group.count} onBlur={e => { const count = Math.max(1, Number.parseInt(e.currentTarget.value, 10) || 1); e.currentTarget.value = String(count); updateManagedGuestCount(group.id, String(count)) }}/></label><button aria-label={`Remove ${group.name}`} onClick={() => removeGuest(group.name)}><X size={15}/></button></div>) : <p className="people-empty">No invitees yet. Add a person or family above.</p>}</div><div className="people-hint"><Users size={14}/> Guest counts include everyone in each invitation group.</div></> : <><div className="add-person-row"><div className="manage-required-field"><input id="manage-new-table-name" aria-label="Table name" required value={newTable} onChange={e => {setNewTable(e.target.value);clearFieldError('manageNewTable')}} onKeyDown={e => {if (e.key === 'Enter') {e.preventDefault();addTable()}}} placeholder="Table name, e.g. Table 1"/>{fieldErrors.manageNewTable && <small className="field-error">{fieldErrors.manageNewTable}</small>}</div><button className="simple-primary" onClick={addTable}><Plus size={15}/> Add table</button></div><div className="table-list">{managedInvite.tables.length ? managedInvite.tables.map((table, index) => <div className="seating-table" key={`${table.name}-${index}`}><div className="seating-table-head"><span className="table-icon">◉</span><strong>{table.name}</strong><span>{peopleSeatedAtTable(managedInvite, table)} people seated</span></div>{table.guests.map((guest) => <div className="seat-person" key={guest}><span>{guest} · {guestGroupsForInvite(managedInvite).find((group) => group.name === guest)?.count ?? 1}</span><button aria-label={`Unassign ${guest}`} onClick={() => unassignGuest(index, guest)}><X size={13}/></button></div>)}<select aria-label={`Add guest to ${table.name}`} defaultValue="" onChange={(event) => {if (event.target.value) assignGuest(index, event.target.value); event.target.value = ''}}><option value="">+ Add a guest to this table</option>{managedInvite.guestNames.filter((guest) => !managedInvite.tables.some((other, otherIndex) => otherIndex !== index && other.guests.includes(guest)) && !table.guests.includes(guest)).map((guest) => <option key={guest} value={guest}>{guest} · {guestGroupsForInvite(managedInvite).find((group) => group.name === guest)?.count ?? 1}</option>)}</select></div>) : <p className="people-empty">No tables yet. Add your first table above.</p>}</div><div className="people-hint"><Users size={14}/> Add guests to the guest list before assigning seats.</div></>}</section></div>}

    <footer className="simple-footer"><span>invitecyprus <span>Made for life’s lovely moments.</span></span><span>Need a hand? &nbsp; Privacy</span></footer>
    {notice && <div className="simple-toast"><Check size={16}/>{notice}<button aria-label="Dismiss" onClick={() => setNotice('')}><X size={14}/></button></div>}
  </div>
}
export default App
