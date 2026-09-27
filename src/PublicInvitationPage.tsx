import { useEffect, useState, type FormEvent } from 'react'
import { CalendarDays, Check, Clock3, MapPin, Users } from 'lucide-react'
import { collection, doc, getDoc, onSnapshot, serverTimestamp, setDoc } from 'firebase/firestore'
import { firebaseDb } from './firebase'

type GuestGroup = { id: string; name: string; count: number; rsvp?: 'pending' | 'attending' | 'declined'; guestsComing?: number }
type ScheduleItem = { id: number; title: string; time: string; place: string }
type PublicInvite = {
  title: string
  type: string
  date: string
  dateInput: string
  place: string
  coverImage: string
  guestNames: string[]
  guestGroups: GuestGroup[]
  schedule: ScheduleItem[]
  shareMessage: string
}
type Response = { rsvp: 'pending' | 'attending' | 'declined'; guestsComing: number }
type SeatingPlan = { tables: { name: string; guests: string[] }[] }

const mapsHref = (place: string) => `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(place)}`

export default function PublicInvitationPage({ token }: { token: string }) {
  const [invite, setInvite] = useState<PublicInvite | null>(null)
  const [responses, setResponses] = useState<Record<string, Response>>({})
  const [seating, setSeating] = useState<SeatingPlan | null>(null)
  const [selectedGroupId, setSelectedGroupId] = useState('')
  const [rsvp, setRsvp] = useState<'attending' | 'declined'>('attending')
  const [guestsComing, setGuestsComing] = useState(1)
  const [loading, setLoading] = useState(true)
  const [seatingLoading, setSeatingLoading] = useState(false)
  const [saving, setSaving] = useState(false)
  const [error, setError] = useState('')
  const [saved, setSaved] = useState(false)

  useEffect(() => {
    if (!firebaseDb) {
      setError('This invitation page is not available right now.')
      setLoading(false)
      return
    }
    return onSnapshot(doc(firebaseDb, 'publicInvitations', token), (snapshot) => {
      if (!snapshot.exists()) {
        setInvite(null)
        setError('This invitation link is no longer available.')
      } else {
        setInvite(snapshot.data() as PublicInvite)
        setError('')
      }
      setLoading(false)
    }, () => {
      setError('Could not load this invitation. Check your connection and try again.')
      setLoading(false)
    })
  }, [token])

  useEffect(() => {
    if (!firebaseDb || !invite) return
    return onSnapshot(collection(firebaseDb, 'publicInvitations', token, 'responses'), (snapshot) => {
      const next: Record<string, Response> = {}
      snapshot.docs.forEach((response) => { next[response.id] = response.data() as Response })
      setResponses(next)
    }, () => setError('Could not load RSVP updates.'))
  }, [invite, token])

  const eventDay = invite?.dateInput ? new Date(`${invite.dateInput}T00:00:00`) : null
  const today = new Date()
  today.setHours(0, 0, 0, 0)
  const seatingAvailable = Boolean(eventDay && today >= eventDay && invite?.guestGroups.some((group) => group.rsvp === 'attending' || responses[group.id]?.rsvp === 'attending'))

  useEffect(() => {
    if (!firebaseDb || !seatingAvailable) {
      setSeating(null)
      return
    }
    let active = true
    setSeatingLoading(true)
    getDoc(doc(firebaseDb, 'publicInvitations', token, 'seating', 'plan'))
      .then((snapshot) => { if (active && snapshot.exists()) setSeating(snapshot.data() as SeatingPlan) })
      .catch(() => { if (active) setSeating(null) })
      .finally(() => { if (active) setSeatingLoading(false) })
    return () => { active = false }
  }, [seatingAvailable, token])

  const chooseGroup = (groupId: string) => {
    setSelectedGroupId(groupId)
    setSaved(false)
    const group = invite?.guestGroups.find((candidate) => candidate.id === groupId)
    const response = responses[groupId]
    setRsvp((response?.rsvp ?? group?.rsvp) === 'declined' ? 'declined' : 'attending')
    setGuestsComing(Math.max(1, response?.guestsComing ?? group?.guestsComing ?? group?.count ?? 1))
  }

  const submitRsvp = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault()
    if (!firebaseDb || !invite || !selectedGroupId) return
    setSaving(true)
    setError('')
    try {
      const group = invite.guestGroups.find((candidate) => candidate.id === selectedGroupId)
      if (!group) throw new Error('Choose the invitation for your person or family.')
      await setDoc(doc(firebaseDb, 'publicInvitations', token, 'responses', group.id), {
        rsvp,
        guestsComing: rsvp === 'attending' ? Math.min(group.count, Math.max(1, guestsComing)) : 0,
        respondedAt: serverTimestamp(),
      })
      setSaved(true)
    } catch {
      setError('Could not save your reply. Please try again or contact your host.')
    } finally {
      setSaving(false)
    }
  }

  if (loading) return <main className="public-invite-page"><section className="public-invite-state"><p className="simple-overline">INVITECYPRUS INVITATION</p><h1>Opening your invitation</h1><p>One moment, please.</p></section></main>
  if (error && !invite) return <main className="public-invite-page"><section className="public-invite-state"><p className="simple-overline">INVITECYPRUS INVITATION</p><h1>Invitation unavailable</h1><p>{error}</p></section></main>
  if (!invite) return null

  const selectedGroup = invite.guestGroups.find((group) => group.id === selectedGroupId)

  return <main className="public-invite-page">
    <header className="public-invite-header"><a className="public-invite-brand" href="/" aria-label="Invitecyprus home"><span className="simple-mark"><i/><i/><i/><i/></span>invitecyprus</a><span className="public-invite-label">YOU’RE INVITED</span></header>
    <article className="public-invite-card">
      {invite.coverImage && <div className="public-invite-cover"><img src={invite.coverImage} alt="Celebration cover"/></div>}
      <div className="public-invite-content">
        <p className="simple-overline">{invite.type.toUpperCase()}</p>
        <h1>{invite.title}</h1>
        {invite.shareMessage && <p className="public-host-message">{invite.shareMessage}</p>}
        <div className="public-invite-details"><p><CalendarDays size={18}/><span>{invite.date}</span></p>{invite.place && <p><MapPin size={18}/><span>{invite.place}</span><a href={mapsHref(invite.place)} target="_blank" rel="noreferrer">Map</a></p>}</div>
        {invite.schedule.length > 0 && <section className="public-schedule"><h2><Clock3 size={17}/> Event-day schedule</h2>{invite.schedule.map((item) => <div key={item.id}><time>{item.time}</time><span><strong>{item.title}</strong>{item.place && <small>{item.place}</small>}</span></div>)}</section>}
        <section className="public-rsvp"><p className="simple-overline">PLEASE REPLY</p><h2>Can you make it?</h2><p>Select your name or family to send your reply to the host.</p><form onSubmit={(event) => void submitRsvp(event)}>
          <label className="public-field"><span>Your invitation</span><select required value={selectedGroupId} onChange={(event) => chooseGroup(event.target.value)}><option value="">Choose your name or family</option>{invite.guestGroups.map((group) => <option key={group.id} value={group.id}>{group.name} · up to {group.count}</option>)}</select></label>
          <fieldset className="public-rsvp-choice"><legend>Your reply</legend><label><input type="radio" name="rsvp" checked={rsvp === 'attending'} onChange={() => setRsvp('attending')}/><span>Joyfully accepts</span></label><label><input type="radio" name="rsvp" checked={rsvp === 'declined'} onChange={() => setRsvp('declined')}/><span>Regretfully declines</span></label></fieldset>
          {rsvp === 'attending' && selectedGroup && <label className="public-field"><span>Guests coming</span><select value={guestsComing} onChange={(event) => setGuestsComing(Number(event.target.value))}>{Array.from({ length: selectedGroup.count }, (_, index) => index + 1).map((count) => <option key={count} value={count}>{count} {count === 1 ? 'guest' : 'guests'}</option>)}</select></label>}
          <button className="simple-primary public-rsvp-submit" type="submit" disabled={!selectedGroup || saving}>{saving ? 'Saving your reply…' : saved ? <><Check size={17}/> Reply saved</> : 'Send my reply'}</button>
          {saved && <p className="public-confirmation" role="status">Thank you. Your reply has been sent to the host.</p>}
          {error && <p className="public-error" role="alert">{error}</p>}
        </form></section>
        <p className="public-seating-note"><Users size={16}/> If seating arrangements are available, you’ll find them at this link on the wedding day.</p>
        {seatingLoading && <p className="public-seating-loading">Checking for seating arrangements…</p>}
        {seating && seating.tables.length > 0 && seatingAvailable && <section className="public-seating"><h2>Seating arrangements</h2>{seating.tables.map((table, index) => <article key={`${table.name}-${index}`}><strong>{table.name}</strong><ul>{table.guests.map((name) => <li key={name}>{name}</li>)}</ul></article>)}</section>}
      </div>
    </article>
    <footer className="public-invite-footer">Made for life’s lovely moments · invitecyprus</footer>
  </main>
}
