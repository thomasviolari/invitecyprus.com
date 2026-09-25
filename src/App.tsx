import { useEffect, useState, type FormEvent } from 'react'
import { ArrowLeft, ArrowRight, CalendarDays, Check, ChevronRight, CircleUserRound, ImagePlus, LockKeyhole, MapPin, Plus, Users, X } from 'lucide-react'

type Invite = { id: number; title: string; type: string; date: string; dateInput?: string; timeInput?: string; place: string; guests: number; guestNames: string[]; tables: { name: string; guests: string[] }[] }
type Screen = 'login' | 'home' | 'create' | 'manage'
type SavedUi = { screen?: Screen; step?: number; kind?: string; title?: string; date?: string; time?: string; place?: string; guestText?: string; guestCount?: string; editingId?: number | null; manageId?: number | null; manageMode?: 'guests' | 'seating' | null; newGuest?: string; newTable?: string }
const readSavedUi = (): SavedUi => {
  try { return JSON.parse(localStorage.getItem('invitecyprus-demo-ui') ?? '{}') as SavedUi } catch { return {} }
}
const starterInvites: Invite[] = [
  { id: 1, title: 'Olivia & James’ wedding', type: 'Wedding', date: 'Saturday, 14 June 2025 · 17:00', dateInput: '2025-06-14', timeInput: '17:00', place: 'Villa Keryneia, Paphos', guests: 4, guestNames: ['Sofia Antoniou', 'Daniel Costa', 'Maya Georgiou', 'Leo Nicolaou'], tables: [{ name: 'Table 1', guests: ['Sofia Antoniou', 'Daniel Costa'] }, { name: 'Table 2', guests: ['Maya Georgiou', 'Leo Nicolaou'] }] },
  { id: 2, title: 'Theo’s first birthday', type: 'Birthday', date: 'Sunday, 22 June 2025 · 15:00', dateInput: '2025-06-22', timeInput: '15:00', place: 'Our garden, Limassol', guests: 0, guestNames: [], tables: [] },
]
const occasions = ['Wedding', 'Birthday', 'Baptism', 'Company event', 'Dinner party', 'Something else']

function App() {
  const savedUi = readSavedUi()
  const [access, setAccess] = useState<'checking' | 'locked' | 'open' | 'setup'>('checking')
  const [accessPassword, setAccessPassword] = useState('')
  const [accessError, setAccessError] = useState('')
  const [screen, setScreen] = useState<Screen>(() => {
    try {
      if (localStorage.getItem('invitecyprus-demo-session') !== 'signed-in') return 'login'
      return savedUi.screen && savedUi.screen !== 'login' ? savedUi.screen : 'home'
    } catch { return 'login' }
  })
  const [invites, setInvites] = useState<Invite[]>(() => {
    try { const saved = localStorage.getItem('invitecyprus-demo-invitations'); return saved ? JSON.parse(saved) as Invite[] : starterInvites } catch { return starterInvites }
  })
  const [step, setStep] = useState(savedUi.step ?? 1)
  const [kind, setKind] = useState(savedUi.kind ?? 'Wedding')
  const [title, setTitle] = useState(savedUi.title ?? '')
  const [date, setDate] = useState(savedUi.date ?? '')
  const [time, setTime] = useState(savedUi.time ?? '')
  const [place, setPlace] = useState(savedUi.place ?? '')
  const [guestText, setGuestText] = useState(savedUi.guestText ?? '')
  const [guestCount, setGuestCount] = useState(savedUi.guestCount ?? '')
  const [editingId, setEditingId] = useState<number | null>(savedUi.editingId ?? null)
  const [manageId, setManageId] = useState<number | null>(savedUi.manageId ?? null)
  const [manageMode, setManageMode] = useState<'guests' | 'seating' | null>(savedUi.manageMode ?? null)
  const [newGuest, setNewGuest] = useState(savedUi.newGuest ?? '')
  const [newTable, setNewTable] = useState(savedUi.newTable ?? '')
  const [notice, setNotice] = useState('')

  useEffect(() => {
    fetch('/__invitecyprus/access', { credentials: 'same-origin' })
      .then(async (response) => ({ response, data: await response.json() as { configured?: boolean; authorized?: boolean } }))
      .then(({ response, data }) => setAccess(!data.configured ? 'setup' : response.ok && data.authorized ? 'open' : 'locked'))
      .catch(() => setAccess('setup'))
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
    try {
      if (screen === 'login') localStorage.removeItem('invitecyprus-demo-session')
      else localStorage.setItem('invitecyprus-demo-session', 'signed-in')
    } catch { /* Storage may be disabled; keep the app usable for this session. */ }
  }, [screen])

  useEffect(() => {
    try {
      if (screen === 'login') localStorage.removeItem('invitecyprus-demo-ui')
      else localStorage.setItem('invitecyprus-demo-ui', JSON.stringify({ screen, step, kind, title, date, time, place, guestText, guestCount, editingId, manageId, manageMode, newGuest, newTable } satisfies SavedUi))
    } catch { /* Storage may be disabled; keep the app usable for this session. */ }
  }, [screen, step, kind, title, date, time, place, guestText, guestCount, editingId, manageId, manageMode, newGuest, newTable])

  useEffect(() => {
    try { localStorage.setItem('invitecyprus-demo-invitations', JSON.stringify(invites)) } catch { /* Storage may be full or disabled. */ }
  }, [invites])

  const create = () => {
    const names = guestText.split(/[\n,;]/).map((g) => g.trim()).filter(Boolean)
    const formattedDate = date ? `${new Date(`${date}T12:00:00`).toLocaleDateString('en-GB', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' })}${time ? ` · ${time}` : ''}` : 'Date to be decided'
    const totalGuests = guestCount.trim() ? Math.max(names.length, Number.parseInt(guestCount, 10) || 0) : names.length
    setInvites((all) => editingId ? all.map((invite) => invite.id === editingId ? { ...invite, title: title || `${kind} invitation`, type: kind, date: date ? formattedDate : invite.date, dateInput: date || invite.dateInput, timeInput: time || invite.timeInput, place: place || invite.place, guests: totalGuests, guestNames: names, tables: invite.tables.map((table) => ({ ...table, guests: table.guests.filter((guest) => names.includes(guest)) })) } : invite) : [{ id: Date.now(), title: title || `${kind} invitation`, type: kind, date: formattedDate, dateInput: date, timeInput: time, place: place || 'Place to be decided', guests: totalGuests, guestNames: names, tables: [] }, ...all])
    setScreen('manage')
    setStep(1); setKind('Wedding'); setTitle(''); setDate(''); setTime(''); setPlace(''); setGuestText(''); setGuestCount(''); setEditingId(null)
    setNotice('Invitation saved. You can keep editing it whenever you’re ready.')
    window.setTimeout(() => setNotice(''), 5000)
  }

  const back = () => {
    if (screen === 'create' && step > 1) setStep(step - 1)
    else setScreen('home')
  }

  const editInvite = (invite: Invite) => {
    setEditingId(invite.id); setKind(invite.type); setTitle(invite.title); setPlace(invite.place === 'Place to be decided' ? '' : invite.place); setGuestText(invite.guestNames.join('\n')); setGuestCount(String(invite.guests)); setDate(invite.dateInput ?? ''); setTime(invite.timeInput ?? ''); setStep(2); setScreen('create')
  }

  const addGuest = () => {
    const name = newGuest.trim()
    if (!name || manageId === null) return
    setInvites((all) => all.map((invite) => invite.id === manageId && !invite.guestNames.includes(name) ? { ...invite, guests: invite.guests + 1, guestNames: [...invite.guestNames, name] } : invite))
    setNewGuest('')
  }

  const removeGuest = (name: string) => {
    if (manageId === null) return
    setInvites((all) => all.map((invite) => invite.id === manageId ? { ...invite, guests: Math.max(0, invite.guests - 1), guestNames: invite.guestNames.filter((guest) => guest !== name), tables: invite.tables.map((table) => ({ ...table, guests: table.guests.filter((guest) => guest !== name) })) } : invite))
  }

  const addTable = () => {
    if (!newTable.trim() || manageId === null) return
    setInvites((all) => all.map((invite) => invite.id === manageId ? { ...invite, tables: [...invite.tables, { name: newTable.trim(), guests: [] }] } : invite))
    setNewTable('')
  }

  const assignGuest = (tableIndex: number, guestName: string) => {
    if (manageId === null) return
    setInvites((all) => all.map((invite) => invite.id === manageId ? { ...invite, tables: invite.tables.map((table, index) => ({ ...table, guests: index === tableIndex ? [...new Set([...table.guests, guestName])] : table.guests.filter((guest) => guest !== guestName) })) } : invite))
  }

  const unassignGuest = (tableIndex: number, guestName: string) => {
    if (manageId === null) return
    setInvites((all) => all.map((invite) => invite.id === manageId ? { ...invite, tables: invite.tables.map((table, index) => ({ ...table, guests: index === tableIndex ? table.guests.filter((guest) => guest !== guestName) : table.guests })) } : invite))
  }

  const managedInvite = invites.find((invite) => invite.id === manageId)

  if (access !== 'open') return <main className="access-screen"><div className="access-card"><div className="access-brand"><span className="simple-mark"><i/><i/><i/><i/></span>invitecyprus</div><span className="login-icon"><LockKeyhole size={19}/></span><p className="simple-overline">PRIVATE PREVIEW</p><h1>{access === 'checking' ? 'Checking access…' : access === 'setup' ? 'Set up preview access' : 'Enter the password'}</h1>{access === 'setup' ? <p className="access-copy">Create a <code>.env.local</code> file in the project folder and add <code>INVITECYPRUS_ACCESS_PASSWORD=your-password</code>. Restart the dev server to apply it.</p> : access === 'checking' ? <p className="access-copy">One moment while we check this preview.</p> : <form onSubmit={unlockPreview}><p className="access-copy">Enter the preview password to continue.</p><label htmlFor="preview-password">Password</label><input id="preview-password" type="password" autoFocus autoComplete="current-password" value={accessPassword} onChange={(e) => setAccessPassword(e.target.value)} placeholder="Enter preview password" required/><button className="simple-primary full-button" type="submit">Open invitecyprus <ArrowRight size={15}/></button>{accessError && <p className="access-error">{accessError}</p>}</form>}</div></main>

  return <div className="simple-app">
    <header className="simple-header"><button className="simple-brand" onClick={() => setScreen(screen === 'login' ? 'login' : 'home')}><span className="simple-mark"><i/><i/><i/><i/></span>invitecyprus</button>{screen !== 'login' && <div className="account-chip"><span className="account-initial">E</span><span>Emma Wilson</span><ChevronRight size={14}/></div>}</header>

    {screen === 'login' && <main className="login-layout"><section className="login-welcome"><div className="welcome-art"><div className="art-photo"></div><div className="art-note"><span>✳</span><strong>A little invite.<br/>A lovely big moment.</strong></div><div className="art-circle"></div><div className="art-caption">INVITECYPRUS · CELEBRATE TOGETHER</div></div><div className="login-welcome-copy"><p className="simple-overline">FOR ALL THE MOMENTS THAT MATTER</p><h1>Bring your people<br/>a little closer.</h1><p>Beautiful invitations and easy RSVPs, all in one place.</p></div></section><section className="login-panel"><div className="login-form"><span className="login-icon"><CircleUserRound size={20}/></span><p className="simple-overline">YOUR INVITECYPRUS ACCOUNT</p><h2>Welcome back</h2><p className="login-sub">Log in to create and manage your invitations.</p><label htmlFor="email">Email address</label><input id="email" type="email" placeholder="you@example.com"/><label htmlFor="password">Password</label><input id="password" type="password" placeholder="Enter your password"/><button className="forgot-link" onClick={() => setNotice('Password reset link requested.')}>Forgot password?</button><button className="simple-primary full-button" onClick={() => setScreen('home')}>Log in <ArrowRight size={16}/></button><div className="login-divider"><span></span>or<span></span></div><button className="google-button" onClick={() => setScreen('home')}><span>G</span> Continue with Google</button><p className="signup-line">New to invitecyprus? <button onClick={() => setScreen('home')}>Create an account</button></p><p className="privacy-line"><LockKeyhole size={12}/> Your guests won’t need an account to RSVP.</p></div></section></main>}

    {screen === 'home' && <main className="simple-main home-screen"><button className="back-to-login" onClick={() => setScreen('login')}><ArrowLeft size={14}/> Log out</button><div className="simple-heading"><p className="simple-overline">YOUR CELEBRATIONS, IN ONE PLACE</p><h1>What would you like to do?</h1><p>It only takes a few minutes to bring everyone together.</p></div><div className="choice-grid"><button className="choice-card choice-create" onClick={() => {setEditingId(null);setScreen('create');setStep(1)}}><span className="choice-icon"><Plus size={20}/></span><span className="choice-label">START SOMETHING NEW</span><strong>What do you want<br/>to invite people to?</strong><span className="choice-description">A wedding, birthday, baptism, work event, or anything worth celebrating.</span><span className="choice-action">Create an invitation <ArrowRight size={16}/></span><span className="choice-flower">✳</span></button><button className="choice-card choice-manage" onClick={() => setScreen('manage')}><span className="choice-icon"><Users size={19}/></span><span className="choice-label">PICK UP WHERE YOU LEFT OFF</span><strong>Manage invitations</strong><span className="choice-description">Check guest replies, update the details, or send your invitation link.</span><span className="choice-action">View my invitations <ArrowRight size={16}/></span><span className="manage-preview"><span>O&J</span><span>✳</span><span>+{invites.length}</span></span></button></div><p className="home-reassurance"><LockKeyhole size={13}/> Guests can open and RSVP to your invitation without logging in.</p></main>}

    {screen === 'create' && <main className="simple-main create-screen"><button className="back-link" onClick={back}><ArrowLeft size={14}/>{step === 1 ? 'Back to choices' : 'Previous step'}</button><div className="wizard-wrap"><div className="wizard-top"><p className="simple-overline">LET’S GET THIS CELEBRATION STARTED</p><div className="wizard-progress"><span className={step >= 1 ? 'current' : ''}></span><span className={step >= 2 ? 'current' : ''}></span><span className={step >= 3 ? 'current' : ''}></span></div><small>STEP {step} OF 3</small></div>
      {step === 1 && <section className="wizard-step"><h1>What are we celebrating?</h1><p>Choose the kind of invitation you’d like to create.</p><div className="occasion-grid">{occasions.map((occasion) => <button key={occasion} className={`occasion-option ${kind === occasion ? 'picked' : ''}`} onClick={() => setKind(occasion)}><span className="occasion-symbol">{occasion === 'Wedding' ? '♡' : occasion === 'Birthday' ? '✳' : occasion === 'Baptism' ? '⌁' : occasion === 'Company event' ? '▧' : occasion === 'Dinner party' ? '◌' : '✦'}</span>{occasion}{kind === occasion && <Check size={15}/>}</button>)}</div><button className="simple-primary wizard-next" onClick={() => setStep(2)}>Next <ArrowRight size={16}/></button></section>}
      {step === 2 && <section className="wizard-step"><h1>{editingId ? 'Edit invitation details.' : 'Add the important details.'}</h1><p>You can change these details later.</p><label htmlFor="invite-title">Give your invitation a name</label><input id="invite-title" value={title} onChange={e => setTitle(e.target.value)} placeholder={kind === 'Wedding' ? 'e.g. Olivia & James’ wedding' : `e.g. My ${kind.toLowerCase()}`}/><div className="form-pair"><div><label htmlFor="invite-date">Date</label><input id="invite-date" type="date" value={date} onChange={e => setDate(e.target.value)}/></div><div><label htmlFor="invite-time">Time</label><input id="invite-time" type="time" value={time} onChange={e => setTime(e.target.value)}/></div></div><label htmlFor="invite-place">Place</label><div className="input-with-icon"><MapPin size={15}/><input id="invite-place" value={place} onChange={e => setPlace(e.target.value)} placeholder="Venue name or address"/></div><button className="simple-primary wizard-next" onClick={() => setStep(3)}>Next <ArrowRight size={16}/></button></section>}
      {step === 3 && <section className="wizard-step"><h1>Who should we invite?</h1><p>Add names one per line, or enter a guest count.</p><label htmlFor="invite-guests">Guest names <span className="optional">Optional · one person per line</span></label><textarea id="invite-guests" rows={4} value={guestText} onChange={e => setGuestText(e.target.value)} placeholder={'Alex Morgan\nSam Taylor'}/><div className="guest-count-field"><label htmlFor="guest-count">Total number of guests</label><input id="guest-count" type="number" min="0" step="1" inputMode="numeric" value={guestCount} onChange={e => setGuestCount(e.target.value)} placeholder={String(guestText.split(/[\n,;]/).filter((name) => name.trim()).length)}/><small>For example, include partners or children in the total.</small></div><div className="private-note"><Users size={14}/> You can add more people later, or remove someone from the guest list at any time.</div><button className="simple-primary wizard-next" onClick={create}>{editingId ? 'Save changes' : 'Create invitation'} <ArrowRight size={16}/></button></section>}
      <div className="wizard-foot"><span><LockKeyhole size={12}/> Your invitation starts as a draft.</span><span><i>1</i> Choose <i>2</i> Details <i>3</i> Guests</span></div></div></main>}

    {screen === 'manage' && <main className="simple-main manage-screen"><button className="back-link" onClick={() => setScreen('home')}><ArrowLeft size={14}/> Back to choices</button><div className="manage-heading"><div><p className="simple-overline">YOUR CELEBRATIONS</p><h1>Manage invitations</h1><p>Keep all your event details and replies together.</p></div><button className="simple-primary" onClick={() => {setEditingId(null);setScreen('create');setStep(1)}}><Plus size={16}/> New invitation</button></div>{invites.length ? <div className="managed-list">{invites.map((invite, i) => <article className="managed-card" key={invite.id}><div className={`managed-cover cover-${i % 3}`}><span>{invite.type === 'Wedding' ? 'O&J' : invite.type === 'Birthday' ? '✳' : invite.type.substring(0,2).toUpperCase()}</span></div><div className="managed-content"><span className="draft-label"><i/> DRAFT</span><h2>{invite.title}</h2><p><CalendarDays size={14}/>{invite.date}</p><p><MapPin size={14}/>{invite.place}</p><div className="managed-bottom"><span><Users size={14}/>{invite.guests ? `${invite.guests} people invited` : 'No guests added yet'}</span><button onClick={() => {setManageId(invite.id);setManageMode('guests')}}>Guests <ArrowRight size={14}/></button></div><div className="invite-actions"><button onClick={() => editInvite(invite)}>Edit invitation</button><button onClick={() => {setManageId(invite.id);setManageMode('guests')}}>Guest list</button><button onClick={() => {setManageId(invite.id);setManageMode('seating')}}>Seating arrangements</button></div></div></article>)}</div> : <div className="empty-invites"><span><ImagePlus size={22}/></span><h2>Your invitations will live here.</h2><p>Create your first invitation and we’ll keep everything organised for you.</p><button className="simple-primary" onClick={() => {setScreen('create');setStep(1)}}><Plus size={16}/> Create invitation</button></div>}<div className="manage-tip"><span className="tip-star">✳</span><p><strong>A little heads-up</strong><br/>Your guests can RSVP from their invitation link. You can add and manage guests at any time.</p></div></main>}

    {managedInvite && manageMode && <div className="simple-modal-backdrop" onMouseDown={(event) => {if (event.target === event.currentTarget) setManageMode(null)}}><section className="people-modal" role="dialog" aria-modal="true" aria-label={manageMode === 'guests' ? 'Manage guest list' : 'Seating arrangements'}><button className="people-close" aria-label="Close" onClick={() => setManageMode(null)}><X size={17}/></button><p className="simple-overline">{managedInvite.title}</p><h2>{manageMode === 'guests' ? 'Who’s on your guest list?' : 'Seating arrangements'}</h2><p className="people-modal-sub">{manageMode === 'guests' ? 'Add people now, and remove anyone whenever you need.' : 'Add tables, then choose where each guest will sit.'}</p>{manageMode === 'guests' ? <><div className="add-person-row"><input aria-label="Guest name" value={newGuest} onChange={e => setNewGuest(e.target.value)} onKeyDown={e => {if (e.key === 'Enter') {e.preventDefault();addGuest()}}} placeholder="Enter a person’s name"/><button className="simple-primary" onClick={addGuest}><Plus size={15}/> Add</button></div><div className="people-list">{managedInvite.guestNames.length ? managedInvite.guestNames.map((guest) => <div className="person-row" key={guest}><span className="person-avatar">{guest.split(' ').map((part) => part[0]).join('').slice(0, 2)}</span><strong>{guest}</strong><button aria-label={`Remove ${guest}`} onClick={() => removeGuest(guest)}><X size={15}/></button></div>) : <p className="people-empty">No guests yet. Add someone above to get started.</p>}</div><div className="people-hint"><Users size={14}/> You can add more people later or remove them from this list.</div></> : <><div className="add-person-row"><input aria-label="Table name" value={newTable} onChange={e => setNewTable(e.target.value)} onKeyDown={e => {if (e.key === 'Enter') {e.preventDefault();addTable()}}} placeholder="Table name, e.g. Table 1"/><button className="simple-primary" onClick={addTable}><Plus size={15}/> Add table</button></div><div className="table-list">{managedInvite.tables.length ? managedInvite.tables.map((table, index) => <div className="seating-table" key={`${table.name}-${index}`}><div className="seating-table-head"><span className="table-icon">◉</span><strong>{table.name}</strong><span>{table.guests.length} seated</span></div>{table.guests.map((guest) => <div className="seat-person" key={guest}><span>{guest}</span><button aria-label={`Unassign ${guest}`} onClick={() => unassignGuest(index, guest)}><X size={13}/></button></div>)}<select aria-label={`Add guest to ${table.name}`} defaultValue="" onChange={(event) => {if (event.target.value) assignGuest(index, event.target.value); event.target.value = ''}}><option value="">+ Add a guest to this table</option>{managedInvite.guestNames.filter((guest) => !managedInvite.tables.some((other, otherIndex) => otherIndex !== index && other.guests.includes(guest)) && !table.guests.includes(guest)).map((guest) => <option key={guest} value={guest}>{guest}</option>)}</select></div>) : <p className="people-empty">No tables yet. Add your first table above.</p>}</div><div className="people-hint"><Users size={14}/> Add guests to the guest list before assigning seats.</div></>}</section></div>}

    <footer className="simple-footer"><span>invitecyprus <span>Made for life’s lovely moments.</span></span><span>Need a hand? &nbsp; Privacy</span></footer>
    {notice && <div className="simple-toast"><Check size={16}/>{notice}<button aria-label="Dismiss" onClick={() => setNotice('')}><X size={14}/></button></div>}
  </div>
}
export default App
