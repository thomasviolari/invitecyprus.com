import { useRef, useState, type FormEvent, type TouchEvent } from 'react'
import { Pencil, Save, Share2, Trash2, X } from 'lucide-react'
import { useLanguage } from './i18n'

type GuestGroup = {
  id: string
  name: string
  count: number
  rsvp?: 'pending' | 'attending' | 'declined'
  invitationSent?: boolean
  sentVia?: string
}

type Props = {
  group: GuestGroup
  index: number
  channels: string[]
  onSave: (id: string, update: Partial<GuestGroup>) => void
  onRemove: (name: string) => void
  onShare: (group: GuestGroup) => void
}

const replyLabel = (rsvp: GuestGroup['rsvp']) => {
  if (rsvp === 'attending') return 'Attending'
  if (rsvp === 'declined') return 'Declined'
  return 'Awaiting reply'
}

export default function MobileGuestCard({ group, index, channels, onSave, onRemove, onShare }: Props) {
  const { t } = useLanguage()
  const [editing, setEditing] = useState(false)
  const [swiped, setSwiped] = useState(false)
  const [draft, setDraft] = useState<GuestGroup>(group)
  const [nameError, setNameError] = useState('')
  const touchStart = useRef<{ x: number; y: number } | null>(null)
  const suppressClick = useRef(false)

  const closeEditor = () => {
    setEditing(false)
    setNameError('')
  }

  const openEditor = () => {
    setDraft({ ...group })
    setNameError('')
    setEditing(true)
  }

  const handleTouchStart = (event: TouchEvent<HTMLButtonElement>) => {
    const touch = event.touches[0]
    touchStart.current = { x: touch.clientX, y: touch.clientY }
    suppressClick.current = false
  }

  const handleTouchEnd = (event: TouchEvent<HTMLButtonElement>) => {
    if (!touchStart.current) return
    const touch = event.changedTouches[0]
    const deltaX = touch.clientX - touchStart.current.x
    const deltaY = touch.clientY - touchStart.current.y
    if (Math.abs(deltaX) > 48 && Math.abs(deltaX) > Math.abs(deltaY) * 1.2) {
      setSwiped(deltaX < 0)
      suppressClick.current = true
    }
    touchStart.current = null
  }

  const handleSummaryClick = () => {
    if (suppressClick.current) {
      suppressClick.current = false
      return
    }
    if (swiped) {
      setSwiped(false)
      return
    }
    openEditor()
  }

  const save = (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault()
    const name = draft.name.trim()
    if (!name) {
      setNameError('Enter a name for this invitee.')
      return
    }
    onSave(group.id, {
      name,
      count: Math.max(1, Number(draft.count) || 1),
      rsvp: draft.rsvp ?? 'pending',
      invitationSent: draft.invitationSent ?? false,
      sentVia: draft.sentVia ?? '',
    })
    closeEditor()
  }

  return <div className={`guest-mobile-card ${swiped ? 'is-swiped' : ''}`}>
    <button type="button" className="guest-mobile-summary" aria-label={`Edit ${group.name || `invitee ${index + 1}`}`} onClick={handleSummaryClick} onTouchStart={handleTouchStart} onTouchEnd={handleTouchEnd}>
      <span className="guest-group-index">{index + 1}</span>
      <span className="guest-mobile-summary-copy">
        <strong>{group.name || `Invitee ${index + 1}`}</strong>
        <span>{group.count} {t(group.count === 1 ? 'guest' : 'guests')} · {t(replyLabel(group.rsvp))}</span>
        <small>{group.invitationSent ? group.sentVia ? `${t('Sent via')} ${group.sentVia}` : t('Invitation sent') : t('Not sent yet')}</small>
      </span>
      <span className="guest-mobile-edit"><Pencil size={16}/> {t('Edit')}</span>
    </button>
    <button type="button" className="guest-mobile-remove" aria-label={`Remove ${group.name}`} onClick={() => onRemove(group.name)}><Trash2 size={17}/></button>
    {editing && <div className="simple-modal-backdrop mobile-guest-edit-backdrop" onMouseDown={(event) => { if (event.target === event.currentTarget) closeEditor() }}>
      <section className="mobile-guest-edit-modal" role="dialog" aria-modal="true" aria-labelledby="mobile-guest-edit-title">
        <button type="button" className="people-close" aria-label="Close guest editor" onClick={closeEditor}><X size={18}/></button>
        <p className="simple-overline">GUEST DETAILS</p>
        <h2 id="mobile-guest-edit-title">{t('Edit invitee')}</h2>
        <p className="mobile-guest-edit-sub">{t('Update this person’s invitation and RSVP details.')}</p>
        <form onSubmit={save}>
          <label className="mobile-guest-edit-field"><span>{t('Name')}</span><input autoFocus required value={draft.name} onChange={(event) => { setDraft({ ...draft, name: event.target.value }); setNameError('') }} placeholder="Person or family name"/></label>
          {nameError && <small className="field-error">{nameError}</small>}
          <div className="mobile-guest-edit-pair">
            <label className="mobile-guest-edit-field"><span>{t('Guests included')}</span><input type="number" min="1" step="1" inputMode="numeric" required value={draft.count} onChange={(event) => setDraft({ ...draft, count: Math.max(1, Number.parseInt(event.target.value, 10) || 1) })}/></label>
            <label className="mobile-guest-edit-field"><span>{t('RSVP')}</span><select value={draft.rsvp ?? 'pending'} onChange={(event) => setDraft({ ...draft, rsvp: event.target.value as GuestGroup['rsvp'] })}><option value="pending">{t('Awaiting reply')}</option><option value="attending">{t('Attending')}</option><option value="declined">{t('Declined')}</option></select></label>
          </div>
          <label className="mobile-guest-edit-sent"><input type="checkbox" checked={draft.invitationSent ?? false} onChange={(event) => setDraft({ ...draft, invitationSent: event.target.checked })}/><span>{t('Invitation sent')}</span></label>
          <label className="mobile-guest-edit-field"><span>{t('Sent via')}</span><select value={draft.sentVia ?? ''} onChange={(event) => setDraft({ ...draft, sentVia: event.target.value })}><option value="">{t('Choose method')}</option>{channels.map((channel) => <option key={channel} value={channel}>{channel}</option>)}</select></label>
          <div className="mobile-guest-edit-actions"><button type="button" className="simple-secondary mobile-guest-share" onClick={() => { onShare({ ...draft, name: draft.name.trim() }); closeEditor() }}><Share2 size={15}/> {t('Share')}</button><button type="button" className="simple-secondary" onClick={closeEditor}>{t('Cancel')}</button><button type="submit" className="simple-primary"><Save size={16}/> {t('Save changes')}</button></div>
        </form>
      </section>
    </div>}
  </div>
}
