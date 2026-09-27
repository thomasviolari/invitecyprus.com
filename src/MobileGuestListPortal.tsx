import { useEffect, useState } from 'react'
import { createPortal } from 'react-dom'
import MobileGuestCard from './MobileGuestCard'

type GuestGroup = {
  id: string
  name: string
  count: number
  rsvp?: 'pending' | 'attending' | 'declined'
  invitationSent?: boolean
  sentVia?: string
}

type Props = {
  active: boolean
  groups: GuestGroup[]
  channels: string[]
  onSave: (id: string, update: Partial<GuestGroup>) => void
  onRemove: (name: string) => void
  onShare: (group: GuestGroup) => void
}

export default function MobileGuestListPortal({ active, groups, channels, onSave, onRemove, onShare }: Props) {
  const [target, setTarget] = useState<HTMLElement | null>(null)

  useEffect(() => {
    const nextTarget = active ? document.querySelector<HTMLElement>('.guest-management-list') : null
    setTarget((current) => current === nextTarget ? current : nextTarget)
  }, [active, groups])

  if (!active || !target) return null

  return createPortal(
    <div className="guest-mobile-list">
        {groups.map((group, index) => <MobileGuestCard key={group.id} group={group} index={index} channels={channels} onSave={onSave} onRemove={onRemove} onShare={onShare}/>)}
    </div>,
    target,
  )
}
