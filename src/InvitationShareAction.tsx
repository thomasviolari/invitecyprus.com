import { useEffect, useState } from 'react'
import { createPortal } from 'react-dom'
import { Share2 } from 'lucide-react'

type Props = { active: boolean; onClick: () => void }

export default function InvitationShareAction({ active, onClick }: Props) {
  const [target, setTarget] = useState<HTMLElement | null>(null)

  useEffect(() => {
    const nextTarget = active ? document.querySelector<HTMLElement>('.dashboard-hero-actions') : null
    setTarget((current) => current === nextTarget ? current : nextTarget)
  }, [active])

  if (!active || !target) return null
  return createPortal(<button type="button" className="simple-secondary invitation-share-action" onClick={onClick}><Share2 size={15}/> Invitation link</button>, target)
}
