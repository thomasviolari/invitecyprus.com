import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App'
import PublicInvitationPage from './PublicInvitationPage'
import { LanguageProvider } from './i18n'
import './styles.css'
import './mobile.css'

const publicInvite = window.location.pathname.match(/^\/invite\/([A-Za-z0-9-]+)\/?$/)
ReactDOM.createRoot(document.getElementById('root')!).render(<React.StrictMode><LanguageProvider>{publicInvite ? <PublicInvitationPage token={publicInvite[1]}/> : <App/>}</LanguageProvider></React.StrictMode>)
