import { defineConfig, loadEnv } from 'vite'
import react from '@vitejs/plugin-react'

function previewPassword(password: string) {
  const sessions = new Map<string, number>()

  return {
    name: 'invitecyprus-preview-password',
    configureServer(server: import('vite').ViteDevServer) {
      server.middlewares.use((req, res, next) => {
        const request = req as unknown as { url?: string; method?: string; headers: { cookie?: string }; on: (event: string, listener: (chunk?: string) => void) => void }
        const response = res as unknown as { statusCode: number; setHeader: (name: string, value: string) => void; end: (body?: string) => void }
        const path = (request.url ?? '/').split('?')[0]
        if (path !== '/__invitecyprus/access') return next()
        response.setHeader('Content-Type', 'application/json; charset=utf-8')
        if (!password) {
          response.statusCode = 503
          response.end(JSON.stringify({ configured: false, authorized: false }))
          return
        }

        const cookies = request.headers.cookie ?? ''
        const token = cookies.split(';').map((cookie) => cookie.trim()).find((cookie) => cookie.startsWith('invitecyprus_preview='))?.split('=')[1]
        const expires = token ? sessions.get(token) : undefined
        if (request.method === 'GET') {
          if (token && expires && expires > Date.now()) {
            response.end(JSON.stringify({ configured: true, authorized: true }))
          } else {
            if (token) sessions.delete(token)
            response.end(JSON.stringify({ configured: true, authorized: false }))
          }
          return
        }

        if (request.method !== 'POST') {
          response.statusCode = 405
          response.end(JSON.stringify({ configured: true, authorized: false }))
          return
        }

        let body = ''
        request.on('data', (chunk) => { body += chunk ?? '' })
        request.on('end', () => {
          let submitted = ''
          try { submitted = String((JSON.parse(body) as { password?: string }).password ?? '') } catch { /* Invalid JSON is treated as a wrong password. */ }
          if (submitted !== password) {
            response.statusCode = 401
            response.end(JSON.stringify({ configured: true, authorized: false }))
            return
          }
          const sessionToken = `${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}-${Math.random().toString(36).slice(2)}`
          sessions.set(sessionToken, Date.now() + 8 * 60 * 60 * 1000)
          response.setHeader('Set-Cookie', `invitecyprus_preview=${sessionToken}; Path=/; HttpOnly; SameSite=Strict; Max-Age=28800`)
          response.end(JSON.stringify({ configured: true, authorized: true }))
        })
      })
    },
  }
}

export default defineConfig(({ mode }) => {
  const env = loadEnv(mode, '.', '')
  return { plugins: [react(), previewPassword(env.INVITECYPRUS_ACCESS_PASSWORD ?? '')] }
})
