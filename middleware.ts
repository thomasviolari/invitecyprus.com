import { next } from '@vercel/functions'

declare const process: { env: Record<string, string | undefined> }

const COOKIE_NAME = 'invitecyprus_preview'
const SESSION_LENGTH_MS = 8 * 60 * 60 * 1000
const ACCESS_PATH = '/__invitecyprus/access'

const json = (body: object, status = 200, headers?: HeadersInit) =>
  new Response(JSON.stringify(body), {
    status,
    headers: { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store', ...Object.fromEntries(new Headers(headers)) },
  })

async function hmac(secret: string, value: string): Promise<Uint8Array> {
  const key = await crypto.subtle.importKey('raw', new TextEncoder().encode(secret), { name: 'HMAC', hash: 'SHA-256' }, false, ['sign'])
  return new Uint8Array(await crypto.subtle.sign('HMAC', key, new TextEncoder().encode(value)))
}

function toBase64Url(bytes: Uint8Array): string {
  let binary = ''
  for (const byte of bytes) binary += String.fromCharCode(byte)
  return btoa(binary).replaceAll('+', '-').replaceAll('/', '_').replaceAll('=', '')
}

function fromBase64Url(value: string): Uint8Array | null {
  try {
    const base64 = value.replaceAll('-', '+').replaceAll('_', '/')
    const binary = atob(base64 + '='.repeat((4 - (base64.length % 4)) % 4))
    return Uint8Array.from(binary, (character) => character.charCodeAt(0))
  } catch { return null }
}

function constantTimeEqual(left: Uint8Array, right: Uint8Array): boolean {
  if (left.length !== right.length) return false
  let difference = 0
  for (let index = 0; index < left.length; index += 1) difference |= left[index] ^ right[index]
  return difference === 0
}

function getCookie(request: Request, name: string): string | null {
  const pair = request.headers.get('cookie')?.split(';').map((part) => part.trim()).find((part) => part.startsWith(`${name}=`))
  return pair ? decodeURIComponent(pair.slice(name.length + 1)) : null
}

async function hasValidSession(request: Request, password: string): Promise<boolean> {
  const token = getCookie(request, COOKIE_NAME)
  if (!token) return false
  const separator = token.indexOf('.')
  if (separator < 1) return false
  const expiresAt = Number(token.slice(0, separator))
  const signature = fromBase64Url(token.slice(separator + 1))
  if (!Number.isSafeInteger(expiresAt) || expiresAt <= Date.now() || !signature) return false
  return constantTimeEqual(signature, await hmac(password, `invitecyprus-preview:${expiresAt}`))
}

function passwordPage(configured: boolean): Response {
  const title = configured ? 'Enter the password' : 'Preview access is not configured'
  const description = configured
    ? 'This invitecyprus preview is private. Enter the password to continue.'
    : 'Set INVITECYPRUS_ACCESS_PASSWORD in your Vercel project environment variables, then redeploy.'
  const form = configured ? `<form id="access-form"><label for="password">Preview password</label><input id="password" name="password" type="password" autocomplete="current-password" required autofocus><button type="submit">Open invitecyprus <span aria-hidden="true">→</span></button><p class="error" id="error" role="alert" aria-live="polite"></p></form>` : ''
  const html = `<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta name="robots" content="noindex,nofollow"><title>${title} · invitecyprus</title><style>
    *{box-sizing:border-box}body{margin:0;min-height:100vh;display:grid;place-items:center;padding:24px;background:radial-gradient(ellipse at 12% 10%,#e8eddf 0,transparent 35%),radial-gradient(ellipse at 95% 90%,#f1e5d9 0,transparent 29%),#faf9f5;color:#34443a;font-family:Inter,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif}.card{width:min(420px,100%);padding:38px;background:#fff;border:1px solid #ebe9e1;border-radius:15px;box-shadow:0 18px 55px #28382c13}.brand{display:flex;align-items:center;gap:10px;margin-bottom:40px;font-family:Georgia,serif;font-size:25px;letter-spacing:-.5px}.mark{width:24px;height:24px;display:grid;grid-template-columns:1fr 1fr;gap:3px;transform:rotate(-9deg)}.mark i{display:block;border-radius:7px 7px 3px 7px;background:#6e816b}.mark i:nth-child(2),.mark i:nth-child(3){background:#d3a889}.mark i:nth-child(4){background:#889780}.eyebrow{margin:0 0 8px;color:#899080;font-size:9px;font-weight:700;letter-spacing:1.45px}.lock{display:grid;place-items:center;width:42px;height:42px;margin-bottom:18px;border-radius:11px;background:#eff1e9;color:#6e8065;font-size:20px}.card h1{margin:0 0 9px;font:400 31px/1.12 Georgia,serif;letter-spacing:-.3px}.description{margin:0 0 23px;color:#828a7f;font-size:12px;line-height:1.65}label{display:block;margin-bottom:7px;color:#596459;font-size:11px;font-weight:600}input{width:100%;height:45px;margin-bottom:14px;padding:0 12px;border:1px solid #e4e5dc;border-radius:6px;font:13px inherit}input:focus{outline:3px solid #dfe6da;border-color:#a4b19b}button{width:100%;height:45px;display:flex;align-items:center;justify-content:space-between;padding:0 14px;border:0;border-radius:6px;background:#405b46;color:#fff;font-size:12px;font-weight:600;cursor:pointer}button:hover{background:#304939}.error{min-height:17px;margin:11px 0 0;color:#a4574d;font-size:11px}@media(max-width:480px){.card{padding:29px 24px}.brand{margin-bottom:32px}.card h1{font-size:28px}}
    </style></head><body><main class="card"><div class="brand"><span class="mark"><i></i><i></i><i></i><i></i></span>invitecyprus</div><div class="lock" aria-hidden="true">⌑</div><p class="eyebrow">PRIVATE PREVIEW</p><h1>${title}</h1><p class="description">${description}</p>${form}</main>${configured ? `<script>const form=document.getElementById('access-form');form.addEventListener('submit',async(event)=>{event.preventDefault();const button=form.querySelector('button');const error=document.getElementById('error');button.disabled=true;button.textContent='Checking…';error.textContent='';try{const response=await fetch('${ACCESS_PATH}',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({password:document.getElementById('password').value})});if(response.ok){location.replace('/');return}error.textContent=response.status===401?'That password doesn’t match. Try again.':'Could not check the password.'}catch{error.textContent='Could not connect. Please try again.'}button.disabled=false;button.innerHTML='Open invitecyprus <span aria-hidden="true">→</span>'});</script>` : ''}</body></html>`
  return new Response(html, { status: configured ? 200 : 503, headers: { 'Content-Type': 'text/html; charset=utf-8', 'Cache-Control': 'no-store', 'X-Robots-Tag': 'noindex, nofollow', 'X-Content-Type-Options': 'nosniff', 'X-Frame-Options': 'DENY' } })
}

export default async function middleware(request: Request): Promise<Response> {
  const password = process.env.INVITECYPRUS_ACCESS_PASSWORD ?? ''
  const url = new URL(request.url)

  if (url.pathname === ACCESS_PATH) {
    if (!password) return json({ configured: false, authorized: false }, 503)
    if (request.method === 'GET') return json({ configured: true, authorized: await hasValidSession(request, password) })
    if (request.method !== 'POST') return json({ configured: true, authorized: false }, 405)

    let submitted = ''
    try { submitted = String((await request.json() as { password?: string }).password ?? '') } catch { return json({ configured: true, authorized: false }, 400) }
    const [expected, candidate] = await Promise.all([hmac(password, password), hmac(password, submitted)])
    if (!constantTimeEqual(expected, candidate)) return json({ configured: true, authorized: false }, 401)

    const expiresAt = Date.now() + SESSION_LENGTH_MS
    const signature = toBase64Url(await hmac(password, `invitecyprus-preview:${expiresAt}`))
    const secure = url.protocol === 'https:' ? '; Secure' : ''
    return json({ configured: true, authorized: true }, 200, { 'Set-Cookie': `${COOKIE_NAME}=${expiresAt}.${signature}; Path=/; HttpOnly; SameSite=Strict; Max-Age=${SESSION_LENGTH_MS / 1000}${secure}` })
  }

  if (!password) return passwordPage(false)
  if (await hasValidSession(request, password)) return next()
  if (request.headers.get('accept')?.includes('text/html')) return passwordPage(true)
  return new Response('Password required', { status: 401, headers: { 'Cache-Control': 'no-store', 'X-Robots-Tag': 'noindex, nofollow' } })
}

export const config = { matcher: '/:path*' }
