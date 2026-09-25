# invitecyprus.com
## Local preview password

The Vite development preview asks for a password before showing the app. Create a `.env.local` file in the project root and add:

```env
INVITECYPRUS_ACCESS_PASSWORD=your-password-here
```

Restart `npm run dev` after changing the password. `.env.local` is ignored by Git; `.env.example` is a safe template.

For production on Vercel, add `INVITECYPRUS_ACCESS_PASSWORD` as an environment variable for the Production environment in Project Settings → Environment Variables, then redeploy. The root Vercel middleware protects the HTML and static app assets and issues an eight-hour signed, HTTP-only session cookie after a correct password. The environment variable is never bundled into the browser app.
