# invitecyprus.com
## Local preview password

The Vite development preview asks for a password before showing the app. Create a `.env.local` file in the project root and add:

```env
INVITECYPRUS_ACCESS_PASSWORD=your-password-here
```

Restart `npm run dev` after changing the password. `.env.local` is ignored by Git; `.env.example` is a safe template.

For production on Vercel, add `INVITECYPRUS_ACCESS_PASSWORD` as an environment variable for the Production environment in Project Settings → Environment Variables, then redeploy. The root Vercel middleware protects the HTML and static app assets and issues an eight-hour signed, HTTP-only session cookie after a correct password. The environment variable is never bundled into the browser app.

## Google Maps previews

Hosts can paste a Google Maps link for the venue and for each schedule location. Guests can open these links in Google Maps without an API key. To show an embedded map preview in the editor, create a browser key for the [Google Maps Embed API](https://developers.google.com/maps/documentation/embed/quickstart), restrict it to your website and that API, then set `VITE_GOOGLE_MAPS_EMBED_KEY` in `.env.local` locally and in your Vercel project settings. This key is public in the browser, so use website/API restrictions; never put a private server key here.

## User accounts

Invitecyprus uses Firebase Authentication for sender accounts. Copy the `VITE_FIREBASE_*` settings from a Firebase web app into `.env.local` for local development and into the Vercel project's environment variables for deployment. The browser Firebase API key is a public app identifier; protect the project with Firebase Authentication configuration and authorized domains, and do not use a server credential as a `VITE_` value.

In Firebase Console, create a project and web app, then open Authentication → Sign-in method and enable Email/Password, Google, and Facebook as needed. For Facebook, add the Facebook App ID and App Secret in Firebase's provider settings and register Firebase's OAuth redirect URL in the Facebook app. Add your local and production site domains to Firebase Authentication's authorized domains. Email sign-up asks the user to verify their email before their first login; password reset emails use Firebase's configured email templates.

The current invitation and draft data is saved in browser storage under the signed-in user's Firebase UID, so accounts on the same browser stay separate. It is not yet synced between devices; that requires connecting the app to a cloud database such as Cloud Firestore.
