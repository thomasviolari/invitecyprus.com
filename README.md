# invitecyprus.com
## Local preview password

The Vite development preview asks for a password before showing the app. Create a `.env.local` file in the project root and add:

```env
INVITECYPRUS_ACCESS_PASSWORD=your-password-here
```

Restart `npm run dev` after changing the password. `.env.local` is ignored by Git; `.env.example` is a safe template.

For production on Vercel, add `INVITECYPRUS_ACCESS_PASSWORD` as an environment variable for the Production environment in Project Settings → Environment Variables, then redeploy. The root Vercel middleware protects the app and issues an eight-hour signed, HTTP-only session cookie after a correct password. Shared `/invite/{token}` pages and app assets are available without that password so invited guests can reply. The environment variable is never bundled into the browser app.

## Event schedule locations

Hosts can add a place or paste a Google Maps link into each event-day schedule moment. Guests can open those links in Google Maps without a Maps API key.

## User accounts

Invitecyprus uses Firebase Authentication for sender accounts. Copy the `VITE_FIREBASE_*` settings from a Firebase web app into `.env.local` for local development and into the Vercel project's environment variables for deployment. The browser Firebase API key is a public app identifier; protect the project with Firebase Authentication configuration and authorized domains, and do not use a server credential as a `VITE_` value.

In Firebase Console, create a project and web app, then open Authentication → Sign-in method and enable Email/Password, Google, and Facebook as needed. For Facebook, add the Facebook App ID and App Secret in Firebase's provider settings and register Firebase's OAuth redirect URL in the Facebook app. Add your local and production site domains to Firebase Authentication's authorized domains. Email sign-up asks the user to verify their email before their first login; password reset emails use Firebase's configured email templates.

Create a Cloud Firestore database and enable Firebase Storage in the same project. Set all `VITE_FIREBASE_*` values, including `VITE_FIREBASE_STORAGE_BUCKET`, in `.env.local` and in Vercel's environment variables. Deploy the included Firestore and Storage rules with the Firebase CLI:

```sh
firebase deploy --project YOUR_FIREBASE_PROJECT_ID --only firestore:rules,storage
```

Each account's invitations are stored at `users/{uid}/invitations/{invitationId}` and update across that account's open sessions in real time. Cover photos are stored under `users/{uid}/covers/`. Owner invitation documents and cover files are private to their account. When a host shares an invitation, its public page is stored at `publicInvitations/{token}`; guests can open it without an account, select their invitation group, and submit or update an RSVP. Public seating details become readable on the event date. The local preview remains browser-only and cannot publish guest links.

After logging in, the host sees a one-time-per-session notice that the service is free for a limited time; payment options are shown as unavailable. From an invitation dashboard, **Send invitation link** creates a share URL and opens a message editor. The copy action includes the optional message, event link, and a note telling guests where to find seating arrangements on the wedding day.

When a signed-in account first connects to Firestore, invitations previously saved in that browser under its Firebase UID are copied to the cloud. Keep that browser data until the migration has completed successfully; if the account already has cloud invitations, cloud data is used and local data is not merged automatically.
