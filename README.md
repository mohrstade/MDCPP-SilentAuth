# MDCPP SilentAUTH

This folder is the clean upload package for the SilentAuth variant of the Team Startpage prototype.

Base references:

- MDCPP samples repo:
  `https://github.com/microsoft/MDCPP-samples`
- Present Live sample README:
  `https://github.com/microsoft/MDCPP-samples/blob/main/present-live-integration-app-sample/README.md`

## What this package is

This package keeps the same main Team Startpage functionality as the earlier prototype, but changes the authentication approach to prefer MDCPP SilentAuth.

Main idea:

- same Team Startpage UI
- same recent-files-first home view
- same folder browser in the left panel
- same preview and file actions
- different authentication strategy: SilentAuth-first

## What SilentAuth is doing here

This version uses MDCPP `trySilentAuth` as the first authentication attempt.

In simple words:

- the app first tries to sign the user in quietly
- if Microsoft already knows the user, the app can continue without immediately showing the normal login flow
- if silent auth does not work, the app falls back to the regular interactive Microsoft sign-in popup

So this variant is:

- SilentAuth first
- normal login second

## Where SilentAuth is explicitly used

Main file:

- `present-live-integration-app-sample/src/aadauth.ts`

Explicit places:

1. `trySilentAuth` is imported from `@microsoft/document-collaboration-sdk`
2. `attemptMdcppSilentAuth(loginHint)` calls:

```ts
trySilentAuth(aadClientId, loginHint, "Web")
```

3. `initializeAuth()` uses SilentAuth when there is no redirect response but a saved login hint may exist
4. `signIn()` tries SilentAuth first before falling back to `loginPopup`
5. `getGraphToken()` retries SilentAuth when token acquisition hits an interaction-required case

## What changed because of SilentAuth

Compared to the non-SilentAuth Team Startpage version:

- the app now prefers a quiet sign-in attempt first
- saved login hint is used to guide that attempt
- cached account reuse is tied more closely to silent-auth success
- normal popup login remains as the fallback path

What did NOT change:

- the Team Startpage UI
- file listing behavior
- recent files home page
- folder browsing behavior
- in-window preview behavior
- edit handoff to Microsoft 365

## Main app location

The actual app is inside:

`present-live-integration-app-sample`

Important source files:

- `present-live-integration-app-sample/src/desktop.ts`
- `present-live-integration-app-sample/src/aadauth.ts`
- `present-live-integration-app-sample/src/spoPicker.ts`
- `present-live-integration-app-sample/src/static/desktop.html`

## What is included

Included because they are needed:

- source code in `src/`
- static assets in `images/`
- project config files like `package.json`, `tsconfig.json`, `webpack.config.js`, `webpack.development.js`
- lock/config files like `.gitignore`, `.npmrc`, `.prettierrc`, `yarn.lock`

## What is intentionally NOT included

Removed from this handoff package:

- `.env`
- `node_modules/`
- `dist/`
- `lib/`
- local log files
- extra working notes
- duplicate or nonessential outer documentation
- support / conduct / security / license metadata files

## Why `.env` is not included

`.env` contains local environment-specific values such as the Entra app ID.
That file should be created by the receiving team in their own environment.

Required local value:

```text
ENTRA_APPID=[your Entra client ID]
```

## How to run locally

1. Open the folder `present-live-integration-app-sample`
2. Create a `.env` file
3. Add:

```text
ENTRA_APPID=[your Entra client ID]
```

4. Install dependencies:

```text
npm install --legacy-peer-deps
```

5. Start the app:

```text
npm run start
```

6. Open:

```text
http://localhost:8080/desktop.html
```

## Current product behavior

- Home shows recent files only
- Folders are browsed from the `Folders` section in the left panel
- Files open in a focused in-window preview
- Full editing should be done through `Edit in Microsoft 365`
- There is also an `Edit in app` attempt for supported Office files, but Microsoft may still require an embedded sign-in depending on tenant/session/browser behavior

## Notes for reviewers

This package is meant to be the safe upload version of the SilentAuth build.
It is intentionally smaller and cleaner than the full working folder.