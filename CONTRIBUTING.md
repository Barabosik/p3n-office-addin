# Contributing

Plain JavaScript with no bundler. Office loads the task pane from `src/` as it is. `dev/` holds the tests and tools, and isn't published.

## Layout

- `manifest.xml`: the add-in as published, with the pane on GitHub Pages. `manifest.dev.xml` is the same add-in served from `https://localhost:3000`; `npm run manifest` writes it from `manifest.xml`, and the tests fail if the two drift apart.
- `src/taskpane.*`: the pane. It calls the API itself.
- `src/lib/word.js`: reads the selection, or the whole document, and keeps a tracked range to it. Replace puts the rewrite back paragraph by paragraph, so each keeps its style, in one batch that Ctrl+Z undoes. It refuses if the text changed meanwhile, or over a table or picture.
- `src/lib/excel.js`: reads the text in the selected cells, one paragraph per cell. Replace writes into the cell only when the text came from one cell that holds plain text, not a formula.
- `src/lib/api.js`: the API client, the same file as in the browser extension.
- `src/lib/store.js`: the API key, in the pane's `localStorage`.
- `src/commands.html`: the function file the manifest needs. The buttons only open the pane.

## Test

```bash
npm install
npm run e2e
```

This loads the pane in headless Chrome with a stand-in for Office.js (`dev/office-stub.js`) that models a Word document and an Excel workbook, and answers API calls with canned results. It covers Check, Rewrite and Replace, the cases where Replace refuses, the limits, the API key, the manifests, and the pane's light and dark (`theme`: Office's dark themes on its own pane colour, with every text checked for WCAG AA contrast; `switch`: Office, Light and Dark at the pane's foot and in the settings, no header button).

- `LIVE=1 npm run e2e` uses the live API without a key: one Check and one Rewrite, 2 of the 10 free calls an hour.
- `ONLY=replace,excelReplace` runs some of the scenarios. `CHROME=` points at another Chrome.

In real Word or Excel for Windows, with the app closed first:

```bash
npx office-addin-dev-certs install
MOCK=1 node dev/live.mjs word
MOCK=1 node dev/live.mjs excel
```

It sideloads the add-in, writes a test document or sheet over COM, drives the pane through WebView2's debugging port, and checks Replace and undo. It never takes focus from what you're doing. Without `MOCK=1` it uses two free calls. Screenshots go to `dev/out/`. `MANIFEST=manifest.xml` runs the published add-in, with its pane from GitHub Pages.

## Run it

Once, first: `npx office-addin-dev-certs install` makes a certificate for `localhost`, and Windows asks you to trust it. Then `npm start` opens Word with the add-in sideloaded from `manifest.dev.xml` and serves `src/` on `https://localhost:3000`. `npm run start:excel` does the same in Excel, and `npm run stop` removes it.

## Release

1. Raise the version in `package.json`, in `manifest.xml` (`<Version>` has four parts) and in `src/lib/api.js` (`VERSION`, sent in `X-P3N-Client`), then run `npm run manifest` and `npm run validate`.
2. Push to `main`. CI runs the tests, and the Pages workflow deploys `src/` and `manifest.xml` to `https://nikita-andreiev.com/p3n-office-addin/` once they pass.

Office loads the pane from Pages each time it opens, so a deploy reaches everyone. A change to the manifest reaches people only when they install it again or the store copy is updated. The add-in's Id can't change once it's on AppSource. The store copy is in `store/listing.md`, and the submission steps are in `store/partner-center-steps.md`. `python dev/icons.py` builds the icons.

Store screenshots: `SHOTS=1 MANIFEST=manifest.xml node dev/live.mjs excel` takes the Excel ones with one live call, and the same with `word` takes a Check and a Rewrite in Word with two. Then `python dev/store-shots.py` writes the 1366×768 images to `store/`. `RECORD=1` saves the live answers to `dev/out/`, and `REPLAY=1` reuses them for the same text instead of calling the API.
