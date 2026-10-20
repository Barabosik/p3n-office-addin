<p align="center">
  <img src="src/assets/icon-128.png" width="88" alt="">
</p>

<h1 align="center">p3n for Word and Excel</h1>

<p align="center">
  Select text in Word, or cells in Excel. See if it reads as AI-written, or rewrite it in place.
</p>

<p align="center">
  AppSource (not listed yet) · <a href="https://p3n.si">p3n.si</a>
</p>

<p align="center">
  <a href="https://github.com/Barabosik/p3n-office-addin/actions/workflows/ci.yml"><img src="https://github.com/Barabosik/p3n-office-addin/actions/workflows/ci.yml/badge.svg" alt="CI"></a>
</p>

![Check for AI in Word](docs/word-check.png)

## What it does

- **Check for AI.** A score, a one-line reason, and the sentences that give it away.
- **Rewrite.** A version that keeps your facts and drops the filler. **Replace** puts it back where it came from. In Word that's your selection, and each paragraph keeps its style. In Excel it's the cell. Ctrl+Z undoes it.
- Open it from **Check for AI** on the Home tab, or **Check with p3n** in the right-click menu.

![Rewrite and Replace in Excel](docs/excel-replaced.png)

## Free, or with a key

You get 10 checks or rewrites an hour without an account, up to 5,000 characters each. For more, make an API key at [p3n.si](https://p3n.si/dashboard/keys) and paste it into the pane (Add API key, top right).

## Privacy

p3n sends the text you chose, only when you click, and only to `api.p3n.si`. It keeps no history and has no analytics. Your API key stays on your device and is never saved in a document. The details are in [PRIVACY.md](PRIVACY.md).

## Install

Until the AppSource listing is live, add it from its manifest, [manifest.xml](https://barabosik.github.io/p3n-office-addin/manifest.xml):

- **Word or Excel on the web:** Home › Add-ins › More Add-ins › My Add-ins › Upload My Add-in, then pick the file.
- **Windows:** put the file in a shared folder, add that folder as a trusted catalog, then find p3n under Shared Folder. [Microsoft's steps](https://learn.microsoft.com/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins).
- **Mac:** copy the file to `~/Library/Containers/com.microsoft.Word/Data/Documents/wef` (for Excel, `com.microsoft.Excel`), then restart the app.

The pane loads from GitHub Pages, so it updates by itself.

## Known gaps

Tested in Word and Excel for Windows (Microsoft 365). Mac and the web run the same code but haven't been tested. In Excel, Replace works on one cell at a time, and never over a formula.

## Develop

Tests and running it locally are in [CONTRIBUTING.md](CONTRIBUTING.md).

## Licence

MIT
