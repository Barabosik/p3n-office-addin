# AppSource listing

What goes into the Partner Center forms for the Office add-in. The package is `manifest.xml`; the pane it points at is on GitHub Pages.

Status: not submitted yet.

## Summary (up to 100 characters)

```
Check text in Word and Excel for AI writing, and rewrite it so it reads as human-written.
```

## Description

```
Select text in Word, or cells in Excel, and click Check for AI on the Home tab. p3n tells you whether the text reads as AI-written and which sentences give it away. Or it rewrites the text so it reads as human, and puts the new version back.

Check for AI: a score, a short reason, and the sentences that read as AI.

Rewrite: a new version that keeps your facts and drops the filler. In Word, Replace puts it over your selection and keeps each paragraph's style. In Excel, it goes back into the cell. Ctrl+Z undoes it.

Free to try: without an account you get 10 checks or rewrites an hour, up to 5,000 characters each. For more, add an API key from p3n.si.

Privacy: p3n sends only the text you chose, only when you click, and only to api.p3n.si. It keeps no history and has no analytics. Your API key stays on your device and is never saved in a document.

Open source: https://github.com/Barabosik/p3n-office-addin
```

## Fields

| Field | Value |
| --- | --- |
| Name | p3n |
| Categories | Productivity, Utilities |
| Search keywords | AI detector, humanize, writing |
| Privacy policy | https://github.com/Barabosik/p3n-office-addin/blob/main/PRIVACY.md (p3n's own: https://p3n.si/privacy) |
| Terms of use | https://p3n.si/terms |
| Support | https://github.com/Barabosik/p3n-office-addin/issues |
| Website | https://p3n.si |
| Logo | `store/logo-300.png` (300×300) |
| Screenshots | 1366×768, from `docs/` (to make) |
| Additional purchases | No: it works without an account; an API key from p3n.si raises the limits. |

## Notes for certification

```
No account is needed. In Word: select a paragraph, then Home > Check for AI, and click Check for AI or Rewrite in the pane. In Excel: select a cell with text and do the same; Replace writes the rewrite back into that cell.

Calls without an API key are free, 10 an hour per IP, up to 5,000 characters each. If the testers share an IP and hit that limit, the pane says so and when more calls come.

The add-in contacts only https://api.p3n.si, and only when the user clicks Check for AI or Rewrite. The pane's files are served from https://barabosik.github.io/p3n-office-addin/. Source: https://github.com/Barabosik/p3n-office-addin.
```
