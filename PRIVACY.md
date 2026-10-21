# Privacy

The p3n add-in for Word and Excel sends text to p3n only when you ask it to.

- **What it sends.** The text you selected, or typed into the pane, when you click Check for AI or Rewrite. If you added an API key, the key goes with it so p3n can bill your account, and the pane uses it to show your balance. Each call also names the add-in and its version (for example `p3n-word/1.1.0`), so your usage page can show which tool made it.
- **Where it goes.** Only to `api.p3n.si`, over HTTPS.
- **What it doesn't do.** While the pane is open it looks at your selection to count its words for the Use selection button. That count stays in the pane. It keeps no history, and has no analytics or ads.
- **What stays on your device.** Your API key, in the add-in's own storage. It's never written into a document or workbook, so it doesn't travel with a file you share. Remove key in the pane deletes it.
- **Where the pane comes from.** Office loads the pane's files from GitHub Pages (`barabosik.github.io`) and Office.js from Microsoft (`appsforoffice.microsoft.com`), like any web page. Those requests carry no document text.

What p3n does with the text it receives:
- It keeps the text for up to 7 days, to return results and debug failures, then deletes it.
- It sends the text through OpenRouter to the AI model providers that run the check or rewrite.
- It logs the time and IP address of each call. That's how the free limit works.
- It doesn't sell data or use your text to train models.

The full policy is at [p3n.si/privacy](https://p3n.si/privacy). Questions go to [p3n.si/contact](https://p3n.si/contact).
