# Submitting to AppSource

**Later.** For now the add-in installs from its manifest (option C below). Listing it on AppSource needs a Microsoft work account and a company profile first, so option A or B.

How to list p3n for Word and Excel in Microsoft Marketplace (AppSource) through Partner Center. **[user]** marks a step only the account holder can do: signing in, account details and legal choices. **[lead]** marks the rest. The text to paste is in [listing.md](listing.md). Microsoft says a review takes up to four weeks and often more than one round.

Checked against Microsoft Learn on 2026-10-20 (sources at the end).

## 0. The account (decide first)

Partner Center lists Office add-ins only through the **Microsoft 365 and Copilot program**, and enrolling needs:

- a **work or school account** (Microsoft Entra ID). A personal Microsoft account (Outlook.com, Hotmail, Live) can't enroll. Microsoft's Q&A confirmed this to an individual developer in January 2026.
- a **company** profile: legal business name, address and primary contact, plus the authority to sign for it. Microsoft verifies these, and the form may ask for payment details even for a free add-in.
- a publisher name that matches the manifest's `ProviderName`, now "Nikita Andreiev".

So the user's personal Microsoft account can't publish it. The options:

- **A.** The user makes a work account in a new Entra tenant (for example through a free Azure sign-up) and enrolls himself as the business. Microsoft doesn't say whether it verifies a sole trader. An enrollment abandoned halfway can lock the account ("Access restricted") until Partner Center support fixes it, so finish it in one go.
- **B.** Publish under a company's existing Partner Center account, such as ZDResearch's, which Abbas offered earlier. `ProviderName` then changes to that company's name: a manifest change and a version bump.
- **C.** Skip AppSource for now. People install from the manifest (README, Install). An organisation's admin can deploy it to everyone in the Microsoft 365 admin center (Settings › Integrated apps › Upload custom apps) with `manifest.xml`. No Partner Center account is needed for this.

Steps 1 to 8 assume A or B.

## 1. Enroll

1. **[user]** Open https://partner.microsoft.com/dashboard/account/v3/enrollment/introduction/office and sign in with the work account.
2. **[user]** Fill in the publisher profile: company, publisher and contact details.
3. **[user]** Accept the Microsoft Office Agreement, and the Microsoft AI Cloud Partner Program Agreement if it's shown.
4. **[user]** Confirm the emails Microsoft sends, and wait until the account is verified.

## 2. Create the offer

1. **[lead]** Partner Center › **Marketplace offers** › the **Microsoft 365 and Copilot** tab › **+ New offer** › **Office add-in**.
2. **[lead]** Name: `p3n` (it must match the manifest's `DisplayName`). Click **Check availability**, keep the enrolled publisher (it can't be changed later), then **Create**.

## 3. Product setup

**[lead]** Answer the four questions, then **Save draft**:

| Question | Answer |
| --- | --- |
| Will your app be listed in the Apple Store? | No. It hasn't been tested on iPad, and that listing needs an Apple Developer ID. |
| Does your app use Microsoft Entra ID or SSO? | No. |
| Does your app require additional purchases? | Yes. It works without an account, but more than 10 calls an hour needs paid p3n credit. The description says so, and the notes tell the reviewer that no purchase is needed to test. |
| Connect a lead management CRM? | No. |

## 4. Packages

**[lead]** Upload `manifest.xml` from the repo. Its **Status** should read **Complete**. `npm run validate` runs the same validation service beforehand.

## 5. Properties

1. **[lead]** Categories (1 to 3): Productivity, Utilities. No industries.
2. **[user]** Legal: tick **Standard Contract**, Microsoft's standard licence for users of the add-in, or give a licence link. The add-in's code is MIT either way. This is a legal choice, and once the add-in is published it can't be undone. Recommended: Standard Contract.
3. **[lead]** Privacy policy link: https://github.com/Barabosik/p3n-office-addin/blob/main/PRIVACY.md. It names the add-in and says what happens to personal data, as Microsoft requires.
4. **[lead]** Support document link: https://github.com/Barabosik/p3n-office-addin/issues
5. **[lead]** **Save**.

## 6. Marketplace listings

1. **[lead]** **Manage additional languages** › English (United States) › **Update**.
2. **[lead]** Click English, then fill in the name, summary, description and search keywords from `listing.md`.
3. **[lead]** Upload the logo `store/logo-300.png`, and the screenshots in this order: `screenshot-word-check.png`, `screenshot-word-replace.png`, `screenshot-excel-replace.png`. They're 1366×768 and under 1024 KB.
4. **[lead]** **Save**.

## 7. Availability

**[lead]** Release as soon as it's certified, and click **Save draft**. The schedule can't be changed after the first publish.

## 8. Review and publish

1. **[lead]** **Review and publish** › paste **Notes for certification** from `listing.md` › **Publish**.
2. Partner Center emails the account owner and posts to its Action Center. If there are problems, reviewers answer within three to four business days. **[user]** Forward those emails to the lead.

## Later updates

- **The pane** (anything in `src/`) goes live with a push to `main`, through the Pages workflow, without a resubmission. Keep the listing true to what the pane does.
- **The manifest** (buttons, hosts, URLs, name): raise `<Version>`, keep the `<Id>`, and upload the new `manifest.xml` under **Packages**, which starts another review.

## Needs the user's screen

None. Nothing above needs Word or Excel in front. Partner Center runs in any browser, and the [user] steps are the sign-in, the account details and the legal choices.

## Sources

- [Step-by-step submission guide](https://learn.microsoft.com/partner-center/marketplace-offers/add-in-submission-guide)
- [Open an Office account in Partner Center](https://learn.microsoft.com/partner-center/marketplace-offers/open-a-developer-account)
- [Publishing checklist](https://learn.microsoft.com/partner-center/marketplace-offers/checklist)
- [Effective listings: name, summary and description limits](https://learn.microsoft.com/partner-center/marketplace-offers/create-effective-office-store-listings)
- [Store images: 1366×768, 1024 KB](https://learn.microsoft.com/partner-center/marketplace-offers/craft-effective-appsource-store-images)
- [Q&A: publishing with a personal account (Jan 2026)](https://learn.microsoft.com/answers/questions/5749417/publish-office-add-in-to-marketplace-with-personal)
