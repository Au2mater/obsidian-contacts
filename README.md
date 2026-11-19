# Contact Links

An Obsidian plugin that turns your contacts into person links. It syncs from Microsoft 365 (your Outlook contacts and the people you work with), from Google (your Google contacts and the people you email in Gmail), or both, and it can import a CSV or vCard file. When the same person turns up in more than one place, the copies are merged into one suggestion. As you type, the plugin suggests people, and picking one inserts a link. The first time you link someone, a minimal person note is created, so links, aliases and backlinks all work the way Obsidian normally does.

## How it works

| Goal | How |
|---|---|
| **Autocomplete** | Suggestions come from the synced and imported contacts and the notes in your people folder. Typing `soren` finds "Søren", `sj` finds Søren Jensen, and nicknames and aliases match too. People you pick often rank higher. |
| **Aliases** | An Outlook nickname becomes an alias. Aliases you add yourself to a person note's `aliases:` are picked up automatically. If you type an alias, the link becomes `[[Full Name\|alias]]`. |
| **Links** | The first time you pick someone, `People/Full Name.md` is created with only frontmatter (`aliases`, `email`, `company`, `role`). From then on the person is a real note, with graph node, backlinks and rename support. People you never mention never get a file. |

### Where suggestions appear

- **Inside `[[`** after 2 characters, when the text matches a person's name or alias. If nothing matches, or you press Esc, Obsidian's own note suggestions take over.
- **After `@`** anywhere in text, for example `@søren je` → `[[Søren Jensen]]`. The `@` is replaced. This doesn't trigger inside email addresses or inline code.
- **On plain capitalised names** if you turn that on in settings. Accept with Tab; Enter still makes a new line.

The link follows your vault's link settings (wikilink or Markdown, path format).

## Syncing from Microsoft 365

The plugin signs in to Microsoft the same way as the Outlook Calendar Panel and Microsoft To Do Panel plugins: an OAuth device code, your own Microsoft Entra app registration, and a session kept in Obsidian's secret storage on your device, outside the vault. Access is read-only.

It syncs two sources, and you can turn either off:
- **Saved contacts** (`/me/contacts`, permission `Contacts.Read`).
- **People you work with** (`/me/people`, permission `People.Read`): colleagues and people you email, ranked by Microsoft by how much you work with them, 500 by default. The ranking also lifts them in suggestions.

A sync runs shortly after Obsidian starts and then every 12 hours (you can change this). You can also use **Sync now** in settings or the command **Contact Links: Sync contacts now**, which syncs every account you're signed in to. Each sync replaces the contacts previously synced from that account; contacts from the other account and from files are kept. Signing out of an account removes the contacts synced from it; person notes are kept.

### Setup

1. Use an existing app registration (the calendar or To Do plugin's works), or register one: in the [Microsoft Entra admin center](https://entra.microsoft.com), go to **App registrations → New registration**, leave the redirect URI empty, and under **Authentication** set **Allow public client flows** to **Yes**.
2. Under **API permissions → Add a permission → Microsoft Graph → Delegated**, add **Contacts.Read** and **People.Read**. If your organisation requires it, grant admin consent.
3. In Obsidian, go to **Settings → Contact Links → Sign in**, enter the client ID and tenant (your directory ID for a single-organisation app), select **Continue**, and enter the code on the Microsoft page.

Each plugin keeps its own sign-in, so you sign in once here even if the other plugins are already signed in.

## Syncing from Google

The plugin signs in with your own Google Cloud OAuth client. Access is read-only. It syncs two sources, and you can turn either off:
- **Saved contacts** (`people/me/connections`, scope `contacts.readonly`).
- **Other contacts** (`otherContacts`, scope `contacts.other.readonly`): people you've emailed in Gmail but never saved. Only the ones with a name are used. If Google refuses one of the two lists (for example because a permission was left unticked when you signed in), the other is still synced and the settings show what failed.

Signing in to Google works in Obsidian on **desktop only**. Google doesn't allow the contacts permissions in the code-on-another-device flow that Microsoft uses, so the browser returns to a short-lived listener on `127.0.0.1`. Contacts synced on desktop still reach your phone with the vault.

### Setup

1. In the [Google Cloud console](https://console.cloud.google.com), create a project (or use an existing one) and turn on the **People API** under **APIs & Services → Library**.
2. Under **APIs & Services → OAuth consent screen**, set up the consent screen and add the scopes `contacts.readonly` and `contacts.other.readonly`.
   - With a Google Workspace account, choose **Internal**.
   - With a personal Gmail account, the user type is **External**. While the app is in **Testing**, add yourself as a test user. Google ends a Testing app's sign-ins after 7 days, so you'd have to sign in again each week. Publishing the app avoids that; without Google's verification you then click through an "unverified app" warning when you sign in, which is fine for your own use.
3. Under **APIs & Services → Credentials → Create credentials → OAuth client ID**, choose **Desktop app**. Copy the client ID and the client secret.
4. In Obsidian, go to **Settings → Contact Links → Google → Sign in**, enter the client ID and secret, select **Continue**, and allow access in the browser.

The client secret and the session are kept in Obsidian's secret storage on your device, not in the vault.

## Duplicates

With more than one source, the same person often shows up twice. Contacts are merged when:
- they share an email address (checked across all their addresses, ignoring case), or
- one of them has no email and exactly one other person has the same name. Two people with the same name and different emails are merged only if you turn on **Merge people with the same name** (off by default); it's useful when someone appears with both a work and a private address.

The merged contact keeps the name and nickname from the first source in this order: Microsoft 365, Google, files. Its email addresses are combined, company and job title come from the first source that has them, and the Microsoft relevance rank is kept. No-reply and notification senders (`noreply@`, `notifications@`, GitHub reply addresses and the like) are left out of every source.

Merging happens only in memory: each account's contacts are stored as synced, so signing out or turning a source off never leaves half a merge behind.

## Importing from a file

1. Export your contacts:
   - **Classic Outlook:** File → Open & Export → Import/Export → Export to a file → Comma Separated Values.
   - **New Outlook / Outlook on the web:** People → Manage contacts → Export contacts.
   - **Google Contacts:** Export → vCard.
2. In Obsidian, run the command **Contact Links: Import contacts from file**, or use the button under Settings → Contact Links. **Clear imported** removes the contacts imported from files.

Both CSV and vCard (`.vcf`) files work. That includes English and Danish column headers, comma or semicolon separators, UTF-8 or Windows ANSI encoding, and Outlook's vCard 2.1 quoted-printable format. Importing again updates contacts by email address; it doesn't create duplicates.

## Notes outside the people folder

If a contact has no note in the people folder, but a link to their full name already opens a note somewhere else (say `Colleagues/Søren Jensen.md`), that note is used as their person note. Suggestions link to it and the panel opens it, so no second note with the same name is created. This uses Obsidian's own link resolution and only reads your notes. To stay on the safe side, it applies only to full names of two or more words, skips a note whose `email` belongs to someone else, and never uses one note for two contacts. You can turn it off under **Use existing notes outside the folder**.

Nothing is moved unless you ask:
- In the contacts panel, these people have a "note in *folder*" tag. Right-click → **Move note to the people folder**.
- The command **Contact Links: Move matching notes to the people folder** lists every such note and moves them all after you confirm.

Moving uses Obsidian's rename, so links to the notes keep working. Right-click → **Add contact details to note** fills in the email, company, role and nickname a note is missing, and changes nothing else.

## Contacts panel

The contact icon in the ribbon, or the command **Contact Links: Open contacts panel**, opens a list of everyone the plugin knows in the right sidebar, with a search field that stays at the top. Each row shows where the person comes from (Microsoft, Google, File), whether they have a note, and their role, company and email.

- The refresh button next to the search field syncs every signed-in account now.
- Click someone with a note to open it (Ctrl/Cmd-click opens it in a new tab).
- Click someone without a note to create one: you're asked to confirm first.
- Right-click for **Open note** / **Create note…**, **Open in new tab**, **Add contact details to note**, **Move note to the people folder** and **Copy email address**.

## Settings

| Setting | Default |
|---|---|
| Microsoft 365: saved contacts / people you work with | on / on (500 people) |
| Google: saved contacts / other contacts | on / on |
| Merge people with the same name | off |
| Use existing notes outside the folder | on |
| Sync every | 12 hours |
| People folder | `People` |
| Create a note on first mention | on |
| Show the alias you typed | on |
| Inside `[[ ]]` | on, minimum 2 characters |
| Mention trigger | `@` |
| Suggest on plain names | off, minimum 3 characters |

## Development

```bash
npm install
npm test          # unit tests for import, mapping, deduplication, matching and trigger detection
npm run build     # type-check and bundle to main.js
npm run dev       # rebuild on change
```

To install, copy `main.js`, `manifest.json` and `styles.css` into `<vault>/.obsidian/plugins/outlook-contacts-autocomplete/`, then enable the plugin under Community plugins.

### Design notes

- Pure logic lives in `src/core/` and has no Obsidian dependency, so it can be unit-tested.
- The index is held in memory. It's rebuilt from the metadata cache only when contacts or people-folder notes change. Each keypress runs a regex and a linear scan over precomputed folded strings, which takes well under 1 ms for thousands of contacts.
- Taking priority inside `[[` relies on Obsidian internals: the ordered `workspace.editorSuggest.suggests` list, where the first suggester to trigger wins. Tab to accept relies on the popover's `suggestions.useSelectedItem`. Both were checked against the Obsidian 1.15.1 app bundle. If a future version changes them, the `@` trigger keeps working, and Tab falls back to its normal behaviour.
- Known limit: in plain-name mode, Enter makes a plain new line and doesn't continue a list bullet.
- Only names, nicknames, emails, company, job title and relevance rank are stored, in the plugin's `data.json` in the vault. If your vault is in Git, consider ignoring that file. The Microsoft and Google sessions and the Google client secret are never stored in the vault.
