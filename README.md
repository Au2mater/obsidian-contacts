# Contact Links

An Obsidian plugin that turns your Microsoft 365 contacts into person links. It syncs your Outlook contacts and the people you work with, or imports a CSV or vCard file. As you type, the plugin suggests people, and picking one inserts a link. The first time you link someone, a minimal person note is created, so links, aliases and backlinks all work the way Obsidian normally does.

## How it works

| Goal | How |
|---|---|
| **Autocomplete** | Suggestions come from the imported contacts and the notes in your people folder. Typing `soren` finds "Søren", `sj` finds Søren Jensen, and nicknames and aliases match too. People you pick often rank higher. |
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

A sync runs shortly after Obsidian starts and then every 12 hours (you can change this). You can also use **Sync now** in settings or the command **Contact Links: Sync contacts from Microsoft 365**. Each sync replaces the previously synced contacts; contacts imported from files are kept.

### Setup

1. Use an existing app registration (the calendar or To Do plugin's works), or register one: in the [Microsoft Entra admin center](https://entra.microsoft.com), go to **App registrations → New registration**, leave the redirect URI empty, and under **Authentication** set **Allow public client flows** to **Yes**.
2. Under **API permissions → Add a permission → Microsoft Graph → Delegated**, add **Contacts.Read** and **People.Read**. If your organisation requires it, grant admin consent.
3. In Obsidian, go to **Settings → Contact Links → Sign in**, enter the client ID and tenant (your directory ID for a single-organisation app), select **Continue**, and enter the code on the Microsoft page.

Each plugin keeps its own sign-in, so you sign in once here even if the other plugins are already signed in.

## Importing from a file

1. Export your contacts from Outlook:
   - **Classic Outlook:** File → Open & Export → Import/Export → Export to a file → Comma Separated Values.
   - **New Outlook / Outlook on the web:** People → Manage contacts → Export contacts.
2. In Obsidian, run the command **Contact Links: Import contacts from file**, or use the button under Settings → Contact Links.

Both CSV and vCard (`.vcf`) files work. That includes English and Danish column headers, comma or semicolon separators, UTF-8 or Windows ANSI encoding, and Outlook's vCard 2.1 quoted-printable format. Importing again updates contacts by email address; it doesn't create duplicates.

## Settings

| Setting | Default |
|---|---|
| Sync saved contacts / people you work with | on / on (500 people) |
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
npm test          # unit tests for import, matching and trigger detection
npm run build     # type-check and bundle to main.js
npm run dev       # rebuild on change
```

To install, copy `main.js`, `manifest.json` and `styles.css` into `<vault>/.obsidian/plugins/outlook-contacts-autocomplete/`, then enable the plugin under Community plugins.

### Design notes

- Pure logic lives in `src/core/` and has no Obsidian dependency, so it can be unit-tested.
- The index is held in memory. It's rebuilt from the metadata cache only when contacts or people-folder notes change. Each keypress runs a regex and a linear scan over precomputed folded strings, which takes well under 1 ms for thousands of contacts.
- Taking priority inside `[[` relies on Obsidian internals: the ordered `workspace.editorSuggest.suggests` list, where the first suggester to trigger wins. Tab to accept relies on the popover's `suggestions.useSelectedItem`. Both were checked against the Obsidian 1.15.1 app bundle. If a future version changes them, the `@` trigger keeps working, and Tab falls back to its normal behaviour.
- Known limit: in plain-name mode, Enter makes a plain new line and doesn't continue a list bullet.
- Only names, nicknames, emails, company, job title and relevance rank are stored, in the plugin's `data.json` in the vault. If your vault is in Git, consider ignoring that file. The Microsoft session is never stored in the vault.
