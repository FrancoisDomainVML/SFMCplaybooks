# Playbooks

Internal playbooks generated from YAML content files and published as PowerPoint decks.

## How it works

```
content/playbook-name.yaml   ← edit this to update a playbook
        ↓  (GitHub Actions)
output/playbook-name.pptx    ← committed back to repo automatically
        ↓  (you download manually for now)
SharePoint folder            ← source of truth for end users
```

**The YAML file is the only thing you ever edit.** The `.pptx` is generated automatically — never edit it directly.

---

## Making a change

1. Edit the relevant file in `content/` (e.g. `content/playbook-incident-response.yaml`)
2. Commit and push to `main`
3. GitHub Actions regenerates the `.pptx` within ~1 minute and commits it to `output/`
4. Download the new `.pptx` from `output/` and upload to SharePoint

To update only one field (e.g. a contact email or a bullet point), just change that line in the YAML — no full rewrite needed.

---

## Creating a new playbook

1. Copy `content/playbook-example.yaml` and rename it
2. Edit the content
3. Push — the PPTX appears in `output/` automatically

---

## Running locally (optional)

```bash
npm install
npm run generate                           # regenerate all playbooks
node scripts/generate_pptx.js content/playbook-name.yaml  # one file only
```

---

## Slide types

Each slide in the YAML has a `type` field:

| Type | Description |
|------|-------------|
| `title` | Cover slide with title, subtitle, owner, version |
| `section` | Divider slide — marks the start of a phase or topic |
| `content` | Heading + bullet points |
| `process` | Numbered step cards in a horizontal flow |
| `table` | Header row + data rows |
| `contact` | Two-column cards for contacts and resources |

---

## YAML structure reference

```yaml
meta:
  title: "Playbook Name"
  subtitle: "One-line description"
  version: "1.0"
  last_updated: "2026-06-24"
  owner: "Team Name"
  contact: "team@company.com"

slides:
  - type: title
    title: "..."
    subtitle: "..."
    notes: "Presenter notes go here"         # optional on any slide

  - type: section
    heading: "Phase 1: ..."
    subheading: "..."

  - type: content
    title: "Step N — ..."
    bullets:
      - "First point"
      - "Second point"

  - type: process
    title: "Flow name"
    steps:
      - label: "Step label"
        description: "Short description"

  - type: table
    title: "Table title"
    headers: ["Col 1", "Col 2", "Col 3"]
    rows:
      - ["Row 1 A", "Row 1 B", "Row 1 C"]

  - type: contact
    title: "Key Contacts"
    contacts:
      - role: "Role label"
        detail: "Contact detail or link"
```

---

## Versioning

- **GitHub** = full change history on every YAML file (who changed what, when, why)
- **SharePoint** = built-in versioning on the `.pptx` once uploaded
- The `version` and `last_updated` fields in `meta:` appear on every slide as a footer

Bump `version` in the YAML whenever you make a meaningful update so readers can tell at a glance which version they're looking at.
