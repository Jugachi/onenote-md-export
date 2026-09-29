# onenote-md

Exports your local OneNote notebooks to a Markdown folder tree.

It talks to the **OneNote desktop app you already have installed** through its
local COM automation interface. There is no Microsoft Graph, no Azure AD app
registration, no OAuth, and no API permission to grant — nothing leaves the
machine.

```
onenote-md                      # export everything to ./onenote-export
onenote-md D:\notes --list      # show the hierarchy, write nothing
onenote-md D:\notes --notebook <name>
onenote-md D:\notes --dry-run
```

## Build

```
build.cmd
```

Produces `bin\onenote-md.exe`. It needs nothing but the .NET Framework that
ships with Windows; the only external reference is
`Microsoft.Office.Interop.OneNote`, which is part of the OneNote/Office
install and is located automatically.

## Requirements

* **OneNote for Windows desktop** (Office 2016/2019/365) — the classic
  `ONENOTE.EXE`, not the Microsoft Store / UWP app. The Store app has no COM
  automation surface at all.
* You must be **signed in** to OneNote, because the tool reads the same
  hierarchy OneNote itself shows. Cloud notebooks come down through your
  existing sync; it does not fetch anything extra.
* Run it as your normal desktop user. An *elevated* shell can fail to attach to
  the already-running OneNote instance.

## Output layout

The hierarchy maps straight onto folders:

```
onenote-export/
  README.md                       index of everything that was written
  <Notebook>/
    <Section>/
      <Page title>.md
      assets/
        p01-img001.png             images extracted from that section
    <Section group>/
      <Section>/
        <Page title>.md
```

Each page gets YAML front matter:

```markdown
---
title: "Title"
notebook: "Notebook"
section: "Sect"
group: "Name"
onenote_id: "{E8DDE53E-...}{1}{E19536...}"
---
```

## Options

| Option | Meaning |
| --- | --- |
| `-o, --out <dir>` | Output directory (default `onenote-export`) |
| `--notebook <name>` | Only notebooks whose name contains `<name>` |
| `--section <name>` | Only sections whose name contains `<name>` |
| `--no-images` | Skip image extraction (much faster, much smaller) |
| `--list` | Print notebooks / sections / pages and exit |
| `--overwrite` | Overwrite existing files (default behaviour) |
| `--skip-existing` | Leave existing files untouched, for incremental runs |
| `-n, --dry-run` | Report what would be written, write nothing |
| `-h, --help` | Usage |


## What gets converted

| OneNote | Markdown |
| --- | --- |
| `quickStyleIndex` → `h1`–`h6` | `#`–`######` |
| `<one:List>` bullet | `-` list item, nested by outline depth |
| `<one:r>` runs with `font-weight:bold` | `**bold**` |
| `<one:r>` runs with `font-style:italic` | `*italic*` |
| `<span style="font-weight:bold">` | `**bold**` |
| `<one:Image>` with inline `<one:Data>` | `![alt](assets/…)` |
| `tableHTML` | GitHub-flavoured Markdown table |
| `<a href>` | `text (href)` |
| HTML entities in CDATA (`&amp;`, `&nbsp;`, `&#xA;`) | decoded characters |

Section groups (the `_Inhaltsbibliothek` / `Collaboration Space` folders inside
a notebook) become an extra directory level.

Page text arrives as a mix of three different encodings depending on how it was
typed, and all three are handled: plain CDATA, structured `<one:r>` runs, and
literal HTML fragments that OneNote stores directly inside `<one:T>`. The HTML
path is a tolerant tag scanner rather than an XML parse, because those
fragments are routinely unclosed (`<span>` without `</span>`), so no content is
ever dropped.

## Notes and limits

* **Image-heavy pages are expensive.** Fetching a page with images uses
  `PageInfo.piAll`, which inlines every image as base64. Some pages are tens of
  megabytes, which is why a large notebook with many screenshots still runs at a
  reasonable speed but will use noticeable memory. `--no-images` avoids it.
* **Attachments and embedded objects.** Inline images are extracted. OLE
  attachments and drawings are referenced by callback and are not written out.
* **Notebooks are read as OneNote sees them**, so only what has synced to the
  desktop appears. The OneNote window will visibly move around while the export
  runs; that is normal.
* **Existing files are overwritten** by default. Use `--skip-existing` for
  incremental re-runs.
* Page titles are used for filenames, so invalid Windows characters are
  replaced and very long titles are truncated to 120 characters.

## How it works

1. `CoCreateInstance` on `OneNote.Application`
   (`{DC67E480-C3CB-49F8-8232-60B0C2056C8E}`), apartment-threaded.
2. `GetHierarchy(hsPages)` once, to get every notebook, section and page id.
3. `GetPageContent(id, PageInfo.piAll)` per page.
4. The page XML is walked and rendered to Markdown; `<one:Data>` payloads are
   base64-decoded to the section's `assets` folder.

Because step 1 and step 2 go through the PIA rather than late binding, the tool
does not depend on the OneNote type library being correctly registered. On
machines where that registration is broken, PowerShell's `New-Object -ComObject
OneNote.Application` fails with `TYPE_E_LIBNOTREGISTERED` while this tool works
normally.
