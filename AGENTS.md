# Agent Instructions

## Scope

This repository publishes University of Nottingham document templates. An
agent working here may build, render and verify freely; licence terms, the
brand palette and the public identity of the repository are not its to change.

## Working style

- Concise and present tense. State the rule, not the history.
- British English in prose; code and identifiers in English.

## Principles

- **One source of truth.** Colours and fonts live in `common/uon-brand.sty`;
  `thesis/` and `report/` load it. Never restate a colour or a face in a
  template that could reference the brand core instead.
- **No system fonts.** Every face must resolve through kpathsea from TeX
  Live (`kpsewhich` prints a path inside the TeX tree). A face outside the
  tree is an opt-in override, never a default.
- **LuaLaTeX is the engine.** `% !TEX program = lualatex` plus `$pdf_mode = 4`.
- **The artefact is the evidence.** For the Office templates the accepted
  rendering is what PowerPoint and LibreOffice produce — not what an agent
  tool reports. For LaTeX, read the built PDF: page count, `pdffonts`,
  overfull boxes and the front-matter order.
- **Binaries stay out of git.** `.pptx`, `.potx`, `.otp` and built `.pdf`
  are released as assets, never committed; the sources are what gets diffed.
- **Public tree, zero private identifiers.** No local paths, machine names,
  internal hostnames or personal mailboxes; the author appears only in
  `LICENSE` and in the commit identity.

## Completion discipline

Before calling work done: rebuild every template the change touches, state
the measured result (pages, errors, overfull boxes, embedded faces), and
update every document the change invalidates — `README.md` for user-facing
behaviour, this file for operating rules.
