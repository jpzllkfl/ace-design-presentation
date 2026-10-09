# A.C.E. Distribution — Website Design Review

Internal leadership deck: five homepage directions (two featured), EOS accountability, and the decision.

**Two versions, same slides:**
- **HTML:** `deck/index.html` (the root `index.html` redirects there). Multi-file: keep the `deck/assets/` folder next to it.
- **Live designs:** `designs/index.html`: all five homepage concepts as full, scrollable sites (plus the booking wizard and Round 2 concepts). Needs internet for photos and fonts.
- **PowerPoint:** `ACE-Design-Review.pptx`: speaker notes in the Notes pane, sections for Opening / Designs / Accountability / Decision / Appendix. Owner and due-date cells are blank table cells to type into. Fonts are Cambria + Calibri, so it looks the same on any Office install.

## Controls
| | Keyboard | iPad |
|---|---|---|
| Next / back | → ← · Space | Swipe, or tap right/left edge |
| Speaker notes | N | Two-finger tap |
| Slide overview | G | — |
| Fullscreen | F | Add to Home Screen |
| Open the live site | W (or the **Open live ↗** button) | Tap **Open live ↗** |

Owner names, due dates and the picked design on slides 16–17 are typed in place and remembered on that device.

## Swapping the featured pair
Default is **Big Sky + Heartland Steel**. To feature any two:
`deck/index.html?pair=piney,openwater`
(names: `bigsky`, `piney`, `openwater`, `heartland`, `hill`)

## Live designs
Every design slide in both decks has an **Open live ↗** link (in the HTML deck, **W** does the same), and the five thumbnails on the "Five directions" slide are links too.

The links open the hosted copy, so they work wherever the deck is opened, as long as there's internet:
- All five: https://raw.githack.com/jpzllkfl/ace-design-presentation/claude/ace-design-presentation-rebuild-b1ft9r/designs/index.html
- Big Sky: https://raw.githack.com/jpzllkfl/ace-design-presentation/claude/ace-design-presentation-rebuild-b1ft9r/designs/bigsky.html
- Piney Woods: https://raw.githack.com/jpzllkfl/ace-design-presentation/claude/ace-design-presentation-rebuild-b1ft9r/designs/piney.html
- Open Water: https://raw.githack.com/jpzllkfl/ace-design-presentation/claude/ace-design-presentation-rebuild-b1ft9r/designs/openwater.html
- Heartland Steel: https://raw.githack.com/jpzllkfl/ace-design-presentation/claude/ace-design-presentation-rebuild-b1ft9r/designs/heartland.html
- Hill Country: https://raw.githack.com/jpzllkfl/ace-design-presentation/claude/ace-design-presentation-rebuild-b1ft9r/designs/hill.html
- HTML deck: https://raw.githack.com/jpzllkfl/ace-design-presentation/claude/ace-design-presentation-rebuild-b1ft9r/deck/index.html

In PowerPoint, links only respond in **Slide Show** mode (or Ctrl/Cmd-click while editing).

To use GitHub Pages instead, enable it in the repo settings and change `LIVE_BASE` in `deck/index.html` and `pptx/build-pptx.js`.

## Files
- `ACE-Design-Review.pptx`: the PowerPoint version (featured pair: Big Sky + Heartland Steel)
- `pptx/build-pptx.js`: rebuilds the .pptx (`npm i pptxgenjs@3 && node pptx/build-pptx.js`); change `FEATURED` there to swap the pair
- `deck/index.html`: the deck
- `designs/`: the five concepts (`five-concepts.html`), booking wizard, Round 2 concepts, and per-design short links. Copied from `jpzllkfl/ACE` `design_handoff_ashercrest_website/`; React is bundled in `designs/vendor/` so it doesn't depend on a CDN
- `deck/SPEAKER-NOTES.md`: talk track and why this pair
- `deck/assets/`: gold/light ACE logo (transparent), five hero shots, five themed booking matrices
- `archive/ACE-Presentation-draft.html`: the earlier draft, kept for reference only
