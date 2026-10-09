# A.C.E. Distribution — Website Design Review

Internal leadership deck: five homepage directions (two featured), EOS accountability, and the decision.

**Two versions, same slides:**
- **HTML:** `deck/index.html` (the root `index.html` redirects there). Multi-file: keep the `deck/assets/` folder next to it.
- **PowerPoint:** `ACE-Design-Review.pptx`: speaker notes in the Notes pane, sections for Opening / Designs / Accountability / Decision / Appendix. Owner and due-date cells are blank table cells to type into. Fonts are Cambria + Calibri, so it looks the same on any Office install.

## Controls
| | Keyboard | iPad |
|---|---|---|
| Next / back | → ← · Space | Swipe, or tap right/left edge |
| Speaker notes | N | Two-finger tap |
| Slide overview | G | — |
| Fullscreen | F | Add to Home Screen |

Owner names, due dates and the picked design on slides 16–17 are typed in place and remembered on that device.

## Swapping the featured pair
Default is **Big Sky + Heartland Steel**. To feature any two:
`deck/index.html?pair=piney,openwater`
(names: `bigsky`, `piney`, `openwater`, `heartland`, `hill`)

## Files
- `ACE-Design-Review.pptx`: the PowerPoint version (featured pair: Big Sky + Heartland Steel)
- `pptx/build-pptx.js`: rebuilds the .pptx (`npm i pptxgenjs@3 && node pptx/build-pptx.js`); change `FEATURED` there to swap the pair
- `deck/index.html`: the deck
- `deck/SPEAKER-NOTES.md`: talk track and why this pair
- `deck/assets/`: gold/light ACE logo (transparent), five hero shots, five themed booking matrices
- `archive/ACE-Presentation-draft.html`: the earlier draft, kept for reference only
