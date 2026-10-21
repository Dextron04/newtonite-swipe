# Newtonite Swipe

Tinder-style candidate screening for the Newtonite engineering challenge responses. Static site, no build, no backend.

1. Drop the responses `.xlsx` (or save it as `data.xlsx` next to `index.html` to auto-load; it is gitignored).
2. **Columns**: every column is classified (name, email, repo, stack, write-up, checklist...) from header and contents. Fix any role with the dropdown.
3. **Review**: swipe right = accept (green), left = reject (red), down = maybe (yellow). Buttons and keys work too: `←` `↓` `→`, `Z` to undo.
4. **Board**: all decisions tinted by outcome, editable, filterable, exportable to CSV.

Cards show contacts, links, stack, write-up, checklist (missing items flagged), plus live GitHub data: profile stats, contribution graph, and repo signals (README, tests, decisions doc, Docker, CI, file count). Duplicate submissions are merged by email, latest wins.

Decisions are stored in `localStorage`; the spreadsheet never leaves your browser. GitHub's unauthenticated API allows 60 requests/hour (3 per candidate, cached 12h); paste a personal access token when prompted to lift that.

```bash
python3 -m http.server 8000   # then open http://localhost:8000
```
