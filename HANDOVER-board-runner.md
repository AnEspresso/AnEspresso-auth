# AnEspresso handover — board-runner branch

Written 2026-08-28 ~18:40 EDT for a **new Grok Build chat**.  
Peter Gottschalk · live OR break board at Corewell / Beaumont Royal Oak.

If you are a new chat: read this whole file before editing. Attach the snapshots in `/workspace/public/reports/` (or this repo’s `board-runner` branch).

---

## 0. Hard rules

1. **Do not ship to production.** Live phones use GitHub Pages `main`. Do not merge, do not run a ship script, do not PUT a new Firebase `minVersion`.
2. **Do not edit `main`.** All new work goes on **`board-runner`**.
3. The live app is used every weekday morning. After ~15:00 usage drops; weekends are a separate mode.
4. No PHI. No Epic. Names on the CRNA assignment sheet are staff names only.
5. Design language: cream / espresso brown, quiet type, tiny pills, no clutter. Match existing AnEspresso, don’t invent a second UI.
6. Grok preview must keep serving the PWA on **`0.0.0.0:8080`** from `/workspace/public`. The python `http.server` dies often — restart it.

---

## 1. The two versions (posted)

| | **Live production** | **Board-runner (unshipped)** |
|---|---|---|
| Git branch | `main` | `board-runner` |
| Commit | `eb9c743` — *Ship AnEspresso 1787844240320* | `74ec276` — *Board-runner preview: assignment sheet, On deck, shift prefixes.* |
| `APP_VERSION` | **1787844240320** | **1787928630719** (preview only; not published as minVersion) |
| Pages | https://anespresso.github.io/AnEspresso-auth/ | not published (Pages source is `main` `/`) |
| Repo | https://github.com/AnEspresso/AnEspresso-auth | same repo, other branch |
| Files | `index.html` + `sw.js` | those **plus** `assignment.js` + `LATE-BOARD.md` |
| Assignment upload | no | yes |
| Snapshot in this workspace | [`live-index.html`](live-index.html) [`live-sw.js`](live-sw.js) | [`board-runner-index.html`](board-runner-index.html) [`board-runner-assignment.js`](board-runner-assignment.js) [`board-runner-sw.js`](board-runner-sw.js) |

Working copies while building:

- Source: `/workspace/anespresso/index.html` + `/workspace/anespresso/assignment.js`
- Preview tree: `/workspace/public/` (must stay in sync; copy after every edit)
- Git clone: `/workspace/AnEspresso-auth` (on `board-runner`)
- This Grok App Builder workspace is a TanStack scaffold around the real app. **The product is the single-file PWA**, not the React starter.

---

## 2. What live production already does

Single-file PWA. Firebase anonymous auth + Realtime Database. Last-write-wins merge. Roles: **Board Runner** vs CRNA (room / breaker).

- Categories: NT Tower, ST Tower, STE.100, CCS, FBC, Endo, NORA, EP Lab.
- Windows: Morning / Lunch / Afternoon / Evening.
- Board runner: Edit Active Anesthesia Sites, deactivate rooms, notes, break taps, Deactivate All.
- Evening: E / D / N kind cycle on rooms.
- Weekend mode with a different POS cap.
- **Day rollover** (the 2026-08-28 morning bug): `checkDayRollover()` + `applyDailyDefaults()`, called on visibility, poll, and a **60s interval** while the app is loaded. Live has this. Do not regress it.
- Telemetry counts (anonymous). `minVersion` force-refresh for stale phones.
- Service worker is **push-only** — it does not cache HTML. Do not turn it into an app-cache SW.

Firebase:

- RTDB: `https://anespresso-auth-default-rtdb.firebaseio.com`
- Project / API key already baked into the share URL in `index.html` (`anespresso.github.io/AnEspresso-auth/?fb=...&key=AIza...`).
- Identity Toolkit sign-in needs `Referer` / `Origin` headers when minting tokens from a script (browser origin is fine).

---

## 3. What board-runner adds (preview only)

Goal: replace the paper the board runner rewrites all day — who is in which room, their shift, open rooms, POS, float/breaker pool, automatic late/dinner lists.

### 3.1 Morning upload

Button under **Edit Active Anesthesia Sites**: “Upload CRNA assignment sheet”. Accepts `.xlsx` only (weekday grid or weekend split). Parsed **in the browser** (ZIP + XML + `DecompressionStream`). No CDN, no SheetJS.

Confirm modal: date, named rooms, CLOSED count, unlisted rooms that will hide, late (M+S), dinner (Q+W+E), on-deck count, unmatched labels. Apply writes into `roomStaff` + `catEditState.deletedRooms` and re-renders.

Sample sheets (in `/workspace/attachments/`):

- `8.27.26 CRNA Assignments.xlsx` — weekday (primary test)
- `8.17.26.xlsx` — weekday
- `8.15-8.16.2026 WEEKEND SPLIT.xlsx` — weekend
- `Blank8.26.26.xlsx` — blank template

### 3.2 Shift lexicon (Peter’s rules)

Listed times include the extra 30 minutes. **Algorithm end times drop that 30.**

| Letter | Listed | Algorithm |
|---|---|---|
| D | 7a–3:30p | 7–15 |
| d | 7a–5:30p | 7–17 |
| M | 6a–6:30p | 6–18 |
| S or s | 7a–7:30p | 7–19 |
| Q | 7a–9:30p | 7–21 |
| W | 7a–11:30p | 7–23 |
| E | 3p–11:30p | 15–23 |
| N | 11p–7:30a next day | overnight |
| t | 7p–7:30a next day | overnight |
| Dr | MD / attending | no CRNA shift |
| `*` prefix or `D*` | same letter, **1 hour earlier** (`*D` = 6–14:30 listed / algo 6–14) | |
| `o/` prefix | orientation (`o/D`, `o/*D`). Resource person is a breaker on deck | |
| `/` | combined (`W/N`, `E/N`) | |

**M gets a late break with S** (both `kind: late`). Dinner = Q+W+E.

### 3.3 Parser (`assignment.js`)

Weekday layout:

- A/B North Tower (+ CCS block)
- E/F South Tower + STE 100 + OB
- C/D Endo, EP, offsite / NORA
- G/H midnight, late stay, call, **NT/ST breakers**
- Loops run rows **4–44** so STE 108 / 109 / BMBx2 are included (earlier bug: loop stopped at 37, so Crudo on 109 never landed)

Apply rules:

- Named + not CLOSED → activate room, show last name + shift pill.
- Sheet says CLOSED → deactivate.
- **Room not on the sheet at all → deactivate** (hide empty tiles / drop POS).
- First named assignment to a room wins; a second person for the same room goes **On deck** (stops offsite WBF from overwriting NT OR 1).
- `rotate` lines skipped (OR 38-39).
- Late stay `#1…#5` = “may stay 2 hours past shift”. **Do not put them on On deck.** They already have a room.
- Breakers (CV, CCS, 3N, 2N, ENDO, ST 1/2, STE 1/2) → On deck, tag `breaker`. Morning POC / “write names outside OR” is not a room.
- Call names sometimes sit in column G on the next row (MN Call / Heart Call).
- Unmatched labels (IR 17, res tee, E/N, OB Resident, Pal Jones / ENDO B) → On deck or the unmatched list.

`staffChipHtml` is called from `renderCatCard` in `index.html`. Shift pill sits **on the name line**, not over the room number. Long last names shrink (`long` / `tiny` classes). No ellipsis if we can help it.

### 3.4 On deck

Coffee-themed float rail at the top of the board.

- Tap a chip, then tap a room → place them. Whoever was in that room comes On deck as `last OR n`.
- Tap the chip again to cancel.
- Deactivating a named room also sends them On deck.
- Afternoon: D / `*D` fade toward “out” via `stillInHouse()`.

### 3.5 Preview bootstrap

`index.html` load path: if `?preview=1`, `skipSetup()`, `currentRole='runner'`, `enterApp()`, `mountUploadButton()`. Grok’s preview host should pass that. If the board is a blank cream page after Apply, `enterApp()` / `buildBoard()` didn’t run.

---

## 4. How to run the preview (next Grok)

```
python3 -m http.server 8080 --bind 0.0.0.0 --directory /workspace/public
```

Wrap in `while true; do …; sleep 1; done` — the process dies and the user sees a black “Couldn’t load preview” spinner. After every `assignment.js` / `index.html` edit:

```
cp /workspace/anespresso/assignment.js /workspace/public/assignment.js
cp /workspace/anespresso/index.html /workspace/public/index.html
```

Do **not** add `frame-ancestors 'none'` to CSP.

---

## 5. Git

```
Repo:  git@ / https://github.com/AnEspresso/AnEspresso-auth.git
Live:  origin/main          → GitHub Pages
Dev:   origin/board-runner  → this work
```

Peter pasted a `GITHUB_TOKEN` (fine-grained PAT) in an earlier chat. **Do not commit it. Repo is public.** If push fails, ask Peter for a fresh token.

Push **only** `board-runner`. Never `git push origin main` unless Peter explicitly says ship.

---

## 6. What is still open

Pinned from Peter (do not lose):

1. Not everyone after 15:00 gets a break. People who leave before 19:00 do not. S (to 19:00) get a **late** break. People to 21:00 or later get a **dinner** break. **M rides with S.**
2. Afternoon still has many rooms running + many late/dinner breaks to give. Do not collapse the board just because D has left.
3. Yes to making the digital board as useful as the paper: open rooms, POS, who is in which room, their shift, automatic break list.
4. Morning upload of the assignment sheet is the source of names/shifts/call. After that the runner only maintains open rooms and who is where.

Not built / incomplete:

- iOS Share Sheet → AnEspresso (PWA cannot be a share target). In-app picker is the path.
- Automatic break **queue** (ordered list of who is due a late/dinner), not just counts.
- Saving assignment staff to Firebase so other phones see names without re-upload (currently apply → `saveShared()`; confirm this actually persists `roomStaff` on the payload).
- Weekend parser is thinner than weekday (no CLOSED-unlisted pass as thoroughly tested).
- Typo names on the sheet (`Fonatine` vs `Fontaine`) will miss last-name match.
- `IR 17` is not an app room (IR 10/11/12/14 are).
- Board-runner paper: live edits of “move Braun from OR 2 to OR 10” works via On deck; no drag.
- Daily usage digest file (Peter asked for a markdown report in the workspace every morning, not email / GrokBot). Script was `daily_digest.py` — **not in this workspace anymore**. Rebuild from telemetry in RTDB if asked.
- `ship.py` is also gone from this sandbox. Recreate from git history / earlier chat before any production ship.
- Epic “room open/closed without PHI” was discussed and **rejected** for now.

---

## 7. Design / UX notes Peter already gave

- Shift and name must read as two things, one line, no cutoff.
- Dr badges must **not** cover the OR number.
- On deck should not look like a second assignment. Only show `last OR n` after someone is pulled off a room; breakers get a quiet `breaker` tag; late stay is not a location.
- Keep it “beautifully simplistic.” Cream tiles, brown type, espresso cups already in the live UI.

---

## 8. Security (already reviewed with Peter)

- Client API key + anonymous auth is the current model (same as SitePlumb-style “leave the token open”).
- Hardcoded board-runner password in the PWA (pre-existing). Don’t expand it.
- No new HIPAA surface. Assignment xlsx never leaves the phone except whatever `saveShared()` already writes (staff names + shifts if wired).
- Public GitHub repo: never commit PATs, Firebase ID tokens, or dump files with tokens.

---

## 9. Suggested first steps for the new chat

1. Confirm preview is up (cream AnEspresso board, not the Grok “BUILD HANDOFF” page and not a black spinner).
2. Have Peter upload `8.27.26 CRNA Assignments.xlsx` (or that day’s sheet).
3. Check: `*D` / `o/D` pills, Dr on the name line, CLOSED + unlisted rooms gone, Crudo on OR 109, no late-stay duplicates on On deck, breakers on deck, tap-to-place still works.
4. Then continue board-runner features. Do not merge to `main`.

---

## 10. Product copy for Peter

Live phones are on **1787844240320**.  
Board-runner preview is **1787928630719** plus `assignment.js`.  
Same GitHub repo, two branches. Production stays as it is until you say ship.
