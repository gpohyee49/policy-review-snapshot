# Policy Review Snapshot 保单检视摘要

A small web app for annual insurance policy reviews in Malaysia. Load a client's policy list and it gives you:

- a dashboard of existing cover against your review standard (Life, TPD, CI, Retirement, Medical),
- a bilingual (English + 中文) keepsake PDF report for the client,
- an Excel copy with the policy table, data checks and a change log.

**Open the app:** see the GitHub Pages link in this repository's *About* section. On Chrome or Edge you can install it (address bar → *Install*) and it then opens like a normal app, offline after the first visit.

## Privacy

Everything runs in your browser. Excel files you load are read on your own device and are **not uploaded anywhere**. This repository contains only the app, a blank template and a fictional sample client.

## How to use

1. **Load** your policy summary Excel, an ePartner export, or `Policy_Review_Template.xlsx` (or try *Load sample client*).
2. **Confirm columns** — matched automatically; check and adjust if needed.
3. **Client details** — name, age, review date, EPF / PRS balances, review standards (fixed RM or × annual income) and report style (Standard / DISC D · I · S · C).
4. **Add a policy** (optional) — e.g. a policy from another insurer.
5. **Insurer & purpose** — pick from the Malaysian life / family takaful list and a purpose for each policy.
6. **Data check** — fix anything marked *Must fix*, review warnings. Nothing is changed without your confirmation.
7. **Generate** the dashboard, then **Save PDF report** and **Download Excel**.

Review standards are planning benchmarks chosen by the adviser, not official figures. Benefits are always governed by each policy contract.

## Files

| File | Purpose |
|---|---|
| `index.html` | The app |
| `Policy_Review_Template.xlsx` | Input template with dropdowns and an Excel-only dashboard |
| `sample/sample-report-fictional.pdf` | Example report (fictional client) |
| `manifest.webmanifest`, `sw.js`, `icons/` | Make the app installable and usable offline |
| `tools/` | Scripts that rebuild the template and this app folder |

Screenshot reading (Claude reads ePartner screenshots) is available only in the claude.ai version of the app.
