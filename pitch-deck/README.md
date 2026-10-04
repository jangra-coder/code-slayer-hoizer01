# RepairOS: CodeSlayer 2K26 submission deck

| File | What it is |
|---|---|
| `RepairOS_CodeSlayer2K26.pdf` | The submission file (the rules accept PDF only) |
| `RepairOS_CodeSlayer2K26.pptx` | Editable deck. Speaker notes hold the talk track and full source URLs |
| `build_deck.js` | Generator. Change the team details here and rebuild |
| `assets/` | Organiser logos, banner and wordmark taken from the official template |

## Slide plan (6 slides max, title slide included)

1. **Title**: project, team, members and roles, track. Includes an illustrative product screen.
2. **Problem Statement & Target Audience**: what the problem is, four sourced statistics, where the decision breaks down, and the B2C and B2B segments.
3. **Our Unique Solution**: the five-step product flow and a comparison matrix against the existing alternatives.
4. **Tech Stack + Architecture**: the stack and the component and data-flow diagram.
5. **Feasibility and Showstoppers**: MVP scope, risks with their mitigations, and a 24-hour build plan.
6. **USP & Business Model**: the four USPs, why now, revenue streams and go-to-market.

The template's separate References slide would have been slide 7, which breaks the 6-slide limit. The sources are cited as footnotes on the slides that use them instead.

## Update team details

1. Edit the `TEAM` block at the top of `build_deck.js`.
2. Run `npm install && npm run build`.
3. Export the PPTX to PDF from PowerPoint, Google Slides or LibreOffice.

The team name in the header oval comes from the slide layout. To change it by hand in PowerPoint, go to **View → Slide Master**, edit the oval once, and every slide updates.

## Sources

1. Counterpoint Research, India smartphone after-sales service consumer survey, 15 Sep 2025. 43% said repair cost was high to very high; 4 in 10 needed repeat visits.
2. CPCB data in the MoHUA written reply, Rajya Sabha, 16 Dec 2024. India generated 1.751 MT of e-waste in FY2023-24, up from 1.01 MT in FY2019-20.
3. European Environmental Bureau / Coolproducts, *Coolproducts don't cost the Earth*, 2019. About 72% of a smartphone's climate impact comes from manufacturing, distribution and disposal.
4. Counterpoint Research, active smartphone installed base, 2025. India has more than 740M.
5. Department of Consumer Affairs, Repairability Index framework for the mobile and electronics sector, committee report, 2025. It rates six factors: disassembly depth, fasteners, tools, spare parts, software updates and repair information.
6. EU Delegated Regulation 2023/1669. Smartphone and tablet labels carry a repairability class from 20 Jun 2025.
7. MoEFCC, E-Waste (Management) Rules, 2022.

Prices, scores and confidence values in the deck are planned or illustrative and are labelled that way.
