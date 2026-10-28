# Housing CT · Prefab Gap Explorer 2

An interactive version of `Housing_CT_Feedback_Response_Model.xlsx` (Oct 14, 2026), the Prefab Modular team's model rebuilt after External Review 1.

**Live page:** https://eytanjda123.github.io/HousingCTModularModel-v2/
**Version 1 (original, unchanged):** https://eytanjda123.github.io/HousingCTModularModel/

## What's new in version 2
- **Two gaps.** The sales gap (appraised value minus what the buyer can pay) takes buyer-side money only. The developer gap (TDC minus sale proceeds) takes developer-side money only.
- **Line-item pro forma** built from the redacted course budgets: site work, soft costs, Phase I, construction-loan interest and a developer fee.
- **Four modular manufacturers with learning curves by order volume,** compared against stick-built calibrated to ~$800K.
- **80% AMI base case** and an appraisal-basis choice.
- **Homeowner affordability** at 60/80/100/120% AMI, with energy burden, MIT living-wage residual and stress tests.
- **Mixed-income portfolio** across Hartford, New Haven and Waterbury.

For New Haven at the workbook's own inputs, the headline numbers match the workbook. Hartford and Waterbury reuse New Haven's itemized site and soft costs, which are flagged. Manufacturer costs are faculty placeholders, and program amounts are caps or precedents, not awards.


## Add-on features and procurement

The inputs panel now has nested, collapsible sub-sections (construction, site work, soft costs). A new "Add-on features and procurement" group covers rooftop solar (upfront cost plus owner bill savings under CT's RRES tariff), battery storage, an efficiency / heat-pump package, two custom add-on slots, other hard-cost line items, and a warehouse for bulk buying. Everything is off or $0 by default, so the base results still match the workbook. Sliders extend when dragged to their end, and the number box accepts any value.


## October 28 update

- **Build it mode:** walks through the 10 input categories one step at a time, with the decisions for each step and a running budget, then shows the full budget.
- **New Markets Tax Credits:** added to the developer-side stack (off by default), sized from the NY Fed, Chicago Development Fund and Novogradac primers.
- **Two appraisal cases side by side:** existing-home value (CAMA) vs. new-build value, for all three cities. Also a choice of assessed value for property tax.
- **Sourced defaults:** CHFA mortgage rate, Housing Authority of New Haven utility allowances and water/sewer, city insurance averages, and prime + 1% construction loan. The Hazel Street and two-unit pilot precedents are cited.
- **Export:** downloads a complete, live Excel model (generated for this explorer) for the focus city: submission summary, sources and uses, pro forma, gaps, homeowner test, appraisal cases, add-ons, NMTC, sensitivity, modular vs. stick-built, three-city snapshot, and an Engine tab. It matches the explorer exactly.
