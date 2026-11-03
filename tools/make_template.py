"""Builds the explorer's export template: a fully live workbook whose Engine tab mirrors compute() in the page.
Output: template.xlsx + xmap.json (input key -> cell) in this folder."""
import openpyxl, json, copy, sys
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as CL
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart import LineChart, BarChart, Reference

D = '/tmp/claude-0/-home-claude/cf885be6-d73a-57f1-84e3-0c4b8ede712e/scratchpad/sub/'
BASEWB = D + 'Housing_CT_Prefab_Modular_Financial_Model.xlsx'
wb = openpyxl.load_workbook(BASEWB)

NAVY = 'FF1E2A32'; GREY = 'FF56636B'; TOT = 'FFEAF1F5'
USD = '\\$#,##0;"($"#,##0\\);\\-'; USD2 = '\\$#,##0.00'; PCT = '0%'; PCT1 = '0.0%'; PCT2 = '0.00%'; NUM = '#,##0'; DEC = '0.00'
F = lambda **k: Font(name='Arial', **k)
title = F(sz=14, b=True, color=NAVY); subf = F(sz=9, color=GREY); hdr = F(sz=11, b=True, color='FFFFFFFF')
hfill = PatternFill('solid', fgColor=NAVY); tfill = PatternFill('solid', fgColor=TOT); ifill = PatternFill('solid', fgColor='FFFFF2CC')
body = F(sz=10); bold = F(sz=10, b=True); link = F(sz=10, color='FF008000'); linkb = F(sz=10, b=True, color='FF008000')
note = F(sz=9, color=GREY); inp = F(sz=10, color='FF0000FF'); calc = F(sz=10)
WRAP = Alignment(wrap_text=True, vertical='top')

for n in ['Summary', 'Submission Summary', 'Sources & Uses', 'Inputs', 'Pro Forma', 'Gaps & Stack', 'Homeowner',
          'Sensitivity', 'Scale & Method', 'Read Me', 'Sources', 'Mixed Income', 'Ownership Models']:
    if n in wb.sheetnames:
        del wb[n]

def setup(ws, widths, grid=False):
    ws.sheet_view.showGridLines = grid
    for k, v in widths.items():
        ws.column_dimensions[k].width = v

def sec(ws, r, text, c0=2, c1=7):
    for c in range(c0, c1 + 1):
        ws.cell(r, c).fill = hfill
    ws.cell(r, c0, text).font = hdr

def put(ws, ref, val, font=body, fmt=None, fill=None, al=None):
    c = ws[ref]; c.value = val; c.font = font
    if fmt: c.number_format = fmt
    if fill: c.fill = fill
    if al: c.alignment = al
    return c

# ============================================================ SOURCES
U = {
 'doh': 'https://portal.ct.gov/-/media/doh/homeownership-initiative/doh-homeownership-developer-subsidy-initiative-published-03-21-2024.pdf',
 'tto': 'https://www.chfa.org/assets/1/19/000_LPO_Final_Rev_11.13.2025_Final1.pdf?14506=', 'hdf': 'https://portal.ct.gov/governor/news/press-releases/2025/04-2025/governor-lamont-announces-ct-home-funds', 'dep26': 'https://portal.ct.gov/-/media/doh/dep-requirements-and-instructions-spring-2026.pdf', 'fhfaAhp': 'https://www.fhfa.gov/programs/affordable-housing-program', 'nhDpa': 'https://newhavenct.gov/government/departments-divisions/economic-development-administration/housing-administration-and-development/homebuyers/down-payment-closing-cost-assistance-financing', 'yale': 'https://your.yale.edu/working-at-yale/benefits/your-finances/homebuyer-program',
 'ahp': 'https://www.fhlbboston.com/?p=10990', 'htcc': 'https://www.chfa.org/assets/1/19/FINAL_2025_Guidelines.pdf?14018=', 'hudil': 'https://www.huduser.gov/portal/datasets/il.html',
 'chfa': 'https://www.chfa.org/homebuyer-mortgage-interest-rates-/', 'pmms': 'https://www.freddiemac.com/pmms',
 'mip': 'https://archives.hud.gov/news/2024/2023-05hsgml.pdf',
 'ccm': 'https://hub.ccm-ct.org/community/new-haven/', 'nw': 'https://www.nerdwallet.com/insurance/homeowners/connecticut-home-insurance',
 'mit': 'https://livingwage.mit.edu/metros/35300', 'cama': 'https://portal.ct.gov/opm/igpp/data-analytics/parcel-data',
 'ecc': 'https://elmcitycommunities.org/wp-content/uploads/2024/02/2024-Utility-allowance-schedule.pdf',
 'prime': 'https://www.bankrate.com/rates/interest-rates/wall-street-prime-rate/',
 'hazel': 'https://ctpublic.org/news/2024-12-03/new-haven-affordable-housing-development-first-in-partnership-between-state-yale',
 'pilot': 'https://portal.ct.gov/doh/doh/main/press-releases/fourteen-developments-across-ct',
 'reframe': 'https://bankerandtradesman.com/ramping-up-housing-production-through-robotics',
 'nyfed': 'https://www.newyorkfed.org/medialibrary/media/outreach-and-education/community-development-finance/investment-connection-puerto-rico/Intro-to-NMTC.pdf',
 'cdf': 'https://chicagodevelopmentfund.org/how-it-works/basic_transaction_structur/', 'smith': 'https://smithnmtc.com/affordable-homeownership/',
 'novo': 'https://www.novoco.com/podcast/dec-9-2025-unlocking-homeownership-nmtc-strategies-for-for-sale-housing',
 'cdfi': 'https://www.cdfifund.gov/sites/cdfi/files/2021-11/2021_Introduction_to_the_NMTC_Program.pdf',
 'olr': 'https://www.cga.ct.gov/2012/rpt/2012-R-0098.htm',
 'nlondon': 'https://news.bloombergtax.com/daily-tax-report/connecticut-superior-court-partially-grants-property-tax-relief-based-on-deed-restriction-value-cap',
 'hfdfix': 'https://hartfordbusiness.com/article/hartford-city-council-considers-10-year-tax-break-for-new-five-unit-rowhouses/',
 'escost': 'https://www.energysage.com/local-data/solar-panel-cost/ct/', 'es25d': 'https://www.energysage.com/news/congress-passes-bill-ending-residential-solar-tax-credit/?rc=p-tcd-w',
 'rres': 'https://www.eversource.com/rres', 'esrate': 'https://www.energysage.com/local-data/electricity-cost/ct/',
 'ess': 'https://solarbuildermag.com/news/connecticut-launches-statewide-battery-storage-incentive-program/',
 'whrent': 'https://www.commercialcafe.com/industrial/us/ct/new-haven/', 'hfdlots': 'https://www.ctpublic.org/news/2025-12-17/new-affordable-homeownership-opportunities-coming-to-hartford-in-2026',
}
INFILL = 'Redacted New Haven 1-2 unit infill budget (course files, Canvas).'
HOB = 'Redacted homeownership budget (course files, Canvas).'
V, SE, CB, PH, TC, PL = 'Verified', 'Sourced estimate', 'Course budget', 'Placeholder', 'Team choice', 'Program rule'
QT, LB, DX = 'Quote', 'Land bank actual', 'DOH example'
NHLB = 'New Haven Land Bank (Barbara Malmberg, Nov 2, 2026): construction estimates, cost estimates table and Infill Homeownership Initiative webinar.'
CHFAM = 'CHFA meeting notes, Oct 30, 2026.'
RFC = 'Reframe Systems call, Oct 30, 2026.'
OWN_T = ['Owner-occupant + rental unit(s)', 'Condo: every unit sold to its own buyer']
FOUND_T = ['Insulated slab on grade', 'Crawl space on stem walls', 'Helical piles, vented crawl space']
HTCC_T = ['Grant toward the developer gap', 'Revolving construction loan fund']
FACT_T = ['Shoreline quote: factory package', 'Reframe public range: all-in building', 'Faculty placeholder: factory + field', 'Custom factory package']
VERT = [('roof', 'Roofing', 8.627, 0.15), ('envelope', 'Exterior envelope', 9.371, 0), ('interior', 'Interior construction', 55.067, 0), ('siding', 'Siding', 9.996, 1), ('windows', 'Windows and doors', 10.520, 0), ('finishes', 'Finishes', 27.146, 0), ('hvac', 'HVAC', 11.246, 1), ('elec', 'Electrical', 12.288, 0), ('plumbing', 'Plumbing', 12.551, 0.2)]
LEVK = [('solar', 'levSolar', None, 'No solar in the base budget'), ('fence', 'levFence', None, 'No fencing and security'), ('pattern', 'levPatternOn', 'levPattern', 'Pattern book (architect fee)'), ('zoning', 'levZoningOn', 'levZoning', 'As-of-right zoning (legal)'), ('permit', 'levPermitOn', 'levPermit', 'As-of-right permits and fee discount'), ('cdfi', 'levCdfiOn', 'levCdfi', 'CDFI or foundation bridge financing'), ('realtor', 'levRealtorOn', 'levRealtor', 'Realtor discount on the commission'), ('fee', 'levFeeOn', 'levFee', 'Developer fee capped'), ('land', 'levLand', None, 'No acquisition cost')]
SIDE_B = {'capBasis', 'allinCap', 'ownerAMI', 'hhSize', 'rentAMI', 'rentShare', 'capMkt', 'voucherOn', 'vpsPct', 'ami50', 'mill', 'assess', 'assess4', 'dpa', 'chfaAmi', 'abateOn', 'abateVal', 'rate', 'term', 'dp', 'mip', 'rentCredit', 'lenderCap', 'ownerCap', 'ins', 'energy', 'common', 'maintPct', 'vac', 'mitType', 'energyThresh', 'taxBasis', 'taxCustom', 'incAssist', 'tto', 'otherAssist', 'useHB', 'empDpa', 'condoFee'}
SIDE_P = {'city', 'units', 'ownerBR', 'rentBR', 'own', 'basis', 'med2', 'med3', 'med46', 'psf2', 'psf3', 'psf46', 'prec', 'custom', 'lb', 'condoVal'}
bfill = PatternFill('solid', fgColor='FFEDE4FA'); dfill = PatternFill('solid', fgColor='FFDDF2EA')
def sidefill(k):
    return ifill if k in SIDE_P else (bfill if k in SIDE_B else dfill)

BRS = ['Studio', '1BR', '2BR', '3BR', '4BR']
METHODS = ['Volumetric modular', 'Stick-built']
BASIS = ['Existing-home value (CAMA median)', 'Size-adjusted new-build value (base case)', 'New-build value (DOH precedent)', 'Land bank market value (new 2,400 sq ft two-family)', 'Sales comps: 75th-percentile $/sq ft \u00d7 our size', "Neighborhood comps: the site's neighborhood $/sq ft \u00d7 our size"]
B_CAMA, B_CUST, B_PREC, B_LB, B_COMPS, B_NBHD = BASIS
TAXB = ['Appraised value (Connecticut law)', 'Existing-home value (restricted-value case)', 'Custom assessed value', 'Price the buyer can pay (affordable-price case)']
CAPB = ['All housing costs (mortgage, tax, insurance, utilities)', 'Mortgage payment only, with an all-in ceiling']
FEEB = ['Total costs before the fee', 'Hard (construction) costs', 'Sales proceeds (land bank method)']
SOMN = 'SOM meeting notes, Nov 2, 2026.'
OPMS = "Team sales-ratio study (ct_sales_comps_v9, Nov 2, 2026): arm's-length sales Sep 2023-Sep 2025, CT OPM Real Estate Sales (data.ct.gov 5mzw-sjtu) matched to CAMA (rny9-6ak2): https://data.ct.gov/d/5mzw-sjtu"
YN = ['Yes', 'No']
MIT = [('1 adult, 0 children', 52660, 16549, 36111), ('2 adults (1 working), 2 children', 94768, 23749, 71019),
       ('2 adults (both working), 1 child', 115580, 23749, 91831), ('2 adults (both working), 2 children', 145113, 23749, 121364)]

# (key, label, default, fmt, status, source, options)
I = []
def sect(t): I.append(('#', t))
def add(k, l, v, f, st, src, opts=None): I.append((k, l, v, f, st, src, opts))

sect('1. SCENARIO AND CITY')
add('city', 'City this workbook models (focus city in the explorer)', 'New Haven', None, TC, 'Set by the explorer at export. City-specific inputs below are for this city.', ['Hartford', 'New Haven', 'Waterbury'])
sect('2. BUILDING AND HOUSEHOLD')
add('units', 'Units in the building', 2, NUM, TC, 'Team design choice. CT DOH Developer Subsidy Initiative excludes structures of 3+ units: ' + U['doh'])
add('ownerBR', 'Owner unit bedrooms', '3BR', None, TC, 'Team choice: a family-sized home. Sets sq ft from the bedroom table.', BRS)
add('rentBR', 'Rental unit bedrooms', '2BR', None, TC, 'Team choice. Sets rent limit factor, utility allowance, market rent and sq ft from the bedroom table.', BRS)
add('ownerAMI', 'Owner household income target (% of AMI)', 1.0, PCT, TC, CHFAM + ' Buyers at 100-120% AMI, where the RKG study finds New Haven\'s ownership shortage; 80% and 120% are comparison cases (Engine columns).')
add('hhSize', 'Owner household size', 4, NUM, TC, 'HUD reference household (4 persons).')
add('voucherOn', 'Rent the unit to a voucher holder at the payment standard?', 'No', None, TC, 'Off by default. Rent = payment standard (bedroom table) x the % below - utility allowance.', YN)
add('vpsPct', 'Voucher payment standard as % of the Small Area FMR', 1, PCT, PH, 'Placeholder until Elm City Communities\' payment standards are in hand (usually 90-110% of FMR).')
add('rentAMI', 'Rental unit rent target (% of AMI; 0 = market rent)', 0.8, PCT, TC, 'Team choice: 80% AMI, with 60% (the DOH rental tier: ' + U['doh'] + ') as a comparison column in the Engine.')
add('own', 'Ownership structure', OWN_T[0], None, TC, 'Condo: every unit sold to its own buyer at the same density; no landlord and no rent. Unit values split the building by sq ft.', OWN_T)
add('condoFee', 'Condo fee per unit ($ per month)', 150, USD, PH, 'Condo only: master insurance, reserves and shared upkeep. Placeholder.')
add('condoVal', 'Condo: total unit sales vs. appraised value of the building', 1, PCT, PH, 'Condo only. Placeholder until there are condo comparables.')
add('rentShare', 'Rent as % of income (gross, incl. utilities) at the rent target', 0.30, PCT1, PL, 'HUD affordability standard: 30% of income; the basis of program rent limits. Lower it for a deeper discount.')
add('capMkt', 'Cap the AMI rent at market rent?', 'Yes', None, TC, 'An owner cannot charge more than the market will pay.', YN)
sect('3. CONSTRUCTION: THE BUILDING')
add('method', 'Construction method', 'Volumetric modular', None, TC, 'The team builds volumetric modular; stick-built is the baseline. Line items are in the table further down.', METHODS)
add('factOpt', 'Volumetric: factory price option', FACT_T[0], None, TC, 'A factory package (Shoreline, custom) leaves the local share of each line to a local contractor; an all-in price (Reframe, faculty) covers the whole building.', FACT_T)
add('factShore', 'Factory price: Shoreline quote ($ per sq ft, factory package)', 143, USD2, QT, "Shoreline's $501,620 for a ~3,500 sq ft 3-story duplex, per sq ft: boxes, delivery, set crew, crane, roof to weathertight. Excludes siding, foundation, site work. Likely higher on a 2,400 sq ft home (Kate's DOH workbook, via NHLB).")
add('factReframe', 'Factory price: Reframe public range ($ per sq ft, all-in building)', 287.5, USD2, SE, 'Reframe $275-300/sq ft all-in: ' + U['reframe'])
add('factFaculty', 'Factory price: faculty placeholder ($ per sq ft, factory $215 + field $88)', 303, USD2, PH, 'F26 Housing CT Model Workshop v2.2, Volume Curve tab.')
add('factCustom', 'Factory price: custom package ($ per sq ft)', 200, USD2, TC, 'Any other quote, treated as a factory package.')
add('factTrans', 'Transport, crane and set outside the factory price ($ per building)', 0, USD, TC, "$0: Shoreline's quote includes them.")
add('learn', 'Factory price drop per doubling of the order', 0, PCT1, PH, 'The quote is for one project; the faculty placeholder assumed 5%.')
add('scaleMod', 'Volumetric multiplier (factory price and local lines)', 1, DEC, TC, 'Set by the explorer\'s volumetric headline $/sq ft. 1.00 = as entered.')
add('scaleStick', 'Stick-built line-item multiplier', 1, DEC, TC, 'Set by the explorer\'s stick-built baseline list or headline $/sq ft. 1.00 = the land bank lines.')
add('solarLine', 'Solar in the base budget ($ per building, both methods)', 10000, USD, LB, NHLB + ' Replaced by the rooftop solar add-on when that is on; removed by the no-solar lever.')
add('bondPct', 'Payment and performance bond (% of local work)', 0.0091, PCT2, LB, NHLB + ' $4,708 on ~$517K of work.')
add('cfeePct', 'Contractor fee (% of local work)', 0.0835, PCT2, LB, NHLB + ' Bids ran 8.35-9.5%. The factory package sits outside the local contract.')
add('cCont', 'Contractor contingency ($ per building)', 3042, USD, LB, NHLB)
add('volume', 'Order volume (buildings ordered together)', 10, NUM, TC, 'Team choice: ~10 available land bank sites (Feedback Tracker #17).')
add('grossup', 'Gross-up (net to gross sq ft)', 1.12, DEC, PH, 'Team assumption (walls, stairs, mechanical); architecture team to confirm.')
add('speedPct', 'Modular build time vs. stick-built (% faster)', 0.30, PCT, SE, RFC + ' About 30% faster than site-built. Modular loan months = stick-built months x (1 - this).')
add('loanStick', 'Construction-loan months, stick-built (the baseline)', 10, DEC, PH, 'Schedule tab.')
add('foundType', 'Foundation type', FOUND_T[0], None, TC, RFC + ' Slab on grade is Reframe\'s choice where there is ledge. The slab is assumed inside the field cost; other types add only the difference.', FOUND_T)
add('foundSlab', 'Foundation price: insulated slab on grade (reference)', 39215, USD, LB, NHLB + ' Base-average foundation for the two-family prototype.')
add('foundCrawl', 'Foundation price: crawl space on stem walls', 45000, USD, PH, 'Placeholder until a quote.')
add('foundPiles', 'Foundation price: helical piles, vented crawl space', 35000, USD, PH, 'Placeholder until a quote.')
add('foundSoilsOn', 'Apply a soils-unknown contingency?', 'Yes', None, PH, 'On until there are soils data for the parcel (USDA Web Soil Survey, CT bedrock maps, geotech).', YN)
add('foundSoilsPct', 'Soils-unknown contingency (% of foundation price)', 0.10, PCT, PH, 'Placeholder.')
add('foundLedgeOn', 'Apply a ledge premium (crawl space or piles only)?', 'No', None, PH, 'Ignored for a slab.', YN)
add('foundLedge', 'Ledge premium', 10000, USD, PH, 'Placeholder.')
add('hxFound', 'Other line item: foundation / soils premium beyond the choice above', 0, USD, TC, '$0 until the design team has a figure.')
add('hxAppl', 'Other line item: appliance package', 0, USD, TC, '$0 until the design team has a figure.')
add('hxFin', 'Other line item: finishes upgrade', 0, USD, TC, '$0 until the design team has a figure.')
add('hxAcc', 'Other line item: accessibility / visitability', 0, USD, TC, '$0 until the design team has a figure.')
sect('4. SITE WORK AND HARD CONTINGENCY (per building)')
add('clearing', 'Site work (clearing, grading)', 15954, USD, LB, NHLB)
add('utilities', 'Street utilities and connections', 33824, USD, LB, NHLB)
add('landscape', 'Landscaping, driveway, walks', 26469, USD, LB, NHLB)
add('fence', 'Fencing and security', 15207, USD, LB, NHLB)
add('hardCont', 'Hard-cost contingency', 0.0727, PCT1, LB, NHLB + ' 7.27% (the redacted course budget used 10%).')
sect('4b. DELAYS AND CONTINGENCIES (each off unless set to Yes; storage, crane and trailer are modular only)')
add('dlyStorageOn', 'Module storage if the site is not ready?', 'No', None, PH, CHFAM + ' CHFA flagged storage.', YN)
add('dlyModules', 'Modules stored', 4, NUM, PH, 'Placeholder.')
add('dlyStorageMo', 'Storage cost per module per month', 1500, USD, PH, 'Placeholder.')
add('dlyStorageMonths', 'Months in storage', 2, DEC, PH, 'Placeholder.')
add('dlyWeatherOn', 'Weather delay?', 'No', None, PH, 'Adds construction-loan months.', YN)
add('dlyWeatherMonths', 'Weather delay (extra loan months)', 1, DEC, PH, 'Placeholder.')
add('dlyCraneOn', 'Crane re-mobilization?', 'No', None, PH, 'A second crane day after a weather or site delay.', YN)
add('dlyCraneVal', 'Crane re-mobilization cost', 8000, USD, PH, 'Placeholder.')
add('dlySoilsOn', 'Soils or ledge surprise?', 'No', None, PH, 'Found during excavation, beyond the foundation contingency.', YN)
add('dlySoilsVal', 'Soils or ledge surprise cost', 15000, USD, PH, 'Placeholder.')
add('dlyPermitOn', 'Permit delay?', 'No', None, PH, 'Adds construction-loan months.', YN)
add('dlyPermitMonths', 'Permit delay (extra loan months)', 2, DEC, PH, 'Placeholder.')
add('dlyUtilityOn', 'Utility or sewer upgrade?', 'No', None, PH, 'Beyond the standard connections.', YN)
add('dlyUtilityVal', 'Utility or sewer upgrade cost', 15000, USD, PH, 'Placeholder.')
add('dlyTrailerOn', 'Street access for a 60 ft trailer?', 'No', None, PH, "Kate's DOH example workbook / Shoreline quote flagged this.", YN)
add('dlyTrailerVal', 'Trailer access cost (police detail, trimming, closure)', 3000, USD, PH, 'Placeholder.')
add('dlyLogisticsOn', 'Extra transportation and logistics (modular)?', 'No', None, PH, SOMN + ' Oversize-load permits, escorts, route survey beyond the delivery in the factory price.', YN)
add('dlyLogisticsVal', 'Extra transportation and logistics cost', 5000, USD, PH, 'Placeholder.')
sect('5. SOFT COSTS, LOAN, LAND AND FEE (per building)')
add('architect', 'Architect', 12000, USD, CB, INFILL + ' $6,000 per unit.')
add('designRepeat', 'Repeated design: share of architect + engineering paid once per design (spread over the order)', 0, PCT, PH, '0% by default; a pattern design reused across the order pays the one-time share once.')
add('engineering', 'Engineering, stick-built (civil, structural, MEP, geotech)', 10000, USD, CB, INFILL + ' $5,000 per unit.')
add('engMod', 'Engineering, volumetric', 0, USD, DX, "$0: the factory includes engineering (Kate's DOH example).")
add('surveys', 'Surveys', 3600, USD, LB, NHLB + ' Land bank actual: $3,600 per site.')
add('phase1', 'Environmental: Phase I ESA', 1900, USD, LB, NHLB + ' Land bank actual: $1,900.')
add('phase2On', 'Include Phase II testing / remediation?', 'Yes', None, LB, NHLB + ' A limited Phase II, pending a new bid.', YN)
add('phase2Val', 'Phase II testing / remediation ($ per building)', 8700, USD, LB, NHLB + ' Land bank actual: $8,700.')
add('permits', 'Building permits and development fees', 15000, USD, CB, INFILL + ' Confirm with each city.')
add('waiverOn', 'Apply a city permit / fee waiver?', 'No', None, PH, 'No city program found yet; ask each city.', YN)
add('waiverVal', 'City permit / fee waiver ($ per building)', 8000, USD, PH, 'Rough guess for testing only.')
add('risk', "Builder's risk insurance", 2310, USD, LB, NHLB + ' The payment and performance bond is in the hard costs.')
add('legal', 'Legal: real estate, closing, title', 7300, USD, CB, HOB + ' $5,000 legal + $2,300 lender counsel.')
add('apprInsp', 'Appraisal and construction inspections', 3450, USD, CB, HOB)
add('marketing', 'Marketing and buyer intake', 2000, USD, CB, HOB)
add('commPct', 'Real estate commission (% of the sales price)', 0.03, PCT2, LB, NHLB + ' About 3% of the $399K price; 3-4% is typical.')
add('coord', 'Project coordination / predevelopment staff', 8000, USD, CB, HOB)
add('softCont', 'Soft-cost contingency', 0.05, PCT1, CB, HOB + ' 5%.')
add('carry', 'Property tax and insurance during construction', 1500, USD, PH, 'Team assumption.')
add('loanRate', 'Construction loan rate', 0.08, PCT2, SE, 'WSJ prime 7.00% (Bankrate, Oct 27, 2026) + ~1% construction spread: ' + U['prime'] + ' A redacted bridge loan was 6.5%.')
add('draw', 'Average share of costs drawn during construction', 0.55, PCT, PH, 'Typical S-curve draw (team assumption). Financing is a soft cost: interest only on the balance drawn, growing with time (' + SOMN + ')')
add('fabOn', 'Volumetric: factory deposit to hold a production slot?', 'Yes', None, PH, SOMN + ' A factory may want about 40% up front to get on the line; ask how they take payment.', YN)
add('fabDep', 'Factory deposit (% of factory price)', 0.40, PCT, PH, SOMN + ' Placeholder.')
add('fabQueue', 'Months from deposit to the start of production', 1.5, DEC, PH, 'Placeholder.')
add('fabLine', 'Months on the production line', 1.5, DEC, PH, 'Placeholder. The deposit carries loan interest for queue + line months.')
add('land', 'Land acquisition and closing ($ per building)', 1000, USD, PH, 'Land banks convey lots for $1 (Hartford: ' + U['hfdlots'] + '); closing cost is a placeholder.')
add('feeBasis', 'Developer fee is a share of', FEEB[0], None, TC, SOMN + ' Industry standard ~15% of TDC or hard costs; the New Haven Land Bank takes 15% of sales proceeds in lieu of a fee.', FEEB)
add('fee', 'Developer fee (% of the basis above; range 6-15%)', 0.10, PCT1, CB, 'Kept inside TDC. The New Haven Land Bank uses 8% and recommends a 6-8% cap; CHFA said fees typically run 10-15% (Oct 30). DOH allows up to 22% of non-acquisition cost: ' + U['doh'])
sect('5b. COST-SAVING LEVERS FROM THE LAND BANK (each off unless set to Yes)')
add('levSolar', 'Lever: no solar in the base budget?', 'No', None, LB, NHLB + ' -$10,000.', YN)
add('levFence', 'Lever: no fencing and security?', 'No', None, LB, NHLB + ' -$15,207.', YN)
add('levPatternOn', 'Lever: pattern book (no custom architect)?', 'No', None, LB, NHLB, YN)
add('levPattern', 'Pattern book: architect fee saved', 10000, USD, LB, NHLB + ' -$10,000.')
add('levZoningOn', 'Lever: as-of-right zoning (less legal)?', 'No', None, LB, NHLB, YN)
add('levZoning', 'As-of-right zoning: legal saved', 15000, USD, LB, NHLB + ' -$15,000; legal cannot go below $0.')
add('levPermitOn', 'Lever: as-of-right permits and fee discount?', 'No', None, LB, NHLB, YN)
add('levPermit', 'As-of-right permits: permit cost saved', 6000, USD, LB, NHLB + ' -$6,000.')
add('levCdfiOn', 'Lever: CDFI or foundation bridge financing?', 'No', None, LB, NHLB, YN)
add('levCdfi', 'CDFI bridge financing: interest saved (capped at the interest)', 20000, USD, LB, NHLB + ' -$20,000.')
add('levRealtorOn', 'Lever: realtor discount on the commission?', 'No', None, PH, 'Land bank webinar (no amount given).', YN)
add('levRealtor', 'Realtor discount: commission rate', 0.015, PCT2, PH, 'Placeholder.')
add('levFeeOn', 'Lever: cap the developer fee?', 'No', None, LB, NHLB, YN)
add('levFee', 'Developer fee cap', 0.08, PCT1, LB, NHLB + ' 8%; it recommends 6-8%.')
add('levLand', 'Lever: no acquisition cost?', 'No', None, LB, NHLB + ' -$3,333.', YN)
sect('6. DESIGN ADD-ONS AND ENERGY FEATURES (each off unless set to Yes)')
add('roofOn', 'Separate roof module (vs. the flat roof built into the top module)?', 'No', None, TC, RFC, YN)
add('roofCost', 'Roof module: added cost per building', 12000, USD, PH, 'Placeholder until Reframe prices it.')
add('roofVal', 'Roof module: added appraised value', 0, USD, PH, '$0 by default.')
add('garOn', 'Site-built garage?', 'No', None, TC, 'Off by default.', YN)
add('garCost', 'Garage: cost per building', 35000, USD, PH, 'Placeholder for a detached one-car garage.')
add('garVal', 'Garage: added appraised value', 0, USD, PH, '$0 by default.')
add('bikeOn', 'Bike / e-bike storage?', 'No', None, TC, 'A room in the building or a standalone shed with charging.', YN)
add('bikeCost', 'Bike storage: cost per building', 6000, USD, PH, 'Placeholder.')
add('bikeVal', 'Bike storage: added appraised value', 0, USD, PH, '$0 by default.')
add('solarOn', 'Rooftop solar: include?', 'No', None, TC, 'Off by default.', YN)
add('solarKw', 'Solar: system size (kW)', 6, DEC, TC, 'What fits on a duplex roof; design team to confirm.')
add('solarCostW', 'Solar: installed cost ($/W)', 2.68, USD2, SE, 'EnergySage CT average before incentives (2026): ' + U['escost'])
add('solarRebateW', 'Solar: upfront rebate ($/W)', 0, USD2, PH, 'No statewide upfront rebate found; RRES pays over time instead.')
add('solarCredit', 'Solar: federal tax credit (share of cost)', 0, PCT, PL, 'Homeowner credit (25D) ended Dec 31, 2025; 30% (48E) only for third-party-owned systems placed in service before 2028: ' + U['es25d'])
add('solarYield', 'Solar: output (kWh per kW per year)', 1150, NUM, PH, 'Typical CT output; check each site in NREL PVWatts.')
add('solarRate', 'Solar: value per kWh (retail rate under RRES Netting)', 0.29, '\\$0.000', SE, 'EnergySage CT average ~29 cents/kWh: ' + U['esrate'] + ' RRES: ' + U['rres'])
add('solarAdder', 'Solar: RRES adder per kWh', 0.0175, '\\$0.0000', PL, 'RRES distressed-community adder under Netting: ' + U['rres'])
add('solarShare', 'Solar: share of value to the owner', 1, PCT, TC, 'Team choice.')
add('solarOm', 'Solar: upkeep and inverter reserve ($/yr)', 150, USD, PH, 'Team assumption.')
add('solarEsc', 'Solar: electricity price rise per year (25-yr view only)', 0.03, PCT1, PH, 'Team assumption.')
add('solarDegr', 'Solar: panel output loss per year (25-yr view only)', 0.005, PCT1, PH, 'Typical panel degradation.')
add('solarVal', 'Solar: added appraised value', 0, USD, PH, '$0 by default; appraisers may not credit owned solar.')
add('batOn', 'Battery storage: include?', 'No', None, TC, 'Off by default.', YN)
add('batKwh', 'Battery: size (kWh)', 10, NUM, TC, 'Team choice.')
add('batCost', 'Battery: installed cost per kWh', 1100, USD, PH, 'Placeholder; get quotes.')
add('batInc', 'Battery: Energy Storage Solutions incentive per kWh', 200, USD, SE, 'CT Energy Storage Solutions launch rate: ' + U['ess'])
add('batCap', 'Battery: incentive cap per project', 7500, USD, PL, U['ess'])
add('batBenefit', 'Battery: owner benefit per year', 300, USD, PH, 'Placeholder.')
add('batVal', 'Battery: added appraised value', 0, USD, PH, '$0 by default.')
add('effOn', 'Efficiency / heat-pump package: include?', 'No', None, TC, 'Off by default.', YN)
add('effCost', 'Efficiency: added cost per building', 15000, USD, PH, 'Placeholder; design-team estimate needed.')
add('effRebate', 'Efficiency: rebates per building', 5000, USD, PH, 'Energize CT incentives; confirm current amounts.')
add('effSave', 'Efficiency: cut in owner energy bill', 0.25, PCT, PH, 'Placeholder.')
add('effVal', 'Efficiency: added appraised value', 0, USD, PH, '$0 by default.')
for x in 'AB':
    add('cust%sOn' % x, 'Custom add-on %s: include?' % x, 'No', None, TC, 'A blank slot for any feature.', YN)
    add('cust%sCost' % x, 'Custom add-on %s: cost per building' % x, 5000, USD, TC, 'Team choice.')
    add('cust%sRebate' % x, 'Custom add-on %s: rebate or grant' % x, 0, USD, TC, 'Team choice.')
    add('cust%sSave' % x, 'Custom add-on %s: owner savings per month' % x, 0, USD, TC, 'Team choice.')
    add('cust%sVal' % x, 'Custom add-on %s: added appraised value' % x, 0, USD, TC, 'Team choice.')
add('lenderCounts', 'Lender counts add-on savings when qualifying the buyer?', 'No', None, TC, 'Lenders usually ignore projected savings; savings still count in Test B and net cost.', YN)
sect('7. WAREHOUSE AND BULK BUYING')
add('whOn', 'Lease a warehouse and buy in bulk?', 'No', None, TC, 'Off by default.', YN)
add('whSf', 'Warehouse size (sq ft)', 5000, NUM, TC, 'Team choice.')
add('whRent', 'Warehouse rent ($ per sq ft per year)', 10, USD2, SE, 'New Haven-area industrial asking rents $5-$16.50/sf/yr (CommercialCafe, 2026): ' + U['whrent'])
add('whMonths', 'Lease length (months)', 12, NUM, TC, 'Team choice.')
add('whSetup', 'Setup, handling, insurance, trucking', 15000, USD, PH, 'Placeholder.')
add('whServed', 'Buildings sharing the warehouse', 10, NUM, TC, 'The 10-building order.')
add('whShare', 'Share of hard cost bought in bulk', 0.15, PCT, PH, 'Placeholder: field materials, appliances, fixtures, solar.')
add('whDisc', 'Bulk discount on those purchases', 0.10, PCT1, PH, 'Placeholder: needs supplier quotes.')
add('whInvMonths', 'Months materials sit in storage', 6, NUM, PH, 'Placeholder.')
add('whInvRate', 'Cost of money tied up in inventory', 0.08, PCT2, PH, 'Same as the construction-loan rate.')
sect('8. CITY DATA AND APPRAISAL')
add('chfaAmi', '100% AMI for buyers, CHFA Time To Own chart (focus city)', 123200, USD, V, 'CHFA Time To Own AMI limits, rev. 6/13/2026: New Haven $123,200, Hartford $129,200, Waterbury $114,000: https://www.chfa.org/assets/1/6/027_Time_To_Own_-_AMI__Rev._6-13-26.pdf. Treated as the 4-person figure; HUD household-size factors apply.')
add('ami50', 'HUD 50% income limit, 4-person (focus city; sets rent limits only)', 62250, USD, V, 'HUD FY2026 income limits: ' + U['hudil'] + ' (Waterbury summary confirms $62,250 / $99,600.)')
add('med2', 'Median sale, existing two-families', 345000, USD, V, OPMS)
add('med3', 'Median sale, existing three-families', 430000, USD, V, 'Same source.')
add('med46', 'Median sale, existing four-families (used for 4-6 units)', 535000, USD, V, 'Same source.')
add('psf2', '75th-percentile price per sq ft, existing two-families', 197.03, USD2, V, 'Same source: CAMA-matched living area. The comps appraisal = this x our net sq ft.')
add('psf3', '75th-percentile price per sq ft, existing three-families', 164.43, USD2, V, 'Same source.')
add('psf46', '75th-percentile price per sq ft, existing four-families', 190.42, USD2, V, 'Same source.')
add('nbName', 'Site neighborhood (focus city)', 'Citywide (no neighborhood)', None, TC, 'Set in the explorer; Waterbury uses assessor neighborhood codes.')
add('nbPsf', 'Neighborhood $/sq ft, new-build adjusted (for the neighborhood-comps basis)', 197.03, USD2, SE, "The neighborhood's median $/sq ft (two-family sales where 5+, else all 2-4 family) x the city's 75th-pct / median $/sq ft. Written by the explorer; citywide 75th pct when no neighborhood is set. " + OPMS)
add('prec', 'New-build value (appraisal precedent)', 420000, USD, CB, 'Redacted New Haven DOH Consolidated Application (2025): fair market value of a new 2,800 sq ft 2-family. Used for every city until each has its own comparable.')
add('custom', 'Size-adjusted new-build value (base case)', 340000, USD, SE, 'The $420K DOH precedent scaled to our 2,274 sq ft; in line with 2024-25 resales of new two-families in Hartford ($133-$151/sq ft: 90-92 Mather St $459,900; 197-199 Bond St $385,000), 2025 CT Parcel and CAMA data: https://catalog.data.gov/dataset/2025-connecticut-parcel-and-cama-data Replace with an appraiser\'s opinion.')
add('lb', 'Land bank market value, new two-family', 429000, USD, LB, NHLB + ' $429,000 for a new 2,400 sq ft two-family; applied to every city until each has its own.')
add('dpa', 'City down-payment assistance per buyer (focus city)', 20000, USD, LB, NHLB + ' City of New Haven page: greater of 6% of price or $20,000, 5-year forgivable, first-time low-income buyers (income cap not stated): ' + U['nhDpa'] + ' Hartford and Waterbury not researched ($0).')
add('basis', 'Appraisal basis used for the sale', B_CUST, None, TC, 'Base case: size-adjusted new-build value. The Appraisal Cases tab runs all four cases side by side.', BASIS)
add('mill', 'Mill rate (focus city)', 39.96, DEC, V, 'Confirmed with each city tax office (team, Oct 2026): New Haven 39.96 (Tax Collector Division), Hartford 69.95 (Tax Assessor), Waterbury 43.34 (Tax Assessor).')
add('assess', 'Assessment ratio, 1-3 unit homes (focus city)', 0.7, PCT1, PL, 'CGS Sec. 12-62a: 70% of fair market value (' + U['olr'] + '). Hartford uses 36.8% for 1-3 unit residential (effective tax rate ~2.57%); law team to review.')
add('assess4', 'Assessment ratio, 4+ unit buildings (focus city)', 0.7, PCT1, PL, 'Used only when the building has 4+ units. Hartford treats 4+ units differently from 1-3 units; confirm with the assessor.')
add('taxBasis', 'Property tax is assessed on', TAXB[0], None, TC, 'CT law taxes market value, not purchase price; a recorded resale cap can lower it (New London, 2026): ' + U['nlondon'], TAXB)
add('taxCustom', 'Custom assessed (market) value', 300000, USD, TC, 'Only used when the line above is set to Custom.')
add('abateOn', 'Apply a property tax abatement?', 'No', None, PH, 'No owner-occupied program found in the three cities; Hartford has used 10-year tax-fixing agreements for new homes: ' + U['hfdfix'], YN)
add('abateVal', 'Abatement (share of tax)', 0.5, PCT, PH, 'Rough guess for testing only.')
sect('9. BUYER FINANCING')
add('rate', 'Mortgage rate', 0.065, '0.000%', SE, 'CHFA government-insured 6.75% less 0.25% Targeted Area discount (CHFA, Oct 2026): ' + U['chfa'] + ' Market average 7.32% (Freddie Mac PMMS, Oct 22, 2026).')
add('term', 'Mortgage term (years)', 30, NUM, TC, 'Standard 30-year fixed.')
add('dp', 'Down payment', 0.035, PCT1, PL, 'FHA minimum for owner-occupied 1-4 unit homes (HUD Handbook 4000.1).')
add('mip', 'FHA mortgage insurance (annual)', 0.0055, PCT2, PL, 'HUD ML 2023-05: 0.55% above 95% LTV: ' + U['mip'] + ' Assumes FHA; a CHFA conventional loan would use private mortgage insurance.')
add('rentCredit', 'Lender credit on rental income', 0.75, PCT, PL, 'CT DOH outline p.9: 75% of future gross rent: ' + U['doh'])
add('lenderCap', 'Test A: lender cap, housing share of qualifying income', 0.40, PCT1, TC, 'Conservative team assumption; FHA allows 31%-46.9%.')
add('ownerCap', "Test B: owner cap, share of own income", 0.30, PCT, PL, 'HUD affordability standard: 30% of income.')
add('capBasis', "The owner's 30% covers", CAPB[0], None, TC, SOMN + ' The 30% usually applies to the mortgage alone; all-in costs (tax, insurance, utilities) run about 33-40%.', CAPB)
add('allinCap', 'All-in ceiling when the 30% covers the mortgage only', 0.38, PCT1, SE, SOMN + ' Middle of 33-40%.')
sect('10. OWNER OPERATING COSTS')
add('ins', 'Homeowner insurance ($ per unit per year, focus city)', 1818, USD, SE, 'City average premium per policy / 2 units (NerdWallet, Mar 2026): ' + U['nw'])
add('energy', 'Owner unit energy bills ($ per month, all-electric)', 312, USD, SE, 'Housing Authority of New Haven allowance, 3BR rowhouse, heat pump, incl. A/C (Jan 2024): ' + U['ecc'] + ' Replace with a design-team energy model.')
add('common', 'Owner-paid water and sewer ($ per unit per month)', 128, USD, SE, 'Same schedule: water + sewer $105 (2BR) and $151 (3BR), averaged per unit.')
add('maintPct', 'Maintenance and capital reserve (% of value per year)', 0.01, PCT1, PH, 'Rule of thumb ~1% a year.')
add('vac', 'Vacancy and collection loss', 0.07, PCT1, PH, 'Standard underwriting assumption.')
add('mitType', 'Household type for the MIT living-wage test', MIT[1][0], None, TC, 'MIT Living Wage Calculator, New Haven metro (Feb 15, 2026): ' + U['mit'], [m[0] for m in MIT])
add('energyThresh', 'High energy burden threshold', 0.06, PCT, V, 'ACEEE: above 6% of income is a high energy burden.')
sect('11. BUYER-SIDE MONEY (fills the sales gap only)')
add('incAssist', 'Include buyer assistance (Time To Own etc.)?', 'Yes', None, TC, '', YN)
add('tto', 'CHFA Time To Own, max per buyer', 18750, USD, PL, 'CHFA Loan Program Outlines rev. 11/13/2025: $25,000 at or below 80% AMI; $18,750 above 80% to 100%; none above 100%. Needs a CHFA first mortgage; CHFA FHA loans exclude 2-4 units (VA only). OPEN: which first mortgage the two-family buyer uses: ' + U['tto'])
add('empDpa', 'Employer down-payment assistance per buyer (e.g., Yale)', 0, USD, PH, 'Yale Homebuyer Program: $30,000 over 10 years ($7,500 at closing, taxable), Yale staff only: ' + U['yale'] + ' $0 by default.')
add('otherAssist', 'Other buyer assistance (e.g., FHLBank Equity Builder)', 0, USD, PH, 'Off by default; Equity Builder up to $32,099.')
add('useHB', 'Fill the sales gap with a DOH homebuyer subsidy?', 'Yes', None, CB, 'INFERENCE from the redacted New Haven DOH application. DOH rules size one gap to an affordable price (100% AMI in the 2024 draft, 80% AMI in the Spring 2026 DEP): ' + U['dep26'] + ' OPEN: does DOH fund a separate homebuyer subsidy?')
sect('12. DEVELOPER-SIDE MONEY (fills the developer gap, in this order)')
add('incHdf', 'Include HDF / CT Home Funds?', 'Yes', None, TC, '', YN)
add('hdf', 'HDF / CT Home Funds grant, per home', 150000, USD, PL, 'CT Home Funds Homeownership Construction Program (Apr 2025): grants up to $150,000 per home, up to ten homes: ' + U['hdf'] + ' OPEN: stacking with DOH; buyer income limit.')
add('incAhp', 'Include FHLBank Boston AHP?', 'No', None, TC, 'Off: AHP owner units must be at or below 80% AMI (FHFA: ' + U['fhfaAhp'] + '); needs a member bank; 2026 round closed.')
add('ahp', 'FHLBank Boston AHP, per home', 100000, USD, PL, 'Up to $100,000 per homeownership unit: ' + U['ahp'])
add('incHtcc', 'Include CHFA HTCC?', 'Yes', None, TC, '', YN)
add('htcc', 'HTCC award to a nonprofit sponsor', 500000, USD, PL, 'CHFA HTCC guidelines rev. 3/2025: up to $500,000 per applicant and per development per year; owners at or below 100% AMI: ' + U['htcc'] + ' New Haven 2-unit pilot received $500K: ' + U['pilot'])
add('htccHomes', 'Homes sharing one HTCC award', 10, NUM, TC, 'One award spread across a 10-building phase.')
add('htccMode', 'How the HTCC award is used', HTCC_T[0], None, TC, 'INFERENCE: HTCC guidelines describe revolving funds for buyer costs (down payment, buydowns, closing, rehab), not construction loans; a 2021 Capital for Change pool did fund construction loans. OPEN: confirm with CHFA.')
add('htccTurns', 'HTCC fund: times it turns over during the phase', 2, DEC, PH, 'Working estimate 2-3.')
add('htccRate', 'HTCC fund: interest it charges', 0, PCT2, PH, 'Placeholder.')
add('incDohRent', 'Include a DOH rental subsidy for the rental unit(s)?', 'No', None, TC, CHFAM + ' Pending the law team. Not counted for a condo.', YN)
add('dohRent', 'DOH rental subsidy per rental unit', 100000, USD, PH, 'Spring 2026 DEP: DOH targets $100K per assisted unit at 60-80% AMI, $150K at or below 50%: ' + U['dep26'] + ' INFERENCE: applying it to the rental unit of an owner-occupied two-family.')
add('incNmtc', 'Include New Markets Tax Credits (when switched on below)?', 'Yes', None, TC, 'Outside the base case: NMTC fits a condo scenario better. LIHTC is off the table for owner-occupied homes and not modeled.', YN)
add('nmtcOn', 'NMTC: use New Markets Tax Credits?', 'No', None, TC, 'Off by default. See the NMTC tab.', YN)
add('nmtcHomes', 'NMTC: homes in the transaction', 10, NUM, SE, 'For-sale NMTC deals typically 10-20 homes, $5-10M: ' + U['novo'])
add('nmtcQei', 'NMTC: share of TDC financed through the structure (QEI)', 1, PCT, TC, 'Team choice.')
add('nmtcCredit', 'NMTC: credit rate over 7 years', 0.39, PCT, PL, 'Set by law (5% x 3 yrs + 6% x 4 yrs): ' + U['cdfi'])
add('nmtcPrice', 'NMTC: investor price per $1 of credit', 0.80, USD2, SE, '$0.77 (Chicago Development Fund) to $0.83 (NY Fed): ' + U['nyfed'])
add('nmtcCosts', 'NMTC: fees and transaction costs (share of QEI)', 0.14, PCT1, SE, 'Backed out of the NY Fed 15-20% net-benefit range at $0.80 pricing.')
add('incDoh', 'Include DOH developer subsidy?', 'Yes', None, TC, '', YN)
add('dohUnit', 'DOH developer subsidy, per unit', 150000, USD, LB, NHLB + ' INFERENCE: $150K per unit is the land bank planning figure, not a DOH cap. DOH sizes the gap to an affordable price (100% AMI in the 2024 draft, 80% in the Spring 2026 DEP); total cost per home must be at or below the CHFA Bond Sales Price Limit: ' + U['dep26'])

# ---- write Inputs
ws = wb.create_sheet('Inputs', 0)
setup(ws, {'A': 2, 'B': 58, 'C': 22, 'D': 16, 'E': 90})
put(ws, 'B2', 'INPUTS: every assumption in the model (blue on yellow = change me)', title)
ws['F2'] = 'Buyer-side input'; ws['F2'].fill = bfill; ws['F3'] = 'Developer-side input'; ws['F3'].fill = dfill; ws['F4'] = 'Shared (design, appraisal)'; ws['F4'].fill = ifill
for a in ('F2', 'F3', 'F4'): ws[a].font = F(sz=9, color='FF0000FF')
put(ws, 'B3', 'The explorer writes its current dials here when you export. Input colors: lavender = buyer side, mint = developer side, yellow = shared. Change any yellow cell and every tab recalculates. Column D says how solid each number is; column E gives the source.', subf)
XM = {}; r = 5
dvs = {}
for e in I:
    if e[0] == '#':
        r += 1 if r > 5 else 0
        sec(ws, r, e[1], 2, 5); ws.cell(r, 3, 'Value').font = hdr; ws.cell(r, 4, 'Status').font = hdr; ws.cell(r, 5, 'Source / note').font = hdr
        r += 1; continue
    k, l, v, f, st, src, opts = e
    ws.cell(r, 2, l).font = body
    c = ws.cell(r, 3, v); c.font = inp; c.fill = sidefill(k)
    if f: c.number_format = f
    ws.cell(r, 4, st).font = F(sz=9, b=True, color={'Verified': 'FF0B6E4F', 'Sourced estimate': 'FF1F5FBF', 'Course budget': 'FF6B4FA0', 'Placeholder': 'FFB45309', 'Team choice': 'FF56636B', 'Program rule': 'FF0B6E4F', 'Quote': 'FF7B3FA6', 'Land bank actual': 'FF0B7480', 'DOH example': 'FF8A5A1F'}[st])
    n = ws.cell(r, 5, src); n.font = note; n.alignment = Alignment(wrap_text=True, vertical='top')
    if opts:
        key = '"' + ','.join(opts) + '"'
        if len(key) < 250:
            dvs.setdefault(key, []).append('C%d' % r)
    XM[k] = 'C%d' % r
    r += 1
for key, cells in dvs.items():
    dv = DataValidation(type='list', formula1=key, allow_blank=False); ws.add_data_validation(dv)
    for cc in cells: dv.add(cc)
# tables
r += 1; sec(ws, r, '13. TABLES: building line items (land bank structure)', 2, 8)
r += 1
for j, h in enumerate(['Line', 'Stick-built $ per sq ft', 'Volumetric: share a local contractor still does']):
    ws.cell(r, 2 + j, h).font = bold
vl_r0 = r + 1
for i, (k, l, psf, vol) in enumerate(VERT):
    rr = r + 1 + i; ws.cell(rr, 2, l).font = body
    c = ws.cell(rr, 3, psf); c.font = inp; c.fill = dfill; c.number_format = USD2; XM['vl_' + k] = 'C%d' % rr
    c = ws.cell(rr, 4, vol); c.font = inp; c.fill = dfill; c.number_format = PCT; XM['vv_' + k] = 'D%d' % rr
ws.cell(r + len(VERT) + 1, 2, "Stick-built $/sq ft: New Haven Land Bank base average of three bids for its 2,400 sq ft two-family (" + NHLB + ") Volumetric shares: Kate's DOH infill example mapping (Shoreline package). 0% = in the factory price; 100% = all local.").font = note
r += len(VERT) + 1
r += 2; sec(ws, r, '14. TABLES: by bedroom count', 2, 8); r += 1
for j, h in enumerate(['Bedrooms', 'Net sq ft', 'HUD rent-limit factor', 'Tenant utility allowance ($/mo)', 'Market asking rent ($/mo, focus city)', 'Voucher payment standard ($/mo, focus city; placeholder = Small Area FMR)']):
    ws.cell(r, 2 + j, h).font = bold
SQ = [480, 620, 880, 1150, 1400]; BF = [0.7, 0.75, 0.9, 1.04, 1.16]; UA = [162, 191, 252, 312, 371]; MK = [1532, 1595, 1750, 2100, 2970]; VPS = [1630, 1770, 2120, 2640, 2980]
br_r0 = r + 1
for i, b in enumerate(BRS):
    rr = r + 1 + i; ws.cell(rr, 2, b).font = body
    for j, (v, f) in enumerate(zip([SQ[i], BF[i], UA[i], MK[i], VPS[i]], [NUM, DEC, USD, USD, USD])):
        c = ws.cell(rr, 3 + j, v); c.font = inp; c.fill = dfill if j == 0 else bfill; c.number_format = f
        XM['br%d_%d' % (i, j)] = '%s%d' % (CL(3 + j), rr)
ws.cell(r + 6, 2, 'Sq ft: course calculator defaults. Rent factor: HUD 1.5 persons per bedroom. Utility allowance: Housing Authority of New Haven, all-electric rowhouse with heat pump, Jan 2024 (' + U['ecc'] + '). Market rent: Zillow (New Haven, Sept 2026) / Redfin (Hartford, Waterbury, Aug 2026).').font = note
BR_ROWS = (br_r0, br_r0 + 4)
r += 8; sec(ws, r, '15. TABLES: HUD household-size factor and MIT living wage', 2, 8); r += 1
ws.cell(r, 2, 'Household size').font = bold; ws.cell(r, 3, 'HUD factor').font = bold
hh_r0 = r + 1
for i, v in enumerate([0.7, 0.8, 0.9, 1.0, 1.08, 1.16]):
    rr = r + 1 + i; ws.cell(rr, 2, i + 1).font = body
    c = ws.cell(rr, 3, v); c.font = inp; c.fill = bfill; c.number_format = DEC; XM['hh%d' % i] = 'C%d' % rr
ws.cell(r + 7, 2, 'HUD income-limit family-size adjustments; above 6 persons the model adds 0.08 per person.').font = note
r += 9
for j, h in enumerate(['MIT household type', 'Required pre-tax income', 'Housing in MIT budget', 'Non-housing needs (incl. taxes)']):
    ws.cell(r, 2 + j, h).font = bold
mit_r0 = r + 1
for i, m in enumerate(MIT):
    rr = r + 1 + i; ws.cell(rr, 2, m[0]).font = body
    for j, v in enumerate(m[1:]):
        c = ws.cell(rr, 3 + j, v); c.font = inp; c.fill = bfill; c.number_format = USD
        XM['mit%d_%d' % (i, j)] = '%s%d' % (CL(3 + j), rr)
ws.cell(r + 5, 2, 'MIT Living Wage Calculator, New Haven metro (Feb 15, 2026): ' + U['mit']).font = note
ws.freeze_panes = 'C5'
IN = lambda k: "Inputs!$%s$%s" % (XM[k][0], XM[k][1:])

# ============================================================ ENGINE
eng = wb.create_sheet('Engine')
setup(eng, {'A': 2, 'B': 46, 'C': 8}, grid=True)
put(eng, 'B1', 'ENGINE: the whole model, one scenario per column (do not edit; change Inputs instead)', title)
VOLS = [1, 2, 5, 10, 20, 50, 100]
AMIS = [0.6, 0.8, 1.0, 1.2]
SC = []  # (id, label, overrides dict)
SC.append(('base', 'Base case (Inputs as set)', {}))
for a in AMIS: SC.append(('ami%g' % a, 'Owner at %d%% AMI' % (a * 100), {'ami': a}))
SC.append(('cama', 'Appraisal: existing-home value', {'basis': 'cama'}))
SC.append(('prec', 'Appraisal: DOH precedent', {'basis': 'prec'}))
SC.append(('cust', 'Appraisal: size-adjusted base case', {'basis': 'custom'}))
SC.append(('lbv', 'Appraisal: land bank value', {'basis': 'lb'}))
SC.append(('cmp', 'Appraisal: sales comps', {'basis': 'comps'}))
SC.append(('nbv', 'Appraisal: neighborhood comps', {'basis': 'nbhd'}))
for tb, tl in [('fmv', 'appraised value'), ('aff', 'affordable price'), ('cama', 'existing-home value')]:
    for a in AMIS: SC.append(('tx_%s_%g' % (tb, a), 'Tax on %s, %d%% AMI' % (tl, a * 100), {'taxb': tb, 'ami': a}))
for cb, cl in [('allin', '30% all-in'), ('pi', '30% mortgage only')]:
    for a in AMIS: SC.append(('cp_%s_%g' % (cb, a), '%s, %d%% AMI' % (cl, a * 100), {'capb': cb, 'ami': a}))
for fb, fl in [('pre', 'total costs'), ('hard', 'hard costs'), ('sales', 'sales price')]:
    SC.append(('fe_%s' % fb, 'Fee on %s (rate as set)' % fl, {'feeb': fb}))
    SC.append(('f15_%s' % fb, 'Fee 15%% of %s' % fl, {'feeb': fb, 'feer': 0.15}))
for m in METHODS:
    for v in VOLS: SC.append(('m_%s_%d' % (METHODS.index(m), v), '%s, %d bldgs' % (m, v), {'method': m, 'volume': v}))
SC.append(('mx_mod', 'Volumetric modular (Inputs volume)', {'method': METHODS[0]}))
SC.append(('mx_stick', 'Stick-built baseline (Inputs volume)', {'method': METHODS[1]}))
RATES = [0.055, 0.065, 0.075]; DPS = [0.035, 0.10, 0.20]; FMVS = [300000, 360000, 420000]
for a in AMIS:
    for x in RATES: SC.append(('ar_%g_%g' % (a, x), '%d%% AMI, rate %.1f%%' % (a * 100, x * 100), {'ami': a, 'rate': x}))
for a in AMIS:
    for x in DPS: SC.append(('ad_%g_%g' % (a, x), '%d%% AMI, %.1f%% down' % (a * 100, x * 100), {'ami': a, 'dp': x}))
for a in AMIS:
    for x in FMVS: SC.append(('af_%g_%d' % (a, x), '%d%% AMI, value $%dK' % (a * 100, x // 1000), {'ami': a, 'fmv': x}))
HM = [0.8, 0.9, 1.0, 1.1, 1.2]; SMM = [0.75, 1.0, 1.25]
for h in HM:
    for x in SMM: SC.append(('hs_%g_%g' % (h, x), 'hard x%.2f, soft x%.2f' % (h, x), {'hard': h, 'soft': x}))
DOHS = [100000, 115000, 150000, 200000, 325000]; HDFS = [0, 75000, 150000]
for d in DOHS:
    for x in HDFS: SC.append(('dh_%d_%d' % (d, x), 'DOH $%dK/unit, HDF $%dK' % (d // 1000, x // 1000), {'doh': d, 'hdf': x}))
SC.append(('rent0.6', 'Rental unit at 60% AMI', {'rentAmi': 0.6}))
SC.append(('lev_none', 'Levers: none on', {'lev_' + k: 'No' for k, _, _, _ in LEVK}))
for k, _, _, l in LEVK: SC.append(('lev_' + k, 'Lever on alone: ' + l, {**{'lev_' + j: 'No' for j, _, _, _ in LEVK}, 'lev_' + k: 'Yes'}))
SC.append(('lev_all', 'Levers: all on', {'lev_' + k: 'Yes' for k, _, _, _ in LEVK}))
SC.append(('ahp_on', 'AHP award counted', {'ahp': 'Yes'}))
SC.append(('htcc_rev', 'HTCC as a revolving fund', {'htcc': HTCC_T[1]}))
SC.append(('condo', 'Condo: every unit sold', {'own': OWN_T[1]}))
SC.append(('speed0.2', 'Modular 20% faster', {'speed': 0.2}))
FEATS = [('roof', 'Separate roof module'), ('garage', 'Site-built garage'), ('bike', 'Bike storage'), ('solar', 'Rooftop solar'), ('bat', 'Battery storage'), ('eff', 'Efficiency package'), ('custA', 'Custom add-on A'), ('custB', 'Custom add-on B'), ('wh', 'Warehouse and bulk buying'), ('nmtc', 'New Markets Tax Credits')]
for k, l in FEATS:
    SC.append(('on_' + k, l + ': on', {k: 'Yes'}))
    SC.append(('off_' + k, l + ': off', {k: 'No'}))
COL = {sid: CL(4 + i) for i, (sid, _, _) in enumerate(SC)}
OVK = ['ami', 'basis', 'method', 'volume', 'rate', 'dp', 'fmv', 'hard', 'soft', 'doh', 'hdf', 'rentAmi', 'ahp', 'htcc', 'own', 'speed', 'taxb', 'capb', 'feeb', 'feer'] + [k for k, _ in FEATS] + ['lev_' + k for k, _, _, _ in LEVK]
ROW = {}
er = 3
eng.cell(2, 2, 'Scenario').font = bold
for i, (sid, lab, ov) in enumerate(SC):
    c = eng.cell(2, 4 + i, lab); c.font = F(sz=8, b=True); c.alignment = Alignment(wrap_text=True, vertical='top')
    eng.column_dimensions[CL(4 + i)].width = 13
eng.row_dimensions[2].height = 40
eng.cell(er, 2, 'OVERRIDES (blank = use Inputs)').font = bold; er += 1
for k in OVK:
    ROW['ov_' + k] = er; eng.cell(er, 2, 'override: ' + k).font = note
    for i, (sid, lab, ov) in enumerate(SC):
        if k in ov:
            c = eng.cell(er, 4 + i, ov[k]); c.font = F(sz=9, color='FF0000FF')
    er += 1
LINES = []  # (name, label, formula-template, fmt)
def L(name, label, f, fmt=USD): LINES.append((name, label, f, fmt))
yn = lambda k: '(%s="Yes")' % IN(k)
def eff_yn(k, ink):  # effective Yes/No (override or input)
    return 'IF({ov_%s}="",%s,{ov_%s}="Yes")' % (k, yn(ink), k)
L('ami', 'Owner income target (% AMI)', 'IF({ov_ami}="",%s,{ov_ami})' % IN('ownerAMI'), PCT)
L('method', 'Construction method', 'IF({ov_method}="",%s,{ov_method})' % IN('method'), '@')
L('volume', 'Order volume', 'IF({ov_volume}="",%s,{ov_volume})' % IN('volume'), NUM)
L('rate', 'Mortgage rate', 'IF({ov_rate}="",%s,{ov_rate})' % IN('rate'), PCT2)
L('dp', 'Down payment', 'IF({ov_dp}="",%s,{ov_dp})' % IN('dp'), PCT1)
L('basis', 'Appraisal basis code', 'IF({ov_basis}="",IF(%s="%s","prec",IF(%s="%s","custom",IF(%s="%s","lb",IF(%s="%s","comps",IF(%s="%s","nbhd","cama"))))),{ov_basis})' % (IN('basis'), B_PREC, IN('basis'), B_CUST, IN('basis'), B_LB, IN('basis'), B_COMPS, IN('basis'), B_NBHD), '@')
L('taxb', 'Property tax basis code', 'IF({ov_taxb}="",IF(%s="%s","cama",IF(%s="%s","custom",IF(%s="%s","aff","fmv"))),{ov_taxb})' % (IN('taxBasis'), TAXB[1], IN('taxBasis'), TAXB[2], IN('taxBasis'), TAXB[3]), '@')
L('capb', 'Owner cap basis code', 'IF({ov_capb}="",IF(%s="%s","pi","allin"),{ov_capb})' % (IN('capBasis'), CAPB[1]), '@')
L('feeb', 'Developer fee basis code', 'IF({ov_feeb}="",IF(%s="%s","hard",IF(%s="%s","sales","pre")),{ov_feeb})' % (IN('feeBasis'), FEEB[1], IN('feeBasis'), FEEB[2]), '@')
L('hardM', 'Hard-cost multiplier', 'IF({ov_hard}="",1,{ov_hard})', DEC)
L('softM', 'Soft-cost multiplier', 'IF({ov_soft}="",1,{ov_soft})', DEC)
L('dohAmt', 'DOH developer subsidy per unit', 'IF({ov_doh}="",%s,{ov_doh})' % IN('dohUnit'))
L('hdfAmt', 'HDF amount', 'IF({ov_hdf}="",%s,{ov_hdf})' % IN('hdf'))
L('rentAmi', 'Rent target (% AMI)', 'IF({ov_rentAmi}="",%s,{ov_rentAmi})' % IN('rentAMI'), PCT)
L('speed', 'Modular: % faster than stick-built', 'IF({ov_speed}="",%s,{ov_speed})' % IN('speedPct'), PCT)
L('on_ahp', 'AHP counted? (1/0)', 'IF(%s,1,0)' % eff_yn('ahp', 'incAhp'), NUM)
L('htccRev', 'HTCC used as a revolving fund? (1/0)', 'IF(AND(%s,IF({ov_htcc}="",%s,{ov_htcc})="%s"),1,0)' % (yn('incHtcc'), IN('htccMode'), HTCC_T[1]), NUM)
L('condo', 'Condo: every unit sold? (1/0)', 'IF(IF({ov_own}="",%s,{ov_own})="%s",1,0)' % (IN('own'), OWN_T[1]), NUM)
for k, ink, _, _ in LEVK:
    L('lv_' + k, 'Lever ' + k + ' on? (1/0)', 'IF(%s,1,0)' % eff_yn('lev_' + k, ink), NUM)
for k, ink in [('roof', 'roofOn'), ('garage', 'garOn'), ('bike', 'bikeOn'), ('solar', 'solarOn'), ('bat', 'batOn'), ('eff', 'effOn'), ('custA', 'custAOn'), ('custB', 'custBOn'), ('wh', 'whOn'), ('nmtc', 'nmtcOn')]:
    L('on_' + k, k + ' on? (1/0)', 'IF(%s,1,0)' % eff_yn(k, ink), NUM)
BRR = 'Inputs!$B$%d:$B$%d' % BR_ROWS
def brcol(j): return 'Inputs!$%s$%d:$%s$%d' % (CL(3 + j), BR_ROWS[0], CL(3 + j), BR_ROWS[1])
L('units', 'Units', IN('units'), NUM)
L('rented', 'Rented units', '{units}-1', NUM)
L('ownerSF', 'Owner unit net sq ft', 'INDEX(%s,MATCH(%s,%s,0))' % (brcol(0), IN('ownerBR'), BRR), NUM)
L('rentSF', 'Rental unit net sq ft', 'INDEX(%s,MATCH(%s,%s,0))' % (brcol(0), IN('rentBR'), BRR), NUM)
L('gsf', 'Gross sq ft', '({ownerSF}+{rented}*{rentSF})*%s' % IN('grossup'), NUM)
L('mod', 'Modular? (1/0)', 'IF({method}="Stick-built",0,1)', NUM)
L('dbl', 'Volume doublings', 'LOG(MAX(1,{volume}),2)', DEC)
L('allin', 'Factory price covers the whole building? (1/0)', 'IF(AND({mod}=1,OR(%s="%s",%s="%s")),1,0)' % (IN('factOpt'), FACT_T[1], IN('factOpt'), FACT_T[2]), NUM)
L('factBase', 'Factory price option, $ per sq ft', 'IF(%s="%s",%s,IF(%s="%s",%s,IF(%s="%s",%s,%s)))' % (IN('factOpt'), FACT_T[1], IN('factReframe'), IN('factOpt'), FACT_T[2], IN('factFaculty'), IN('factOpt'), FACT_T[3], IN('factCustom'), IN('factShore')), USD2)
L('facPSF', 'Factory price per sq ft at this volume', 'IF({mod}=1,{factBase}*(1-%s)^{dbl}*%s*{hardM},0)' % (IN('learn'), IN('scaleMod')), USD2)
L('trans', 'Transport, crane and set outside the factory price', 'IF({mod}=1,%s,0)' % IN('factTrans'))
L('factory', 'FACTORY PRICE (volumetric)', '{facPSF}*{gsf}+{trans}')
L('vs_roof', 'Stick-built line: Roofing', '%s*{gsf}*%s*{hardM}' % (IN('vl_roof'), IN('scaleStick')))
L('v_roof', 'Line done locally: Roofing', 'IF({mod}=1,IF({allin}=1,0,%s*{gsf}*%s*%s*{hardM}),{vs_roof})' % (IN('vl_roof'), IN('vv_roof'), IN('scaleMod')))
L('vs_envelope', 'Stick-built line: Exterior envelope', '%s*{gsf}*%s*{hardM}' % (IN('vl_envelope'), IN('scaleStick')))
L('v_envelope', 'Line done locally: Exterior envelope', 'IF({mod}=1,IF({allin}=1,0,%s*{gsf}*%s*%s*{hardM}),{vs_envelope})' % (IN('vl_envelope'), IN('vv_envelope'), IN('scaleMod')))
L('vs_interior', 'Stick-built line: Interior construction', '%s*{gsf}*%s*{hardM}' % (IN('vl_interior'), IN('scaleStick')))
L('v_interior', 'Line done locally: Interior construction', 'IF({mod}=1,IF({allin}=1,0,%s*{gsf}*%s*%s*{hardM}),{vs_interior})' % (IN('vl_interior'), IN('vv_interior'), IN('scaleMod')))
L('vs_siding', 'Stick-built line: Siding', '%s*{gsf}*%s*{hardM}' % (IN('vl_siding'), IN('scaleStick')))
L('v_siding', 'Line done locally: Siding', 'IF({mod}=1,IF({allin}=1,0,%s*{gsf}*%s*%s*{hardM}),{vs_siding})' % (IN('vl_siding'), IN('vv_siding'), IN('scaleMod')))
L('vs_windows', 'Stick-built line: Windows and doors', '%s*{gsf}*%s*{hardM}' % (IN('vl_windows'), IN('scaleStick')))
L('v_windows', 'Line done locally: Windows and doors', 'IF({mod}=1,IF({allin}=1,0,%s*{gsf}*%s*%s*{hardM}),{vs_windows})' % (IN('vl_windows'), IN('vv_windows'), IN('scaleMod')))
L('vs_finishes', 'Stick-built line: Finishes', '%s*{gsf}*%s*{hardM}' % (IN('vl_finishes'), IN('scaleStick')))
L('v_finishes', 'Line done locally: Finishes', 'IF({mod}=1,IF({allin}=1,0,%s*{gsf}*%s*%s*{hardM}),{vs_finishes})' % (IN('vl_finishes'), IN('vv_finishes'), IN('scaleMod')))
L('vs_hvac', 'Stick-built line: HVAC', '%s*{gsf}*%s*{hardM}' % (IN('vl_hvac'), IN('scaleStick')))
L('v_hvac', 'Line done locally: HVAC', 'IF({mod}=1,IF({allin}=1,0,%s*{gsf}*%s*%s*{hardM}),{vs_hvac})' % (IN('vl_hvac'), IN('vv_hvac'), IN('scaleMod')))
L('vs_elec', 'Stick-built line: Electrical', '%s*{gsf}*%s*{hardM}' % (IN('vl_elec'), IN('scaleStick')))
L('v_elec', 'Line done locally: Electrical', 'IF({mod}=1,IF({allin}=1,0,%s*{gsf}*%s*%s*{hardM}),{vs_elec})' % (IN('vl_elec'), IN('vv_elec'), IN('scaleMod')))
L('vs_plumbing', 'Stick-built line: Plumbing', '%s*{gsf}*%s*{hardM}' % (IN('vl_plumbing'), IN('scaleStick')))
L('v_plumbing', 'Line done locally: Plumbing', 'IF({mod}=1,IF({allin}=1,0,%s*{gsf}*%s*%s*{hardM}),{vs_plumbing})' % (IN('vl_plumbing'), IN('vv_plumbing'), IN('scaleMod')))
L('solarLineU', 'Solar in the base budget', 'IF(OR({lv_solar}=1,{on_solar}=1),0,%s)' % IN('solarLine'))
L('vertStick', 'Stick-built building lines', '+'.join('{vs_%s}' % k for k, _, _, _ in VERT) + '+{solarLineU}')
L('vertLocal', 'Building lines done locally', '+'.join('{v_%s}' % k for k, _, _, _ in VERT) + '+{solarLineU}')
L('bldg', 'BUILDING HARD COST (factory + local lines)', '{factory}+{vertLocal}')
L('hcPSF', 'Building cost per gross sq ft', '{bldg}/{gsf}', USD2)
L('fieldPSF', 'Local building work per gross sq ft', '{vertLocal}/{gsf}', USD2)
L('fenceUsed', 'Fencing and security', 'IF({lv_fence}=1,0,%s)' % IN('fence'))
L('site', 'Site work', '%s+%s+%s+{fenceUsed}' % (IN('clearing'), IN('utilities'), IN('landscape')))
L('hx', 'Other hard-cost line items', '%s+%s+%s+%s' % (IN('hxFound'), IN('hxAppl'), IN('hxFin'), IN('hxAcc')))
L('foundPrice', 'Foundation price (chosen type)', 'IF(%s="%s",%s,IF(%s="%s",%s,%s))' % (IN('foundType'), FOUND_T[1], IN('foundCrawl'), IN('foundType'), FOUND_T[2], IN('foundPiles'), IN('foundSlab')))
L('foundPrem', 'Foundation: difference from the slab', '{foundPrice}-%s' % IN('foundSlab'))
L('foundCont', 'Foundation: soils-unknown contingency', 'IF(%s,{foundPrice}*%s,0)' % (yn('foundSoilsOn'), IN('foundSoilsPct')))
L('ledgeC', 'Foundation: ledge premium', 'IF(AND(%s,%s<>"%s"),%s,0)' % (yn('foundLedgeOn'), IN('foundType'), FOUND_T[0], IN('foundLedge')))
L('found', 'FOUNDATION (price + soils contingency + ledge)', '{foundPrice}+{foundCont}+{ledgeC}')
L('dStorage', 'Delay: module storage (modular only)', 'IF(AND({mod}=1,%s),%s*%s*%s,0)' % (yn('dlyStorageOn'), IN('dlyModules'), IN('dlyStorageMo'), IN('dlyStorageMonths')))
L('dCrane', 'Delay: crane re-mobilization (modular only)', 'IF(AND({mod}=1,%s),%s,0)' % (yn('dlyCraneOn'), IN('dlyCraneVal')))
L('dTrailer', 'Delay: 60 ft trailer street access (modular only)', 'IF(AND({mod}=1,%s),%s,0)' % (yn('dlyTrailerOn'), IN('dlyTrailerVal')))
L('dSoils', 'Delay: soils or ledge surprise', 'IF(%s,%s,0)' % (yn('dlySoilsOn'), IN('dlySoilsVal')))
L('dUtility', 'Delay: utility or sewer upgrade', 'IF(%s,%s,0)' % (yn('dlyUtilityOn'), IN('dlyUtilityVal')))
L('dLogist', 'Delay: extra transportation and logistics (modular only)', 'IF(AND({mod}=1,%s),%s,0)' % (yn('dlyLogisticsOn'), IN('dlyLogisticsVal')))
L('dlyHard', 'DELAYS AND CONTINGENCIES (hard cost)', '{dStorage}+{dCrane}+{dTrailer}+{dSoils}+{dUtility}+{dLogist}')
L('dlyMonths', 'Delay months added to the construction loan', 'IF(%s,%s,0)+IF(%s,%s,0)' % (yn('dlyWeatherOn'), IN('dlyWeatherMonths'), yn('dlyPermitOn'), IN('dlyPermitMonths')), DEC)
L('fSolarCost', 'Solar: net upfront cost', 'IF({on_solar}=1,%s*1000*(%s-%s)*(1-%s),0)' % (IN('solarKw'), IN('solarCostW'), IN('solarRebateW'), IN('solarCredit')))
L('fSolarSave', 'Solar: owner savings per month', 'IF({on_solar}=1,(%s*%s*(%s+%s)*%s-%s)/12,0)' % (IN('solarKw'), IN('solarYield'), IN('solarRate'), IN('solarAdder'), IN('solarShare'), IN('solarOm')))
L('fBatCost', 'Battery: net upfront cost', 'IF({on_bat}=1,MAX(0,%s*%s-MIN(%s*%s,%s)),0)' % (IN('batKwh'), IN('batCost'), IN('batKwh'), IN('batInc'), IN('batCap')))
L('fBatSave', 'Battery: owner savings per month', 'IF({on_bat}=1,%s/12,0)' % IN('batBenefit'))
L('fEffCost', 'Efficiency: net upfront cost', 'IF({on_eff}=1,%s-%s,0)' % (IN('effCost'), IN('effRebate')))
L('fEffSave', 'Efficiency: owner savings per month', 'IF({on_eff}=1,%s*%s,0)' % (IN('energy'), IN('effSave')))
L('fCaCost', 'Custom A: net cost', 'IF({on_custA}=1,%s-%s,0)' % (IN('custACost'), IN('custARebate')))
L('fCaSave', 'Custom A: savings per month', 'IF({on_custA}=1,%s,0)' % IN('custASave'))
L('fCbCost', 'Custom B: net cost', 'IF({on_custB}=1,%s-%s,0)' % (IN('custBCost'), IN('custBRebate')))
L('fCbSave', 'Custom B: savings per month', 'IF({on_custB}=1,%s,0)' % IN('custBSave'))
L('fRoofCost', 'Roof module: cost', 'IF({on_roof}=1,%s,0)' % IN('roofCost'))
L('fGarCost', 'Garage: cost', 'IF({on_garage}=1,%s,0)' % IN('garCost'))
L('fBikeCost', 'Bike storage: cost', 'IF({on_bike}=1,%s,0)' % IN('bikeCost'))
L('featCost', 'Add-ons: total net upfront cost', '{fSolarCost}+{fBatCost}+{fEffCost}+{fCaCost}+{fCbCost}+{fRoofCost}+{fGarCost}+{fBikeCost}')
L('featSaveMo', 'Add-ons: owner savings per month', '{fSolarSave}+{fBatSave}+{fEffSave}+{fCaSave}+{fCbSave}', USD2)
L('featVal', 'Add-ons: added appraised value', '{on_solar}*%s+{on_bat}*%s+{on_eff}*%s+{on_custA}*%s+{on_custB}*%s+{on_roof}*%s+{on_garage}*%s+{on_bike}*%s' % (IN('solarVal'), IN('batVal'), IN('effVal'), IN('custAVal'), IN('custBVal'), IN('roofVal'), IN('garVal'), IN('bikeVal')))
L('whEligible', 'Warehouse: spend eligible for bulk buying', '({bldg}+{featCost}+{hx})*%s' % IN('whShare'))
L('whFixed', 'Warehouse: rent + setup (whole lease)', '%s*%s*%s/12+%s' % (IN('whSf'), IN('whRent'), IN('whMonths'), IN('whSetup')))
L('whPerBldg', 'Warehouse: fixed cost per building', '{whFixed}/MAX(1,%s)' % IN('whServed'))
L('whSave', 'Warehouse: bulk discount per building', '{whEligible}*%s' % IN('whDisc'))
L('whCarry', 'Warehouse: inventory carrying cost', '{whEligible}*(1-%s)*%s*%s/12' % (IN('whDisc'), IN('whInvRate'), IN('whInvMonths')))
L('whBreakEven', 'Warehouse: break-even buildings', 'IF({whSave}-{whCarry}>0,{whFixed}/({whSave}-{whCarry}),"never")', DEC)
L('whNet', 'Warehouse: net cost per building (negative = saves)', 'IF({on_wh}=1,{whPerBldg}+{whCarry}-{whSave},0)')
L('localWork', 'Local contract work (building lines, site, foundation, delays)', '{vertLocal}+{site}+{found}+{dlyHard}')
L('bond', 'Payment and performance bond', '{localWork}*%s' % IN('bondPct'))
L('cCont', 'Contractor contingency', IN('cCont'))
L('cFee', 'Contractor fee', '({localWork}+{bond}+{cCont})*%s' % IN('cfeePct'))
L('hardBase', 'Hard cost before contingency', '{factory}+{localWork}+{bond}+{cCont}+{cFee}+{featCost}+{hx}+{whNet}')
L('hardContC', 'Hard-cost contingency', '{hardBase}*%s' % IN('hardCont'))
L('construction', 'CONSTRUCTION (HARD) COST', '{hardBase}+{hardContC}')
L('hardPSF', 'All-in hard cost per gross sq ft', '{construction}/{gsf}', USD2)
L('repF', 'Repeated design: share of design cost each building pays', '(1-%s)+%s/MAX(1,{volume})' % (IN('designRepeat'), IN('designRepeat')), PCT1)
L('sArch', 'Architect (less any pattern-book saving)', 'IF({lv_pattern}=1,MAX(0,%s*{softM}*{repF}-%s),%s*{softM}*{repF})' % (IN('architect'), IN('levPattern'), IN('architect')))
L('sEng', 'Engineering (volumetric: factory includes it)', 'IF({mod}=1,%s*{repF},%s*{softM}*{repF})' % (IN('engMod'), IN('engineering')))
L('sSurv', 'Surveys', '%s*{softM}' % IN('surveys'))
L('sPh1', 'Phase I ESA', '%s*{softM}' % IN('phase1'))
L('phase2', 'Phase II / remediation', 'IF(%s,%s,0)' % (yn('phase2On'), IN('phase2Val')))
L('permitsUsed', 'Permits (less any as-of-right saving)', 'IF({lv_permit}=1,MAX(0,%s*{softM}-%s),%s*{softM})' % (IN('permits'), IN('levPermit'), IN('permits')))
L('waiver', 'City fee waiver', 'IF(%s,MIN(%s*{softM},{permitsUsed}),0)' % (yn('waiverOn'), IN('waiverVal')))
L('sRisk', "Builder's risk insurance", '%s*{softM}' % IN('risk'))
L('sLegal', 'Legal (less any as-of-right zoning saving)', 'IF({lv_zoning}=1,MAX(0,%s*{softM}-%s),%s*{softM})' % (IN('legal'), IN('levZoning'), IN('legal')))
L('sAppr', 'Appraisal and inspections', '%s*{softM}' % IN('apprInsp'))
L('sMkt', 'Marketing and coordination', '(%s+%s)*{softM}' % (IN('marketing'), IN('coord')))
L('commRate', 'Real estate commission rate', 'IF({lv_realtor}=1,%s,%s)' % (IN('levRealtor'), IN('commPct')), PCT2)
L('commission', 'Real estate commission', '{salesTotal}*{commRate}')
L('softBase', 'Soft cost line items', '{sArch}+{sEng}+{sSurv}+{sPh1}+{phase2}+({permitsUsed}-{waiver})+{sRisk}+{sLegal}+{sAppr}+{sMkt}+{commission}')
L('softContC', 'Soft-cost contingency', '{softBase}*%s' % IN('softCont'))
L('soft', 'SOFT COSTS', '{softBase}+{softContC}')
L('loanMonthsStick', 'Construction-loan months, stick-built baseline', '%s+{dlyMonths}' % IN('loanStick'), DEC)
L('loanMonths', 'Construction-loan months', 'IF({mod}=1,%s*(1-{speed}),%s)+{dlyMonths}' % (IN('loanStick'), IN('loanStick')), DEC)
L('loanBase', 'Average loan balance', '({construction}+{soft})*%s' % IN('draw'))
L('htccLoan', 'HTCC revolving fund: loan per home', 'IF({htccRev}=1,MIN(%s*%s/MAX(1,%s),{loanBase}),0)' % (IN('htcc'), IN('htccTurns'), IN('htccHomes')))
L('intPerMo', 'Construction-loan interest per month', '(({loanBase}-{htccLoan})*%s+{htccLoan}*%s)/12' % (IN('loanRate'), IN('htccRate')), USD2)
L('interestRaw', 'Construction-loan interest before any CDFI saving', '{intPerMo}*{loanMonths}')
L('cdfiSave', 'CDFI bridge financing saving', 'IF({lv_cdfi}=1,MIN(%s,{interestRaw}),0)' % IN('levCdfi'))
L('depCash', 'Factory deposit: cash needed before construction (volumetric)', 'IF(AND({mod}=1,%s),{factory}*%s,0)' % (yn('fabOn'), IN('fabDep')))
L('depMonths', 'Factory deposit: months until delivery', '%s+%s' % (IN('fabQueue'), IN('fabLine')), DEC)
L('depInt', 'Factory deposit: interest until delivery', '{depCash}*%s*{depMonths}/12' % IN('loanRate'))
L('interest', 'FINANCING: construction-loan interest (a soft cost)', '{interestRaw}-{cdfiSave}+{depInt}')
L('htccIntSave', 'Interest saved by the HTCC fund', '{htccLoan}*(%s-%s)*{loanMonths}/12' % (IN('loanRate'), IN('htccRate')))
L('speedSave', 'Interest saved by building faster than stick-built', 'IF({mod}=1,{intPerMo}*({loanMonthsStick}-{loanMonths}),0)')
L('stickBldg', 'Stick-built building lines at this size', '{vertStick}')
L('modPremium', 'Factory premium vs. stick-built (building only)', 'IF({mod}=1,{bldg}-{stickBldg},0)')
L('carry', 'Taxes and insurance during construction', IN('carry'))
L('land', 'Land and closing', 'IF({lv_land}=1,0,%s)' % IN('land'))
L('preFee', 'Total before developer fee', '{construction}+{soft}+{interest}+{carry}+{land}')
L('feeRate', 'Developer fee rate', 'IF({lv_fee}=1,%s,IF({ov_feer}="",%s,{ov_feer}))' % (IN('levFee'), IN('fee')), PCT1)
L('feeBase', 'Developer fee basis ($)', 'IF({feeb}="hard",{construction},IF({feeb}="sales",{salesTotal},{preFee}))')
L('feeC', 'Developer fee', '{feeBase}*{feeRate}')
L('finance', 'Financing: loan interest + carry', '{interest}+{carry}')
L('softAll', 'SOFT COSTS INCLUDING FINANCING', '{soft}+{finance}')
L('tdc', 'TOTAL DEVELOPMENT COST (TDC)', '{preFee}+{feeC}')
L('tdcUnit', 'TDC per unit', '{tdc}/{units}')
L('tdcPSF', 'TDC per gross sq ft', '{tdc}/{gsf}', USD2)
L('ami100', '100% AMI for rent limits (2 x HUD 50% limit)', '2*%s' % IN('ami50'))
L('amiBuy', '100% AMI for buyers (CHFA chart)', IN('chfaAmi'))
hhr = 'Inputs!$C$%d:$C$%d' % (hh_r0, hh_r0 + 5)
L('hhN', 'Household size (rounded)', 'MAX(1,ROUND(%s,0))' % IN('hhSize'), NUM)
L('hhF', 'HUD household-size factor', 'IF({hhN}<=6,INDEX(%s,{hhN}),INDEX(%s,6)+0.08*({hhN}-6))' % (hhr, hhr), DEC)
L('income', 'Owner household income', '{amiBuy}*{hhF}*{ami}')
L('brF', 'Rent-limit factor (rental BR)', 'INDEX(%s,MATCH(%s,%s,0))' % (brcol(1), IN('rentBR'), BRR), DEC)
L('ua', 'Tenant utility allowance', 'INDEX(%s,MATCH(%s,%s,0))' % (brcol(2), IN('rentBR'), BRR))
L('market', 'Market asking rent', 'INDEX(%s,MATCH(%s,%s,0))' % (brcol(3), IN('rentBR'), BRR))
L('condoNR', 'Condo: other-unit buyers', 'IF({condo}=1,{rented},0)', NUM)
L('rentNet', 'AMI rent limit, net of utility allowance', '{ami100}*{brF}*{rentAmi}*%s/12-{ua}' % IN('rentShare'), USD2)
L('vps', 'Voucher payment standard (gross)', 'INDEX(%s,MATCH(%s,%s,0))*%s' % (brcol(4), IN('rentBR'), BRR, IN('vpsPct')), USD2)
L('rentVoucher', 'Voucher rent, net of utility allowance', '{vps}-{ua}', USD2)
L('rBase', 'Rent before the market cap', 'IF(%s,{rentVoucher},{rentNet})' % yn('voucherOn'), USD2)
L('rent', 'Rent charged', 'IF(AND({rentAmi}=0,NOT(%s)),{market},IF(%s,MIN({rBase},{market}),{rBase}))' % (yn('voucherOn'), yn('capMkt')), USD2)
L('grossRent', 'Gross rent (all rental units; 0 for a condo)', 'IF({condo}=1,0,{rent}*{rented})', USD2)
L('collected', 'Rent collected after vacancy', '{grossRent}*(1-%s)' % IN('vac'), USD2)
L('lenderRentM', 'Rent a lender counts', '{grossRent}*%s' % IN('rentCredit'), USD2)
L('med', 'Existing-home median for this building size', 'IF({units}<=2,%s,IF({units}=3,%s,%s))' % (IN('med2'), IN('med3'), IN('med46')))
L('nbVal', 'Neighborhood comps value: neighborhood $/sq ft x net sq ft', '%s*{tsf}' % IN('nbPsf'))
L('compsVal', 'Sales comps value: 75th-pct $/sq ft x net sq ft', 'IF({units}<=2,%s,IF({units}=3,%s,%s))*{tsf}' % (IN('psf2'), IN('psf3'), IN('psf46')))
L('fmvBase', 'Appraised value before add-ons', 'IF({ov_fmv}<>"",{ov_fmv},IF({basis}="prec",%s,IF({basis}="custom",%s,IF({basis}="lb",%s,IF({basis}="comps",{compsVal},IF({basis}="nbhd",{nbVal},{med}))))))' % (IN('prec'), IN('custom'), IN('lb')))
L('fmv', 'APPRAISED VALUE (FMV)', '{fmvBase}+{featVal}')
L('salesTotal', 'GROSS SALES PRICE (all units)', 'IF({condo}=1,{fmv}*%s,{fmv})' % IN('condoVal'))
L('tsf', 'Net sq ft, all units', '{ownerSF}+{rented}*{rentSF}', NUM)
L('shO', 'Owner unit share of the sales price', 'IF({condo}=1,{ownerSF}/{tsf},1)', PCT1)
L('shR', 'Each other unit share (condo)', 'IF({condoNR}>0,{rentSF}/{tsf},0)', PCT1)
L('unitValue', 'Price of the owner unit (the whole building if not a condo)', '{salesTotal}*{shO}')
L('tax1', 'Affordable-price case, pass 1: tax on the appraised value (per month)', '{salesTotal}*%s*IF({units}>=4,%s,%s)/1000/12*{shO}*(1-{abatePct})' % (IN('mill'), IN('assess4'), IN('assess')), USD2)
L('og1', 'Pass 1: non-mortgage costs before savings', '{tax1}+{insM}+{commonM}+{energyM}+{maintM}', USD2)
L('o1', 'Pass 1: non-mortgage costs after savings', '{og1}-{featSaveOwn}', USD2)
L('oL1', 'Pass 1: costs a lender uses', 'IF(%s,{o1},{og1})' % yn('lenderCounts'), USD2)
L('a1', 'Pass 1: Test A', '(({income}/12+{lenderRentM})*%s-{oL1})/{k}/(1-{dp})' % IN('lenderCap'))
L('b1', 'Pass 1: Test B', 'IF({capb}="pi",MIN(({income}/12*%s+{collected})/{k}/(1-{dp}),({income}/12*%s+{collected}-{o1})/{k}/(1-{dp})),({income}/12*%s+{collected}-{o1})/{k}/(1-{dp}))' % (IN('ownerCap'), IN('allinCap'), IN('ownerCap')))
L('cap1', 'Pass 1: buyer capacity', 'MAX(0,MIN({a1},{b1}))')
L('tto1', 'Pass 1: buyer assistance', 'IF(%s,MIN({assistCap},MAX(0,{unitValue}-{cap1})),0)' % yn('incAssist'))
L('afford1', 'Affordable-price case: what the owner-buyer can pay (pass 1)', 'IF({taxb}="aff",{cap1}+{tto1},0)')
L('taxValue', 'Value property tax is assessed on', 'IF({taxb}="cama",{med},IF({taxb}="custom",%s,IF({taxb}="aff",MIN({salesTotal},{afford1}/{shO}),{salesTotal})))' % IN('taxCustom'))
L('taxBldgFull', 'Property tax on the building before abatement (per month)', '{taxValue}*%s*IF({units}>=4,%s,%s)/1000/12' % (IN('mill'), IN('assess4'), IN('assess')), USD2)
L('abatePct', 'Abatement share', 'IF(%s,%s,0)' % (yn('abateOn'), IN('abateVal')), PCT)
L('taxFull', "Owner's property tax before abatement (per month)", '{taxBldgFull}*{shO}', USD2)
L('tax', "Owner's property tax (per month)", '{taxFull}*(1-{abatePct})', USD2)
L('insM', 'Insurance (per month)', 'IF({condo}=1,%s/12,%s*{units}/12)' % (IN('ins'), IN('ins')), USD2)
L('commonM', 'Water and sewer, plus condo fee in a condo (per month)', 'IF({condo}=1,%s+%s,%s*{units})' % (IN('common'), IN('condoFee'), IN('common')), USD2)
L('energyM', 'Owner energy (per month)', IN('energy'), USD2)
L('maintM', 'Maintenance reserve (per month)', '%s*{unitValue}/12' % IN('maintPct'), USD2)
L('opexGross', 'Non-mortgage costs before add-on savings', '{tax}+{insM}+{commonM}+{energyM}+{maintM}', USD2)
L('featSaveOwn', "Owner's share of add-on savings (per month)", '{featSaveMo}*{shO}', USD2)
L('opex', 'Non-mortgage costs after add-on savings', '{opexGross}-{featSaveOwn}', USD2)
L('opexL', 'Non-mortgage costs a lender uses', 'IF(%s,{opex},{opexGross})' % yn('lenderCounts'), USD2)
L('energyNet', 'Owner energy after add-on savings', 'MAX(0,%s-{featSaveOwn})' % IN('energy'), USD2)
L('k', 'Monthly payment per $1 borrowed (P&I + MIP)', 'IF({rate}=0,1/(%s*12),({rate}/12)/(1-(1+{rate}/12)^(-%s*12)))+%s/12' % (IN('term'), IN('term'), IN('mip')), '0.000000')
L('testA', 'Test A: lender max price', '(({income}/12+{lenderRentM})*%s-{opexL})/{k}/(1-{dp})' % IN('lenderCap'))
L('testBpi', 'Test B (mortgage-only cap): mortgage limit', '({income}/12*%s+{collected})/{k}/(1-{dp})' % IN('ownerCap'))
L('testBall', 'Test B (mortgage-only cap): all-in ceiling', '({income}/12*%s+{collected}-{opex})/{k}/(1-{dp})' % IN('allinCap'))
L('testB', 'Test B: owner max price', 'IF({capb}="pi",MIN({testBpi},{testBall}),({income}/12*%s+{collected}-{opex})/{k}/(1-{dp}))' % IN('ownerCap'))
L('capacity', 'BUYER CAPACITY (lower of the tests)', 'MAX(0,MIN({testA},{testB}))')
L('assistCap', 'Buyer assistance available per buyer (TTO + other + City DPA + employer DPA)', '%s+%s+%s+%s' % (IN('tto'), IN('otherAssist'), IN('dpa'), IN('empDpa')))
L('tto', 'Buyer assistance used', 'IF(%s,MIN({assistCap},MAX(0,{unitValue}-{capacity})),0)' % yn('incAssist'))
L('afford', 'Most the buyer can pay', '{capacity}+{tto}')
L('salesGap', 'SALES GAP (price - affordable price)', 'MAX(0,{unitValue}-{afford})')
L('hb', 'DOH homebuyer subsidy', 'IF(%s,{salesGap},0)' % yn('useHB'))
L('financed', "Buyer's own financed price", 'MIN({unitValue},{afford})-{tto}')
L('incomeR', 'Condo: other-unit buyer income', 'IF({condoNR}>0,{amiBuy}*{brF}*{ami},0)')
L('unitValueR', 'Condo: price of each other unit', '{salesTotal}*{shR}')
L('taxR', 'Condo: other-unit property tax (per month)', '{taxBldgFull}*{shR}*(1-{abatePct})', USD2)
L('energyR', 'Condo: other-unit energy (per month)', '%s*IF({ownerSF}>0,{rentSF}/{ownerSF},1)' % IN('energy'), USD2)
L('opexGrossR', 'Condo: other-unit costs before savings', 'IF({condoNR}>0,{taxR}+%s/12+%s+%s+{energyR}+%s*{unitValueR}/12,0)' % (IN('ins'), IN('common'), IN('condoFee'), IN('maintPct')), USD2)
L('opexR', 'Condo: other-unit costs after savings', '{opexGrossR}-{featSaveMo}*{shR}', USD2)
L('opexLR', 'Condo: other-unit costs a lender uses', 'IF(%s,{opexR},{opexGrossR})' % yn('lenderCounts'), USD2)
L('testAR', 'Condo: other-unit Test A', 'IF({condoNR}>0,({incomeR}/12*%s-{opexLR})/{k}/(1-{dp}),0)' % IN('lenderCap'))
L('testBR', 'Condo: other-unit Test B', 'IF({condoNR}>0,IF({capb}="pi",MIN(({incomeR}/12*%s)/{k}/(1-{dp}),({incomeR}/12*%s-{opexR})/{k}/(1-{dp})),({incomeR}/12*%s-{opexR})/{k}/(1-{dp})),0)' % (IN('ownerCap'), IN('allinCap'), IN('ownerCap')))
L('capacityR', 'Condo: other-unit buyer capacity', 'MAX(0,MIN({testAR},{testBR}))')
L('ttoR', 'Condo: other-unit buyer assistance', 'IF(AND({condoNR}>0,%s),MIN({assistCap},MAX(0,{unitValueR}-{capacityR})),0)' % yn('incAssist'))
L('affordR', 'Condo: other-unit affordable price', '{capacityR}+{ttoR}')
L('salesGapR', 'Condo: other-unit sales gap', 'MAX(0,{unitValueR}-{affordR})')
L('hbR', 'Condo: other-unit DOH homebuyer subsidy', 'IF(%s,{salesGapR},0)' % yn('useHB'))
L('financedR', 'Condo: other-unit own financed price', 'MIN({unitValueR},{affordR})-{ttoR}')
L('financedAll', 'All buyers: own mortgage + down payment', '{financed}+{condoNR}*{financedR}')
L('ttoAll', 'All buyers: Time To Own etc.', '{tto}+{condoNR}*{ttoR}')
L('hbAll', 'All buyers: DOH homebuyer subsidy', '{hb}+{condoNR}*{hbR}')
L('salesGapAll', 'All buyers: sales gap', '{salesGap}+{condoNR}*{salesGapR}')
L('buyerSubsidy', 'BUYER SUBSIDIES USED', '{ttoAll}+{hbAll}')
L('buyerRemain', 'REMAINING BUYER GAP', '{salesGapAll}-{hbAll}')
L('buyerCanPay', 'GROSS SALES PRICE THE BUYER CAN PAY', '{salesTotal}-{buyerRemain}')
L('proceeds', 'Gross sales price counted in the developer budget', '{salesTotal}')
L('devGap', 'DEVELOPER GAP (TDC - gross sales price)', 'MAX(0,{tdc}-{proceeds})')
L('nmtcQEI', 'NMTC: qualified equity investment per building', '{tdc}*%s' % IN('nmtcQei'))
L('nmtcRate', 'NMTC: net benefit rate (credit x price - costs)', '%s*%s-%s' % (IN('nmtcCredit'), IN('nmtcPrice'), IN('nmtcCosts')), PCT1)
L('nmtcDeal', 'NMTC: total deal size (QEI x homes)', '{nmtcQEI}*%s' % IN('nmtcHomes'))
L('nmtcNet', 'NMTC: net benefit available per building', 'IF(AND({on_nmtc}=1,%s),MAX(0,{nmtcQEI}*{nmtcRate}),0)' % yn('incNmtc'))
L('dohRentAvail', 'DOH rental subsidy available', 'IF(%s,%s*IF({condo}=1,0,{rented}),0)' % (yn('incDohRent'), IN('dohRent')))
L('left0', 'Gap left before HDF', '{devGap}')
L('sHdf', 'HDF / CT Home Funds used', 'MAX(0,MIN(IF(%s,{hdfAmt},0),{left0}))' % yn('incHdf'))
L('left1', 'Gap left after HDF', '{left0}-{sHdf}')
L('sAhp', 'FHLBank AHP used', 'MAX(0,MIN(IF({on_ahp}=1,%s,0),{left1}))' % IN('ahp'))
L('left2', 'Gap left after AHP', '{left1}-{sAhp}')
L('sHtcc', 'HTCC share used (grant mode)', 'MAX(0,MIN(IF(AND(%s,{htccRev}=0),%s/%s,0),{left2}))' % (yn('incHtcc'), IN('htcc'), IN('htccHomes')))
L('left3', 'Gap left after HTCC', '{left2}-{sHtcc}')
L('sDohRent', 'DOH rental subsidy used', 'MAX(0,MIN({dohRentAvail},{left3}))')
L('left3b', 'Gap left after DOH rental', '{left3}-{sDohRent}')
L('sNmtc', 'NMTC net benefit used', 'MAX(0,MIN({nmtcNet},{left3b}))')
L('left4', 'Gap left after NMTC', '{left3b}-{sNmtc}')
L('dohCapBldg', 'DOH developer subsidy ceiling for this building', '{dohAmt}*{units}')
L('sDoh', 'DOH developer subsidy used', 'MAX(0,MIN(IF(%s,{dohCapBldg},0),{left4}))' % yn('incDoh'))
L('remain', 'REMAINING DEVELOPER GAP', '{left4}-{sDoh}')
L('devSubsidy', 'Developer-side money used', '{devGap}-{remain}')
L('dohFull', 'DOH needed to close the gap alone', '{devGap}-{sHdf}-{sAhp}-{sHtcc}-{sDohRent}-{sNmtc}')
L('buyerPays', 'Buyer pays (own mortgage + down payment)', '{financed}')
L('mortgage', 'First mortgage', '{financed}*(1-{dp})')
L('down', 'Cash down payment', '{financed}*{dp}')
L('pi', 'Mortgage payment + FHA insurance (per month)', '{mortgage}*{k}', USD2)
L('ownTotal', 'Total monthly cost of the building', '{pi}+{opex}', USD2)
L('net', "OWNER'S NET MONTHLY COST", '{ownTotal}-{collected}', USD2)
L('netPct', 'Net cost as % of income', '{net}/({income}/12)', PCT1)
L('piPct', 'Mortgage payment as % of income', '{pi}/({income}/12)', PCT1)
L('tenantShare', "Share of costs paid by the tenant's rent", '{collected}/{ownTotal}', PCT1)
L('energyBurden', 'Energy burden', '{energyNet}*12/{income}', PCT1)
L('mitNeeds', 'MIT non-housing needs', 'INDEX(Inputs!$E$%d:$E$%d,MATCH(%s,Inputs!$B$%d:$B$%d,0))' % (mit_r0, mit_r0 + 3, IN('mitType'), mit_r0, mit_r0 + 3))
L('residual', 'Surplus / shortfall vs MIT budget (per year)', '{income}-{net}*12-{mitNeeds}')
L('vac3', 'Stress: 3-month vacancy (per month, averaged)', '{grossRent}*3/12', USD2)
L('tax20', 'Stress: 20% reassessment (per month)', '{taxFull}*0.2', USD2)
L('repair', 'Stress: $10K repair (per month, over a year)', '10000/12', USD2)
L('stressPct', 'Net cost % of income if all three hit', '({net}+{vac3}+{tax20}+{repair})/({income}/12)', PCT1)
L('repairMonths', 'Months of surplus to absorb a $10K repair', 'IF({residual}>0,10000/({residual}/12),"none")', DEC)
L('subNeed', 'Total subsidy across both sides (TDC - what buyers finance; reference only)', '{tdc}-{financedAll}')
er += 1
eng.cell(er, 2, 'CALCULATION').font = bold; er += 1
for name, label, f, fmt in LINES:
    ROW[name] = er; er += 1
import re
def expand(tmpl, col):
    return re.sub(r'\{(\w+)\}', lambda m: '%s%d' % (col, ROW[m.group(1)]), tmpl)
for name, label, f, fmt in LINES:
    rr = ROW[name]; c = eng.cell(rr, 2, label); c.font = bold if label.isupper() or label.startswith(('CONSTRUCTION', 'TOTAL', 'APPRAISED', 'BUYER CAP', 'SALES', 'DEVELOPER', 'REMAINING', "OWNER'S")) else body
    eng.cell(rr, 3, name).font = note
    for i, (sid, _, _) in enumerate(SC):
        cc = eng.cell(rr, 4 + i, '=' + expand(f, CL(4 + i))); cc.number_format = fmt; cc.font = F(sz=9)
eng.freeze_panes = 'D3'
def E(name, sid='base'): return "Engine!$%s$%d" % (COL[sid], ROW[name])

# ============================================================ display helpers
def line(ws, r, label, f, fmt=USD, b=False, tot=False, nt=None, ncol=7):
    ws.cell(r, 2, label).font = bold if b else body
    c = ws.cell(r, 3, f); c.number_format = fmt; c.font = linkb if b else link
    if tot: c.fill = tfill
    if nt: n = ws.cell(r, ncol, nt); n.font = note

# ---------------- Pro Forma
pf = wb.create_sheet('TDC')
setup(pf, {'A': 2, 'B': 56, 'C': 16, 'D': 12, 'E': 4, 'F': 4, 'G': 70})
put(pf, 'B2', 'TOTAL DEVELOPMENT COST (TDC): one building, line by line (live)', title)
put(pf, 'B3', '="Focus city: "&' + IN('city') + '&" · "&' + IN('method') + '&" · order of "&' + IN('volume') + '&" buildings"', subf)
r = 5; sec(pf, r, 'A. HARD CONSTRUCTION COST', 2, 7)
rows = [('Factory price per sq ft (volumetric)', 'facPSF', USD2), ('Factory price (volumetric: boxes, delivery, set, crane, roof to weathertight)', 'factory', USD), ('Building lines done by a local contractor (all of them for stick-built)', 'vertLocal', USD),
        ('Site work (site, utilities, landscaping, fencing)', 'site', USD), ('Foundation (with soils contingency and ledge)', 'found', USD), ('Delays and contingencies (if on)', 'dlyHard', USD),
        ('Payment and performance bond', 'bond', USD), ('Contractor contingency', 'cCont', USD), ('Contractor fee (local work only)', 'cFee', USD), ('Other hard-cost line items', 'hx', USD),
        ('Design add-ons and energy features (net of rebates)', 'featCost', USD), ('Warehouse / bulk buying, net', 'whNet', USD), ('Hard-cost contingency', 'hardContC', USD)]
for i, (l, k, f) in enumerate(rows): line(pf, r + 1 + i, l, '=' + E(k), f)
r += len(rows) + 1; line(pf, r, 'CONSTRUCTION COST (hard)', '=' + E('construction'), USD, True, True); line(pf, r + 1, 'Construction cost per gross sq ft', '=' + E('construction') + '/' + E('gsf'), USD2)
PF = {'construction': r}
r += 3; sec(pf, r, 'B. SOFT COSTS (FINANCING IS A SOFT COST), LAND, FEE', 2, 7)
softs = [('Architecture, engineering, surveys', '=%s+%s+%s' % (E('sArch'), E('sEng'), E('sSurv'))), ('Environmental (Phase I and Phase II)', '=%s+%s' % (E('sPh1'), E('phase2'))), ('Permits and development fees, less any waiver', '=%s-%s' % (E('permitsUsed'), E('waiver'))),
         ("Builder's risk insurance", '=' + E('sRisk')), ('Legal, title, appraisal, inspections', '=%s+%s' % (E('sLegal'), E('sAppr'))), ('Marketing, buyer intake, coordination', '=' + E('sMkt')),
         ('Real estate commission', '=' + E('commission')), ('Soft-cost contingency', '=' + E('softContC'))]
for i, (l, f) in enumerate(softs): line(pf, r + 1 + i, l, f)
r += len(softs) + 1; line(pf, r, 'Soft costs subtotal', '=' + E('soft'), USD, True); PF['soft'] = r
line(pf, r + 1, 'Construction-loan months (stick-built baseline less the modular speed-up, plus delays)', '=' + E('loanMonths'), DEC); line(pf, r + 2, 'Financing: construction-loan interest (on the drawn balance, plus any factory deposit)', '=' + E('interest'), nt='="Factory deposit "&TEXT(' + E('depCash') + ',"$#,##0")&" up front; its interest "&TEXT(' + E('depInt') + ',"$#,##0")&" is included."'); PF['interest'] = r + 2
line(pf, r + 3, 'Financing: taxes and insurance during construction', '=' + E('carry')); PF['carry'] = r + 3
line(pf, r + 4, 'Land acquisition and closing', '=' + E('land')); PF['land'] = r + 4
line(pf, r + 5, 'Total costs before developer fee', '=' + E('preFee'))
line(pf, r + 6, 'Developer fee', '=' + E('feeC'), nt='=TEXT(' + E('feeRate') + ',"0.0%")&" of "&' + IN('feeBasis')); PF['fee'] = r + 6
line(pf, r + 7, 'TOTAL DEVELOPMENT COST (TDC)', '=' + E('tdc'), USD, True, True); PF['tdc'] = r + 7
line(pf, r + 8, 'TDC per unit', '=' + E('tdcUnit')); line(pf, r + 9, 'TDC per gross sq ft', '=' + E('tdcPSF'), USD2)
line(pf, r + 10, 'Non-construction share of TDC', '=1-' + E('construction') + '/' + E('tdc'), PCT)
line(pf, r + 11, 'Soft costs including financing', '=' + E('softAll'), USD, True, nt='Soft costs + loan interest + carry (SOM meeting, Nov 2: roll the construction loan into soft costs).')
r += 13; sec(pf, r, 'C. CHECKS AGAINST BENCHMARKS', 2, 7)
line(pf, r + 1, 'Hartford partner: 2-family stick-built, all-in', 800000, USD, nt='Verbal, External Review 1 (Oct 9, 2026)')
line(pf, r + 2, 'Redacted New Haven DOH application: new 2-family (2,800 sq ft)', 650000, USD, nt='HO Subsidy Worksheet, Model 4')
line(pf, r + 3, 'Hazel Street (New Haven, 2024): state funds per two-family', 650000, USD, nt='$2.6M for four buildings: ' + U['hazel'])
line(pf, r + 4, 'Our TDC vs. Hartford $800K', '=' + E('tdc') + '/C%d-1' % (r + 1), PCT, nt='Positive = we cost more than the partner\'s stick-built figure.')
for rr in range(r + 1, r + 4): pf.cell(rr, 3).font = body

# ---------------- Gaps & Stack
gs = wb.create_sheet('Gaps & Stack')
setup(gs, {'A': 2, 'B': 58, 'C': 16, 'D': 12, 'E': 4, 'F': 4, 'G': 60})
put(gs, 'B2', 'TWO GAPS, TWO KINDS OF MONEY (live)', title)
put(gs, 'B3', 'Developer gap = TDC - gross sales price (developer-side money only, in this order: HDF, AHP, HTCC grant, DOH rental, NMTC, DOH). The buyer analysis is separate: buyer-side money fills the sales gap only, and nothing combines the two gaps. Detail: Buyer Analysis tab.', subf)
r = 5; sec(gs, r, 'A. BUYER ANALYSIS (kept apart from the developer budget)', 2, 7)
GA = [('Owner household income', 'income', USD), ('Rent charged for the rental unit', 'rent', USD2), ('Test A: lender max price', 'testA', USD), ('Test B: owner max price', 'testB', USD), ('Buyer capacity (lower of the tests)', 'capacity', USD),
      ('Buyer assistance used (Time To Own etc.)', 'ttoAll', USD), ('Most the buyer can pay', 'afford', USD), ('Gross sales price', 'salesTotal', USD), ('SALES GAP', 'salesGapAll', USD), ('DOH homebuyer subsidy', 'hbAll', USD), ('REMAINING BUYER GAP', 'buyerRemain', USD), ('Gross sales price the buyer can pay', 'buyerCanPay', USD)]
GSR = {}
for i, (l, k, f) in enumerate(GA): line(gs, r + 1 + i, l, '=' + E(k), f, b=k in ('buyerRemain',), tot=k == 'buyerRemain'); GSR[k] = r + 1 + i
r += len(GA) + 2; sec(gs, r, 'B. DEVELOPER GAP AND DEVELOPER-SIDE STACK', 2, 7)
GB = [('Total development cost', 'tdc'), ('Gross sales price', 'proceeds'), ('DEVELOPER GAP = TDC - gross sales price', 'devGap'), ('HDF / CT Home Funds', 'sHdf'), ('FHLBank Boston AHP (if counted)', 'sAhp'), ('HTCC share (grant mode)', 'sHtcc'), ('DOH rental subsidy (if counted)', 'sDohRent'), ('New Markets Tax Credits (net benefit)', 'sNmtc'),
      ('DOH developer subsidy (gap-based, up to cap)', 'sDoh'), ('REMAINING DEVELOPER GAP', 'remain'), ('DOH needed to close the gap alone', 'dohFull')]
for i, (l, k) in enumerate(GB): line(gs, r + 1 + i, l, '=' + E(k), USD, b=k in ('devGap', 'remain'), tot=k in ('devGap', 'remain')); GSR[k] = r + 1 + i
r += len(GB) + 2; sec(gs, r, 'C. SOURCES = USES (one building)', 2, 7); gs.cell(r, 3, 'Amount').font = hdr; gs.cell(r, 4, 'Side').font = hdr
SU = [('Uses: total development cost', '=' + E('tdc'), ''), ("Buyers' mortgages + down payments", '=' + E('financedAll'), 'Buyer'), ('Buyer assistance (Time To Own etc.)', '=' + E('ttoAll'), 'Buyer'),
      ('DOH homebuyer subsidy', '=' + E('hbAll'), 'Buyer'), ('Remaining buyer gap (unfunded on the buyer side)', '=' + E('buyerRemain'), 'Buyer: open'), ('HDF / CT Home Funds', '=' + E('sHdf'), 'Developer'), ('FHLBank AHP', '=' + E('sAhp'), 'Developer'), ('HTCC share', '=' + E('sHtcc'), 'Developer'), ('DOH rental subsidy', '=' + E('sDohRent'), 'Developer'),
      ('New Markets Tax Credits', '=' + E('sNmtc'), 'Developer'), ('DOH developer subsidy', '=' + E('sDoh'), 'Developer'), ('Remaining gap (unfunded)', '=' + E('remain'), 'Open')]
su0 = r + 1
for i, (l, f, side) in enumerate(SU):
    line(gs, su0 + i, l, f); gs.cell(su0 + i, 4, side).font = note
line(gs, su0 + len(SU), 'Check: sources - uses (should be $0 when the price is below TDC)', '=SUM(C%d:C%d)-C%d' % (su0 + 1, su0 + len(SU) - 1, su0))
ch = BarChart(); ch.type = 'bar'; ch.grouping = 'stacked'; ch.overlap = 100; ch.title = 'One building: how TDC is covered'; ch.height = 7; ch.width = 18
for i in range(1, len(SU)):
    ref = Reference(gs, min_col=3, min_row=su0 + i, max_row=su0 + i)
    from openpyxl.chart import Series
    sr = Series(ref, title=SU[i][0]); ch.series.append(sr)
ch.y_axis.numFmt = '$#,##0'; gs.add_chart(ch, 'B%d' % (su0 + len(SU) + 2))

# ---------------- Buyer Analysis (item 22: kept apart from the developer budget)
ba = wb.create_sheet('Buyer Analysis')
setup(ba, {'A': 2, 'B': 56, 'C': 16, 'D': 16, 'E': 16, 'F': 16, 'G': 16, 'H': 50})
put(ba, 'B2', 'BUYER ANALYSIS: what the buyer can pay, kept apart from the developer budget (live)', title)
put(ba, 'B3', 'The developer budget counts the full gross sales price. This tab works out how buyers pay it: their own mortgage and down payment, then Time To Own and the DOH homebuyer subsidy. Whatever is left is the remaining buyer gap, never added to the developer gap.', subf)
BAC = [('base', 'This scenario (Inputs)'), ('ami0.8', 'Owner at 80% AMI'), ('ami1.2', 'Owner at 120% AMI'), ('rent0.6', 'Rental unit at 60% AMI'), ('condo', 'Condo: every unit sold')]
for j, (sid, h) in enumerate(BAC):
    c = ba.cell(5, 3 + j, h); c.font = hdr; c.fill = hfill; c.alignment = WRAP
ba.cell(5, 2, 'Result').font = hdr; ba.cell(5, 2).fill = hfill; ba.row_dimensions[5].height = 30
BAR = [('#', 'A. THE HEADLINE'), ('GROSS SALES PRICE THE BUYER CAN PAY', 'buyerCanPay', USD), ('Gross sales price (all units)', 'salesTotal', USD), ('Remaining buyer gap', 'buyerRemain', USD), ('Buyer subsidies used (Time To Own + DOH homebuyer)', 'buyerSubsidy', USD),
       ('#', 'B. THE OWNER-OCCUPANT (or the owner-unit buyer in a condo)'), ('Owner income target (% AMI)', 'ami', PCT), ('Household income', 'income', USD), ('Price of the home they buy', 'unitValue', USD), ('Rent a lender counts (per month)', 'lenderRentM', USD), ('Rent collected after vacancy (per month)', 'collected', USD),
       ('Non-mortgage costs (per month)', 'opex', USD), ('Test A: lender max price', 'testA', USD), ('Test B: owner max price', 'testB', USD), ('Buyer capacity (own mortgage + down payment)', 'capacity', USD), ('Time To Own etc. used', 'tto', USD), ('Sales gap', 'salesGap', USD), ('DOH homebuyer subsidy', 'hb', USD),
       ("Owner's net monthly cost", 'net', USD), ('Net cost as % of income', 'netPct', PCT1),
       ('#', 'C. CONDO ONLY: EACH OTHER-UNIT BUYER'), ('Number of other-unit buyers', 'condoNR', NUM), ('Household income', 'incomeR', USD), ('Price of each other unit', 'unitValueR', USD), ('Non-mortgage costs incl. condo fee (per month)', 'opexR', USD), ('Test A: lender max price', 'testAR', USD), ('Test B: owner max price', 'testBR', USD),
       ('Buyer capacity', 'capacityR', USD), ('Time To Own etc. used', 'ttoR', USD), ('Sales gap', 'salesGapR', USD), ('DOH homebuyer subsidy', 'hbR', USD),
       ('#', 'D. ALL BUYERS'), ("Buyers' own mortgages + down payments", 'financedAll', USD), ('Time To Own etc.', 'ttoAll', USD), ('DOH homebuyer subsidy', 'hbAll', USD), ('Sales gap', 'salesGapAll', USD), ('REMAINING BUYER GAP', 'buyerRemain', USD)]
rr = 6
for e in BAR:
    if e[0] == '#':
        sec(ba, rr, e[1], 2, 8); rr += 1; continue
    l, k, f = e; key = l.isupper()
    ba.cell(rr, 2, l).font = bold if key else body
    for j, (sid, _) in enumerate(BAC):
        c = ba.cell(rr, 3 + j, '=' + E(k, sid)); c.number_format = f; c.font = linkb if key else link
        if key: c.fill = tfill
    rr += 1
ba.cell(rr + 1, 2, 'Condo: no landlord and no rent; each buyer qualifies on their own income. Unit prices split the building by net sq ft (Inputs: condo sales vs. appraised value); each pays the condo fee plus water and sewer. Other-unit buyers are sized to the unit\'s bedrooms (HUD factor) at the same AMI target. Time To Own is income-tiered: $25,000 at or below 80% AMI, $18,750 at 80-100% (the Inputs value applies to every column).').font = note
ba.cell(rr + 1, 2).alignment = WRAP; ba.row_dimensions[rr + 1].height = 54
ba.freeze_panes = 'C6'

# ---------------- Method Comparison (items 21-22: Kate's layout, live)
mc = wb.create_sheet('Method Comparison')
setup(mc, {'A': 2, 'B': 50, 'C': 16, 'D': 16, 'E': 16, 'F': 16, 'G': 16, 'H': 60})
put(mc, 'B2', "METHOD COMPARISON: volumetric modular vs. the stick-built baseline, line by line (live)", title)
put(mc, 'B3', "Kate's DOH infill table in the land bank's cost structure. For volumetric, each line shows what a local contractor still does; the factory price covers the rest. Same site, foundation, soft costs and sales price for both. Columns use the Engine's mx_stick and mx_mod scenarios (the Inputs order volume).", subf)
for j, h in enumerate(['Per building', 'Stick-built baseline', 'Volumetric: local contractor', 'Volumetric: factory price', 'Volumetric total', 'Volumetric - stick', 'How the cost shifts with modular']):
    c = mc.cell(5, 2 + j, h); c.font = hdr; c.fill = hfill; c.alignment = WRAP
mc.row_dimensions[5].height = 30
NOTE = {'roof': 'Shoreline shingles the roof; gutters stay local.', 'envelope': 'Insulation and waterproofing come with the boxes.', 'interior': 'Framing, drywall and paint done in the factory.', 'siding': "Shoreline's quote excludes Hardie siding.", 'windows': 'Installed in the factory.', 'finishes': "Shoreline includes cabinets and finishes.", 'hvac': "Kept local, as in Kate's example.", 'elec': 'Wired in the factory.', 'plumbing': '20% local to marry the modules.'}
ES, EM = (lambda k: E(k, 'mx_stick')), (lambda k: E(k, 'mx_mod'))
rr = 6
def mrow(l, st, loc, fac, note_='', b=False, fmt=USD):
    global rr
    mc.cell(rr, 2, l).font = bold if b else body
    for j, f in enumerate([st, loc, fac]):
        if f is None: continue
        c = mc.cell(rr, 3 + j, f); c.number_format = fmt; c.font = linkb if b else link
    c = mc.cell(rr, 6, '=IFERROR(D%d+E%d,0)' % (rr, rr) if fac is not None or loc is not None else None); c.number_format = fmt; c.font = bold if b else body
    c = mc.cell(rr, 7, '=F%d-C%d' % (rr, rr)); c.number_format = fmt; c.font = bold if b else body
    if b:
        for cc in range(2, 8): mc.cell(rr, cc).fill = tfill
    if note_: n = mc.cell(rr, 8, note_); n.font = note
    rr += 1
def msec(t):
    global rr
    sec(mc, rr, t, 2, 8); rr += 1
msec('A. BUILDING (VERTICAL CONSTRUCTION)')
for k, l, _, _ in VERT: mrow(l, '=' + ES('vs_' + k), '=' + EM('v_' + k), '=0', NOTE[k])
mrow('Factory price', '=0', '=0', '=' + EM('factory'), '="Option: "&' + IN('factOpt') + '&" at "&TEXT(' + EM('facPSF') + ',"$0.00")&"/sq ft"')
mrow('Solar (same in both)', '=' + ES('solarLineU'), '=' + EM('solarLineU'), '=0')
msec('B. SITE, FOUNDATION AND THE LOCAL CONTRACT')
for l, k in [('Site work, utilities, landscaping, fencing', 'site'), ('Foundation (with soils contingency and ledge)', 'found'), ('Delays and contingencies switched on', 'dlyHard'), ('Payment and performance bond', 'bond'), ('Contractor contingency', 'cCont'), ('Contractor fee (local work only)', 'cFee'), ('Hard-cost contingency', 'hardContC')]:
    mrow(l, '=' + ES(k), '=' + EM(k), '=0')
mrow('Add-ons, other line items, warehouse', '=%s+%s+%s' % (ES('featCost'), ES('hx'), ES('whNet')), '=%s+%s+%s' % (EM('featCost'), EM('hx'), EM('whNet')), '=0')
mrow('CONSTRUCTION (HARD) COST', '=' + ES('construction'), '=' + EM('construction') + '-' + EM('factory'), '=' + EM('factory'), b=True)
mrow('All-in hard cost per sq ft', '=' + ES('hardPSF'), None, None, fmt=USD2); mc.cell(rr - 1, 6, '=' + EM('hardPSF')).number_format = USD2
msec('C. SOFT COSTS, LOAN, LAND AND FEE')
mrow('Soft costs (volumetric: no engineering)', '=' + ES('soft'), '=' + EM('soft'), '=0')
mrow('Construction-loan interest', '=' + ES('interest'), '=' + EM('interest'), '=0', '="Loan months: "&TEXT(' + ES('loanMonths') + ',"0.0")&" stick vs. "&TEXT(' + EM('loanMonths') + ',"0.0")&" volumetric"')
mrow('Carry and land', '=%s+%s' % (ES('carry'), ES('land')), '=%s+%s' % (EM('carry'), EM('land')), '=0')
mrow('Developer fee', '=' + ES('feeC'), '=' + EM('feeC'), '=0')
mrow('TOTAL DEVELOPMENT COST', '=' + ES('tdc'), '=' + EM('tdc') + '-' + EM('factory'), '=' + EM('factory'), b=True)
mrow('TDC per sq ft', '=' + ES('tdcPSF'), None, None, fmt=USD2); mc.cell(rr - 1, 6, '=' + EM('tdcPSF')).number_format = USD2
rr += 1; msec('D. FACTORY PREMIUM VS. INTEREST SAVED')
for l, f in [('Factory premium on the building (volumetric building - stick-built lines)', '=%s-%s' % (EM('bldg'), ES('vertStick'))), ('Loan interest saved by building faster', '=' + EM('speedSave')), ('Total development cost difference', '=%s-%s' % (EM('tdc'), ES('tdc')))]:
    line(mc, rr, l, f, USD, ncol=8); rr += 1
mc.cell(rr + 1, 2, "Size-adjusted factory price: Shoreline's $501,620 package is priced per sq ft here. The land bank's own modular figure ($900,686 TDC, $375/sq ft) uses the whole package, quoted for a larger ~3,500 sq ft 3-story duplex, on its 2,400 sq ft home; ask the land bank which design Shoreline priced. Kate's DOH example: $785,000 volumetric vs. $775,500 stick-built (2,800 sq ft, hard and soft only).").font = note
mc.cell(rr + 1, 2).alignment = WRAP; mc.merge_cells('B%d:H%d' % (rr + 1, rr + 3))
mc.freeze_panes = 'C6'

# ---------------- Cost-Saving Levers (item 28)
lv = wb.create_sheet('Cost-Saving Levers')
setup(lv, {'A': 2, 'B': 46, 'C': 14, 'D': 18, 'E': 18, 'F': 60})
put(lv, 'B2', "COST-SAVING LEVERS: the land bank's measures, one at a time (live)", title)
put(lv, 'B3', 'Each row switches one measure on with the others off (Engine lev_ scenarios); the last row switches them all on. Switch them on for real on the Inputs tab, section 5b.', subf)
for j, h in enumerate(['Measure', 'On the Inputs tab', 'Cut in total development cost', 'Cut in remaining developer gap']):
    c = lv.cell(5, 2 + j, h); c.font = hdr; c.fill = hfill; c.alignment = WRAP
for i, (k, ink, vk, l) in enumerate(LEVK):
    r_ = 6 + i; lv.cell(r_, 2, l).font = body; lv.cell(r_, 3, '=' + IN(ink)).font = link
    c = lv.cell(r_, 4, '=%s-%s' % (E('tdc', 'lev_none'), E('tdc', 'lev_' + k))); c.number_format = USD; c.font = link
    c = lv.cell(r_, 5, '=%s-%s' % (E('remain', 'lev_none'), E('remain', 'lev_' + k))); c.number_format = USD; c.font = link
r_ = 6 + len(LEVK); lv.cell(r_, 2, 'ALL OF THEM TOGETHER').font = bold
c = lv.cell(r_, 4, '=%s-%s' % (E('tdc', 'lev_none'), E('tdc', 'lev_all'))); c.number_format = USD; c.font = linkb; c.fill = tfill
c = lv.cell(r_, 5, '=%s-%s' % (E('remain', 'lev_none'), E('remain', 'lev_all'))); c.number_format = USD; c.font = linkb; c.fill = tfill
lv.cell(r_ + 2, 2, "Land bank webinar: these measures cut its gap from $390K to $242K per home. Where the developer stack already closes the gap, a lever lowers TDC but not the remaining gap.").font = note

# ---------------- Homeowner (rows kept compatible with Ownership Models / Mixed Income)
ho = wb.create_sheet('Homeowner')
setup(ho, {'A': 2, 'B': 56, 'C': 15, 'D': 15, 'E': 15, 'F': 15, 'G': 50})
put(ho, 'B2', 'HOMEOWNER AFFORDABILITY (stand-alone): can the buyer sustain ownership? (live)', title)
put(ho, 'B3', 'Same building and appraised value at four income targets. Costs include tax, insurance, energy, water/sewer, a maintenance reserve, less any add-on savings.', subf)
ho.cell(5, 2, 'Owner income target (% AMI)').font = bold
AC = ['ami0.6', 'ami0.8', 'ami1', 'ami1.2']
for j, sid in enumerate(AC):
    c = ho.cell(5, 3 + j, '=' + E('ami', sid)); c.number_format = PCT; c.font = linkb
HR = {7: 'income', 8: 'capacity', 9: 'tto', 10: 'afford', 11: 'fmv', 12: 'salesGap', 13: 'financed', 14: 'mortgage', 15: 'down', 18: 'pi', 19: 'tax', 20: 'insM', 21: 'energyM', 22: 'commonM', 23: 'maintM', 24: 'ownTotal', 25: 'collected', 26: 'net', 27: 'netPct', 28: 'tenantShare', 29: 'energyBurden', 32: 'income', 33: 'mitNeeds', 34: 'residual', 37: 'vac3', 38: 'tax20', 39: 'repair', 40: 'stressPct', 41: 'repairMonths', 42: 'featSaveMo'}
HL = {7: 'Household income', 8: 'Price supported by mortgage + down payment', 9: 'Buyer assistance used', 10: 'Affordable price (with assistance)', 11: 'Appraised value (FMV)', 12: 'Sales gap (needs a buyer subsidy)', 13: "Buyer's own financed price (excl. subsidies)", 14: 'First mortgage', 15: 'Cash down payment',
      18: 'Mortgage payment + FHA insurance', 19: 'Property tax', 20: 'Insurance', 21: 'Owner unit energy bills', 22: 'Water and sewer (both units)', 23: 'Maintenance and capital reserve', 24: 'Total monthly cost (after add-on savings, row 42)', 25: 'Rent collected from the rental unit (after vacancy)', 26: "OWNER'S NET MONTHLY COST",
      27: 'Net cost as % of income', 28: "Share of total cost paid by the tenant's rent", 29: 'Energy burden (owner energy / income)', 32: 'Annual income', 33: 'MIT non-housing needs (incl. taxes)', 34: 'Surplus (+) or shortfall (-) per year', 37: 'Stress: rental unit vacant 3 months (per month)', 38: 'Stress: reassessment raises taxes 20%', 39: 'Stress: one $10,000 repair spread over a year', 40: 'Net cost % of income if all three hit at once', 41: 'Months of residual surplus to absorb a $10K repair', 42: 'Add-on feature savings per month (already subtracted in row 24)'}
sec(ho, 6, 'A. PURCHASE', 2, 7); sec(ho, 17, 'B. MONTHLY COST OF OWNING THE WHOLE BUILDING', 2, 7); sec(ho, 31, 'C. RESIDUAL INCOME (MIT living wage)', 2, 7); sec(ho, 36, 'D. STRESS TESTS', 2, 7)
for rr, k in HR.items():
    ho.cell(rr, 2, HL[rr]).font = bold if rr == 26 else body
    for j, sid in enumerate(AC):
        fm = PCT1 if k in ('netPct', 'tenantShare', 'energyBurden', 'stressPct') else (DEC if k == 'repairMonths' else USD)
        c = ho.cell(rr, 3 + j, '=' + E(k, sid)); c.number_format = fm; c.font = linkb if rr == 26 else link
        if rr == 26: c.fill = tfill
ho.cell(27, 7, 'HUD standard: housing costs at or under 30% of income.').font = note
ho.cell(29, 7, '="High energy burden above "&TEXT(' + IN('energyThresh') + ',"0%")&" of income (ACEEE)."').font = note

# ---------------- Appraisal cases
ap = wb.create_sheet('Appraisal Cases')
setup(ap, {'A': 2, 'B': 52, 'C': 20, 'D': 20, 'E': 16, 'F': 4, 'G': 60})
put(ap, 'B2', 'FOUR APPRAISAL CASES, SIDE BY SIDE (live, focus city)', title)
put(ap, 'B3', 'Low: median sale of existing homes this size (CAMA). Base: the DOH precedent scaled to our size ($340K). DOH precedent: the $420K new-build value in the redacted New Haven DOH application. Land bank: $429K for a new 2,400 sq ft two-family. Every other input as set on the Inputs tab.', subf)
for j, h in enumerate(['Result', 'Existing homes (low)', 'Base case: size-adjusted', 'DOH precedent', 'Land bank value']): ap.cell(5, 2 + j, h).font = hdr; ap.cell(5, 2 + j).fill = hfill
AR = [('Appraised value', 'fmv'), ('Most the buyer can pay', 'afford'), ('Sales gap (buyer-side)', 'salesGapAll'), ('DOH homebuyer subsidy', 'hbAll'), ('Remaining buyer gap', 'buyerRemain'), ('Gross sales price', 'proceeds'), ('Developer gap (TDC - gross sales price)', 'devGap'), ('Developer-side money used', None),
      ('Remaining developer gap', 'remain'), ('Property tax per month', 'tax'), ("Owner's net monthly cost", 'net'), ('Net cost as % of income', 'netPct')]
for i, (l, k) in enumerate(AR):
    rr = 6 + i; ap.cell(rr, 2, l).font = body
    for j, sid in enumerate(['cama', 'cust', 'prec', 'lbv']):
        f = ('=' + E('devGap', sid) + '-' + E('remain', sid)) if k is None else '=' + E(k, sid)
        c = ap.cell(rr, 3 + j, f); c.number_format = PCT1 if k == 'netPct' else USD; c.font = link
ap.cell(19, 2, 'Why it matters: a higher appraisal lowers the developer gap but raises the sales gap, property tax and the maintenance reserve. Connecticut assesses 70% of market value, not the purchase price; a recorded resale-price restriction can lower the assessed value (CT Superior Court, New London, 2026). Test it with "Property tax is assessed on" on the Inputs tab and confirm with each assessor.').font = note
ap.cell(19, 2).alignment = WRAP; ap.merge_cells('B19:G21')

# ---------------- Add-ons & Procurement
ad = wb.create_sheet('Add-ons & Procurement')
setup(ad, {'A': 2, 'B': 34, 'C': 14, 'D': 16, 'E': 16, 'F': 16, 'G': 18, 'H': 16, 'I': 50})
put(ad, 'B2', 'ADD-ON FEATURES AND PROCUREMENT: what each would do on its own (live)', title)
put(ad, 'B3', 'Each row compares the model with that one feature switched on vs. off, all else as set on Inputs. The project pays the upfront cost (raising the developer gap); the owner keeps the savings.', subf)
for j, h in enumerate(['Feature', 'Status on Inputs', 'Net upfront cost / bldg', 'Change in remaining gap', 'Owner savings / month', "Change in owner's net cost", 'Simple payback (yrs)']):
    c = ad.cell(5, 2 + j, h); c.font = hdr; c.fill = hfill; c.alignment = WRAP
ad.row_dimensions[5].height = 30
stat_in = {'roof': 'roofOn', 'garage': 'garOn', 'bike': 'bikeOn', 'solar': 'solarOn', 'bat': 'batOn', 'eff': 'effOn', 'custA': 'custAOn', 'custB': 'custBOn', 'wh': 'whOn', 'nmtc': 'nmtcOn'}
for i, (k, l) in enumerate(FEATS):
    rr = 6 + i; ad.cell(rr, 2, l).font = body
    ad.cell(rr, 3, '=' + IN(stat_in[k])).font = link
    if k == 'wh':
        cost = '=' + E('whNet', 'on_wh'); save = '=0'
    elif k == 'nmtc':
        cost = '=-' + E('sNmtc', 'on_nmtc'); save = '=0'
    elif k in ('roof', 'garage', 'bike'):
        ck = {'roof': 'fRoof', 'garage': 'fGar', 'bike': 'fBike'}[k]
        cost = '=%s-%s' % (E(ck + 'Cost', 'on_' + k), E(ck + 'Cost', 'off_' + k)); save = '=0'
    else:
        ck = {'solar': 'fSolar', 'bat': 'fBat', 'eff': 'fEff', 'custA': 'fCa', 'custB': 'fCb'}[k]
        cost = '=%s-%s' % (E(ck + 'Cost', 'on_' + k), E(ck + 'Cost', 'off_' + k)); save = '=%s-%s' % (E(ck + 'Save', 'on_' + k), E(ck + 'Save', 'off_' + k))
    for j, (f, fm) in enumerate([(cost, USD), ('=%s-%s' % (E('remain', 'on_' + k), E('remain', 'off_' + k)), USD), (save, USD2), ('=%s-%s' % (E('net', 'on_' + k), E('net', 'off_' + k)), USD2), ('=IF(F%d>0.5,D%d/(F%d*12),"-")' % (rr, rr, rr), DEC)]):
        c = ad.cell(rr, 4 + j, f); c.number_format = fm; c.font = link if j < 4 else body
ad.cell(7 + len(FEATS), 2, 'NMTC row: "upfront cost" is negative because it is a source of money; see the NMTC tab.').font = note
r = 9 + len(FEATS); sec(ad, r, 'WAREHOUSE DETAIL (per building, with the warehouse on)', 2, 9)
for i, (l, k) in enumerate([('Spend eligible for bulk buying', 'whEligible'), ('Rent + setup for the whole lease', 'whFixed'), ('Fixed cost per building', 'whPerBldg'), ('Bulk discount per building', 'whSave'), ('Inventory carrying cost', 'whCarry'), ('Net cost per building (negative = saves)', 'whNet'), ('Break-even number of buildings', 'whBreakEven')]):
    line(ad, r + 1 + i, l, '=' + E(k, 'on_wh'), DEC if k == 'whBreakEven' else USD, ncol=9)
r = 18 + len(FEATS); sec(ad, r, 'ROOFTOP SOLAR: 25-YEAR OWNER SAVINGS (with the Inputs solar settings)', 2, 9)
for j, h in enumerate(['Year', 'Output (kWh)', 'Value per kWh', 'Savings to owner']): ad.cell(r + 1, 2 + j, h).font = bold
for y in range(25):
    rr = r + 2 + y
    ad.cell(rr, 2, y + 1).font = body
    c = ad.cell(rr, 3, '=%s*%s*(1-%s)^%d' % (IN('solarKw'), IN('solarYield'), IN('solarDegr'), y)); c.number_format = NUM
    c = ad.cell(rr, 4, '=%s*(1+%s)^%d+%s' % (IN('solarRate'), IN('solarEsc'), y, IN('solarAdder'))); c.number_format = '\\$0.000'
    c = ad.cell(rr, 5, '=C%d*D%d*%s-%s' % (rr, rr, IN('solarShare'), IN('solarOm'))); c.number_format = USD
line(ad, r + 27, '25-year total owner savings', '=SUM(E%d:E%d)' % (r + 2, r + 26), USD, True)
ad.cell(r + 27, 3).value = None; c = ad.cell(r + 27, 5, '=SUM(E%d:E%d)' % (r + 2, r + 26)); c.number_format = USD; c.font = bold

# ---------------- NMTC tab
nm = wb.create_sheet('NMTC')
setup(nm, {'A': 2, 'B': 60, 'C': 18, 'D': 4, 'E': 4, 'F': 4, 'G': 70})
put(nm, 'B2', 'NEW MARKETS TAX CREDITS: how they could fill the developer gap (live)', title)
put(nm, 'B3', 'Treasury allocates credits to Community Development Entities (CDEs). The land bank or a nonprofit developer would be the QALICB that receives the investment; it does not apply to Treasury.', subf)
r = 5; sec(nm, r, 'A. TRANSACTION SIZING (one building, then the whole deal)', 2, 7)
NM = [('Total development cost per building', '=' + E('tdc', 'on_nmtc'), USD), ('Qualified equity investment (QEI) per building', '=' + E('nmtcQEI', 'on_nmtc'), USD),
      ('Homes in the transaction', '=' + IN('nmtcHomes'), NUM), ('Total QEI for the deal', '=' + E('nmtcDeal', 'on_nmtc'), USD),
      ('Credits claimed by the investor over 7 years (39% of QEI)', '=' + E('nmtcDeal', 'on_nmtc') + '*' + IN('nmtcCredit'), USD),
      ('Investor equity paid for the credits', '=C10*' + IN('nmtcPrice'), USD), ('Fees and transaction costs', '=' + E('nmtcDeal', 'on_nmtc') + '*' + IN('nmtcCosts'), USD),
      ('Net benefit to the project (whole deal)', '=C11-C12', USD), ('Net benefit per building', '=' + E('nmtcQEI', 'on_nmtc') + '*' + E('nmtcRate', 'on_nmtc'), USD),
      ('Net benefit as % of QEI', '=' + E('nmtcRate', 'on_nmtc'), PCT1)]
for i, (l, f, fm) in enumerate(NM): line(nm, 6 + i, l, f, fm, b=i in (7, 8), tot=i == 8)
nm.cell(9, 7, '=IF(C9<5000000,"Below the typical $5-10M for-sale NMTC deal: CDEs may not take it on.",IF(C9>10000000,"Larger than typical for-sale deals ($5-10M).","Within the typical $5-10M range for for-sale NMTC deals."))').font = note
r = 17; sec(nm, r, 'B. EFFECT ON THIS BUILDING (NMTC on vs. off; everything else as on Inputs)', 2, 7)
for j, h in enumerate(['', 'NMTC off', 'NMTC on']): nm.cell(r, 2 + j, h).font = hdr
for i, (l, k) in enumerate([('NMTC net benefit used', 'sNmtc'), ('DOH developer subsidy used', 'sDoh'), ('Remaining developer gap', 'remain'), ('DOH needed to close the gap alone', 'dohFull')]):
    rr = r + 1 + i; nm.cell(rr, 2, l).font = body
    for j, sid in enumerate(['off_nmtc', 'on_nmtc']): c = nm.cell(rr, 3 + j, '=' + E(k, sid)); c.number_format = USD; c.font = link
nm.column_dimensions['D'].width = 18
r = 24; sec(nm, r, 'C. HOW IT WORKS AND WHAT TO CHECK', 2, 7)
txt = ['1. The investor (usually a bank) puts equity plus a leverage loan into a CDE that holds NMTC allocation (the QEI).',
       '2. The CDE lends to the QALICB (the land bank or nonprofit developer) as an A loan (mirrors the leverage loan) and a B loan (mirrors the equity, net of fees).',
       '3. The investor claims 39% of the QEI in credits over 7 years: 5% a year for 3 years, then 6% a year for 4 years.',
       '4. After 7 years the B loan is typically forgiven or bought for a nominal price (put/call), so the project keeps ~15-20% of the QEI as subsidy.',
       '5. For-sale homes: build-and-sell in one project (10-20 homes) or in serial phases; sale proceeds and carryback notes must keep the business active for 7 years.',
       'Eligibility to check: each lot must be in a low-income census tract (poverty rate >= 20% or median family income <= 80% of area median; CDFI Fund mapping tool).',
       'Buyers: statute requires at least 20% of buyers at or below 80% AMI; CDEs usually want all buyers low-income. Resale restrictions via deed or carryback note.',
       'Stacking: no published rule found on combining NMTC with DOH, HDF, AHP or HTCC. The law team must confirm before relying on it.',
       'Program status: made permanent in 2025; $10B allocated to 142 CDEs in Dec 2025. Find CDEs active in Connecticut that fund for-sale housing.']
for i, t in enumerate(txt):
    c = nm.cell(r + 1 + i, 2, t); c.font = body if i < 5 else note; c.alignment = WRAP; nm.merge_cells(start_row=r + 1 + i, start_column=2, end_row=r + 1 + i, end_column=7); nm.row_dimensions[r + 1 + i].height = 26
srcs = ['NY Fed primer: ' + U['nyfed'], 'Chicago Development Fund: ' + U['cdf'], 'Smith NMTC (homeownership): ' + U['smith'], 'Novogradac podcast (Dec 9, 2025): ' + U['novo'], 'CDFI Fund intro: ' + U['cdfi']]
for i, t in enumerate(srcs): nm.cell(r + 11 + i, 2, t).font = note

# ---------------- Sensitivity
se = wb.create_sheet('Sensitivity')
setup(se, {'A': 2, 'B': 30, 'C': 15, 'D': 15, 'E': 15, 'F': 15, 'G': 15, 'H': 40})
put(se, 'B2', 'SENSITIVITY: change one or two drivers at a time (live; each cell is a full model run on the Engine tab)', title)
def grid(r0, ttl, rowsv, colsv, sidf, k, rowlab, collab, fmtv, rowfmt, colfmt):
    sec(se, r0, ttl, 2, 2 + len(colsv))
    se.cell(r0 + 1, 2, rowlab + ' ↓  /  ' + collab + ' →').font = bold
    for j, cv in enumerate(colsv): c = se.cell(r0 + 1, 3 + j, cv); c.number_format = colfmt; c.font = bold
    for i, rv in enumerate(rowsv):
        c = se.cell(r0 + 2 + i, 2, rv); c.number_format = rowfmt; c.font = bold
        for j, cv in enumerate(colsv):
            c = se.cell(r0 + 2 + i, 3 + j, '=' + E(k, sidf(rv, cv))); c.number_format = fmtv; c.font = link
    return r0 + 3 + len(rowsv)
r = 4
r = grid(r, "1. Owner's net cost (% of income): AMI target x mortgage rate", AMIS, RATES, lambda a, x: 'ar_%g_%g' % (a, x), 'netPct', 'AMI', 'Rate', PCT1, PCT, PCT1)
r = grid(r, '2. Sales gap per building: AMI target x mortgage rate', AMIS, RATES, lambda a, x: 'ar_%g_%g' % (a, x), 'salesGap', 'AMI', 'Rate', USD, PCT, PCT1)
r = grid(r, "3. Owner's net cost (% of income): AMI target x down payment", AMIS, DPS, lambda a, x: 'ad_%g_%g' % (a, x), 'netPct', 'AMI', 'Down payment', PCT1, PCT, PCT1)
r = grid(r, '4. Sales gap per building: AMI target x appraised value', AMIS, FMVS, lambda a, x: 'af_%g_%d' % (a, x), 'salesGap', 'AMI', 'Appraised value', USD, PCT, USD)
r = grid(r, '5. Developer gap per building: building hard cost x soft costs (multipliers)', HM, SMM, lambda h, x: 'hs_%g_%g' % (h, x), 'devGap', 'Hard x', 'Soft x', USD, DEC, DEC)
r = grid(r, '6. Remaining gap per building: DOH per unit x HDF grant', DOHS, HDFS, lambda d, x: 'dh_%d_%d' % (d, x), 'remain', 'DOH per unit', 'HDF', USD, USD, USD)
se.cell(r, 2, 'Grid 6: $150K per unit is the land bank figure; $325K per unit matches Hazel Street (~$650K per two-family). Other inputs as set on the Inputs tab.').font = note
r += 2
TXB = [('fmv', 'Appraised (market) value'), ('aff', 'Affordable price'), ('cama', 'Existing-home median')]
r = grid(r, '7. Property tax: sales gap the buyer side must fill, by what the home is assessed on x AMI (lower tax = bigger mortgage)', [t[1] for t in TXB], AMIS, lambda t, a: 'tx_%s_%g' % ([x[0] for x in TXB if x[1] == t][0], a), 'salesGapAll', 'Assessed on', 'AMI', USD, '@', PCT)
r = grid(r, '8. Property tax per month, same cases', [t[1] for t in TXB], AMIS, lambda t, a: 'tx_%s_%g' % ([x[0] for x in TXB if x[1] == t][0], a), 'tax', 'Assessed on', 'AMI', USD, '@', PCT)
sec(se, r, '8b. Policy view: buyer underwritten at the affordable-price tax, then billed on the market value', 2, 2 + len(AMIS))
se.cell(r + 1, 2, 'AMI →').font = bold
for j, a in enumerate(AMIS): c = se.cell(r + 1, 3 + j, a); c.number_format = PCT; c.font = bold
for i, (lab, fm) in enumerate([('All-in % of income as underwritten', '=%s'), ('All-in % of income once billed on market value', '=(%s+%s-%s)/(%s/12)'), ('Extra tax per month', '=%s-%s')]):
    se.cell(r + 2 + i, 2, lab).font = bold
    for j, a in enumerate(AMIS):
        ta, tf = 'tx_aff_%g' % a, 'tx_fmv_%g' % a
        f = fm % (E('netPct', ta),) if i == 0 else (fm % (E('net', ta), E('tax', tf), E('tax', ta), E('income', ta)) if i == 1 else fm % (E('tax', tf), E('tax', ta)))
        c = se.cell(r + 2 + i, 3 + j, f); c.number_format = USD if i == 2 else PCT1; c.font = link
r += 6
se.cell(r, 2, 'Grids 7-8: some programs assess at the affordable price when affordability is kept in perpetuity (SOM meeting, Nov 2); CT law otherwise taxes market value. Above 30% of income = cost-burdened (HUD); above 50% = severely.').font = note
r += 2
CPB = [('allin', '30% covers all housing costs'), ('pi', '30% covers the mortgage; all-in ceiling')]
r = grid(r, '9. Price the buyer can pay: what the owner cap covers x AMI', [t[1] for t in CPB], AMIS, lambda t, a: 'cp_%s_%g' % ([x[0] for x in CPB if x[1] == t][0], a), 'afford', 'Owner cap', 'AMI', USD, '@', PCT)
r = grid(r, "10. Owner's net cost (% of income), same cases", [t[1] for t in CPB], AMIS, lambda t, a: 'cp_%s_%g' % ([x[0] for x in CPB if x[1] == t][0], a), 'netPct', 'Owner cap', 'AMI', PCT1, '@', PCT)
r += 1; sec(se, r, '11. Developer fee basis: fee, TDC and remaining developer gap', 2, 8)
for j, h in enumerate(['Fee $', 'TDC', 'Developer gap', 'Remaining gap']): c = se.cell(r + 1, 3 + j, h); c.font = bold
FBS = [('fe_pre', 'Total costs, rate as set'), ('fe_hard', 'Hard costs, rate as set'), ('fe_sales', 'Sales price, rate as set'), ('f15_pre', '15% of total costs'), ('f15_hard', '15% of hard costs (industry standard)'), ('f15_sales', '15% of the sales price (land bank)')]
for i, (sid, lab) in enumerate(FBS):
    se.cell(r + 2 + i, 2, lab).font = bold
    for j, k in enumerate(['feeC', 'tdc', 'devGap', 'remain']): c = se.cell(r + 2 + i, 3 + j, '=' + E(k, sid)); c.number_format = USD; c.font = link
r += 9

# ---------------- Scale & Method
sm = wb.create_sheet('Scale & Method')
setup(sm, {'A': 2, 'B': 44, 'C': 13, 'D': 13, 'E': 13, 'F': 13, 'G': 13, 'H': 13, 'I': 13, 'J': 50})
put(sm, 'B2', 'VOLUMETRIC VS. STICK-BUILT: TDC per building by order volume (live)', title)
put(sm, 'B3', 'The factory price falls by its learning rate (Inputs) each time the order doubles; stick-built does not. Line by line: Method Comparison tab.', subf)
sm.cell(5, 2, 'Order volume (buildings)').font = bold
for j, v in enumerate(VOLS): sm.cell(5, 3 + j, v).font = bold
for i, m in enumerate(METHODS):
    sm.cell(6 + i, 2, m + ' - TDC').font = body
    for j, v in enumerate(VOLS):
        c = sm.cell(6 + i, 3 + j, '=' + E('tdc', 'm_%d_%d' % (i, v))); c.number_format = USD; c.font = link
sm.cell(8, 2, 'Land bank base average, stick-built ($819K)').font = body
for j in range(len(VOLS)): c = sm.cell(8, 3 + j, 819328); c.number_format = USD
sm.cell(13, 2, 'Breakpoint: smallest volume where modular TDC <= stick-built').font = bold
for i, m in enumerate(METHODS[:1]):
    sm.cell(14 + i, 2, m).font = body
    c = sm.cell(14 + i, 3, '=IFERROR(INDEX($C$5:$I$5,MATCH(TRUE,INDEX(C%d:I%d<=C7:I7,0),0)),"none at 100")' % (6 + i, 6 + i)); c.font = body
    c = sm.cell(14 + i, 4, '=F%d-F7' % (6 + i)); c.number_format = USD; c.font = body
sm.cell(13, 4, 'Gap vs stick at 10').font = bold
lc = LineChart(); lc.title = 'TDC per building vs. order volume'; lc.height = 8; lc.width = 18; lc.y_axis.numFmt = '$#,##0'; lc.x_axis.title = 'Buildings ordered'
data = Reference(sm, min_col=2, max_col=9, min_row=6, max_row=8)
lc.add_data(data, from_rows=True, titles_from_data=True); lc.set_categories(Reference(sm, min_col=3, max_col=9, min_row=5))
sm.add_chart(lc, 'B20')

# ---------------- Three Cities (static, written by the explorer)
tc = wb.create_sheet('Three Cities')
setup(tc, {'A': 2, 'B': 46, **{CL(c): 14 for c in range(3, 12)}})
put(tc, 'B2', 'THREE CITIES: same building, same decisions (static values written by the explorer at export)', title)
put(tc, 'B3', 'The live model on the other tabs runs the focus city only. These values come from the explorer for every city with the same settings, so they do not update if you change Inputs here.', subf)
CITIES = ['Hartford', 'New Haven', 'Waterbury']
for j, cty in enumerate(CITIES):
    c = tc.cell(5, 3 + 3 * j, cty); c.font = hdr; c.fill = hfill; tc.cell(5, 4 + 3 * j).fill = hfill; tc.cell(5, 5 + 3 * j).fill = hfill
    for b, h in enumerate(['Existing homes', 'Base case', 'DOH precedent']):
        x = tc.cell(6, 3 + 3 * j + b, h); x.font = bold; x.alignment = WRAP
TCM = [('City inputs: CHFA 100% AMI (buyers)', 'chfaAmi', USD), ('City inputs: HUD 50% limit (4-person)', 'ami50', USD), ('City inputs: mill rate', 'mill', DEC), ('City inputs: assessment ratio', 'assess', PCT1), ('City inputs: insurance per unit per year', 'ins', USD), ('City inputs: market rent, rental unit', 'market', USD),
       ('Total development cost', 'tdc', USD), ('Appraised value', 'fmv', USD), ('Most the buyer can pay', 'afford', USD), ('Sales gap', 'salesGap', USD), ('Remaining buyer gap', 'buyerRemain', USD), ('Developer gap (TDC - gross sales price)', 'devGap', USD),
       ('Developer-side money used', 'devUsed', USD), ('Remaining developer gap', 'remain', USD), ('Rent charged, rental unit', 'rent', USD), ("Owner's net monthly cost", 'net', USD), ('Net cost as % of income', 'netPct', PCT1), ('Surplus / shortfall vs MIT (per year)', 'residual', USD)]
for i, (l, k, fm) in enumerate(TCM):
    rr = 7 + i; tc.cell(rr, 2, l).font = body
    for j, cty in enumerate(CITIES):
        for b, bs in enumerate(['cama', 'custom', 'prec']):
            c = tc.cell(rr, 3 + 3 * j + b, 0); c.number_format = fm; c.font = body
            XM['tc_%s_%d_%s' % (k, j, bs)] = '%s%d' % (CL(3 + 3 * j + b), rr)
tc.cell(23, 2, 'Hartford and Waterbury reuse New Haven\'s itemized site and soft costs (flagged). The base case ($340K) and the DOH precedent ($420K) apply to every city until each has its own comparable.').font = note
put(tc, 'B25', 'Exported from the explorer:', bold); c = tc.cell(25, 3, 'not exported yet'); c.font = body; XM['tc_stamp'] = 'C25'

# ---------------- Mixed Income (rebuilt on new Homeowner rows)
mi = wb.create_sheet('Mixed Income')
setup(mi, {'A': 2, 'B': 46, 'C': 12, 'D': 16, 'E': 16, 'F': 16, 'G': 16})
put(mi, 'B2', 'MIXED-INCOME PHASE: 10 buildings at different owner AMI targets (live)', title)
put(mi, 'B3', 'The developer gap per building is the same at every AMI when a buyer subsidy fills the sales gap; the buyer-side money changes. Change the building counts in column C.', subf)
for j, h in enumerate(['Owner AMI', 'Buildings', 'Sales gap each', 'Total sales gap', 'Net cost % income', 'Residual vs MIT']): mi.cell(5, 2 + j, h).font = bold
for i, col in enumerate('CDEF'):
    rr = 6 + i
    c = mi.cell(rr, 2, '=Homeowner!%s5' % col); c.number_format = PCT; c.font = link
    c = mi.cell(rr, 3, [2, 4, 3, 1][i]); c.font = inp; c.fill = ifill
    c = mi.cell(rr, 4, '=Homeowner!%s12' % col); c.number_format = USD; c.font = link
    c = mi.cell(rr, 5, '=C%d*D%d' % (rr, rr)); c.number_format = USD
    c = mi.cell(rr, 6, '=Homeowner!%s27' % col); c.number_format = PCT1; c.font = link
    c = mi.cell(rr, 7, '=Homeowner!%s34' % col); c.number_format = USD; c.font = link
mi.cell(10, 2, 'Total').font = bold; mi['C10'] = '=SUM(C6:C9)'; mi['E10'] = '=SUM(E6:E9)'; mi['E10'].number_format = USD
line(mi, 11, 'Developer gap, whole phase (same per building)', "=%s*C10" % E('devGap'))
line(mi, 12, 'Buyer-side subsidy, whole phase (this mix)', '=E10')
line(mi, 13, 'Buyer-side subsidy if all 10 at 60% AMI', '=10*Homeowner!C12')
line(mi, 14, 'Buyer-side subsidy if all 10 at 80% AMI', '=10*Homeowner!D12')

# ---------------- Ownership Models (rebuilt, same logic as before)
om = wb.create_sheet('Ownership Models')
setup(om, {'A': 2, 'B': 50, 'C': 22, 'D': 22, 'E': 22, 'F': 50})
put(om, 'B2', 'OWNERSHIP MODELS: fee simple with deed restriction vs. CLT ground lease vs. shared equity (live, 80% AMI buyer)', title)
put(om, 'B3', 'Illustrative. Inputs in yellow are PLACEHOLDERS to replace with a CLT partner\'s terms.', subf)
for rr, l, v, f in [(5, 'Years until resale', 10, NUM), (6, 'Market appreciation per year', 0.03, PCT1), (7, 'Land value in appraisal (removed under a CLT)', 40000, USD), (8, "Owner's share of appreciation (CLT / shared equity)", 0.25, PCT), (9, 'CLT ground lease fee per month', 50, USD)]:
    om.cell(rr, 2, l).font = body; c = om.cell(rr, 3, v); c.font = inp; c.fill = ifill; c.number_format = f
for j, h in enumerate(['Fee simple + DOH deed restriction', 'CLT ground lease (99-yr)', 'Shared equity (silent 2nd)']): om.cell(11, 3 + j, h).font = bold
OMF = {12: ('Price the buyer pays (excl. assistance)', ['=Homeowner!D13', '=MAX(0,Homeowner!D13-$C$7)', '=Homeowner!D13']),
       13: ('Market value at purchase', ['=Homeowner!D11'] * 3),
       14: ('Monthly net housing cost', ['=Homeowner!D26', '=Homeowner!D26-$C$7*(1-%s)*%s+$C$9' % (IN('dp'), E('k', 'ami0.8')), '=Homeowner!D26']),
       15: ('Market value at resale', ['=Homeowner!D11*(1+$C$6)^$C$5'] * 3),
       16: ("Owner's share of appreciation", [1, '=$C$8', '=$C$8']),
       17: ('Mortgage principal repaid by resale', ['=Homeowner!D14+FV(%s/12,$C$5*12,PMT(%s/12,%s*12,Homeowner!D14),Homeowner!D14)' % (IN('rate'), IN('rate'), IN('term'))] * 3),
       18: ('OWNER WEALTH AT RESALE (down payment + principal + appreciation share)', ['=Homeowner!D15+C17+(C15-C13)*C16', '=Homeowner!D15+D17+(D15-D13)*D16', '=Homeowner!D15+E17+(E15-E13)*E16']),
       19: ('Price to the NEXT buyer (affordability kept?)', ['=C15', '=D12+(D15-D13)*D16', '=E15'])}
for rr, (l, fs) in OMF.items():
    om.cell(rr, 2, l).font = bold if rr == 18 else body
    for j, f in enumerate(fs):
        c = om.cell(rr, 3 + j, f); c.number_format = PCT if rr == 16 else USD; c.font = link if isinstance(f, str) and '!' in f else body
om.cell(21, 2, 'Support programs to pair with any model: CHFA homebuyer education and HUD counseling; CT Emergency Mortgage Assistance Program (EMAP); landlord training and a rental-unit reserve; partner with an existing CLT or shared-equity steward.').font = note

# ---------------- Benchmarks: repoint row 6 + add Hazel Street
bm = wb['Benchmarks']
bm['D6'] = '=' + E('gsf'); bm['E6'] = '=' + E('tdc'); bm['H6'] = '=' + E('devGap')
bm['B16'] = 'Hazel Street, New Haven (DOH-Yale Housing Clinic, 2024): 4 two-families, owner + rental'
bm['C16'] = 2; bm['H16'] = 650000; bm['I16'] = 'State funds for the homes: $2.6M / 4 buildings (CT Public, Dec 3, 2024). TDC not published.'

# ============================================================ Submission Summary
ss = wb.create_sheet('Submission Summary')
setup(ss, {'A': 2, 'B': 60, 'C': 16, 'D': 16, 'E': 16, 'F': 16, 'G': 62})
put(ss, 'B2', 'HOUSING CT: PREFAB MODULAR TEAM, FINANCIAL MODEL SUMMARY (live)', title)
put(ss, 'B3', '="One owner-occupied "&' + IN('units') + '&"-unit "&' + IN('method') + '&" building on a land bank lot in "&' + IN('city') + '&". Change the yellow cells on the Inputs tab; sources for every assumption are on the Inputs tab (column E) and the Sources tab."', subf)
r = 5; sec(ss, r, '1. UNITS, SIZE AND RENTS', 2, 7)
for j, h in enumerate(['Unit', 'Bedrooms', 'Net sq ft', 'Monthly rent', 'Rent per sq ft', 'Basis']): ss.cell(6, 2 + j, h).font = bold
ss['B7'] = 'Owner unit'; ss['C7'] = '=' + IN('ownerBR'); ss['D7'] = '=' + E('ownerSF'); ss['E7'] = 'Owner-occupied'; ss['G7'] = 'No rent; the owner pays the mortgage (section 4).'
ss['B8'] = '=IF(%s>1,"Rental unit"&IF(%s>2," (each of "&(%s-1)&")",""),"No rental unit")' % (IN('units'), IN('units'), IN('units'))
ss['C8'] = '=' + IN('rentBR'); ss['D8'] = '=' + E('rentSF'); ss['E8'] = '=' + E('rent'); ss['F8'] = '=IFERROR(E8/D8,0)'
ss['G8'] = '=IF(%s=0,"Market rent",IF(AND(%s="Yes",%s>%s),"Market rent (below the AMI limit)",TEXT(%s,"0%%")&" AMI rent limit, net of utility allowance"))' % (IN('rentAMI'), IN('capMkt'), E('rentNet'), E('market'), IN('rentAMI'))
for a in ['B7', 'B8', 'C7', 'C8', 'E7']: ss[a].font = body
for a, f in [('D7', NUM), ('D8', NUM), ('E8', USD), ('F8', '\\$0.00')]: ss[a].number_format = f; ss[a].font = link
ss['G7'].font = note; ss['G8'].font = note
line(ss, 10, 'Units in the building', '=' + IN('units'), NUM); line(ss, 11, 'Gross building area (sq ft)', '=' + E('gsf'), NUM)
line(ss, 12, 'AMI rent limit for the rental unit, net of utility allowance', '=' + E('rentNet'), USD, nt='HUD limit x 30% / 12, less the tenant utility allowance.')
line(ss, 13, 'Market asking rent, rental unit', '=' + E('market')); line(ss, 14, 'Construction method', '=' + IN('method'), '@'); line(ss, 15, 'Order volume (buildings ordered together)', '=' + IN('volume'), NUM)
r = 17; sec(ss, r, '2. TARGET BUYER AND PRICE POINT', 2, 7)
line(ss, 18, 'Target buyer income (% of AMI)', '=' + E('ami'), PCT); line(ss, 19, 'Household size', '=' + IN('hhSize'), NUM)
line(ss, 20, 'Household income at that target', '=' + E('income'), nt='HUD FY2026 income limits.')
line(ss, 21, 'Price point: gross sales price (= appraised value)', '=' + E('proceeds'), USD, True, True, nt='The developer budget counts the full price; buyer-side money fills any sales gap (Buyer Analysis tab).')
line(ss, 22, 'Most the buyer could pay (lender and 30% tests, with assistance)', '=' + E('afford'))
line(ss, 23, 'Buyer first mortgage', '=' + E('mortgage')); line(ss, 24, 'Buyer cash down payment', '=' + E('down'))
line(ss, 25, 'Buyer-side subsidy (Time To Own + DOH homebuyer subsidy)', '=' + E('buyerSubsidy'), nt='="Remaining buyer gap: "&TEXT(' + E('buyerRemain') + ',"$#,##0")')
r = 27; sec(ss, r, '3. TOTAL DEVELOPMENT COST (TDC), ONE BUILDING', 2, 7)
items = [('Building hard cost (factory price + local building lines)', 'bldg'), ('Foundation, delays, bond, contractor fee and contingency, other items, add-ons, warehouse', None), ('Site work outside the building', 'site'), ('Hard-cost contingency', 'hardContC'), ('Construction (hard) cost', 'construction'),
         ('Soft costs (design, permits, insurance, legal, sales, contingency)', 'soft'), ('Construction loan interest', 'interest'), ('Taxes and insurance during construction', 'carry'), ('Land acquisition and closing', 'land'), ('Developer fee', 'feeC'), ('TOTAL DEVELOPMENT COST (TDC)', 'tdc')]
for i, (l, k) in enumerate(items):
    f = '=%s+%s+%s+%s+%s+%s+%s+%s' % (E('hx'), E('featCost'), E('whNet'), E('found'), E('dlyHard'), E('bond'), E('cCont'), E('cFee')) if k is None else '=' + E(k)
    line(ss, 28 + i, l, f, USD, b=k in ('construction', 'tdc'), tot=k in ('construction', 'tdc'))
line(ss, 39, 'TDC per unit', '=' + E('tdcUnit')); line(ss, 40, 'TDC per gross sq ft', '=' + E('tdcPSF'), USD2)
line(ss, 41, 'Benchmark: Hartford partner 2-family, stick-built', 800000, nt='Verbal, External Review 1.')
line(ss, 42, 'Benchmark: Hazel Street (New Haven), state funds per two-family', 650000, nt='$2.6M / 4 buildings, CT Public Dec 3, 2024.')
for a in ['C41', 'C42']: ss[a].font = body
r = 44; sec(ss, r, '4. HOMEOWNER AFFORDABILITY (stand-alone test)', 2, 7)
ss.cell(45, 2, 'Owner income (% of AMI)').font = bold
for j, col in enumerate('CDEF'):
    c = ss.cell(45, 3 + j, '=Homeowner!%s5' % col); c.number_format = PCT; c.font = linkb
HH = [('Household income', 7, USD), ('Price the buyer can pay (with assistance)', 10, USD), ('Sales gap (needs a buyer subsidy)', 12, USD), ('Total monthly cost of the building', 24, USD), ('Less: rent from the rental unit (after vacancy)', 25, USD),
      ("Owner's net monthly cost", 26, USD), ('Net cost as % of income (30% standard)', 27, PCT1), ('Energy burden (6% = high)', 29, PCT1), ('Surplus / shortfall vs MIT living-wage budget, per year', 34, USD), ('Net cost % of income if vacancy, reassessment and a $10K repair hit at once', 40, PCT1)]
for i, (l, hr, fm) in enumerate(HH):
    rr = 46 + i; ss.cell(rr, 2, l).font = bold if hr == 26 else body
    for j, col in enumerate('CDEF'):
        c = ss.cell(rr, 3 + j, '=Homeowner!%s%d' % (col, hr)); c.number_format = fm; c.font = linkb if hr == 26 else link
        if hr == 26: c.fill = tfill
ss.cell(56, 2, 'Monthly costs: mortgage + FHA insurance, property tax, insurance, energy, water and sewer, 1% maintenance reserve, less any add-on savings. Detail: Homeowner tab.').font = note
r = 58; sec(ss, r, '5. APPRAISAL GAP', 2, 7)
ss.cell(59, 3, 'This scenario').font = bold; ss.cell(59, 4, 'Existing homes (low)').font = bold; ss.cell(59, 5, 'DOH precedent (high)').font = bold
APG = [('Total development cost', 'tdc'), ('Appraised value', 'fmv'), ('DEVELOPER GAP = TDC - gross sales price', 'devGap'), ('Sales gap (appraised value - what the buyer can pay)', 'salesGapAll'), ('Remaining gap after identified developer subsidies', 'remain'), ('Remaining buyer gap (kept separate)', 'buyerRemain')]
for i, (l, k) in enumerate(APG):
    rr = 60 + i; ss.cell(rr, 2, l).font = bold if 'DEVELOPER GAP' in l else body
    for j, sid in enumerate(['base', 'cama', 'prec']):
        c = ss.cell(rr, 3 + j, '=' + E(k, sid)); c.number_format = USD; c.font = linkb if 'DEVELOPER GAP' in l else link
        if 'DEVELOPER GAP' in l: c.fill = tfill
ss.cell(66, 2, '="Appraisal basis in this scenario: "&' + IN('basis') + '&". Both cases run live on the Appraisal Cases tab; three cities on the Three Cities tab."').font = note
r = 68; sec(ss, r, '6. SOURCES AND USES (detail on the Sources & Uses tab)', 2, 7)
line(ss, 69, 'Uses: total development cost', "='Sources & Uses'!$C$22", USD, True, True)
line(ss, 70, 'Buyer-side sources (mortgage, down payment, buyer assistance; remaining buyer gap shown separately)', "=SUM('Sources & Uses'!$C$26:$C$29)")
line(ss, 71, 'Developer-side subsidies (HDF, AHP, HTCC, DOH rental, NMTC, DOH)', "=SUM('Sources & Uses'!$C$31:$C$36)")
line(ss, 72, 'Remaining developer gap (unfunded)', "='Sources & Uses'!$C$37", USD, True)
line(ss, 73, 'Check: sources - uses (should be $0)', "='Sources & Uses'!$C$39")
ss.freeze_panes = 'A4'

# ============================================================ Sources & Uses
su = wb.create_sheet('Sources & Uses')
setup(su, {'A': 2, 'B': 50, 'C': 16, 'D': 11, 'E': 14, 'F': 26, 'G': 70})
put(su, 'B2', 'SOURCES AND USES: one building (live)', title)
put(su, 'B3', 'Program amounts are caps or precedents, not awards. Eligibility and stacking of developer subsidies are for the law team to confirm.', subf)
sec(su, 5, 'USES', 2, 7)
for j, h in enumerate(['Amount', '% of TDC', '', '', 'Source']): su.cell(5, 3 + j, h or None).font = hdr
UU = [('Land acquisition and closing', 'land', 'Inputs: land'), ('Building hard cost (factory price + local building lines)', 'bldg', 'Inputs: factory price and the line-item table (land bank structure)'), ('Foundation, delays, bond, contractor fee and contingency, other items, add-ons, warehouse', None, 'Inputs sections 3, 4b, 6, 7'),
      ('Site work (clearing, utilities, landscaping, fencing)', 'site', NHLB), ('Hard-cost contingency', 'hardContC', NHLB), ('Architecture, engineering, surveys', 'ae', INFILL), ('Environmental (Phase I and Phase II)', 'p1', NHLB),
      ('Permits and development fees (less any waiver)', 'pm', INFILL), ("Builder's risk insurance", 'rk', NHLB), ('Legal, title, appraisal, inspections', 'lg', HOB), ('Marketing, coordination and real estate commission', 'mk', HOB + ' Commission: ' + NHLB),
      ('Soft-cost contingency', 'sc', HOB), ('Construction loan interest', 'interest', 'Inputs: loan rate, draw, months'), ('Taxes and insurance during construction', 'carry', 'Placeholder'), ('Developer fee', 'feeC', 'Inputs: developer fee')]
SPEC = {'ae': '=%s+%s+%s' % (E('sArch'), E('sEng'), E('sSurv')), 'p1': '=%s+%s' % (E('sPh1'), E('phase2')), 'pm': '=%s-%s' % (E('permitsUsed'), E('waiver')), 'rk': '=' + E('sRisk'), 'lg': '=%s+%s' % (E('sLegal'), E('sAppr')), 'mk': '=%s+%s' % (E('sMkt'), E('commission'))}
for i, (l, k, s_) in enumerate(UU):
    rr = 6 + i; su.cell(rr, 2, l).font = body
    if k is None: f = '=%s+%s+%s+%s+%s+%s+%s+%s' % (E('hx'), E('featCost'), E('whNet'), E('found'), E('dlyHard'), E('bond'), E('cCont'), E('cFee'))
    elif k == 'sc': f = '=%s-(C11+C12+C13+C14+C15+C16)' % E('soft')
    elif k in SPEC: f = SPEC[k]
    else: f = '=' + E(k)
    c = su.cell(rr, 3, f); c.number_format = USD; c.font = link
    d = su.cell(rr, 4, '=IFERROR(C%d/$C$22,0)' % rr); d.number_format = PCT1
    su.cell(rr, 7, s_).font = note
su['B22'] = 'TOTAL USES (TDC)'; su['B22'].font = bold; su['C22'] = '=SUM(C6:C20)'; su['C22'].number_format = USD; su['C22'].font = bold; su['C22'].fill = tfill
su['B23'] = 'Check: equals the Engine TDC (should be $0)'; su['B23'].font = note; su['C23'] = '=C22-' + E('tdc'); su['C23'].number_format = USD
sec(su, 25, 'SOURCES (permanent, at sale to the buyer)', 2, 7)
for j, h in enumerate(['Amount', '% of TDC', 'Paid to', 'Status', 'Source / administrator']): su.cell(25, 3 + j, h).font = hdr
SS2 = [('Buyer first mortgage (FHA / CHFA, 30-year)', '=' + E('financedAll') + '*(1-' + E('dp') + ')', 'Buyer', 'Underwritten', 'CHFA or FHA lender. Rate: CHFA Oct 2026; terms: HUD Handbook 4000.1.'),
       ('Buyer cash down payment', '=' + E('financedAll') + '*' + E('dp'), 'Buyer', 'Underwritten', 'FHA 3.5% minimum.'),
       ('CHFA Time To Own (buyer assistance)', '=' + E('ttoAll'), 'Buyer', 'Program cap', 'CHFA: up to $25,000 at or below 80% AMI; $18,750 at 80-100%. ' + U['tto']),
       ('DOH homebuyer subsidy', '=' + E('hbAll'), 'Buyer', 'Precedent', 'CT DOH (redacted New Haven application). ' + U['doh']),
       ('Remaining buyer gap (buyer side, unfunded)', '=' + E('buyerRemain'), 'Buyer', 'Unfunded', 'Kept apart from the developer gap.'),
       ('HDF / CT Home Funds grant', '=' + E('sHdf'), 'Developer', 'Competitive', 'Up to $150,000 per home. ' + U['hdf']),
       ('FHLBank Boston AHP', '=' + E('sAhp'), 'Developer', 'Competitive', 'Up to $100,000 per unit, via a member bank. ' + U['ahp']),
       ('CHFA HTCC share (grant mode)', '=' + E('sHtcc'), 'Developer', 'Competitive', 'Up to $500,000 per nonprofit; the New Haven 2-unit pilot got $500K. In revolving mode it lowers loan interest instead. ' + U['htcc']),
       ('DOH rental subsidy (rental unit)', '=' + E('sDohRent'), 'Developer', 'To confirm', 'CHFA meeting: about $150K/unit at 50% AMI, $100K at 60%. Off unless switched on.'),
       ('New Markets Tax Credits (net benefit, via a CDE)', '=' + E('sNmtc'), 'Developer', 'To explore', 'Off unless switched on. See the NMTC tab. ' + U['cdfi']),
       ('CT DOH developer subsidy', '=' + E('sDoh'), 'Developer', 'Precedent', '$230K precedent; Hazel Street ~$650K per building. ' + U['doh']),
       ('REMAINING GAP (unfunded)', '=' + E('remain'), 'Open', 'Unfunded', 'See candidate sources below.')]
for i, (l, f, who, st, s_) in enumerate(SS2):
    rr = 26 + i; b = l.startswith('REMAINING')
    su.cell(rr, 2, l).font = bold if b else body; c = su.cell(rr, 3, f); c.number_format = USD; c.font = linkb if b else link
    d = su.cell(rr, 4, '=IFERROR(C%d/$C$22,0)' % rr); d.number_format = PCT1
    su.cell(rr, 5, who).font = body; su.cell(rr, 6, st).font = body; g = su.cell(rr, 7, s_); g.font = note; g.alignment = WRAP
    if b:
        for cc in range(2, 8): su.cell(rr, cc).fill = tfill
su['B38'] = 'TOTAL SOURCES'; su['B38'].font = bold; su['C38'] = '=SUM(C26:C37)'; su['C38'].number_format = USD; su['C38'].font = bold
su['B39'] = 'Check: sources - uses (should be $0)'; su['B39'].font = note; su['C39'] = '=C38-C22'; su['C39'].number_format = USD
sec(su, 41, 'CONSTRUCTION-PERIOD FINANCING (repaid from the sources above at sale)', 2, 7)
line(su, 42, 'Construction loan, peak need (hard + soft costs)', '=%s+%s' % (E('construction'), E('soft')))
su['G42'] = 'Candidate lenders: CHFA Build for CT, CDFIs, local banks; an NMTC structure would replace this with A/B loans from the CDE.'; su['G42'].font = note
line(su, 43, 'Construction loan rate', '=' + IN('loanRate'), PCT2); line(su, 44, 'Months the loan is outstanding', '=' + E('loanMonths'), DEC)
su['B45'] = 'Predevelopment / bridge capital before subsidies close'; su['B45'].font = body; su['C45'] = 'Open'
su['G45'] = 'Feedback #15: subsidies arrive after permits, so design, permits and factory deposits need bridge money (e.g., a CDFI revolving facility).'; su['G45'].font = note
sec(su, 47, 'CANDIDATE SOURCES FOR THE REMAINING GAP (to explore; not counted above)', 2, 7)
cands = [('New Markets Tax Credits through a CDE (if not yet on)', '=%s*%s' % (E('nmtcQEI'), E('nmtcRate')), 'Net benefit per building at the Inputs NMTC settings. Needs a qualifying census tract and a CDE partner.'),
         ('Higher appraisal from new-construction comparables', '=' + E('remain', 'prec'), 'Remaining developer gap in the new-build case (Appraisal Cases tab); the sales gap rises.'),
         ('DOH developer subsidy at the Hazel Street level (~$650K)', '=MAX(0,%s-650000)' % E('dohFull'), 'Remaining gap if DOH funded up to $650K per building.'),
         ('Larger order volume / real manufacturer quotes', None, 'Scale & Method tab.'),
         ('City of New Haven federal block grants (HOME, CDBG)', None, 'To confirm eligibility and timing with the city. Not modeled.'),
         ('Developer fee deferral', '=' + E('feeC'), 'Last resort only: reviewers said foregoing the fee is not sustainable (Feedback #13).')]
for i, (l, f, s_) in enumerate(cands):
    rr = 48 + i; su.cell(rr, 2, l).font = body
    if f: c = su.cell(rr, 3, f); c.number_format = USD; c.font = link
    g = su.cell(rr, 7, s_); g.font = note; g.alignment = WRAP
su['B55'] = 'Stacking rules and eligibility have not been assumed; the law team will confirm them.'; su['B55'].font = note

# ============================================================ Read Me
rm = wb.create_sheet('Read Me', 0)
setup(rm, {'A': 2, 'B': 120})
lines = [('HOUSING CT · PREFAB MODULAR TEAM · LIVE FINANCIAL MODEL', title),
         ('Exported from Prefab Gap Explorer 2. Every tab recalculates from the Inputs tab; the Engine tab runs the full model once per scenario column.', subf)] + [('', None)] * 15 + [
         ('HOW TO USE', bold), ('Start with the four tabs after this one: Homeowner Affordability (change the AMI in one cell and everything pulls through), TDC (the budget line by line), Sensitivity (one or two drivers at a time), then Inputs.', body),
         ('Submission Summary: the six items for the class submission (units and rents, buyer and price, TDC, homeowner affordability, appraisal gap, sources and uses).', body),
         ('Inputs: every assumption with its status (Verified, Sourced estimate, Course budget, Program rule, Placeholder, Team choice) and source. Change any yellow cell.', body),
         ('TDC, Method Comparison, Cost-Saving Levers, Gaps & Stack, Buyer Analysis, Homeowner, Appraisal Cases, Add-ons & Procurement, NMTC, Sensitivity, Scale & Method: live detail tabs.', body),
         ('Three Cities: static values for Hartford, New Haven and Waterbury written by the explorer at export (does not recalculate).', body),
         ('Mixed Income, Ownership Models, Schedule, Benchmarks, Comps, Constraints, Feedback Tracker: carried from the team workbook (Oct 14, 2026).', body), ('', None),
         ('KEY CAVEATS', bold),
         ('Manufacturer costs are faculty placeholders, not quotes. Program amounts are caps or precedents, not awards. Stacking rules are for the law team to confirm.', body),
         ('Hartford and Waterbury reuse New Haven\'s itemized site and soft costs.', body), ('', None),
         ('THIS EXPORT', bold)]
for i, (t, f) in enumerate(lines):
    if t: c = rm.cell(2 + i, 2, t); c.font = f
XM['rm_stamp'] = 'B%d' % (2 + len(lines)); rm.cell(2 + len(lines), 2, 'Template (not yet exported from the explorer).').font = body
XM['rm_link'] = 'B%d' % (3 + len(lines)); rm.cell(3 + len(lines), 2, '-').font = body
XM['rm_logrow'] = 5 + len(lines)

# ============================================================ Sources tab
so = wb.create_sheet('Sources')
setup(so, {'A': 2, 'B': 5, 'C': 80, 'D': 70})
put(so, 'B2', 'SOURCES', title)
SL = [('New Haven Land Bank (Barbara Malmberg, Nov 2, 2026): construction estimates spreadsheet, cost estimates table, Infill Homeownership Initiative webinar (Apr 8, 2026), RKG Associates housing study', 'Shared with the team (not public)'),
      ('CHFA meeting notes (Oct 30, 2026): AMI targets, AHP sponsor, developer fees, DOH rental subsidy tiers, module storage', 'Team notes'),
      ('Reframe Systems call (Oct 30, 2026, Ben and Chris Eggert): ~30% faster than site-built, slab on grade for ledge, roof module option', 'Team notes'),
      ("Kate's Infill_DOH_example_2025.xlsx with the Shoreline modular quote (via NHLB): 60 ft trailer access", 'Shared with the team (not public)'),
      ('CHFA 2021 HTCC awards, incl. $500K to the Capital for Change Loan Pool', 'https://www.ncsha.org/hfa-news/tax-credits-will-provide-financing-for-more-than-786-units-of-affordable-housing-fund-loan-program/'),
      ('CHFA 2024 HTCC awards, incl. three revolving down-payment funds', 'https://www.ncsha.org/hfa-news/chfa-awards-10-6-million-in-tax-credits-to-boost-affordable-housing-across-connecticut/'),
      ('Redacted course files (Canvas): New Haven DOH Consolidated Application (2025), 1-2 unit infill budget, homeownership budget', 'Course files (not public)'),
      ('F26 Housing CT Model Workshop v2.2, Volume Curve tab (faculty placeholders)', 'Course files (not public)'), ('Partner feedback, External Review 1 (Oct 9, 2026)', 'Review notes'),
      ('CT DOH Homeownership Developer Subsidy Initiative outline (Mar 21, 2024)', U['doh']), ('CT Public: Hazel Street, first DOH-Yale Housing Clinic project (Dec 3, 2024)', U['hazel']),
      ('CT DOH press release: Fourteen Developments (New Haven two-unit pilot, $3.5M DOH + $500K HTCC)', U['pilot']), ('CT Public: Hartford $1 lots and up to $150K per unit (Dec 17, 2025)', U['hfdlots']),
      ('CT Parcel & CAMA data, recorded sales 2023-26', U['cama']), ('HUD FY2026 income limits', U['hudil']), ('CHFA homebuyer mortgage rates (Oct 2026)', U['chfa']),
      ('Freddie Mac PMMS (Oct 22, 2026)', U['pmms']), ('Bankrate: WSJ prime rate (Oct 27, 2026)', U['prime']), ('Housing Authority of New Haven (Elm City Communities) utility allowance schedule, Jan 2024', U['ecc']),
      ('NerdWallet: Connecticut home insurance by city (Mar 2026)', U['nw']), ('CCM municipal profile, New Haven (FY2027 mill rate)', U['ccm']), ('CT OLR Report 2012-R-0098: property assessment', U['olr']),
      ('Bloomberg Tax: CT court on deed-restricted fair market value (May 2026)', U['nlondon']), ('Hartford Business Journal: Hartford 10-year tax-fixing agreement', U['hfdfix']),
      ('MIT Living Wage Calculator, New Haven metro (Feb 15, 2026)', U['mit']), ('FHA annual mortgage insurance premium', U['mip']), ('CHFA Time To Own', U['tto']), ('CT Home Funds / HDF', U['hdf']),
      ('FHLBank Boston 2026 AHP', U['ahp']), ('CHFA HTCC 2026 round', U['htcc']), ('Banker & Tradesman: Reframe Systems costs', U['reframe']),
      ('NY Fed: Intro to NMTC', U['nyfed']), ('Chicago Development Fund: NMTC structure', U['cdf']), ('Smith NMTC: affordable homeownership', U['smith']), ('Novogradac: NMTC for-sale housing (Dec 9, 2025)', U['novo']),
      ('CDFI Fund: Introduction to the NMTC Program', U['cdfi']), ('EnergySage: CT solar cost', U['escost']), ('EnergySage: 25D credit ended', U['es25d']), ('Eversource RRES', U['rres']), ('EnergySage: CT electricity cost', U['esrate']),
      ('Solar Builder: CT Energy Storage Solutions', U['ess']), ('CommercialCafe: New Haven industrial rents', U['whrent'])]
for i, (t, u) in enumerate(SL):
    rr = 4 + i; so.cell(rr, 2, i + 1).font = body; so.cell(rr, 3, t).font = body; c = so.cell(rr, 4, u); c.font = note
    if u.startswith('http'): c.hyperlink = u; c.font = F(sz=9, color='FF1F5FBF', u='single')

# ============================================================ Homeowner Affordability: change one cell
ha = wb.create_sheet('Homeowner Affordability')
setup(ha, {'A': 2, 'B': 58, 'C': 18, 'D': 4, 'E': 70})
put(ha, 'B2', 'HOMEOWNER AFFORDABILITY: change the income target and everything pulls through (live)', title)
put(ha, 'B3', 'This cell drives the owner income target for the whole workbook (Inputs points here). Everything else is set on the Inputs tab.', subf)
ha['B5'] = "Owner household income target (% of AMI): CHANGE THIS CELL"; ha['B5'].font = bold
c = ha.cell(5, 3, 1.0); c.font = inp; c.fill = bfill; c.number_format = PCT
ha['E5'] = '80-100% AMI = workforce housing; above 120% = market rate (SOM meeting, Nov 2). CHFA and the RKG study point to 100-120%.'; ha['E5'].font = note
ha['B6'] = 'Income band'; ha['B6'].font = body
c = ha.cell(6, 3, '=IF(C5<=0.5,"Very low income",IF(C5<=0.8,"Low income",IF(C5<=1.0001,"Workforce",IF(C5<=1.2001,"Upper workforce","Market rate"))))'); c.font = linkb
HAR = [('A. WHO IS BUYING', None), ('Household income', 'income', USD), ('Household size', 'hhN', NUM),
       ('B. WHAT THEY CAN PAY', None), ('Test A: lender max price', 'testA', USD), ('Test B: owner max price', 'testB', USD), ('Buyer capacity (lower of the tests)', 'capacity', USD),
       ('Buyer assistance used (Time To Own, City DPA, other)', 'tto', USD), ('Most the buyer can pay', 'afford', USD), ('Sales price of the owner unit (the whole building if not a condo)', 'unitValue', USD),
       ('Sales gap (price - what the buyer can pay)', 'salesGap', USD), ('DOH homebuyer subsidy', 'hbAll', USD), ('REMAINING BUYER GAP', 'buyerRemain', USD),
       ('C. MONTHLY COST OF OWNING', None), ('Mortgage payment + FHA insurance', 'pi', USD2), ('Property tax', 'tax', USD2), ('Insurance', 'insM', USD2), ('Water and sewer', 'commonM', USD2),
       ('Owner energy', 'energyM', USD2), ('Maintenance reserve', 'maintM', USD2), ('Rent collected from the rental unit(s)', 'collected', USD2), ("OWNER'S NET MONTHLY COST", 'net', USD2),
       ('All-in housing cost as % of income', 'netPct', PCT1), ('Mortgage payment alone as % of income', 'piPct', PCT1),
       ('D. THE DEVELOPER SIDE AT THIS PRICE', None), ('Total development cost (TDC)', 'tdc', USD), ('Developer gap (TDC - sales price)', 'devGap', USD), ('Remaining developer gap after developer-side money', 'remain', USD)]
rr = 8
for row in HAR:
    if row[1] is None: sec(ha, rr, row[0], 2, 5); rr += 1; continue
    l, k, f = row; ha.cell(rr, 2, l).font = bold if l.isupper() else body
    c = ha.cell(rr, 3, '=' + E(k)); c.number_format = f; c.font = linkb if l.isupper() else link
    if l.isupper(): c.fill = tfill
    rr += 1
ha.cell(rr, 2, 'Over 30% of income all-in = cost-burdened (HUD); the 30% usually applies to the mortgage alone, and all-in costs run about 33-40% (SOM meeting, Nov 2). See Sensitivity grids 7-10.').font = note
rr += 2; sec(ha, rr, 'E. THE SAME HOME AT FOUR INCOME TARGETS (fixed columns)', 2, 7)
for j, a in enumerate(AMIS): c = ha.cell(rr + 1, 3 + j if j == 0 else 3 + j, a); c.number_format = PCT; c.font = bold
ha.column_dimensions['D'].width = 16; ha.column_dimensions['F'].width = 16
for i, (l, k, f) in enumerate([('Household income', 'income', USD), ('Most the buyer can pay', 'afford', USD), ('Remaining buyer gap', 'buyerRemain', USD), ("Owner's net monthly cost", 'net', USD), ('All-in % of income', 'netPct', PCT1)]):
    ha.cell(rr + 2 + i, 2, l).font = body
    for j, a in enumerate(AMIS): c = ha.cell(rr + 2 + i, 3 + j, '=' + E(k, 'ami%g' % a)); c.number_format = f; c.font = link
ha.cell(rr + 1, 2, 'Income target (% AMI)').font = bold
for j in range(len(AMIS)): ha.cell(rr + 1, 3 + j).number_format = PCT
ws[XM['ownerAMI']] = "='Homeowner Affordability'!$C$5"; ws[XM['ownerAMI']].font = link
ws.cell(int(XM['ownerAMI'][1:]), 5).value = 'Set on the Homeowner Affordability tab (change it there). ' + (ws.cell(int(XM['ownerAMI'][1:]), 5).value or '')

# ============================================================ Cover: headline results (live)
rm.title = 'Cover'
rm.column_dimensions['B'].width = 62; rm.column_dimensions['C'].width = 18
put(rm, 'B4', '="Focus city: "&' + IN('city') + '&" · "&' + IN('method') + '&" · "&TEXT(' + E('ami') + ',"0%")&" AMI owner"', subf)
CV = [('HEADLINE RESULTS (live, one building)', None, None), ('Total development cost (TDC)', 'tdc', USD), ('Gross sales price', 'salesTotal', USD), ('Developer gap (TDC - sales price)', 'devGap', USD),
      ('Developer-side money used (HDF, AHP, HTCC, DOH...)', 'devSubsidy', USD), ('REMAINING DEVELOPER GAP', 'remain', USD), ('Price the buyer can pay', 'buyerCanPay', USD), ('REMAINING BUYER GAP', 'buyerRemain', USD),
      ("Owner's net monthly cost", 'net', USD), ('All-in housing cost as % of income', 'netPct', PCT1)]
for i, (l, k, f) in enumerate(CV):
    if k is None: sec(rm, 5 + i, l, 2, 3); continue
    rm.cell(5 + i, 2, l).font = bold if l.isupper() else body
    c = rm.cell(5 + i, 3, '=' + E(k)); c.number_format = f; c.font = linkb if l.isupper() else link
    if l.isupper(): c.fill = tfill
rm.cell(5 + len(CV) + 1, 2, 'Placeholders are placeholders: manufacturer prices, deposits, delays and program amounts are not quotes or awards (Inputs column D says how solid each is).').font = note

# ============================================================ order + save
order = ['Cover', 'Homeowner Affordability', 'TDC', 'Sensitivity', 'Submission Summary', 'Sources & Uses', 'Inputs', 'Method Comparison', 'Cost-Saving Levers', 'Gaps & Stack', 'Buyer Analysis', 'Homeowner', 'Appraisal Cases', 'Add-ons & Procurement', 'NMTC', 'Scale & Method',
         'Three Cities', 'Mixed Income', 'Ownership Models', 'Schedule', 'Benchmarks', 'Comps', 'Constraints', 'Feedback Tracker', 'Sources', 'Engine']
wb._sheets = [wb[n] for n in order] + [s for s in wb._sheets if s.title not in order]
for w in wb.worksheets: w.sheet_view.tabSelected = False
wb.active = 1; wb.worksheets[1].sheet_view.tabSelected = True
wb.save(D + 'template.xlsx')
XM['ownerAMI'] = 'Homeowner Affordability!C5'
json.dump({'cells': XM, 'engineBase': {k: E(k) for k in ROW if not k.startswith('ov_')}, 'scenarios': [s[0] for s in SC]}, open(D + 'xmap.json', 'w'), indent=0)
print('saved; inputs', len([e for e in I if e[0] != '#']), 'scenarios', len(SC), 'engine rows', len(LINES))
