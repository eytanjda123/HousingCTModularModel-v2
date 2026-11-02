# Export 2 uses the team's own "City Models and Charts" workbook as the template (same layout, styles and 18 charts).
# This script writes sub/cmc2_map.json: which template cells the explorer overwrites, and with what.
# e = JS expression (context: R s c T city AMI U SC SCR CROSS FOCUS RF ...). t: n number, s string. link: keep the
# template's link formula unless the city's value differs from the shared value.
import json, shutil
SRC='/mnt/user-data/attachments/cf885be6-d73a-57f1-84e3-0c4b8ede712e/6bfba090-Housing_CT_Prefab_City_Models_and_Charts.xlsx'
shutil.copy(SRC,'sub/cmc2_template.xlsx')
CITIES=["Hartford","New Haven","Waterbury"]
M=[]
def N(s,r,e,c=None,**k): M.append(dict(s=s,r=r,e=e,c=c,t='n',**k))
def S_(s,r,e,c=None,**k): M.append(dict(s=s,r=r,e=e,c=c,t='s',**k))
def T_(s,r,txt): M.append(dict(s=s,r=r,e=json.dumps(txt),c=None,t='s'))
COLS='CDEFGHI'
# ---------- Read Me
S_('Read Me','B3','"Exported "+STAMP+" from Prefab Gap Explorer 2 (focus city "+FOCUS+", design "+designOf(T)+"). Same layout as the team workbook of Oct 8; every number now comes from the explorer as set at export."')
T_('Read Me','B14','Blue text on yellow = an input as set in the explorer at export.  Green = linked from Shared Inputs (replaced by a value where a city has its own setting).  Black = a result calculated by the explorer, or a formula.')
T_('Read Me','B16','HOW THE MODEL WORKS (Prefab Gap Explorer 2; the city tabs hold the explorer\'s results)')
T_('Read Me','B17','1. Rent: the rental unit charges the lower of the AMI rent (rent + utilities at the set share of a renter household\'s income, less the utility allowance) and market asking rent.')
T_('Read Me','B18','2. Buyer capacity: the lower of Test A (lender: 40% of own income + 75% of rent) and Test B (owner: 30% of own income + rent collected), after tax, insurance, utilities and reserves.')
T_('Read Me','B19','3. Price: the building sells at its appraised value (the basis set in the explorer: existing-home median, size-adjusted new-build value, the DOH precedent or the land bank value). The buyer pays what their own mortgage and down payment cover; Time To Own, City DPA and the DOH homebuyer subsidy fill the buyer gap.')
T_('Read Me','B20','4. Cost: the New Haven Land Bank cost structure. Volumetric = a factory price per sq ft plus the share of each building line a local contractor still does; stick-built = every line. Both add site work, the foundation, bond, contractor fee and contingency, then soft costs, the construction loan (shorter for modular), carry, land and the developer fee: total development cost (TDC).')
T_('Read Me','B21','5. Two separate gaps. Developer gap = TDC - gross sales price, filled by HDF, AHP (if counted), HTCC, DOH rental (if on), NMTC (if on) and DOH. Buyer gap = sales price - what the buyer pays from their own mortgage and down payment, filled only by buyer-side money. The two are never added together.')
T_('Read Me','B24','Factory price (Shoreline quote unless changed), line items (land bank base average), land closing, warehouse and add-on figures, and every funding amount. Mill rates and assessment ratios are confirmed with each city.')
T_('Read Me','B25','The funding stack uses caps and precedents, not awards. Whether DOH, HDF, AHP and HTCC can fund the same home is for the law team to confirm.')
T_('Read Me','B26','For the fully live model with every formula, use Export model in the explorer. Changes from the explorer defaults are listed at the bottom of this tab.')
# ---------- Shared Inputs (focus-city scenario, shared values)
SI='Shared Inputs'
for r,e,lab,unit in [(6,'s.rate',None,None),(7,'s.term',None,None),(8,'s.dp',None,None),(9,'s.mip',None,None),(10,'s.rentCredit',None,None),(11,'s.lenderCap',None,None),(12,'s.ownerCap',None,None),
    (13,'c.ins','Homeowner insurance (focus city; each city tab uses its own)',None),(14,'s.common','Owner-paid common utilities (water and sewer)',None),
    (15,'s.maintPct','Maintenance and reserves','% of value per year'),(16,'s.vac',None,None),(17,'s.rentShare','Rent + utilities as % of renter income','% of income'),
    (20,'U(s.units,1).hardPSF','Stick-built baseline: all-in hard cost ($ per gross sq ft)',None),(21,'s.grossup',None,None),(22,'R.soft','Soft costs (itemized, incl. soft contingency)','$ per building'),
    (23,'s.hardCont','Hard-cost contingency','% of hard cost'),(24,'s.loanRate','Construction loan rate','% per year on the average draw'),(25,'s.fee','Developer fee','% of cost before fee')]:
    N(SI,f'C{r}',e)
    if lab: T_(SI,f'B{r}',lab)
    if unit: T_(SI,f'D{r}',unit)
for r,e in [(6,'CTRL_BY.rate.why'),(8,'CTRL_BY.dp.why'),(9,'CTRL_BY.mip.why'),(13,'CTRL_BY.ins.why'),(14,'CTRL_BY.common.why'),(15,'CTRL_BY.maintPct.why'),(20,'CTRL_BY.stickPSF.why'),(22,'"Architect, engineering, surveys, Phase I, permits, risk, legal, appraisal and inspection, marketing and coordination, from the redacted course budgets."'),(25,'"DOH accepts profit up to 22% of non-acquisition cost; a nonprofit may defer."')]:
    S_(SI,f'E{r}',e)
for j in range(5):
    N(SI,f'{"CDEFG"[j]}29',f's.sqft[{j}]'); N(SI,f'{"CDEFG"[j]}30',f's.brFactor[{j}]'); N(SI,f'{"CDEFG"[j]}31',f's.ua[{j}]')
T_(SI,'H29',"Course sizes, or the Reframe design picked in the explorer.")
T_(SI,'H31','Housing Authority of New Haven schedule (Jan 2024), all-electric with heat pump.')
for j in range(6): N(SI,f'{"CDEFGH"[j]}35',f's.hhFactor[{j}]')
T_(SI,'B37','VOLUMETRIC MODULAR COST BY BUILDING SIZE (at the order volume set in the explorer)')
T_(SI,'B39','Factory price ($ per gross sq ft)'); T_(SI,'B40','Local building work ($ per gross sq ft)'); T_(SI,'B41','Transport, crane and set outside the factory price ($ per building)'); T_(SI,'B42','Volumetric building cost ($ per gross sq ft)')
for j,n in enumerate([2,3,4]):
    col="CDE"[j]; N(SI,f'{col}39',f'U({n},0).facPSF'); N(SI,f'{col}40',f'U({n},0).fieldPSF'); N(SI,f'{col}41',f'U({n},0).trans')
S_(SI,'B43','"Factory price: "+(FACT_T.find(x=>x[0]===s.cost.fact.opt)||FACT_T[0])[1]+" at a "+s.volume+"-building order; local work from the land bank line items (shares from Kate\'s DOH example)."')
for r,e in [(46,'s.tto'),(47,'s.otherAssist'),(48,'s.dohUnit*s.units'),(49,'s.hdf'),(50,'(s.inc.ahp?s.ahp:0)'),(51,'s.htcc'),(52,'s.htccHomes')]: N(SI,f'C{r}',e)
S_(SI,'B48','"DOH Developer Subsidy per building ("+usdK(s.dohUnit)+" per unit, land bank)"')
# ---------- City tabs
for city in CITIES:
    def n(r,e,**k): N(city,r,e,city,**k)
    def st(r,e,**k): S_(city,r,e,city,**k)
    def lab(r,t): T_(city,r,t)
    st('B2','city.toUpperCase()+": one "+s.units+"-unit owner-plus-rental building, "+(s.method==="Stick-built"?"stick-built":"prefab modular")+" ("+designOf(T)+")"')
    lab('B3','Values exported from Prefab Gap Explorer 2. Blue on yellow = inputs as set. Green = linked to Shared Inputs. Black = results the explorer calculated. The Compare and Charts tabs link to this tab.')
    st('C6','city'); n('C7','s.units'); st('C10','s.ownerBR'); st('C11','s.rentBR'); n('C12','s.ownerAMI'); n('C13','s.hhSize'); n('C14','s.rentAMI')
    st('C15','s.capMkt?"Yes":"No"'); st('C16','s.method'); lab('D16','Construction method in the explorer')
    n('C21','c.ami50')
    for j in range(5): n(f'{"CDEFG"[j]}24',f'c.mkt[{j}]')
    for j in range(3): n(f'{"CDE"[j]}26',f'c.med[{j}]')
    lab('B26','Median sale, existing buildings (CT CAMA; the low appraisal case)')
    n('C27','R.assessUsed'); st('E27','(s.units>=4?"Ratio for 4+ unit buildings.":"Ratio for 1-3 unit homes.")+(city==="Hartford"?" Hartford uses 36.8% for 1-3 units; 4+ units differ (law-team question).":" CT statutory 70%.")')
    n('C28','c.mill'); lab('E28','Confirmed by the team with the city tax office.')
    # section 3 links
    for r,e in [(33,'s.rate'),(34,'s.term'),(35,'s.dp'),(36,'s.mip'),(37,'s.rentCredit'),(38,'s.lenderCap'),(39,'s.ownerCap'),(40,'c.ins'),(41,'s.common'),(42,'s.maintPct'),(43,'s.vac'),(44,'s.rentShare'),(45,'U(s.units,1).hardPSF'),(46,'s.grossup'),(47,'R.soft'),(48,'s.hardCont'),(49,'s.loanRate'),(50,'s.fee')]:
        N(city,f'C{r}',e,city,link=True)
    for r,t in [(40,'Insurance ($/unit/yr, this city)'),(41,'Owner-paid common utilities ($/unit/mo)'),(42,'Maintenance and reserves (% of value/yr)'),(44,'Rent + utilities as % of renter income'),(47,'Soft costs ($ per building)'),(48,'Hard-cost contingency'),(49,'Construction loan rate'),(50,'Developer fee')]:
        lab(f'B{r}',t)
    # section 4
    n('C53','R.ownerSF'); n('C54','R.rentSF'); n('C55','R.gsf'); st('C56','["2-unit","3-unit","4-6 unit"][R.bucket]')
    lab('B57','Appraised value used (basis set in the explorer)'); n('C57','R.fmv'); st('E57','(BASIS.find(b=>b[0]===s.basis)||[0,s.basis])[1]')
    n('C60','s.brFactor[BR.indexOf(s.rentBR)]'); lab('B61','AMI rent limit, gross (rent + utilities at the set % of income)'); n('C61','R.tenantInc*s.rentShare/12')
    n('C62','s.ua[BR.indexOf(s.rentBR)]'); n('C63','R.rentNet'); n('C64','R.market'); n('C65','R.rent'); n('C66','R.grossRent'); n('C67','R.collected'); n('C68','R.lenderRentM*12')
    n('C71','R.income/R.ami'); n('C72','R.income')
    lab('B73','Property tax per month (on the assessed value)'); n('C73','R.tax'); n('C74','R.insM')
    lab('B75','Owner-paid utilities per month (common + owner energy)'); n('C75','R.commonM+R.energyM')
    lab('B76','Maintenance and reserves per month (less add-on savings)'); n('C76','R.maintM-R.featSaveMo')
    n('C77','R.opex'); n('C78','R.k'); n('C79','Math.max(0,R.testA)'); n('C80','Math.max(0,R.testB)'); n('C81','R.capacity')
    lab('B82',"Buyer's price from own mortgage + down payment (lower of capacity and appraisal)"); n('C82','R.financedAll')
    st('C83','R.capacity<R.fmv?"Buyer income":"Appraisal"')
    lab('B84','Buyer gap (sales price minus the buyer\'s own price)'); n('C84','Math.max(0,R.salesTotal-R.financedAll)')
    n('C87','R.bldg>0&&R.gsf>0?(R.bldg-R.trans)/R.gsf:0'); n('C88','R.trans')
    lab('B89','Hard cost (factory, local work, site, foundation, bond, contractor fee, add-ons)'); n('C89','R.construction-R.hardContC')
    lab('B90','Soft costs (itemized, incl. soft contingency)'); n('C90','R.soft')
    lab('B91','Hard-cost contingency'); n('C91','R.hardContC')
    lab('B92','Construction loan interest, carry and land'); n('C92','R.interest+R.carry+R.land')
    lab('B93','Development cost before developer fee'); n('C93','R.preFee')
    lab('B94','Development cost before fee, per unit'); n('C94','R.preFee/s.units')
    lab('B95','Developer fee'); n('C95','R.feeC')
    lab('B96','Total development cost (TDC, incl. developer fee)'); n('C96','R.tdc')
    lab('B99','Developer gap per building (TDC minus gross sales price; developer-side money fills it)'); n('C99','R.devGap'); lab('B100','Developer gap per unit'); n('C100','R.devGap/s.units'); lab('B101','Developer gap as % of TDC'); n('C101','R.devGap/R.tdc')
    lab('B102','Buyer gap (sales price minus the buyer\'s own price; buyer-side money fills it, kept separate)'); n('C102','Math.max(0,R.salesTotal-R.financedAll)')
    lab('B103',"Buyer's price as % of the existing-building median"); n('C103','Math.min(R.capacity,R.fmv)/c.med[R.bucket]')
    n('C106','R.pi'); n('C107','R.tax'); n('C108','R.insM'); n('C109','R.commonM+R.energyM'); lab('B110','Plus maintenance and reserves (less add-on savings)'); n('C110','R.maintM-R.featSaveMo'); n('C112','R.collected')
    lab('B118','Hard cost (building, add-ons, site work)'); lab('B121','Construction loan, carry and land'); lab('B122','Developer fee')
    lab('B123','Buyer gap (buyer-side subsidies)'); N(city,'D123','Math.max(0,R.salesTotal-R.financedAll)',city); lab('B124',"Buyer's own mortgage + down payment"); N(city,'D124','R.financedAll',city)
    lab('B125','Developer gap (developer-side subsidies)'); N(city,'D125','R.devGap',city)
    lab('B131','Total development cost'); n('C131','R.tdc')
    n('C132','R.financedAll')
    lab('B133','Buyer gap, paid with buyer-side subsidies (Time To Own, DPA, DOH homebuyer); any shortfall is shown separately'); n('C133','Math.max(0,R.salesTotal-R.financedAll)')
    lab('B134','HDF / C4C grant'); n('C134','R.sHdf')
    lab('B135','FHLBank Boston AHP'); n('C135','R.sAhp')
    lab('B136','HTCC share'); n('C136','R.sHtcc')
    st('B137','[R.sDohRent>0.5?"DOH rental subsidy":"",R.sNmtc>0.5?"New Markets Tax Credits":"","DOH Developer Subsidy (cap per home)"].filter(Boolean).join(" + ")'); n('C137','R.sDohRent+R.sNmtc+R.sDoh')
    lab('B138','Remaining developer gap (buyer gap shown separately)'); n('C138','R.remain'); n('D138','0')
    st('B139','"Developer-side amounts are what each program covers in the explorer, in order (HDF, AHP, HTCC, DOH rental, NMTC if on, DOH), up to each cap. Separately, the remaining buyer gap after buyer subsidies is "+usdK(R.buyerRemain)+". Caps and precedents, not awards; stacking is for the law team to confirm."')
    for j,x in enumerate([0.5,0.6,0.7,0.8,0.9,1.0,1.2]):
        col=COLS[j]
        n(f'{col}143',f'AMI({x}).income'); n(f'{col}144',f'Math.max(0,AMI({x}).testA)'); n(f'{col}145',f'Math.max(0,AMI({x}).testB)'); n(f'{col}146',f'AMI({x}).capacity')
        n(f'{col}147',f'AMI({x}).fmv'); n(f'{col}148',f'Math.min(AMI({x}).capacity,AMI({x}).fmv)'); n(f'{col}149',f'Math.max(0,AMI({x}).salesTotal-AMI({x}).financedAll)')
    lab('B147','Appraised value used'); lab('B148',"Buyer's price (own mortgage + down payment)"); lab('B149','Buyer gap at this income (filled by buyer-side subsidies)')
    lab('B150','The crossover is where buyer capacity first exceeds the appraised value; above it the appraisal, not income, sets the price.')
    for j,u in enumerate([2,3,4]):
        col="CDE"[j]
        n(f'{col}154',f'U({u},0).gsf'); n(f'{col}155',f'U({u},0).fmv'); n(f'{col}156',f'U({u},0).construction'); n(f'{col}157',f'U({u},0).tdc'); n(f'{col}158',f'U({u},0).tdcUnit')
        n(f'{col}159',f'U({u},1).construction'); n(f'{col}160',f'U({u},1).tdc'); n(f'{col}161',f'U({u},1).tdcUnit')
        n(f'{col}163',f'U({u},0).grossRent'); n(f'{col}164',f'U({u},0).opex'); n(f'{col}165',f'Math.max(0,U({u},0).testA)'); n(f'{col}166',f'Math.max(0,U({u},0).testB)')
        n(f'{col}167',f'Math.min(U({u},0).capacity,U({u},0).fmv)'); n(f'{col}168',f'U({u},0).devGap/{u}')
    lab('B155','Appraised value used'); lab('B156','Modular construction cost (with site work and contingency)');     lab('B160','Stick-built TDC'); lab('B168','Developer gap per unit (modular)'); lab('B159','Stick-built construction cost (land bank baseline, with site work and contingency)'); lab('B169','Modular = the construction method set in the explorer. The appraisal stays at the value set for the base building. DOH excludes 3+ unit structures.')
    # section 10
    for t,(c_,r_) in {'Developer fee':('D',172),'Appraised value':('E',172),'Total development cost':('L',172),'Buyer assistance (TTO + DOH homebuyer)':('O',172)}.items(): lab(f'{c_}{r_}',t)
    lab('B176','Developer fee deferred (0%)'); lab('B177','Developer fee 5%'); lab('B185','Hard cost -10% and fee 5%'); lab('B183','AHP award counted (needs a member bank)'); lab('B181','DOH $200K per unit')
    for i in range(13):
        r=173+i
        for col,e in [('C','SCR(%d).hardMult'),('D','SCR(%d).T.shared.fee'),('E','SCR(%d).R.fmv'),('F','SCR(%d).T.shared.ownerAMI'),('G','SCR(%d).T.shared.rate'),
                      ('H','SCR(%d).T.shared.dohUnit*SCR(%d).T.shared.units'),('I','SCR(%d).T.shared.hdf'),('J','(SCR(%d).T.shared.inc.ahp?SCR(%d).T.shared.ahp:0)'),('K','SCR(%d).T.shared.htcc/SCR(%d).T.shared.htccHomes'),
                      ('L','SCR(%d).R.tdc'),('M','SCR(%d).R.capacity'),('N','Math.min(SCR(%d).R.capacity,SCR(%d).R.fmv)'),('O','SCR(%d).R.ttoAll+SCR(%d).R.hbAll'),('P','SCR(%d).R.remain')]:
            n(f'{col}{r}',e.replace('%d',str(i)))
    lab('B186','Each row reruns the explorer model with one change from the scenario as set. A remaining gap of $0 means the stack covers it under that assumption.')
# ---------- Compare labels
for r,t in [(8,'Appraised value used'),(9,"Buyer's price as % of the existing-building median"),(10,"Buyer gap (appraisal minus the buyer's own price)"),(15,'Development cost before developer fee'),(16,'Total development cost (incl. developer fee)'),(17,'Developer gap per building (TDC minus gross sales price)'),(18,'Developer gap per unit'),(19,"Buyer gap (sales price minus the buyer's own price)"),(20,'Remaining developer gap after the funding stack')]:
    T_('Compare',f'B{r}',t)
# ---------- Market Data (rows the explorer controls)
for j,city in enumerate(CITIES):
    col="CDE"[j]
    for r,e in [(24,'2*c.ami50*0.6'),(25,'2*c.ami50*0.8'),(26,'c.mkt[1]'),(27,'c.mkt[2]'),(28,'c.mkt[3]'),(29,'2*c.ami50*s.brFactor[2]*0.6*s.rentShare/12'),(31,'c.med[0]')]:
        N('Market Data',f'{col}{r}',e,city)
MD='Market Data'
T_(MD,'B42','NEW HAVEN DEMAND: RKG ASSOCIATES HOUSING STUDY (for the New Haven Land Bank)')
T_(MD,'B43','Source: RKG Associates, New Haven Housing and Economic Development study (HUD income limits, ACS 2022, RKG calculations). Shared with the team by the land bank, Nov 2, 2026.')
T_(MD,'B45','Median sale price, New Haven'); T_(MD,'C44','2015'); T_(MD,'D44','2024'); T_(MD,'E44','Change')
M.append(dict(s=MD,r='C45',e='160000',c=None,t='n')); M.append(dict(s=MD,r='D45',e='337000',c=None,t='n')); M.append(dict(s=MD,r='E45',e='337000/160000-1',c=None,t='n'))
T_(MD,'B47','Ownership homes short (-) or in surplus (+), by AMI'); T_(MD,'C47','FHA lending'); T_(MD,'D47','Conventional')
for i,(l,a,b) in enumerate([('Under 50% AMI',-1802,-945),('50-80% AMI',1389,3289),('80-100% AMI',-468,373),('Above 100% AMI',-1803,-2717)]):
    T_(MD,'B%d'%(48+i),l); M.append(dict(s=MD,r='C%d'%(48+i),e=str(a),c=None,t='n')); M.append(dict(s=MD,r='D%d'%(48+i),e=str(b),c=None,t='n'))
T_(MD,'B53','Affordable purchase price without rental income (single family)'); T_(MD,'C53','Household income'); T_(MD,'D53','FHA'); T_(MD,'E53','Conventional')
for i,(l,inc,f,cv) in enumerate([('80% AMI',83650,205979,239181),('100% AMI',104500,257320,298798),('120% AMI',125400,308784,358557)]):
    T_(MD,'B%d'%(54+i),l)
    for col,v in zip('CDE',[inc,f,cv]): M.append(dict(s=MD,r='%s%d'%(col,54+i),e=str(v),c=None,t='n'))
T_(MD,'B58','Rental shortages at both ends: 7,283 homes for households under 30% AMI and 5,671 above 120% AMI. Elm City Communities closed its voucher waitlist May 1, 2026. The ownership shortage above 100% AMI backs the 100-120% AMI buyer target.')
# ---------- Charts labels
S_('Charts','B318','"H. One building: what a "+Math.round(RF_S.ownerAMI*100)+"% AMI buyer pays, the buyer gap and the developer gap (deck: prefab slide)"')
S_('Charts','B336','"I. New Haven: illustrative sources for a "+usdK(RNH.tdc)+" building; developer gap "+usdK(RNH.devGap)+", buyer gap "+usdK(Math.max(0,RNH.salesTotal-RNH.financedAll))+" (deck: prefab slide)"')
T_('Charts','B5','A. Price a buyer can carry vs. the appraised value (deck: affordability curve slide)')
for i,cty in enumerate(CITIES): T_('Charts',f'B{8+2*i}',f'{cty} appraised value')
T_('Charts','B304','Buyer reaches the appraised value (with rent)')
# ---------- Chart H: three stacked series that add up to TDC (buyer's own money, buyer gap, developer gap); never one combined "subsidy needed"
T_('Charts','B320',"Buyer's own mortgage + down payment"); T_('Charts','B321','Buyer gap (buyer-side subsidies)'); T_('Charts','B322','Developer gap (developer-side subsidies)')
for j,cty in enumerate(CITIES):
    col="CDE"[j]; q="'"+cty+"'" if ' ' in cty else cty
    N('Charts',f'{col}320','R.financedAll',cty); N('Charts',f'{col}321','Math.max(0,R.salesTotal-R.financedAll)',cty); N('Charts',f'{col}322','R.devGap',cty)
T_('Charts','B125','Modular developer gap per unit')
T_('Charts','B323','The three bars add up to total development cost. The buyer gap is filled only by buyer-side money and the developer gap only by developer-side money; the two are never added together.')
import zipfile, re, os
def patch_chart_h(path):
    zin=zipfile.ZipFile(path); items=[(i,zin.read(i.filename)) for i in zin.infolist()]; zin.close()
    out=[]
    for info,data in items:
        if info.filename=='xl/charts/chart17.xml':
            x=data.decode()
            if '$B$322' not in x:
                x=x.replace('Per building: buyer pays vs. subsidy','Per building: buyer pays, buyer gap and developer gap')
                sers=re.findall(r'<c:ser>.*?</c:ser>',x,flags=re.S); assert len(sers)==2
                s2=sers[1].replace('<c:v>Subsidy needed</c:v>','<c:v>Buyer gap (buyer-side subsidies)</c:v>').replace('B5452F','7A5BB5')
                s3=sers[1].replace('<c:idx val="1"/><c:order val="1"/>','<c:idx val="2"/><c:order val="2"/>').replace('$B$321','$B$322').replace('$C$321:$E$321','$C$322:$E$322')\
                    .replace('<c:v>Subsidy needed</c:v>','<c:v>Developer gap (developer-side subsidies)</c:v>').replace('{00000001-','{00000002-')
                s3=re.sub(r'(<c:numCache>.*?<c:ptCount val="3"/>).*?(</c:numCache>)',r'\1\2',s3,flags=re.S)
                s2=re.sub(r'(<c:numCache>.*?<c:ptCount val="3"/>).*?(</c:numCache>)',r'\1\2',s2,flags=re.S)
                x=x.replace(sers[1],s2+s3,1)
                if '<c:legend>' not in x:
                    x=x.replace('</c:plotArea>','</c:plotArea><c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>',1)
            data=x.encode()
        out.append((info,data))
    tmp=path+'.tmp'; zo=zipfile.ZipFile(tmp,'w',zipfile.ZIP_DEFLATED)
    for info,data in out: zo.writestr(info,data)
    zo.close(); os.replace(tmp,path)
patch_chart_h('sub/cmc2_template.xlsx')
json.dump({"slots":M,"readmeLog":35},open('sub/cmc2_map.json','w'))
print(len(M))
