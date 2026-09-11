"use strict";
const XLSX = require("./node_modules/xlsx");

const rules = [
  // ── RULES THAT DIFFER BETWEEN UPGRADE AND CONVERSION ─────────────────────
  { rule:"Note 2438131 — MATNR_LONG BAPI/RFC", trigger:"note/fields contain 2438131", condition:"Any", upgrade:"False Positive", conversion:"Needs Remediation", notes:"Already on S/4HANA — _LONG fields already in use. No code change required." },
  { rule:"Note 2438006 — RFC/BAPI Long Fields", trigger:"note = 2438006", condition:"Any", upgrade:"False Positive", conversion:"Needs Remediation", notes:"Long field parameters already in use on S/4HANA." },
  { rule:"Note 2628704 — S/4HANA On-Premise 1809", trigger:"note = 2628704", condition:"Any", upgrade:"False Positive", conversion:"Needs Remediation", notes:"Already on 1809+ — finding no longer applicable." },
  { rule:"Note 2669857 — BAPI Object Number parameter", trigger:"note = 2669857", condition:"Any", upgrade:"False Positive", conversion:"Needs Remediation", notes:"BAPI available on S/4HANA — check not required." },
  { rule:"Note 2689873 — Deprecated DTEL (not CHAR02/CHAR05)", trigger:"note = 2689873, refObjType=DTEL, refObj not CHAR02/CHAR05", condition:"refObjType = DTEL, refObj ≠ CHAR02/CHAR05", upgrade:"False Positive", conversion:"Needs Remediation", notes:"Deprecated DTEL no longer flagged in upgrade scenario." },
  { rule:"Note 2296016 — Deprecated Objects (nr1 set)", trigger:"note = 2296016, not literal, refObj/_cs in nr1 set", condition:"WEEK_DAY / KNC1 / LFC1 / VBUP / EKPO / KEKO / MCHBH / SHLP / etc.", upgrade:"False Positive", conversion:"Needs Remediation", notes:"Already on S/4HANA — deprecated object available or no longer flagged." },
  { rule:"S/4H: Field Length Extensions prio1/2 (non-RFC)", trigger:"Title contains FIELD LENGTH EXTENSIONS, prio1/2, not RFC-FUNCTION PARAMETER", condition:"Priority 1 or 2, not RFC parameter reference", upgrade:"False Positive", conversion:"Needs Remediation", notes:"Already on S/4HANA — FLE finding not applicable in upgrade scenario." },
  { rule:"S/4H: Simplified Objects — FUNC found in TFDIR", trigger:"Simplified Objects check, FUNC/FUGR type, FM found in tfdirSet", condition:"FUNC refObjType, found in TFDIR lookup", upgrade:"False Positive", conversion:"Needs Remediation", notes:"Object already available on S/4HANA." },
  { rule:"S/4H: Simplified Objects — SAP-standard FM (not Y/Z namespace)", trigger:"Simplified Objects check, FUNC type, not in tfdirSet, SAP-standard namespace", condition:"FM name NOT starting with Y/Z or /namespace/", upgrade:"False Positive", conversion:"Needs Remediation", notes:"SAP-standard FM exists in real TFDIR on S/4HANA." },
  { rule:"S/4H: Simplified Objects — DDIC type in TADIR, non-RFC, prio3", trigger:"Simplified Objects check, DTEL/DOMA/TTYP type, found in TADIR, not RFC, prio3", condition:"DTEL/DOMA/TTYP refObjType, found in TADIR, Priority 3", upgrade:"Can be ignored", conversion:"Needs Remediation", notes:"Object available but low severity in upgrade." },
  { rule:"HCA default prio1/2 (no specific rule matched)", trigger:"not isS4H, no remType set by any rule, priority 1 or 2", condition:"Priority 1 or 2, HCA check (not S/4H), no note/pattern match", upgrade:"Optional", conversion:"Mandatory", notes:"Objects already exist in S/4HANA env — relax Mandatory to Optional. Exception: CRITICAL STATEMENTS title stays Mandatory in both modes." },
  // ── RULES SAME IN BOTH MODES ──────────────────────────────────────────────
  { rule:"Syntax / Internal Error", trigger:"checkMsg contains INTERNAL ERROR / SYNTAX ERROR / DOES NOT EXIST / SCAN ERROR / MISSING INCLUDE / SYMBOL...NOT FOUND", condition:"Any (except AQQU/AQSG/IDOC obj types → Fit Gap)", upgrade:"Syntax Error", conversion:"Syntax Error", notes:"Returns immediately." },
  { rule:"SELECT on MLHD/MLCR/MLIT/BSEG/KONV/T881", trigger:"DB OPERATION SELECT FOUND, refObj in {MLHD,MLCR,MLIT,BSEG,KONV,T881}", condition:"", upgrade:"Fit Gap", conversion:"Fit Gap", notes:"Material Ledger tables replaced by MLDOC/MLDOCCCS." },
  { rule:"SELECT on BPEG/GLPCP/CKMLCR", trigger:"DB OPERATION SELECT FOUND, refObj in {BPEG,GLPCP,CKMLCR}", condition:"", upgrade:"Fit Gap", conversion:"Fit Gap", notes:"" },
  { rule:"SELECT on GLPCA/GLPCT", trigger:"DB OPERATION SELECT FOUND, refObj in {GLPCA,GLPCT}", condition:"", upgrade:"False Positive", conversion:"False Positive", notes:"" },
  { rule:"SELECT on SKA1/SKB1/T004", trigger:"DB OPERATION SELECT FOUND, refObj in {SKA1,SKB1,T004}", condition:"", upgrade:"False Positive", conversion:"False Positive", notes:"" },
  { rule:"SELECT on VBFA", trigger:"DB OPERATION SELECT FOUND, refObj = VBFA", condition:"", upgrade:"False Positive", conversion:"False Positive", notes:"Primary key changed in VBFA." },
  { rule:"CURSOR on VBUK/VBUP", trigger:"DB OPERATION CURSOR, checkMsg contains VBUK or VBUP", condition:"", upgrade:"Needs Remediation", conversion:"Needs Remediation", notes:"" },
  { rule:"CURSOR on VBFA", trigger:"DB OPERATION CURSOR, refObj = VBFA", condition:"", upgrade:"False Positive", conversion:"False Positive", notes:"" },
  { rule:"IDOC CHECK", trigger:"Check Title/Message contains S/4HANA: IDOC CHECK", condition:"", upgrade:"Fit Gap", conversion:"Fit Gap", notes:"" },
  { rule:"ORDER BY missing prio1/2", trigger:"checkMsg contains ORDER BY + (WRITE/EXIT/LEAVE/RETURN/SELECT/ENDSELECT)", condition:"Not LOOP AT EMPTY in checkMsg", upgrade:"Mandatory", conversion:"Mandatory", notes:"" },
  { rule:"LOOP AT + ORDER BY/EQUALITY/etc — prio1/2", trigger:"LOOP AT + (ORDER BY/EQUALITY TEST/AT/ON/MODIFY/EMPTY)", condition:"Priority 1 or 2", upgrade:"Mandatory", conversion:"Mandatory", notes:"" },
  { rule:"LOOP AT + ORDER BY/EQUALITY/etc — prio3", trigger:"LOOP AT + (ORDER BY/EQUALITY TEST/AT/ON/MODIFY/EMPTY)", condition:"Priority 3", upgrade:"Optional", conversion:"Optional", notes:"" },
  { rule:"LOOP AT + WRITE/EXIT/RETURN/LEAVE — prio1/2", trigger:"LOOP AT + (WRITE/EXIT/RETURN/LEAVE)", condition:"Priority 1 or 2", upgrade:"Mandatory", conversion:"Mandatory", notes:"Sort statement required before LOOP." },
  { rule:"LOOP AT + WRITE/EXIT/RETURN/LEAVE — prio3", trigger:"LOOP AT + (WRITE/EXIT/RETURN/LEAVE)", condition:"Priority 3", upgrade:"Optional", conversion:"Optional", notes:"" },
  { rule:"WRITE IN LOOP FOR — prio1/2", trigger:"checkMsg contains WRITE IN LOOP FOR", condition:"Priority 1 or 2", upgrade:"Mandatory", conversion:"Mandatory", notes:"" },
  { rule:"WRITE IN LOOP FOR — prio3", trigger:"checkMsg contains WRITE IN LOOP FOR", condition:"Priority 3", upgrade:"Optional", conversion:"Optional", notes:"" },
  { rule:"ALV CALL AT", trigger:"checkMsg contains ALV CALL AT", condition:"Any priority", upgrade:"Optional", conversion:"Optional", notes:"ALV statement correct. No change required." },
  { rule:"READ TABLE INDEX (not INDEX 1)", trigger:"checkMsg contains READ TABLE + INDEX, not INDEX 1", condition:"Priority 1/2/3 → Mandatory", upgrade:"Mandatory", conversion:"Mandatory", notes:"Sort statement required." },
  { rule:"READ TABLE INDEX 1", trigger:"checkMsg contains READ TABLE + INDEX 1", condition:"Any priority", upgrade:"Optional", conversion:"Optional", notes:"" },
  { rule:"CONCATENATE IN LOOP", trigger:"checkMsg contains CONCATENATE IN LOOP", condition:"", upgrade:"Mandatory", conversion:"Mandatory", notes:"" },
  { rule:"TYPE/LENGTH CONFLICT / WRITE TO / TRANSFER / etc.", trigger:"TYPE CONFLICT / LENGTH CONFLICT / WRITE TO / TRANSFER / REPLACE / OFFSET / STRUCTURE COMPONENT / READ DATASET", condition:"No remType set yet", upgrade:"Needs Remediation", conversion:"Needs Remediation", notes:"" },
  { rule:"SEARCH...IN ITAB FOR RESULT", trigger:"checkMsg contains SEARCH ... IN ITAB FOR RESULT", condition:"", upgrade:"Mandatory", conversion:"Mandatory", notes:"" },
  { rule:"Note 2993220 — Classical PCA", trigger:"note/fields contain 2993220", condition:"", upgrade:"False Positive", conversion:"False Positive", notes:"Classical PCA in S/4HANA compat scope." },
  { rule:"Note 2602107 — Compatibility Views", trigger:"note = 2602107, not DML/Query, refObj in compat view list", condition:"BSAD/BSAK/BSAS/BSID/BSIK/BSIS/COEP/COSP/GLT0/FAGLFLEXT/etc.", upgrade:"False Positive", conversion:"False Positive", notes:"Select works with CDS Views." },
  { rule:"Note 2610650 — Amount Field Length prio3", trigger:"note/fields contain 2610650, not USED BY RFC, prio3", condition:"Priority 3", upgrade:"Can be ignored", conversion:"Can be ignored", notes:"" },
  { rule:"Note 2610650 — Amount Field Length prio1/2", trigger:"note/fields contain 2610650, not USED BY RFC, prio1/2", condition:"Priority 1 or 2", upgrade:"Needs Remediation", conversion:"Needs Remediation", notes:"" },
  { rule:"Note 2768887 — SD Billing Draft", trigger:"note/fields contain 2768887, not DML/Query", condition:"", upgrade:"Needs Remediation", conversion:"Needs Remediation", notes:"" },
  { rule:"Note 2270387 — Asset Accounting", trigger:"note/fields contain 2270387, not DML", condition:"", upgrade:"Needs Remediation", conversion:"Needs Remediation", notes:"" },
  { rule:"Note 2431747 — GL (BSEG/CSKA/TKA02/T881/etc.)", trigger:"note contain 2431747, not DML", condition:"refObj in {BSEG,CSKA,TKA02,T881,FAGL_ACTIVEC,T882G,CSKB,FAGL_LEDGER_SCEN}", upgrade:"Needs Remediation", conversion:"Needs Remediation", notes:"" },
  { rule:"Note 2431747 — GL (SKA1/SKB1/T001)", trigger:"note contain 2431747, not DML", condition:"refObj in {SKA1,SKB1,T001,RPL_S130_CHANGE_IN_UPDATE_TASK}", upgrade:"False Positive", conversion:"False Positive", notes:"" },
  { rule:"Note 2431747 — GL (AQQU/AQSG)", trigger:"note contain 2431747, not DML", condition:"refObj in {AQQU,AQSG}", upgrade:"Fit Gap", conversion:"Fit Gap", notes:"" },
  { rule:"Note 2365665/2535093 — Seasonal Fields DTEL", trigger:"note 2365665 or 2535093, refObjType=DTEL", condition:"", upgrade:"Needs Remediation", conversion:"Needs Remediation", notes:"" },
  { rule:"Note 2365665/2535093 — Seasonal Fields TABL/MARA", trigger:"note 2365665 or 2535093, refObjType=TABL, refObj=MARA", condition:"", upgrade:"Fit Gap", conversion:"Fit Gap", notes:"isFitGapDelta=true" },
  { rule:"Note 2215424 — Material Number Length prio1/2", trigger:"note = 2215424, prio1/2", condition:"", upgrade:"Needs Remediation", conversion:"Needs Remediation", notes:"" },
  { rule:"Note 2215424 — Material Number Length prio3", trigger:"note = 2215424, prio3", condition:"", upgrade:"Can be ignored", conversion:"Can be ignored", notes:"" },
  { rule:"Note 2389136 — Cost Element (CSKA/CSKB/CSKU)", trigger:"note = 2389136, not Query", condition:"refObj in {CSKA,CSKB,CSKU}", upgrade:"Needs Remediation", conversion:"Needs Remediation", notes:"" },
  { rule:"Note 2389136 — Cost Element (SKA1/SKB1/SKAT/T004)", trigger:"note = 2389136, not Query", condition:"refObj in {SKA1,SKB1,SKAT,T004}", upgrade:"False Positive", conversion:"False Positive", notes:"" },
  { rule:"Note 2389136 — Cost Element TKSKA", trigger:"note = 2389136, not Query, refObj=TKSKA", condition:"", upgrade:"Fit Gap", conversion:"Fit Gap", notes:"" },
  { rule:"Note 2226131 — BP (FD06/FK06/XD01/etc.)", trigger:"note = 2226131, not literal", condition:"refObj in fitGapTrans set (FD06/FK06/XD01/XD02/XK01/FK01/FK02/etc.)", upgrade:"Fit Gap", conversion:"Fit Gap", notes:"T-code redirected to transaction BP." },
  { rule:"Note 2226131 — BP TABL", trigger:"note = 2226131, not literal, refObjType=TABL", condition:"", upgrade:"Needs Remediation", conversion:"Needs Remediation", notes:"" },
  { rule:"Note 2226131 — BP (KNA1/LFA1/SHLP)", trigger:"note = 2226131, not literal", condition:"refObj in {KNA1,LFA1,SHLP}", upgrade:"False Positive", conversion:"False Positive", notes:"" },
  { rule:"Note 1803189 — Obsolete PO APIs FUNC", trigger:"note = 1803189, refObjType=FUNC or contains BAPI_PO_GETDETAIL", condition:"", upgrade:"Needs Remediation", conversion:"Needs Remediation", notes:"" },
  { rule:"Note 1803189 — Obsolete PO ME53/ME23/ME26", trigger:"note = 1803189, TRAN refObjType, refObj in {ME53,ME23,ME26}", condition:"", upgrade:"Needs Remediation", conversion:"Needs Remediation", notes:"" },
  { rule:"Note 1803189 — Obsolete PO other TRAN", trigger:"note = 1803189, TRAN refObjType, other", condition:"", upgrade:"Fit Gap", conversion:"Fit Gap", notes:"" },
  { rule:"Note 2265093 — Obsolete T-codes (needsRem)", trigger:"note = 2265093, not DML", condition:"refObj in {FD03,FK03,MAP3,MK03,VAP3,VD03,XD03,XK06,XK03}", upgrade:"Needs Remediation", conversion:"Needs Remediation", notes:"" },
  { rule:"Note 2265093 — Obsolete T-codes (fitGap)", trigger:"note = 2265093, not DML", condition:"refObj in large fitGap set (FD01/VD01/XK01/LFA1/KNA1/etc.)", upgrade:"Fit Gap", conversion:"Fit Gap", notes:"" },
  { rule:"Note 2689873 — Deprecated FUNC", trigger:"note = 2689873, refObjType=FUNC or _eq(FUNC)", condition:"", upgrade:"Fit Gap", conversion:"Fit Gap", notes:"DEPRECATED FM in use." },
  { rule:"Note 2689873 — Deprecated CHAR02/CHAR05 DTEL", trigger:"note = 2689873, refObjType=DTEL, refObj in {CHAR02,CHAR05}", condition:"", upgrade:"False Positive", conversion:"False Positive", notes:"" },
  { rule:"Note 2689873 — Deprecated SICH TRAN", trigger:"note = 2689873, refObjType=TRAN, refObj=SICH", condition:"", upgrade:"False Positive", conversion:"False Positive", notes:"" },
  { rule:"Note 2689873 — Deprecated DOMA/TTYP/MSAG", trigger:"note = 2689873, _eq(DOMA/TTYP/MSAG)", condition:"", upgrade:"Needs Remediation", conversion:"Needs Remediation", notes:"" },
  { rule:"Note 2689873 — Deprecated PARA/TRAN/TABL/NC10", trigger:"note = 2689873, _eq(PARA/TRAN/TABL/NC10/V_NMARC)", condition:"", upgrade:"Fit Gap", conversion:"Fit Gap", notes:"" },
  { rule:"Note 2689873 — Deprecated NLEI", trigger:"note = 2689873, _eq(NLEI)", condition:"", upgrade:"Fit Gap", conversion:"Fit Gap", notes:"IS-H not planned for S/4HANA conversion." },
  { rule:"Note 2198647 — Sales Doc Status VBFA", trigger:"note = 2198647, not DML/Query, _eq(VBFA)", condition:"", upgrade:"False Positive", conversion:"False Positive", notes:"" },
  { rule:"Note 2198647 — Sales Doc Status VBUK/VBUP/etc.", trigger:"note = 2198647, not DML/Query", condition:"refObj in {VBUK,VBUP,VBAKUK,LIKPUK,LIPSUP,RVVBTYP}", upgrade:"Needs Remediation", conversion:"Needs Remediation", notes:"" },
  { rule:"Note 2203518 — WBS PK (WBRF)", trigger:"note = 2203518, checkMsg contains WBRF", condition:"", upgrade:"False Positive", conversion:"False Positive", notes:"" },
  { rule:"Note 2203518 — WBS PK (other)", trigger:"note = 2203518", condition:"", upgrade:"Needs Remediation", conversion:"Needs Remediation", notes:"" },
  { rule:"Note 1804812 — MB Transactions", trigger:"note = 1804812", condition:"MB1C/MCHA/MCHB/MB1B/MBST/MB01/MB02/MBUS/MB04/MB05/etc.", upgrade:"Fit Gap", conversion:"Fit Gap", notes:"MB transactions replaced by MIGO." },
  { rule:"Note 2268085 — MRP Live (MD04/MD4C/MDBS)", trigger:"note = 2268085, not literal, _eq(MD04/MD4C/MDBS)", condition:"", upgrade:"False Positive", conversion:"False Positive", notes:"" },
  { rule:"Note 2268085 — MRP Live (MD_STOCK_REQUIREMENTS)", trigger:"note = 2268085, not literal, contains MD_STOCK_REQUIREMENTS_LIST_API", condition:"", upgrade:"Needs Remediation", conversion:"Needs Remediation", notes:"" },
  { rule:"Note 2268085 — MRP Live (other)", trigger:"note = 2268085, not literal, unmatched", condition:"", upgrade:"Fit Gap", conversion:"Fit Gap", notes:"" },
  { rule:"Note 2227014 — Credit Management (Simplified Objects block)", trigger:"note = 2227014, isS4H Simplified Objects check", condition:"Overrides TADIR lookup", upgrade:"Fit Gap", conversion:"Fit Gap", notes:"" },
  { rule:"Note 2469385 — ISM/SAP Media (ISM_CUSTOMER_PI/A386/etc.)", trigger:"note = 2469385, not literal", condition:"ISM_CUSTOMER_PI / A386 / JPTTITLERELATTR / refObj=TABL", upgrade:"Fit Gap", conversion:"Fit Gap", notes:"" },
  { rule:"Note 2469385 — ISM/SAP Media (J1ST/JSE4/etc.)", trigger:"note = 2469385, not literal", condition:"_eq(J1ST/JSE4/JSE2/JHC2/JKK1/JSB1/JSB2/TBDLS)", upgrade:"False Positive", conversion:"False Positive", notes:"" },
  { rule:"Note 2469385 — ISM/SAP Media (other)", trigger:"note = 2469385, not literal, unmatched", condition:"", upgrade:"Needs Remediation", conversion:"Needs Remediation", notes:"" },
  { rule:"Note 2223144 — Foreign Trade SD/MM", trigger:"note = 2223144", condition:"", upgrade:"Fit Gap", conversion:"Fit Gap", notes:"" },
  { rule:"Note 2340247 — WM WQ transactions", trigger:"note = 2340247, refObj in {WC01,WQ01..WQ21,WVAL}", condition:"", upgrade:"Fit Gap", conversion:"Fit Gap", notes:"" },
  { rule:"Note 2340247 — WM other", trigger:"note = 2340247, other", condition:"", upgrade:"False Positive", conversion:"False Positive", notes:"" },
  { rule:"Note 2217202 — LE Replication TTYP", trigger:"note = 2217202, not DML, refObjType=TTYP", condition:"", upgrade:"Needs Remediation", conversion:"Needs Remediation", notes:"" },
  { rule:"Note 2217202 — LE Replication other", trigger:"note = 2217202, not DML, other", condition:"", upgrade:"Fit Gap", conversion:"Fit Gap", notes:"" },
  { rule:"Note 2217205/2217206 — Occupational Health/EC", trigger:"note = 2217205 or 2217206", condition:"", upgrade:"Fit Gap", conversion:"Fit Gap", notes:"" },
  { rule:"Note 2224144 — /BEV* (falsePos set)", trigger:"note = 2224144, refObj in falsePos set", condition:"CMM_MTM_ANTCP.../KOMG/KOMP/VBAP/etc.", upgrade:"False Positive", conversion:"False Positive", notes:"" },
  { rule:"Note 2224144 — /BEV* MSEG", trigger:"note = 2224144, refObj=MSEG", condition:"", upgrade:"Fit Gap", conversion:"Fit Gap", notes:"" },
  { rule:"Note 2224144 — /BEV* /BEV DTEL/DOMA", trigger:"note = 2224144, contains /BEV, refObjType=DTEL/DOMA", condition:"", upgrade:"Needs Remediation", conversion:"Needs Remediation", notes:"" },
  { rule:"Note 2224144 — /BEV* other /BEV", trigger:"note = 2224144, contains /BEV, unmatched", condition:"", upgrade:"Fit Gap", conversion:"Fit Gap", notes:"" },
  { rule:"Note 2354101 — WB Transactions", trigger:"note = 2354101", condition:"", upgrade:"Fit Gap", conversion:"Fit Gap", notes:"WB01/WB02/WB03 no longer accessible." },
  { rule:"Note 2468869 — Deleted DB Objects", trigger:"note = 2468869", condition:"", upgrade:"Needs Remediation", conversion:"Needs Remediation", notes:"" },
  { rule:"Note 2468834 — Prepayment/External Interfacing", trigger:"note = 2468834", condition:"", upgrade:"Needs Remediation", conversion:"Needs Remediation", notes:"" },
  { rule:"Note 2209696 — VBBS/VBBE Structures", trigger:"note = 2209696", condition:"", upgrade:"Fit Gap", conversion:"Fit Gap", notes:"" },
  { rule:"Note 2371631 — Tax Functionality", trigger:"note = 2371631", condition:"", upgrade:"False Positive", conversion:"False Positive", notes:"" },
  { rule:"Note 2206980 — ERP Product Availability DML+MARC", trigger:"note = 2206980, isDML, refObj=MARC", condition:"", upgrade:"Fit Gap", conversion:"Fit Gap", notes:"" },
  { rule:"Note 2206980 — ERP Product Availability other", trigger:"note = 2206980, other", condition:"", upgrade:"Needs Remediation", conversion:"Needs Remediation", notes:"" },
  { rule:"Note 2296016 — Deprecated Objects (fp1 set)", trigger:"note = 2296016, not literal, refObj in fp1 set", condition:"LAST_DAY_OF_MONTHS/BSAK/GLT0/MARA/MARD/MSEG/MBEW/etc.", upgrade:"False Positive", conversion:"False Positive", notes:"" },
  { rule:"Note 2296016 — Deprecated Objects (fg1 set)", trigger:"note = 2296016, not literal, refObj in fg1 set", condition:"IBSP/DTEL/TTYP/DOMA/SER01/SER02/NLEI/etc.", upgrade:"Fit Gap", conversion:"Fit Gap", notes:"" },
  { rule:"Transactions in Literals (fp set)", trigger:"isLiteral, _eq in {FORM,UPDA,J1ST,CPUB,SICH,GP12N,FORW,UPGRADE,WAST,KNC1,LFC1,SDIN}", condition:"", upgrade:"False Positive", conversion:"False Positive", notes:"" },
  { rule:"Transactions in Literals TRAN or default", trigger:"isLiteral, _eq(TRAN) or unmatched", condition:"", upgrade:"Fit Gap", conversion:"Fit Gap", notes:"" },
  { rule:"Note 2877717 — J_1IMOCUST/J_1IMOVEND", trigger:"note = 2877717", condition:"", upgrade:"Needs Remediation", conversion:"Needs Remediation", notes:"" },
  { rule:"Note 2437332/2445654 — DTEL ENHO", trigger:"note 2437332 or 2445654, refObjType=DTEL, objType=ENHO", condition:"", upgrade:"False Positive", conversion:"False Positive", notes:"" },
  { rule:"Note 2437332/2445654 — DTEL", trigger:"note 2437332 or 2445654, _eq(DTEL)", condition:"", upgrade:"Needs Remediation", conversion:"Needs Remediation", notes:"" },
  { rule:"Note 2354768 — Material Ledger ACDOCA", trigger:"note = 2354768", condition:"", upgrade:"Fit Gap", conversion:"Fit Gap", notes:"" },
  { rule:"Note 2227532/2227579 — Subcontracting FUNC", trigger:"note 2227532 or 2227579, refObjType=FUNC", condition:"", upgrade:"Needs Remediation", conversion:"Needs Remediation", notes:"" },
  { rule:"Note 2227532/2227579 — Subcontracting other", trigger:"note 2227532 or 2227579, other", condition:"", upgrade:"Fit Gap", conversion:"Fit Gap", notes:"" },
  { rule:"Note 2378796 — MARC STAWN/EXPME DML+MARC", trigger:"note = 2378796, isDML, refObj=MARC", condition:"", upgrade:"Fit Gap", conversion:"Fit Gap", notes:"" },
  { rule:"Note 2378796 — MARC STAWN/EXPME other", trigger:"note = 2378796, other", condition:"", upgrade:"Needs Remediation", conversion:"Needs Remediation", notes:"" },
  { rule:"Note 2337368 — Inventory Valuation", trigger:"note = 2337368", condition:"", upgrade:"Needs Remediation", conversion:"Needs Remediation", notes:"" },
  { rule:"Note 2522971 — Segment field length prio1/2", trigger:"note = 2522971, not USED BY RFC, prio1/2", condition:"", upgrade:"Needs Remediation", conversion:"Needs Remediation", notes:"" },
  { rule:"Note 2522971 — Segment field length prio3", trigger:"note = 2522971, not USED BY RFC, prio3", condition:"", upgrade:"Can be ignored", conversion:"Can be ignored", notes:"" },
  { rule:"Note 2628699 — Extended fields _LONG", trigger:"note = 2628699", condition:"", upgrade:"Needs Remediation", conversion:"Needs Remediation", notes:"" },
  { rule:"Note 3211383 — MRP MD_STOCK_REQUIREMENTS", trigger:"note = 3211383, contains MD_STOCK_REQUIREMENTS_LIST_API", condition:"", upgrade:"Needs Remediation", conversion:"Needs Remediation", notes:"" },
  { rule:"Note 3211383 — MRP other", trigger:"note = 3211383, other", condition:"", upgrade:"False Positive", conversion:"False Positive", notes:"" },
  { rule:"Note 2438110 — Syntactically incompatible FUNC", trigger:"note = 2438110, refObjType=FUNC or _eq(FUNC)", condition:"", upgrade:"Needs Remediation", conversion:"Needs Remediation", notes:"" },
  { rule:"Note 2438110 — Syntactically incompatible non-FUNC", trigger:"note = 2438110, non-FUNC refObjType", condition:"", upgrade:"False Positive", conversion:"False Positive", notes:"" },
  { rule:"Note 2694441 — ISU Payment FUNC", trigger:"note = 2694441, refObjType=FUNC or objType=FUNC/FUGR", condition:"", upgrade:"Needs Remediation", conversion:"Needs Remediation", notes:"ISU FM — verify availability in target system." },
  { rule:"Note 2694441 — ISU Payment non-FUNC", trigger:"note = 2694441, non-FUNC refObjType", condition:"", upgrade:"False Positive", conversion:"False Positive", notes:"IS-U object available in S/4HANA." },
  { rule:"S/4H: FLE prio3", trigger:"Title contains FIELD LENGTH EXTENSIONS or S/4HANA: FLE, prio3", condition:"Priority 3", upgrade:"Can be ignored", conversion:"Can be ignored", notes:"" },
  { rule:"S/4H: FLE prio1/2 RFC-FUNCTION PARAMETER", trigger:"FLE check, prio1/2, checkMsg contains RFC-FUNCTION PARAMETER", condition:"Priority 1 or 2, RFC param ref", upgrade:"False Positive", conversion:"False Positive", notes:"Informational — no code change." },
  { rule:"S/4H: Readiness Check for SAP Queries", trigger:"Title contains READINESS CHECK FOR SAP QUERIES", condition:"", upgrade:"Fit Gap", conversion:"Fit Gap", notes:"" },
  { rule:"S/4H: Simplified Transactions Literals (fp set)", trigger:"isS4H, title SIMPLIFIED TRANSACTIONS IN LITERALS", condition:"_eq in {UPDA,SICH,FORM,MCHA,MCHB,MCH1,MCHP,FOUR,MR01}", upgrade:"False Positive", conversion:"False Positive", notes:"" },
  { rule:"S/4H: Simplified Transactions Literals (nr set)", trigger:"isS4H, title SIMPLIFIED TRANSACTIONS IN LITERALS", condition:"_eq in {XD03,XK03,FD03,FK03,VD03,VK03,MB03,ME23,ME53,ME26}", upgrade:"Needs Remediation", conversion:"Needs Remediation", notes:"" },
  { rule:"S/4H: Simplified Transactions Literals (other)", trigger:"isS4H, title SIMPLIFIED TRANSACTIONS IN LITERALS, unmatched", condition:"", upgrade:"Fit Gap", conversion:"Fit Gap", notes:"" },
  { rule:"S/4H: Simplified Objects — customer FM orphan (knownOrphans)", trigger:"Simplified Objects, FUNC, not in TFDIR, Y/Z namespace, in knownOrphans list", condition:"DR_GET_COUNTRY_NAME / DATE_TO_DAY / NEXT_WEEK / ISP_GET_MONTH_NAME / etc.", upgrade:"Needs Remediation", conversion:"Needs Remediation", notes:"FM not in S/4HANA — create custom." },
  { rule:"S/4H: Simplified Objects — customer FM orphan (other Y/Z)", trigger:"Simplified Objects, FUNC, not in TFDIR, Y/Z namespace, not knownOrphans", condition:"", upgrade:"Fit Gap", conversion:"Fit Gap", notes:"FM not in S/4HANA — functional help needed." },
  { rule:"S/4H: Simplified Objects — DDIC type in TADIR, RFC ref", trigger:"Simplified Objects, DTEL/DOMA/TTYP/INTF/CLAS, found in TADIR, contains RFC", condition:"", upgrade:"False Positive", conversion:"False Positive", notes:"RFC informational — no code change." },
  { rule:"S/4H: Simplified Objects — INTF/CLAS found in TADIR", trigger:"Simplified Objects, INTF/CLAS type, found in TADIR", condition:"", upgrade:"Needs Remediation", conversion:"Needs Remediation", notes:"" },
  { rule:"S/4H: Simplified Objects — INTF/CLAS NOT in TADIR", trigger:"Simplified Objects, INTF/CLAS type, not in TADIR", condition:"", upgrade:"Fit Gap", conversion:"Fit Gap", notes:"" },
  { rule:"S/4H: Simplified Objects — DTEL/DOMA/TTYP prio1/2 in TADIR", trigger:"Simplified Objects, DTEL/DOMA/TTYP, found in TADIR, not RFC, prio1/2", condition:"Priority 1 or 2", upgrade:"Needs Remediation", conversion:"Needs Remediation", notes:"" },
  { rule:"S/4H: Simplified Objects — DTEL/DOMA/TTYP NOT in TADIR", trigger:"Simplified Objects, DTEL/DOMA/TTYP, not in TADIR", condition:"", upgrade:"Needs Remediation", conversion:"Needs Remediation", notes:"" },
  { rule:"S/4H: DML on simplified table (fallback)", trigger:"isS4H, isDML, title SEARCH FOR DATABASE OPERATIONS, no remType set", condition:"", upgrade:"Fit Gap", conversion:"Fit Gap", notes:"" },
  { rule:"HCA prio3: PROBLEMATIC title + SELECT in message", trigger:"not isS4H, prio3, no remType, title contains PROBLEMATIC, checkMsg contains SELECT", condition:"", upgrade:"Mandatory", conversion:"Mandatory", notes:"" },
  { rule:"HCA prio3: ADBC title", trigger:"not isS4H, prio3, no remType, title contains ADBC", condition:"", upgrade:"Mandatory", conversion:"Mandatory", notes:"" },
  { rule:"S/4H default prio3 (fallback)", trigger:"isS4H, no remType, prio3", condition:"", upgrade:"Can be ignored", conversion:"Can be ignored", notes:"" },
  { rule:"S/4H default prio1/2 (fallback)", trigger:"isS4H, no remType, prio1/2", condition:"", upgrade:"Needs Remediation", conversion:"Needs Remediation", notes:"" },
  { rule:"HCA default prio1/2: CRITICAL STATEMENTS", trigger:"not isS4H, no remType, prio1/2, title contains CRITICAL STATEMENTS", condition:"Priority 1 or 2, title has CRITICAL STATEMENTS", upgrade:"Mandatory", conversion:"Mandatory", notes:"Database Hint — always Mandatory in both modes." },
  { rule:"HCA default prio3 (fallback)", trigger:"not isS4H, no remType, prio3", condition:"", upgrade:"Optional", conversion:"Optional", notes:"" },
  { rule:"HCA default no-prio fallback", trigger:"not isS4H, no remType, priority not 1/2/3", condition:"", upgrade:"Can be ignored", conversion:"Needs Remediation", notes:"" },
];

const HDR = ["#","Rule / Pattern","Trigger Condition","Sub-condition / Object Match","Upgrade Result","Conversion Result","Notes / Rationale"];
const COL_W = [{wch:4},{wch:52},{wch:74},{wch:46},{wch:22},{wch:22},{wch:65}];

function makeSheet(rows) {
  const data = [HDR, ...rows.map((r,i)=>[i+1,r.rule,r.trigger,r.condition,r.upgrade,r.conversion,r.notes])];
  const ws = XLSX.utils.aoa_to_sheet(data);
  ws["!cols"] = COL_W;
  return ws;
}

const diffRules = rules.filter(r => r.upgrade !== r.conversion);
const sameRules = rules.filter(r => r.upgrade === r.conversion);

const wb = XLSX.utils.book_new();

// Sheet 1 — rules that differ (most important)
XLSX.utils.book_append_sheet(wb, makeSheet(diffRules), "Upgrade vs Conversion (Diff)");

// Sheet 2 — all rules
XLSX.utils.book_append_sheet(wb, makeSheet(rules), "All Rules (Complete)");

// Sheet 3 — grouped by impact direction
const toNR   = diffRules.filter(r => r.conversion === "Needs Remediation");
const toMand = diffRules.filter(r => r.conversion === "Mandatory");

const s3 = [
  ["SECTION 1 — Conversion: Needs Remediation  →  Upgrade: False Positive / Optional / Can be ignored"],
  ["These are the rules where conversion requires code change but upgrade does NOT (object already exists on S/4HANA)"],
  [],
  HDR,
  ...toNR.map((r,i)=>[i+1,r.rule,r.trigger,r.condition,r.upgrade,r.conversion,r.notes]),
  [],
  ["SECTION 2 — Conversion: Mandatory  →  Upgrade: Optional"],
  ["HCA default fallback for prio1/2: upgrade relaxes to Optional since objects already in S/4HANA env"],
  [],
  HDR,
  ...toMand.map((r,i)=>[i+1,r.rule,r.trigger,r.condition,r.upgrade,r.conversion,r.notes]),
  [],
  ["SECTION 3 — Same in both modes (reference)"],
  ["These rules produce identical results for upgrade and conversion"],
  [],
  HDR,
  ...sameRules.map((r,i)=>[i+1,r.rule,r.trigger,r.condition,r.upgrade,r.conversion,r.notes]),
];
const ws3 = XLSX.utils.aoa_to_sheet(s3);
ws3["!cols"] = COL_W;
XLSX.utils.book_append_sheet(wb, ws3, "Impact by Direction");

const out = "C:/Users/I755599/Downloads/classification_rules_upgrade_vs_conversion.xlsx";
XLSX.writeFile(wb, out);
console.log("Saved:", out);
console.log("Total rules:", rules.length, "| Differ:", diffRules.length, "| Same:", sameRules.length);
console.log("  Conversion NR → different in upgrade:", toNR.length);
console.log("  Conversion Mandatory → different in upgrade:", toMand.length);
