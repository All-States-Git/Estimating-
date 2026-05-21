/**
 * ASM Preconstruction Pipeline Tracker V9
 * All States Mechanical — Excel workbook generator (Node.js / ExcelJS)
 */
const ExcelJS = require('exceljs');
const path = require('path');

const OUTPUT = path.join(__dirname, 'ASM_Preconstruction_Pipeline_Tracker_V9.xlsx');

// ─── Palette ──────────────────────────────────────────────────────────────────
const C = {
  navy:    'FF003B4C', blue:    'FF2E75B6', ltblue:  'FFD9E2F3', // navy = ASM Teal PMS 548
  dkgrn:   'FF375623', ltgrn:   'FFE2EFDA', dkred:   'FFD03238', // dkred = ASM Red PMS 1797
  ltred:   'FFFFE0E0', orange:  'FFC55A11', ltorang: 'FFFCE4D6',
  dkyel:   'FF7F6000', ltyel:   'FFFFEB9C', purple:  'FF4B0082',
  ltpurp:  'FFEAD5F5', teal:    'FF003B4C', ltteal:  'FFD4EEF1', // teal = ASM Teal PMS 548
  white:   'FFFFFFFF', lgray:   'FFF2F2F2', mgray:   'FFD9D9D9',
  dgray:   'FF58595B', black:   'FF231F20', amber:   'FFFFC000', // dgray = ASM Grey; black = ASM Black
  rust:    'FF7F3300', lrust:   'FFFFCC66',
};

// ─── Style helpers ────────────────────────────────────────────────────────────
function font(opts = {}) {
  return { name: 'Arial', size: opts.size||10, bold: !!opts.bold,
           color: { argb: opts.color||C.black }, italic: !!opts.italic };
}
function fill(argb) {
  return { type: 'pattern', pattern: 'solid', fgColor: { argb } };
}
function align(h='left', v='center', wrap=false) {
  return { horizontal: h, vertical: v, wrapText: wrap };
}
function border(color = C.mgray) {
  const s = { style: 'thin', color: { argb: color } };
  return { left: s, right: s, top: s, bottom: s };
}

function applyHeader(cell, text, opts = {}) {
  cell.value = text;
  cell.font  = font({ size: opts.size||10, bold: true, color: opts.fg||C.white });
  cell.fill  = fill(opts.bg||C.navy);
  cell.alignment = align(opts.h||'center', 'center', opts.wrap !== false);
  cell.border = border();
}

function applyValue(cell, val, opts = {}) {
  cell.value = val;
  cell.font  = font({ size: opts.size||10, bold: !!opts.bold, color: opts.color||C.black });
  if (opts.bg) cell.fill = fill(opts.bg);
  cell.alignment = align(opts.h||'left', 'center', !!opts.wrap);
  cell.border = border();
  if (opts.numFmt) cell.numFmt = opts.numFmt;
}

function titleBar(ws, row, fromCol, toCol, text, opts = {}) {
  ws.mergeCells(row, fromCol, row, toCol);
  const cell = ws.getCell(row, fromCol);
  cell.value = text;
  cell.font  = font({ size: opts.size||14, bold: true, color: C.white });
  cell.fill  = fill(opts.bg||C.navy);
  cell.alignment = align('center', 'center');
  ws.getRow(row).height = opts.height||32;
}

function sectionBar(ws, row, fromCol, toCol, text, opts = {}) {
  ws.mergeCells(row, fromCol, row, toCol);
  const cell = ws.getCell(row, fromCol);
  cell.value = text;
  cell.font  = font({ size: 11, bold: true, color: opts.fg||C.white });
  cell.fill  = fill(opts.bg||C.navy);
  cell.alignment = align('left', 'center');
  cell.border = border();
  ws.getRow(row).height = opts.height||22;
}

// ══════════════════════════════════════════════════════════════════════════════
// SETTINGS TAB
// ══════════════════════════════════════════════════════════════════════════════
const SETTINGS_LISTS = [
  { title: 'PLUMB ESTIMATORS',   col: 1, values: [
    'Patrick','Cristhian','Les','Sam','No Bid'
  ]},
  { title: 'BID STATUS',         col: 3, values: [
    'In Progress','ROM Submitted','Budget Submitted','Bid Submitted','Awarded','Lost','No Bid','On Hold'
  ]},
  { title: 'MARKET SEGMENTS',    col: 5, values: [
    'Correctional','Education','Federal / Government','Food Service',
    'Healthcare','Hotel / Resort','Industrial / Manufacturing','Mission Critical',
    'Multi-Family','Municipal','Professional Office','Religious / Cultural',
    'Retail','Transportation','Warehouse / Distribution'
  ]},
  { title: 'SCOPE TYPES',        col: 7, values: [
    'Plumbing Only','HVAC Only','Full Mechanical','Design-Build','Service','Budget Only'
  ]},
  { title: 'BID STAGES',         col: 9, values: [
    'ROM','Budget','DD','Permit Set','Hard Bid','GMP','Design-Build'
  ]},
  { title: 'DELIVERY METHODS',   col: 11, values: [
    'Hard Bid','GMP','Design-Build','Negotiated','CM at Risk','JOC'
  ]},
  { title: 'LOST REASONS',       col: 13, values: [
    'Price','Relationship','Scope Gap','Schedule Conflict','No Award',
    'Withdrawn by Owner','Competition Too Strong','Not Qualified','Other'
  ]},
  { title: 'GO / NO-GO',         col: 15, values: ['Go','Review','No-Go'] },
  { title: 'GENERAL CONTRACTORS',col: 17, values: [
    'BCC','Big-D','Bonneville','Brinkmann','Cameron','Campbell',
    'EK Bailey','Engage Contracting','Forge Contractors','Fortis',
    'Gramoll','Headwaters','HITT','Hughes','Jacobsen Construction',
    'JN Beach','Judd','Kier','Layton','Lotus','MINT',
    'Nabholz Construction','Poulson','R&O','Revival Development','Rimrock',
    'Robinson','Rod-Lewis','Sheiner Commercial Group','Stallings',
    'Stout','Wadman','Wasatch','Westland','Other'
  ]},
  { title: 'FOLLOW-UP OWNERS',   col: 19, values: [
    'Patrick','Cristhian','Les','Sam','Stephen','VP Preconstruction','President'
  ]},
  { title: 'HVAC ESTIMATORS',    col: 21, values: [
    'Stephen','No Bid'
  ]},
];

function buildSettings(ws) {
  ws.properties.tabColor = { argb: 'FF808080' };
  titleBar(ws, 1, 1, 26, 'ASM PRECONSTRUCTION PIPELINE TRACKER V9 — SETTINGS & DROPDOWN LISTS');

  ws.mergeCells(2, 1, 2, 26);
  const warn = ws.getCell(2, 1);
  warn.value = 'WARNING: Do not delete, reorder, or rename items in this tab. These lists control all dropdown menus in the Bid Log.';
  warn.font  = font({ size: 10, bold: true, color: C.orange });
  warn.fill  = fill('FFFFF2CC');
  warn.alignment = align('left','center');
  ws.getRow(2).height = 16;

  for (const list of SETTINGS_LISTS) {
    const { title, col, values } = list;
    ws.getRow(4).height = 22;
    const hdr = ws.getCell(4, col);
    applyHeader(hdr, title, { size: 9 });

    values.forEach((val, i) => {
      ws.getRow(5 + i).height = 16;
      const cell = ws.getCell(5 + i, col);
      cell.value = val;
      cell.font  = font({ size: 10 });
      cell.fill  = fill(i % 2 === 0 ? C.lgray : C.white);
      cell.alignment = align('left','center');
      cell.border = border();
    });

    ws.getColumn(col).width = 22;
    if (col > 1) ws.getColumn(col - 1).width = 2; // spacer
  }
  ws.getColumn(23).width = 2;
  ws.getColumn(24).width = 2;
}

// ══════════════════════════════════════════════════════════════════════════════
// BID LOG TAB
// ══════════════════════════════════════════════════════════════════════════════
const BID_COLS = [
  { name: 'Bid ID',              width: 9,  numFmt: '@'          }, // col 1  = A
  { name: 'Project Name',        width: 32, numFmt: '@'          }, // col 2  = B
  { name: 'General Contractor',  width: 24, numFmt: '@'          }, // col 3  = C
  { name: 'Location',            width: 20, numFmt: '@'          }, // col 4  = D
  { name: 'Market Segment',      width: 18, numFmt: '@'          }, // col 5  = E
  { name: 'Scope Type',          width: 20, numFmt: '@'          }, // col 6  = F
  { name: 'Bid Stage',           width: 14, numFmt: '@'          }, // col 7  = G
  { name: 'Delivery Method',     width: 16, numFmt: '@'          }, // col 8  = H
  { name: 'Plumb Estimator',     width: 16, numFmt: '@'          }, // col 9  = I
  { name: 'HVAC Estimator',      width: 16, numFmt: '@'          }, // col 10 = J  ← NEW
  { name: 'Bid Received',        width: 14, numFmt: 'm/d/yyyy'   }, // col 11 = K
  { name: 'Bid Due Date',        width: 14, numFmt: 'm/d/yyyy'   }, // col 12 = L
  { name: 'Submitted Date',      width: 14, numFmt: 'm/d/yyyy'   }, // col 13 = M
  { name: 'Plumb Bid ($)',       width: 16, numFmt: '$#,##0'     }, // col 14 = N  ← split
  { name: 'HVAC Bid ($)',        width: 16, numFmt: '$#,##0'     }, // col 15 = O  ← split
  { name: 'Total Bid ($)',       width: 16, numFmt: '$#,##0'     }, // col 16 = P  ← formula
  { name: 'Status',              width: 16, numFmt: '@'          }, // col 17 = Q
  { name: 'Awarded Amount',      width: 16, numFmt: '$#,##0'     }, // col 18 = R
  { name: 'Winning Competitor',  width: 22, numFmt: '@'          }, // col 19 = S
  { name: 'Competitor Amount',   width: 16, numFmt: '$#,##0'     }, // col 20 = T
  { name: 'Lost Reason',         width: 18, numFmt: '@'          }, // col 21 = U
  { name: 'Expected Award Date', width: 18, numFmt: 'm/d/yyyy'   }, // col 22 = V
  { name: 'Probability %',       width: 14, numFmt: '0%'         }, // col 23 = W
  { name: 'Next Follow-Up Date', width: 17, numFmt: 'm/d/yyyy'   }, // col 24 = X
  { name: 'Follow-Up Owner',     width: 18, numFmt: '@'          }, // col 25 = Y
  { name: 'Last Contact Date',   width: 16, numFmt: 'm/d/yyyy'   }, // col 26 = Z
  { name: 'Notes',               width: 38, numFmt: '@'          }, // col 27 = AA
  { name: 'Go/No-Go Score',      width: 14, numFmt: '0'          }, // col 28 = AB
  { name: 'Go/No-Go Rec',        width: 14, numFmt: '@'          }, // col 29 = AC
  { name: 'Days Pending',        width: 14, numFmt: '0'          }, // col 30 = AD
  { name: 'Follow-Up Status',    width: 16, numFmt: '@'          }, // col 31 = AE
  { name: 'Data Quality Flag',   width: 44, numFmt: '@'          }, // col 32 = AF
];

const COL = {};
BID_COLS.forEach((c, i) => { COL[c.name] = i + 1; });

// Real 2026 bids migrated from ASM_Bid_Tracker_v8.xlsx
// Fields: id, proj, gc, loc, mkt, scope, stage, deliv,
//         plumbEst, hvacEst, recv, due, subm,
//         plumbBid, hvacBid, (Total is formula col P)
//         status, awardAmt, winComp, compAmt, lostRsn,
//         expAward, prob, nextFu, fuOwner, lastCt, notes
const REAL_BIDS = [
  // ── Patrick / Stephen bids ────────────────────────────────────────────────
  {id:'ASM-2026-001',proj:'SLCC LAB Remodel',gc:'Judd',loc:'',mkt:'Professional Office',scope:'Full Mechanical',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Patrick',hvacEst:'Stephen',recv:new Date(2026,0,6),due:null,subm:null,plumbBid:159751,hvacBid:59515,status:'Bid Submitted',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.50,nextFu:null,fuOwner:null,lastCt:null,notes:'GCs: Judd, Stout, Rod-Lewis, Wasatch-West'},
  {id:'ASM-2026-002',proj:'Granary Lofts',gc:'Kier',loc:'',mkt:'Multi-Family',scope:'Full Mechanical',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Patrick',hvacEst:'Stephen',recv:new Date(2026,0,9),due:null,subm:null,plumbBid:1271316,hvacBid:1278337,status:'Bid Submitted',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.50,nextFu:null,fuOwner:null,lastCt:null,notes:'GC: Kier (owner-managed)'},
  {id:'ASM-2026-003',proj:'Everhome Suites',gc:'Big-D',loc:'',mkt:'Hotel / Resort',scope:'Full Mechanical',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Patrick',hvacEst:'Stephen',recv:new Date(2026,0,13),due:null,subm:null,plumbBid:1469827,hvacBid:809252,status:'Lost',awardAmt:null,winComp:null,compAmt:1972079,lostRsn:null,expAward:null,prob:0,nextFu:null,fuOwner:null,lastCt:new Date(2026,4,5),notes:'GCs: Big-D, Zwick, R&O'},
  {id:'ASM-2026-004',proj:'Bluffdale City Admin',gc:'Rod-Lewis',loc:'Bluffdale, UT',mkt:'Municipal',scope:'Full Mechanical',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Patrick',hvacEst:'Stephen',recv:new Date(2026,0,15),due:null,subm:null,plumbBid:223680,hvacBid:92569,status:'Bid Submitted',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.50,nextFu:null,fuOwner:null,lastCt:null,notes:'GCs: Rod-Lewis, Judd, Mahas, Bailey, Silverleaf'},
  {id:'ASM-2026-005',proj:'Crown Apt Remodel',gc:'Lotus',loc:'',mkt:'Multi-Family',scope:'Full Mechanical',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Patrick',hvacEst:'Stephen',recv:new Date(2026,0,21),due:null,subm:null,plumbBid:390920,hvacBid:340868,status:'Bid Submitted',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.50,nextFu:null,fuOwner:null,lastCt:null,notes:''},
  {id:'ASM-2026-006',proj:'The Collective @ Morning Vista',gc:'Robinson',loc:'',mkt:'Multi-Family',scope:'Full Mechanical',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Patrick',hvacEst:'Stephen',recv:new Date(2026,0,21),due:null,subm:null,plumbBid:1330294,hvacBid:880986,status:'Bid Submitted',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.50,nextFu:null,fuOwner:null,lastCt:null,notes:'GCs: Robinson, R&O'},
  {id:'ASM-2026-007',proj:'Jordan Valley Lots 8, 9, 10',gc:'BCC',loc:'West Jordan, UT',mkt:'Multi-Family',scope:'Full Mechanical',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Patrick',hvacEst:'Stephen',recv:new Date(2026,0,22),due:null,subm:null,plumbBid:6376876,hvacBid:6874008,status:'Bid Submitted',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.50,nextFu:null,fuOwner:null,lastCt:null,notes:'GC: BCC'},
  {id:'ASM-2026-008',proj:'Pacific Yard Apt',gc:'Kier',loc:'',mkt:'Multi-Family',scope:'Full Mechanical',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Patrick',hvacEst:'Stephen',recv:new Date(2026,1,5),due:null,subm:null,plumbBid:5398931,hvacBid:4298204,status:'Bid Submitted',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.50,nextFu:null,fuOwner:null,lastCt:null,notes:'GCs: Kier, Headwaters, RVC'},
  {id:'ASM-2026-009',proj:'Antelope Station',gc:'Revival Development',loc:'',mkt:'Municipal',scope:'Full Mechanical',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Patrick',hvacEst:'Stephen',recv:new Date(2026,0,28),due:null,subm:null,plumbBid:43556,hvacBid:80889,status:'Bid Submitted',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.50,nextFu:null,fuOwner:null,lastCt:null,notes:''},
  {id:'ASM-2026-010',proj:'Red Sky Apt',gc:'BCC',loc:'',mkt:'Multi-Family',scope:'Full Mechanical',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Patrick',hvacEst:'Stephen',recv:new Date(2026,1,3),due:null,subm:null,plumbBid:3166896,hvacBid:3511430,status:'Bid Submitted',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.50,nextFu:null,fuOwner:null,lastCt:null,notes:'GC: BCC (owner-managed)'},
  {id:'ASM-2026-011',proj:'HH Parking Structure',gc:'Wadman',loc:'',mkt:'Professional Office',scope:'Plumbing Only',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Patrick',hvacEst:'No Bid',recv:new Date(2026,0,29),due:null,subm:null,plumbBid:167139,hvacBid:null,status:'Bid Submitted',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.50,nextFu:null,fuOwner:null,lastCt:null,notes:'GCs: Wadman, Forge'},
  {id:'ASM-2026-012',proj:'The Royce',gc:'Wadman',loc:'',mkt:'Multi-Family',scope:'Full Mechanical',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Patrick',hvacEst:'Stephen',recv:new Date(2026,0,29),due:null,subm:null,plumbBid:4134360,hvacBid:4534706,status:'Bid Submitted',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.50,nextFu:null,fuOwner:null,lastCt:null,notes:'GC: Wadman'},
  {id:'ASM-2026-013',proj:'Crossroads Urban Center',gc:'Cameron',loc:'',mkt:'Professional Office',scope:'Full Mechanical',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Patrick',hvacEst:'Stephen',recv:new Date(2026,1,5),due:null,subm:null,plumbBid:205754,hvacBid:130206,status:'Bid Submitted',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.50,nextFu:null,fuOwner:null,lastCt:null,notes:'GC: Cameron'},
  {id:'ASM-2026-014',proj:'Mojave Crossing',gc:'EK Bailey',loc:'',mkt:'Multi-Family',scope:'Full Mechanical',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Patrick',hvacEst:'Stephen',recv:new Date(2026,0,27),due:null,subm:null,plumbBid:1332893,hvacBid:778874,status:'Bid Submitted',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.50,nextFu:null,fuOwner:null,lastCt:null,notes:'GC: EK Bailey'},
  {id:'ASM-2026-015',proj:'Goodyear Flex Bldg',gc:'Campbell',loc:'Goodyear, AZ',mkt:'Professional Office',scope:'Full Mechanical',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Patrick',hvacEst:'Stephen',recv:new Date(2026,1,10),due:null,subm:null,plumbBid:114514,hvacBid:72736,status:'Bid Submitted',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.50,nextFu:null,fuOwner:null,lastCt:null,notes:'GC: Campbell'},
  {id:'ASM-2026-016',proj:'Salt Lake Excavating',gc:'Campbell',loc:'Salt Lake City, UT',mkt:'Professional Office',scope:'Full Mechanical',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Patrick',hvacEst:'Stephen',recv:new Date(2026,1,18),due:null,subm:null,plumbBid:339849,hvacBid:276430,status:'Bid Submitted',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.50,nextFu:null,fuOwner:null,lastCt:null,notes:'GC: Campbell'},
  {id:'ASM-2026-017',proj:'The Hive on 11th',gc:'Bonneville',loc:'Salt Lake City, UT',mkt:'Multi-Family',scope:'Full Mechanical',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Patrick',hvacEst:'Stephen',recv:new Date(2026,1,11),due:null,subm:null,plumbBid:2809180,hvacBid:3496373,status:'Bid Submitted',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.50,nextFu:null,fuOwner:null,lastCt:null,notes:'GC: Bonneville'},
  {id:'ASM-2026-018',proj:'Radnet Corp. Office',gc:'Big-D',loc:'',mkt:'Healthcare',scope:'Full Mechanical',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Patrick',hvacEst:'Stephen',recv:new Date(2026,1,23),due:null,subm:null,plumbBid:3417972,hvacBid:2417425,status:'Bid Submitted',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.50,nextFu:null,fuOwner:null,lastCt:null,notes:'GCs: Big-D, Westland, Bonneville, R&O, Eckman'},
  {id:'ASM-2026-019',proj:'Kari Sue Hamilton Remodel',gc:'Judd',loc:'',mkt:'Education',scope:'Full Mechanical',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Patrick',hvacEst:'Stephen',recv:new Date(2026,1,4),due:null,subm:null,plumbBid:131701,hvacBid:153676,status:'Bid Submitted',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.50,nextFu:null,fuOwner:null,lastCt:null,notes:'GCs: Judd, Mahas'},
  {id:'ASM-2026-020',proj:'St. George AP Add',gc:'Big-D',loc:'St. George, UT',mkt:'Professional Office',scope:'Full Mechanical',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Patrick',hvacEst:'',recv:new Date(2026,1,4),due:null,subm:null,plumbBid:885121,hvacBid:4277586,status:'Bid Submitted',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.50,nextFu:null,fuOwner:null,lastCt:null,notes:'GCs: Big-D, Westland, Bonneville, R&O'},
  {id:'ASM-2026-021',proj:'Stryker Kitchen & Café Remodel',gc:'Engage Contracting',loc:'',mkt:'Professional Office',scope:'Full Mechanical',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Patrick',hvacEst:'Stephen',recv:new Date(2026,2,17),due:null,subm:null,plumbBid:355155,hvacBid:232824,status:'Bid Submitted',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.50,nextFu:null,fuOwner:null,lastCt:null,notes:'GC: Engage Contracting'},
  {id:'ASM-2026-022',proj:'HCA West Jordan FSER',gc:'Nabholz Construction',loc:'West Jordan, UT',mkt:'Healthcare',scope:'Full Mechanical',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Patrick',hvacEst:'',recv:new Date(2026,2,4),due:null,subm:null,plumbBid:2326498,hvacBid:1192678,status:'Bid Submitted',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.50,nextFu:null,fuOwner:null,lastCt:null,notes:'Free Standing Emergency Room'},
  {id:'ASM-2026-023',proj:'UTA Mt. Ogden Admin',gc:'Poulson',loc:'Ogden, UT',mkt:'Professional Office',scope:'Full Mechanical',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Patrick',hvacEst:'Stephen',recv:new Date(2026,1,19),due:null,subm:null,plumbBid:602661,hvacBid:797531,status:'Bid Submitted',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.50,nextFu:null,fuOwner:null,lastCt:null,notes:'GCs: Poulson, Westland'},
  {id:'ASM-2026-024',proj:'Delta Cargo',gc:'Big-D',loc:'',mkt:'Professional Office',scope:'Full Mechanical',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Patrick',hvacEst:'Stephen',recv:new Date(2026,1,20),due:null,subm:null,plumbBid:142132,hvacBid:329764,status:'Bid Submitted',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.50,nextFu:null,fuOwner:null,lastCt:null,notes:'GC: Big-D'},
  {id:'ASM-2026-025',proj:'American Cruise Lines',gc:'Engage Contracting',loc:'',mkt:'Professional Office',scope:'Full Mechanical',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Patrick',hvacEst:'Stephen',recv:new Date(2026,1,26),due:null,subm:null,plumbBid:68011,hvacBid:36426,status:'Bid Submitted',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.50,nextFu:null,fuOwner:null,lastCt:null,notes:'GCs: Engage, Stout'},
  {id:'ASM-2026-026',proj:'IH Southridge',gc:'Layton',loc:'',mkt:'Professional Office',scope:'Full Mechanical',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Patrick',hvacEst:'Stephen',recv:new Date(2026,1,26),due:null,subm:null,plumbBid:222700,hvacBid:157906,status:'Bid Submitted',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.50,nextFu:null,fuOwner:null,lastCt:null,notes:'GC: Layton'},
  {id:'ASM-2026-027',proj:'Canyons Innovation Center',gc:'Hughes',loc:'',mkt:'Professional Office',scope:'Plumbing Only',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Patrick',hvacEst:'',recv:new Date(2026,2,19),due:null,subm:null,plumbBid:4367489,hvacBid:null,status:'Bid Submitted',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.50,nextFu:null,fuOwner:null,lastCt:null,notes:'GC: Hughes'},
  {id:'ASM-2026-028',proj:'UofU Bus Garage',gc:'Fortis',loc:'Salt Lake City, UT',mkt:'Education',scope:'Plumbing Only',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Patrick',hvacEst:'Stephen',recv:new Date(2026,2,19),due:null,subm:null,plumbBid:34977,hvacBid:17303,status:'Bid Submitted',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.50,nextFu:null,fuOwner:null,lastCt:null,notes:'GC: Fortis'},
  {id:'ASM-2026-029',proj:'AF Police Station Remodel',gc:'Stout',loc:'American Fork, UT',mkt:'Municipal',scope:'Full Mechanical',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Patrick',hvacEst:'Stephen',recv:new Date(2026,2,20),due:null,subm:null,plumbBid:90847,hvacBid:125683,status:'Bid Submitted',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.50,nextFu:null,fuOwner:null,lastCt:null,notes:'GCs: Stout, Miller, Warner, Wasatch-West, Big-D, KDK, Judd'},
  {id:'ASM-2026-030',proj:'Ogden Weber Tech',gc:'Gramoll',loc:'Ogden, UT',mkt:'Education',scope:'Plumbing Only',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Patrick',hvacEst:'Stephen',recv:new Date(2026,2,23),due:null,subm:null,plumbBid:12110932,hvacBid:3451361,status:'Bid Submitted',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.50,nextFu:null,fuOwner:null,lastCt:null,notes:'GCs: Gramoll, R&O'},
  {id:'ASM-2026-031',proj:'BYU Tanner',gc:'Stallings',loc:'Provo, UT',mkt:'Education',scope:'Full Mechanical',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Patrick',hvacEst:'Stephen',recv:new Date(2026,2,27),due:null,subm:null,plumbBid:84641,hvacBid:511064,status:'Bid Submitted',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.50,nextFu:null,fuOwner:null,lastCt:null,notes:'GC: Stallings'},
  {id:'ASM-2026-032',proj:'Ford Training Center',gc:'Judd',loc:'',mkt:'Professional Office',scope:'Full Mechanical',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Patrick',hvacEst:'Stephen',recv:new Date(2026,2,27),due:null,subm:null,plumbBid:115057,hvacBid:166286,status:'Bid Submitted',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.50,nextFu:null,fuOwner:null,lastCt:null,notes:'GC: Judd'},
  {id:'ASM-2026-033',proj:'Coleman Orthodontics',gc:'EK Bailey',loc:'',mkt:'Professional Office',scope:'Full Mechanical',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Patrick',hvacEst:'Stephen',recv:new Date(2026,3,1),due:null,subm:null,plumbBid:135032,hvacBid:70834,status:'Bid Submitted',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.50,nextFu:null,fuOwner:null,lastCt:null,notes:'GC: EK Bailey'},
  // ── Les bids ──────────────────────────────────────────────────────────────
  {id:'ASM-2026-034',proj:'UofU PET',gc:'Fortis',loc:'Salt Lake City, UT',mkt:'Education',scope:'Full Mechanical',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Les',hvacEst:'',recv:new Date(2026,3,2),due:null,subm:null,plumbBid:73232,hvacBid:17303,status:'Bid Submitted',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.50,nextFu:null,fuOwner:null,lastCt:null,notes:'GC: Fortis'},
  {id:'ASM-2026-035',proj:'Project Bedrock',gc:'Brinkmann',loc:'',mkt:'Mission Critical',scope:'Full Mechanical',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Les',hvacEst:'Stephen',recv:new Date(2026,3,10),due:null,subm:null,plumbBid:8813255,hvacBid:4172214,status:'Bid Submitted',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.50,nextFu:null,fuOwner:null,lastCt:null,notes:'GC: Brinkmann Construction'},
  {id:'ASM-2026-036',proj:'Northwest Yardworks',gc:'Forge Contractors',loc:'',mkt:'Professional Office',scope:'Plumbing Only',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Les',hvacEst:'',recv:new Date(2026,3,6),due:null,subm:null,plumbBid:645427,hvacBid:null,status:'Lost',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0,nextFu:null,fuOwner:null,lastCt:null,notes:'GC: Forge Contractors'},
  {id:'ASM-2026-037',proj:'Skims',gc:'Sheiner Commercial Group',loc:'',mkt:'Professional Office',scope:'Full Mechanical',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Les',hvacEst:'Stephen',recv:new Date(2026,3,9),due:null,subm:null,plumbBid:311643,hvacBid:null,status:'Lost',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0,nextFu:null,fuOwner:null,lastCt:null,notes:'GC: Sheiner Commercial Group'},
  {id:'ASM-2026-038',proj:'U of U Fan Coil Replacement',gc:'Fortis',loc:'Salt Lake City, UT',mkt:'Healthcare',scope:'HVAC Only',stage:'Budget',deliv:'Negotiated',plumbEst:'',hvacEst:'Les',recv:new Date(2026,3,20),due:null,subm:null,plumbBid:null,hvacBid:28000,status:'Lost',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0,nextFu:null,fuOwner:null,lastCt:null,notes:'GC: Fortis'},
  {id:'ASM-2026-039',proj:'UVH Sorenson 9th Floor Remodel',gc:'Big-D',loc:'Salt Lake City, UT',mkt:'Healthcare',scope:'Full Mechanical',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Les',hvacEst:'Stephen',recv:new Date(2026,4,4),due:null,subm:null,plumbBid:1358760,hvacBid:960690,status:'Bid Submitted',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.50,nextFu:null,fuOwner:null,lastCt:null,notes:'GC: Big-D'},
  {id:'ASM-2026-040',proj:'Sharon Elementary Remodel',gc:'Westland',loc:'',mkt:'Education',scope:'Full Mechanical',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Les',hvacEst:'Stephen',recv:new Date(2026,3,30),due:null,subm:null,plumbBid:313879,hvacBid:543475,status:'Bid Submitted',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.50,nextFu:null,fuOwner:null,lastCt:null,notes:'GCs: Westland & Judd'},
  // ── Sam bids ──────────────────────────────────────────────────────────────
  {id:'ASM-2026-041',proj:'The Point Parcel H3A',gc:'R&O',loc:'',mkt:'Multi-Family',scope:'Budget Only',stage:'Budget',deliv:'Negotiated',plumbEst:'Sam',hvacEst:'',recv:new Date(2026,2,13),due:null,subm:null,plumbBid:11897000,hvacBid:null,status:'Bid Submitted',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.50,nextFu:null,fuOwner:null,lastCt:null,notes:'GC: R&O — budget pricing'},
  {id:'ASM-2026-042',proj:'Xenter',gc:'Jacobsen Construction',loc:'',mkt:'Industrial / Manufacturing',scope:'Full Mechanical',stage:'Budget',deliv:'Negotiated',plumbEst:'Sam',hvacEst:'',recv:new Date(2026,3,23),due:null,subm:null,plumbBid:1950000,hvacBid:null,status:'Bid Submitted',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.50,nextFu:null,fuOwner:null,lastCt:null,notes:'GC: Jacobsen Construction — budget pricing'},
  // ── Patrick + ARA Southwest (In Progress) ─────────────────────────────────
  {id:'ASM-2026-043',proj:'ARA Southwest (St. George)',gc:'Campbell',loc:'St. George, UT',mkt:'Professional Office',scope:'Full Mechanical',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Patrick',hvacEst:'No Bid',recv:new Date(2026,3,20),due:null,subm:null,plumbBid:885352,hvacBid:null,status:'In Progress',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.25,nextFu:null,fuOwner:null,lastCt:null,notes:'GC: Campbell'},
  {id:'ASM-2026-044',proj:'Moab Hotel',gc:'Bonneville',loc:'Moab, UT',mkt:'Hotel / Resort',scope:'Full Mechanical',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Patrick',hvacEst:'Stephen',recv:new Date(2026,3,6),due:null,subm:null,plumbBid:null,hvacBid:null,status:'In Progress',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.25,nextFu:null,fuOwner:null,lastCt:null,notes:'GC: Bonneville'},
  {id:'ASM-2026-045',proj:'Trove Office Bldg',gc:'Stout',loc:'',mkt:'Professional Office',scope:'Full Mechanical',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Patrick',hvacEst:'Stephen',recv:new Date(2026,3,9),due:null,subm:null,plumbBid:null,hvacBid:null,status:'In Progress',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.25,nextFu:null,fuOwner:null,lastCt:null,notes:'GC: Stout'},
  // ── Patrick — Lost bids ───────────────────────────────────────────────────
  {id:'ASM-2026-046',proj:'IH St. George AHC',gc:'Big-D',loc:'St. George, UT',mkt:'Professional Office',scope:'Full Mechanical',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Patrick',hvacEst:'Stephen',recv:new Date(2026,3,7),due:null,subm:null,plumbBid:599945,hvacBid:484763,status:'Lost',awardAmt:null,winComp:'Local Plumber',compAmt:922708,lostRsn:'Relationship',expAward:null,prob:0,nextFu:null,fuOwner:null,lastCt:new Date(2026,4,5),notes:'GC: Big-D'},
  {id:'ASM-2026-047',proj:'Ford Land TI',gc:'HITT',loc:'',mkt:'Professional Office',scope:'Full Mechanical',stage:'Design-Build',deliv:'Design-Build',plumbEst:'Patrick',hvacEst:'Stephen',recv:new Date(2026,2,15),due:null,subm:null,plumbBid:247641,hvacBid:341169,status:'Lost',awardAmt:null,winComp:'Harris Mechanical',compAmt:588810,lostRsn:'Price',expAward:null,prob:0,nextFu:null,fuOwner:null,lastCt:new Date(2026,3,14),notes:'GC: HITT'},
  // ── Cristhian bids ────────────────────────────────────────────────────────
  {id:'ASM-2026-048',proj:'Dirks Field / Book Cliffs',gc:'',loc:'',mkt:'Multi-Family',scope:'Full Mechanical',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Cristhian',hvacEst:'Stephen',recv:new Date(2025,9,27),due:null,subm:null,plumbBid:927548,hvacBid:1330682,status:'Awarded',awardAmt:927548,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:1.0,nextFu:null,fuOwner:null,lastCt:new Date(2026,1,9),notes:'Awarded Feb 2026'},
  {id:'ASM-2026-049',proj:'Amelia Apartments',gc:'',loc:'',mkt:'Multi-Family',scope:'Full Mechanical',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Cristhian',hvacEst:'Stephen',recv:new Date(2025,11,1),due:null,subm:null,plumbBid:2240430,hvacBid:3245170,status:'Bid Submitted',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.50,nextFu:null,fuOwner:null,lastCt:null,notes:''},
  {id:'ASM-2026-050',proj:'Brooklyn Yards',gc:'',loc:'',mkt:'Multi-Family',scope:'Plumbing Only',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Cristhian',hvacEst:'',recv:new Date(2025,11,17),due:null,subm:null,plumbBid:2526170,hvacBid:null,status:'Awarded',awardAmt:2526170,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:1.0,nextFu:null,fuOwner:null,lastCt:new Date(2026,1,26),notes:'Awarded Feb 2026'},
  {id:'ASM-2026-051',proj:'Wasatch Rock',gc:'',loc:'',mkt:'Multi-Family',scope:'Plumbing Only',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Cristhian',hvacEst:'',recv:new Date(2025,11,19),due:null,subm:null,plumbBid:5246900,hvacBid:null,status:'Bid Submitted',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.50,nextFu:null,fuOwner:null,lastCt:null,notes:''},
  {id:'ASM-2026-052',proj:'Malakyte Apartments (Adams Ave)',gc:'',loc:'',mkt:'Multi-Family',scope:'Plumbing Only',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Cristhian',hvacEst:'',recv:new Date(2025,4,23),due:null,subm:null,plumbBid:1786648,hvacBid:null,status:'Awarded',awardAmt:1786648,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:1.0,nextFu:null,fuOwner:null,lastCt:new Date(2026,1,9),notes:'Awarded Feb 2026'},
  {id:'ASM-2026-053',proj:'Ascend at Riverview',gc:'',loc:'',mkt:'Multi-Family',scope:'Plumbing Only',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Cristhian',hvacEst:'',recv:new Date(2026,1,9),due:null,subm:null,plumbBid:3632233,hvacBid:null,status:'Bid Submitted',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.50,nextFu:null,fuOwner:null,lastCt:null,notes:''},
  {id:'ASM-2026-054',proj:'Merced Creek Apartments',gc:'Wasatch',loc:'',mkt:'Multi-Family',scope:'Full Mechanical',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Cristhian',hvacEst:'Stephen',recv:new Date(2026,1,24),due:null,subm:null,plumbBid:4381716,hvacBid:5974387,status:'Bid Submitted',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.50,nextFu:null,fuOwner:null,lastCt:null,notes:'GC: Wasatch'},
  {id:'ASM-2026-055',proj:'Euclid Apartments',gc:'Stout',loc:'',mkt:'Multi-Family',scope:'Plumbing Only',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Cristhian',hvacEst:'',recv:new Date(2026,2,4),due:null,subm:null,plumbBid:499239,hvacBid:null,status:'Bid Submitted',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.50,nextFu:null,fuOwner:null,lastCt:null,notes:'GC: Stout'},
  {id:'ASM-2026-056',proj:'aQui 2194',gc:'MINT',loc:'',mkt:'Multi-Family',scope:'Full Mechanical',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Cristhian',hvacEst:'Stephen',recv:new Date(2026,2,17),due:null,subm:null,plumbBid:748908,hvacBid:597262,status:'Bid Submitted',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.50,nextFu:null,fuOwner:null,lastCt:null,notes:'GC: MINT'},
  {id:'ASM-2026-057',proj:'Emeril Apartments',gc:'Headwaters',loc:'',mkt:'Multi-Family',scope:'Full Mechanical',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Cristhian',hvacEst:'Stephen',recv:new Date(2026,2,24),due:null,subm:null,plumbBid:3397199,hvacBid:2142548,status:'Bid Submitted',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.50,nextFu:null,fuOwner:null,lastCt:null,notes:'GC: Headwaters'},
  {id:'ASM-2026-058',proj:'Northwest Pipeline',gc:'Bonneville',loc:'',mkt:'Multi-Family',scope:'Full Mechanical',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Cristhian',hvacEst:'Stephen',recv:new Date(2026,2,25),due:null,subm:null,plumbBid:2578560,hvacBid:4387570,status:'Bid Submitted',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.50,nextFu:null,fuOwner:null,lastCt:null,notes:'GC: Bonneville'},
  {id:'ASM-2026-059',proj:'Murray Block One',gc:'Rimrock',loc:'Murray, UT',mkt:'Multi-Family',scope:'Full Mechanical',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Cristhian',hvacEst:'Stephen',recv:new Date(2026,2,31),due:null,subm:null,plumbBid:3020768,hvacBid:2618756,status:'Bid Submitted',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.50,nextFu:null,fuOwner:null,lastCt:new Date(2026,4,5),notes:'GCs: Rimrock / RVC — result TBD'},
  {id:'ASM-2026-060',proj:'The Grove (Lupine Building)',gc:'Bonneville',loc:'',mkt:'Multi-Family',scope:'Full Mechanical',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Cristhian',hvacEst:'Stephen',recv:new Date(2026,3,8),due:null,subm:null,plumbBid:2912975,hvacBid:3264667,status:'Bid Submitted',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.50,nextFu:null,fuOwner:null,lastCt:null,notes:'GCs: Bonneville / Big-D / Shaw'},
  {id:'ASM-2026-061',proj:'Utopia Apartments',gc:'MINT',loc:'',mkt:'Multi-Family',scope:'Full Mechanical',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Cristhian',hvacEst:'Stephen',recv:new Date(2026,3,20),due:null,subm:null,plumbBid:5376708,hvacBid:3442844,status:'Bid Submitted',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.50,nextFu:null,fuOwner:null,lastCt:null,notes:'GC: MINT'},
  {id:'ASM-2026-062',proj:'300 West Apartments',gc:'JN Beach',loc:'Salt Lake City, UT',mkt:'Multi-Family',scope:'Full Mechanical',stage:'Budget',deliv:'Negotiated',plumbEst:'Cristhian',hvacEst:'Stephen',recv:new Date(2026,3,20),due:null,subm:null,plumbBid:2918916,hvacBid:2715700,status:'Bid Submitted',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.50,nextFu:null,fuOwner:null,lastCt:null,notes:'GC: JN Beach — budget pricing'},
  {id:'ASM-2026-063',proj:'Herriman Commons',gc:'Kier',loc:'Herriman, UT',mkt:'Multi-Family',scope:'Plumbing Only',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Cristhian',hvacEst:'',recv:new Date(2026,3,27),due:null,subm:null,plumbBid:3571515,hvacBid:null,status:'In Progress',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.25,nextFu:null,fuOwner:null,lastCt:null,notes:'GC: Kier'},
  {id:'ASM-2026-064',proj:'Lofts North Salt Lake',gc:'Headwaters',loc:'North Salt Lake, UT',mkt:'Multi-Family',scope:'Full Mechanical',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Cristhian',hvacEst:'Stephen',recv:new Date(2025,4,1),due:null,subm:null,plumbBid:848181,hvacBid:706093,status:'In Progress',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.25,nextFu:null,fuOwner:null,lastCt:null,notes:'GC: Headwaters'},
  {id:'ASM-2026-065',proj:'Lotus Forge Apartments',gc:'Lotus',loc:'',mkt:'Multi-Family',scope:'Full Mechanical',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Cristhian',hvacEst:'Stephen',recv:new Date(2026,4,25),due:null,subm:null,plumbBid:null,hvacBid:null,status:'In Progress',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.25,nextFu:null,fuOwner:null,lastCt:null,notes:'GC: Lotus Built Group'},
  {id:'ASM-2026-066',proj:'Provo Dual Brand Hotel',gc:'Bonneville',loc:'Provo, UT',mkt:'Hotel / Resort',scope:'Full Mechanical',stage:'Design-Build',deliv:'Design-Build',plumbEst:'Cristhian',hvacEst:'Stephen',recv:new Date(2026,4,5),due:null,subm:null,plumbBid:32000,hvacBid:32000,status:'Bid Submitted',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0.50,nextFu:null,fuOwner:null,lastCt:null,notes:'GC: Bonneville'},
  // ── Cristhian — Lost bids ─────────────────────────────────────────────────
  {id:'ASM-2026-067',proj:'Jefferson Apartments',gc:'',loc:'',mkt:'Multi-Family',scope:'Plumbing Only',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Cristhian',hvacEst:'Stephen',recv:new Date(2026,0,29),due:null,subm:null,plumbBid:2484078,hvacBid:2442654,status:'Lost',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0,nextFu:null,fuOwner:null,lastCt:new Date(2026,4,24),notes:''},
  {id:'ASM-2026-068',proj:'Granary Lofts Apartments',gc:'',loc:'',mkt:'Multi-Family',scope:'Plumbing Only',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Cristhian',hvacEst:'',recv:new Date(2026,0,12),due:null,subm:null,plumbBid:1118554,hvacBid:null,status:'Lost',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0,nextFu:null,fuOwner:null,lastCt:null,notes:''},
  {id:'ASM-2026-069',proj:'Ameyalli Wellbeing',gc:'',loc:'',mkt:'Multi-Family',scope:'Plumbing Only',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Cristhian',hvacEst:'',recv:new Date(2026,0,14),due:null,subm:null,plumbBid:1714760,hvacBid:null,status:'Lost',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0,nextFu:null,fuOwner:null,lastCt:null,notes:''},
  {id:'ASM-2026-070',proj:'Jordan Valley',gc:'',loc:'West Jordan, UT',mkt:'Multi-Family',scope:'Plumbing Only',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Cristhian',hvacEst:'',recv:new Date(2026,0,23),due:null,subm:null,plumbBid:6489325,hvacBid:null,status:'Lost',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0,nextFu:null,fuOwner:null,lastCt:null,notes:''},
  {id:'ASM-2026-071',proj:'Holladay Parking',gc:'',loc:'Holladay, UT',mkt:'Multi-Family',scope:'Plumbing Only',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Cristhian',hvacEst:'',recv:new Date(2026,0,29),due:null,subm:null,plumbBid:158514,hvacBid:null,status:'Lost',awardAmt:null,winComp:null,compAmt:null,lostRsn:null,expAward:null,prob:0,nextFu:null,fuOwner:null,lastCt:null,notes:''},
  {id:'ASM-2026-072',proj:'SLC Student Housing',gc:'Headwaters',loc:'Salt Lake City, UT',mkt:'Multi-Family',scope:'Full Mechanical',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Cristhian',hvacEst:'Stephen',recv:new Date(2026,1,19),due:null,subm:null,plumbBid:6729530,hvacBid:4268902,status:'Lost',awardAmt:null,winComp:'Shamrock Plumbing',compAmt:6500000,lostRsn:'Price',expAward:null,prob:0,nextFu:null,fuOwner:null,lastCt:new Date(2026,3,9),notes:'GC: Headwaters'},
  {id:'ASM-2026-073',proj:'Lakeside Apartments',gc:'Rimrock',loc:'',mkt:'Multi-Family',scope:'Full Mechanical',stage:'Hard Bid',deliv:'Hard Bid',plumbEst:'Cristhian',hvacEst:'Stephen',recv:new Date(2026,3,2),due:null,subm:null,plumbBid:5358517,hvacBid:3970420,status:'Lost',awardAmt:null,winComp:'Harris Mechanical & Shamrock',compAmt:7858517,lostRsn:'Price',expAward:null,prob:0,nextFu:null,fuOwner:null,lastCt:new Date(2026,4,5),notes:'GC: Rimrock'},
];

function buildBidLog(ws) {
  ws.properties.tabColor = { argb: 'FF2E75B6' };
  ws.views = [{ state: 'frozen', xSplit: 1, ySplit: 1 }];
  ws.getRow(1).height = 40;

  // Set column widths and headers
  BID_COLS.forEach((col, i) => {
    const c = i + 1;
    ws.getColumn(c).width = col.width;
    const cell = ws.getCell(1, c);
    applyHeader(cell, col.name, { size: 10, wrap: true });
  });

  // Data rows
  REAL_BIDS.forEach((bid, i) => {
    const r = i + 2;
    ws.getRow(r).height = 18;
    const bg = r % 2 === 0 ? C.lgray : C.white;

    // cols 1-27 (A-AA): direct data values; col 16 (P = Total Bid) is a formula added below
    const vals = [
      bid.id, bid.proj, bid.gc, bid.loc, bid.mkt, bid.scope, bid.stage,
      bid.deliv, bid.plumbEst, bid.hvacEst, bid.recv, bid.due, bid.subm,
      bid.plumbBid, bid.hvacBid, null,           // col 14 N, 15 O, 16 P (formula)
      bid.status, bid.awardAmt, bid.winComp, bid.compAmt, bid.lostRsn,
      bid.expAward, bid.prob, bid.nextFu, bid.fuOwner, bid.lastCt, bid.notes,
      null, null,                                 // cols 28-29 AB-AC  (GNG formulas)
    ];

    vals.forEach((val, ci) => {
      const colDef = BID_COLS[ci];
      const cell = ws.getCell(r, ci + 1);
      if (val !== null && val !== undefined) {
        cell.value = val;
      }
      cell.font  = font({ size: 10 });
      cell.fill  = fill(bg);
      cell.alignment = align('left', 'center', ci === 26); // Notes = col 27 (index 26)
      cell.border = border();
      if (colDef && colDef.numFmt && colDef.numFmt !== '@') {
        cell.numFmt = colDef.numFmt;
      }
    });

    // Total Bid (col 16 = P) — formula: plumb + HVAC
    const totalBid = ws.getCell(r, 16);
    totalBid.value = { formula: `IFERROR(IF(AND(N${r}="",O${r}=""),"",IFERROR(N${r},0)+IFERROR(O${r},0)),"")` };
    totalBid.font = font({ size: 10 }); totalBid.fill = fill(bg);
    totalBid.alignment = align('right', 'center'); totalBid.border = border();
    totalBid.numFmt = '$#,##0';

    // Go/No-Go Score (col 28 = AB) — VLOOKUP from Go-No-Go tab
    const gngScore = ws.getCell(r, 28);
    gngScore.value = { formula: `IFERROR(VLOOKUP(A${r},'Go-No-Go'!$A:$M,12,0),"")` };
    gngScore.font  = font({ size: 10, color: C.dgray });
    gngScore.fill  = fill(bg);
    gngScore.alignment = align('center');
    gngScore.border = border();
    gngScore.numFmt = '0';

    // Go/No-Go Rec (col 29 = AC) — VLOOKUP from Go-No-Go tab
    const gngRec = ws.getCell(r, 29);
    gngRec.value = { formula: `IFERROR(VLOOKUP(A${r},'Go-No-Go'!$A:$N,14,0),"")` };
    gngRec.font  = font({ size: 10, color: C.dgray });
    gngRec.fill  = fill(bg);
    gngRec.alignment = align('center');
    gngRec.border = border();

    // Days Pending (col 30 = AD)
    // Counts days since Submitted Date (M), or Bid Received (K) if not yet submitted
    const daysPend = ws.getCell(r, 30);
    daysPend.value = {
      formula: `IF(OR(Q${r}="Bid Submitted",Q${r}="Budget Submitted",Q${r}="ROM Submitted",Q${r}="In Progress"),IF(M${r}<>"",TODAY()-M${r},IF(K${r}<>"",TODAY()-K${r},"")),"")`
    };
    daysPend.font = font({ size: 10 }); daysPend.fill = fill(bg);
    daysPend.alignment = align('center'); daysPend.border = border();
    daysPend.numFmt = '0';

    // Follow-Up Status (col 31 = AE)
    const fuStat = ws.getCell(r, 31);
    const pend = `OR(Q${r}="Bid Submitted",Q${r}="Budget Submitted",Q${r}="ROM Submitted")`;
    fuStat.value = {
      formula: `IF(AND(${pend},X${r}<>"",X${r}<TODAY()),"OVERDUE",IF(AND(${pend},X${r}<>"",X${r}<=TODAY()+3),"DUE SOON",IF(AND(${pend},X${r}=""),"MISSING","OK")))`
    };
    fuStat.font = font({ size: 10 }); fuStat.fill = fill(bg);
    fuStat.alignment = align('center'); fuStat.border = border();

    // Data Quality Flag (col 32 = AF)
    const dq = ws.getCell(r, 32);
    const dqInner =
        `IF(B${r}="","[Missing Project Name] ","")&`+
        `IF(C${r}="","[Missing GC] ","")&`+
        `IF(I${r}="","[Missing Plumb Estimator] ","")&`+
        `IF(AND(Q${r}<>"In Progress",Q${r}<>"No Bid",Q${r}<>"On Hold",P${r}=""),"[Missing Bid Amount] ","")&`+
        `IF(AND(Q${r}="Awarded",R${r}=""),"[Awarded-No Award Amount] ","")&`+
        `IF(AND(Q${r}="Lost",U${r}=""),"[Lost-No Lost Reason] ","")&`+
        `IF(AND(OR(Q${r}="Bid Submitted",Q${r}="Budget Submitted"),X${r}=""),"[Pending-No Follow-Up Date] ","")&`+
        `IF(AND(OR(Q${r}="Bid Submitted",Q${r}="Budget Submitted"),Y${r}=""),"[Pending-No Follow-Up Owner] ","")&`+
        `IF(AND(L${r}<>"",L${r}<TODAY(),OR(Q${r}="In Progress",Q${r}="")),"[Bid Due Date Passed] ","")`;
    dq.value = { formula: `IF(TRIM(${dqInner})="","OK",TRIM(${dqInner}))` };
    dq.font = font({ size: 9 }); dq.fill = fill(bg);
    dq.alignment = align('left', 'center', true); dq.border = border();
  });

  // Data validation (dropdown lists referencing Settings)
  // General Contractor is free-text (no DV) — multiple GCs per bid are common
  const dvMappings = [
    { col: COL['Market Segment'],     settingsCol: 5,  start: 5, count: 15 },
    { col: COL['Scope Type'],         settingsCol: 7,  start: 5, count: 6  },
    { col: COL['Bid Stage'],          settingsCol: 9,  start: 5, count: 7  },
    { col: COL['Delivery Method'],    settingsCol: 11, start: 5, count: 6  },
    { col: COL['Plumb Estimator'],    settingsCol: 1,  start: 5, count: 5  },
    { col: COL['HVAC Estimator'],     settingsCol: 21, start: 5, count: 2  },
    { col: COL['Status'],             settingsCol: 3,  start: 5, count: 8  },
    { col: COL['Lost Reason'],        settingsCol: 13, start: 5, count: 9  },
    { col: COL['Follow-Up Owner'],    settingsCol: 19, start: 5, count: 7  },
  ];

  dvMappings.forEach(({ col, settingsCol, start, count }) => {
    const colLetter = ws.getColumn(col).letter || colToLetter(col);
    const sColLetter = colToLetter(settingsCol);
    ws.dataValidations.add(`${colLetter}2:${colLetter}2001`, {
      type: 'list',
      allowBlank: true,
      showErrorMessage: true,
      errorStyle: 'stop',
      errorTitle: 'Invalid Entry',
      error: 'Please select from the dropdown list.',
      formulae: [`Settings!$${sColLetter}$${start}:$${sColLetter}$${start + count - 1}`],
    });
  });

  // Conditional formatting
  const lastRow = REAL_BIDS.length + 1; // 73 data rows + 1 header = 74
  const dRange  = `A2:AF${lastRow}`;    // full row highlight
  const aeRange = `AE2:AE${lastRow}`;   // Follow-Up Status (col 31)
  const adRange = `AD2:AD${lastRow}`;   // Days Pending (col 30)
  const afRange = `AF2:AF${lastRow}`;   // Data Quality Flag (col 32)
  const acRange = `AC2:AC${lastRow}`;   // Go/No-Go Rec (col 29)

  // Awarded rows → light green
  ws.addConditionalFormatting({ ref: dRange, rules: [{
    type: 'expression', formulae: ['$Q2="Awarded"'], priority: 1,
    style: { fill: { type:'pattern', pattern:'solid', fgColor:{argb:C.ltgrn} },
             font: { color:{argb:C.dkgrn}, bold:false } }
  }]});

  // Lost rows → light red
  ws.addConditionalFormatting({ ref: dRange, rules: [{
    type: 'expression', formulae: ['$Q2="Lost"'], priority: 2,
    style: { fill: { type:'pattern', pattern:'solid', fgColor:{argb:C.ltred} },
             font: { color:{argb:C.dkred}, bold:false } }
  }]});

  // Follow-Up Status: OVERDUE → orange
  ws.addConditionalFormatting({ ref: aeRange, rules: [{
    type: 'expression', formulae: ['AE2="OVERDUE"'], priority: 3,
    style: { fill: { type:'pattern', pattern:'solid', fgColor:{argb:C.lrust} },
             font: { color:{argb:C.rust}, bold:true } }
  }]});

  // Follow-Up Status: DUE SOON → yellow
  ws.addConditionalFormatting({ ref: aeRange, rules: [{
    type: 'expression', formulae: ['AE2="DUE SOON"'], priority: 4,
    style: { fill: { type:'pattern', pattern:'solid', fgColor:{argb:C.ltyel} },
             font: { color:{argb:C.dkyel}, bold:true } }
  }]});

  // Follow-Up Status: MISSING → light purple
  ws.addConditionalFormatting({ ref: aeRange, rules: [{
    type: 'expression', formulae: ['AE2="MISSING"'], priority: 5,
    style: { fill: { type:'pattern', pattern:'solid', fgColor:{argb:C.ltpurp} },
             font: { color:{argb:C.purple}, bold:true } }
  }]});

  // Days Pending > 60 → red
  ws.addConditionalFormatting({ ref: adRange, rules: [{
    type: 'expression', formulae: ['AND(AD2>60,AD2<>"")'], priority: 6,
    style: { fill: { type:'pattern', pattern:'solid', fgColor:{argb:C.ltred} },
             font: { color:{argb:C.dkred}, bold:true } }
  }]});

  // Days Pending 31-60 → yellow
  ws.addConditionalFormatting({ ref: adRange, rules: [{
    type: 'expression', formulae: ['AND(AD2>30,AD2<=60,AD2<>"")'], priority: 7,
    style: { fill: { type:'pattern', pattern:'solid', fgColor:{argb:C.ltyel} },
             font: { color:{argb:C.dkyel}, bold:true } }
  }]});

  // Data Quality not OK → light red
  ws.addConditionalFormatting({ ref: afRange, rules: [{
    type: 'expression', formulae: ['AF2<>"OK"'], priority: 8,
    style: { fill: { type:'pattern', pattern:'solid', fgColor:{argb:C.ltred} },
             font: { color:{argb:C.dkred} } }
  }]});

  // Go/No-Go Rec coloring (col 29 = AC)
  ws.addConditionalFormatting({ ref: acRange, rules: [
    { type:'expression', formulae:['AC2="Go"'],     priority:9,  style:{fill:{type:'pattern',pattern:'solid',fgColor:{argb:C.ltgrn}},font:{color:{argb:C.dkgrn},bold:true}} },
    { type:'expression', formulae:['AC2="No-Go"'],  priority:10, style:{fill:{type:'pattern',pattern:'solid',fgColor:{argb:C.ltred}},font:{color:{argb:C.dkred},bold:true}} },
    { type:'expression', formulae:['AC2="Review"'], priority:11, style:{fill:{type:'pattern',pattern:'solid',fgColor:{argb:C.ltyel}},font:{color:{argb:C.dkyel},bold:true}} },
  ]});
}

// ══════════════════════════════════════════════════════════════════════════════
// GO / NO-GO TAB
// ══════════════════════════════════════════════════════════════════════════════
const GNG_COLS = [
  { name: 'Bid ID',                       width: 12 },
  { name: 'Project Name',                  width: 30 },
  { name: 'GC',                            width: 22 },
  { name: 'GC Relationship\n(1-5)',        width: 14 },
  { name: 'Project Fit\n(1-5)',            width: 13 },
  { name: 'Drawing Quality\n(1-5)',        width: 14 },
  { name: 'Schedule Risk\n(1-5)',          width: 13 },
  { name: 'Labor Available\n(1-5)',        width: 14 },
  { name: 'Margin Potential\n(1-5)',       width: 14 },
  { name: 'Strategic Value\n(1-5)',        width: 14 },
  { name: 'Competition\n(1-5)',            width: 13 },
  { name: 'Total Score\n(/40)',            width: 13 },
  { name: 'Score %',                       width: 11 },
  { name: 'Recommendation',               width: 16 },
];

const GNG_SCORES = [
  { id:'ASM-2026-001', scores:[4,5,4,4,4,4,4,3] },
  { id:'ASM-2026-003', scores:[3,5,3,3,4,4,5,3] },
  { id:'ASM-2026-005', scores:[5,5,4,4,5,4,5,4] },
  { id:'ASM-2026-007', scores:[2,4,3,3,3,4,5,2] },
  { id:'ASM-2026-009', scores:[3,4,5,4,3,4,3,3] },
  { id:'ASM-2026-011', scores:[3,3,2,4,4,3,3,4] },
  { id:'ASM-2026-012', scores:[4,4,3,2,3,5,5,3] },
  { id:'ASM-2026-013', scores:[3,4,3,4,4,4,4,4] },
  { id:'ASM-2026-015', scores:[2,5,2,2,3,5,5,2] },
];

const GNG_TIPS = [
  '5 = Strong / partner GC',
  '5 = Ideal project type & size',
  '5 = Complete clear drawings',
  '5 = Low risk flexible sched',
  '5 = Crew fully available',
  '5 = Strong margin opp.',
  '5 = High strategic importance',
  '5 = Low competition / sole-src',
];

function buildGonogo(ws) {
  ws.properties.tabColor = { argb: 'FFC55A11' };
  ws.views = [{ state:'frozen', xSplit:3, ySplit:2 }];
  ws.getRow(1).height = 50;

  GNG_COLS.forEach((col, i) => {
    ws.getColumn(i+1).width = col.width;
    const cell = ws.getCell(1, i+1);
    applyHeader(cell, col.name, { size:9, wrap:true });
  });

  // Scoring tip row
  ws.getRow(2).height = 30;
  ws.mergeCells(2, 1, 2, 3);
  const tipLbl = ws.getCell(2, 1);
  tipLbl.value = 'SCORING GUIDE →';
  tipLbl.font  = font({ size:8, bold:true, color:C.dgray });
  tipLbl.fill  = fill('FFFFF2CC');
  tipLbl.alignment = align('center','center');

  GNG_TIPS.forEach((tip, i) => {
    const cell = ws.getCell(2, 4+i);
    cell.value = tip;
    cell.font  = font({ size:7, italic:true, color:C.dgray });
    cell.fill  = fill('FFFFF2CC');
    cell.alignment = align('center','center',true);
  });

  GNG_SCORES.forEach((entry, i) => {
    const r = i + 3;
    ws.getRow(r).height = 18;
    const bg = r % 2 === 0 ? C.lgray : C.white;

    // Bid ID
    const idCell = ws.getCell(r, 1);
    idCell.value = entry.id;
    idCell.font = font({size:10}); idCell.fill = fill(bg);
    idCell.alignment = align('center'); idCell.border = border();

    // Project (VLOOKUP)
    const projCell = ws.getCell(r, 2);
    projCell.value = { formula: `IFERROR(VLOOKUP(A${r},'Bid Log'!$A:$B,2,0),"")` };
    projCell.font = font({size:10}); projCell.fill = fill(bg);
    projCell.alignment = align('left'); projCell.border = border();

    // GC (VLOOKUP)
    const gcCell = ws.getCell(r, 3);
    gcCell.value = { formula: `IFERROR(VLOOKUP(A${r},'Bid Log'!$A:$C,3,0),"")` };
    gcCell.font = font({size:10}); gcCell.fill = fill(bg);
    gcCell.alignment = align('left'); gcCell.border = border();

    // Scores (cols 4-11)
    entry.scores.forEach((score, si) => {
      const cell = ws.getCell(r, 4+si);
      cell.value = score;
      cell.font = font({size:10,bold:true}); cell.fill = fill(bg);
      cell.alignment = align('center'); cell.border = border();
      cell.numFmt = '0';
    });

    // Total Score (col 12)
    const totCell = ws.getCell(r, 12);
    totCell.value = { formula: `SUM(D${r}:K${r})` };
    totCell.font = font({size:10,bold:true}); totCell.fill = fill(bg);
    totCell.alignment = align('center'); totCell.border = border();
    totCell.numFmt = '0';

    // Score % (col 13)
    const pctCell = ws.getCell(r, 13);
    pctCell.value = { formula: `IF(L${r}>0,L${r}/40,"")` };
    pctCell.font = font({size:10,bold:true}); pctCell.fill = fill(bg);
    pctCell.alignment = align('center'); pctCell.border = border();
    pctCell.numFmt = '0%';

    // Recommendation (col 14)
    const recCell = ws.getCell(r, 14);
    recCell.value = { formula: `IF(M${r}="","",IF(M${r}>=0.70,"Go",IF(M${r}>=0.50,"Review","No-Go")))` };
    recCell.font = font({size:10,bold:true}); recCell.fill = fill(bg);
    recCell.alignment = align('center'); recCell.border = border();
  });

  const lastR = GNG_SCORES.length + 2;

  // Conditional formatting
  ws.addConditionalFormatting({ ref:`N3:N${lastR}`, rules:[
    {type:'expression',formulae:['N3="Go"'],    priority:1, style:{fill:{type:'pattern',pattern:'solid',fgColor:{argb:C.ltgrn}},font:{color:{argb:C.dkgrn},bold:true}}},
    {type:'expression',formulae:['N3="No-Go"'], priority:2, style:{fill:{type:'pattern',pattern:'solid',fgColor:{argb:C.ltred}},font:{color:{argb:C.dkred},bold:true}}},
    {type:'expression',formulae:['N3="Review"'],priority:3, style:{fill:{type:'pattern',pattern:'solid',fgColor:{argb:C.ltyel}},font:{color:{argb:C.dkyel},bold:true}}},
  ]});

  ws.addConditionalFormatting({ ref:`M3:M${lastR}`, rules:[
    {type:'expression',formulae:['AND(M3>=0.70,M3<>"")'],priority:4,style:{fill:{type:'pattern',pattern:'solid',fgColor:{argb:C.ltgrn}}}},
    {type:'expression',formulae:['AND(M3<0.50,M3<>"")'], priority:5,style:{fill:{type:'pattern',pattern:'solid',fgColor:{argb:C.ltred}}}},
  ]});

  ws.addConditionalFormatting({ ref:`D3:K${lastR}`, rules:[
    {type:'expression',formulae:['D3<=2'],priority:6,style:{fill:{type:'pattern',pattern:'solid',fgColor:{argb:'FFFFD0D0'}}}},
    {type:'expression',formulae:['D3>=4'],priority:7,style:{fill:{type:'pattern',pattern:'solid',fgColor:{argb:'FFD0F0D0'}}}},
  ]});

  // Key row
  const noteRow = lastR + 2;
  ws.mergeCells(noteRow, 1, noteRow, 14);
  const noteCell = ws.getCell(noteRow, 1);
  noteCell.value = 'SCORING KEY:  Each category scored 1–5.  Max = 40.  |  GO = 70%+ (≥28 pts)   |   REVIEW = 50–69% (20–27 pts)   |   NO-GO = below 50% (<20 pts)';
  noteCell.font  = font({ size:9, bold:true, color:C.dgray });
  noteCell.fill  = fill('FFFFF2CC');
  noteCell.alignment = align('left','center');
  ws.getRow(noteRow).height = 20;
}

// ══════════════════════════════════════════════════════════════════════════════
// DASHBOARD TAB
// ══════════════════════════════════════════════════════════════════════════════
function buildDashboard(ws) {
  ws.properties.tabColor = { argb: 'FF375623' };
  ws.views = [{ showGridLines: false }];

  // Zone 1 (KPI cards): 5 cards × 2 cols = cols 2–11, each card 28 wide
  // Zone 2 (tables):    cols 2–12, merged differently per table
  //   Market Seg:  [2-4]=label  [5]=# Bids  [6-7]=Bid Vol  [8]=# Awd  [9]=# Lost  [10-11]=Awd$  [12]=WR
  //   Estimator:   [2-3]=label  [4]=Active  [5-6]=ActVol   [7]=Awd    [8-9]=Awd$  [10]=Lost      [11]=WR
  //   Losses:      [2-5]=reason [6-7]=# Lost                           [8-12]=Lost Vol
  const widths = [2, 14,14,14, 14,14, 14,14, 14,14,14,14, 2];
  widths.forEach((w, i) => { ws.getColumn(i + 1).width = w; });

  // ── Helpers ────────────────────────────────────────────────────────────────
  function mergeHdr(row, c1, c2, text) {
    if (c1 !== c2) ws.mergeCells(row, c1, row, c2);
    applyHeader(ws.getCell(row, c1), text, { bg: C.blue, size: 9 });
  }
  function mergeVal(row, c1, c2, formula, bg, numFmt) {
    if (c1 !== c2) ws.mergeCells(row, c1, row, c2);
    const cell = ws.getCell(row, c1);
    cell.value = { formula };
    cell.font = font({ size: 9 }); cell.fill = fill(bg);
    cell.alignment = align('center', 'center'); cell.border = border();
    if (numFmt) cell.numFmt = numFmt;
  }
  function mergeLabel(row, c1, c2, text, bg) {
    if (c1 !== c2) ws.mergeCells(row, c1, row, c2);
    applyValue(ws.getCell(row, c1), text, { bg, size: 9 });
  }

  // ── Title ──────────────────────────────────────────────────────────────────
  ws.getRow(1).height = 36;
  titleBar(ws, 1, 2, 12, 'ASM PRECONSTRUCTION PIPELINE TRACKER V9  —  DASHBOARD');

  ws.getRow(2).height = 20;
  ws.mergeCells(2, 2, 2, 12);
  const sub = ws.getCell(2, 2);
  sub.value = { formula: `"All States Mechanical  |  Data as of: "&TEXT(TODAY(),"mmmm d, yyyy")` };
  sub.font = font({ size: 10, color: C.dgray });
  sub.alignment = align('center', 'center');

  // ── KPI rows ───────────────────────────────────────────────────────────────
  // 5 equal cards: [2-3], [4-5], [6-7], [8-9], [10-11]
  const kCards = [[2,3],[4,5],[6,7],[8,9],[10,12]];

  ws.getRow(3).height = 18;
  ws.getRow(4).height = 36;
  ws.getRow(5).height = 8;
  ws.getRow(6).height = 18;
  ws.getRow(7).height = 36;

  function kpiBox(lRow, vRow, idx, label, formula, bg, numFmt = '$#,##0', valSize = 18) {
    const [c1, c2] = kCards[idx];
    ws.mergeCells(lRow, c1, lRow, c2);
    const lc = ws.getCell(lRow, c1);
    lc.value = label; lc.font = font({ size: 9, bold: true, color: C.white });
    lc.fill = fill(bg); lc.alignment = align('center', 'center', true);
    ws.mergeCells(vRow, c1, vRow, c2);
    const vc = ws.getCell(vRow, c1);
    vc.value = { formula }; vc.font = font({ size: valSize, bold: true, color: C.white });
    vc.fill = fill(bg); vc.alignment = align('center', 'center');
    vc.numFmt = numFmt;
  }

  kpiBox(3,4, 0, 'TOTAL BID VOLUME YTD',
    `IFERROR(SUMIF('Bid Log'!Q:Q,"<>No Bid",'Bid Log'!P:P),0)`, C.navy);
  kpiBox(3,4, 1, 'AWARDED VOLUME YTD',
    `IFERROR(SUMIF('Bid Log'!Q:Q,"Awarded",'Bid Log'!R:R),0)`, C.dkgrn);
  kpiBox(3,4, 2, 'LOST VOLUME YTD',
    `IFERROR(SUMIF('Bid Log'!Q:Q,"Lost",'Bid Log'!P:P),0)`, C.dkred);
  kpiBox(3,4, 3, 'PENDING VOLUME',
    `IFERROR(SUMIF('Bid Log'!Q:Q,"Bid Submitted",'Bid Log'!P:P)+SUMIF('Bid Log'!Q:Q,"Budget Submitted",'Bid Log'!P:P)+SUMIF('Bid Log'!Q:Q,"ROM Submitted",'Bid Log'!P:P),0)`, C.blue);
  kpiBox(3,4, 4, 'WIN RATE (BY $)',
    `IFERROR(SUMIF('Bid Log'!Q:Q,"Awarded",'Bid Log'!P:P)/(SUMIF('Bid Log'!Q:Q,"Awarded",'Bid Log'!P:P)+SUMIF('Bid Log'!Q:Q,"Lost",'Bid Log'!P:P)),0)`, C.teal, '0.0%', 22);

  kpiBox(6,7, 0, 'HIGH-PROB PENDING (≥70%)',
    `IFERROR(SUMPRODUCT((('Bid Log'!Q:Q="Bid Submitted")+('Bid Log'!Q:Q="Budget Submitted")>0)*('Bid Log'!W:W>=0.7)*'Bid Log'!P:P),0)`, 'FF003B4C');
  kpiBox(6,7, 1, 'BIDS DUE NEXT 7 DAYS',
    `IFERROR(COUNTIFS('Bid Log'!L:L,">="&TODAY(),'Bid Log'!L:L,"<="&TODAY()+7,'Bid Log'!Q:Q,"In Progress"),0)`, C.orange, '0', 26);
  kpiBox(6,7, 2, 'FOLLOW-UPS PAST DUE',
    `IFERROR(COUNTIF('Bid Log'!AE:AE,"OVERDUE"),0)`, C.dkred, '0', 26);
  kpiBox(6,7, 3, 'EXPECTED AWARDS THIS MONTH',
    `IFERROR(COUNTIFS('Bid Log'!V:V,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),'Bid Log'!V:V,"<="&EOMONTH(TODAY(),0),'Bid Log'!Q:Q,"Bid Submitted"),0)`, C.purple, '0', 26);
  kpiBox(6,7, 4, 'WIN RATE (BY COUNT)',
    `IFERROR(COUNTIF('Bid Log'!Q:Q,"Awarded")/(COUNTIF('Bid Log'!Q:Q,"Awarded")+COUNTIF('Bid Log'!Q:Q,"Lost")),0)`, C.teal, '0.0%', 22);

  ws.getRow(8).height = 12;

  // ── Market Segment table ───────────────────────────────────────────────────
  let r = 9;
  sectionBar(ws, r, 2, 12, '  BID VOLUME BY MARKET SEGMENT', { bg: C.navy }); r++;
  ws.getRow(r).height = 18;
  mergeHdr(r, 2,  4, 'Market Segment');
  mergeHdr(r, 5,  5, '# Bids');
  mergeHdr(r, 6,  7, 'Bid Volume');
  mergeHdr(r, 8,  8, '# Awarded');
  mergeHdr(r, 9,  9, '# Lost');
  mergeHdr(r, 10, 11, 'Awarded $');
  mergeHdr(r, 12, 12, 'Win Rate');
  r++;

  SETTINGS_LISTS.find(l => l.title === 'MARKET SEGMENTS').values.forEach((seg, i) => {
    ws.getRow(r).height = 16;
    const bg = i % 2 === 0 ? C.lgray : C.white;
    const q  = `"${seg}"`;
    mergeLabel(r, 2,  4, seg, bg);
    mergeVal(r, 5,  5,  `COUNTIFS('Bid Log'!E:E,${q},'Bid Log'!Q:Q,"<>No Bid")`,                              bg, '0');
    mergeVal(r, 6,  7,  `IFERROR(SUMIFS('Bid Log'!P:P,'Bid Log'!E:E,${q},'Bid Log'!Q:Q,"<>No Bid"),0)`,       bg, '$#,##0');
    mergeVal(r, 8,  8,  `COUNTIFS('Bid Log'!E:E,${q},'Bid Log'!Q:Q,"Awarded")`,                               bg, '0');
    mergeVal(r, 9,  9,  `COUNTIFS('Bid Log'!E:E,${q},'Bid Log'!Q:Q,"Lost")`,                                  bg, '0');
    mergeVal(r, 10, 11, `IFERROR(SUMIFS('Bid Log'!R:R,'Bid Log'!E:E,${q},'Bid Log'!Q:Q,"Awarded"),0)`,        bg, '$#,##0');
    mergeVal(r, 12, 12, `IFERROR(H${r}/(H${r}+I${r}),"")`,                                                    bg, '0%');
    r++;
  });

  // ── Plumb Estimator table ──────────────────────────────────────────────────
  r++;
  sectionBar(ws, r, 2, 12, '  PLUMB ESTIMATOR WORKLOAD & PERFORMANCE', { bg: C.navy }); r++;
  ws.getRow(r).height = 18;
  mergeHdr(r, 2,  3, 'Plumb Estimator');
  mergeHdr(r, 4,  4, 'Active Bids');
  mergeHdr(r, 5,  6, 'Active Volume');
  mergeHdr(r, 7,  7, 'Awarded');
  mergeHdr(r, 8,  9, 'Awarded $');
  mergeHdr(r, 10, 10, 'Lost');
  mergeHdr(r, 11, 11, 'Win Rate');
  r++;

  SETTINGS_LISTS.find(l => l.title === 'PLUMB ESTIMATORS').values.forEach((est, i) => {
    ws.getRow(r).height = 16;
    const bg = i % 2 === 0 ? C.lgray : C.white;
    const q  = `"${est}"`;
    mergeLabel(r, 2, 3, est, bg);
    mergeVal(r, 4, 4,
      `COUNTIFS('Bid Log'!I:I,${q},'Bid Log'!Q:Q,"Bid Submitted")+COUNTIFS('Bid Log'!I:I,${q},'Bid Log'!Q:Q,"Budget Submitted")+COUNTIFS('Bid Log'!I:I,${q},'Bid Log'!Q:Q,"In Progress")`,
      bg, '0');
    mergeVal(r, 5, 6,
      `IFERROR(SUMIFS('Bid Log'!P:P,'Bid Log'!I:I,${q},'Bid Log'!Q:Q,"Bid Submitted")+SUMIFS('Bid Log'!P:P,'Bid Log'!I:I,${q},'Bid Log'!Q:Q,"Budget Submitted")+SUMIFS('Bid Log'!P:P,'Bid Log'!I:I,${q},'Bid Log'!Q:Q,"In Progress"),0)`,
      bg, '$#,##0');
    mergeVal(r, 7,  7, `COUNTIFS('Bid Log'!I:I,${q},'Bid Log'!Q:Q,"Awarded")`,                             bg, '0');
    mergeVal(r, 8,  9, `IFERROR(SUMIFS('Bid Log'!R:R,'Bid Log'!I:I,${q},'Bid Log'!Q:Q,"Awarded"),0)`,      bg, '$#,##0');
    mergeVal(r, 10, 10, `COUNTIFS('Bid Log'!I:I,${q},'Bid Log'!Q:Q,"Lost")`,                               bg, '0');
    mergeVal(r, 11, 11, `IFERROR(G${r}/(G${r}+J${r}),"")`,                                                 bg, '0%');
    r++;
  });

  // ── Losses by Reason ───────────────────────────────────────────────────────
  r++;
  sectionBar(ws, r, 2, 12, '  LOSSES BY REASON', { bg: C.navy }); r++;
  ws.getRow(r).height = 18;
  mergeHdr(r, 2, 5, 'Lost Reason');
  mergeHdr(r, 6, 7, '# Lost');
  mergeHdr(r, 8, 12, 'Lost Volume');
  r++;

  SETTINGS_LISTS.find(l => l.title === 'LOST REASONS').values.forEach((rsn, i) => {
    ws.getRow(r).height = 16;
    const bg = i % 2 === 0 ? C.lgray : C.white;
    const q  = `"${rsn}"`;
    mergeLabel(r, 2, 5, rsn, bg);
    mergeVal(r, 6,  7,  `COUNTIFS('Bid Log'!U:U,${q},'Bid Log'!Q:Q,"Lost")`,                              bg, '0');
    mergeVal(r, 8,  12, `IFERROR(SUMIFS('Bid Log'!P:P,'Bid Log'!U:U,${q},'Bid Log'!Q:Q,"Lost"),0)`,       bg, '$#,##0');
    r++;
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// FRIDAY MEETING TAB
// ══════════════════════════════════════════════════════════════════════════════
function buildFriday(ws) {
  ws.properties.tabColor = { argb: 'FF4B0082' };
  ws.views = [{ showGridLines: false }];
  const widths = [2,25,20,16,18,18,14,14,2];
  widths.forEach((w,i)=>{ ws.getColumn(i+1).width = w; });

  ws.getRow(1).height = 34;
  ws.mergeCells(1,2,1,8);
  const t = ws.getCell(1,2);
  t.value = { formula: `"ASM FRIDAY ESTIMATING MEETING  |  "&TEXT(TODAY(),"mmmm d, yyyy")` };
  t.font  = font({size:14,bold:true,color:C.white});
  t.fill  = fill(C.navy);
  t.alignment = align('center','center');

  ws.getRow(2).height = 22;
  ws.mergeCells(2,2,2,8);
  const sub = ws.getCell(2,2);
  sub.value = 'Review each section with the team. Use the Bid Log AutoFilter to pull detailed lists for each item.';
  sub.font  = font({size:10,italic:true,color:C.dgray});
  sub.alignment = align('left','center');

  let r = 3;

  function friSection(title, bg) {
    ws.getRow(r).height = 8; r++;
    sectionBar(ws, r, 2, 8, `  ${title}`, {bg}); r++;
  }

  function metric(label, formula, numFmt='0') {
    ws.getRow(r).height = 20;
    ws.mergeCells(r,2,r,4);
    const c1 = ws.getCell(r,2);
    c1.value = label; c1.font = font({size:10}); c1.fill = fill(C.lgray);
    c1.alignment = align('left','center'); c1.border = border();
    const c2 = ws.getCell(r,5);
    c2.value = {formula}; c2.font = font({size:11,bold:true});
    c2.fill = fill(C.white); c2.alignment = align('center'); c2.border = border();
    c2.numFmt = numFmt;
    r++;
  }

  function hint(text) {
    ws.getRow(r).height = 16;
    ws.mergeCells(r,2,r,8);
    const c = ws.getCell(r,2);
    c.value = text; c.font = font({size:9,italic:true,color:C.dgray});
    c.fill = fill('FFFFFFFA'); c.alignment = align('left','center');
    r++;
  }

  friSection('1.  BIDS DUE IN NEXT 14 DAYS', C.blue);
  metric('Bids due in next 7 days (In Progress)',   `COUNTIFS('Bid Log'!L:L,">="&TODAY(),'Bid Log'!L:L,"<="&TODAY()+7,'Bid Log'!Q:Q,"In Progress")`);
  metric('Bids due in days 8–14 (In Progress)',     `COUNTIFS('Bid Log'!L:L,">="&TODAY()+8,'Bid Log'!L:L,"<="&TODAY()+14,'Bid Log'!Q:Q,"In Progress")`);
  metric('Total bid volume due in next 14 days',    `IFERROR(SUMIFS('Bid Log'!P:P,'Bid Log'!L:L,">="&TODAY(),'Bid Log'!L:L,"<="&TODAY()+14,'Bid Log'!Q:Q,"In Progress"),0)`, '$#,##0');
  hint('▶  Filter Bid Log: Status = In Progress | Bid Due Date between today and today+14');

  friSection('2.  FOLLOW-UP STATUS', C.dkred);
  metric('Follow-ups OVERDUE',                      `COUNTIF('Bid Log'!AE:AE,"OVERDUE")`);
  metric('Follow-ups due this week (3 days)',        `COUNTIF('Bid Log'!AE:AE,"DUE SOON")`);
  metric('Pending bids with NO follow-up date',     `COUNTIF('Bid Log'!AE:AE,"MISSING")`);
  metric('Pending bids older than 30 days',         `COUNTIFS('Bid Log'!AD:AD,">30",'Bid Log'!AD:AD,"<>")`);
  metric('Pending bids older than 60 days (URGENT)',`COUNTIFS('Bid Log'!AD:AD,">60",'Bid Log'!AD:AD,"<>")`);
  hint('▶  Filter Bid Log: Follow-Up Status = OVERDUE  |  Sort: Next Follow-Up Date ascending');

  friSection('3.  EXPECTED AWARDS THIS MONTH', C.dkgrn);
  metric('Bids expecting award this month',
    `COUNTIFS('Bid Log'!V:V,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),'Bid Log'!V:V,"<="&EOMONTH(TODAY(),0),'Bid Log'!Q:Q,"Bid Submitted")`);
  metric('Expected award volume this month',
    `IFERROR(SUMPRODUCT((MONTH('Bid Log'!V2:V2001)=MONTH(TODAY()))*(YEAR('Bid Log'!V2:V2001)=YEAR(TODAY()))*('Bid Log'!Q2:Q2001="Bid Submitted")*'Bid Log'!P2:P2001),0)`, '$#,##0');
  metric('High-probability pending bids (≥70%)',
    `IFERROR(SUMPRODUCT((('Bid Log'!Q2:Q2001="Bid Submitted")+('Bid Log'!Q2:Q2001="Budget Submitted")>0)*('Bid Log'!W2:W2001>=0.7)),0)`);
  metric('High-probability pending volume (≥70%)',
    `IFERROR(SUMPRODUCT((('Bid Log'!Q2:Q2001="Bid Submitted")+('Bid Log'!Q2:Q2001="Budget Submitted")>0)*('Bid Log'!W2:W2001>=0.7)*'Bid Log'!P2:P2001),0)`, '$#,##0');
  hint('▶  Filter Bid Log: Status = Bid Submitted | Expected Award Date = this month | Sort: Probability descending');

  friSection('4.  GO / NO-GO ITEMS TO REVIEW', C.orange);
  metric('Bids scored No-Go in Go-No-Go tab',       `COUNTIF('Bid Log'!AC:AC,"No-Go")`);
  metric('Bids scored Review (decision needed)',     `COUNTIF('Bid Log'!AC:AC,"Review")`);
  metric('Active bids with no Go/No-Go score',      `COUNTIFS('Bid Log'!Q:Q,"In Progress",'Bid Log'!AB:AB,"")`);
  hint('▶  Go to Go-No-Go tab  |  Filter Bid Log: Go/No-Go Rec = No-Go or Review');

  friSection('5.  RECENT AWARDS & LOSSES (LAST 30 DAYS)', C.teal);
  metric('Jobs awarded in last 30 days (count)',    `COUNTIFS('Bid Log'!Z:Z,">="&TODAY()-30,'Bid Log'!Q:Q,"Awarded")`);
  metric('Jobs awarded in last 30 days (volume)',   `IFERROR(SUMPRODUCT(('Bid Log'!Z2:Z2001>=TODAY()-30)*('Bid Log'!Q2:Q2001="Awarded")*'Bid Log'!R2:R2001),0)`, '$#,##0');
  metric('Jobs lost in last 30 days (count)',       `COUNTIFS('Bid Log'!Z:Z,">="&TODAY()-30,'Bid Log'!Q:Q,"Lost")`);
  metric('Jobs lost in last 30 days (volume)',      `IFERROR(SUMPRODUCT(('Bid Log'!Z2:Z2001>=TODAY()-30)*('Bid Log'!Q2:Q2001="Lost")*'Bid Log'!P2:P2001),0)`, '$#,##0');
  hint('▶  Filter Bid Log: Status = Awarded or Lost | Last Contact Date >= today-30');

  friSection('6.  ESTIMATOR WORKLOAD SNAPSHOT', C.navy);
  ws.getRow(r).height = 18;
  ['','Estimator','Active Bids','Active Volume','Due Next 14 Days','','',''].forEach((h,i)=>{
    if(h) applyHeader(ws.getCell(r,i+2), h, {bg:C.blue,size:9});
  });
  r++;

  estimators.forEach((est,i) => {
    ws.getRow(r).height = 16;
    const bg = i%2===0 ? C.lgray : C.white;
    const q = `"${est}"`;
    applyValue(ws.getCell(r,3), est, {bg,size:9});
    ws.getCell(r,4).value  = {formula:`COUNTIFS('Bid Log'!I:I,${q},'Bid Log'!Q:Q,"Bid Submitted")+COUNTIFS('Bid Log'!I:I,${q},'Bid Log'!Q:Q,"Budget Submitted")+COUNTIFS('Bid Log'!I:I,${q},'Bid Log'!Q:Q,"In Progress")`};
    ws.getCell(r,5).value  = {formula:`IFERROR(SUMIFS('Bid Log'!P:P,'Bid Log'!I:I,${q},'Bid Log'!Q:Q,"Bid Submitted")+SUMIFS('Bid Log'!P:P,'Bid Log'!I:I,${q},'Bid Log'!Q:Q,"Budget Submitted")+SUMIFS('Bid Log'!P:P,'Bid Log'!I:I,${q},'Bid Log'!Q:Q,"In Progress"),0)`};
    ws.getCell(r,6).value  = {formula:`COUNTIFS('Bid Log'!I:I,${q},'Bid Log'!L:L,">="&TODAY(),'Bid Log'!L:L,"<="&TODAY()+14)`};
    [3,4,5,6].forEach(c=>{
      const cell=ws.getCell(r,c);
      cell.font=font({size:9}); cell.fill=fill(bg);
      cell.alignment=align(c===3?'left':'center'); cell.border=border();
    });
    ws.getCell(r,5).numFmt = '$#,##0';
    r++;
  });

  ws.getRow(r).height = 24;
  ws.mergeCells(r,2,r,8);
  const alertCell = ws.getCell(r,2);
  alertCell.value = 'ALERT: Any estimator with 3+ active bids AND 2+ due in the next 14 days is at capacity. Discuss workload rebalancing at this meeting.';
  alertCell.font  = font({size:9,bold:true,color:C.orange});
  alertCell.fill  = fill(C.ltorang);
  alertCell.alignment = align('left','center',true);
}

// ══════════════════════════════════════════════════════════════════════════════
// FOLLOW-UP CONTROL TAB
// ══════════════════════════════════════════════════════════════════════════════
function buildFollowup(ws) {
  ws.properties.tabColor = { argb: 'FFC00000' };
  ws.views = [{ showGridLines: false }];
  const widths = [2,28,20,16,16,16,16,14,2];
  widths.forEach((w,i)=>{ ws.getColumn(i+1).width = w; });

  titleBar(ws, 1, 2, 8, 'ASM — FOLLOW-UP CONTROL CENTER', {bg:C.dkred});

  ws.getRow(2).height = 18;
  ws.mergeCells(2,2,2,8);
  const sub = ws.getCell(2,2);
  sub.value = 'Use this tab to manage follow-up on pending bids. All data pulls from the Bid Log.';
  sub.font = font({size:10,italic:true,color:C.dgray}); sub.alignment = align('left','center');

  let r = 3;
  function fuSection(title) {
    ws.getRow(r).height = 8; r++;
    sectionBar(ws, r, 2, 7, `  ${title}`, {bg:C.dkred}); r++;
  }
  function fuMetric(label, formula, numFmt='0') {
    ws.getRow(r).height = 20;
    ws.mergeCells(r,2,r,5);
    const c1 = ws.getCell(r,2);
    c1.value = label; c1.font = font({size:10}); c1.fill = fill(C.lgray);
    c1.alignment = align('left','center',true); c1.border = border();
    const c2 = ws.getCell(r,6);
    c2.value = {formula}; c2.font = font({size:11,bold:true});
    c2.fill = fill(C.white); c2.alignment = align('center'); c2.border = border();
    c2.numFmt = numFmt;
    r++;
  }

  fuSection('FOLLOW-UP SUMMARY');
  fuMetric('Total pending bids (Bid/Budget/ROM Submitted + In Progress)',
    `COUNTIF('Bid Log'!Q:Q,"Bid Submitted")+COUNTIF('Bid Log'!Q:Q,"Budget Submitted")+COUNTIF('Bid Log'!Q:Q,"ROM Submitted")+COUNTIF('Bid Log'!Q:Q,"In Progress")`);
  fuMetric('Follow-ups OVERDUE',                     `COUNTIF('Bid Log'!AE:AE,"OVERDUE")`);
  fuMetric('Follow-ups DUE within 3 days',           `COUNTIF('Bid Log'!AE:AE,"DUE SOON")`);
  fuMetric('Pending bids with NO follow-up date',    `COUNTIF('Bid Log'!AE:AE,"MISSING")`);
  fuMetric('Pending bids with NO follow-up owner',   `COUNTIFS('Bid Log'!Y:Y,"",'Bid Log'!Q:Q,"Bid Submitted")+COUNTIFS('Bid Log'!Y:Y,"",'Bid Log'!Q:Q,"Budget Submitted")`);

  fuSection('AGING — DAYS SINCE SUBMITTED OR RECEIVED');
  fuMetric('Pending bids: 0–30 days old (count)',    `COUNTIFS('Bid Log'!AD:AD,">0",'Bid Log'!AD:AD,"<=30")`);
  fuMetric('Pending bids: 31–60 days old (ATTENTION)',`COUNTIFS('Bid Log'!AD:AD,">30",'Bid Log'!AD:AD,"<=60")`);
  fuMetric('Pending bids: over 60 days old (URGENT)',`COUNTIFS('Bid Log'!AD:AD,">60",'Bid Log'!AD:AD,"<>")`);
  fuMetric('Pending volume: 0–30 days',              `IFERROR(SUMPRODUCT(('Bid Log'!AD2:AD2001>0)*('Bid Log'!AD2:AD2001<=30)*'Bid Log'!P2:P2001),0)`, '$#,##0');
  fuMetric('Pending volume: 31–60 days',             `IFERROR(SUMPRODUCT(('Bid Log'!AD2:AD2001>30)*('Bid Log'!AD2:AD2001<=60)*'Bid Log'!P2:P2001),0)`, '$#,##0');
  fuMetric('Pending volume: over 60 days',           `IFERROR(SUMPRODUCT(('Bid Log'!AD2:AD2001>60)*'Bid Log'!P2:P2001),0)`, '$#,##0');

  fuSection('HIGH-VALUE PENDING BIDS');
  fuMetric('Pending bids over $1,000,000 (count)',
    `COUNTIFS('Bid Log'!P:P,">1000000",'Bid Log'!Q:Q,"Bid Submitted")+COUNTIFS('Bid Log'!P:P,">1000000",'Bid Log'!Q:Q,"Budget Submitted")`);
  fuMetric('Total pending volume over $1M',
    `IFERROR(SUMPRODUCT(('Bid Log'!P2:P2001>1000000)*(('Bid Log'!Q2:Q2001="Bid Submitted")+('Bid Log'!Q2:Q2001="Budget Submitted")>0)*'Bid Log'!P2:P2001),0)`, '$#,##0');
  fuMetric('Bids over $500K with no follow-up date',
    `IFERROR(SUMPRODUCT(('Bid Log'!P2:P2001>500000)*(('Bid Log'!Q2:Q2001="Bid Submitted")+('Bid Log'!Q2:Q2001="Budget Submitted")>0)*('Bid Log'!X2:X2001="")),0)`);

  fuSection('FOLLOW-UP BY OWNER');
  ws.getRow(r).height = 18;
  ['Follow-Up Owner','Overdue','Due Soon','Missing Date','Total Pending'].forEach((h,i)=>{
    applyHeader(ws.getCell(r,3+i), h, {bg:C.dkred,size:9});
  });
  r++;

  const fuOwners = SETTINGS_LISTS.find(l=>l.title==='FOLLOW-UP OWNERS').values;
  fuOwners.forEach((owner,i)=>{
    ws.getRow(r).height = 16;
    const bg = i%2===0 ? C.lgray : C.white;
    const q = `"${owner}"`;
    applyValue(ws.getCell(r,3), owner, {bg,size:9});
    ws.getCell(r,4).value = {formula:`COUNTIFS('Bid Log'!Y:Y,${q},'Bid Log'!AE:AE,"OVERDUE")`};
    ws.getCell(r,5).value = {formula:`COUNTIFS('Bid Log'!Y:Y,${q},'Bid Log'!AE:AE,"DUE SOON")`};
    ws.getCell(r,6).value = {formula:`COUNTIFS('Bid Log'!Y:Y,${q},'Bid Log'!AE:AE,"MISSING")`};
    ws.getCell(r,7).value = {formula:`COUNTIFS('Bid Log'!Y:Y,${q},'Bid Log'!Q:Q,"Bid Submitted")+COUNTIFS('Bid Log'!Y:Y,${q},'Bid Log'!Q:Q,"Budget Submitted")`};
    [3,4,5,6,7].forEach(c=>{
      const cell=ws.getCell(r,c);
      cell.font=font({size:9}); cell.fill=fill(bg);
      cell.alignment=align(c===3?'left':'center'); cell.border=border();
    });
    r++;
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// DATA QUALITY TAB
// ══════════════════════════════════════════════════════════════════════════════
function buildDataQuality(ws) {
  ws.properties.tabColor = { argb: 'FFFFC000' };
  ws.views = [{ showGridLines: false }];
  [2,32,24,16,2].forEach((w,i)=>{ ws.getColumn(i+1).width = w; });

  titleBar(ws, 1, 2, 4, 'ASM — DATA QUALITY REPORT', {bg:C.dkyel,size:13});

  ws.getRow(2).height = 18;
  ws.mergeCells(2,2,2,4);
  const sub = ws.getCell(2,2);
  sub.value = 'Fix all flagged items before the Friday estimating meeting.  Filter Bid Log: Data Quality Flag ≠ OK.';
  sub.font = font({size:10,italic:true,color:C.dgray}); sub.alignment = align('left','center');

  const checks = [
    { section:'REQUIRED FIELDS', bg:C.navy, items:[
      ['Bids missing Project Name',         `COUNTIFS('Bid Log'!A:A,"<>",'Bid Log'!B:B,"")`],
      ['Bids missing General Contractor',   `COUNTIFS('Bid Log'!B:B,"<>",'Bid Log'!C:C,"")`],
      ['Bids missing Plumb Estimator',      `COUNTIFS('Bid Log'!B:B,"<>",'Bid Log'!I:I,"")`],
      ['Bids missing Bid Due Date',         `COUNTIFS('Bid Log'!B:B,"<>",'Bid Log'!L:L,"")`],
      ['Bids missing Total Bid Amount',     `COUNTIFS('Bid Log'!B:B,"<>",'Bid Log'!P:P,"",'Bid Log'!Q:Q,"<>In Progress",'Bid Log'!Q:Q,"<>No Bid",'Bid Log'!Q:Q,"<>On Hold")`],
    ]},
    { section:'AWARDED BIDS', bg:C.dkgrn, items:[
      ['Awarded bids with no Award Amount', `COUNTIFS('Bid Log'!Q:Q,"Awarded",'Bid Log'!R:R,"")`],
    ]},
    { section:'LOST BIDS', bg:C.dkred, items:[
      ['Lost bids with no Lost Reason',     `COUNTIFS('Bid Log'!Q:Q,"Lost",'Bid Log'!U:U,"")`],
    ]},
    { section:'FOLLOW-UP COMPLETENESS', bg:C.orange, items:[
      ['Pending bids with no Follow-Up Date',  `COUNTIF('Bid Log'!AE:AE,"MISSING")`],
      ['Pending bids with no Follow-Up Owner', `COUNTIFS('Bid Log'!Y:Y,"",'Bid Log'!Q:Q,"Bid Submitted")+COUNTIFS('Bid Log'!Y:Y,"",'Bid Log'!Q:Q,"Budget Submitted")`],
      ['Follow-ups past due',                  `COUNTIF('Bid Log'!AE:AE,"OVERDUE")`],
    ]},
    { section:'BID DUE DATE ALERTS', bg:C.blue, items:[
      ['Active bids with Bid Due Date already passed', `COUNTIFS('Bid Log'!L:L,"<"&TODAY(),'Bid Log'!Q:Q,"In Progress")`],
      ['Pending bids older than 60 days',              `COUNTIFS('Bid Log'!AD:AD,">60",'Bid Log'!AD:AD,"<>")`],
    ]},
    { section:'OVERALL SUMMARY', bg:C.dgray, items:[
      ['Total rows with ANY data quality issue', `COUNTIFS('Bid Log'!A:A,"<>",'Bid Log'!AF:AF,"<>OK")`],
    ]},
  ];

  let r = 3;
  let cfPri = 1;  // global priority counter — must be unique across entire worksheet
  checks.forEach(({ section, bg, items }) => {
    ws.getRow(r).height = 8; r++;
    sectionBar(ws, r, 2, 4, `  ${section}`, {bg}); r++;
    ws.getRow(r).height = 18;
    applyHeader(ws.getCell(r,2), 'Check',  {bg:C.blue,size:9});
    applyHeader(ws.getCell(r,3), 'Count',  {bg:C.blue,size:9});
    applyHeader(ws.getCell(r,4), 'Status', {bg:C.blue,size:9});
    r++;
    items.forEach(([label, formula]) => {
      ws.getRow(r).height = 20;
      const rbg = r%2===0 ? C.lgray : C.white;
      const lc = ws.getCell(r,2);
      lc.value = label; lc.font = font({size:9}); lc.fill = fill(rbg);
      lc.alignment = align('left','center',true); lc.border = border();
      const cc = ws.getCell(r,3);
      cc.value = {formula}; cc.font = font({size:11,bold:true});
      cc.fill = fill(rbg); cc.alignment = align('center'); cc.border = border();
      cc.numFmt = '0';
      const sc = ws.getCell(r,4);
      sc.value = {formula:`IF(C${r}=0,"OK","*** FIX NEEDED")`};
      sc.font = font({size:9,bold:true}); sc.fill = fill(rbg);
      sc.alignment = align('center'); sc.border = border();
      ws.addConditionalFormatting({ ref:`D${r}:D${r}`, rules:[
        {type:'expression',formulae:[`C${r}=0`], priority:cfPri++, style:{fill:{type:'pattern',pattern:'solid',fgColor:{argb:C.ltgrn}},font:{color:{argb:C.dkgrn},bold:true}}},
        {type:'expression',formulae:[`C${r}>0`], priority:cfPri++, style:{fill:{type:'pattern',pattern:'solid',fgColor:{argb:C.ltred}},font:{color:{argb:C.dkred},bold:true}}},
      ]});
      r++;
    });
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// INSTRUCTIONS TAB
// ══════════════════════════════════════════════════════════════════════════════
function buildInstructions(ws) {
  ws.properties.tabColor = { argb: 'FF4472C4' };
  ws.views = [{ showGridLines: false }];
  [2,24,68].forEach((w,i)=>{ ws.getColumn(i+1).width = w; });

  titleBar(ws, 1, 2, 3, 'ASM PRECONSTRUCTION PIPELINE TRACKER V9  —  INSTRUCTIONS & SOP', {size:13});

  let r = 2;
  function instSection(title, bg=C.navy) {
    ws.getRow(r).height = 8; r++;
    sectionBar(ws, r, 2, 3, `  ${title}`, {bg,height:26}); r++;
  }
  function instRow(label, content) {
    ws.getRow(r).height = 50;
    const c1 = ws.getCell(r,2);
    c1.value = label; c1.font = font({size:10,bold:true});
    c1.fill = fill(r%2===0?C.lgray:C.white);
    c1.alignment = align('left','top',true); c1.border = border();
    const c2 = ws.getCell(r,3);
    c2.value = content; c2.font = font({size:10});
    c2.fill = fill(r%2===0?C.lgray:C.white);
    c2.alignment = align('left','top',true); c2.border = border();
    r++;
  }

  instSection('WHO UPDATES THIS TRACKER');
  instRow('Primary Owner',   'The assigned estimator updates their own bids. The Preconstruction Manager or VP of Preconstruction reviews and spot-checks weekly.');
  instRow('Backup / Admin',  'If an estimator is unavailable, the Preconstruction Manager or office admin updates the tracker based on team notes.');

  instSection('WHEN TO UPDATE');
  instRow('Bid received',    'Add a new row immediately. Fill in: Project Name, GC, Location, Market Segment, Scope Type, Bid Stage, Estimator, Bid Received Date, Bid Due Date. Set Status = In Progress.');
  instRow('Bid submitted',   'Update: Submitted Date, ASM Bid Amount, Status = Bid Submitted. Set Next Follow-Up Date (5–7 business days out). Assign Follow-Up Owner.');
  instRow('Following up',    'Update: Last Contact Date, Notes (brief summary). Reset Next Follow-Up Date to next expected contact date.');
  instRow('Job awarded',     'Update: Status = Awarded, ASM Awarded Amount, Last Contact Date. Clear follow-up fields.');
  instRow('Job lost',        'Update: Status = Lost, Winning Competitor, Competitor Amount (if known), Lost Reason. Add note with any intel gathered.');
  instRow('Before Friday mtg','All estimators update their bids by Thursday EOD. Check Data Quality tab — all flags should read OK before the meeting.');

  instSection('STATUS DEFINITIONS');
  instRow('In Progress',       'Drawings received, takeoff and pricing underway. Bid not yet submitted.');
  instRow('ROM Submitted',     'Rough Order of Magnitude pricing provided. Not a formal bid.');
  instRow('Budget Submitted',  'Preliminary budget pricing submitted. Used in early design phases (SD/DD).');
  instRow('Bid Submitted',     'Formal bid submitted. Awaiting decision. Follow-up required.');
  instRow('Awarded',           'ASM selected for the project. Contract or LOI received.');
  instRow('Lost',              'GC or owner awarded work to a competitor. Fill in Winning Competitor, Amount, and Lost Reason.');
  instRow('No Bid',            'ASM chose not to bid. Note reason in Notes column.');
  instRow('On Hold',           'Project paused or decision delayed. Monitor for reactivation.');

  instSection('GO / NO-GO SCORING GUIDE');
  instRow('Purpose',           'Use the Go-No-Go tab to score each bid before submitting. This helps leadership decide whether ASM should invest estimating time on a project.');
  instRow('How to score',      'Score each of the 8 categories from 1 (very low) to 5 (very high). Total out of 40. Recommendation calculates automatically.');
  instRow('Recommendations',   'GO = 70%+ (score ≥28).  REVIEW = 50–69% (score 20–27) — bring to leadership.  NO-GO = below 50% (score <20) — recommend passing unless strategic reason exists.');
  instRow('Override policy',   'Leadership can override No-Go to Go for strategic reasons. Document the reason in the Notes column of the Bid Log.');

  instSection('FRIDAY MEETING AGENDA GUIDE');
  instRow('Thursday EOD prep', '1. Update all bid statuses.\n2. Set next follow-up dates.\n3. Check Data Quality tab — all rows should show OK.\n4. Score new bids in the Go-No-Go tab.');
  instRow('Review at meeting', '1. Dashboard — overall pipeline health.\n2. Friday Meeting tab — bids due, follow-ups, expected awards.\n3. Go-No-Go tab — No-Go or Review items.\n4. Follow-Up Control — overdue or missing follow-ups.');
  instRow('Key questions',     '1. Are we bidding enough to feed operations?\n2. Does any estimator need workload relief?\n3. Which pending bids are most likely to convert this month?\n4. Any stale bids (60+ days) to close out?\n5. What did we learn from recent losses?');

  instSection('TIPS FOR KEEPING THE TRACKER CLEAN');
  instRow('Use dropdowns',     'Always select from the dropdown lists. Do not type statuses, segments, or names manually — this breaks filters and dashboard formulas.');
  instRow('One row per bid',   'Each bid gets one row. If a project re-bids, add a new row with a new Bid ID and note the relationship in Notes.');
  instRow('Don\'t delete rows','Lost and No Bid rows must stay — they are needed for win rate calculations and historical analysis.');
  instRow('Probability guide', '0% = Lost/No Bid.  25% = Long shot.  50% = Competitive, no clear advantage.  75% = Strong position.  100% = Awarded.');
}

// ══════════════════════════════════════════════════════════════════════════════
// UTILITIES
// ══════════════════════════════════════════════════════════════════════════════
function colToLetter(n) {
  let s = '';
  while (n > 0) { s = String.fromCharCode(65 + ((n - 1) % 26)) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

// Returns excel column letter by 1-based index (for formula strings)
function colToLetter2(n) { return colToLetter(n); }

// Also attach to ExcelJS column objects since getColumn().letter may not exist
const _clCache = {};
function cl(n) {
  if (_clCache[n]) return _clCache[n];
  _clCache[n] = colToLetter(n);
  return _clCache[n];
}

// ── Patch column.letter onto ExcelJS (it may not expose this) ─────────────────
const estimators = SETTINGS_LISTS.find(l=>l.title==='PLUMB ESTIMATORS').values;

// ══════════════════════════════════════════════════════════════════════════════
// MAIN
// ══════════════════════════════════════════════════════════════════════════════
async function main() {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'All States Mechanical';
  wb.created  = new Date();
  wb.modified = new Date();

  const wsDash  = wb.addWorksheet('Dashboard');
  const wsBid   = wb.addWorksheet('Bid Log');
  const wsGng   = wb.addWorksheet('Go-No-Go');
  const wsFri   = wb.addWorksheet('Friday Meeting');
  const wsFu    = wb.addWorksheet('Follow-Up Control');
  const wsDq    = wb.addWorksheet('Data Quality');
  const wsSet   = wb.addWorksheet('Settings');
  const wsInst  = wb.addWorksheet('Instructions');

  console.log('Building Settings...');      buildSettings(wsSet);
  console.log('Building Bid Log...');       buildBidLog(wsBid);
  console.log('Building Go-No-Go...');      buildGonogo(wsGng);
  console.log('Building Dashboard...');     buildDashboard(wsDash);
  console.log('Building Friday Meeting...'); buildFriday(wsFri);
  console.log('Building Follow-Up...');     buildFollowup(wsFu);
  console.log('Building Data Quality...'); buildDataQuality(wsDq);
  console.log('Building Instructions...'); buildInstructions(wsInst);

  // Set Dashboard as default active tab
  wsDash.state = 'visible';
  wb.views = [{ activeTab: 0 }];

  console.log('Saving workbook...');
  await wb.xlsx.writeFile(OUTPUT);
  console.log(`✓ Saved: ${OUTPUT}`);
}

main().catch(err => { console.error('ERROR:', err.message); process.exit(1); });

