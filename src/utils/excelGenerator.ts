import ExcelJS from 'exceljs';
import type { TimesheetData } from '../types';
import { getMonthName, isWeekend, calculateHours, timeToMinutes } from './dateUtils';

// ── Row structure matching F012 exactly ──────────────────────────────────
interface TimeBlock {
  startMin: number;
  endMin: number;
  label: string;
  excelRows: number[];
  isLunch: boolean;
}

// Grid rows 4-47 all have height 14.1pt (matching F012 Base template)
const TIME_BLOCKS: TimeBlock[] = [
  { startMin: 420,  endMin: 450,  label: '7',     excelRows: [4],         isLunch: false },
  { startMin: 450,  endMin: 480,  label: '7.30',  excelRows: [5, 6],      isLunch: false },
  { startMin: 480,  endMin: 510,  label: '8',     excelRows: [7, 8],      isLunch: false },
  { startMin: 510,  endMin: 540,  label: '8.30',  excelRows: [9, 10],     isLunch: false },
  { startMin: 540,  endMin: 570,  label: '9',     excelRows: [11, 12],    isLunch: false },
  { startMin: 570,  endMin: 600,  label: '9.30',  excelRows: [13, 14],    isLunch: false },
  { startMin: 600,  endMin: 630,  label: '10',    excelRows: [15, 16],    isLunch: false },
  { startMin: 630,  endMin: 660,  label: '10.30', excelRows: [17, 18],    isLunch: false },
  { startMin: 660,  endMin: 690,  label: '11',    excelRows: [19, 20],    isLunch: false },
  { startMin: 690,  endMin: 720,  label: '11.30', excelRows: [21, 22],    isLunch: false },
  { startMin: 720,  endMin: 780,  label: '12\n12.15\n13', excelRows: [23, 24, 25, 26], isLunch: true },
  { startMin: 810,  endMin: 840,  label: '13.30', excelRows: [27, 28],    isLunch: false },
  { startMin: 840,  endMin: 870,  label: '14',    excelRows: [29, 30],    isLunch: false },
  { startMin: 870,  endMin: 900,  label: '14.30', excelRows: [31, 32],    isLunch: false },
  { startMin: 900,  endMin: 930,  label: '15',    excelRows: [33, 34],    isLunch: false },
  { startMin: 930,  endMin: 960,  label: '15.30', excelRows: [35, 36],    isLunch: false },
  { startMin: 960,  endMin: 990,  label: '16',    excelRows: [37, 38],    isLunch: false },
  { startMin: 990,  endMin: 1020, label: '16.30', excelRows: [39, 40],    isLunch: false },
  { startMin: 1020, endMin: 1050, label: '17',    excelRows: [41, 42],    isLunch: false },
  { startMin: 1050, endMin: 1080, label: '17.30', excelRows: [43, 44],    isLunch: false },
  { startMin: 1080, endMin: 1110, label: '18',    excelRows: [45, 46],    isLunch: false },
  { startMin: 1110, endMin: 1110, label: '18.30', excelRows: [47],        isLunch: false },
];

// ── Column mapping (1-indexed, matching F012 exactly) ────────────────────
// Col  1 (A):     time label "H"
// Cols 2-31:      days 1-15, 2 cols each  (day d → cols 2d, 2d+1)
// Col 32 (AF):    second "H" separator
// Cols 33-64:     days 16-31, 2 cols each (day d → cols 32+2(d-15)-1, 32+2(d-15))
// Col 65 (BM):    "Temps bloqué" markers + Réel/Dû totals
// Col 66 (BN):    merged with BM for Réel/Dû; start of Remarques header
// Cols 66-72 (BN-BT): Remarques

function dayCol1(d: number): number {
  // Returns 1-indexed first column for this day
  if (d <= 15) return 2 * d;
  return 31 + 2 * (d - 15);  // d=16 → 33, d=31 → 63
}

// ── Fills ─────────────────────────────────────────────────────────────────
const WORKED_FILL = {
  type: 'pattern' as const,
  pattern: 'solid' as const,
  fgColor: { theme: 0 as const, tint: -0.249977111117893 },
};
const RED_FILL   = { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: 'FFFF0000' } };
const MGRAY_FILL = { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: 'FFBfBFBF' } };
const VACN_FILL  = { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: 'FFA8CCE8' } };
const HOLI_FILL  = { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: 'FFFFD966' } };
// "Temps bloqué" uses same theme color as worked cells
const TBLK_FILL  = {
  type: 'pattern' as const,
  pattern: 'solid' as const,
  fgColor: { theme: 0 as const, tint: -0.3499862666707358 },
};

// ── Borders ───────────────────────────────────────────────────────────────
const MED  = { style: 'medium' as const };
const THIN = { style: 'thin'   as const };
const HAIR = { style: 'hair'   as const };

function medBox()  { return { top: MED,  left: MED,  bottom: MED,  right: MED  }; }
function medLR()   { return { left: MED, right: MED }; }

// ── Font helpers ──────────────────────────────────────────────────────────
function arial(size: number, bold = false, color?: { argb: string }): Partial<ExcelJS.Font> {
  const f: Partial<ExcelJS.Font> = { name: 'Arial', size, bold };
  if (color) f.color = color;
  return f;
}

// ── Status of a time block for a day ─────────────────────────────────────
function blockStatus(
  block: TimeBlock,
  mStart: number | null, mEnd: number | null,
  aStart: number | null, aEnd: number | null
): 'worked' | 'lunch' | 'empty' {
  if (block.isLunch) return 'lunch';
  const mid = (block.startMin + block.endMin) / 2;
  if (mStart !== null && mEnd !== null && mid >= mStart && mid < mEnd) return 'worked';
  if (aStart !== null && aEnd !== null && mid >= aStart && mid < aEnd) return 'worked';
  if (mEnd !== null && aStart !== null && mid >= mEnd && mid < aStart) return 'lunch';
  return 'empty';
}

// ── Main generator ────────────────────────────────────────────────────────
export async function generateExcel(data: TimesheetData): Promise<void> {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'SRW App';
  wb.created = new Date();

  const sheetName = `${getMonthName(data.month)} ${data.year}`;
  const ws = wb.addWorksheet(sheetName, {
    pageSetup: {
      paperSize: 9,         // A4
      orientation: 'landscape',
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 1,
      margins: { left: 0.31, right: 0.31, top: 0.31, bottom: 0.31, header: 0, footer: 0 },
    },
  });

  const daysInMonth = new Date(data.year, data.month, 0).getDate();

  // ── Column widths (exact F012 values) ────────────────────────────────
  ws.getColumn(1).width = 5.5;                              // A: time "H"
  for (let c = 2; c <= 31; c++) ws.getColumn(c).width = 1.5;  // B-AE: days 1-15
  ws.getColumn(32).width = 5.875;                           // AF: H separator
  for (let c = 33; c <= 64; c++) ws.getColumn(c).width = 1.5; // AG-BL: days 16-31
  ws.getColumn(65).width = 5.75;                            // BM: markers
  ws.getColumn(66).width = 5.375;                           // BN: Remarques start
  ws.getColumn(67).width = 3.875;                           // BO
  for (let c = 68; c <= 71; c++) ws.getColumn(c).width = 5;
  ws.getColumn(72).width = 17;                              // BT

  // ── Row heights (exact F012 values) ──────────────────────────────────
  ws.getRow(1).height = 15.75;
  ws.getRow(2).height = 20.1;
  ws.getRow(3).height = 17.25;
  for (let r = 4; r <= 47; r++) ws.getRow(r).height = 14.1;
  ws.getRow(48).height = 9.95;
  ws.getRow(49).height = 9.95;
  ws.getRow(50).height = 19.5;
  ws.getRow(51).height = 61.5;
  ws.getRow(52).height = 61.5;
  ws.getRow(53).height = 4.5;
  ws.getRow(54).height = 30;
  ws.getRow(55).height = 32.25;

  // ── Row 1: Company + logo ─────────────────────────────────────────────
  ws.mergeCells(1, 1, 1, 24); // A1:X1
  const r1c1 = ws.getCell(1, 1);
  r1c1.value = 'INGENIEURS-CONSEILS SCHERLER SA';
  r1c1.font = arial(12, true);
  r1c1.alignment = { vertical: 'middle', horizontal: 'left' };

  const r1cAG = ws.getCell(1, 33); // AG1
  r1cAG.value = 'srg | engineering';
  r1cAG.font = { ...arial(8), italic: true };
  r1cAG.alignment = { vertical: 'middle', horizontal: 'right' };

  // ── Row 2: CONTRÔLE DES HEURES — NOM — MOIS — F012 ───────────────────
  ws.mergeCells(2, 1, 2, 16);  // A2:P2 — "CONTRÔLE DES HEURES"
  const r2ctrl = ws.getCell(2, 1);
  r2ctrl.value = 'CONTRÔLE DES HEURES';
  r2ctrl.font = arial(10, true);
  r2ctrl.alignment = { vertical: 'middle', horizontal: 'left' };

  ws.mergeCells(2, 33, 2, 63); // AG2:BK2 — NOM
  const r2nom = ws.getCell(2, 33);
  r2nom.value = `NOM :   ${data.employeeName}`;
  r2nom.font = arial(10, true);
  r2nom.alignment = { vertical: 'middle', horizontal: 'left' };

  ws.mergeCells(2, 67, 2, 70); // BO2:BR2 — MOIS
  const r2mois = ws.getCell(2, 67);
  r2mois.value = `MOIS :  ${getMonthName(data.month)} ${data.year}`;
  r2mois.font = arial(10, true);
  r2mois.alignment = { vertical: 'middle', horizontal: 'left' };

  const r2f012 = ws.getCell(2, 72); // BT2 — F012
  r2f012.value = data.reference;
  r2f012.font = arial(10, true);
  r2f012.alignment = { vertical: 'middle', horizontal: 'right' };

  // ── Row 3: Day headers ────────────────────────────────────────────────
  const r3A = ws.getCell(3, 1);
  r3A.value = 'H';
  r3A.font = arial(8, true);
  r3A.alignment = { vertical: 'middle', horizontal: 'center' };
  r3A.border = medBox();

  for (let d = 1; d <= daysInMonth; d++) {
    const col1 = dayCol1(d);
    const date = new Date(data.year, data.month - 1, d);
    const wknd = isWeekend(date);
    ws.mergeCells(3, col1, 3, col1 + 1);
    const cell = ws.getCell(3, col1);
    cell.value = d;
    cell.font = arial(7, wknd, wknd ? { argb: 'FFFFFFFF' } : undefined);
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.fill = wknd ? RED_FILL : { type: 'pattern', pattern: 'none' };
    cell.border = {
      left: d === 1 ? MED : THIN,
      top: MED,
      bottom: MED,
    };
  }

  const r3AF = ws.getCell(3, 32); // AF3
  r3AF.value = 'H';
  r3AF.font = arial(8, true);
  r3AF.alignment = { vertical: 'middle', horizontal: 'center' };
  r3AF.border = medBox();

  ws.mergeCells(3, 66, 3, 72); // BN3:BT3 — Remarques
  const r3rem = ws.getCell(3, 66);
  r3rem.value = 'Remarques';
  r3rem.font = arial(8, true);
  r3rem.alignment = { vertical: 'middle', horizontal: 'center' };
  r3rem.border = { left: MED, top: MED, bottom: MED };

  // ── Time grid (rows 4-47) ─────────────────────────────────────────────
  for (const block of TIME_BLOCKS) {
    const firstRow = block.excelRows[0];
    const lastRow  = block.excelRows[block.excelRows.length - 1];

    // Time label: col A
    if (firstRow < lastRow) ws.mergeCells(firstRow, 1, lastRow, 1);
    const labelA = ws.getCell(firstRow, 1);
    labelA.value = block.label;
    labelA.font = arial(7);
    labelA.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    labelA.border = medLR();

    // Mirror label: col AF (32)
    if (firstRow < lastRow) ws.mergeCells(firstRow, 32, lastRow, 32);
    const labelAF = ws.getCell(firstRow, 32);
    labelAF.value = block.label;
    labelAF.font = arial(7);
    labelAF.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    labelAF.border = medLR();

    // Day columns
    for (let d = 1; d <= daysInMonth; d++) {
      const col1 = dayCol1(d);
      const date = new Date(data.year, data.month - 1, d);
      const wknd = isWeekend(date);
      const dateStr = `${data.year}-${String(data.month).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
      const entry = data.entries[dateStr];

      ws.mergeCells(firstRow, col1, lastRow, col1 + 1);
      const cell = ws.getCell(firstRow, col1);
      cell.border = {
        bottom: HAIR,
        ...(d === daysInMonth ? { right: MED } : {}),
      };

      if (wknd) {
        cell.fill = RED_FILL;
      } else if (entry?.isVacation) {
        cell.fill = VACN_FILL;
      } else if (entry?.isHoliday) {
        cell.fill = HOLI_FILL;
      } else if (block.isLunch) {
        cell.fill = MGRAY_FILL;
      } else if (entry) {
        const mStart = entry.morning ? timeToMinutes(entry.morning.start) : null;
        const mEnd   = entry.morning ? timeToMinutes(entry.morning.end)   : null;
        const aStart = entry.afternoon ? timeToMinutes(entry.afternoon.start) : null;
        const aEnd   = entry.afternoon ? timeToMinutes(entry.afternoon.end)   : null;

        if (mStart === null && aStart === null) {
          // no time entered
        } else {
          const status = blockStatus(block, mStart, mEnd, aStart, aEnd);
          if (status === 'worked') cell.fill = WORKED_FILL;
          else if (status === 'lunch') cell.fill = MGRAY_FILL;
        }
      }
    }

    // Top border on first grid row
    if (firstRow === 4) {
      for (let d = 1; d <= daysInMonth; d++) {
        const col1 = dayCol1(d);
        const cell = ws.getCell(4, col1);
        cell.border = { ...(cell.border as object || {}), top: MED, bottom: HAIR };
      }
    }
  }

  // ── BM column (col 65): Temps bloqué + borders ────────────────────────
  // Medium left/right borders on all grid rows
  for (let r = 4; r <= 47; r++) {
    const cell = ws.getCell(r, 65);
    const existing = (cell.border as object) || {};
    cell.border = { ...existing, left: MED, right: MED };
    if (r === 4)  cell.border = { ...cell.border, top: MED };
    if (r === 47) cell.border = { ...cell.border, bottom: MED };
  }

  // Time markers (representing "Temps bloqué" contracted hours windows)
  // Morning block: 8:30–11:45, Afternoon block: 14:00–16:30
  ws.getCell(10, 65).value = 8.3;   // end of 8:30 slot
  ws.getCell(10, 65).fill  = TBLK_FILL;
  ws.getCell(10, 65).font  = arial(6);
  ws.getCell(10, 65).alignment = { vertical: 'bottom', horizontal: 'right' };
  ws.getCell(10, 65).border = { left: MED, right: MED, bottom: MED };

  ws.mergeCells(11, 65, 21, 65); // 9:00–11:30 mandatory morning
  const bmMorn = ws.getCell(11, 65);
  bmMorn.value = 'Temps bloqué';
  bmMorn.fill  = TBLK_FILL;
  bmMorn.font  = { ...arial(6), color: { argb: 'FFFFFFFF' } };
  bmMorn.alignment = { vertical: 'middle', horizontal: 'center', textRotation: 90 };
  bmMorn.border = { left: MED, right: MED };

  ws.getCell(22, 65).value = 11.45;
  ws.getCell(22, 65).fill  = TBLK_FILL;
  ws.getCell(22, 65).font  = arial(6);
  ws.getCell(22, 65).alignment = { vertical: 'top', horizontal: 'right' };
  ws.getCell(22, 65).border = { left: MED, right: MED, top: MED };

  ws.getCell(30, 65).value = 14;
  ws.getCell(30, 65).fill  = TBLK_FILL;
  ws.getCell(30, 65).font  = arial(6);
  ws.getCell(30, 65).alignment = { vertical: 'bottom', horizontal: 'right' };
  ws.getCell(30, 65).border = { left: MED, right: MED, bottom: MED };

  ws.mergeCells(31, 65, 38, 65); // 14:30–16:00 mandatory afternoon
  const bmAftn = ws.getCell(31, 65);
  bmAftn.value = 'Temps bloqué';
  bmAftn.fill  = TBLK_FILL;
  bmAftn.font  = { ...arial(6), color: { argb: 'FFFFFFFF' } };
  bmAftn.alignment = { vertical: 'middle', horizontal: 'center', textRotation: 90 };
  bmAftn.border = { left: MED, right: MED };

  ws.getCell(39, 65).value = 16.3;
  ws.getCell(39, 65).fill  = TBLK_FILL;
  ws.getCell(39, 65).font  = arial(6);
  ws.getCell(39, 65).alignment = { vertical: 'top', horizontal: 'right' };
  ws.getCell(39, 65).border = { left: MED, right: MED, top: MED };

  // ── Rows 48-53: Presence / signature area ─────────────────────────────
  // (leave content empty — user fills in manually; just set BM borders)
  for (let r = 48; r <= 53; r++) {
    ws.getCell(r, 65).border = { right: THIN };
  }

  // ── Row 54: Daily actual hours ────────────────────────────────────────
  let totalReal = 0;
  let totalDu   = 0;

  for (let d = 1; d <= daysInMonth; d++) {
    const col1 = dayCol1(d);
    const date = new Date(data.year, data.month - 1, d);
    const wknd = isWeekend(date);
    const dateStr = `${data.year}-${String(data.month).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
    const entry = data.entries[dateStr];

    ws.mergeCells(54, col1, 54, col1 + 1);
    const cell = ws.getCell(54, col1);
    cell.font = arial(7);
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border = { left: THIN, top: MED, bottom: THIN };

    if (wknd) {
      cell.value = 'X';
      cell.fill = RED_FILL;
      cell.font = { ...arial(7), color: { argb: 'FFFFFFFF' } };
    } else if (entry) {
      const mH = calculateHours(entry.morning);
      const aH = calculateHours(entry.afternoon);
      const total = mH + aH;
      if (total > 0) {
        cell.value = total;
        totalReal += total;
        totalDu   += entry.contractedHours;
      }
    }
  }

  // BM54:BN54 — Réel total
  ws.mergeCells(54, 65, 54, 66);
  const realCell = ws.getCell(54, 65);
  realCell.value = `Réel: ${totalReal.toFixed(2)}`;
  realCell.font = arial(8, true);
  realCell.alignment = { vertical: 'middle', horizontal: 'left' };
  realCell.border = { left: MED, top: MED };

  // ── Row 55: Contracted hours ──────────────────────────────────────────
  for (let d = 1; d <= daysInMonth; d++) {
    const col1 = dayCol1(d);
    const date = new Date(data.year, data.month - 1, d);
    const wknd = isWeekend(date);
    const dateStr = `${data.year}-${String(data.month).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
    const entry = data.entries[dateStr];

    ws.mergeCells(55, col1, 55, col1 + 1);
    const cell = ws.getCell(55, col1);
    cell.font = arial(7);
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border = { left: THIN, bottom: MED };

    if (wknd) {
      cell.value = 'X';
      cell.fill = RED_FILL;
      cell.font = { ...arial(7), color: { argb: 'FFFFFFFF' } };
    } else if (entry) {
      cell.value = entry.contractedHours;
    }
  }

  // BM55:BN55 — Dû total
  ws.mergeCells(55, 65, 55, 66);
  const duCell = ws.getCell(55, 65);
  duCell.value = `Dû: ${totalDu.toFixed(2)}`;
  duCell.font = arial(8, true);
  duCell.alignment = { vertical: 'middle', horizontal: 'left' };
  duCell.border = { left: MED, bottom: MED };

  // ── Download ──────────────────────────────────────────────────────────
  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `CDH_${data.employeeName.replace(/\s+/g, '_')}_${getMonthName(data.month)}_${data.year}.xlsx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
