import ExcelJS from 'exceljs';
import type { TimesheetData } from '../types';
import { getMonthName, isWeekend, calculateHours, timeToMinutes } from './dateUtils';

// ── Row structure matching F012 exactly ──────────────────────────────────
// Each entry: [startMin, endMin, excelRows]
// The grid rows 4-47 all have height 14.1pt
interface TimeBlock {
  startMin: number;
  endMin: number;
  label: string;
  excelRows: number[]; // 1-indexed Excel row numbers covered by this block
  isLunch: boolean;
}

// Row 4 = 7:00 (1 row), Row 5-6 = 7:30, Row 7-8 = 8:00 ... Row 23-26 = lunch, Row 27-28 = 13:30 ...
const TIME_BLOCKS: TimeBlock[] = [
  { startMin: 420,  endMin: 450,  label: '7',              excelRows: [4],               isLunch: false },
  { startMin: 450,  endMin: 480,  label: '7.30',           excelRows: [5, 6],             isLunch: false },
  { startMin: 480,  endMin: 510,  label: '8',              excelRows: [7, 8],             isLunch: false },
  { startMin: 510,  endMin: 540,  label: '8.30',           excelRows: [9, 10],            isLunch: false },
  { startMin: 540,  endMin: 570,  label: '9',              excelRows: [11, 12],           isLunch: false },
  { startMin: 570,  endMin: 600,  label: '9.30',           excelRows: [13, 14],           isLunch: false },
  { startMin: 600,  endMin: 630,  label: '10',             excelRows: [15, 16],           isLunch: false },
  { startMin: 630,  endMin: 660,  label: '10.30',          excelRows: [17, 18],           isLunch: false },
  { startMin: 660,  endMin: 690,  label: '11',             excelRows: [19, 20],           isLunch: false },
  { startMin: 690,  endMin: 720,  label: '11.30',          excelRows: [21, 22],           isLunch: false },
  { startMin: 720,  endMin: 780,  label: '12.00\n12.15\n13.00', excelRows: [23, 24, 25, 26], isLunch: true },
  { startMin: 810,  endMin: 840,  label: '13.30',          excelRows: [27, 28],           isLunch: false },
  { startMin: 840,  endMin: 870,  label: '14',             excelRows: [29, 30],           isLunch: false },
  { startMin: 870,  endMin: 900,  label: '14.30',          excelRows: [31, 32],           isLunch: false },
  { startMin: 900,  endMin: 930,  label: '15',             excelRows: [33, 34],           isLunch: false },
  { startMin: 930,  endMin: 960,  label: '15.30',          excelRows: [35, 36],           isLunch: false },
  { startMin: 960,  endMin: 990,  label: '16',             excelRows: [37, 38],           isLunch: false },
  { startMin: 990,  endMin: 1020, label: '16.30',          excelRows: [39, 40],           isLunch: false },
  { startMin: 1020, endMin: 1050, label: '17',             excelRows: [41, 42],           isLunch: false },
  { startMin: 1050, endMin: 1080, label: '17.30',          excelRows: [43, 44],           isLunch: false },
  { startMin: 1080, endMin: 1110, label: '18',             excelRows: [45, 46],           isLunch: false },
  { startMin: 1110, endMin: 1110, label: '18.30',          excelRows: [47],               isLunch: false },
];

// ── Column mapping ────────────────────────────────────────────────────────
// Col A (idx 0): time labels
// Cols B–AE (idx 1–30): days 1–15, each uses 2 columns
// Col AF (idx 31): second H label
// Cols AG–BL (idx 32–63): days 16–31, each uses 2 columns
// Col BM (idx 64): markers
// Col BN (idx 65): day number for remarks
// Col BO–BV (idx 66–73): remarks text

function dayColIdx(d: number): number {
  if (d <= 15) return 1 + (d - 1) * 2;   // B=1, D=3, F=5... AE=29 (0-indexed)
  return 32 + (d - 16) * 2;              // AG=32, AI=34... BL=62 (0-indexed)
}

// ── Colors ────────────────────────────────────────────────────────────────
const RED    = { argb: 'FFFF0000' };      // Weekend
const BLACK  = { argb: 'FF000000' };      // Worked
const LGRAY  = { argb: 'FFD9D9D9' };      // Header background
const MGRAY  = { argb: 'FFBfBFBF' };      // Lunch block
const WHITE  = { argb: 'FFFFFFFF' };
const GREEN  = { argb: 'FFE2EFDA' };      // Réel total
const BLUE   = { argb: 'FFDCE6F1' };      // Dû total

function solidFill(color: { argb: string }) {
  return { type: 'pattern' as const, pattern: 'solid' as const, fgColor: color };
}

function thinBorder() {
  const side = { style: 'thin' as const, color: { argb: 'FF808080' } };
  return { top: side, left: side, bottom: side, right: side };
}

function cellFont(size = 7, bold = false) {
  return { name: 'Calibri', size, bold };
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
      paperSize: 9,        // A4
      orientation: 'landscape',
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 1,
      margins: { left: 0.31, right: 0.31, top: 0.31, bottom: 0.31, header: 0, footer: 0 },
    },
  });

  const daysInMonth = new Date(data.year, data.month, 0).getDate();

  // ── Column widths (matching F012 exactly) ─────────────────────────────
  ws.getColumn(1).width = 5.5;    // A: time label
  for (let i = 2; i <= 31; i++) ws.getColumn(i).width = 1.5;  // B–AE day cols 1-15
  ws.getColumn(32).width = 5.875; // AF: H separator
  for (let i = 33; i <= 64; i++) ws.getColumn(i).width = 1.5; // AG–BL day cols 16-31
  ws.getColumn(65).width = 5.75;  // BM: markers
  ws.getColumn(66).width = 5.375; // BN: remark day #
  ws.getColumn(67).width = 3.875; // BO: remark text start
  for (let i = 68; i <= 74; i++) ws.getColumn(i).width = 5;   // rest of remarks

  // Merge remarks text cols BO–BV for a wide text cell
  // (done per-row below)

  // ── Row heights ────────────────────────────────────────────────────────
  ws.getRow(1).height = 15.75;
  ws.getRow(2).height = 20.1;
  ws.getRow(3).height = 17.25;
  for (let r = 4; r <= 47; r++) ws.getRow(r).height = 14.1;
  ws.getRow(48).height = 9.95;
  ws.getRow(49).height = 9.95;
  ws.getRow(54).height = 30;
  ws.getRow(55).height = 32.25;

  // ── Row 1: Company name ───────────────────────────────────────────────
  const r1 = ws.getRow(1);
  const c1a = r1.getCell(1);
  c1a.value = 'INGENIEURS-CONSEILS SCHERLER SA';
  c1a.font = cellFont(9, true);
  c1a.alignment = { vertical: 'middle', horizontal: 'left' };
  // Merge A1 to end
  ws.mergeCells(1, 1, 1, 66);
  const c1b = r1.getCell(67);
  c1b.value = 'srg | engineering';
  c1b.font = { ...cellFont(8), italic: true };
  c1b.alignment = { vertical: 'middle', horizontal: 'right' };
  ws.mergeCells(1, 67, 1, 74);

  // ── Row 2: CONTRÔLE DES HEURES header ─────────────────────────────────
  const r2 = ws.getRow(2);
  const hdrStyle = { font: cellFont(10, true), alignment: { vertical: 'middle' as const } };

  ws.mergeCells(2, 1, 2, 10);
  const c2a = r2.getCell(1);
  c2a.value = 'CONTRÔLE DES HEURES';
  Object.assign(c2a, hdrStyle);
  c2a.alignment = { vertical: 'middle', horizontal: 'left' };

  ws.mergeCells(2, 11, 2, 31);
  const c2b = r2.getCell(11);
  c2b.value = `NOM :   ${data.employeeName}`;
  Object.assign(c2b, hdrStyle);

  ws.mergeCells(2, 32, 2, 58);
  const c2c = r2.getCell(32);
  c2c.value = `MOIS :  ${getMonthName(data.month)} ${data.year}`;
  Object.assign(c2c, hdrStyle);

  ws.mergeCells(2, 59, 2, 74);
  const c2d = r2.getCell(59);
  c2d.value = data.reference;
  Object.assign(c2d, hdrStyle);
  c2d.alignment = { vertical: 'middle', horizontal: 'right' };

  // ── Row 3: Day numbers ────────────────────────────────────────────────
  const r3 = ws.getRow(3);

  // H label (col A)
  ws.mergeCells(3, 1, 3, 1);
  const c3h1 = r3.getCell(1);
  c3h1.value = 'H';
  c3h1.font = cellFont(8, true);
  c3h1.alignment = { vertical: 'middle', horizontal: 'center' };
  c3h1.fill = solidFill(LGRAY);
  c3h1.border = thinBorder();

  for (let d = 1; d <= daysInMonth; d++) {
    const colIdx = dayColIdx(d);
    const date = new Date(data.year, data.month - 1, d);
    const wknd = isWeekend(date);

    // Merge the 2 cols for this day (except day 1 which uses B only for 31-day months)
    const colsForDay = d === 1 || d === 16 ? 2 : 2;
    ws.mergeCells(3, colIdx + 1, 3, colIdx + colsForDay);
    const cell = r3.getCell(colIdx + 1);
    cell.value = d;
    cell.font = cellFont(7, wknd);
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.fill = solidFill(wknd ? RED : LGRAY);
    cell.border = thinBorder();
    if (wknd) cell.font = { ...cellFont(7, true), color: WHITE };
  }

  // H separator col AF (idx 31 = col 32)
  ws.mergeCells(3, 32, 3, 32);
  const c3h2 = r3.getCell(32);
  c3h2.value = 'H';
  c3h2.font = cellFont(8, true);
  c3h2.alignment = { vertical: 'middle', horizontal: 'center' };
  c3h2.fill = solidFill(LGRAY);
  c3h2.border = thinBorder();

  // Remarques header (cols BN-BV = 66-74)
  ws.mergeCells(3, 66, 3, 74);
  const c3r = r3.getCell(66);
  c3r.value = 'Remarques';
  c3r.font = cellFont(8, true);
  c3r.alignment = { vertical: 'middle', horizontal: 'center' };
  c3r.fill = solidFill(LGRAY);
  c3r.border = thinBorder();

  // ── Time grid (rows 4-47) ─────────────────────────────────────────────
  for (const block of TIME_BLOCKS) {
    const firstRow = block.excelRows[0];
    const lastRow  = block.excelRows[block.excelRows.length - 1];

    // Time label (col A), merge all rows in block
    ws.mergeCells(firstRow, 1, lastRow, 1);
    const labelCell = ws.getCell(firstRow, 1);
    labelCell.value = block.label;
    labelCell.font = cellFont(6.5);
    labelCell.alignment = { vertical: 'middle', wrapText: true };
    labelCell.fill = solidFill(LGRAY);
    labelCell.border = thinBorder();

    // Mirror label in col AF
    ws.mergeCells(firstRow, 32, lastRow, 32);
    const labelCell2 = ws.getCell(firstRow, 32);
    labelCell2.value = block.label;
    labelCell2.font = cellFont(6.5);
    labelCell2.alignment = { vertical: 'middle', wrapText: true };
    labelCell2.fill = solidFill(LGRAY);
    labelCell2.border = thinBorder();

    // Day columns
    for (let d = 1; d <= daysInMonth; d++) {
      const colIdx = dayColIdx(d) + 1; // 1-indexed column
      const date = new Date(data.year, data.month - 1, d);
      const wknd = isWeekend(date);
      const dateStr = `${data.year}-${String(data.month).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
      const entry = data.entries[dateStr];

      let fillColor = WHITE;

      if (wknd) {
        fillColor = RED;
      } else if (entry?.isVacation) {
        fillColor = { argb: 'FFA8CCE8' };
      } else if (entry?.isHoliday) {
        fillColor = { argb: 'FFFFD966' };
      } else if (entry && !block.isLunch) {
        const mStart = entry.morning ? timeToMinutes(entry.morning.start) : null;
        const mEnd   = entry.morning ? timeToMinutes(entry.morning.end)   : null;
        const aStart = entry.afternoon ? timeToMinutes(entry.afternoon.start) : null;
        const aEnd   = entry.afternoon ? timeToMinutes(entry.afternoon.end)   : null;

        const status = blockStatus(block, mStart, mEnd, aStart, aEnd);
        if (status === 'worked') fillColor = BLACK;
        else if (status === 'lunch') fillColor = MGRAY;
      } else if (block.isLunch && entry && !entry.isVacation && !entry.isHoliday) {
        fillColor = MGRAY;
      }

      // Merge the 2 sub-columns for this day across all rows of this block
      ws.mergeCells(firstRow, colIdx, lastRow, colIdx + 1);
      const cell = ws.getCell(firstRow, colIdx);
      cell.fill = solidFill(fillColor);
      cell.border = thinBorder();
    }
  }

  // ── Remarks column (BN-BV, rows 4-47) per day ─────────────────────────
  // Each day gets a proportional slice of the grid height
  // We'll use 1 row per day for remarks (keep it simple)
  // Since we have 44 rows for up to 31 days, give each day ~1.4 rows
  // For simplicity: use the row nearest to day's proportional position

  const gridRows = 44; // rows 4-47
  for (let d = 1; d <= daysInMonth; d++) {
    const relRow = Math.round(((d - 1) / daysInMonth) * gridRows);
    const excelRow = 4 + relRow;

    const date = new Date(data.year, data.month - 1, d);
    const wknd = isWeekend(date);
    const dateStr = `${data.year}-${String(data.month).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
    const entry = data.entries[dateStr];

    // Day number cell (BN = col 66)
    const numCell = ws.getCell(excelRow, 66);
    numCell.value = d;
    numCell.font = { ...cellFont(6, true), color: wknd ? RED : BLACK };
    numCell.alignment = { vertical: 'middle', horizontal: 'center' };

    // Remark text (BO-BV = cols 67-74)
    ws.mergeCells(excelRow, 67, excelRow, 74);
    const txtCell = ws.getCell(excelRow, 67);
    txtCell.font = cellFont(6);
    txtCell.alignment = { vertical: 'middle', wrapText: false };

    if (!wknd && entry) {
      const mH = calculateHours(entry.morning);
      const aH = calculateHours(entry.afternoon);
      let text = '';

      if (entry.isVacation) text = 'Vacances';
      else if (entry.isHoliday) text = 'Férié';
      else {
        const parts: string[] = [];
        if (entry.morningWorksiteId && mH > 0) {
          const ws2 = data.worksites.find(w => w.id === entry.morningWorksiteId);
          if (ws2) parts.push(`${ws2.number} ${mH}h`);
        }
        if (entry.afternoonWorksiteId && aH > 0) {
          const ws2 = data.worksites.find(w => w.id === entry.afternoonWorksiteId);
          if (ws2) parts.push(`${ws2.number} ${aH}h`);
        }
        text = parts.join(' / ');
        if (entry.remark) text += text ? ` (${entry.remark})` : entry.remark;
      }
      txtCell.value = text;
    }
  }

  // ── Row 54: Daily totals ───────────────────────────────────────────────
  let totalReal = 0;
  let totalDu   = 0;
  const wsHours: Record<string, number> = {};

  for (let d = 1; d <= daysInMonth; d++) {
    const colIdx = dayColIdx(d) + 1;
    const date = new Date(data.year, data.month - 1, d);
    const wknd = isWeekend(date);
    const dateStr = `${data.year}-${String(data.month).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
    const entry = data.entries[dateStr];

    ws.mergeCells(54, colIdx, 54, colIdx + 1);
    const cell = ws.getCell(54, colIdx);
    cell.font = cellFont(7);
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border = thinBorder();

    if (wknd) {
      cell.fill = solidFill(RED);
      cell.font = { ...cellFont(7), color: WHITE };
      cell.value = 'X';
    } else if (entry) {
      const mH = calculateHours(entry.morning);
      const aH = calculateHours(entry.afternoon);
      const total = mH + aH;
      if (total > 0) {
        cell.value = total;
        totalReal += total;
        totalDu   += entry.contractedHours;
        if (entry.morningWorksiteId) wsHours[entry.morningWorksiteId] = (wsHours[entry.morningWorksiteId] || 0) + mH;
        if (entry.afternoonWorksiteId) wsHours[entry.afternoonWorksiteId] = (wsHours[entry.afternoonWorksiteId] || 0) + aH;
      }
    }
  }

  // Réel label (BM col = 65)
  ws.mergeCells(54, 65, 54, 74);
  const realCell = ws.getCell(54, 65);
  realCell.value = `Réel: ${totalReal.toFixed(2)}`;
  realCell.font = cellFont(8, true);
  realCell.fill = solidFill(GREEN);
  realCell.alignment = { vertical: 'middle' };
  realCell.border = thinBorder();

  // ── Row 55: Contracted hours ───────────────────────────────────────────
  for (let d = 1; d <= daysInMonth; d++) {
    const colIdx = dayColIdx(d) + 1;
    const date = new Date(data.year, data.month - 1, d);
    const wknd = isWeekend(date);
    const dateStr = `${data.year}-${String(data.month).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
    const entry = data.entries[dateStr];

    ws.mergeCells(55, colIdx, 55, colIdx + 1);
    const cell = ws.getCell(55, colIdx);
    cell.font = cellFont(7);
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border = thinBorder();

    if (wknd) {
      cell.fill = solidFill(RED);
      cell.font = { ...cellFont(7), color: WHITE };
      cell.value = 'X';
    } else if (entry) {
      cell.value = entry.contractedHours;
    }
  }

  // Dû label
  ws.mergeCells(55, 65, 55, 74);
  const duCell = ws.getCell(55, 65);
  duCell.value = `Dû: ${totalDu.toFixed(2)}`;
  duCell.font = cellFont(8, true);
  duCell.fill = solidFill(BLUE);
  duCell.alignment = { vertical: 'middle' };
  duCell.border = thinBorder();

  // ── Rows 56-60: Per-chantier summary ──────────────────────────────────
  ws.getRow(56).height = 14;
  const sumCell = ws.getCell(56, 1);
  ws.mergeCells(56, 1, 56, 74);
  sumCell.font = cellFont(7);

  const parts = data.worksites
    .filter(s => (wsHours[s.id] || 0) > 0)
    .map(s => `${s.number} ${s.abbreviation}: ${(wsHours[s.id] || 0).toFixed(2)}h`)
    .join('     ');
  const overtime = totalReal - totalDu;
  sumCell.value = parts + `     Sup.: ${overtime >= 0 ? '+' : ''}${overtime.toFixed(2)}h`;

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
