import jsPDF from 'jspdf';
import type { TimesheetData } from '../types';
import { getMonthName, timeToMinutes, calculateHours, isWeekend } from './dateUtils';

// ── Time block definitions (matching the Excel grid exactly) ──────────────
// Each block: startMin, endMin, label, heightUnits (1 unit = slotH mm)
// Lunch block uses 4 units to match Excel rows 23-26

interface Block {
  label: string;
  startMin: number;
  endMin: number;
  units: number;      // relative height
  isLunch: boolean;
}

const BLOCKS: Block[] = [
  { label: '7.00',  startMin: 420,  endMin: 450,  units: 1, isLunch: false },
  { label: '7.30',  startMin: 450,  endMin: 480,  units: 2, isLunch: false },
  { label: '8.00',  startMin: 480,  endMin: 510,  units: 2, isLunch: false },
  { label: '8.30',  startMin: 510,  endMin: 540,  units: 2, isLunch: false },
  { label: '9.00',  startMin: 540,  endMin: 570,  units: 2, isLunch: false },
  { label: '9.30',  startMin: 570,  endMin: 600,  units: 2, isLunch: false },
  { label: '10.00', startMin: 600,  endMin: 630,  units: 2, isLunch: false },
  { label: '10.30', startMin: 630,  endMin: 660,  units: 2, isLunch: false },
  { label: '11.00', startMin: 660,  endMin: 690,  units: 2, isLunch: false },
  { label: '11.30', startMin: 690,  endMin: 720,  units: 2, isLunch: false },
  // Lunch block (12.00–13.00) compressed into 4 units
  { label: '12.00\n12.15\n13.00', startMin: 720, endMin: 780, units: 4, isLunch: true },
  { label: '13.30', startMin: 810,  endMin: 840,  units: 2, isLunch: false },
  { label: '14.00', startMin: 840,  endMin: 870,  units: 2, isLunch: false },
  { label: '14.30', startMin: 870,  endMin: 900,  units: 2, isLunch: false },
  { label: '15.00', startMin: 900,  endMin: 930,  units: 2, isLunch: false },
  { label: '15.30', startMin: 930,  endMin: 960,  units: 2, isLunch: false },
  { label: '16.00', startMin: 960,  endMin: 990,  units: 2, isLunch: false },
  { label: '16.30', startMin: 990,  endMin: 1020, units: 2, isLunch: false },
  { label: '17.00', startMin: 1020, endMin: 1050, units: 2, isLunch: false },
  { label: '17.30', startMin: 1050, endMin: 1080, units: 2, isLunch: false },
  { label: '18.00', startMin: 1080, endMin: 1110, units: 2, isLunch: false },
  { label: '18.30', startMin: 1110, endMin: 1110, units: 1, isLunch: false }, // end marker only
];

const TOTAL_UNITS = BLOCKS.reduce((s, b) => s + b.units, 0); // 44 units

type RGB = [number, number, number];
const setFill = (doc: jsPDF, c: RGB) => doc.setFillColor(c[0], c[1], c[2]);
const setDraw = (doc: jsPDF, c: RGB, lw = 0.1) => {
  doc.setDrawColor(c[0], c[1], c[2]);
  doc.setLineWidth(lw);
};

function isWorked(block: Block, entry: { morning: { start: string; end: string } | null; afternoon: { start: string; end: string } | null } | undefined): 'worked' | 'lunch' | 'empty' {
  if (!entry) return 'empty';
  const { startMin, endMin } = block;
  const mid = (startMin + endMin) / 2;

  const mStart = entry.morning ? timeToMinutes(entry.morning.start) : null;
  const mEnd   = entry.morning ? timeToMinutes(entry.morning.end)   : null;
  const aStart = entry.afternoon ? timeToMinutes(entry.afternoon.start) : null;
  const aEnd   = entry.afternoon ? timeToMinutes(entry.afternoon.end)   : null;

  if (mStart !== null && mEnd !== null && mid >= mStart && mid <= mEnd) return 'worked';
  if (aStart !== null && aEnd !== null && mid >= aStart && mid <= aEnd) return 'worked';
  // Lunch gap
  if (mEnd !== null && aStart !== null && mid > mEnd && mid < aStart) return 'lunch';
  return 'empty';
}

export function generatePDF(data: TimesheetData): void {
  const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });

  const PW = 297;
  const PH = 210;
  const ML = 4;
  const MT = 4;

  // ── Layout ──────────────────────────────────────────────────
  const COMPANY_H  = 6;   // company name row
  const HEADER_H   = 8;   // CONTRÔLE DES HEURES row
  const DAYNUM_H   = 5;   // day numbers row
  const FOOTER1_H  = 5;   // totals row
  const FOOTER2_H  = 5;   // scheduled row
  const FOOTER3_H  = 5;   // Réel/Dû row
  const SUMMARY_H  = 22;  // per-chantier summary

  const gridTop    = MT + COMPANY_H + HEADER_H + DAYNUM_H;
  const gridAvailH = PH - MT - COMPANY_H - HEADER_H - DAYNUM_H - FOOTER1_H - FOOTER2_H - FOOTER3_H - SUMMARY_H - ML;
  const UNIT_H     = gridAvailH / TOTAL_UNITS;   // height of 1 unit
  const gridH      = TOTAL_UNITS * UNIT_H;
  const gridBot    = gridTop + gridH;

  const daysInMonth = new Date(data.year, data.month, 0).getDate();
  const TIME_COL_W  = 11;
  const RIGHT_W     = 72;
  const dayColW     = (PW - ML * 2 - TIME_COL_W - RIGHT_W) / daysInMonth;
  const gridW       = dayColW * daysInMonth;
  const RIGHT_X     = ML + TIME_COL_W + gridW + 1;

  // Pre-compute block Y positions
  const blockY: number[] = [];
  let yAcc = gridTop;
  for (const b of BLOCKS) {
    blockY.push(yAcc);
    yAcc += b.units * UNIT_H;
  }

  // ── Row 1: Company name ─────────────────────────────────────
  setFill(doc, [255, 255, 255]);
  setDraw(doc, [0, 0, 0], 0.4);
  doc.rect(ML, MT, PW - 2 * ML, COMPANY_H, 'FD');
  doc.setFontSize(7);
  doc.setFont('helvetica', 'bold');
  doc.setTextColor(0, 0, 0);
  doc.text('INGENIEURS-CONSEILS SCHERLER SA', ML + 3, MT + 4);
  doc.setFont('helvetica', 'normal');
  doc.text('srg | engineering', PW - ML - 3, MT + 4, { align: 'right' });

  // ── Row 2: CONTRÔLE DES HEURES header ──────────────────────
  const hY = MT + COMPANY_H;
  setFill(doc, [255, 255, 255]);
  setDraw(doc, [0, 0, 0], 0.4);
  doc.rect(ML, hY, PW - 2 * ML, HEADER_H, 'FD');
  doc.setFontSize(9);
  doc.setFont('helvetica', 'bold');
  doc.text('CONTRÔLE DES HEURES', ML + 3, hY + 5.5);
  doc.setFontSize(8);
  doc.text(`NOM :   ${data.employeeName}`, ML + 60, hY + 5.5);
  doc.text(`MOIS :  ${getMonthName(data.month)} ${data.year}`, ML + 140, hY + 5.5);
  doc.text(data.reference, PW - ML - 12, hY + 5.5);

  // ── Row 3: Day numbers ──────────────────────────────────────
  const dnY = hY + HEADER_H;

  // Time column
  setFill(doc, [210, 210, 210]);
  setDraw(doc, [0, 0, 0], 0.3);
  doc.rect(ML, dnY, TIME_COL_W, DAYNUM_H, 'FD');
  doc.setFontSize(6);
  doc.setFont('helvetica', 'bold');
  doc.setTextColor(0, 0, 0);
  doc.text('H', ML + TIME_COL_W / 2, dnY + 3.5, { align: 'center' });

  for (let d = 1; d <= daysInMonth; d++) {
    const x = ML + TIME_COL_W + (d - 1) * dayColW;
    const date = new Date(data.year, data.month - 1, d);
    const wknd = isWeekend(date);
    setFill(doc, wknd ? [200, 60, 60] : [210, 210, 210]);
    setDraw(doc, [0, 0, 0], 0.2);
    doc.rect(x, dnY, dayColW, DAYNUM_H, 'FD');
    doc.setFontSize(5);
    doc.setFont('helvetica', wknd ? 'bold' : 'normal');
    doc.setTextColor(wknd ? 255 : 0, wknd ? 255 : 0, wknd ? 255 : 0);
    doc.text(`${d}`, x + dayColW / 2, dnY + 3.5, { align: 'center' });
  }

  // Right-column header
  setFill(doc, [210, 210, 210]);
  setDraw(doc, [0, 0, 0], 0.3);
  doc.rect(RIGHT_X, dnY, RIGHT_W, DAYNUM_H, 'FD');
  doc.setFontSize(5.5);
  doc.setFont('helvetica', 'bold');
  doc.setTextColor(0, 0, 0);
  doc.text('Remarques', RIGHT_X + RIGHT_W / 2, dnY + 3.5, { align: 'center' });

  // ── Time grid ───────────────────────────────────────────────
  for (let bi = 0; bi < BLOCKS.length; bi++) {
    const block = BLOCKS[bi];
    const by = blockY[bi];
    const bh = block.units * UNIT_H;

    // Time label column
    if (block.isLunch) {
      setFill(doc, [230, 230, 230]);
      setDraw(doc, [80, 80, 80], 0.15);
      doc.rect(ML, by, TIME_COL_W, bh, 'FD');
      doc.setFontSize(4);
      doc.setFont('helvetica', 'normal');
      doc.setTextColor(80, 80, 80);
      // Three labels stacked
      const labels = ['12.00', '12.15', '13.00'];
      labels.forEach((lbl, li) => {
        doc.text(lbl, ML + 1, by + bh * (li + 0.5) / 3 + 0.5);
      });
    } else {
      const isHalfHour = bi % 2 === 0 || bi === 0;
      setFill(doc, [248, 248, 248]);
      setDraw(doc, isHalfHour ? [100, 100, 100] : [200, 200, 200], isHalfHour ? 0.2 : 0.1);
      doc.rect(ML, by, TIME_COL_W, bh, 'FD');
      if (block.startMin !== block.endMin) { // skip end marker for label
        doc.setFontSize(4.5);
        doc.setFont('helvetica', 'normal');
        doc.setTextColor(0, 0, 0);
        doc.text(block.label, ML + 1, by + bh * 0.6);
      }
    }

    // Day columns
    for (let d = 1; d <= daysInMonth; d++) {
      const x = ML + TIME_COL_W + (d - 1) * dayColW;
      const date = new Date(data.year, data.month - 1, d);
      const wknd = isWeekend(date);
      const dateStr = `${data.year}-${String(data.month).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
      const entry = data.entries[dateStr];

      let fill: RGB;
      if (wknd) {
        fill = [210, 70, 70];
      } else if (!entry || entry.isHoliday || entry.isVacation) {
        fill = entry?.isVacation ? [180, 210, 240] : entry?.isHoliday ? [255, 230, 150] : [255, 255, 255];
      } else if (block.isLunch) {
        fill = [190, 190, 190]; // gray lunch block for worked days
      } else {
        const state = isWorked(block, entry);
        if (state === 'worked') fill = [30, 30, 30];
        else if (state === 'lunch') fill = [180, 180, 180];
        else fill = [255, 255, 255];
      }

      setFill(doc, fill);
      setDraw(doc, [190, 190, 190], 0.05);
      doc.rect(x, by, dayColW, bh, 'FD');
    }
  }

  // ── Grid outer borders ──────────────────────────────────────
  setDraw(doc, [0, 0, 0], 0.35);
  doc.rect(ML, gridTop, TIME_COL_W, gridH);
  doc.rect(ML + TIME_COL_W, gridTop, gridW, gridH);

  // Horizontal lines at each block boundary (every 2 units = 30 min)
  setDraw(doc, [80, 80, 80], 0.2);
  let lineY = gridTop;
  for (const b of BLOCKS) {
    doc.line(ML, lineY, ML + TIME_COL_W + gridW, lineY);
    lineY += b.units * UNIT_H;
  }
  doc.line(ML, lineY, ML + TIME_COL_W + gridW, lineY); // bottom

  // ── Right column: "Temps bloqué" markers + day remarks ─────
  setDraw(doc, [0, 0, 0], 0.3);
  doc.rect(RIGHT_X, gridTop, RIGHT_W, gridH);

  // Temps bloqué zones (approximate: 8:30-11:45 and 14:00-16:30)
  const tBlock1Y = blockY[3]; // 8:30
  const tBlock1H = blockY[10] - blockY[3]; // to lunch
  const tBlock2Y = blockY[12]; // 14:00
  const tBlock2H = blockY[16] - blockY[12]; // to 16:30

  const markerX = RIGHT_X - 6;
  setFill(doc, [245, 245, 245]);
  setDraw(doc, [150, 150, 150], 0.15);
  doc.rect(markerX, tBlock1Y, 5, tBlock1H, 'FD');
  doc.rect(markerX, tBlock2Y, 5, tBlock2H, 'FD');

  // "Temps bloqué" rotated text
  doc.setFontSize(4);
  doc.setFont('helvetica', 'normal');
  doc.setTextColor(100, 100, 100);
  // Use small horizontal text as approximation
  doc.text('Temps bloqué', markerX + 2.5, tBlock1Y + tBlock1H / 2, { align: 'center', angle: 90 });
  doc.text('Temps bloqué', markerX + 2.5, tBlock2Y + tBlock2H / 2, { align: 'center', angle: 90 });

  // Time boundary markers on right (8.30, 11.45, 14.00, 16.30)
  const markerLabels: [number, string][] = [
    [blockY[3], '8.30'],
    [blockY[10], '11.45'],
    [blockY[12], '14.00'],
    [blockY[16], '16.30'],
  ];
  doc.setFontSize(4);
  doc.setTextColor(80, 80, 80);
  for (const [y, lbl] of markerLabels) {
    doc.text(lbl, markerX + 0.3, y + 1.2);
  }

  // Day remarks (evenly spaced 1–31)
  const remarkLineH = gridH / daysInMonth;
  let totalWorked = 0;
  let totalDu = 0;
  const worksiteHours: Record<string, number> = {};

  for (let d = 1; d <= daysInMonth; d++) {
    const date = new Date(data.year, data.month - 1, d);
    const wknd = isWeekend(date);
    const dateStr = `${data.year}-${String(data.month).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
    const entry = data.entries[dateStr];
    const ry = gridTop + (d - 1) * remarkLineH;

    // Subtle line
    setDraw(doc, [210, 210, 210], 0.08);
    doc.line(RIGHT_X, ry, RIGHT_X + RIGHT_W, ry);

    // Day number
    doc.setFontSize(4.5);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(wknd ? 180 : 60, 0, 0);
    doc.text(`${d}`, RIGHT_X + 1.5, ry + remarkLineH * 0.7);

    if (!wknd && entry) {
      const mH = calculateHours(entry.morning);
      const aH = calculateHours(entry.afternoon);
      const dayTotal = mH + aH;
      totalWorked += dayTotal;
      totalDu += entry.contractedHours;

      if (entry.morningWorksiteId && mH > 0) {
        worksiteHours[entry.morningWorksiteId] = (worksiteHours[entry.morningWorksiteId] || 0) + mH;
      }
      if (entry.afternoonWorksiteId && aH > 0) {
        worksiteHours[entry.afternoonWorksiteId] = (worksiteHours[entry.afternoonWorksiteId] || 0) + aH;
      }

      let text = '';
      if (entry.isVacation) {
        text = 'Vacances';
      } else if (entry.isHoliday) {
        text = 'Férié';
      } else {
        const parts: string[] = [];
        if (entry.morningWorksiteId && mH > 0) {
          const ws = data.worksites.find(w => w.id === entry.morningWorksiteId);
          if (ws) parts.push(`${ws.number} ${mH}h`);
        }
        if (entry.afternoonWorksiteId && aH > 0) {
          const ws = data.worksites.find(w => w.id === entry.afternoonWorksiteId);
          if (ws) parts.push(`${ws.number} ${aH}h`);
        }
        text = parts.join(' / ');
        if (entry.remark) text += text ? ` (${entry.remark})` : entry.remark;
      }

      if (text) {
        doc.setFont('helvetica', 'normal');
        doc.setFontSize(4.0);
        doc.setTextColor(0, 0, 0);
        doc.text(text, RIGHT_X + 5, ry + remarkLineH * 0.7, { maxWidth: RIGHT_W - 6.5 });
      }
    }
  }

  // ── Footer row 1: Daily totals ──────────────────────────────
  const f1Y = gridBot;

  setFill(doc, [240, 240, 240]);
  setDraw(doc, [0, 0, 0], 0.3);
  doc.rect(ML, f1Y, TIME_COL_W, FOOTER1_H, 'FD');
  doc.setFontSize(4.5);
  doc.setFont('helvetica', 'bold');
  doc.setTextColor(0, 0, 0);
  doc.text('Réel', ML + 1, f1Y + 3.5);

  for (let d = 1; d <= daysInMonth; d++) {
    const x = ML + TIME_COL_W + (d - 1) * dayColW;
    const date = new Date(data.year, data.month - 1, d);
    const wknd = isWeekend(date);
    const dateStr = `${data.year}-${String(data.month).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
    const entry = data.entries[dateStr];

    setFill(doc, wknd ? [210, 70, 70] : [255, 255, 255]);
    setDraw(doc, [0, 0, 0], 0.2);
    doc.rect(x, f1Y, dayColW, FOOTER1_H, 'FD');

    doc.setFontSize(4);
    doc.setFont('helvetica', 'normal');
    if (wknd) {
      doc.setTextColor(255, 255, 255);
      doc.text('X', x + dayColW / 2, f1Y + 3.5, { align: 'center' });
    } else if (entry) {
      const t = calculateHours(entry.morning) + calculateHours(entry.afternoon);
      if (t > 0) {
        doc.setTextColor(0, 0, 0);
        doc.text(`${t}`, x + dayColW / 2, f1Y + 3.5, { align: 'center' });
      }
    }
  }

  // Total réel label in right col
  setFill(doc, [220, 255, 220]);
  setDraw(doc, [0, 0, 0], 0.3);
  doc.rect(RIGHT_X, f1Y, RIGHT_W, FOOTER1_H, 'FD');
  doc.setFontSize(5.5);
  doc.setFont('helvetica', 'bold');
  doc.setTextColor(0, 0, 0);
  doc.text(`Réel: ${totalWorked.toFixed(2)}`, RIGHT_X + 2, f1Y + 3.5);

  // ── Footer row 2: Scheduled (Du) ────────────────────────────
  const f2Y = f1Y + FOOTER1_H;

  setFill(doc, [240, 240, 240]);
  setDraw(doc, [0, 0, 0], 0.3);
  doc.rect(ML, f2Y, TIME_COL_W, FOOTER2_H, 'FD');
  doc.setFontSize(4.5);
  doc.setFont('helvetica', 'bold');
  doc.setTextColor(0, 0, 0);
  doc.text('Dû', ML + 1, f2Y + 3.5);

  for (let d = 1; d <= daysInMonth; d++) {
    const x = ML + TIME_COL_W + (d - 1) * dayColW;
    const date = new Date(data.year, data.month - 1, d);
    const wknd = isWeekend(date);
    const dateStr = `${data.year}-${String(data.month).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
    const entry = data.entries[dateStr];

    setFill(doc, wknd ? [210, 70, 70] : [255, 255, 255]);
    setDraw(doc, [0, 0, 0], 0.2);
    doc.rect(x, f2Y, dayColW, FOOTER2_H, 'FD');

    doc.setFontSize(4);
    doc.setFont('helvetica', 'normal');
    if (wknd) {
      doc.setTextColor(255, 255, 255);
      doc.text('X', x + dayColW / 2, f2Y + 3.5, { align: 'center' });
    } else if (entry) {
      doc.setTextColor(0, 0, 0);
      doc.text(`${entry.contractedHours}`, x + dayColW / 2, f2Y + 3.5, { align: 'center' });
    }
  }

  // Total Dû in right col
  setFill(doc, [220, 230, 255]);
  setDraw(doc, [0, 0, 0], 0.3);
  doc.rect(RIGHT_X, f2Y, RIGHT_W, FOOTER2_H, 'FD');
  doc.setFontSize(5.5);
  doc.setFont('helvetica', 'bold');
  doc.setTextColor(0, 0, 0);
  doc.text(`Dû: ${totalDu.toFixed(2)}`, RIGHT_X + 2, f2Y + 3.5);

  // ── Footer row 3: Sup (overtime) ────────────────────────────
  const f3Y = f2Y + FOOTER2_H;
  const overtime = totalWorked - totalDu;

  setFill(doc, [255, 255, 255]);
  setDraw(doc, [0, 0, 0], 0.3);
  doc.rect(ML, f3Y, PW - 2 * ML, FOOTER3_H, 'FD');
  doc.setFontSize(5.5);
  doc.setFont('helvetica', 'bold');
  doc.setTextColor(0, 0, 0);
  doc.text(`Sup. : ${overtime >= 0 ? '+' : ''}${overtime.toFixed(2)}h`, ML + 3, f3Y + 3.5);

  // ── Summary: per-worksite totals ────────────────────────────
  const sumY = f3Y + FOOTER3_H + 1;
  setFill(doc, [255, 255, 255]);
  setDraw(doc, [0, 0, 0], 0.3);
  doc.rect(ML, sumY, PW - 2 * ML, SUMMARY_H, 'FD');

  doc.setFontSize(5.5);
  doc.setFont('helvetica', 'bold');
  doc.setTextColor(0, 0, 0);
  doc.text('Récapitulatif par chantier :', ML + 2, sumY + 4.5);

  let sx = ML + 2;
  let sy = sumY + 9;
  const colWidth = 55;

  for (const ws of data.worksites) {
    const h = worksiteHours[ws.id] || 0;
    if (h > 0) {
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(5);
      doc.text(`${ws.number}  ${ws.abbreviation} : ${h.toFixed(2)}h`, sx, sy);
      sx += colWidth;
      if (sx + colWidth > PW - ML) {
        sx = ML + 2;
        sy += 6;
      }
    }
  }

  // Outer border
  setDraw(doc, [0, 0, 0], 0.5);
  doc.rect(ML, MT, PW - 2 * ML, PH - 2 * MT);

  const filename = `CDH_${data.employeeName.replace(/\s+/g, '_')}_${getMonthName(data.month)}_${data.year}.pdf`;
  doc.save(filename);
}

export { getMonthName };
