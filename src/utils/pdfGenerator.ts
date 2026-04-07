import jsPDF from 'jspdf';
import type { TimesheetData } from '../types';
import {
  getMonthName,
  timeToMinutes,
  calculateHours,
  isWeekend,
} from './dateUtils';

// Time grid: 15-min slots from 7:00 to 18:30
const TIME_START_MIN = 7 * 60;       // 420
const TIME_END_MIN = 18 * 60 + 30;   // 1110
const SLOT_DURATION = 15;
const TOTAL_SLOTS = (TIME_END_MIN - TIME_START_MIN) / SLOT_DURATION; // 46

function slotMinutes(slot: number): number {
  return TIME_START_MIN + slot * SLOT_DURATION;
}

function minutesToLabel(minutes: number): string {
  const h = Math.floor(minutes / 60);
  const m = minutes % 60;
  return `${h}.${m === 0 ? '00' : m.toString().padStart(2, '0')}`;
}

type RGB = [number, number, number];

function setFill(doc: jsPDF, color: RGB) {
  doc.setFillColor(color[0], color[1], color[2]);
}

function setDraw(doc: jsPDF, color: RGB, lw = 0.1) {
  doc.setDrawColor(color[0], color[1], color[2]);
  doc.setLineWidth(lw);
}

export function generatePDF(data: TimesheetData): void {
  const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });

  const PW = 297;
  const PH = 210;
  const ML = 5; // left margin
  const MT = 5; // top margin

  // ── Layout ─────────────────────────────────────────────────
  const HEADER_H = 10;
  const SUB_H = 6; // day-numbers sub-header
  const TIME_COL_W = 11;
  const daysInMonth = new Date(data.year, data.month, 0).getDate();
  const FOOTER_H = 18;
  const SUMMARY_H = 30;

  const gridTop = MT + HEADER_H + SUB_H;
  const gridAvailH = PH - MT - HEADER_H - SUB_H - FOOTER_H - SUMMARY_H;
  const SLOT_H = gridAvailH / TOTAL_SLOTS;
  const gridBottom = gridTop + TOTAL_SLOTS * SLOT_H;

  // Right column width based on remaining space
  const DAY_COL_W = (PW - ML * 2 - TIME_COL_W - 75) / daysInMonth;
  const gridW = DAY_COL_W * daysInMonth;
  const RIGHT_X = ML + TIME_COL_W + gridW + 2;
  const RIGHT_W = PW - RIGHT_X - ML;

  // ── Header ─────────────────────────────────────────────────
  setFill(doc, [255, 255, 255]);
  setDraw(doc, [0, 0, 0], 0.4);
  doc.rect(ML, MT, PW - 2 * ML, HEADER_H, 'FD');

  doc.setFontSize(10);
  doc.setFont('helvetica', 'bold');
  doc.text('CONTRÔLE DES HEURES', ML + 3, MT + 6.5);

  doc.setFontSize(9);
  doc.text(`NOM : ${data.employeeName}`, PW * 0.36, MT + 6.5);
  doc.text(`MOIS : ${getMonthName(data.month)} ${data.year}`, PW * 0.58, MT + 6.5);
  doc.text(data.reference, PW - ML - 15, MT + 6.5);

  // ── Day-number sub-header ───────────────────────────────────
  const subY = MT + HEADER_H;

  // Time column label
  setFill(doc, [230, 230, 230]);
  setDraw(doc, [0, 0, 0], 0.3);
  doc.rect(ML, subY, TIME_COL_W, SUB_H, 'FD');
  doc.setFontSize(6);
  doc.setFont('helvetica', 'bold');
  doc.text('H', ML + TIME_COL_W / 2, subY + 4, { align: 'center' });

  for (let d = 1; d <= daysInMonth; d++) {
    const x = ML + TIME_COL_W + (d - 1) * DAY_COL_W;
    const date = new Date(data.year, data.month - 1, d);
    const weekend = isWeekend(date);
    setFill(doc, weekend ? [220, 80, 80] : [230, 230, 230]);
    setDraw(doc, [0, 0, 0], 0.2);
    doc.rect(x, subY, DAY_COL_W, SUB_H, 'FD');
    doc.setFontSize(5.5);
    doc.setFont('helvetica', weekend ? 'bold' : 'normal');
    doc.setTextColor(weekend ? 255 : 0, 255, 255);
    if (weekend) doc.setTextColor(255, 255, 255);
    else doc.setTextColor(0, 0, 0);
    doc.text(`${d}`, x + DAY_COL_W / 2, subY + 4, { align: 'center' });
  }

  // Right-column header
  setFill(doc, [230, 230, 230]);
  setDraw(doc, [0, 0, 0], 0.3);
  doc.rect(RIGHT_X, subY, RIGHT_W, SUB_H, 'FD');
  doc.setFontSize(6);
  doc.setFont('helvetica', 'bold');
  doc.setTextColor(0, 0, 0);
  doc.text('Remarques', RIGHT_X + RIGHT_W / 2, subY + 4, { align: 'center' });

  // ── Time grid (slots × days) ────────────────────────────────
  for (let slot = 0; slot < TOTAL_SLOTS; slot++) {
    const mins = slotMinutes(slot);
    const y = gridTop + slot * SLOT_H;
    const isHalfHour = mins % 30 === 0;

    // Time label every 30 min (spans 2 slots)
    if (isHalfHour) {
      setFill(doc, [245, 245, 245]);
      setDraw(doc, [100, 100, 100], 0.15);
      doc.rect(ML, y, TIME_COL_W, SLOT_H * 2, 'FD');
      doc.setFontSize(4.5);
      doc.setFont('helvetica', 'normal');
      doc.setTextColor(0, 0, 0);
      doc.text(minutesToLabel(mins), ML + 1, y + SLOT_H + 0.8);
    }

    // Day columns
    for (let d = 1; d <= daysInMonth; d++) {
      const x = ML + TIME_COL_W + (d - 1) * DAY_COL_W;
      const date = new Date(data.year, data.month - 1, d);
      const weekend = isWeekend(date);
      const dateStr = `${data.year}-${String(data.month).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
      const entry = data.entries[dateStr];

      let fill: RGB;

      if (weekend) {
        fill = [220, 100, 100];
      } else if (entry?.isHoliday || entry?.isVacation) {
        fill = [200, 230, 255];
      } else if (entry) {
        const slotStart = mins;
        const slotEnd = mins + SLOT_DURATION;

        const mStart = entry.morning ? timeToMinutes(entry.morning.start) : null;
        const mEnd = entry.morning ? timeToMinutes(entry.morning.end) : null;
        const aStart = entry.afternoon ? timeToMinutes(entry.afternoon.start) : null;
        const aEnd = entry.afternoon ? timeToMinutes(entry.afternoon.end) : null;

        const inMorning =
          mStart !== null &&
          mEnd !== null &&
          slotStart >= mStart &&
          slotEnd <= mEnd;
        const inAfternoon =
          aStart !== null &&
          aEnd !== null &&
          slotStart >= aStart &&
          slotEnd <= aEnd;

        // Lunch break (between morning end and afternoon start)
        const inLunch =
          mEnd !== null &&
          aStart !== null &&
          slotStart >= mEnd &&
          slotEnd <= aStart;

        if (inMorning || inAfternoon) {
          fill = [40, 40, 40];
        } else if (inLunch) {
          fill = [180, 180, 180];
        } else {
          fill = [255, 255, 255];
        }
      } else {
        fill = [255, 255, 255];
      }

      setFill(doc, fill);
      setDraw(doc, [200, 200, 200], 0.05);
      doc.rect(x, y, DAY_COL_W, SLOT_H, 'FD');
    }
  }

  // Grid border lines (horizontal every 30 min, thicker)
  setDraw(doc, [80, 80, 80], 0.2);
  for (let slot = 0; slot <= TOTAL_SLOTS; slot += 2) {
    const y = gridTop + slot * SLOT_H;
    doc.line(ML, y, ML + TIME_COL_W + gridW, y);
  }

  // Outer borders
  setDraw(doc, [0, 0, 0], 0.4);
  doc.rect(ML, gridTop, TIME_COL_W, TOTAL_SLOTS * SLOT_H);
  doc.rect(ML + TIME_COL_W, gridTop, gridW, TOTAL_SLOTS * SLOT_H);

  // ── Right column — day remarks ──────────────────────────────
  setDraw(doc, [0, 0, 0], 0.3);
  doc.rect(RIGHT_X, gridTop, RIGHT_W, TOTAL_SLOTS * SLOT_H);

  const entryLineH = (TOTAL_SLOTS * SLOT_H) / daysInMonth;

  for (let d = 1; d <= daysInMonth; d++) {
    const date = new Date(data.year, data.month - 1, d);
    const weekend = isWeekend(date);
    const dateStr = `${data.year}-${String(data.month).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
    const entry = data.entries[dateStr];
    const ry = gridTop + (d - 1) * entryLineH;

    // Subtle separator line
    setDraw(doc, [200, 200, 200], 0.1);
    doc.line(RIGHT_X, ry, RIGHT_X + RIGHT_W, ry);

    // Day number
    doc.setFontSize(5);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(weekend ? 180 : 0, 0, 0);
    doc.text(`${d}`, RIGHT_X + 1.5, ry + entryLineH * 0.65);

    if (entry && !weekend) {
      doc.setFont('helvetica', 'normal');
      doc.setTextColor(0, 0, 0);

      let parts: string[] = [];

      if (entry.isVacation) {
        parts.push('Vacances');
      } else if (entry.isHoliday) {
        parts.push('Férié');
      } else {
        if (entry.morningWorksiteId && entry.morning) {
          const ws = data.worksites.find(w => w.id === entry.morningWorksiteId);
          if (ws) {
            const h = calculateHours(entry.morning);
            parts.push(`${ws.number} ${h}h`);
          }
        }
        if (
          entry.afternoonWorksiteId &&
          entry.afternoon &&
          entry.afternoonWorksiteId !== entry.morningWorksiteId
        ) {
          const ws = data.worksites.find(w => w.id === entry.afternoonWorksiteId);
          if (ws) {
            const h = calculateHours(entry.afternoon);
            parts.push(`${ws.number} ${h}h`);
          }
        } else if (
          entry.afternoonWorksiteId &&
          entry.afternoon &&
          entry.afternoonWorksiteId === entry.morningWorksiteId
        ) {
          // Same site all day — show combined
          const ws = data.worksites.find(w => w.id === entry.afternoonWorksiteId);
          if (ws && entry.morning) {
            parts = [];
            const mH = calculateHours(entry.morning);
            const aH = calculateHours(entry.afternoon);
            parts.push(`${ws.number} ${mH}h / ${ws.number} ${aH}h`);
          }
        }
      }

      if (entry.remark) parts.push(`(${entry.remark})`);

      doc.setFontSize(4.2);
      doc.text(parts.join(' / '), RIGHT_X + 4.5, ry + entryLineH * 0.65, {
        maxWidth: RIGHT_W - 6,
      });
    }
  }

  // ── Footer — daily totals ───────────────────────────────────
  const footerY = gridBottom + 1;

  // Time column label
  setFill(doc, [245, 245, 245]);
  setDraw(doc, [0, 0, 0], 0.3);
  doc.rect(ML, footerY, TIME_COL_W, FOOTER_H / 2);
  doc.setFontSize(5);
  doc.setFont('helvetica', 'bold');
  doc.setTextColor(0, 0, 0);
  doc.text('Total', ML + 1, footerY + FOOTER_H / 4 + 1);

  let grandTotal = 0;
  const worksiteHours: Record<string, number> = {};

  for (let d = 1; d <= daysInMonth; d++) {
    const x = ML + TIME_COL_W + (d - 1) * DAY_COL_W;
    const date = new Date(data.year, data.month - 1, d);
    const weekend = isWeekend(date);
    const dateStr = `${data.year}-${String(data.month).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
    const entry = data.entries[dateStr];

    setFill(doc, weekend ? [220, 100, 100] : [255, 255, 255]);
    setDraw(doc, [0, 0, 0], 0.2);
    doc.rect(x, footerY, DAY_COL_W, FOOTER_H / 2, 'FD');

    doc.setFontSize(4);
    doc.setFont('helvetica', 'normal');

    if (weekend) {
      doc.setTextColor(255, 255, 255);
      doc.text('X', x + DAY_COL_W / 2, footerY + FOOTER_H / 4 + 1, { align: 'center' });
    } else if (entry) {
      const mH = calculateHours(entry.morning);
      const aH = calculateHours(entry.afternoon);
      const total = mH + aH;
      grandTotal += total;

      if (entry.morningWorksiteId && mH > 0) {
        worksiteHours[entry.morningWorksiteId] = (worksiteHours[entry.morningWorksiteId] || 0) + mH;
      }
      if (entry.afternoonWorksiteId && aH > 0) {
        worksiteHours[entry.afternoonWorksiteId] = (worksiteHours[entry.afternoonWorksiteId] || 0) + aH;
      }

      doc.setTextColor(0, 0, 0);
      if (total > 0) {
        doc.text(`${total}`, x + DAY_COL_W / 2, footerY + FOOTER_H / 4 + 1, { align: 'center' });
      }
    } else {
      doc.setTextColor(0, 0, 0);
    }
  }

  // Grand total
  setFill(doc, [230, 255, 230]);
  setDraw(doc, [0, 0, 0], 0.3);
  doc.rect(RIGHT_X, footerY, RIGHT_W, FOOTER_H / 2, 'FD');
  doc.setFontSize(6);
  doc.setFont('helvetica', 'bold');
  doc.setTextColor(0, 0, 0);
  doc.text(`Réel : ${grandTotal.toFixed(2)}h`, RIGHT_X + 2, footerY + FOOTER_H / 4 + 1);

  // ── Summary — per-worksite totals ───────────────────────────
  const sumY = footerY + FOOTER_H / 2 + 2;
  doc.setFontSize(6);
  doc.setFont('helvetica', 'bold');
  doc.text('Récapitulatif par chantier :', ML, sumY + 4);

  let sumX = ML + 50;
  doc.setFont('helvetica', 'normal');
  for (const ws of data.worksites) {
    const h = worksiteHours[ws.id] || 0;
    if (h > 0) {
      doc.text(`${ws.number} ${ws.abbreviation} : ${h.toFixed(2)}h`, sumX, sumY + 4);
      sumX += 45;
      if (sumX + 44 > PW - ML) {
        sumX = ML + 50;
        // second line if needed
      }
    }
  }

  // Outer page border
  setDraw(doc, [0, 0, 0], 0.5);
  doc.rect(ML, MT, PW - 2 * ML, PH - 2 * MT);

  // Save
  const filename = `CDH_${data.employeeName.replace(/\s+/g, '_')}_${getMonthName(data.month)}_${data.year}.pdf`;
  doc.save(filename);
}

export { getMonthName };
