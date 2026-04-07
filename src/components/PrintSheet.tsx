import type { TimesheetData } from '../types';
import { getMonthName, isWeekend, calculateHours } from '../utils/dateUtils';

// ── Time grid structure matching the F012 Excel exactly ───────────────────
// Each entry: label shown, start time (min), height in grid units
interface Row {
  label: string;
  startMin: number;
  endMin: number;
  isLunch: boolean;
}

const GRID_ROWS: Row[] = [
  { label: '7',     startMin: 420, endMin: 450, isLunch: false },
  { label: '7.30',  startMin: 450, endMin: 480, isLunch: false },
  { label: '8',     startMin: 480, endMin: 510, isLunch: false },
  { label: '8.30',  startMin: 510, endMin: 540, isLunch: false },
  { label: '9',     startMin: 540, endMin: 570, isLunch: false },
  { label: '9.30',  startMin: 570, endMin: 600, isLunch: false },
  { label: '10',    startMin: 600, endMin: 630, isLunch: false },
  { label: '10.30', startMin: 630, endMin: 660, isLunch: false },
  { label: '11',    startMin: 660, endMin: 690, isLunch: false },
  { label: '11.30', startMin: 690, endMin: 720, isLunch: false },
  { label: '12.00\u00a012.15\u00a013.00', startMin: 720, endMin: 780, isLunch: true },
  { label: '13.30', startMin: 810, endMin: 840, isLunch: false },
  { label: '14',    startMin: 840, endMin: 870, isLunch: false },
  { label: '14.30', startMin: 870, endMin: 900, isLunch: false },
  { label: '15',    startMin: 900, endMin: 930, isLunch: false },
  { label: '15.30', startMin: 930, endMin: 960, isLunch: false },
  { label: '16',    startMin: 960, endMin: 990, isLunch: false },
  { label: '16.30', startMin: 990, endMin: 1020, isLunch: false },
  { label: '17',    startMin: 1020, endMin: 1050, isLunch: false },
  { label: '17.30', startMin: 1050, endMin: 1080, isLunch: false },
  { label: '18',    startMin: 1080, endMin: 1110, isLunch: false },
  { label: '18.30', startMin: 1110, endMin: 1110, isLunch: false },
];

function timeToMin(t: string) {
  const [h, m] = t.split(':').map(Number);
  return h * 60 + m;
}

function cellStatus(row: Row, entry: { morning: { start: string; end: string } | null; afternoon: { start: string; end: string } | null } | undefined): 'worked' | 'lunch' | 'empty' {
  if (!entry) return 'empty';
  if (row.isLunch) return 'lunch';

  const mid = (row.startMin + row.endMin) / 2;
  const mS = entry.morning ? timeToMin(entry.morning.start) : null;
  const mE = entry.morning ? timeToMin(entry.morning.end)   : null;
  const aS = entry.afternoon ? timeToMin(entry.afternoon.start) : null;
  const aE = entry.afternoon ? timeToMin(entry.afternoon.end)   : null;

  if (mS !== null && mE !== null && mid >= mS && mid < mE) return 'worked';
  if (aS !== null && aE !== null && mid >= aS && mid < aE) return 'worked';
  if (mE !== null && aS !== null && mid >= mE && mid < aS) return 'lunch';
  return 'empty';
}

// ── React component (renders into a new window for printing) ──────────────
interface Props {
  data: TimesheetData;
}

export default function PrintSheet({ data }: Props) {
  const daysInMonth = new Date(data.year, data.month, 0).getDate();
  const days = Array.from({ length: daysInMonth }, (_, i) => i + 1);

  // Pre-compute per-day info
  const dayInfo = days.map(d => {
    const dateStr = `${data.year}-${String(data.month).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
    const date = new Date(data.year, data.month - 1, d);
    const entry = data.entries[dateStr];
    const wknd = isWeekend(date);
    const mH = entry ? calculateHours(entry.morning) : 0;
    const aH = entry ? calculateHours(entry.afternoon) : 0;
    return { d, dateStr, date, entry, wknd, mH, aH, total: mH + aH };
  });

  // Per-worksite totals
  const wsHours: Record<string, number> = {};
  let totalReal = 0;
  let totalDu = 0;
  for (const { entry, mH, aH, wknd } of dayInfo) {
    if (wknd || !entry) continue;
    totalReal += mH + aH;
    totalDu += entry.contractedHours;
    if (entry.morningWorksiteId)   wsHours[entry.morningWorksiteId]   = (wsHours[entry.morningWorksiteId]   || 0) + mH;
    if (entry.afternoonWorksiteId) wsHours[entry.afternoonWorksiteId] = (wsHours[entry.afternoonWorksiteId] || 0) + aH;
  }

  // Colors
  const C = {
    worked:  '#1a1a1a',
    weekend: '#c83232',
    lunch:   '#b4b4b4',
    vacay:   '#a8cce8',
    holiday: '#ffd966',
    empty:   '#ffffff',
    border:  '#888888',
    hdrBg:   '#d4d4d4',
    red_txt: '#ffffff',
  };

  const cellBorder = `1px solid ${C.border}`;
  const thinBorder = `0.5px solid #cccccc`;

  // Temps bloqué zones (rows index 3-9 = 8:30-11:30, rows 12-16 = 14:00-16:30)
  const tempsBloqueTop1 = 3;   // 8:30 index
  const tempsBloqueBot1 = 9;   // 11:30 index
  const tempsBloqueTop2 = 12;  // 14:00 index
  const tempsBloqueBot2 = 16;  // 16:30 index

  return (
    <div style={{
      fontFamily: 'Arial, Helvetica, sans-serif',
      fontSize: '7px',
      width: '277mm',
      minHeight: '190mm',
      margin: '0',
      padding: '0',
      backgroundColor: '#fff',
      color: '#000',
    }}>
      {/* Row 1: Company */}
      <div style={{
        display: 'flex',
        justifyContent: 'space-between',
        alignItems: 'center',
        padding: '2px 4px',
        borderBottom: cellBorder,
        fontSize: '7px',
        fontWeight: 'bold',
      }}>
        <span>INGENIEURS-CONSEILS SCHERLER SA</span>
        <span style={{ fontStyle: 'italic' }}>srg | engineering</span>
      </div>

      {/* Row 2: CONTRÔLE DES HEURES */}
      <div style={{
        display: 'grid',
        gridTemplateColumns: '1fr 1fr 1fr auto',
        alignItems: 'center',
        padding: '2px 4px',
        borderBottom: cellBorder,
        fontSize: '7.5px',
        fontWeight: 'bold',
        gap: '8px',
      }}>
        <span>CONTRÔLE DES HEURES</span>
        <span>NOM :&nbsp;&nbsp;&nbsp;{data.employeeName}</span>
        <span>MOIS :&nbsp;&nbsp;{getMonthName(data.month)} {data.year}</span>
        <span>{data.reference}</span>
      </div>

      {/* Main grid */}
      <div style={{ display: 'flex', width: '100%' }}>

        {/* Left time label column */}
        <div style={{ display: 'flex', flexDirection: 'column', width: '24px', flexShrink: 0 }}>
          {/* Header */}
          <div style={{
            background: C.hdrBg, border: cellBorder, textAlign: 'center',
            fontWeight: 'bold', fontSize: '7px', height: '14px',
            display: 'flex', alignItems: 'center', justifyContent: 'center',
          }}>H</div>
          {/* Time rows */}
          {GRID_ROWS.map((row, ri) => (
            <div key={ri} style={{
              background: C.hdrBg,
              border: thinBorder,
              borderTop: row.isLunch ? cellBorder : thinBorder,
              borderBottom: row.isLunch ? cellBorder : thinBorder,
              fontSize: row.isLunch ? '5.5px' : '6px',
              padding: '1px',
              height: row.isLunch ? '28px' : '10px',
              lineHeight: '1.2',
              display: 'flex',
              alignItems: 'center',
              flexShrink: 0,
              whiteSpace: 'pre',
              color: '#333',
            }}>
              {row.label.replace('\u00a0', '\n')}
            </div>
          ))}
          {/* Footer rows */}
          <div style={{ background: C.hdrBg, border: cellBorder, height: '10px', fontSize: '6px', padding: '1px', display: 'flex', alignItems: 'center' }}>Réel</div>
          <div style={{ background: C.hdrBg, border: cellBorder, height: '10px', fontSize: '6px', padding: '1px', display: 'flex', alignItems: 'center' }}>Dû</div>
        </div>

        {/* Day columns */}
        <div style={{ display: 'flex', flex: 1, minWidth: 0 }}>
          {dayInfo.map(({ d, entry, wknd, total }) => {
            const colBg = wknd ? C.weekend : C.empty;
            return (
              <div key={d} style={{ display: 'flex', flexDirection: 'column', flex: 1, minWidth: 0 }}>
                {/* Day number header */}
                <div style={{
                  background: wknd ? C.weekend : C.hdrBg,
                  color: wknd ? C.red_txt : '#000',
                  border: thinBorder,
                  textAlign: 'center',
                  fontWeight: wknd ? 'bold' : 'normal',
                  fontSize: '6px',
                  height: '14px',
                  display: 'flex', alignItems: 'center', justifyContent: 'center',
                }}>{d}</div>

                {/* Grid cells */}
                {GRID_ROWS.map((row, ri) => {
                  let bg = colBg;
                  if (!wknd && entry) {
                    if (entry.isVacation) bg = C.vacay;
                    else if (entry.isHoliday) bg = C.holiday;
                    else {
                      const status = cellStatus(row, entry);
                      if (status === 'worked') bg = C.worked;
                      else if (status === 'lunch') bg = C.lunch;
                      else bg = C.empty;
                    }
                  }
                  const isTempsBloq = ri >= tempsBloqueTop1 && ri <= tempsBloqueBot1 ||
                                      ri >= tempsBloqueTop2 && ri <= tempsBloqueBot2;
                  return (
                    <div key={ri} style={{
                      background: bg,
                      border: thinBorder,
                      borderTop: row.isLunch ? cellBorder : (isTempsBloq && ri === tempsBloqueTop1 || ri === tempsBloqueTop2) ? '1px solid #666' : thinBorder,
                      borderBottom: row.isLunch ? cellBorder : thinBorder,
                      height: row.isLunch ? '28px' : '10px',
                      flexShrink: 0,
                    }} />
                  );
                })}

                {/* Footer: total actual */}
                <div style={{
                  background: wknd ? C.weekend : C.empty,
                  color: wknd ? C.red_txt : '#000',
                  border: thinBorder,
                  textAlign: 'center',
                  fontSize: '6px',
                  height: '10px',
                  display: 'flex', alignItems: 'center', justifyContent: 'center',
                  fontWeight: 'bold',
                }}>
                  {wknd ? 'X' : total > 0 ? String(total) : ''}
                </div>

                {/* Footer: contracted hours */}
                <div style={{
                  background: wknd ? C.weekend : C.empty,
                  color: wknd ? C.red_txt : '#000',
                  border: thinBorder,
                  textAlign: 'center',
                  fontSize: '6px',
                  height: '10px',
                  display: 'flex', alignItems: 'center', justifyContent: 'center',
                }}>
                  {wknd ? 'X' : entry ? String(entry.contractedHours) : ''}
                </div>
              </div>
            );
          })}
        </div>

        {/* Right: Temps bloqué + Remarques */}
        <div style={{ display: 'flex', flexShrink: 0, width: '78mm' }}>
          {/* Temps bloqué side markers */}
          <div style={{ width: '14px', display: 'flex', flexDirection: 'column', position: 'relative', flexShrink: 0 }}>
            <div style={{ height: '14px', background: C.hdrBg, border: thinBorder }} />
            {GRID_ROWS.map((row, ri) => {
              const inZone1 = ri >= tempsBloqueTop1 && ri <= tempsBloqueBot1;
              const inZone2 = ri >= tempsBloqueTop2 && ri <= tempsBloqueBot2;
              return (
                <div key={ri} style={{
                  height: row.isLunch ? '28px' : '10px',
                  background: (inZone1 || inZone2) ? '#e8e8e8' : 'transparent',
                  borderLeft: (inZone1 || inZone2) ? '1px solid #999' : 'none',
                  borderRight: (inZone1 || inZone2) ? '1px solid #999' : 'none',
                  borderTop: (ri === tempsBloqueTop1 || ri === tempsBloqueTop2) ? '1px solid #666' : 'none',
                  borderBottom: (ri === tempsBloqueBot1 || ri === tempsBloqueBot2) ? '1px solid #666' : 'none',
                  fontSize: '5px',
                  overflow: 'hidden',
                  flexShrink: 0,
                }} />
              );
            })}
            <div style={{ height: '10px', background: C.hdrBg, border: thinBorder }} />
            <div style={{ height: '10px', background: C.hdrBg, border: thinBorder }} />
          </div>

          {/* Remarques column */}
          <div style={{ flex: 1, display: 'flex', flexDirection: 'column', border: `1px solid ${C.border}` }}>
            {/* Header */}
            <div style={{
              background: C.hdrBg, fontWeight: 'bold', fontSize: '6px',
              textAlign: 'center', height: '14px',
              display: 'flex', alignItems: 'center', justifyContent: 'center',
              borderBottom: cellBorder,
            }}>Remarques</div>

            {/* Grid area for remarks (same total height as grid rows) */}
            <div style={{
              flex: 1,
              display: 'flex',
              flexDirection: 'column',
            }}>
              {dayInfo.map(({ d, entry, wknd, mH, aH }) => {
                let text = '';
                if (!wknd && entry) {
                  if (entry.isVacation) text = 'Vacances';
                  else if (entry.isHoliday) text = 'Férié';
                  else {
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
                }

                return (
                  <div key={d} style={{
                    flex: 1,
                    display: 'flex',
                    alignItems: 'center',
                    borderBottom: `0.5px solid #ddd`,
                    padding: '0 2px',
                    minHeight: 0,
                    overflow: 'hidden',
                  }}>
                    <span style={{ fontWeight: 'bold', marginRight: '3px', color: wknd ? '#999' : '#333', fontSize: '5.5px', flexShrink: 0 }}>
                      {d}
                    </span>
                    <span style={{ fontSize: '5.5px', color: '#000', overflow: 'hidden', whiteSpace: 'nowrap', textOverflow: 'ellipsis' }}>
                      {text}
                    </span>
                  </div>
                );
              })}
            </div>

            {/* Footer Réel/Dû */}
            <div style={{
              borderTop: cellBorder, padding: '1px 2px', fontSize: '6px', fontWeight: 'bold',
              height: '10px', display: 'flex', alignItems: 'center',
              background: '#e8ffe8',
            }}>
              Réel : {totalReal.toFixed(2)}h
            </div>
            <div style={{
              borderTop: thinBorder, padding: '1px 2px', fontSize: '6px', fontWeight: 'bold',
              height: '10px', display: 'flex', alignItems: 'center',
              background: '#e8eaff',
            }}>
              Dû : {totalDu.toFixed(2)}h
            </div>
          </div>
        </div>
      </div>

      {/* Summary: per-chantier totals */}
      <div style={{
        borderTop: cellBorder,
        padding: '3px 4px',
        fontSize: '6.5px',
        display: 'flex',
        flexWrap: 'wrap',
        gap: '8px',
        alignItems: 'center',
      }}>
        <strong>Récapitulatif :</strong>
        {data.worksites.filter(ws => (wsHours[ws.id] || 0) > 0).map(ws => (
          <span key={ws.id}>
            <strong>{ws.number}</strong> {ws.abbreviation} : <strong>{(wsHours[ws.id] || 0).toFixed(2)}h</strong>
          </span>
        ))}
        <span style={{ marginLeft: 'auto', fontWeight: 'bold' }}>
          Sup. : {totalReal - totalDu >= 0 ? '+' : ''}{(totalReal - totalDu).toFixed(2)}h
        </span>
      </div>
    </div>
  );
}
