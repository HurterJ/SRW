import { useState, useCallback } from 'react';
import type { TimesheetData, Worksite, DayEntry } from './types';
import {
  getDaysInMonth,
  isWeekend,
  formatDate,
  getDayNameShort,
  getMonthName,
  calculateHours,
} from './utils/dateUtils';
import { generatePDF } from './utils/pdfGenerator';

const DEFAULT_WORKSITES: Worksite[] = [
  { id: '1', number: '20834A02', abbreviation: 'SCII', fullName: 'SCII' },
  { id: '2', number: '03607A42', abbreviation: 'P41', fullName: 'Parking P41' },
  { id: '3', number: '24975A01', abbreviation: 'PLO', fullName: 'PLO Groupe Secours' },
  { id: '4', number: '22910A01', abbreviation: 'P47', fullName: 'Parking P47' },
  { id: '5', number: '12397A01', abbreviation: 'GS CA', fullName: 'Groupe de Secours Carouge' },
  { id: '6', number: '03607A39', abbreviation: 'CTA', fullName: 'Central Thermique' },
];

function makeDefaultEntry(
  dateStr: string,
  data: Pick<TimesheetData, 'defaultMorningStart' | 'defaultMorningEnd' | 'defaultAfternoonStart' | 'defaultAfternoonEnd' | 'defaultContractedHours'>
): DayEntry {
  return {
    date: dateStr,
    morning: { start: data.defaultMorningStart, end: data.defaultMorningEnd },
    afternoon: { start: data.defaultAfternoonStart, end: data.defaultAfternoonEnd },
    morningWorksiteId: null,
    afternoonWorksiteId: null,
    remark: '',
    isHoliday: false,
    isVacation: false,
    contractedHours: data.defaultContractedHours,
  };
}

function buildEntries(
  year: number,
  month: number,
  existing: Record<string, DayEntry>,
  defaults: Pick<TimesheetData, 'defaultMorningStart' | 'defaultMorningEnd' | 'defaultAfternoonStart' | 'defaultAfternoonEnd' | 'defaultContractedHours'>
): Record<string, DayEntry> {
  const days = getDaysInMonth(year, month);
  const entries: Record<string, DayEntry> = {};
  for (const d of days) {
    const key = formatDate(d);
    if (existing[key]) {
      entries[key] = existing[key];
    } else if (!isWeekend(d)) {
      entries[key] = makeDefaultEntry(key, defaults);
    }
  }
  return entries;
}

const now = new Date();

const initialData: TimesheetData = {
  employeeName: 'HURTER Julien',
  month: now.getMonth() + 1,
  year: now.getFullYear(),
  reference: 'F012',
  worksites: DEFAULT_WORKSITES,
  entries: {},
  defaultContractedHours: 8.5,
  defaultMorningStart: '08:00',
  defaultMorningEnd: '12:15',
  defaultAfternoonStart: '13:00',
  defaultAfternoonEnd: '17:30',
};
initialData.entries = buildEntries(initialData.year, initialData.month, {}, initialData);

// ─── Small UI helpers ──────────────────────────────────────────────────────

function Input({
  label,
  value,
  onChange,
  type = 'text',
  className = '',
}: {
  label: string;
  value: string | number;
  onChange: (v: string) => void;
  type?: string;
  className?: string;
}) {
  return (
    <label className={`flex flex-col gap-1 ${className}`}>
      <span className="text-xs font-semibold text-gray-600 uppercase tracking-wide">
        {label}
      </span>
      <input
        type={type}
        value={value}
        onChange={e => onChange(e.target.value)}
        className="border border-gray-300 rounded px-2 py-1 text-sm focus:outline-none focus:ring-2 focus:ring-blue-400"
      />
    </label>
  );
}

function TimeInput({
  value,
  onChange,
}: {
  value: string;
  onChange: (v: string) => void;
}) {
  return (
    <input
      type="time"
      value={value}
      onChange={e => onChange(e.target.value)}
      className="w-20 border border-gray-300 rounded px-1 py-0.5 text-xs focus:outline-none focus:ring-2 focus:ring-blue-400"
    />
  );
}

// ─── Main App ──────────────────────────────────────────────────────────────

export default function App() {
  const [data, setData] = useState<TimesheetData>(initialData);

  const updateField = useCallback(
    <K extends keyof TimesheetData>(key: K, value: TimesheetData[K]) => {
      setData(prev => ({ ...prev, [key]: value }));
    },
    []
  );

  const changeMonthYear = useCallback((month: number, year: number) => {
    setData(prev => ({
      ...prev,
      month,
      year,
      entries: buildEntries(year, month, prev.entries, prev),
    }));
  }, []);

  const updateEntry = useCallback((dateStr: string, patch: Partial<DayEntry>) => {
    setData(prev => ({
      ...prev,
      entries: {
        ...prev.entries,
        [dateStr]: {
          ...(prev.entries[dateStr] ?? makeDefaultEntry(dateStr, prev)),
          ...patch,
        },
      },
    }));
  }, []);

  const addWorksite = () => {
    const ws: Worksite = { id: Date.now().toString(), number: '', abbreviation: '', fullName: '' };
    updateField('worksites', [...data.worksites, ws]);
  };

  const updateWorksite = (id: string, patch: Partial<Worksite>) => {
    updateField('worksites', data.worksites.map(w => (w.id === id ? { ...w, ...patch } : w)));
  };

  const removeWorksite = (id: string) => {
    updateField('worksites', data.worksites.filter(w => w.id !== id));
  };

  const resetDefaults = () => {
    setData(prev => ({ ...prev, entries: buildEntries(prev.year, prev.month, {}, prev) }));
  };

  // ── Summary computation ─────────────────────────────────────

  const worksiteHours: Record<string, number> = {};
  let totalHours = 0;
  for (const entry of Object.values(data.entries)) {
    const mH = calculateHours(entry.morning);
    const aH = calculateHours(entry.afternoon);
    totalHours += mH + aH;
    if (entry.morningWorksiteId)
      worksiteHours[entry.morningWorksiteId] = (worksiteHours[entry.morningWorksiteId] || 0) + mH;
    if (entry.afternoonWorksiteId)
      worksiteHours[entry.afternoonWorksiteId] = (worksiteHours[entry.afternoonWorksiteId] || 0) + aH;
  }

  const days = getDaysInMonth(data.year, data.month);

  return (
    <div className="min-h-screen bg-gray-50">
      {/* Page header */}
      <header className="bg-blue-800 text-white px-6 py-4 shadow">
        <h1 className="text-xl font-bold tracking-wide">Contrôle des Heures — Générateur PDF</h1>
      </header>

      <main className="max-w-screen-2xl mx-auto p-4 space-y-4">
        {/* Section 1: Identity */}
        <section className="bg-white rounded shadow p-4">
          <h2 className="font-bold text-gray-700 mb-3 text-sm uppercase tracking-wide border-b pb-2">
            Informations générales
          </h2>
          <div className="flex flex-wrap gap-4">
            <Input
              label="Nom de l'employé"
              value={data.employeeName}
              onChange={v => updateField('employeeName', v)}
              className="flex-1 min-w-40"
            />
            <label className="flex flex-col gap-1">
              <span className="text-xs font-semibold text-gray-600 uppercase tracking-wide">Mois</span>
              <select
                value={data.month}
                onChange={e => changeMonthYear(Number(e.target.value), data.year)}
                className="border border-gray-300 rounded px-2 py-1 text-sm focus:outline-none focus:ring-2 focus:ring-blue-400"
              >
                {Array.from({ length: 12 }, (_, i) => (
                  <option key={i + 1} value={i + 1}>{getMonthName(i + 1)}</option>
                ))}
              </select>
            </label>
            <Input
              label="Année"
              value={data.year}
              onChange={v => changeMonthYear(data.month, Number(v))}
              type="number"
              className="w-24"
            />
            <Input
              label="Référence"
              value={data.reference}
              onChange={v => updateField('reference', v)}
              className="w-24"
            />
            <Input
              label="H/jour contractuel"
              value={data.defaultContractedHours}
              onChange={v => updateField('defaultContractedHours', Number(v))}
              type="number"
              className="w-28"
            />
          </div>
          <div className="flex flex-wrap gap-4 mt-3">
            <div>
              <span className="text-xs font-semibold text-gray-600 uppercase tracking-wide block mb-1">
                Horaires par défaut
              </span>
              <div className="flex items-center gap-2 text-sm flex-wrap">
                <span className="text-gray-500">Matin</span>
                <TimeInput value={data.defaultMorningStart} onChange={v => updateField('defaultMorningStart', v)} />
                <span>–</span>
                <TimeInput value={data.defaultMorningEnd} onChange={v => updateField('defaultMorningEnd', v)} />
                <span className="text-gray-500 ml-4">Après-midi</span>
                <TimeInput value={data.defaultAfternoonStart} onChange={v => updateField('defaultAfternoonStart', v)} />
                <span>–</span>
                <TimeInput value={data.defaultAfternoonEnd} onChange={v => updateField('defaultAfternoonEnd', v)} />
                <button
                  onClick={resetDefaults}
                  className="ml-4 text-xs bg-gray-100 hover:bg-gray-200 border border-gray-300 rounded px-2 py-1"
                >
                  Réinitialiser les jours
                </button>
              </div>
            </div>
          </div>
        </section>

        {/* Section 2: Worksites */}
        <section className="bg-white rounded shadow p-4">
          <div className="flex items-center justify-between mb-3 border-b pb-2">
            <h2 className="font-bold text-gray-700 text-sm uppercase tracking-wide">Chantiers</h2>
            <button
              onClick={addWorksite}
              className="text-xs bg-blue-600 text-white hover:bg-blue-700 rounded px-3 py-1"
            >
              + Ajouter
            </button>
          </div>
          <div className="space-y-2">
            {data.worksites.map(ws => (
              <div key={ws.id} className="flex items-center gap-2">
                <input
                  placeholder="N° (ex: 20834A02)"
                  value={ws.number}
                  onChange={e => updateWorksite(ws.id, { number: e.target.value })}
                  className="border border-gray-300 rounded px-2 py-1 text-xs w-32 focus:outline-none focus:ring-2 focus:ring-blue-400"
                />
                <input
                  placeholder="Abrév. (ex: SCII)"
                  value={ws.abbreviation}
                  onChange={e => updateWorksite(ws.id, { abbreviation: e.target.value })}
                  className="border border-gray-300 rounded px-2 py-1 text-xs w-24 focus:outline-none focus:ring-2 focus:ring-blue-400"
                />
                <input
                  placeholder="Nom complet"
                  value={ws.fullName}
                  onChange={e => updateWorksite(ws.id, { fullName: e.target.value })}
                  className="border border-gray-300 rounded px-2 py-1 text-xs flex-1 focus:outline-none focus:ring-2 focus:ring-blue-400"
                />
                <span className="text-xs text-blue-700 font-semibold w-16 text-right">
                  {(worksiteHours[ws.id] || 0).toFixed(2)}h
                </span>
                <button
                  onClick={() => removeWorksite(ws.id)}
                  className="text-red-400 hover:text-red-600 text-sm px-1"
                >
                  ✕
                </button>
              </div>
            ))}
          </div>
        </section>

        {/* Section 3: Daily entries table */}
        <section className="bg-white rounded shadow p-4">
          <div className="flex items-center justify-between mb-3 border-b pb-2">
            <h2 className="font-bold text-gray-700 text-sm uppercase tracking-wide">
              Saisie journalière — {getMonthName(data.month)} {data.year}
            </h2>
            <span className="text-sm font-bold text-blue-800">Total : {totalHours.toFixed(2)}h</span>
          </div>

          <div className="overflow-x-auto">
            <table className="text-xs border-collapse w-full">
              <thead>
                <tr className="bg-gray-100 text-gray-600">
                  <th className="border border-gray-300 px-2 py-1 text-left">#</th>
                  <th className="border border-gray-300 px-2 py-1 text-left">Jour</th>
                  <th className="border border-gray-300 px-2 py-1 text-center" colSpan={2}>Matin</th>
                  <th className="border border-gray-300 px-2 py-1 text-center" colSpan={2}>Après-midi</th>
                  <th className="border border-gray-300 px-2 py-1 text-center">H mat.</th>
                  <th className="border border-gray-300 px-2 py-1 text-center">H après.</th>
                  <th className="border border-gray-300 px-2 py-1 text-center">Total</th>
                  <th className="border border-gray-300 px-2 py-1 text-center">Chantier matin</th>
                  <th className="border border-gray-300 px-2 py-1 text-center">Chantier après-midi</th>
                  <th className="border border-gray-300 px-2 py-1 text-center">Vac.</th>
                  <th className="border border-gray-300 px-2 py-1 text-center">Fér.</th>
                  <th className="border border-gray-300 px-2 py-1 text-left">Remarques</th>
                </tr>
              </thead>
              <tbody>
                {days.map(d => {
                  const dateStr = formatDate(d);
                  const weekend = isWeekend(d);
                  const entry = data.entries[dateStr];
                  const dayNum = d.getDate();
                  const dayName = getDayNameShort(d);
                  const mH = entry ? calculateHours(entry.morning) : 0;
                  const aH = entry ? calculateHours(entry.afternoon) : 0;
                  const total = mH + aH;

                  const rowBg = weekend
                    ? 'bg-red-50 text-red-400'
                    : entry?.isVacation
                    ? 'bg-blue-50'
                    : entry?.isHoliday
                    ? 'bg-yellow-50'
                    : '';

                  if (weekend) {
                    return (
                      <tr key={dateStr} className={rowBg}>
                        <td className="border border-gray-200 px-2 py-0.5 font-bold">{dayNum}</td>
                        <td className="border border-gray-200 px-2 py-0.5">{dayName}</td>
                        <td colSpan={12} className="border border-gray-200 px-2 py-0.5 text-center text-red-300">
                          Week-end
                        </td>
                      </tr>
                    );
                  }

                  return (
                    <tr key={dateStr} className={`${rowBg} hover:bg-gray-50`}>
                      <td className="border border-gray-200 px-2 py-0.5 font-semibold text-gray-700">{dayNum}</td>
                      <td className="border border-gray-200 px-2 py-0.5 text-gray-600">{dayName}</td>

                      <td className="border border-gray-200 px-1 py-0.5">
                        <TimeInput
                          value={entry?.morning?.start ?? data.defaultMorningStart}
                          onChange={v => updateEntry(dateStr, { morning: { start: v, end: entry?.morning?.end ?? data.defaultMorningEnd } })}
                        />
                      </td>
                      <td className="border border-gray-200 px-1 py-0.5">
                        <TimeInput
                          value={entry?.morning?.end ?? data.defaultMorningEnd}
                          onChange={v => updateEntry(dateStr, { morning: { start: entry?.morning?.start ?? data.defaultMorningStart, end: v } })}
                        />
                      </td>
                      <td className="border border-gray-200 px-1 py-0.5">
                        <TimeInput
                          value={entry?.afternoon?.start ?? data.defaultAfternoonStart}
                          onChange={v => updateEntry(dateStr, { afternoon: { start: v, end: entry?.afternoon?.end ?? data.defaultAfternoonEnd } })}
                        />
                      </td>
                      <td className="border border-gray-200 px-1 py-0.5">
                        <TimeInput
                          value={entry?.afternoon?.end ?? data.defaultAfternoonEnd}
                          onChange={v => updateEntry(dateStr, { afternoon: { start: entry?.afternoon?.start ?? data.defaultAfternoonStart, end: v } })}
                        />
                      </td>

                      <td className="border border-gray-200 px-2 py-0.5 text-center font-mono text-blue-700">
                        {mH > 0 ? mH : ''}
                      </td>
                      <td className="border border-gray-200 px-2 py-0.5 text-center font-mono text-blue-700">
                        {aH > 0 ? aH : ''}
                      </td>
                      <td className="border border-gray-200 px-2 py-0.5 text-center font-mono font-bold text-blue-900">
                        {total > 0 ? total : ''}
                      </td>

                      <td className="border border-gray-200 px-1 py-0.5">
                        <select
                          value={entry?.morningWorksiteId ?? ''}
                          onChange={e => updateEntry(dateStr, { morningWorksiteId: e.target.value || null })}
                          className="w-full border border-gray-200 rounded px-1 py-0.5 text-xs focus:outline-none focus:ring-2 focus:ring-blue-400"
                        >
                          <option value="">—</option>
                          {data.worksites.map(ws => (
                            <option key={ws.id} value={ws.id}>{ws.abbreviation} ({ws.number})</option>
                          ))}
                        </select>
                      </td>

                      <td className="border border-gray-200 px-1 py-0.5">
                        <select
                          value={entry?.afternoonWorksiteId ?? ''}
                          onChange={e => updateEntry(dateStr, { afternoonWorksiteId: e.target.value || null })}
                          className="w-full border border-gray-200 rounded px-1 py-0.5 text-xs focus:outline-none focus:ring-2 focus:ring-blue-400"
                        >
                          <option value="">—</option>
                          {data.worksites.map(ws => (
                            <option key={ws.id} value={ws.id}>{ws.abbreviation} ({ws.number})</option>
                          ))}
                        </select>
                      </td>

                      <td className="border border-gray-200 px-2 py-0.5 text-center">
                        <input
                          type="checkbox"
                          checked={entry?.isVacation ?? false}
                          onChange={e => updateEntry(dateStr, { isVacation: e.target.checked })}
                          className="cursor-pointer"
                        />
                      </td>
                      <td className="border border-gray-200 px-2 py-0.5 text-center">
                        <input
                          type="checkbox"
                          checked={entry?.isHoliday ?? false}
                          onChange={e => updateEntry(dateStr, { isHoliday: e.target.checked })}
                          className="cursor-pointer"
                        />
                      </td>
                      <td className="border border-gray-200 px-1 py-0.5">
                        <input
                          type="text"
                          value={entry?.remark ?? ''}
                          onChange={e => updateEntry(dateStr, { remark: e.target.value })}
                          placeholder="Remarque..."
                          className="w-full border-0 text-xs focus:outline-none bg-transparent min-w-24"
                        />
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </section>

        {/* Section 4: Summary + PDF button */}
        <section className="bg-white rounded shadow p-4">
          <h2 className="font-bold text-gray-700 text-sm uppercase tracking-wide border-b pb-2 mb-3">
            Récapitulatif
          </h2>
          <div className="flex flex-wrap gap-3 mb-4">
            {data.worksites
              .filter(ws => (worksiteHours[ws.id] || 0) > 0)
              .map(ws => (
                <div key={ws.id} className="bg-blue-50 border border-blue-200 rounded px-3 py-2 text-sm">
                  <div className="font-bold text-blue-800">{ws.abbreviation}</div>
                  <div className="text-xs text-gray-500">{ws.number}</div>
                  <div className="font-mono font-bold text-blue-700">{(worksiteHours[ws.id] || 0).toFixed(2)}h</div>
                </div>
              ))}
            <div className="bg-green-50 border border-green-200 rounded px-3 py-2 text-sm ml-auto">
              <div className="font-bold text-green-800">TOTAL</div>
              <div className="font-mono font-bold text-green-700 text-lg">{totalHours.toFixed(2)}h</div>
            </div>
          </div>

          <button
            onClick={() => generatePDF(data)}
            className="w-full bg-blue-700 hover:bg-blue-800 text-white font-bold py-3 rounded text-sm tracking-wide transition-colors shadow"
          >
            Générer le PDF — {getMonthName(data.month)} {data.year}
          </button>
        </section>
      </main>
    </div>
  );
}
