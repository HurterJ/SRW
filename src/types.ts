export interface Worksite {
  id: string;
  number: string;       // e.g. "20834A02"
  abbreviation: string; // e.g. "SCII"
  fullName: string;     // e.g. "SCII Chantier principal"
}

export interface TimeRange {
  start: string; // "HH:MM"
  end: string;   // "HH:MM"
}

export interface DayEntry {
  date: string;                    // "YYYY-MM-DD"
  morning: TimeRange | null;
  afternoon: TimeRange | null;
  morningWorksiteId: string | null;
  afternoonWorksiteId: string | null;
  remark: string;
  isHoliday: boolean;
  isVacation: boolean;
  contractedHours: number;         // "Du" column (e.g. 8.5)
}

export interface TimesheetData {
  employeeName: string;
  month: number;  // 1–12
  year: number;
  reference: string;               // e.g. "F012"
  worksites: Worksite[];
  entries: Record<string, DayEntry>; // keyed by "YYYY-MM-DD"
  defaultContractedHours: number;  // default 8.5
  defaultMorningStart: string;     // "08:00"
  defaultMorningEnd: string;       // "12:15"
  defaultAfternoonStart: string;   // "13:00"
  defaultAfternoonEnd: string;     // "17:30"
}
