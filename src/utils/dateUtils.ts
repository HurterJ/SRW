export function getDaysInMonth(year: number, month: number): Date[] {
  const days: Date[] = [];
  const daysInMonth = new Date(year, month, 0).getDate();
  for (let d = 1; d <= daysInMonth; d++) {
    days.push(new Date(year, month - 1, d));
  }
  return days;
}

export function isWeekend(date: Date): boolean {
  const day = date.getDay();
  return day === 0 || day === 6;
}

export function formatDate(date: Date): string {
  return date.toISOString().split('T')[0];
}

export function getDayName(date: Date): string {
  const names = ['Dimanche', 'Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi'];
  return names[date.getDay()];
}

export function getDayNameShort(date: Date): string {
  const names = ['Dim', 'Lun', 'Mar', 'Mer', 'Jeu', 'Ven', 'Sam'];
  return names[date.getDay()];
}

export function getMonthName(month: number): string {
  const months = [
    'Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin',
    'Juillet', 'Août', 'Septembre', 'Octobre', 'Novembre', 'Décembre',
  ];
  return months[month - 1];
}

export function timeToMinutes(time: string): number {
  const [h, m] = time.split(':').map(Number);
  return h * 60 + m;
}

export function minutesToDecimalHours(minutes: number): number {
  return Math.round((minutes / 60) * 4) / 4; // round to nearest 0.25h
}

export function calculateHours(range: { start: string; end: string } | null): number {
  if (!range) return 0;
  const diff = timeToMinutes(range.end) - timeToMinutes(range.start);
  if (diff <= 0) return 0;
  return minutesToDecimalHours(diff);
}

