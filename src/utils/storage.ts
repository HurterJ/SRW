import type { TimesheetData } from '../types';

const KEY_PROFILES = 'srw_profiles';
const KEY_CURRENT  = 'srw_current_profile';
const dataKey = (name: string) => `srw_data_${name}`;

export function getProfiles(): string[] {
  try {
    return JSON.parse(localStorage.getItem(KEY_PROFILES) ?? '[]');
  } catch { return []; }
}

export function getCurrentProfile(): string | null {
  return localStorage.getItem(KEY_CURRENT);
}

export function setCurrentProfile(name: string) {
  localStorage.setItem(KEY_CURRENT, name);
}

export function addProfile(name: string) {
  const profiles = getProfiles();
  if (!profiles.includes(name)) {
    profiles.push(name);
    localStorage.setItem(KEY_PROFILES, JSON.stringify(profiles));
  }
}

export function deleteProfile(name: string) {
  const profiles = getProfiles().filter(p => p !== name);
  localStorage.setItem(KEY_PROFILES, JSON.stringify(profiles));
  localStorage.removeItem(dataKey(name));
  if (getCurrentProfile() === name) {
    localStorage.removeItem(KEY_CURRENT);
  }
}

export function loadData(profile: string): Partial<TimesheetData> | null {
  try {
    const raw = localStorage.getItem(dataKey(profile));
    return raw ? JSON.parse(raw) : null;
  } catch { return null; }
}

export function saveData(profile: string, data: TimesheetData) {
  localStorage.setItem(dataKey(profile), JSON.stringify(data));
}
