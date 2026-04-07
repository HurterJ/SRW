import { useState } from 'react';
import { getProfiles, addProfile, deleteProfile } from '../utils/storage';

interface Props {
  onSelect: (name: string) => void;
}

export default function ProfileSelector({ onSelect }: Props) {
  const [profiles, setProfiles] = useState<string[]>(getProfiles);
  const [newName, setNewName] = useState('');
  const [error, setError] = useState('');

  const handleSelect = (name: string) => onSelect(name);

  const handleAdd = () => {
    const trimmed = newName.trim();
    if (!trimmed) { setError('Entrez un nom.'); return; }
    if (profiles.includes(trimmed)) { setError('Ce profil existe déjà.'); return; }
    addProfile(trimmed);
    setProfiles(getProfiles());
    setNewName('');
    setError('');
    onSelect(trimmed);
  };

  const handleDelete = (name: string, e: React.MouseEvent) => {
    e.stopPropagation();
    if (!confirm(`Supprimer le profil "${name}" et toutes ses données ?`)) return;
    deleteProfile(name);
    setProfiles(getProfiles());
  };

  return (
    <div className="min-h-screen bg-blue-900 flex items-center justify-center p-4">
      <div className="bg-white rounded-xl shadow-2xl w-full max-w-md p-8">
        <h1 className="text-2xl font-bold text-gray-800 mb-1 text-center">
          Contrôle des Heures
        </h1>
        <p className="text-gray-400 text-sm text-center mb-8">Sélectionnez votre profil</p>

        {/* Existing profiles */}
        {profiles.length > 0 ? (
          <div className="space-y-2 mb-6">
            {profiles.map(name => (
              <button
                key={name}
                onClick={() => handleSelect(name)}
                className="w-full flex items-center justify-between bg-blue-50 hover:bg-blue-100 border border-blue-200 rounded-lg px-4 py-3 text-left transition-colors group"
              >
                <span className="font-semibold text-blue-900">{name}</span>
                <div className="flex items-center gap-2">
                  <span className="text-xs text-blue-400 group-hover:text-blue-600">→ Ouvrir</span>
                  <span
                    onClick={e => handleDelete(name, e)}
                    className="text-red-300 hover:text-red-500 text-xs px-1 cursor-pointer"
                    title="Supprimer ce profil"
                  >✕</span>
                </div>
              </button>
            ))}
          </div>
        ) : (
          <p className="text-gray-400 text-sm text-center mb-6">Aucun profil — créez le vôtre ci-dessous.</p>
        )}

        {/* Create new profile */}
        <div className="border-t pt-6">
          <p className="text-xs font-semibold text-gray-500 uppercase tracking-wide mb-2">
            Nouveau profil
          </p>
          <div className="flex gap-2">
            <input
              type="text"
              placeholder="Prénom Nom (ex: Julien Hurter)"
              value={newName}
              onChange={e => { setNewName(e.target.value); setError(''); }}
              onKeyDown={e => e.key === 'Enter' && handleAdd()}
              className="flex-1 border border-gray-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-400"
            />
            <button
              onClick={handleAdd}
              className="bg-blue-700 hover:bg-blue-800 text-white font-bold px-4 py-2 rounded-lg text-sm transition-colors"
            >
              Créer
            </button>
          </div>
          {error && <p className="text-red-500 text-xs mt-1">{error}</p>}
        </div>
      </div>
    </div>
  );
}
