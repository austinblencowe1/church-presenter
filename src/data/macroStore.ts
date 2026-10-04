import type { ShowMacro } from "../types/presentation";

const storageKey = "church-presenter-macros";

export function listMacros(): ShowMacro[] {
  try {
    const value: unknown = JSON.parse(localStorage.getItem(storageKey) ?? "[]");
    return Array.isArray(value) ? value as ShowMacro[] : [];
  } catch {
    return [];
  }
}

export function saveMacro(macro: ShowMacro): ShowMacro[] {
  const macros = [macro, ...listMacros().filter((entry) => entry.id !== macro.id)];
  localStorage.setItem(storageKey, JSON.stringify(macros));
  return macros;
}

export function deleteMacro(macroId: string): ShowMacro[] {
  const macros = listMacros().filter((macro) => macro.id !== macroId);
  localStorage.setItem(storageKey, JSON.stringify(macros));
  return macros;
}
