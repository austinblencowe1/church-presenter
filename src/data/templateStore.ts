import type { SlideTemplate } from "../types/presentation";

const storageKey = "church-presenter-slide-templates";

const builtInTemplates: SlideTemplate[] = [
  { id: "builtin-scripture", name: "Scripture passage", content: "{{verse}}\n\n{{reference}} · {{translation}}", background: "#17211f", fontSize: 42, textAlign: "center", transition: "fade", transitionDuration: 300 },
  { id: "builtin-announcement", name: "Announcement", content: "{{title}}\n\n{{details}}", background: "#17211f", fontSize: 48, textAlign: "center", transition: "fade", transitionDuration: 300 },
  { id: "builtin-lower-third", name: "Lower third", content: "{{name}}\n{{role}}", background: "#17211f", fontSize: 36, textAlign: "left", transition: "fade", transitionDuration: 300 },
];

export function listSlideTemplates(): SlideTemplate[] {
  try {
    const saved = localStorage.getItem(storageKey);
    if (saved === null) {
      localStorage.setItem(storageKey, JSON.stringify(builtInTemplates));
      return builtInTemplates;
    }
    const parsed: unknown = JSON.parse(saved);
    return Array.isArray(parsed) ? parsed as SlideTemplate[] : [...builtInTemplates];
  } catch {
    return [...builtInTemplates];
  }
}

export function saveSlideTemplate(template: SlideTemplate): SlideTemplate[] {
  const templates = [template, ...listSlideTemplates().filter((entry) => entry.id !== template.id)];
  localStorage.setItem(storageKey, JSON.stringify(templates));
  return templates;
}

export function deleteSlideTemplate(templateId: string): SlideTemplate[] {
  const templates = listSlideTemplates().filter((template) => template.id !== templateId);
  localStorage.setItem(storageKey, JSON.stringify(templates));
  return templates;
}
