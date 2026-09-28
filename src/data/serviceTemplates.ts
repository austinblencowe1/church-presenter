import type { Service, ServiceItemType } from "../types/presentation";

export type TemplateName = "blank" | "sunday" | "evening";

interface TemplateItem {
  title: string;
  type: ServiceItemType;
}

const templates: Record<TemplateName, TemplateItem[]> = {
  blank: [],
  sunday: [
    { title: "Welcome", type: "welcome" },
    { title: "Song", type: "song" },
    { title: "Bible Reading", type: "bible" },
    { title: "Sermon", type: "sermon" },
    { title: "Song", type: "song" },
    { title: "Closing", type: "other" },
  ],
  evening: [
    { title: "Welcome", type: "welcome" },
    { title: "Song", type: "song" },
    { title: "Bible Reading", type: "bible" },
    { title: "Sermon", type: "sermon" },
    { title: "Closing", type: "other" },
  ],
};

export function createServiceFromTemplate(title: string, date: string, template: TemplateName): Service {
  return {
    id: crypto.randomUUID(),
    title: title.trim() || "New Service",
    date,
    items: templates[template].map(({ title: itemTitle, type }) => ({
      id: crypto.randomUUID(),
      title: itemTitle,
      type,
      slides: [{
        id: crypto.randomUUID(),
        type: "text",
        text: itemTitle,
        background: "#18231f",
        textAlign: "center",
        fontSize: 56,
      }],
    })),
  };
}

export function createInitialService(date: string): Service {
  return createServiceFromTemplate("Sunday Service", date, "sunday");
}