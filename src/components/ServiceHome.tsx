import { useState } from "react";
import type { ChangeEvent, SyntheticEvent } from "react";
import type { Service, ServiceSummary } from "../types/presentation";
import { createServiceFromTemplate, type TemplateName } from "../data/serviceTemplates";

interface ServiceHomeProps {
  readonly services: ServiceSummary[];
  readonly error: string;
  readonly onOpen: (serviceId: string) => void;
  readonly onCreate: (service: Service) => void;
  readonly onRename: (serviceId: string, title: string) => void;
  readonly onDuplicate: (serviceId: string) => void;
  readonly onDelete: (service: ServiceSummary) => void;
  readonly onImport: (service: Service) => void;
  readonly onMoveToFolder: (serviceId: string, folder: string) => void;
}

function localDate(): string {
  const date = new Date();
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function formatDate(value: string): string {
  if (!value) return "Date not set";
  return new Intl.DateTimeFormat(undefined, { day: "numeric", month: "long", year: "numeric" })
    .format(new Date(`${value}T12:00:00`));
}

function parseServiceBackup(value: unknown): Service {
  const isRecord = (entry: unknown): entry is Record<string, unknown> => typeof entry === "object" && entry !== null;
  const validTypes = new Set(["welcome", "song", "bible", "sermon", "announcement", "other"]);
  if (!isRecord(value) || typeof value.title !== "string" || typeof value.date !== "string"
    || (value.folder !== undefined && typeof value.folder !== "string") || !Array.isArray(value.items)) {
    throw new Error("This file does not contain a Church Presenter service backup.");
  }
  const valid = value.items.every((item) => isRecord(item)
    && typeof item.title === "string"
    && typeof item.type === "string"
    && validTypes.has(item.type)
    && (item.reference === undefined || typeof item.reference === "string")
    && (item.bibleVersion === undefined || typeof item.bibleVersion === "string")
    && (item.groups === undefined || (Array.isArray(item.groups) && item.groups.every((group) => isRecord(group)
      && typeof group.id === "string" && typeof group.name === "string" && typeof group.color === "string"
      && Array.isArray(group.slideIds) && group.slideIds.every((id) => typeof id === "string"))))
    && (item.arrangements === undefined || (Array.isArray(item.arrangements) && item.arrangements.every((arrangement) => isRecord(arrangement)
      && typeof arrangement.id === "string" && typeof arrangement.name === "string"
      && Array.isArray(arrangement.groupIds) && arrangement.groupIds.every((id) => typeof id === "string"))))
    && (item.activeArrangementId === undefined || item.activeArrangementId === null || typeof item.activeArrangementId === "string")
    && Array.isArray(item.slides)
    && item.slides.every((slide) => isRecord(slide)
      && (slide.type === undefined || slide.type === "text")
      && typeof slide.text === "string"
      && typeof slide.background === "string"
      && typeof slide.fontSize === "number"
      && (slide.transition === undefined || ["cut", "fade", "dissolve", "push"].includes(String(slide.transition)))
      && (slide.transitionDuration === undefined || typeof slide.transitionDuration === "number")
      && ["left", "center", "right"].includes(String(slide.textAlign))
      && (slide.sectionLabel === undefined || slide.sectionLabel === null || typeof slide.sectionLabel === "string")
      && (slide.notes === undefined || slide.notes === null || typeof slide.notes === "string")
      && (slide.backgroundMedia === undefined || slide.backgroundMedia === null || typeof slide.backgroundMedia === "string")
      && (slide.cueMacroIds === undefined || (Array.isArray(slide.cueMacroIds) && slide.cueMacroIds.every((id) => typeof id === "string")))
      && (slide.objects === undefined || (Array.isArray(slide.objects) && slide.objects.every((object) => isRecord(object)
        && typeof object.id === "string"
        && ["text", "rectangle", "ellipse", "image", "video"].includes(String(object.kind))
        && ["x", "y", "width", "height", "rotation", "opacity", "strokeWidth", "fontSize"].every((key) => typeof object[key] === "number")
        && ["fill", "stroke", "color", "text"].every((key) => typeof object[key] === "string")
        && ["left", "center", "right"].includes(String(object.textAlign))
        && (object.source === undefined || object.source === null || typeof object.source === "string"))))));
  if (!valid) throw new Error("The service backup has invalid items or slides.");
  return value as unknown as Service;
}

export function ServiceHome({ services, error, onOpen, onCreate, onRename, onDuplicate, onDelete, onImport, onMoveToFolder }: ServiceHomeProps) {
  const [dialog, setDialog] = useState<"create" | ServiceSummary | null>(null);
  const [title, setTitle] = useState("");
  const [date, setDate] = useState(localDate);
  const [template, setTemplate] = useState<TemplateName>("sunday");
  const [importError, setImportError] = useState("");
  const [search, setSearch] = useState("");
  const [activeFolder, setActiveFolder] = useState("all");
  const [newFolderName, setNewFolderName] = useState("");
  const [folders, setFolders] = useState<string[]>(() => {
    try {
      const value: unknown = JSON.parse(localStorage.getItem("church-presenter-folders") ?? "[]");
      return Array.isArray(value) ? value.filter((folder): folder is string => typeof folder === "string") : [];
    } catch {
      return [];
    }
  });
  const folderNames = [...new Set([...folders, ...services.map((service) => service.folder).filter((folder): folder is string => Boolean(folder))])]
    .sort((first, second) => first.localeCompare(second));
  const filteredServices = services.filter((service) => service.title.toLocaleLowerCase().includes(search.trim().toLocaleLowerCase())
    && (activeFolder === "all" || (service.folder ?? "") === activeFolder));

  function createFolder(event: SyntheticEvent<HTMLFormElement>) {
    event.preventDefault();
    const name = newFolderName.trim();
    if (!name) return;
    const next = [...new Set([...folders, name])];
    setFolders(next);
    localStorage.setItem("church-presenter-folders", JSON.stringify(next));
    setNewFolderName("");
    setActiveFolder("all");
  }

  function openCreateDialog() {
    setTitle("");
    setDate(localDate());
    setTemplate("sunday");
    setDialog("create");
  }

  function handleSubmit(event: SyntheticEvent<HTMLFormElement>) {
    event.preventDefault();
    if (dialog === "create") {
      onCreate(createServiceFromTemplate(title, date, template));
    } else if (dialog) {
      onRename(dialog.id, title.trim());
    }
    setDialog(null);
  }

  function importBackup(event: ChangeEvent<HTMLInputElement>) {
    const file = event.currentTarget.files?.[0];
    event.currentTarget.value = "";
    if (!file) return;
    void file.text().then((text) => {
      const service = parseServiceBackup(JSON.parse(text) as unknown);
      setImportError("");
      onImport(service);
    }).catch((cause: unknown) => {
      setImportError(cause instanceof Error ? cause.message : "The selected file is not a valid service backup.");
    });
  }

  return (
    <main className="service-home console-theme-light">
      <header className="home-topbar">
        <div className="brand-lockup">
          <div className="brand-mark" aria-hidden="true">CP</div>
          <div className="brand-name">Church Presenter</div>
          <span className="version-label">SERVICE LIBRARY</span>
        </div>
        <div className="home-actions">
          <label className="import-backup-button">Import backup<input type="file" accept=".json,application/json" onChange={importBackup} /></label>
          <button className="present-button" onClick={openCreateDialog}><span>+</span> New service</button>
        </div>
      </header>

      <section className="service-library">
        <div className="library-heading">
          <div>
            <span className="eyebrow">YOUR WORKSPACE</span>
            <h1>Services</h1>
          </div>
          <span className="library-count">{services.length} {services.length === 1 ? "service" : "services"}</span>
        </div>

        {services.length > 0 && (
          <div className="service-library-tools">
            <label className="service-search">
            <span aria-hidden="true">⌕</span>
            <input value={search} onChange={(event) => setSearch(event.currentTarget.value)} placeholder="Search services..." aria-label="Search services" />
            {search && <button type="button" onClick={() => setSearch("")} aria-label="Clear search">×</button>}
            </label>
            <label className="service-folder-filter">
              <span>FOLDER</span>
              <select value={activeFolder} onChange={(event) => setActiveFolder(event.currentTarget.value)}>
                <option value="all">All folders</option>
                <option value="">Unfiled</option>
                {folderNames.map((folder) => <option value={folder} key={folder}>{folder}</option>)}
              </select>
            </label>
            <form className="create-folder-form" onSubmit={createFolder}>
              <input maxLength={48} value={newFolderName} onChange={(event) => setNewFolderName(event.currentTarget.value)} placeholder="New folder name" aria-label="New folder name" />
              <button type="submit" disabled={!newFolderName.trim()}>+ Folder</button>
            </form>
          </div>
        )}

        {error && <p className="library-error" role="alert">{error}</p>}
        {importError && <p className="library-error" role="alert">{importError}</p>}
        {filteredServices.length > 0 ? (
          <div className="service-grid">
            {filteredServices.map((service) => (
              <article className="service-card" key={service.id}>
                <button className="service-card__open" onClick={() => onOpen(service.id)}>
                  <span className="service-card__date">{formatDate(service.date)}</span>
                  <strong>{service.title}</strong>
                  <span className="service-card__enter">Open service <span>→</span></span>
                </button>
                <div className="service-card__actions">
                  <button onClick={() => { setTitle(service.title); setDialog(service); }}>Rename</button>
                  <button onClick={() => onDuplicate(service.id)}>Duplicate</button>
                  <label className="service-card__folder">
                    <span>Folder</span>
                    <select aria-label={`Move ${service.title} to folder`} value={service.folder ?? ""} onChange={(event) => onMoveToFolder(service.id, event.currentTarget.value)}>
                      <option value="">Unfiled</option>
                      {folderNames.map((folder) => <option value={folder} key={folder}>{folder}</option>)}
                    </select>
                  </label>
                  <button className="service-card__delete" onClick={() => onDelete(service)}>Delete</button>
                </div>
              </article>
            ))}
          </div>
        ) : services.length === 0 ? (
          <div className="library-empty">
            <div className="library-empty__mark" aria-hidden="true">CP</div>
            <h2>Your service list is ready</h2>
            <p>Create a blank service or start from a simple order template.</p>
            <button className="present-button" onClick={openCreateDialog}>+ Create service</button>
          </div>
        ) : <div className="library-empty library-empty--search"><h2>No matching services</h2><p>Try a different title.</p></div>}
      </section>

      {dialog && (
        <div className="dialog-backdrop">
          <dialog open className="service-dialog" aria-labelledby="service-dialog-title">
            <form onSubmit={handleSubmit}>
              <div className="dialog-heading">
                <div>
                  <span className="eyebrow">{dialog === "create" ? "NEW SERVICE" : "SERVICE DETAILS"}</span>
                  <h2 id="service-dialog-title">{dialog === "create" ? "Create service" : "Rename service"}</h2>
                </div>
                <button type="button" className="dialog-close" aria-label="Close" onClick={() => setDialog(null)}>×</button>
              </div>
              <label className="dialog-field">
                <span>Service name</span>
                <input required maxLength={80} value={title} onChange={(event) => setTitle(event.currentTarget.value)} placeholder="Sunday PM" />
              </label>
              {dialog === "create" ? (
                <>
                  <label className="dialog-field">
                    <span>Date</span>
                    <input type="date" required value={date} onChange={(event) => setDate(event.currentTarget.value)} />
                  </label>
                  <label className="dialog-field">
                    <span>Template</span>
                    <select value={template} onChange={(event) => setTemplate(event.currentTarget.value as TemplateName)}>
                      <option value="blank">Blank service</option>
                      <option value="sunday">Sunday service</option>
                      <option value="evening">Evening service</option>
                    </select>
                  </label>
                </>
              ) : (
                <p className="dialog-note">The service date stays {formatDate(dialog.date)}.</p>
              )}
              <div className="dialog-actions">
                <button type="button" className="secondary-button" onClick={() => setDialog(null)}>Cancel</button>
                <button type="submit" className="present-button">{dialog === "create" ? "Create service" : "Save name"}</button>
              </div>
            </form>
          </dialog>
        </div>
      )}
    </main>
  );
}
