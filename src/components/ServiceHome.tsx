import { useState } from "react";
import type { SyntheticEvent } from "react";
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

export function ServiceHome({ services, error, onOpen, onCreate, onRename, onDuplicate, onDelete }: ServiceHomeProps) {
  const [dialog, setDialog] = useState<"create" | ServiceSummary | null>(null);
  const [title, setTitle] = useState("");
  const [date, setDate] = useState(localDate);
  const [template, setTemplate] = useState<TemplateName>("sunday");

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

  return (
    <main className="service-home">
      <header className="home-topbar">
        <div className="brand-lockup">
          <div className="brand-mark" aria-hidden="true">CP</div>
          <div className="brand-name">Church Presenter</div>
          <span className="version-label">SERVICE LIBRARY</span>
        </div>
        <button className="present-button" onClick={openCreateDialog}><span>+</span> New service</button>
      </header>

      <section className="service-library">
        <div className="library-heading">
          <div>
            <span className="eyebrow">YOUR WORKSPACE</span>
            <h1>Services</h1>
          </div>
          <span className="library-count">{services.length} {services.length === 1 ? "service" : "services"}</span>
        </div>

        {error && <p className="library-error" role="alert">{error}</p>}
        {services.length > 0 ? (
          <div className="service-grid">
            {services.map((service) => (
              <article className="service-card" key={service.id}>
                <button className="service-card__open" onClick={() => onOpen(service.id)}>
                  <span className="service-card__date">{formatDate(service.date)}</span>
                  <strong>{service.title}</strong>
                  <span className="service-card__enter">Open service <span>→</span></span>
                </button>
                <div className="service-card__actions">
                  <button onClick={() => { setTitle(service.title); setDialog(service); }}>Rename</button>
                  <button onClick={() => onDuplicate(service.id)}>Duplicate</button>
                  <button className="service-card__delete" onClick={() => onDelete(service)}>Delete</button>
                </div>
              </article>
            ))}
          </div>
        ) : (
          <div className="library-empty">
            <div className="library-empty__mark" aria-hidden="true">CP</div>
            <h2>Your service list is ready</h2>
            <p>Create a blank service or start from a simple order template.</p>
            <button className="present-button" onClick={openCreateDialog}>+ Create service</button>
          </div>
        )}
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