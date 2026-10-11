import { usePresenterStore } from '../data/serviceStore';

function formatTitle(name?: string, tag?: string) {
  const t = (tag || '').trim();
  const n = (name || '').trim();
  if (t && n) return `${t} - ${n}`;
  return t || n || 'Service';
}

export function HomeServices() {
  const services = usePresenterStore((s) => s.services);
  const openService = usePresenterStore((s) => s.openService);

  return (
    <section className="home">
      <div className="home-head">
        <div>
          <h1>Services</h1>
          <p className="note">Create and organise your services.</p>
        </div>
      </div>
      <div className="service-list">
        {services.length === 0 && <p className="note">No services yet.</p>}
        {services.map((svc, i) => (
          <div className="service-card" key={i}>
            <div className="info">
              <h3>{formatTitle(svc.name, svc.tag)}</h3>
              <small>{svc.date || 'No date'}</small>
            </div>
            <button type="button" className="pri" onClick={() => openService(i)}>
              Open
            </button>
          </div>
        ))}
      </div>
    </section>
  );
}
