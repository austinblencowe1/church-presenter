import { usePresenterStore } from '../data/serviceStore';
import { slideList } from '../types/presentation';
import { SlideView } from './SlideView';

interface Props {
  onEnd: () => void;
}

export function PresenterView({ onEnd }: Props) {
  const slides = usePresenterStore((s) => s.slides);
  const sel = usePresenterStore((s) => s.sel);
  const aj = usePresenterStore((s) => s.aj);
  const mediaURLs = usePresenterStore((s) => s.mediaURLs);
  const go = usePresenterStore((s) => s.go);
  const selectItem = usePresenterStore((s) => s.selectItem);
  const selectSlide = usePresenterStore((s) => s.selectSlide);
  const toggleClear = usePresenterStore((s) => s.toggleClear);
  const clear = usePresenterStore((s) => s.clear);
  const services = usePresenterStore((s) => s.services);
  const activeService = usePresenterStore((s) => s.activeService);

  const item = slides[sel];
  const parts = item ? slideList(item) : [];
  const part = parts[aj];

  // next slide across service
  let nextItem = item;
  let nextPart = parts[aj + 1];
  if (!nextPart) {
    const nextIt = slides[sel + 1];
    if (nextIt) {
      nextItem = nextIt;
      nextPart = slideList(nextIt)[0];
    } else {
      nextItem = null as unknown as typeof item;
      nextPart = undefined;
    }
  }

  const title = services[activeService]
    ? [services[activeService].tag, services[activeService].name].filter(Boolean).join(' - ')
    : 'Service';

  return (
    <div id="presenter-view" className="presenter-view">
      <div className="pv-top">
        <span className="pv-title">{title}</span>
        <span className="sp" />
        <button type="button" onClick={() => go(-1)}>
          Previous
        </button>
        <button type="button" onClick={() => go(1)}>
          Next
        </button>
        <button type="button" className={clear === 'black' ? 'on' : ''} onClick={() => toggleClear('black')}>
          B
        </button>
        <button type="button" className={clear === 'white' ? 'on' : ''} onClick={() => toggleClear('white')}>
          W
        </button>
        <button type="button" className={clear === 'logo' ? 'on' : ''} onClick={() => toggleClear('logo')}>
          L
        </button>
        <button type="button" className="pri" onClick={onEnd}>
          End
        </button>
      </div>

      <div className="pv-outline">
        <h3>Service order</h3>
        {slides.map((it, i) => (
          <div
            key={i}
            className={`pv-oi ${i === sel ? 'on' : ''}`}
            onClick={() => selectItem(i)}
            role="button"
            tabIndex={0}
          >
            <b>
              {i + 1}. {it.t}
            </b>
            <small>
              {it.k} · {slideList(it).length} slides
            </small>
          </div>
        ))}
      </div>

      <div className="pv-strip">
        <div className="pv-strip-heading">Slides</div>
        {parts.map((p, j) => (
          <button
            key={p.id}
            type="button"
            className={`pv-thumb ${j === aj ? 'on' : ''}`}
            onClick={() => selectSlide(sel, j)}
          >
            <div className="pv-slide-frame">
              <SlideView item={item!} part={p} mediaURLs={mediaURLs} />
            </div>
            <span className="pn">{p.n}</span>
          </button>
        ))}
      </div>

      <div className="pv-main">
        <div>
          <h3>Current</h3>
          {item && part && <SlideView item={item} part={part} mediaURLs={mediaURLs} stage />}
        </div>
        <div>
          <h3>Next</h3>
          {nextItem && nextPart ? (
            <SlideView item={nextItem} part={nextPart} mediaURLs={mediaURLs} stage />
          ) : (
            <p className="note">End of service</p>
          )}
        </div>
      </div>
    </div>
  );
}
