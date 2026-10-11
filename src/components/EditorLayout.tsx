import { usePresenterStore } from '../data/serviceStore';
import { slideList } from '../types/presentation';
import { SlideView } from './SlideView';

export function EditorLayout() {
  const slides = usePresenterStore((s) => s.slides);
  const sel = usePresenterStore((s) => s.sel);
  const aj = usePresenterStore((s) => s.aj);
  const mediaURLs = usePresenterStore((s) => s.mediaURLs);
  const selectItem = usePresenterStore((s) => s.selectItem);
  const selectSlide = usePresenterStore((s) => s.selectSlide);
  const go = usePresenterStore((s) => s.go);
  const clear = usePresenterStore((s) => s.clear);
  const toggleClear = usePresenterStore((s) => s.toggleClear);
  const addBlankSlide = usePresenterStore((s) => s.addBlankSlide);

  const item = slides[sel];
  const parts = item ? slideList(item) : [];
  const part = parts[aj];

  return (
    <main className="editor-main">
      <nav id="outline" aria-label="Service overview">
        <h3>Service order</h3>
        {slides.map((it, i) => (
          <div
            key={i}
            className={`oi ${i === sel ? 'on' : ''}`}
            onClick={() => selectItem(i)}
            role="button"
            tabIndex={0}
          >
            <span className="ot">
              {i + 1}. {it.t}
            </span>
          </div>
        ))}
      </nav>

      <nav id="strip" aria-label="Service slides">
        {slides.map((it, i) => {
          const ss = slideList(it);
          return (
            <div key={i}>
              <div className={`hb ${i === sel ? 'on' : ''}`}>
                <button type="button" className="hd" onClick={() => selectItem(i)}>
                  <b>{it.t}</b>
                  <small>
                    {it.k}, {ss.length} slide{ss.length === 1 ? '' : 's'}
                  </small>
                </button>
              </div>
              {i === sel &&
                ss.map((p, j) => (
                  <button
                    key={p.id}
                    type="button"
                    className={`th ${j === aj ? 'on' : ''}`}
                    onClick={() => selectSlide(i, j)}
                  >
                    <SlideView item={it} part={p} mediaURLs={mediaURLs} />
                    <span className="pn">{p.n}</span>
                  </button>
                ))}
            </div>
          );
        })}
      </nav>

      <section className="mid">
        <div id="stage" className="stage-preview">
          {item && part ? (
            <div className={`cv ${clear ? 'cleared' : ''}`}>
              <SlideView item={item} part={part} mediaURLs={mediaURLs} editable />
              {clear === 'black' && <div className="ov black" />}
              {clear === 'white' && <div className="ov white" />}
              {clear === 'logo' && <div className="ov logo">Logo</div>}
            </div>
          ) : (
            <p className="note">Select a slide</p>
          )}
        </div>
        <div className="bar">
          <div className="g">
            <button type="button" onClick={() => go(-1)}>
              Previous
            </button>
            <button type="button" id="nx" onClick={() => go(1)}>
              Next
            </button>
          </div>
          <div className="g">
            <button type="button" className={clear === 'black' ? 'on' : ''} onClick={() => toggleClear('black')}>
              Black (B)
            </button>
            <button type="button" className={clear === 'white' ? 'on' : ''} onClick={() => toggleClear('white')}>
              White (W)
            </button>
            <button type="button" className={clear === 'logo' ? 'on' : ''} onClick={() => toggleClear('logo')}>
              Logo (L)
            </button>
          </div>
        </div>
      </section>

      <aside id="panel">
        <div className="tabs">
          <button type="button" className="on">
            Style
          </button>
          <button type="button">Edit</button>
          <button type="button" onClick={() => addBlankSlide()}>
            Add Slide
          </button>
          <button type="button">Team</button>
        </div>
        <p className="note">
          Style and freeform tools will port here from the HTML app (Phase 2). Shared slide rendering is already live in the
          preview and strip.
        </p>
        {item && (
          <>
            <h3>{item.t}</h3>
            <p className="note">
              {item.k} · {parts.length} slide{parts.length === 1 ? '' : 's'}
            </p>
          </>
        )}
      </aside>
    </main>
  );
}
