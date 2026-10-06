import { useRef } from "react";
import type { CSSProperties, PointerEvent, ReactNode } from "react";
import { fontStackFor } from "../data/fonts";
import type { ScreenMode, Slide, SlideObject } from "../types/presentation";

interface SlideCanvasProps {
  readonly slide: Slide;
  readonly screenMode: ScreenMode;
  readonly variant: "preview" | "thumbnail" | "stage" | "overview";
  readonly clearSlide?: boolean;
  readonly clearBackground?: boolean;
  readonly interactive?: boolean;
  readonly selectedObjectId?: string | null;
  readonly onObjectSelect?: (objectId: string) => void;
  readonly onObjectChange?: (objectId: string, update: Partial<SlideObject>) => void;
}

interface ObjectDrag {
  object: SlideObject;
  pointerX: number;
  pointerY: number;
  mode: "move" | "resize";
}

export function SlideCanvas({
  slide,
  screenMode,
  variant,
  clearSlide = false,
  clearBackground = false,
  interactive = false,
  selectedObjectId = null,
  onObjectSelect,
  onObjectChange,
}: SlideCanvasProps): ReactNode {
  const canvasRef = useRef<HTMLDivElement>(null);
  const dragRef = useRef<ObjectDrag | null>(null);
  const media = clearBackground ? "" : slide.backgroundMedia ?? "";
  const selectedFont = fontStackFor(slide.fontFamily);
  const transform = slide.textTransform ?? "none";
  const textColor = slide.color ?? "#ffffff";
  const shadow = slide.textShadow !== false ? "0 2px 10px rgba(0,0,0,0.7), 0 1px 3px rgba(0,0,0,0.9)" : "none";

  const style = {
    "--slide-background": clearBackground ? "#000000" : slide.background,
    "--slide-media": media ? `url(${JSON.stringify(media)})` : "none",
    "--slide-font-size": `${slide.fontSize}px`,
    "--slide-font-family": selectedFont,
    "--slide-font-weight": String(slide.fontWeight ?? 600),
    "--slide-text-transform": transform,
    "--slide-text-color": textColor,
    "--slide-text-shadow": shadow,
    "--slide-alignment": slide.textAlign,
    "--slide-transition-duration": `${slide.transitionDuration ?? 300}ms`,
  } as CSSProperties;

  function beginObjectPointer(event: PointerEvent<HTMLElement>, object: SlideObject, mode: "move" | "resize") {
    if (!interactive || !canvasRef.current) return;
    event.preventDefault();
    event.stopPropagation();
    onObjectSelect?.(object.id);
    dragRef.current = { object, pointerX: event.clientX, pointerY: event.clientY, mode };
    event.currentTarget.setPointerCapture(event.pointerId);
  }

  function moveObjectPointer(event: PointerEvent<HTMLDivElement>) {
    const drag = dragRef.current;
    const bounds = canvasRef.current?.getBoundingClientRect();
    if (!drag || !bounds || !onObjectChange) return;
    const deltaX = ((event.clientX - drag.pointerX) / bounds.width) * 100;
    const deltaY = ((event.clientY - drag.pointerY) / bounds.height) * 100;
    if (drag.mode === "move") {
      onObjectChange(drag.object.id, {
        x: Math.max(0, Math.min(100 - drag.object.width, drag.object.x + deltaX)),
        y: Math.max(0, Math.min(100 - drag.object.height, drag.object.y + deltaY)),
      });
    } else {
      onObjectChange(drag.object.id, {
        width: Math.max(4, Math.min(100 - drag.object.x, drag.object.width + deltaX)),
        height: Math.max(4, Math.min(100 - drag.object.y, drag.object.height + deltaY)),
      });
    }
    drag.pointerX = event.clientX;
    drag.pointerY = event.clientY;
    drag.object = {
      ...drag.object,
      ...(drag.mode === "move"
        ? { x: Math.max(0, Math.min(100 - drag.object.width, drag.object.x + deltaX)), y: Math.max(0, Math.min(100 - drag.object.height, drag.object.y + deltaY)) }
        : { width: Math.max(4, Math.min(100 - drag.object.x, drag.object.width + deltaX)), height: Math.max(4, Math.min(100 - drag.object.y, drag.object.height + deltaY)) }),
    };
  }

  function endObjectPointer() {
    dragRef.current = null;
  }

  return (
    <div ref={canvasRef} key={slide.id} className={`slide-canvas slide-canvas--${variant} slide-canvas--${screenMode} slide-canvas--transition-${slide.transition ?? "cut"}${interactive ? " slide-canvas--interactive" : ""}`} style={style} onPointerMove={moveObjectPointer} onPointerUp={endObjectPointer} onPointerCancel={endObjectPointer}>
      {screenMode === "slide" && media.startsWith("data:video/") && <video className="slide-canvas__media" src={media} autoPlay muted loop playsInline tabIndex={-1} />}
      {screenMode === "slide" && media && !media.startsWith("data:video/") && <div className="slide-canvas__media" aria-hidden="true" />}
      {screenMode === "slide" && media && <div className="slide-canvas__shade" aria-hidden="true" />}
      {screenMode === "slide" && !clearSlide && <div className="slide-canvas__text">{slide.text}</div>}
      {interactive && <div className="slide-canvas__safe-area" aria-hidden="true" />}
      {screenMode === "slide" && !clearSlide && (slide.objects ?? []).map((object) => {
        const objectStyle = {
          left: `${object.x}%`,
          top: `${object.y}%`,
          width: `${object.width}%`,
          height: `${object.height}%`,
          opacity: object.opacity,
          transform: `rotate(${object.rotation}deg)`,
          "--object-fill": object.fill,
          "--object-stroke": object.stroke,
          "--object-stroke-width": `${object.strokeWidth}px`,
          "--object-color": object.color,
          "--object-font-size": `${object.fontSize}px`,
          "--object-align": object.textAlign,
          "--object-scale": "var(--slide-scale, 1)",
        } as CSSProperties;
        return (
          <div
            key={object.id}
            className={`slide-canvas__object slide-canvas__object--${object.kind}${selectedObjectId === object.id ? " is-selected" : ""}`}
            style={objectStyle}
            onPointerDown={(event) => beginObjectPointer(event, object, "move")}
            onClick={() => interactive && onObjectSelect?.(object.id)}
            title={interactive ? "Drag to move" : undefined}
          >
            {object.kind === "text" && object.text}
            {object.kind === "image" && object.source && <img src={object.source} alt="" draggable={false} />}
            {object.kind === "video" && object.source && <video src={object.source} autoPlay muted loop playsInline />}
            {interactive && selectedObjectId === object.id && <button className="slide-canvas__resize-handle" aria-label="Resize object" onPointerDown={(event) => beginObjectPointer(event, object, "resize")} />}
          </div>
        );
      })}
    </div>
  );
}

