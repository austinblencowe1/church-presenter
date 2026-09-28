import type { CSSProperties } from "react";
import type { ScreenMode, Slide } from "../types/presentation";

interface SlideCanvasProps {
  slide: Slide;
  screenMode: ScreenMode;
  variant: "preview" | "thumbnail" | "stage";
}

export function SlideCanvas({ slide, screenMode, variant }: SlideCanvasProps) {
  const style = {
    "--slide-background": slide.background,
    "--slide-font-size": `${slide.fontSize}px`,
    "--slide-alignment": slide.textAlign,
  } as CSSProperties;

  return (
    <div className={`slide-canvas slide-canvas--${variant} slide-canvas--${screenMode}`} style={style}>
      {screenMode === "slide" && <div className="slide-canvas__text">{slide.text}</div>}
    </div>
  );
}