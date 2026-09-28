export type SlideType = "text";
export type TextAlignment = "left" | "center" | "right";
export type ScreenMode = "slide" | "black" | "white";

export interface Slide {
  id: string;
  type: SlideType;
  text: string;
  background: string;
  textAlign: TextAlignment;
  fontSize: number;
}

export interface Presentation {
  id: string;
  title: string;
  slides: Slide[];
}