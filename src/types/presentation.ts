export type SlideType = "text";
export type TextAlignment = "left" | "center" | "right";
export type ScreenMode = "slide" | "black" | "white";
export type ServiceItemType = "welcome" | "song" | "bible" | "sermon" | "announcement" | "other";

export interface Slide {
  id: string;
  type: SlideType;
  text: string;
  background: string;
  textAlign: TextAlignment;
  fontSize: number;
}

export interface ServiceItem {
  id: string;
  title: string;
  type: ServiceItemType;
  slides: Slide[];
}

export interface Service {
  id: string;
  title: string;
  date: string;
  items: ServiceItem[];
}

export interface ServiceSummary {
  id: string;
  title: string;
  date: string;
}

export interface LiveSlide {
  serviceId: string;
  serviceTitle: string;
  itemId: string;
  itemTitle: string;
  itemType: ServiceItemType;
  itemSlidePosition: number;
  itemSlideTotal: number;
  overallPosition: number;
  overallTotal: number;
  slide: Slide;
}

export interface PresentationUpdate {
  liveSlide: LiveSlide | null;
  screenMode: ScreenMode;
}