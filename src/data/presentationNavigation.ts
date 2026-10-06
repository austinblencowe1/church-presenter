import type { LiveSlide, Service, ServiceItem, Slide } from "../types/presentation";

export function getItemSequenceSlides(item: ServiceItem): Slide[] {
  if (item.slides && item.slides.length > 0) return item.slides;
  return [];
}

export function getSlideSequence(service: Service): LiveSlide[] {
  const sequence: LiveSlide[] = [];
  const total = service.items.reduce((count, item) => count + getItemSequenceSlides(item).length, 0);

  for (const item of service.items) {
    const slides = getItemSequenceSlides(item);
    slides.forEach((slide, index) => {
      sequence.push({
        serviceId: service.id,
        serviceTitle: service.title,
        itemId: item.id,
        itemTitle: item.title,
        itemType: item.type,
        itemSlidePosition: index + 1,
        itemSlideTotal: slides.length,
        overallPosition: sequence.length + 1,
        overallTotal: total,
        slide,
      });
    });
  }

  return sequence;
}

export function sameLiveSlide(first: LiveSlide | null, second: LiveSlide | null): boolean {
  return first?.serviceId === second?.serviceId && first?.slide.id === second?.slide.id;
}