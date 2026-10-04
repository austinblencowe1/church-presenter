import type { LiveSlide, Service, ServiceItem, Slide } from "../types/presentation";

export function getItemSequenceSlides(item: ServiceItem): Slide[] {
  const groups = item.groups ?? [];
  const arrangements = item.arrangements ?? [];
  const arrangement = arrangements.find((entry) => entry.id === item.activeArrangementId) ?? arrangements[0];
  if (!arrangement || groups.length === 0) return item.slides;
  const slidesById = new Map(item.slides.map((slide) => [slide.id, slide]));
  const ordered = arrangement.groupIds.flatMap((groupId) => {
    const group = groups.find((entry) => entry.id === groupId);
    return group?.slideIds.map((slideId) => slidesById.get(slideId)).filter((slide): slide is Slide => Boolean(slide)) ?? [];
  });
  const included = new Set(ordered.map((slide) => slide.id));
  return [...ordered, ...item.slides.filter((slide) => !included.has(slide.id))];
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