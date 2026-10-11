import { useMemo } from 'react';
import type { ServiceItem, SlidePart } from '../types/presentation';
import { renderSlideHtml } from '../data/slideRender';
import './SlideView.css';

interface Props {
  item: ServiceItem;
  part: SlidePart;
  editable?: boolean;
  selectedElementId?: string | null;
  stage?: boolean;
  mediaURLs?: Record<string, string>;
  className?: string;
}

/** Renders one slide via shared HTML renderer (editor + audience stay identical). */
export function SlideView({
  item,
  part,
  editable,
  selectedElementId,
  stage,
  mediaURLs,
  className,
}: Props) {
  const html = useMemo(
    () =>
      renderSlideHtml(item, part, {
        editable,
        selectedElementId,
        stage,
        mediaURLs,
      }),
    [item, part, editable, selectedElementId, stage, mediaURLs]
  );

  return (
    <div
      className={className ? `slide-view ${className}` : 'slide-view'}
      dangerouslySetInnerHTML={{ __html: html }}
    />
  );
}
