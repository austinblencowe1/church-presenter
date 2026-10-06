import type { Slide, SlideTheme } from "../types/presentation";

export interface PresetFont {
  readonly value: string;
  readonly label: string;
  readonly stack: string;
}

/** Built-in choices. These keep working for slides saved before custom fonts existed. */
export const PRESET_FONTS: readonly PresetFont[] = [
  { value: "sans-serif", label: "Sans (system default)", stack: '"Segoe UI Variable", "Segoe UI", ui-sans-serif, -apple-system, BlinkMacSystemFont, "SF Pro Display", Roboto, Arial, sans-serif' },
  { value: "serif", label: "Serif (Georgia)", stack: 'Georgia, "Times New Roman", ui-serif, "New York", Cambria, serif' },
  { value: "display", label: "Rounded / Display", stack: '"Outfit", "Segoe UI Variable Display", "Trebuchet MS", "SF Pro Rounded", system-ui, sans-serif' },
  { value: "mono", label: "Monospace", stack: 'Consolas, ui-monospace, "SF Mono", Menlo, Monaco, "Courier New", monospace' },
];

/**
 * Fonts commonly found on Windows / macOS / Linux. Nothing is downloaded: each name is
 * tested against the fonts actually installed on this computer and only matches are shown.
 */
const CANDIDATE_FONTS: readonly string[] = [
  "Arial", "Arial Black", "Arial Narrow", "Arial Rounded MT Bold", "Avenir", "Avenir Next", "Bahnschrift",
  "Baskerville", "Bell MT", "Book Antiqua", "Bookman Old Style", "Calibri", "Cambria", "Candara",
  "Century Gothic", "Century Schoolbook", "Comic Sans MS", "Consolas", "Constantia", "Copperplate",
  "Corbel", "Courier New", "Didot", "DejaVu Sans", "DejaVu Serif", "Franklin Gothic Medium", "Futura",
  "Garamond", "Geneva", "Georgia", "Gill Sans", "Gill Sans MT", "Helvetica", "Helvetica Neue", "Hoefler Text",
  "Impact", "Lato", "Lucida Bright", "Lucida Grande", "Lucida Sans Unicode", "Merriweather", "Montserrat",
  "Noto Sans", "Noto Serif", "Open Sans", "Optima", "Oswald", "Palatino", "Palatino Linotype", "Playfair Display",
  "Poppins", "Rockwell", "Roboto", "Segoe Print", "Segoe Script", "Segoe UI", "Segoe UI Black",
  "Segoe UI Light", "Segoe UI Semibold", "Source Sans Pro", "Tahoma", "Times New Roman", "Trebuchet MS",
  "Verdana", "Cinzel", "Inter", "Bebas Neue", "Ubuntu",
];

const SAMPLE = "mmmmmmmmmmlli0O1WwQq";
const BASELINES = ["monospace", "sans-serif", "serif"] as const;
let installedCache: string[] | null = null;

/** Strip characters that could break out of a CSS font-family declaration. */
export function sanitizeFontName(name: string): string {
  return name.replace(/["'\\;{}<>]/g, "").replace(/\s+/g, " ").trim().slice(0, 80);
}

/** True when the browser actually has a font with this name (it renders differently to every fallback). */
export function isFontInstalled(name: string): boolean {
  const clean = sanitizeFontName(name);
  if (!clean || typeof document === "undefined") return false;
  const context = document.createElement("canvas").getContext("2d");
  if (!context) return false;
  return BASELINES.some((baseline) => {
    context.font = `72px ${baseline}`;
    const baselineWidth = context.measureText(SAMPLE).width;
    context.font = `72px "${clean}", ${baseline}`;
    return context.measureText(SAMPLE).width !== baselineWidth;
  });
}

export function getInstalledFonts(): string[] {
  if (!installedCache) installedCache = CANDIDATE_FONTS.filter(isFontInstalled);
  return installedCache;
}

export function isPresetFont(value: string | undefined): boolean {
  return PRESET_FONTS.some((font) => font.value === value);
}

/** The CSS font-family value for a stored slide font (preset key or a real font name). */
export function fontStackFor(value: string | undefined): string {
  const preset = PRESET_FONTS.find((font) => font.value === (value ?? "sans-serif"));
  if (preset) return preset.stack;
  const clean = sanitizeFontName(value ?? "");
  return clean ? `"${clean}", ${PRESET_FONTS[0].stack}` : PRESET_FONTS[0].stack;
}

/** Slide properties that make up "typography" and are copied by Apply to all / themes. */
export const TYPOGRAPHY_KEYS = ["fontFamily", "fontSize", "fontWeight", "textAlign", "textTransform", "color", "textShadow"] as const;

export function typographyOf(slide: Slide): Partial<Slide> {
  const picked: Partial<Slide> = {};
  for (const key of TYPOGRAPHY_KEYS) {
    if (slide[key] !== undefined) (picked as Record<string, unknown>)[key] = slide[key];
  }
  return picked;
}

/** Font settings a saved theme carries. Only defined values, so older themes never blank a slide's font. */
export function typographyFromTheme(theme: SlideTheme): Partial<Slide> {
  const picked: Partial<Slide> = {};
  if (theme.fontFamily !== undefined) picked.fontFamily = theme.fontFamily;
  if (theme.textTransform !== undefined) picked.textTransform = theme.textTransform;
  if (theme.color !== undefined) picked.color = theme.color;
  return picked;
}
