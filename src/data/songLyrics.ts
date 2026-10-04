import type { Slide } from "../types/presentation";

interface ParsedSection {
  label: string | null;
  text: string;
}

export type ReflowMode = "markers" | "blank-lines" | "four-lines";

export function parseSongLyrics(lyrics: string, mode: ReflowMode = "markers"): ParsedSection[] {
  const sections: ParsedSection[] = [];
  let label: string | null = null;
  let lines: string[] = [];
  let lineCount = 0;

  function flush() {
    const text = lines.join("\n").trim();
    if (text) sections.push({ label, text });
    lines = [];
    lineCount = 0;
  }

  for (const sourceLine of lyrics.replace(/\r\n?/g, "\n").split("\n")) {
    const line = sourceLine.replace(/^\s*#{1,6}\s?/, "").trimEnd();
    const heading = line.trim().match(/^["']?\[([^\]]+)\]["']?$/);
    if (heading) {
      flush();
      label = heading[1].trim();
    } else if (/^\s*---+\s*$/.test(line)) {
      flush();
    } else if (mode === "blank-lines" && !line.trim()) {
      flush();
    } else if (mode === "four-lines" && !line.trim()) {
      continue;
    } else {
      lines.push(line);
      lineCount += 1;
      if (mode === "four-lines" && lineCount >= 4) {
        flush();
      }
    }
  }
  flush();
  return sections;
}

export function songSlides(lyrics: string, mode: ReflowMode = "markers"): Slide[] {
  return parseSongLyrics(lyrics, mode).map(({ label, text }) => ({
    id: crypto.randomUUID(),
    type: "text",
    text,
    background: "#18231f",
    textAlign: "center",
    fontSize: 52,
    sectionLabel: label,
  }));
}
