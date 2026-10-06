import type { Slide, Song, SongArrangementPreset, SongSection, SongSectionType } from "../types/presentation";

export type ReflowMode = "markers" | "blank-lines" | "four-lines";

export interface SectionMeta {
  type: SongSectionType;
  label: string;
  color: string;
  bgRgba: string;
}

export const SECTION_PALETTES: Record<SongSectionType, SectionMeta> = {
  verse: { type: "verse", label: "V", color: "#388bfd", bgRgba: "rgba(56, 139, 253, 0.16)" },
  chorus: { type: "chorus", label: "C", color: "#a371f7", bgRgba: "rgba(163, 113, 247, 0.16)" },
  bridge: { type: "bridge", label: "B", color: "#f0883e", bgRgba: "rgba(240, 136, 62, 0.16)" },
  tag: { type: "tag", label: "T", color: "#f778ba", bgRgba: "rgba(247, 120, 186, 0.16)" },
  intro: { type: "intro", label: "Intro", color: "#3fb950", bgRgba: "rgba(63, 185, 80, 0.16)" },
  outro: { type: "outro", label: "Outro", color: "#2ea043", bgRgba: "rgba(46, 160, 67, 0.16)" },
  vamp: { type: "vamp", label: "Vamp", color: "#d29922", bgRgba: "rgba(210, 153, 34, 0.16)" },
  "pre-chorus": { type: "pre-chorus", label: "PC", color: "#db61a2", bgRgba: "rgba(219, 97, 162, 0.16)" },
  blank: { type: "blank", label: "Blank", color: "#8b949e", bgRgba: "rgba(139, 148, 158, 0.16)" },
  other: { type: "other", label: "Sec", color: "#79c0ff", bgRgba: "rgba(121, 192, 255, 0.16)" },
};

export function detectSectionType(name: string): SongSectionType {
  const clean = name.toLowerCase().trim();
  if (clean.includes("verse") || /^v\d+/i.test(clean)) return "verse";
  if (clean.includes("pre-chorus") || clean.includes("pre chorus")) return "pre-chorus";
  if (clean.includes("chorus") || /^c\d+/i.test(clean)) return "chorus";
  if (clean.includes("bridge") || /^b\d+/i.test(clean)) return "bridge";
  if (clean.includes("tag") || clean.includes("ending")) return "tag";
  if (clean.includes("intro")) return "intro";
  if (clean.includes("outro")) return "outro";
  if (clean.includes("vamp") || clean.includes("interlude") || clean.includes("instrumental")) return "vamp";
  if (clean.includes("blank") || clean.includes("clear") || clean.includes("pad")) return "blank";
  return "other";
}

export function detectSectionBadge(name: string, type: SongSectionType): string {
  const match = name.match(/\d+/);
  const num = match ? match[0] : "";
  switch (type) {
    case "verse":
      return num ? `V${num}` : "V";
    case "chorus":
      return num ? `C${num}` : "C";
    case "bridge":
      return num ? `B${num}` : "B";
    case "pre-chorus":
      return num ? `PC${num}` : "PC";
    case "tag":
      return "Tag";
    case "intro":
      return "Intro";
    case "outro":
      return "Outro";
    case "vamp":
      return "Vamp";
    case "blank":
      return "Blank";
    default:
      return name.slice(0, 3).toUpperCase();
  }
}

export function parseLyricsToSections(lyrics: string, mode: ReflowMode = "markers"): SongSection[] {
  const sections: SongSection[] = [];
  let currentName: string = "Verse 1";
  let currentLines: string[] = [];
  let sectionIndex = 1;

  function flush() {
    const text = currentLines.join("\n").trim();
    if (text || currentName) {
      const type = detectSectionType(currentName);
      const label = detectSectionBadge(currentName, type);
      const meta = SECTION_PALETTES[type] ?? SECTION_PALETTES.other;
      
      // Split section text into slides if markers exist or long text
      const rawSlides = text
        ? text.split(/\n\s*---+\s*\n/).map((s) => s.trim()).filter(Boolean)
        : [];
      
      const slides = rawSlides.length > 0 ? rawSlides : (text ? [text] : []);

      sections.push({
        id: `sec-${type}-${sectionIndex++}`,
        type,
        name: currentName,
        label,
        lines: text,
        slides,
        color: meta.color,
      });
    }
    currentLines = [];
  }

  for (const sourceLine of lyrics.replace(/\r\n?/g, "\n").split("\n")) {
    const line = sourceLine.replace(/^\s*#{1,6}\s?/, "").trimEnd();
    const heading = line.trim().match(/^["']?\[([^\]]+)\]["']?$/);
    if (heading) {
      if (currentLines.length > 0) flush();
      currentName = heading[1].trim();
    } else if (/^\s*---+\s*$/.test(line) && mode !== "markers") {
      flush();
    } else {
      currentLines.push(line);
    }
  }
  flush();

  return sections.length > 0 ? sections : [{
    id: "sec-v1",
    type: "verse",
    name: "Verse 1",
    label: "V1",
    lines: lyrics.trim(),
    slides: [lyrics.trim()],
    color: SECTION_PALETTES.verse.color,
  }];
}

export function generateSlidesFromArrangement(
  sections: SongSection[],
  sequence: string[],
  baseStyle?: Partial<Slide>
): Slide[] {
  const sectionMap = new Map(sections.map((sec) => [sec.id, sec]));
  const slides: Slide[] = [];

  sequence.forEach((sectionId, seqIndex) => {
    const section = sectionMap.get(sectionId);
    if (!section) return;

    const sectionSlides = section.slides && section.slides.length > 0
      ? section.slides
      : (section.lines.trim() ? section.lines.split(/\n\s*---+\s*\n/).map((s) => s.trim()).filter(Boolean) : [""]);

    sectionSlides.forEach((slideText, subIndex) => {
      const partLabel = sectionSlides.length > 1
        ? `${section.name} (${subIndex + 1}/${sectionSlides.length})`
        : section.name;

      slides.push({
        id: `slide-${section.id}-${seqIndex}-${subIndex}`,
        type: "text",
        text: slideText,
        background: baseStyle?.background ?? "#14171a",
        textAlign: baseStyle?.textAlign ?? "center",
        fontSize: baseStyle?.fontSize ?? 54,
        fontWeight: baseStyle?.fontWeight,
        fontFamily: baseStyle?.fontFamily ?? "sans-serif",
        textTransform: baseStyle?.textTransform ?? "none",
        color: baseStyle?.color ?? "#ffffff",
        textShadow: baseStyle?.textShadow ?? true,
        sectionLabel: partLabel,
        sectionType: section.type,
        notes: baseStyle?.notes ?? null,
        backgroundMedia: baseStyle?.backgroundMedia ?? null,
        transition: baseStyle?.transition ?? "cut",
        transitionDuration: baseStyle?.transitionDuration ?? 300,
        themeId: baseStyle?.themeId ?? null,
      });
    });
  });

  return slides;
}

export function parseSongLyrics(lyrics: string, mode: ReflowMode = "markers"): Array<{ label: string | null; text: string }> {
  const sections = parseLyricsToSections(lyrics, mode);
  const result: Array<{ label: string | null; text: string }> = [];
  sections.forEach((sec) => {
    const list = sec.slides && sec.slides.length > 0 ? sec.slides : [sec.lines];
    list.forEach((t) => {
      if (t.trim()) result.push({ label: sec.name, text: t });
    });
  });
  return result;
}

export function songSlides(lyrics: string, mode: ReflowMode = "markers"): Slide[] {
  const sections = parseLyricsToSections(lyrics, mode);
  const defaultSeq = sections.map((sec) => sec.id);
  return generateSlidesFromArrangement(sections, defaultSeq);
}

// Curated library of Smart Songs with full arrangements and metadata
export const DEFAULT_SMART_SONGS: Song[] = [
  {
    id: "song-goodness-of-god",
    title: "Goodness of God",
    artist: "Jenn Johnson",
    author: "Bethel Music",
    copyright: "2019 Bethel Music Publishing",
    ccli: "7117726",
    key: "A",
    bpm: 63,
    timeSignature: "4/4",
    lyrics: `[Verse 1]
I love You Lord
Oh Your mercy never fails me
All my days
I've been held in Your hands
From the moment that I wake up
Until I lay my head
I will sing of the goodness of God

[Chorus]
All my life You have been faithful
All my life You have been so so good
With every breath that I am able
I will sing of the goodness of God

[Verse 2]
I love Your voice
You have led me through the fire
In darkest night
You are close like no other
I've known You as a father
I've known You as a friend
I have lived in the goodness of God

[Bridge]
Your goodness is running after
It's running after me
Your goodness is running after
It's running after me
With my life laid down
I'm surrendered now
I give You everything
Your goodness is running after
It's running after me

[Tag]
I will sing of the goodness of God`,
    sections: [
      {
        id: "gog-v1",
        type: "verse",
        name: "Verse 1",
        label: "V1",
        lines: "I love You Lord\nOh Your mercy never fails me\nAll my days\nI've been held in Your hands\n\nFrom the moment that I wake up\nUntil I lay my head\nI will sing of the goodness of God",
        slides: [
          "I love You Lord\nOh Your mercy never fails me\nAll my days\nI've been held in Your hands",
          "From the moment that I wake up\nUntil I lay my head\nI will sing of the goodness of God",
        ],
        color: SECTION_PALETTES.verse.color,
      },
      {
        id: "gog-c",
        type: "chorus",
        name: "Chorus",
        label: "C",
        lines: "All my life You have been faithful\nAll my life You have been so so good\nWith every breath that I am able\nI will sing of the goodness of God",
        slides: [
          "All my life You have been faithful\nAll my life You have been so so good\nWith every breath that I am able\nI will sing of the goodness of God",
        ],
        color: SECTION_PALETTES.chorus.color,
      },
      {
        id: "gog-v2",
        type: "verse",
        name: "Verse 2",
        label: "V2",
        lines: "I love Your voice\nYou have led me through the fire\nIn darkest night\nYou are close like no other\n\nI've known You as a father\nI've known You as a friend\nI have lived in the goodness of God",
        slides: [
          "I love Your voice\nYou have led me through the fire\nIn darkest night\nYou are close like no other",
          "I've known You as a father\nI've known You as a friend\nI have lived in the goodness of God",
        ],
        color: SECTION_PALETTES.verse.color,
      },
      {
        id: "gog-b",
        type: "bridge",
        name: "Bridge",
        label: "B",
        lines: "Your goodness is running after\nIt's running after me\nYour goodness is running after\nIt's running after me\n\nWith my life laid down\nI'm surrendered now\nI give You everything\nYour goodness is running after\nIt's running after me",
        slides: [
          "Your goodness is running after\nIt's running after me\nYour goodness is running after\nIt's running after me",
          "With my life laid down\nI'm surrendered now\nI give You everything\nYour goodness is running after\nIt's running after me",
        ],
        color: SECTION_PALETTES.bridge.color,
      },
      {
        id: "gog-t",
        type: "tag",
        name: "Tag",
        label: "Tag",
        lines: "I will sing of the goodness of God",
        slides: ["I will sing of the goodness of God"],
        color: SECTION_PALETTES.tag.color,
      },
      {
        id: "gog-blank",
        type: "blank",
        name: "Instrumental / Blank",
        label: "Blank",
        lines: "",
        slides: [""],
        color: SECTION_PALETTES.blank.color,
      },
    ],
    arrangements: [
      {
        id: "arr-gog-standard",
        name: "Standard Arrangement",
        sectionIds: ["gog-v1", "gog-c", "gog-v2", "gog-c", "gog-b", "gog-b", "gog-c", "gog-t"],
      },
      {
        id: "arr-gog-acoustic",
        name: "Acoustic Sunday",
        sectionIds: ["gog-v1", "gog-v2", "gog-c", "gog-b", "gog-c", "gog-t"],
      },
      {
        id: "arr-gog-extended",
        name: "Extended Worship",
        sectionIds: ["gog-v1", "gog-c", "gog-v2", "gog-c", "gog-b", "gog-b", "gog-c", "gog-c", "gog-t", "gog-blank"],
      },
    ],
    defaultArrangementId: "arr-gog-standard",
  },
  {
    id: "song-way-maker",
    title: "Way Maker",
    artist: "Sinach / Leeland",
    author: "Osinachi Kalu Okoro Egbu",
    copyright: "2016 Integrity Music",
    ccli: "7115744",
    key: "E",
    bpm: 68,
    timeSignature: "4/4",
    lyrics: `[Verse 1]
You are here moving in our midst
I worship You I worship You
You are here working in this place
I worship You I worship You

[Chorus]
Way Maker Miracle Worker
Promise Keeper Light in the darkness
My God that is who You are

[Verse 2]
You are here touching every heart
I worship You I worship You
You are here healing every heart
I worship You I worship You

[Bridge]
Even when I don't see it You're working
Even when I don't feel it You're working
You never stop You never stop working
You never stop You never stop working`,
    sections: [
      {
        id: "wm-v1",
        type: "verse",
        name: "Verse 1",
        label: "V1",
        lines: "You are here moving in our midst\nI worship You I worship You\nYou are here working in this place\nI worship You I worship You",
        slides: [
          "You are here moving in our midst\nI worship You I worship You",
          "You are here working in this place\nI worship You I worship You",
        ],
        color: SECTION_PALETTES.verse.color,
      },
      {
        id: "wm-c",
        type: "chorus",
        name: "Chorus",
        label: "C",
        lines: "Way Maker Miracle Worker\nPromise Keeper Light in the darkness\nMy God that is who You are",
        slides: [
          "Way Maker Miracle Worker\nPromise Keeper Light in the darkness\nMy God that is who You are",
        ],
        color: SECTION_PALETTES.chorus.color,
      },
      {
        id: "wm-v2",
        type: "verse",
        name: "Verse 2",
        label: "V2",
        lines: "You are here touching every heart\nI worship You I worship You\nYou are here healing every heart\nI worship You I worship You",
        slides: [
          "You are here touching every heart\nI worship You I worship You",
          "You are here healing every heart\nI worship You I worship You",
        ],
        color: SECTION_PALETTES.verse.color,
      },
      {
        id: "wm-b",
        type: "bridge",
        name: "Bridge",
        label: "B",
        lines: "Even when I don't see it You're working\nEven when I don't feel it You're working\nYou never stop You never stop working\nYou never stop You never stop working",
        slides: [
          "Even when I don't see it You're working\nEven when I don't feel it You're working\nYou never stop You never stop working",
        ],
        color: SECTION_PALETTES.bridge.color,
      },
      {
        id: "wm-tag",
        type: "tag",
        name: "Tag",
        label: "Tag",
        lines: "That is who You are\nThat is who You are",
        slides: ["That is who You are\nThat is who You are"],
        color: SECTION_PALETTES.tag.color,
      },
    ],
    arrangements: [
      {
        id: "arr-wm-standard",
        name: "Standard Arrangement",
        sectionIds: ["wm-v1", "wm-c", "wm-v2", "wm-c", "wm-b", "wm-b", "wm-c", "wm-tag"],
      },
    ],
    defaultArrangementId: "arr-wm-standard",
  },
  {
    id: "song-amazing-grace",
    title: "Amazing Grace (My Chains Are Gone)",
    artist: "Chris Tomlin",
    author: "John Newton / Chris Tomlin / Louie Giglio",
    copyright: "2006 worshiptogether.com songs",
    ccli: "4768434",
    key: "G",
    bpm: 70,
    timeSignature: "3/4",
    lyrics: `[Verse 1]
Amazing grace how sweet the sound
That saved a wretch like me
I once was lost but now I'm found
Was blind but now I see

[Verse 2]
'Twas grace that taught my heart to fear
And grace my fears relieved
How precious did that grace appear
The hour I first believed

[Chorus]
My chains are gone I've been set free
My God my Savior has ransomed me
And like a flood His mercy reigns
Unending love amazing grace

[Verse 3]
The Lord has promised good to me
His word my hope secures
He will my shield and portion be
As long as life endures`,
    sections: [
      {
        id: "ag-v1",
        type: "verse",
        name: "Verse 1",
        label: "V1",
        lines: "Amazing grace how sweet the sound\nThat saved a wretch like me\nI once was lost but now I'm found\nWas blind but now I see",
        slides: ["Amazing grace how sweet the sound\nThat saved a wretch like me\nI once was lost but now I'm found\nWas blind but now I see"],
        color: SECTION_PALETTES.verse.color,
      },
      {
        id: "ag-v2",
        type: "verse",
        name: "Verse 2",
        label: "V2",
        lines: "'Twas grace that taught my heart to fear\nAnd grace my fears relieved\nHow precious did that grace appear\nThe hour I first believed",
        slides: ["'Twas grace that taught my heart to fear\nAnd grace my fears relieved\nHow precious did that grace appear\nThe hour I first believed"],
        color: SECTION_PALETTES.verse.color,
      },
      {
        id: "ag-c",
        type: "chorus",
        name: "Chorus",
        label: "C",
        lines: "My chains are gone I've been set free\nMy God my Savior has ransomed me\nAnd like a flood His mercy reigns\nUnending love amazing grace",
        slides: ["My chains are gone I've been set free\nMy God my Savior has ransomed me\nAnd like a flood His mercy reigns\nUnending love amazing grace"],
        color: SECTION_PALETTES.chorus.color,
      },
      {
        id: "ag-v3",
        type: "verse",
        name: "Verse 3",
        label: "V3",
        lines: "The Lord has promised good to me\nHis word my hope secures\nHe will my shield and portion be\nAs long as life endures",
        slides: ["The Lord has promised good to me\nHis word my hope secures\nHe will my shield and portion be\nAs long as life endures"],
        color: SECTION_PALETTES.verse.color,
      },
    ],
    arrangements: [
      {
        id: "arr-ag-standard",
        name: "Standard Flow",
        sectionIds: ["ag-v1", "ag-v2", "ag-c", "ag-v3", "ag-c", "ag-c"],
      },
    ],
    defaultArrangementId: "arr-ag-standard",
  },
  {
    id: "song-living-hope",
    title: "Living Hope",
    artist: "Phil Wickham",
    author: "Phil Wickham / Brian Johnson",
    copyright: "2017 Phil Wickham Music / Bethel Music Publishing",
    ccli: "7106807",
    key: "Eb",
    bpm: 71,
    timeSignature: "4/4",
    lyrics: `[Verse 1]
How great the chasm that lay between us
How high the mountain I could not climb
In desperation I turned to heaven
And spoke Your name into the night

[Chorus]
Hallelujah praise the One who set me free
Hallelujah death has lost its grip on me
You have broken every chain
There's salvation in Your name
Jesus Christ my living hope

[Verse 2]
Who could imagine so great a mercy
What heart could fathom such boundless grace
The God of ages stepped down from glory
To wear my sin and bear my shame`,
    sections: [
      {
        id: "lh-v1",
        type: "verse",
        name: "Verse 1",
        label: "V1",
        lines: "How great the chasm that lay between us\nHow high the mountain I could not climb\nIn desperation I turned to heaven\nAnd spoke Your name into the night",
        slides: ["How great the chasm that lay between us\nHow high the mountain I could not climb\nIn desperation I turned to heaven\nAnd spoke Your name into the night"],
        color: SECTION_PALETTES.verse.color,
      },
      {
        id: "lh-c",
        type: "chorus",
        name: "Chorus",
        label: "C",
        lines: "Hallelujah praise the One who set me free\nHallelujah death has lost its grip on me\nYou have broken every chain\nThere's salvation in Your name\nJesus Christ my living hope",
        slides: ["Hallelujah praise the One who set me free\nHallelujah death has lost its grip on me\nYou have broken every chain\nThere's salvation in Your name\nJesus Christ my living hope"],
        color: SECTION_PALETTES.chorus.color,
      },
      {
        id: "lh-v2",
        type: "verse",
        name: "Verse 2",
        label: "V2",
        lines: "Who could imagine so great a mercy\nWhat heart could fathom such boundless grace\nThe God of ages stepped down from glory\nTo wear my sin and bear my shame",
        slides: ["Who could imagine so great a mercy\nWhat heart could fathom such boundless grace\nThe God of ages stepped down from glory\nTo wear my sin and bear my shame"],
        color: SECTION_PALETTES.verse.color,
      },
    ],
    arrangements: [
      {
        id: "arr-lh-standard",
        name: "Standard Flow",
        sectionIds: ["lh-v1", "lh-c", "lh-v2", "lh-c", "lh-c"],
      },
    ],
    defaultArrangementId: "arr-lh-standard",
  },
];
