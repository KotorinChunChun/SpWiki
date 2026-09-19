/**
 * Colours for the tag chips.
 *
 * The hue is derived from the tag text rather than drawn at random, so a tag
 * keeps the same colour on every page and every reload — the point of colouring
 * them at all is that a reader recognises a tag by its colour. Backgrounds stay
 * pale and the foreground is picked as black or white by contrast.
 */

export interface ITagColor {
  background: string;
  foreground: string;
}

/** Pale backgrounds, alternating so neighbouring hues stay distinguishable. */
const SATURATIONS: number[] = [72, 58];
const LIGHTNESSES: number[] = [87, 80];

/** Stable, well-spread hash of a tag name. */
function hashTag(tag: string): number {
  let hash: number = 5381;

  for (let i: number = 0; i < tag.length; i++) {
    // `| 0` keeps it in 32-bit range instead of drifting into floats.
    hash = ((hash << 5) + hash + tag.charCodeAt(i)) | 0;
  }

  return Math.abs(hash);
}

/** Background and foreground for one tag. */
export function tagColor(tag: string): ITagColor {
  const hash: number = hashTag(normalize(tag));
  const hue: number = hash % 360;
  const saturation: number = SATURATIONS[(hash >> 9) % SATURATIONS.length];
  const lightness: number = LIGHTNESSES[(hash >> 17) % LIGHTNESSES.length];
  const rgb: number[] = hslToRgb(hue, saturation / 100, lightness / 100);

  return {
    background: 'hsl(' + hue + ', ' + saturation + '%, ' + lightness + '%)',
    foreground: relativeLuminance(rgb) > 0.5 ? '#1b1b1b' : '#ffffff'
  };
}

/** Case and surrounding space must not change the colour. */
function normalize(tag: string): string {
  return tag.replace(/^\s+|\s+$/g, '').toLowerCase();
}

function hslToRgb(hue: number, saturation: number, lightness: number): number[] {
  const c: number = (1 - Math.abs(2 * lightness - 1)) * saturation;
  const x: number = c * (1 - Math.abs(((hue / 60) % 2) - 1));
  const m: number = lightness - c / 2;

  let rgb: number[];
  if (hue < 60) {
    rgb = [c, x, 0];
  } else if (hue < 120) {
    rgb = [x, c, 0];
  } else if (hue < 180) {
    rgb = [0, c, x];
  } else if (hue < 240) {
    rgb = [0, x, c];
  } else if (hue < 300) {
    rgb = [x, 0, c];
  } else {
    rgb = [c, 0, x];
  }

  return [rgb[0] + m, rgb[1] + m, rgb[2] + m];
}

/** WCAG relative luminance, used only to choose between black and white text. */
function relativeLuminance(rgb: number[]): number {
  const channels: number[] = [];

  for (let i: number = 0; i < 3; i++) {
    const value: number = rgb[i];
    channels.push(value <= 0.03928 ? value / 12.92 : Math.pow((value + 0.055) / 1.055, 2.4));
  }

  return 0.2126 * channels[0] + 0.7152 * channels[1] + 0.0722 * channels[2];
}
