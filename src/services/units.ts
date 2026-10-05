/** CSS fixes 96 px per inch; OOXML measures in twips, a twentieth of a point. */

export const MM_PER_INCH = 25.4
export const PT_PER_INCH = 72
export const CSS_PX_PER_INCH = 96
export const TWIPS_PER_INCH = 1440

/** Half an inch (720 twips): Word's indent step. */
export const INDENT_STEP_MM = MM_PER_INCH / 2

export function mmToInches(mm: number): number {
  return mm / MM_PER_INCH
}

export function mmToPx(mm: number): number {
  return (mm * CSS_PX_PER_INCH) / MM_PER_INCH
}

export function pxToMm(px: number): number {
  return (px * MM_PER_INCH) / CSS_PX_PER_INCH
}

export function pxToPt(px: number): number {
  return px * (PT_PER_INCH / CSS_PX_PER_INCH)
}

export function twipsToMm(twips: number): number {
  return (twips * MM_PER_INCH) / TWIPS_PER_INCH
}

export function twipsToPx(twips: number): number {
  return (twips * CSS_PX_PER_INCH) / TWIPS_PER_INCH
}
