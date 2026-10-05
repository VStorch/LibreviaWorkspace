/**
 * Caps the schemas repeat. They keep an absurd input from freezing the UI; they do not restrict
 * legitimate use.
 */

export const MAX_COLOR_LENGTH = 32
export const MAX_CSS_VALUE_LENGTH = 16
export const MAX_ID_LENGTH = 120
export const MAX_FIELD_LENGTH = 100
export const MAX_NAME_LENGTH = 200
export const MAX_FILE_NAME_LENGTH = 255
export const MAX_PATH_LENGTH = 4096

export const MAX_FONT_FAMILY_LENGTH = 100
export const MAX_FONT_FAMILIES = 4000
export const MAX_USER_TEMPLATES = 1000
export const MAX_IMAGE_SIDE_PX = 4000

/** The `docProps/` fields: the description is free text, the others fit on one line. */
export const MAX_PROPERTY_LENGTH = 2_000
export const MAX_DESCRIPTION_LENGTH = 100_000

/** Word stores the start number in a 16-bit integer. */
export const MAX_START_NUMBER = 32767

/** Word does not split a page into more than 45 columns. */
export const MAX_SECTION_COLUMNS = 45
