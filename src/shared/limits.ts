/**
 * Tetos que se repetem nos esquemas. Existem para que uma entrada absurda não
 * trave a interface, não para restringir uso legítimo.
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

/** Os campos de `docProps/`: a descrição é texto livre, os outros cabem numa linha. */
export const MAX_PROPERTY_LENGTH = 2_000
export const MAX_DESCRIPTION_LENGTH = 100_000

/** O Word guarda o número inicial num inteiro de 16 bits. */
export const MAX_START_NUMBER = 32767

/** O Word não divide a página em mais de 45 colunas. */
export const MAX_SECTION_COLUMNS = 45
