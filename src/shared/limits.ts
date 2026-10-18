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
/** Word's dialog has no limit; a list longer than this is no longer something a person keeps. */
export const MAX_CUSTOM_PROPERTIES = 500
export const MAX_DESCRIPTION_LENGTH = 100_000

/** Word stores the start number in a 16-bit integer. */
export const MAX_START_NUMBER = 32767

/** Word does not split a page into more than 45 columns. */
export const MAX_SECTION_COLUMNS = 45

/** A long chapter; a whole book would cost money on a paid service and stall a local model. */
export const MAX_AI_SELECTION_LENGTH = 20_000
/** The open document goes along as context; past this, only the text around the selection. */
export const MAX_AI_DOCUMENT_LENGTH = 200_000
/** Ollama reads 4096 tokens by default and drops the rest without a word: about 6000 characters fit. */
export const DEFAULT_LOCAL_AI_CONTEXT_LENGTH = 6_000
export const MAX_AI_INSTRUCTION_LENGTH = 2_000
export const MAX_AI_RESPONSE_LENGTH = 100_000
/** Twenty exchanges; the oldest leave the panel first. */
export const MAX_AI_CHAT_MESSAGES = 40
export const MAX_AI_CHAT_MESSAGE_LENGTH = MAX_AI_INSTRUCTION_LENGTH
export const MAX_AI_KEY_LENGTH = 500
export const MAX_AI_URL_LENGTH = 2_000
export const MAX_AI_PROVIDERS = 20
export const MAX_AI_MODELS = 1_000
export const MAX_AI_STREAM_LINE_LENGTH = 1_000_000
export const MAX_AI_AGENT_ARGS_LENGTH = 1_000
/** The chat's document tools: what the model may read and propose in one reply. */
export const MAX_AI_PARAGRAPHS = 50_000
export const MAX_AI_TOOL_ROUNDS = 8
export const MAX_AI_TOOL_CALLS = 10
/** The JSON of one call's arguments: a find and a replace at their limits fit with room. */
export const MAX_AI_TOOL_ARGUMENTS_LENGTH = 20_000
export const MAX_AI_PROPOSALS = 30
export const MAX_AI_FIND_LENGTH = 2_000
export const MAX_AI_REPLACE_LENGTH = 10_000
export const MAX_AI_COMMENT_LENGTH = 2_000
export const MAX_AI_REASON_LENGTH = 500
export const MAX_AI_SEARCH_QUERY_LENGTH = 200
export const MAX_AI_SEARCH_RESULTS = 20
export const MAX_AI_READ_PARAGRAPHS = 50
/** The local MCP server: the name a client gives itself, shown as the author of its changes. */
export const MAX_MCP_CLIENT_NAME_LENGTH = 64
export const MAX_MCP_CLIENTS = 2
/** The spreadsheet assistant: the cells sent as context and how many rows one fill writes. */
export const MAX_AI_SHEET_CONTEXT_LENGTH = MAX_AI_DOCUMENT_LENGTH
export const MAX_AI_FILL_ROWS = 500
/** Excel's own limit for the text of a formula. */
export const MAX_AI_FORMULA_LENGTH = 8_192
