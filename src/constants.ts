export const DEFAULT_EXCEL_NAME = "Untitled";

export const EXCEL_EXTENSION = "xlsx";
export const EXCEL_EXTENSIONS = ['.xlsx', '.xls', '.xlsb', '.xlsm'];

export const EXCEL_VIEW_EL_NAME = "excel-embedded-container";
export const WORKBOOK_EL_NAME = "excel-embedded-workbook";
export const WORKSHEET_EL_NAME = "excel-embedded-worksheet";

export const EMBED_LINK_EXCEL_REGEX = /(?:!?\[\[)?([^#]+\.xls[xmb]?)(?:#([^!|]+)?)?(?:!([^|]+))?(?:\|([^|\]]+))?(?:\]\])?/
export const EXCEL_RANGE_A1_REGEX = /(?:\$)?([a-zA-Z]+)(?:\$)?([0-9]+)(?:[:：](?:\$)?([a-zA-Z]+)(?:\$)?([0-9]+))?/
export const EXCEL_RANGE_R1C1_REGEX = /(?:[rR])(\[?[0-9]+\]?)(?:[cC])(\[?[0-9]+\]?)(?:[:：](?:[rR])(\[?[0-9]+\]?)(?:[cC])(\[?[0-9]+\]?))?/

export const EMBED_SIZE_REGEX = /(\d+)(?:\D(\d+))?/
export const EMBED_SIZE_REGEX_STRICT = /^(\d+)(?:\D(\d+))?$/


/**
 * Matches an extension with a leading period.
 * @example
 * .loom
 */
export const EXTENSION_REGEX = new RegExp(/\.[a-z]*$/);

/**
 * Matches all wiki links
 * @example
 * [[my-file]]
 * @example
 * [[my-file|alias]]
 * @example
 * [[my-file|]]
 */
export const WIKI_LINK_REGEX = new RegExp(/\[\[([^|\]]+)(\|([\w-]*))?\]\]/g);

export const removeDotFromExtensions = (extensions: string[]): string[] => {
    return extensions.map(extension => extension.startsWith('.') ? extension.substring(1) : extension);
};