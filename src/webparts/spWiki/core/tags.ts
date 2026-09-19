/**
 * Tag handling. A "tag" is a value stored in a column of the document library;
 * which column is a web part setting, so the code has to cope with the three
 * column types that people realistically use for this:
 *
 *   Text / Note  -> one string holding several tags, separated by ; or ,
 *   Choice       -> one value
 *   MultiChoice  -> several values
 *
 * Everything in this file is pure so it can be unit tested without SharePoint.
 */

export type TagFieldKind = 'text' | 'choice' | 'multiChoice';

export interface ITagFieldInfo {
  /** Internal name of the column — how the web part settings refer to it. */
  internalName: string;
  /**
   * Name the *list item* exposes the column under, which is not always the
   * internal name. See {@link toEntityPropertyName}.
   */
  entityPropertyName: string;
  /** Display name shown in the property pane and the tag editor. */
  title: string;
  kind: TagFieldKind;
  /** Allowed values for choice columns; empty for text columns. */
  choices: string[];
  /**
   * Choice columns: the column's "allow custom values" setting
   * (`FillInChoice`). SharePoint rejects a value outside `choices` unless this
   * is on, so it decides whether a tag can be typed in rather than picked.
   * Always true for text columns, which take anything.
   */
  allowFillIn: boolean;
}

/** True when a tag that is not among the column's choices can be saved. */
export function acceptsNewTags(field: ITagFieldInfo): boolean {
  return field.kind === 'text' || field.allowFillIn;
}

/**
 * SharePoint's own columns start with an underscore (`_UIVersionString`,
 * `_ComplianceTag`), so an internal name beginning with one is normally not a
 * column anybody would tag with.
 *
 * A column *created with a non-ASCII display name* looks exactly the same at
 * first glance: SharePoint derives the internal name by escaping each character
 * it cannot use, so 「タグ」 becomes `_x30bf__x30b0_`. Rejecting every name that
 * starts with an underscore therefore hid every Japanese-named column from the
 * property pane dropdowns — the columns people are most likely to have made.
 *
 * The escape is what tells them apart: `_x` followed by exactly four hex digits
 * and a closing underscore. No system column is named that way.
 */
export function isSystemFieldName(internalName: string): boolean {
  if (internalName.charAt(0) !== '_') {
    return false;
  }

  return !/^_x[0-9a-fA-F]{4}_/.test(internalName);
}

/**
 * The name a list item exposes a column under, which is what `$select`, the
 * item's own JSON and a write payload all have to use.
 *
 * It is the internal name for an ordinary column, but SharePoint cannot start
 * an entity property with an underscore, so it prefixes those with `OData_`.
 * A column *created with a Japanese display name* is exactly that case: 「業種」
 * is stored as `_x696d__x7a2e_`, and the item therefore carries it as
 * `OData__x696d__x7a2e_`.
 *
 * Using the internal name for those columns is why they read back empty and why
 * saving them answered:
 *
 *   HTTP 400 プロパティ '_x696d__x7a2e_' は型 'SP.Data.ExampleItem' に存在しません
 *
 * The field list gives the real name in `EntityPropertyName`; this is the rule
 * it follows, used when that value is missing.
 */
export function toEntityPropertyName(internalName: string): string {
  return internalName.charAt(0) === '_' ? 'OData_' + internalName : internalName;
}

/**
 * Applies the columns just edited on top of the ones the document already
 * carries, matching on internal name.
 *
 * The editor opens one column at a time, so what comes back out of it is a
 * subset. Merging rather than replacing is what keeps the other columns' tags
 * on screen — and out of the save payload, which only names the columns it is
 * given.
 */
export function mergeTagGroups(existing: ITagGroup[], edited: ITagGroup[]): ITagGroup[] {
  const merged: ITagGroup[] = [];

  for (let i: number = 0; i < existing.length; i++) {
    let replacement: ITagGroup | undefined;

    for (let e: number = 0; e < edited.length; e++) {
      if (edited[e].field.internalName === existing[i].field.internalName) {
        replacement = edited[e];
        break;
      }
    }

    merged.push(replacement || existing[i]);
  }

  return merged;
}

/**
 * One tag column together with the tags a document carries in it. Up to three
 * columns can be configured, which is why tags travel grouped everywhere.
 */
export interface ITagGroup {
  field: ITagFieldInfo;
  tags: string[];
}

/** The tags of every group as one de-duplicated list. */
export function flattenTagGroups(groups: ITagGroup[]): string[] {
  let all: string[] = [];

  for (let i: number = 0; i < groups.length; i++) {
    all = all.concat(groups[i].tags);
  }

  return normalizeTags(all);
}

/** ASCII and Japanese separators people actually type into a text column. */
const SEPARATORS: RegExp = /[;,、，；]/;

/** Trims whitespace. JavaScript's \s covers U+3000 IDEOGRAPHIC SPACE too. */
function trim(value: string): string {
  return value.replace(/^\s+|\s+$/g, '');
}

/** Drops empties and case-insensitive duplicates, keeping the first spelling. */
export function normalizeTags(tags: string[]): string[] {
  const out: string[] = [];
  const seen: { [lower: string]: boolean } = {};

  for (let i: number = 0; i < tags.length; i++) {
    const tag: string = trim(String(tags[i]));
    if (tag === '') {
      continue;
    }
    const key: string = tag.toLowerCase();
    if (seen[key]) {
      continue;
    }
    seen[key] = true;
    out.push(tag);
  }

  return out;
}

/**
 * Reads whatever the REST API returned for the tag column.
 *
 * Depending on the column type and the OData metadata level this is a string,
 * a plain array, or `{ results: [...] }`, so all three are accepted.
 */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
export function parseTagValue(value: any): string[] {
  if (value === undefined || value === null) {
    return [];
  }

  if (typeof value === 'string') {
    return normalizeTags(value.split(SEPARATORS));
  }

  if (Object.prototype.toString.call(value) === '[object Array]') {
    return normalizeTags(value);
  }

  if (value.results && Object.prototype.toString.call(value.results) === '[object Array]') {
    return normalizeTags(value.results);
  }

  return [];
}

/**
 * How a multi-choice value is written on the wire.
 *
 * SharePoint accepts a collection in two shapes and which one it wants depends
 * on the OData flavour of the request: `{ results: [...] }` is the older
 * verbose form, a plain array the one that goes with `odata=nometadata`. A
 * tenant that wants the other one answers HTTP 400 rather than saying which, so
 * the caller tries both.
 */
export type MultiChoiceFormat = 'results' | 'array';

/** Builds the value to send back to SharePoint for the given column type. */
export function formatTagValue(
  tags: string[],
  kind: TagFieldKind,
  multiChoiceFormat: MultiChoiceFormat = 'results'
  // The return type is the column type's wire format: a string, null, an array
  // or a wrapper object.
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
): any {
  const normalized: string[] = normalizeTags(tags);

  if (kind === 'multiChoice') {
    return multiChoiceFormat === 'array' ? normalized : { results: normalized };
  }

  if (kind === 'choice') {
    return normalized.length > 0 ? normalized[0] : null;
  }

  return normalized.join('; ');
}

/** True when any of the columns is written as a collection. */
export function hasMultiChoice(groups: ITagGroup[]): boolean {
  for (let i: number = 0; i < groups.length; i++) {
    if (groups[i].field.kind === 'multiChoice') {
      return true;
    }
  }
  return false;
}

/** Case-insensitive tag comparison. */
export function isSameTag(a: string, b: string): boolean {
  return trim(a).toLowerCase() === trim(b).toLowerCase();
}

/** True when the two tag sets have at least one tag in common. */
export function sharesTag(a: string[], b: string[]): boolean {
  for (let i: number = 0; i < a.length; i++) {
    for (let j: number = 0; j < b.length; j++) {
      if (isSameTag(a[i], b[j])) {
        return true;
      }
    }
  }
  return false;
}

/** Maps a SharePoint field type name onto the kind this web part supports. */
export function toTagFieldKind(typeAsString: string): TagFieldKind | undefined {
  switch (typeAsString) {
    case 'Text':
    case 'Note':
      return 'text';
    case 'Choice':
      return 'choice';
    case 'MultiChoice':
      return 'multiChoice';
    default:
      return undefined;
  }
}
