/**
 * When the property pane has to (re)read the library's columns, and what it
 * should say about the tag column dropdowns.
 *
 * This lives in `core/` because getting it wrong is invisible until someone
 * opens the property pane on a real tenant: the dropdowns simply stay greyed
 * out. The rules are small enough to state exactly, so they are stated here and
 * unit tested, and the web part only wires them up.
 */

export interface ITagFieldCacheState {
  /** Identifies the library the pane is configured against right now. */
  libraryKey: string;
  /** Library the cached columns were read from. */
  loadedKey: string;
  /** A read has completed for `loadedKey`, successfully or not. */
  loaded: boolean;
  /** A read is in flight. */
  loading: boolean;
  /** A library is configured at all. */
  hasLibrary: boolean;
  /** Number of columns the last read produced. */
  optionCount: number;
}

/**
 * True when the columns have to be read.
 *
 * The `loadedKey` comparison is the important one: the property pane re-renders
 * on every keystroke and every toggle, and a cache filled for the library
 * currently configured must survive that. Re-reading only when the library
 * actually changed is what keeps the dropdowns usable while other settings are
 * being edited.
 */
export function shouldLoadTagFields(state: ITagFieldCacheState): boolean {
  if (!state.hasLibrary || state.loading) {
    return false;
  }
  return !state.loaded || state.loadedKey !== state.libraryKey;
}

/** True when the dropdowns can be operated. */
export function canPickTagField(state: ITagFieldCacheState): boolean {
  return state.loaded && !state.loading && state.optionCount > 0;
}

/** What the pane tells the reader about the state of the dropdowns. */
export type TagFieldStatus = 'needs-library' | 'loading' | 'none-found' | 'ready';

export function tagFieldStatus(state: ITagFieldCacheState): TagFieldStatus {
  if (!state.hasLibrary) {
    return 'needs-library';
  }
  if (state.loading || !state.loaded) {
    return 'loading';
  }
  return state.optionCount > 0 ? 'ready' : 'none-found';
}
