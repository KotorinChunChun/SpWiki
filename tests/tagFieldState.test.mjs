import assert from 'node:assert/strict';
import { describe, it } from 'node:test';

import { core } from './setup.mjs';

const { canPickTagField, shouldLoadTagFields, tagFieldStatus } = core;

/** The state right after the columns of `|ドキュメント` were read. */
const loaded = (overrides) => ({
  libraryKey: '|ドキュメント',
  loadedKey: '|ドキュメント',
  loaded: true,
  loading: false,
  hasLibrary: true,
  optionCount: 3,
  ...overrides
});

describe('shouldLoadTagFields', () => {
  it('reads the columns when nothing has been read yet', () => {
    assert.equal(shouldLoadTagFields(loaded({ loaded: false, loadedKey: '', optionCount: 0 })), true);
  });

  it('does not read them again for the library already cached', () => {
    assert.equal(shouldLoadTagFields(loaded()), false);
  });

  it('reads them again when the library changed', () => {
    assert.equal(shouldLoadTagFields(loaded({ libraryKey: '|Shared Documents/md' })), true);
  });

  it('reads them again when the site changed', () => {
    assert.equal(shouldLoadTagFields(loaded({ libraryKey: 'https://other|ドキュメント' })), true);
  });

  it('waits instead of starting a second read', () => {
    assert.equal(shouldLoadTagFields(loaded({ loaded: false, loading: true })), false);
  });

  it('does nothing until a library is configured', () => {
    assert.equal(
      shouldLoadTagFields(loaded({ hasLibrary: false, loaded: false, loadedKey: '', optionCount: 0 })),
      false
    );
  });

  // The regression this file exists for: editing any other setting re-renders
  // the web part and the pane, and used to clear the cache flag, which left the
  // dropdowns disabled until the pane was closed and opened again.
  it('survives a re-render caused by an unrelated setting', () => {
    const state = loaded();

    assert.equal(shouldLoadTagFields(state), false, 'no reload');
    assert.equal(canPickTagField(state), true, 'and still usable');
  });

  it('a library that turned out to be unreadable is not retried in a loop', () => {
    // The read completed, it just produced nothing.
    assert.equal(shouldLoadTagFields(loaded({ optionCount: 0 })), false);
  });
});

describe('canPickTagField', () => {
  it('allows picking once columns are known', () => {
    assert.equal(canPickTagField(loaded()), true);
  });

  it('keeps the dropdowns disabled while reading, and when nothing was found', () => {
    assert.equal(canPickTagField(loaded({ loaded: false, loading: true })), false);
    assert.equal(canPickTagField(loaded({ optionCount: 0 })), false);
  });
});

describe('tagFieldStatus', () => {
  it('explains every state the pane can be in', () => {
    assert.equal(tagFieldStatus(loaded({ hasLibrary: false })), 'needs-library');
    assert.equal(tagFieldStatus(loaded({ loading: true })), 'loading');
    assert.equal(tagFieldStatus(loaded({ loaded: false })), 'loading');
    assert.equal(tagFieldStatus(loaded({ optionCount: 0 })), 'none-found');
    assert.equal(tagFieldStatus(loaded()), 'ready');
  });

  it('says which library is missing before it says the columns are missing', () => {
    assert.equal(tagFieldStatus(loaded({ hasLibrary: false, optionCount: 0 })), 'needs-library');
  });
});
