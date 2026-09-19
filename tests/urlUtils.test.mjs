import assert from 'node:assert/strict';
import { describe, it } from 'node:test';

import { core } from './setup.mjs';

const { buildLibraryViewUrl, getQueryParam, removeQueryParam, setQueryParam } = core;

describe('getQueryParam', () => {
  it('reads the file parameter', () => {
    assert.equal(getQueryParam('?file=README.md', 'file'), 'README.md');
    assert.equal(getQueryParam('?a=1&file=guide/install.md&b=2', 'file'), 'guide/install.md');
  });

  it('decodes percent-encoded values', () => {
    assert.equal(getQueryParam('?file=guide%2Finstall.md', 'file'), 'guide/install.md');
  });

  it('returns an empty string when absent', () => {
    assert.equal(getQueryParam('', 'file'), '');
    assert.equal(getQueryParam('?a=1', 'file'), '');
  });
});

describe('setQueryParam', () => {
  it('adds the parameter to an empty query string (design 6.3)', () => {
    assert.equal(setQueryParam('', 'file', 'README.md'), '?file=README.md');
  });

  it('replaces an existing value in place', () => {
    assert.equal(
      setQueryParam('?file=guide/install.md', 'file', 'guide/update.md'),
      '?file=guide/update.md'
    );
  });

  it('preserves the other parameters and their order', () => {
    assert.equal(
      setQueryParam('?env=WebView&file=README.md&x=1', 'file', 'guide/faq.md'),
      '?env=WebView&file=guide/faq.md&x=1'
    );
  });

  it('keeps slashes readable but escapes everything else', () => {
    assert.equal(setQueryParam('', 'file', 'a b/c&d.md'), '?file=a%20b/c%26d.md');
  });

  it('round-trips through getQueryParam', () => {
    const search = setQueryParam('?a=1', 'file', 'ガイド/導入 手順.md');
    assert.equal(getQueryParam(search, 'file'), 'ガイド/導入 手順.md');
  });
});

describe('buildLibraryViewUrl', () => {
  const ORIGIN = 'https://contoso.sharepoint.com';
  const VIEW = '/sites/docs/Shared Documents/Forms/AllItems.aspx';

  it('opens the file in the library view, in the folder that holds it', () => {
    const url = buildLibraryViewUrl(ORIGIN, VIEW, '/sites/docs/Shared Documents/md/README.md');

    assert.equal(
      url,
      'https://contoso.sharepoint.com/sites/docs/Shared%20Documents/Forms/AllItems.aspx' +
        '?id=%2Fsites%2Fdocs%2FShared%20Documents%2Fmd%2FREADME.md' +
        '&parent=%2Fsites%2Fdocs%2FShared%20Documents%2Fmd'
    );
  });

  it('points parent at the library root for a file sitting there', () => {
    const url = buildLibraryViewUrl(ORIGIN, VIEW, '/sites/docs/Shared Documents/README.md');

    assert.match(url, /&parent=%2Fsites%2Fdocs%2FShared%20Documents$/);
  });

  it('accepts a view URL that is already percent-encoded', () => {
    const encoded = '/sites/docs/Shared%20Documents/Forms/AllItems.aspx';

    assert.equal(
      buildLibraryViewUrl(ORIGIN, encoded, '/sites/docs/Shared Documents/README.md'),
      buildLibraryViewUrl(ORIGIN, VIEW, '/sites/docs/Shared Documents/README.md')
    );
  });

  it('encodes Japanese folder and file names', () => {
    const url = buildLibraryViewUrl(ORIGIN, VIEW, '/sites/docs/Shared Documents/手順/導入.html');

    assert.equal(url.indexOf('/手順/'), -1);
    assert.equal(
      url.slice(url.indexOf('?id=') + 4, url.indexOf('&parent=')),
      encodeURIComponent('/sites/docs/Shared Documents/手順/導入.html')
    );
  });

  it('returns an empty string when there is nothing to build from', () => {
    assert.equal(buildLibraryViewUrl(ORIGIN, '', '/sites/x/Shared Documents/README.md'), '');
    assert.equal(buildLibraryViewUrl(ORIGIN, VIEW, ''), '');
  });
});

describe('removeQueryParam', () => {
  it('drops the parameter and keeps the rest in order', () => {
    assert.equal(removeQueryParam('?a=1&file=x.md&b=2', 'file'), '?a=1&b=2');
  });

  it('returns an empty string when nothing is left', () => {
    assert.equal(removeQueryParam('?file=x.md', 'file'), '');
    assert.equal(removeQueryParam('', 'file'), '');
  });

  it('leaves a query string that does not carry the parameter alone', () => {
    assert.equal(removeQueryParam('?a=1', 'file'), '?a=1');
  });

  it('lets a tag search and a document path replace each other', () => {
    const search = setQueryParam(removeQueryParam('?file=guide/install.md', 'file'), 'tag', '設計');

    assert.equal(getQueryParam(search, 'tag'), '設計');
    assert.equal(getQueryParam(search, 'file'), '');
  });
});
