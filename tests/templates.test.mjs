import assert from 'node:assert/strict';
import { describe, it } from 'node:test';

import { core, render } from './setup.mjs';

const { documentTitle, newDocumentTemplate } = core;

describe('documentTitle', () => {
  it('drops the extension the viewer knows', () => {
    assert.equal(documentTitle('手順.md'), '手順');
    assert.equal(documentTitle('page.html'), 'page');
    assert.equal(documentTitle('notes.markdown'), 'notes');
  });

  it('keeps a name that carries no extension of ours', () => {
    assert.equal(documentTitle('release-1.2'), 'release-1.2');
    assert.equal(documentTitle('手順'), '手順');
  });

  it('uses the file name, not the folder', () => {
    assert.equal(documentTitle('guide/api/手順.md'), '手順');
  });
});

describe('newDocumentTemplate', () => {
  it('starts a Markdown page with its title and the 目次 marker', () => {
    const template = newDocumentTemplate('インストール手順.md', 'markdown');

    assert.equal(template.indexOf('# インストール手順\n'), 0);
    assert.match(template, /\n## 目次\n/);
  });

  it('starts an HTML page with a whole document', () => {
    const template = newDocumentTemplate('手順.html', 'html');

    assert.match(template, /^<!doctype html>/);
    assert.match(template, /<title>手順<\/title>/);
    assert.match(template, /<h1>手順<\/h1>/);
    assert.match(template, /<h2>目次<\/h2>/);
  });

  it('escapes a title that would otherwise break the HTML', () => {
    const template = newDocumentTemplate('a<b>&"c".html', 'html');

    assert.equal(template.indexOf('<b>'), -1);
    assert.match(template, /a&lt;b&gt;&amp;&quot;c&quot;/);
  });

  // The template is only worth anything if the viewer renders it as intended.
  it('renders as a titled page whose table of contents is ready to fill', () => {
    const container = render(newDocumentTemplate('手順.md', 'markdown'), { currentPath: '手順.md' });

    assert.equal(container.querySelector('h1').textContent, '手順');
    assert.deepEqual(
      [...container.querySelectorAll('h2')].map((h) => h.textContent),
      ['目次', '概要']
    );
  });

  it('renders the html template the same way', () => {
    const container = render(newDocumentTemplate('手順.html', 'html'), { currentPath: '手順.html' });

    assert.equal(container.querySelector('h1').textContent, '手順');
    assert.deepEqual(
      [...container.querySelectorAll('h2')].map((h) => h.textContent),
      ['目次', '概要']
    );
  });
});
