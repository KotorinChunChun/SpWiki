import assert from 'node:assert/strict';
import { describe, it } from 'node:test';

import { core, render } from './setup.mjs';

const { COLLAPSED_CLASS, HIDDEN_CLASS, findHeadingAnchor, toggleHeadingAt } = core;

const isHidden = (element) => element.className.indexOf(HIDDEN_CLASS) >= 0;
const isCollapsed = (element) => element.className.indexOf(COLLAPSED_CLASS) >= 0;

describe('heading controls', () => {
  it('gives every heading with an id a copy button carrying that id', () => {
    const container = render('# 題名\n\n## 手順');
    const buttons = [...container.querySelectorAll('.mdv-heading-copy')];

    assert.equal(buttons.length, 2);
    assert.equal(buttons[0].getAttribute('data-mdv-anchor'), '題名');
    assert.equal(buttons[1].getAttribute('data-mdv-anchor'), '手順');
    assert.notEqual(buttons[0].querySelector('svg'), null);
  });

  it('keeps the button out of the heading text used elsewhere', () => {
    const container = render('## 目次\n\n## 手順');
    const entry = container.querySelector('.mdv-toc a');

    assert.equal(entry.textContent, '手順');
  });

  it('finds the anchor from a node inside the button', () => {
    const container = render('# 題名');
    const icon = container.querySelector('.mdv-heading-copy svg');

    assert.equal(findHeadingAnchor(icon, container), '題名');
  });

  it('returns an empty anchor for a click in the body', () => {
    const container = render('# 題名\n\n本文');

    assert.equal(findHeadingAnchor(container.querySelector('p'), container), '');
  });
});

describe('heading folding', () => {
  it('hides everything under the heading until the next one of the same level', () => {
    const container = render('## 一\n\n本文1\n\n## 二\n\n本文2');
    const [first, second] = container.querySelectorAll('h2');
    const paragraphs = container.querySelectorAll('p');

    assert.equal(toggleHeadingAt(first, container), true);

    assert.equal(isCollapsed(first), true);
    assert.equal(isHidden(paragraphs[0]), true);
    assert.equal(isHidden(second), false);
    assert.equal(isHidden(paragraphs[1]), false);
  });

  it('hides deeper headings and their bodies too', () => {
    const container = render('## 一\n\n### 一.一\n\n本文\n\n## 二');
    const h2 = container.querySelector('h2');
    const h3 = container.querySelector('h3');

    toggleHeadingAt(h2, container);

    assert.equal(isHidden(h3), true);
    assert.equal(isHidden(container.querySelector('p')), true);
    assert.equal(isHidden(container.querySelectorAll('h2')[1]), false);
  });

  it('unfolds again on a second click', () => {
    const container = render('## 一\n\n本文');
    const h2 = container.querySelector('h2');

    toggleHeadingAt(h2, container);
    toggleHeadingAt(h2, container);

    assert.equal(isCollapsed(h2), false);
    assert.equal(isHidden(container.querySelector('p')), false);
  });

  it('keeps an inner heading folded when the outer one is unfolded', () => {
    const container = render('## 一\n\n### 一.一\n\n本文\n\n### 一.二\n\n本文2');
    const h2 = container.querySelector('h2');
    const [inner] = container.querySelectorAll('h3');
    const [innerBody, otherBody] = container.querySelectorAll('p');

    toggleHeadingAt(inner, container);
    toggleHeadingAt(h2, container);
    toggleHeadingAt(h2, container);

    assert.equal(isCollapsed(inner), true);
    assert.equal(isHidden(innerBody), true, 'the inner section stays folded');
    assert.equal(isHidden(otherBody), false, 'the rest of the outer section is back');
  });

  it('leaves a link inside a heading to the browser', () => {
    const container = render('## [参照](install.md)', { currentPath: 'guide/index.md' });
    const anchor = container.querySelector('h2 a');

    assert.equal(toggleHeadingAt(anchor, container), false);
    assert.equal(isCollapsed(container.querySelector('h2')), false);
  });

  it('ignores a click that is not on a heading', () => {
    const container = render('## 一\n\n本文');

    assert.equal(toggleHeadingAt(container.querySelector('p'), container), false);
  });

  it('folds headings that sit inside a wrapper element', () => {
    const container = render('<section><h2>一</h2><p>本文</p><h2>二</h2></section>', {
      currentPath: 'page.html'
    });
    const first = container.querySelector('h2');

    assert.equal(toggleHeadingAt(first, container), true);
    assert.equal(isHidden(container.querySelector('p')), true);
    assert.equal(isHidden(container.querySelectorAll('h2')[1]), false);
  });
});
