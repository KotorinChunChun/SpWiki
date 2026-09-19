import assert from 'node:assert/strict';
import { describe, it } from 'node:test';

import { core } from './setup.mjs';

const {
  formatTagValue,
  hasMultiChoice,
  isSameTag,
  isSystemFieldName,
  mergeTagGroups,
  normalizeTags,
  parseTagValue,
  sharesTag,
  toEntityPropertyName,
  toTagFieldKind
} = core;

const field = (internalName, kind = 'text') => ({
  internalName,
  entityPropertyName: internalName,
  title: internalName,
  kind,
  choices: [],
  allowFillIn: true
});

describe('parseTagValue', () => {
  it('splits a text column on ; and ,', () => {
    assert.deepEqual(parseTagValue('設計; SPFx, 運用'), ['設計', 'SPFx', '運用']);
  });

  it('splits on Japanese separators too', () => {
    assert.deepEqual(parseTagValue('設計、SPFx；運用'), ['設計', 'SPFx', '運用']);
  });

  it('reads a plain array (MultiChoice, nometadata)', () => {
    assert.deepEqual(parseTagValue(['設計', 'SPFx']), ['設計', 'SPFx']);
  });

  it('reads a results wrapper (MultiChoice, verbose)', () => {
    assert.deepEqual(parseTagValue({ results: ['設計', 'SPFx'] }), ['設計', 'SPFx']);
  });

  it('reads a single Choice value', () => {
    assert.deepEqual(parseTagValue('設計'), ['設計']);
  });

  it('returns an empty array for an unset column', () => {
    assert.deepEqual(parseTagValue(null), []);
    assert.deepEqual(parseTagValue(undefined), []);
    assert.deepEqual(parseTagValue(''), []);
  });

  it('drops blanks and duplicates', () => {
    assert.deepEqual(parseTagValue('a;; b ; A ;b'), ['a', 'b']);
  });
});

describe('formatTagValue', () => {
  it('joins tags for a text column', () => {
    assert.equal(formatTagValue(['設計', 'SPFx'], 'text'), '設計; SPFx');
  });

  it('keeps only the first tag for a single choice column', () => {
    assert.equal(formatTagValue(['設計', 'SPFx'], 'choice'), '設計');
  });

  it('clears a single choice column when nothing is selected', () => {
    assert.equal(formatTagValue([], 'choice'), null);
  });

  it('wraps a multi choice column in results', () => {
    assert.deepEqual(formatTagValue(['設計', 'SPFx'], 'multiChoice'), { results: ['設計', 'SPFx'] });
  });

  it('can write a multi choice column as a plain array instead', () => {
    // The shape the retry uses when the tenant rejected the other one.
    assert.deepEqual(formatTagValue(['設計', 'SPFx'], 'multiChoice', 'array'), ['設計', 'SPFx']);
  });

  it('ignores the collection shape for the other column types', () => {
    assert.equal(formatTagValue(['設計'], 'text', 'array'), '設計');
    assert.equal(formatTagValue(['設計'], 'choice', 'array'), '設計');
  });

  it('round-trips through parseTagValue', () => {
    for (const kind of ['text', 'multiChoice']) {
      assert.deepEqual(parseTagValue(formatTagValue(['a', 'b'], kind)), ['a', 'b']);
    }
  });
});

describe('normalizeTags', () => {
  it('trims ASCII and ideographic whitespace', () => {
    assert.deepEqual(normalizeTags(['  a ', '　b　']), ['a', 'b']);
  });
});

describe('tag comparison', () => {
  it('ignores case and surrounding space', () => {
    assert.equal(isSameTag(' SPFx ', 'spfx'), true);
    assert.equal(isSameTag('設計', '運用'), false);
  });

  it('detects an overlap between two tag sets', () => {
    assert.equal(sharesTag(['設計', '運用'], ['SPFx', 'ウンヨウ']), false);
    assert.equal(sharesTag(['設計', '運用'], ['SPFx', '運用']), true);
    assert.equal(sharesTag([], ['設計']), false);
  });
});

describe('toTagFieldKind', () => {
  it('maps the SharePoint column types this web part supports', () => {
    assert.equal(toTagFieldKind('Text'), 'text');
    assert.equal(toTagFieldKind('Note'), 'text');
    assert.equal(toTagFieldKind('Choice'), 'choice');
    assert.equal(toTagFieldKind('MultiChoice'), 'multiChoice');
  });

  it('rejects the types it does not', () => {
    assert.equal(toTagFieldKind('TaxonomyFieldType'), undefined);
    assert.equal(toTagFieldKind('Number'), undefined);
    assert.equal(toTagFieldKind('User'), undefined);
  });
});

describe('isSystemFieldName', () => {
  it('rejects the columns SharePoint owns', () => {
    assert.equal(isSystemFieldName('_UIVersionString'), true);
    assert.equal(isSystemFieldName('_ComplianceTag'), true);
    assert.equal(isSystemFieldName('_ModerationStatus'), true);
    assert.equal(isSystemFieldName('_IsRecord'), true);
  });

  it('keeps a column created with a Japanese name', () => {
    // 「タグ」 and 「分類」 — SharePoint escapes each character it cannot use in
    // an internal name, so these start with an underscore without being system
    // columns. Dropping them hid every Japanese column from the dropdowns.
    assert.equal(isSystemFieldName('_x30bf__x30b0_'), false);
    assert.equal(isSystemFieldName('_x5206__x985e_'), false);
    // 「担当者」 with an ASCII suffix, which is what a second column of the same
    // name looks like.
    assert.equal(isSystemFieldName('_x62c5__x5f53__x8005_0'), false);
  });

  it('keeps an ordinary column', () => {
    assert.equal(isSystemFieldName('DocTags'), false);
    assert.equal(isSystemFieldName('Title'), false);
  });
});

describe('hasMultiChoice', () => {
  it('spots a column written as a collection', () => {
    assert.equal(hasMultiChoice([{ field: field('A', 'text'), tags: [] }]), false);
    assert.equal(
      hasMultiChoice([
        { field: field('A', 'text'), tags: [] },
        { field: field('B', 'multiChoice'), tags: [] }
      ]),
      true
    );
    assert.equal(hasMultiChoice([]), false);
  });
});

describe('mergeTagGroups', () => {
  const existing = [
    { field: field('DocTags', 'multiChoice'), tags: ['設計'] },
    { field: field('DocOwner'), tags: ['山田'] },
    { field: field('DocKind', 'choice'), tags: ['手順書'] }
  ];

  it('replaces only the column that was edited', () => {
    const merged = mergeTagGroups(existing, [
      { field: field('DocOwner'), tags: ['田中', '鈴木'] }
    ]);

    assert.deepEqual(merged.map((g) => g.field.internalName), ['DocTags', 'DocOwner', 'DocKind']);
    assert.deepEqual(merged.map((g) => g.tags), [['設計'], ['田中', '鈴木'], ['手順書']]);
  });

  it('clears a column that was emptied, rather than treating it as absent', () => {
    const merged = mergeTagGroups(existing, [{ field: field('DocTags', 'multiChoice'), tags: [] }]);

    assert.deepEqual(merged[0].tags, []);
    assert.deepEqual(merged[1].tags, ['山田']);
  });

  it('ignores a column the document does not carry', () => {
    const merged = mergeTagGroups(existing, [{ field: field('Gone'), tags: ['x'] }]);

    assert.deepEqual(merged, existing);
  });

  it('leaves everything alone when nothing was edited', () => {
    assert.deepEqual(mergeTagGroups(existing, []), existing);
  });
});

describe('toEntityPropertyName', () => {
  it('leaves an ordinary column alone', () => {
    assert.equal(toEntityPropertyName('DocTags'), 'DocTags');
    assert.equal(toEntityPropertyName('Title'), 'Title');
  });

  it('prefixes a column created with a Japanese name', () => {
    // 「業種」. SharePoint cannot start an entity property with an underscore,
    // so the item carries it as OData__x696d__x7a2e_ — using the internal name
    // read back empty and answered HTTP 400 on save.
    assert.equal(toEntityPropertyName('_x696d__x7a2e_'), 'OData__x696d__x7a2e_');
    // 「タグ」
    assert.equal(toEntityPropertyName('_x30bf__x30b0_'), 'OData__x30bf__x30b0_');
  });

  it('prefixes any other underscore column the same way', () => {
    assert.equal(toEntityPropertyName('_Status'), 'OData__Status');
  });
});
