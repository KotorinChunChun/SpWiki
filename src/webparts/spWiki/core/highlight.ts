import hljs from 'highlight.js/lib/core';
import javascript from 'highlight.js/lib/languages/javascript';
import typescript from 'highlight.js/lib/languages/typescript';
import shell from 'highlight.js/lib/languages/powershell';
import json from 'highlight.js/lib/languages/json';
import xml from 'highlight.js/lib/languages/xml';
import css from 'highlight.js/lib/languages/css';
import sql from 'highlight.js/lib/languages/sql';
import csharp from 'highlight.js/lib/languages/csharp';

/**
 * Only the languages listed in the design document are registered, so the
 * bundle stays small instead of pulling in all ~190 highlight.js grammars.
 */
hljs.registerLanguage('javascript', javascript);
hljs.registerLanguage('typescript', typescript);
hljs.registerLanguage('powershell', shell);
hljs.registerLanguage('json', json);
hljs.registerLanguage('xml', xml);
hljs.registerLanguage('css', css);
hljs.registerLanguage('sql', sql);
hljs.registerLanguage('csharp', csharp);

hljs.configure({ classPrefix: 'hljs-' });

/** Language names in a fenced code block that map onto a registered grammar. */
function resolveLanguage(className: string): string | undefined {
  const match: RegExpExecArray | null = /(?:^|\s)language-([\w#+.-]+)/.exec(className);
  if (!match) {
    return undefined;
  }

  const name: string = match[1].toLowerCase();
  return hljs.getLanguage(name) ? name : undefined;
}

/**
 * Highlights every fenced code block inside `container` whose language is one
 * of the registered ones. Blocks with an unknown or missing language are left
 * as plain text.
 *
 * highlight.js escapes the code it is given, so the HTML it produces is safe to
 * assign even though this runs after sanitization.
 */
export function highlightCodeBlocks(container: HTMLElement): void {
  const blocks: NodeListOf<Element> = container.querySelectorAll('pre > code');

  for (let i: number = 0; i < blocks.length; i++) {
    const block: Element = blocks[i];
    const language: string | undefined = resolveLanguage(block.className || '');
    const source: string = block.textContent || '';

    if (language) {
      block.innerHTML = hljs.highlight(source, { language: language, ignoreIllegals: true }).value;
      block.className = block.className + ' hljs';
    } else {
      block.className = (block.className ? block.className + ' ' : '') + 'hljs';
    }
  }
}
