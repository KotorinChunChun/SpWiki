/**
 * Inline SVG shared by the chrome and the page trees.
 *
 * Drawn rather than imported so no icon font has to be shipped, and shared so
 * the padlock in the header and the padlock beside a file in a listing are the
 * same shape — they mean the same thing, and two drawings of it would drift.
 */

const NS: string = 'http://www.w3.org/2000/svg';

/**
 * A padlock: closed when the item carries permissions of its own, open when it
 * still follows its parent.
 *
 * The shackle is the whole message. Closed, it comes down into the body on both
 * sides; open, the right leg is missing so it hangs to one side. That reads at
 * a glance even at 12px, which a colour difference alone does not.
 */
export function buildPadlockIcon(doc: Document, locked: boolean, size: number): SVGElement {
  const svg: SVGElement = doc.createElementNS(NS, 'svg') as SVGElement;
  svg.setAttribute('viewBox', '0 0 16 16');
  svg.setAttribute('width', String(size));
  svg.setAttribute('height', String(size));
  svg.setAttribute('aria-hidden', 'true');
  svg.setAttribute('focusable', 'false');

  const body: Element = doc.createElementNS(NS, 'rect');
  body.setAttribute('x', '3');
  body.setAttribute('y', '7');
  body.setAttribute('width', '10');
  body.setAttribute('height', '7');
  body.setAttribute('rx', '1');
  body.setAttribute('fill', 'currentColor');
  svg.appendChild(body);

  const shackle: Element = doc.createElementNS(NS, 'path');
  shackle.setAttribute(
    'd',
    locked
      ? 'M5.2 7V5a2.8 2.8 0 0 1 5.6 0v2H9.4V5a1.4 1.4 0 0 0-2.8 0v2z'
      : 'M5.2 7V5a2.8 2.8 0 0 1 5.6 0h-1.4a1.4 1.4 0 0 0-2.8 0v2z'
  );
  shackle.setAttribute('fill', 'currentColor');
  svg.appendChild(shackle);

  return svg;
}
