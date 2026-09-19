import DOMPurify from 'dompurify';

import { SANITIZE_CONFIG } from './sanitizeConfig';
import { Sanitizer } from './types';

/**
 * Sanitizer bound to the browser `window`. This is what the web part uses.
 * The local test harness and the unit tests build their own via
 * {@link createSanitizer} so the same rules are exercised everywhere.
 */
export const sanitizeHtml: Sanitizer = (html: string): string => {
  return DOMPurify.sanitize(html, SANITIZE_CONFIG);
};

/** Builds a sanitizer bound to a specific window-like object. */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
export function createSanitizer(windowLike: any): Sanitizer {
  const instance: typeof DOMPurify = DOMPurify(windowLike);
  return (html: string): string => instance.sanitize(html, SANITIZE_CONFIG);
}
