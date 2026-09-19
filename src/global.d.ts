// Type declarations for static asset imports used in gulp-based SPFx projects.
declare module '*.png' {
  const value: string;
  export default value;
}

// highlight.js ships runtime code but no typings for the per-language entry
// points, which is how we keep only the languages we actually need in the bundle.
declare module 'highlight.js/lib/languages/*' {
  import { LanguageFn } from 'highlight.js';
  const language: LanguageFn;
  export default language;
}
