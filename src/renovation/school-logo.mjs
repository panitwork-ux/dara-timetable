// Absolute URL keeps the crest available in standalone print windows on GitHub Pages.
export const SCHOOL_LOGO = new URL((import.meta.env?.BASE_URL || './') + 'dara-logo.png', typeof window !== 'undefined' ? window.location.href : 'http://localhost/').href;
export function schoolLogo(value) { return !value || /^https?:\/\/drive\.google\.com\//i.test(value) ? SCHOOL_LOGO : value; }
