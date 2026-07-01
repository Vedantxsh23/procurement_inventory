/* WeRoCon Lab — Procurement & Inventory System
   Access gate.
   IMPORTANT — read this if you're editing config:
   This is a STATIC site with no server, so this is NOT real authentication.
   It's a shared password that hides the UI until entered. Anyone who has
   the password (or who opens browser devtools) can get in. It stops casual
   access, not a determined person. Good enough for a small lab tool;
   not for sensitive data. When you're ready for real per-person logins
   and shared live data, swap this for Supabase Auth (see README.md).
*/
const AUTH_KEY = 'werocon_auth_ok';
// Change these to your own values before publishing.
// 'Sauravk' -> full access (add/edit/delete).
// 'SauravView' -> view-only, for anyone who should see status but not change it.
const ACCESS_CODES = {
  'Sauravk': 'Lab member',
  'SauravView': 'Professor (view)'
};
const Auth = (() => {
  function isUnlocked() {
    return sessionStorage.getItem(AUTH_KEY) !== null;
  }
  function role() {
    return sessionStorage.getItem(AUTH_KEY) || null;
  }
  function tryUnlock(code) {
    const match = ACCESS_CODES[code];
    if (match) {
      sessionStorage.setItem(AUTH_KEY, match);
      return true;
    }
    return false;
  }
  function lock() {
    sessionStorage.removeItem(AUTH_KEY);
  }
  function isViewOnly() {
    return role() === 'Professor (view)';
  }
  return { isUnlocked, tryUnlock, lock, role, isViewOnly };
})();
