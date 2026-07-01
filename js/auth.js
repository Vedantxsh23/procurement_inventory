/* WeRoCon Lab — Procurement & Inventory System
   Access gate — Supabase Auth (email + password).

   This replaces the old shared-password gate. Real per-person accounts,
   backed by Supabase Auth, with the session persisted by Supabase itself
   (it uses localStorage under the hood) so people stay signed in across
   reloads.

   ROLE / VIEW-ONLY:
   Supabase Auth doesn't have a built-in "role" field, so we read it from
   the user's metadata (set when the account is created / invited).
   In the Supabase dashboard: Authentication -> Users -> select a user ->
   "User Metadata" and add:
     { "role": "Professor (view)" }
   Any user WITHOUT that metadata is treated as full access ("Lab member").
   If you'd rather manage roles in a table instead of user metadata, see
   the note above getRole() below.

   Requires: config.js loaded first (creates the `sb` Supabase client),
   and Email/Password sign-in enabled in Supabase Auth settings, with the
   relevant users created there (Authentication -> Users -> Add user).
*/

const Auth = (() => {
  let currentUser = null; // Supabase auth user object, or null

  function getRole() {
    // Reads role out of user_metadata. Swap this for a DB lookup
    // (e.g. a `profiles` table keyed by user id) if you'd rather manage
    // roles there instead of in Supabase Auth's metadata.
    const meta = currentUser?.user_metadata || {};
    return meta.role || 'Lab member';
  }

  async function init() {
    const { data, error } = await sb.auth.getSession();
    if (error) {
      console.error('Auth.init session error:', error);
      return;
    }
    currentUser = data?.session?.user || null;

    // Keep currentUser in sync if the session changes in another tab,
    // gets refreshed, or expires.
    sb.auth.onAuthStateChange((_event, session) => {
      currentUser = session?.user || null;
    });
  }

  function isUnlocked() {
    return !!currentUser;
  }

  function email() {
    return currentUser?.email || null;
  }

  function role() {
    if (!currentUser) return null;
    return getRole();
  }

  function isViewOnly() {
    return role() === 'Professor (view)';
  }

  async function tryUnlock(emailInput, password) {
    if (!emailInput || !password) {
      return { ok: false, message: 'Enter both email and password.' };
    }
    const { data, error } = await sb.auth.signInWithPassword({
      email: emailInput.trim(),
      password
    });
    if (error) {
      return { ok: false, message: error.message || 'Sign in failed.' };
    }
    currentUser = data?.user || null;
    return { ok: true };
  }

  async function lock() {
    const { error } = await sb.auth.signOut();
    if (error) console.error('Auth.lock signOut error:', error);
    currentUser = null;
  }

  return { init, isUnlocked, tryUnlock, lock, email, role, isViewOnly };
})();
