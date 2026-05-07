/*
 * workbookUsers.js
 *
 * Tracks which Student Retention Kit users have opened this workbook.
 *
 * On SSO login (or auto-login from a cached SSO_USER), App.jsx calls
 * `registerWorkbookUser(username)` once per session. The user is upserted
 * into the workbook's `users` array under the existing `workbookSettings`
 * document setting, with a `dateJoined` ISO timestamp on first sight.
 *
 * Existing entries are left alone so the original join date is preserved
 * even if the same user reopens the workbook later. The list is surfaced
 * in Settings → Workbook → Users and can be used for future automation
 * (e.g. routing, distribution lists).
 */

const DOC_KEY = 'workbookSettings';
const USERS_KEY = 'users';

function isOfficeReady() {
    return typeof window !== 'undefined'
        && window.Office
        && Office.context
        && Office.context.document
        && Office.context.document.settings;
}

function readWorkbookSettings() {
    if (!isOfficeReady()) return null;
    try {
        return Office.context.document.settings.get(DOC_KEY) || {};
    } catch (err) {
        console.warn('workbookUsers: failed to read document settings', err);
        return null;
    }
}

function writeWorkbookSettings(mapping) {
    if (!isOfficeReady()) return Promise.resolve(false);
    return new Promise(resolve => {
        try {
            Office.context.document.settings.set(DOC_KEY, mapping);
            Office.context.document.settings.saveAsync(result => {
                const ok = result && result.status === Office.AsyncResultStatus.Succeeded;
                if (!ok && result && result.error) {
                    console.warn('workbookUsers: saveAsync failed', result.error);
                }
                resolve(!!ok);
            });
        } catch (err) {
            console.warn('workbookUsers: failed to set document settings', err);
            resolve(false);
        }
    });
}

/**
 * Returns the list of registered users (always an array; empty if missing).
 * @returns {Array<{name: string, dateJoined: string}>}
 */
export function getWorkbookUsers() {
    const settings = readWorkbookSettings();
    if (!settings) return [];
    const users = settings[USERS_KEY];
    return Array.isArray(users) ? users : [];
}

/**
 * Upserts the given user into the workbook's users list. If the user is
 * already registered, this is a no-op (preserves the original dateJoined).
 * Safe to call repeatedly — intended to run once per session in the background.
 * @param {string} username - SSO username (display name).
 * @returns {Promise<boolean>} True if a new user was added.
 */
export async function registerWorkbookUser(username) {
    if (!username || typeof username !== 'string') return false;
    const trimmed = username.trim();
    if (!trimmed) return false;
    if (!isOfficeReady()) return false;

    const settings = readWorkbookSettings();
    if (!settings) return false;

    const users = Array.isArray(settings[USERS_KEY]) ? settings[USERS_KEY] : [];
    const exists = users.some(u => u && typeof u.name === 'string' && u.name.toLowerCase() === trimmed.toLowerCase());
    if (exists) return false;

    const next = {
        ...settings,
        [USERS_KEY]: [
            ...users,
            { name: trimmed, dateJoined: new Date().toISOString() },
        ],
    };
    const ok = await writeWorkbookSettings(next);
    if (ok) {
        console.log(`workbookUsers: registered "${trimmed}" in workbook settings`);
    }
    return ok;
}
