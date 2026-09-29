import { DEV_MODE, LICENSE_API_BASE, defaultLicense } from './config.js';
import { loadState, persistLicense } from './state.js';
export async function checkLicense() {
    if (DEV_MODE) {
        await persistLicense(defaultLicense);
        return defaultLicense;
    }
    const current = await loadState();
    const token = current.license?.token;
    if (!LICENSE_API_BASE)
        return { active: false, message: 'Licensing service is not configured.' };
    try {
        const response = await fetch(`${LICENSE_API_BASE}/v1/license/status`, {
            headers: token ? { Authorization: `Bearer ${token}` } : {},
            credentials: 'include',
        });
        const data = await response.json();
        const state = {
            active: response.ok && data.active === true,
            company: data.company,
            email: data.email,
            expiresAt: data.expiresAt,
            plan: data.plan,
            token,
            message: response.ok ? data.message : (data.message || 'Subscription validation failed.'),
        };
        await persistLicense(state);
        return state;
    }
    catch (err) {
        return { active: false, token, message: `License server unavailable: ${err.message}` };
    }
}
export async function activateLicense(code) {
    if (DEV_MODE)
        return defaultLicense;
    if (!LICENSE_API_BASE)
        return { active: false, message: 'Licensing service is not configured.' };
    const response = await fetch(`${LICENSE_API_BASE}/v1/license/activate`, {
        method: 'POST', headers: { 'Content-Type': 'application/json' }, credentials: 'include', body: JSON.stringify({ code })
    });
    const data = await response.json();
    const state = {
        active: response.ok && data.active === true,
        company: data.company,
        email: data.email,
        expiresAt: data.expiresAt,
        plan: data.plan,
        token: data.token,
        message: data.message,
    };
    await persistLicense(state);
    return state;
}
