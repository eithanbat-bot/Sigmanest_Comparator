export const APP_NAME = 'SigmaNEST Comparator';
export const VERSION = '0.3.0';
// Development mode deliberately allows comparison without a licensing server.
// Set this to false for any deployed/customer build.
export const DEV_MODE = true;
export const LICENSE_API_BASE = '';
export const LICENSE_CACHE_KEY = 'sncomparator.license.v1';
export const STATE_KEY = 'sncomparator.state.v2';
export const defaultLicense = {
    active: DEV_MODE,
    plan: DEV_MODE ? 'DEVELOPMENT TEST' : undefined,
    message: DEV_MODE ? 'Development mode: subscription check bypassed.' : 'Subscription not verified.'
};
