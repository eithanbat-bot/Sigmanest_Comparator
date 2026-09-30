export const APP_NAME = 'SigmaNEST Comparator';
export const VERSION = '0.4.0';
// Development mode remains enabled for internal testing. Set to false before customer release.
export const DEV_MODE = true;
export const LICENSE_API_BASE = '';
export const LICENSE_CACHE_KEY = 'sncomparator.license.v1';
export const STATE_KEY = 'sncomparator.state.v3';
export const DEFAULT_UI_SETTINGS = {
  headerFillColor: '#1F4E78',
  headerTextColor: '#FFFFFF',
  headerFontFamily: 'Calibri',
  headerFontSize: 10,
};
export const defaultLicense = {
  active: DEV_MODE,
  plan: DEV_MODE ? 'DEVELOPMENT TEST' : undefined,
  message: DEV_MODE ? 'Development mode: test license is active.' : 'Subscription not verified.'
};
