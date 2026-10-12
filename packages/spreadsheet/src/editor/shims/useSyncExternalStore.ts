/**
 * react-i18next imports useSyncExternalStore from use-sync-external-store's
 * CommonJS shim, whose require('react') can't work in the React build, where
 * React is the host's. React 18 has the hook itself.
 */
export { useSyncExternalStore } from 'react';
