import { useEffect, useReducer } from 'react';

/**
 * A module loaded the first time it is wanted, or sooner with preload(),
 * and kept from then on, so a dialog that is seldom opened isn't in the
 * code every page loads. Unlike React.lazy, once loaded it renders at once.
 */
export function lazyModule<T>(load: () => Promise<T>) {
  let mod: T | undefined;
  let loading: Promise<T> | undefined;
  const get = () => (loading ??= load().then((m) => (mod = m)));
  return {
    preload: () => void get().catch(() => (loading = undefined)),
    /** The module, once loaded; asks for it while `wanted`, and renders again when it arrives. */
    useModule(wanted: boolean): T | undefined {
      const [, loaded] = useReducer((n: number) => n + 1, 0);
      useEffect(() => {
        if (wanted && !mod) void get().then(() => loaded(), () => (loading = undefined));
      }, [wanted]);
      return mod;
    },
  };
}
