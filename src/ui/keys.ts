/** The modifier for shortcuts as this platform writes it: ⌘ on Apple devices, Ctrl+ elsewhere. */
export const MOD = typeof navigator !== 'undefined' && /Mac|iPhone|iPad/.test(navigator.platform) ? '⌘' : 'Ctrl+';
