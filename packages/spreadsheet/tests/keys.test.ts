import { describe, expect, it } from 'vitest';
import { commandKey, shortcutKey } from '../src/ui/keys';

describe('shortcut keys', () => {
  it('follow the letter the key types, so they move with the keyboard layout', () => {
    expect(shortcutKey({ key: 'z', code: 'KeyZ' })).toBe('z');
    // AZERTY: the key that types Z sits where a US keyboard has W.
    expect(shortcutKey({ key: 'z', code: 'KeyW' })).toBe('z');
    expect(shortcutKey({ key: 'G', code: 'KeyG' })).toBe('g');
  });

  it('fall back to the key’s place where it types no Latin letter', () => {
    // Russian: Ctrl+Я is Ctrl+Z.
    expect(shortcutKey({ key: 'я', code: 'KeyZ' })).toBe('z');
    // Some browsers say nothing about the key.
    expect(shortcutKey({ key: '', code: 'KeyS' })).toBe('s');
  });

  it('take / by its place, as many layouts type it with Shift', () => {
    expect(shortcutKey({ key: '/', code: 'Slash' })).toBe('/');
    // German: the key in the US / place types -; / is Shift+7.
    expect(shortcutKey({ key: '-', code: 'Slash' })).toBe('/');
    expect(shortcutKey({ key: '/', code: 'Digit7' })).toBe('/');
  });

  it('leave other keys as they are', () => {
    expect(shortcutKey({ key: 'F6', code: 'F6' })).toBe('f6');
    expect(shortcutKey({ key: 'Enter', code: 'Enter' })).toBe('enter');
  });
});

describe('the command key', () => {
  it('is ⌘ or ⌃ on Apple devices, and Ctrl elsewhere, where Meta (Windows, Super) is the system’s', () => {
    expect([commandKey({ ctrlKey: false, metaKey: true }, true), commandKey({ ctrlKey: true, metaKey: false }, true)]).toEqual([true, true]);
    expect([commandKey({ ctrlKey: false, metaKey: true }, false), commandKey({ ctrlKey: true, metaKey: false }, false)]).toEqual([false, true]);
    expect(commandKey({ ctrlKey: false, metaKey: false }, true)).toBe(false);
  });
});
