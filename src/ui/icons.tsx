const PATHS: Record<string, string> = {
  chart: 'M4 20V10M10 20V4M16 20v-7M22 20H2',
  image: 'M4 5h16v14H4zM4 15l4-4 4 4 3-3 5 5M15 9.5a1.5 1.5 0 1 0 0-.01',
  shape: 'M4 4h10v10H4zM14 10a6 6 0 1 1-4 9.5',
  undo: 'M9 14 4 9l5-5M4 9h10a6 6 0 0 1 0 12h-3',
  redo: 'M15 14l5-5-5-5M20 9H10a6 6 0 0 0 0 12h3',
  download: 'M12 3v12M7 10l5 5 5-5M4 21h16',
  open: 'M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z',
  info: 'M12 22a10 10 0 1 0 0-20 10 10 0 0 0 0 20zM12 16v-5M12 8h.01',
  warn: 'M12 3 2 20h20zM12 10v4M12 17h.01',
  close: 'M6 6l12 12M18 6 6 18',
  check: 'M5 12l5 5 9-10',
  lock: 'M6 11h12v10H6zM8 11V7a4 4 0 0 1 8 0v4',
  sheet: 'M4 4h16v16H4zM4 10h16M4 15h16M10 4v16',
  edit: 'M4 20h4L19 9l-4-4L4 16zM14 6l4 4',
  formula: 'M14 4h-2a3 3 0 0 0-3 3v13M6 10h7M15 13l5 6M20 13l-5 6',
  eye: 'M2 12s4-7 10-7 10 7 10 7-4 7-10 7S2 12 2 12zM12 15a3 3 0 1 0 0-6 3 3 0 0 0 0 6z',
  plus: 'M12 5v14M5 12h14',
  minus: 'M5 12h14',
  schema: 'M4 5h6v5H4zM14 14h6v5h-6zM7 10v6.5h7M14 7.5h3.5V14',
  upload: 'M12 21V9M7 14l5-5 5 5M4 3h16',
  sparkle: 'M12 3l1.8 5.2L19 10l-5.2 1.8L12 17l-1.8-5.2L5 10l5.2-1.8z',
  trash: 'M4 7h16M9 7V4h6v3M6 7l1 13h10l1-13',
  note: 'M4 5h16v11H9l-5 4z',
};

export function Icon({ name, size = 16 }: { name: keyof typeof PATHS | string; size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={1.8} strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
      <path d={PATHS[name] ?? PATHS.info} />
    </svg>
  );
}
