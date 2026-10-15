import { useEffect, useId, useRef, useState, type MouseEvent } from 'react';
import { useT } from '../i18n/react';
import type { CommandContext, IconSource } from '../plugin/types';
import { pluginText, type RegisteredPanel, type RegisteredSource } from '../state/plugins';
import type { WorkbookStore } from '../state/store';
import { ContextMenu, type MenuItem } from './ContextMenu';
import { Icon } from './icons';
import { useDialog } from './useDialog';

/**
 * What plugins add to the editor: the Open and "Save a copy" menus, their
 * commands and panels, and their dialogs. Loaded only for an editor that has
 * plugins running.
 */

/** A plugin's icon: SVG path data on a 24×24 grid, drawn as the editor's own are. */
function PathIcon({ d, size = 16 }: { d: IconSource; size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" style={{ strokeWidth: 'var(--wbe-icon-stroke-width, 1.8)' }} strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
      <path d={d} />
    </svg>
  );
}

/** A button (named `text`) that opens a menu (named `label`) under itself, telling plugins first so they can load what they need. */
function MenuButton({ label, title, icon, items, onShow, className = 'wbe-btn', text }: { label: string; title: string; icon: 'open' | 'download'; items: () => MenuItem[]; onShow: () => void; className?: string; text: string }) {
  const [at, setAt] = useState<{ x: number; y: number } | null>(null);
  const show = (e: MouseEvent<HTMLButtonElement>) => {
    const r = e.currentTarget.getBoundingClientRect();
    onShow();
    setAt({ x: r.left, y: r.bottom + 4 });
  };
  return (
    <>
      <button className={className} aria-haspopup="menu" aria-expanded={!!at} onClick={show} onPointerEnter={onShow} title={title} aria-label={text}>
        <Icon name={icon} /> <span className="wbe-label">{text}</span> <Icon name="chevronDown" size={12} />
      </button>
      {at && <ContextMenu x={at.x} y={at.y} items={items()} label={label} onClose={() => setAt(null)} />}
    </>
  );
}

/** Open, as a menu: this device, then each source. */
export function OpenMenu({ store, onOpen }: { store: WorkbookStore; onOpen: () => void }) {
  const t = useT();
  const sources = store.plugins!.sources;
  return (
    <MenuButton
      label={t('toolbar.openMenu')}
      title={t('toolbar.openTitle')}
      text={t('toolbar.open')}
      icon="open"
      onShow={() => store.plugins!.preload(sources)}
      items={() => [{ label: t('toolbar.thisDevice'), onSelect: onOpen }, ...sources.map((s) => sourceEntry(store, s, t('toolbar.browse')))]}
    />
  );
}

/**
 * A source in the Open menu: an entry that picks, or a submenu of the files
 * it offers by name and then "Browse…" to pick. A source offering no files,
 * with nothing to pick, is shown disabled.
 */
function sourceEntry(store: WorkbookStore, source: RegisteredSource, browseLabel: string): MenuItem {
  const host = store.plugins!;
  const label = store.sourceLabel(source);
  const files = host.items(source);
  const pick = () => void host.openFrom(source);
  if (!files.length) return { label, onSelect: pick, disabled: !source.pick };
  return {
    label,
    items: [
      ...files.map((origin) => ({ label: origin.name, title: origin.location, onSelect: () => void host.openItem(origin) })),
      ...(source.pick ? [{ separator: true as const }, { label: browseLabel, onSelect: pick }] : []),
    ],
  };
}

/**
 * The saves that aren't the primary one: "Save to <source>", and "Save a
 * copy" to this device or a source. They sit beside the primary Save, or in
 * More on a phone.
 */
export function OtherSaves({ store }: { store: WorkbookStore }) {
  const t = useT();
  const ready = store.status === 'ready';
  const target = store.saveTarget;
  const source = store.originSource;
  const copies = store.copySources;
  const device = store.features.download && target !== 'download';
  return (
    <>
      {store.canSaveToSource && target !== 'source' && source && (
        <button
          className="wbe-btn"
          disabled={!ready}
          aria-disabled={store.busy}
          onClick={() => !store.busy && void store.plugins!.saveToSource()}
          title={t('toolbar.saveToTitle', { fileName: store.fileName, source: store.sourceLabel(source) })}
          aria-label={t(store.sourceUnsaved ? 'toolbar.saveToUnsaved' : 'toolbar.saveTo', { source: store.sourceLabel(source) })}
        >
          <Icon name="upload" /> <span className="wbe-label">{t('toolbar.saveTo', { source: store.sourceLabel(source) })}</span>
          {store.sourceUnsaved && <span className="wbe-dot" aria-hidden="true" />}
        </button>
      )}
      {ready && copies.length > 0 && (
        <MenuButton
          label={t('toolbar.saveCopy')}
          title={t('toolbar.saveCopy')}
          text={t('toolbar.saveCopy')}
          icon="download"
          onShow={() => store.plugins!.preload(copies)}
          items={() => [
            ...(device ? [{ label: t('toolbar.thisDevice'), onSelect: () => !store.busy && void store.exportCopy() }] : []),
            ...copies.map((s) => ({ label: store.sourceLabel(s), onSelect: () => !store.busy && void store.plugins!.saveCopyTo(s) })),
          ]}
        />
      )}
    </>
  );
}

/** Plugins' commands for the toolbar, or for More. */
export function CommandButtons({ store, where }: { store: WorkbookStore; where: 'toolbar' | 'more' }) {
  const t = useT();
  const context: CommandContext = {};
  return store.plugins!.commands
    .filter((c) => c.where === where)
    .map((c) => {
      const state = store.plugins!.commandState(c, context);
      if (!state.visible) return null;
      const label = t.text(pluginText(c.label));
      return (
        <button key={c.key} className="wbe-btn" disabled={!state.enabled} onClick={() => store.plugins!.run(c, context)} onPointerEnter={() => store.plugins!.preload([c])} title={label} aria-label={label}>
          {c.icon && <PathIcon d={c.icon} />} <span className="wbe-label">{label}</span>
        </button>
      );
    });
}

/** A button for each plugin panel, showing or hiding it, as the report's does. */
export function PanelButtons({ store }: { store: WorkbookStore }) {
  const t = useT();
  return store.plugins!.panels.map((p) => {
    const id = `plugin:${p.key}` as const;
    const title = t.text(pluginText(p.title));
    return (
      <button key={p.key} className={`wbe-btn ${store.panel === id ? 'wbe-active' : ''}`} aria-expanded={store.panel === id} onClick={() => store.set('panel', store.panel === id ? null : id)} title={title} aria-label={title}>
        {p.icon && <PathIcon d={p.icon} />} <span className="wbe-label">{title}</span>
      </button>
    );
  });
}

const noClose = () => undefined;

/**
 * An element a plugin fills with its own DOM, emptied (and the plugin told)
 * when it goes, and filled again in a new language. A dialog's body can
 * close it.
 */
function PluginBody({ store, plugin, render, close = noClose }: { store: WorkbookStore; plugin: string; render: (body: HTMLElement, close: (action: string | null) => void) => void | (() => void); close?: (action: string | null) => void }) {
  const body = useRef<HTMLDivElement>(null);
  const t = useT();
  useEffect(() => {
    const el = body.current!;
    const cleanup = store.plugins!.call(plugin, () => render(el, close), undefined);
    return () => {
      if (typeof cleanup === 'function') store.plugins!.call(plugin, cleanup, undefined);
      el.replaceChildren();
    };
    // `t` changes with the language.
  }, [store, plugin, render, close, t]);
  return <div ref={body} className="wbe-plugin-body" />;
}

export function PanelBody({ store, panel }: { store: WorkbookStore; panel: RegisteredPanel }) {
  return <PluginBody store={store} plugin={panel.plugin} render={panel.render} />;
}

/** A plugin's dialog: the editor's frame, title and buttons around the plugin's DOM. */
export function PluginDialogView({ store }: { store: WorkbookStore }) {
  const t = useT();
  const id = useId();
  const dialog = useRef<HTMLDivElement>(null);
  const d = store.pluginDialog;
  useDialog(dialog, !!d, () => d?.resolve(null), { modal: true });
  if (!d) return null;
  return (
    <div className="wbe-modal-backdrop" onMouseDown={() => d.resolve(null)}>
      <div ref={dialog} className="wbe-modal wbe-modal-wide" role="dialog" aria-modal="true" aria-labelledby={`${id}-title`} tabIndex={-1} onMouseDown={(e) => e.stopPropagation()}>
        <div className="wbe-modal-content">
          <h2 id={`${id}-title`}>{t.text(pluginText(d.title))}</h2>
          <PluginBody store={store} plugin={d.plugin} render={d.render} close={d.resolve} />
          <div className="wbe-modal-actions">
            {d.actions?.length ? (
              d.actions.map((a) => (
                <button key={a.id} className={`wbe-btn ${a.primary ? 'wbe-primary' : ''}`} onClick={() => d.resolve(a.id)}>
                  {t.text(pluginText(a.label))}
                </button>
              ))
            ) : (
              <button className="wbe-btn" onClick={() => d.resolve(null)}>
                {t('dialog.close')}
              </button>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}
