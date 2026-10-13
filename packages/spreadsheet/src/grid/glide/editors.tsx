import type { GridCell, ProvideEditorComponent, TextCell } from '@glideapps/glide-data-grid';
import type { ComponentType } from 'react';
import type { CellEditorProps } from '../contract';
import { GLIDE_CLICK_OUTSIDE_IGNORE } from './glideDom';

/** A text cell whose field only allows a fixed list of values. */
export type EnumCell = TextCell & { enumOptions?: readonly string[] };

export function enumOptionsOf(cell: GridCell): readonly string[] | undefined {
  return (cell as EnumCell).enumOptions;
}

/**
 * One of the editor's cell editors, opened by Glide: Glide's cell and its
 * callbacks in, the editor's props out. When typing opens the editor, Glide
 * has already put the key typed in the cell's text.
 */
export function glideEditor(Editor: ComponentType<CellEditorProps>): ProvideEditorComponent<GridCell> {
  return function GlideCellEditor({ value, onChange, onFinishedEditing, initialValue, isHighlighted, target }) {
    const cell = value as EnumCell;
    return (
      <Editor
        text={cell.data}
        onText={(data) => onChange({ ...cell, data })}
        finish={(text, move) => onFinishedEditing(text === undefined ? undefined : { ...cell, data: text }, move)}
        typed={initialValue !== undefined}
        selectAll={!!isHighlighted}
        readOnly={cell.readonly === true}
        wrap={!!cell.allowWrapping}
        box={target}
        options={cell.enumOptions}
        popupClass={GLIDE_CLICK_OUTSIDE_IGNORE}
      />
    );
  };
}
