import { forwardRef, type SelectHTMLAttributes } from 'react';
import { Icon } from './icons';

/**
 * A drop-down list: the browser's own select, so keyboards, screen readers
 * and phones' pickers work as they always do, with the browser's arrow
 * (spaced differently in each browser, and blind to the theme) replaced by
 * the theme's chevron. `className` sizes the whole control.
 */
export const Select = forwardRef<HTMLSelectElement, SelectHTMLAttributes<HTMLSelectElement>>(function Select({ className, ...rest }, ref) {
  return (
    <span className={className ? `wbe-select ${className}` : 'wbe-select'}>
      <select ref={ref} {...rest} />
      <Icon name="chevronDown" size={14} />
    </span>
  );
});
