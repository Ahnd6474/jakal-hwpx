import { useEffect, useRef, type RefObject } from 'react';

const stack: HTMLElement[] = [];
let previousOverflow = '';
const selector = 'button, a[href], input, select, textarea, [tabindex], [contenteditable="true"]';

export function trapModalFocus(element: HTMLElement, close: () => void) {
  const previous = document.activeElement as HTMLElement | null;
  const tabIndex = element.getAttribute('tabindex');
  if (tabIndex === null) element.tabIndex = -1;
  if (!stack.length) { previousOverflow = document.body.style.overflow; document.body.style.overflow = 'hidden'; }
  stack.push(element);
  const top = () => stack.at(-1) === element;
  const controls = () => {
    const visible = new WeakMap<HTMLElement, boolean>();
    return [...element.querySelectorAll<HTMLElement>(selector)].filter(item => {
    if (item.tabIndex < 0 || item.matches(':disabled, input[type="hidden"]') || item.closest('[hidden], [inert]')) return false;
    for (let node: HTMLElement | null = item; node; node = node.parentElement) {
      let shown = visible.get(node);
      if (shown === undefined) {
        const style = window.getComputedStyle(node);
        shown = style.display !== 'none' && style.visibility !== 'hidden' && style.visibility !== 'collapse';
        visible.set(node, shown);
      }
      if (!shown) return false;
      if (node === element) break;
    }
    return true;
    });
  };
  const focusFirst = () => {
    const items = controls();
    (items.find(item => item.hasAttribute('data-autofocus')) || items[0] || element).focus();
  };
  const keyboard = (event: KeyboardEvent) => {
    if (!top() || event.defaultPrevented) return;
    if (event.key === 'Escape') { event.preventDefault(); close(); return; }
    if (event.key !== 'Tab') return;
    const items = controls();
    if (!items.length) { event.preventDefault(); element.focus(); return; }
    const index = items.indexOf(document.activeElement as HTMLElement);
    if (index === -1 || (event.shiftKey ? index === 0 : index === items.length - 1)) {
      event.preventDefault(); (event.shiftKey ? items.at(-1) : items[0])?.focus();
    }
  };
  const contain = (event: FocusEvent) => { if (top() && !element.contains(event.target as Node)) focusFirst(); };
  focusFirst();
  document.addEventListener('keydown', keyboard);
  document.addEventListener('focusin', contain);
  let released = false;
  return () => {
    if (released) return;
    released = true;
    document.removeEventListener('keydown', keyboard);
    document.removeEventListener('focusin', contain);
    const wasTop = top();
    stack.splice(stack.indexOf(element), 1);
    if (!stack.length) document.body.style.overflow = previousOverflow;
    if (tabIndex === null) element.removeAttribute('tabindex');
    if (wasTop && previous?.isConnected) previous.focus();
  };
}

export function useModalFocus(root: RefObject<HTMLElement | null>, close: () => void) {
  const latestClose = useRef(close);
  latestClose.current = close;
  useEffect(() => {
    if (root.current) return trapModalFocus(root.current, () => latestClose.current());
  }, [root]);
}
