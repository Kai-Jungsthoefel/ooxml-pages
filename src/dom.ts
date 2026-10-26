// SPDX-License-Identifier: MIT

type Child = Node | string | null | undefined | false;
type Props = Record<string, unknown>;

/** Tiny element factory: properties go onto the element, `on*` functions become listeners, the rest become attributes. */
export function h<K extends keyof HTMLElementTagNameMap>(
  tag: K,
  props: Props = {},
  ...children: Child[]
): HTMLElementTagNameMap[K] {
  const el = document.createElement(tag);
  for (const [key, value] of Object.entries(props)) {
    if (value === undefined || value === null || value === false) continue;
    if (key.startsWith('on') && typeof value === 'function') {
      el.addEventListener(key.slice(2), value as EventListener);
    } else if (key === 'class') {
      el.className = String(value);
    } else if (key in el && !key.includes('-')) {
      (el as unknown as Props)[key] = value;
    } else {
      el.setAttribute(key, value === true ? '' : String(value));
    }
  }
  el.append(...children.filter((c): c is Node | string => c !== null && c !== undefined && c !== false));
  return el;
}

/** Toolbar button; one- or two-character labels (−, +, ↑) get the square icon-button shape. */
export function toolButton(label: string, title: string, onclick: () => void): HTMLButtonElement {
  const shape = label.length <= 2 ? 'ibtn s' : 'btn sm';
  return h('button', { class: shape, type: 'button', title, 'aria-label': title, onclick }, label);
}

export function setPressed(button: HTMLButtonElement, pressed: boolean): void {
  button.setAttribute('aria-pressed', String(pressed));
}
