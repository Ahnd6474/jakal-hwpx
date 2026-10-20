// Appended to the pinned Studio entry at build time; uses its existing event bus.
// Only the enclosing same-origin Jakal window receives state and save requests.
const jakalHost = window.parent !== window;
let jakalSequence = 0;
let jakalChangeFrame = 0;
function jakalPost(type: string, extra: Record<string, unknown> = {}) {
  if (jakalHost) window.parent.postMessage({ channel: 'jakal-studio-v1', type, sequence: jakalSequence, ...extra }, location.origin);
}
eventBus.on('document-dirty-changed', () => jakalPost('dirty', { dirty: documentState.isDirty() }));
function jakalChanged() {
  jakalSequence++;
  if (!jakalChangeFrame) jakalChangeFrame = requestAnimationFrame(() => {
    jakalChangeFrame = 0;
    jakalPost('changed', { dirty: documentState.isDirty() });
  });
}
eventBus.on('document-mutated', jakalChanged);
eventBus.on('document-changed', jakalChanged);
window.addEventListener('message', event => {
  if (event.source !== window.parent || event.origin !== location.origin || event.data?.channel !== 'jakal-studio-v1') return;
  if (event.data.type === 'tools') {
    const mode = event.data.mode;
    if (!['essential', 'full', 'focus'].includes(mode)) return;
    document.documentElement.classList.toggle('jakal-full-tools', mode === 'full');
    setChromeVisibility({ menu: mode !== 'focus', toolbar: mode !== 'focus', statusbar: true });
  }
});
document.addEventListener('keydown', (event) => {
  if (jakalHost && (event.ctrlKey || event.metaKey) && !event.altKey && event.code === 'KeyS') {
    event.preventDefault(); event.stopImmediatePropagation(); jakalPost('save');
  }
}, true);
// The embed profile gives file lifecycle to the host. Keep the document light
// and the runtime's native menus, selection, IME and undo stack intact.
if (jakalHost) {
  setThemeMode('light');
  document.documentElement.classList.add('jakal-embed');
}
