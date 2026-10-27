// Font rules and loaded faces cannot change during a synchronous paint. Cache
// only inside that scope; a later paint always resolves against current fonts.
type Scope = { owner: object | null; values: Map<string, string> };
let current: Scope | null = null;
const LIMIT = 1024;

export function withFontResolutionCache<T>(owner: object | null, operation: () => T): T {
  if (current?.owner === owner) return operation();
  const previous = current;
  current = { owner, values: new Map() };
  try { return operation(); }
  finally { current = previous; }
}

export function resolveCachedFont(font: string, resolve: () => string): string {
  if (!current) return resolve();
  const cached = current.values.get(font);
  if (cached !== undefined) return cached;
  const resolved = resolve();
  if (current.values.size >= LIMIT) current.values.delete(current.values.keys().next().value!);
  current.values.set(font, resolved);
  return resolved;
}
