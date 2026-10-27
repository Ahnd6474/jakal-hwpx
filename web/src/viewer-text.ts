/** Index painted SVG glyphs rather than paragraph text spilling across pages. */
export function renderedText(svg: string): string {
  return [...svg.matchAll(/<text\b[^>]*>([\s\S]*?)<\/text>/g)].map(match => match[1].replace(/<[^>]*>/g, '').replace(/&#(x[0-9a-f]+|\d+);|&(amp|lt|gt|quot|apos);/gi, (_all, numeric, named) => {
    if (named) return ({ amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" } as Record<string, string>)[named.toLowerCase()];
    const point = numeric[0].toLowerCase() === 'x' ? parseInt(numeric.slice(1), 16) : Number(numeric);
    return point >= 0 && point <= 0x10ffff ? String.fromCodePoint(point) : '';
  })).join('').replace(/\s/g, '');
}
