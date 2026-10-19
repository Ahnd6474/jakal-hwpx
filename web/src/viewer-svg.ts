import DOMPurify from 'dompurify';
import type { Figure } from './viewer-types';

export function safeSvg(raw: string): string {
  const clean = DOMPurify.sanitize(raw, { USE_PROFILES: { svg: true, svgFilters: true }, FORBID_TAGS: ['script', 'style', 'foreignObject', 'a', 'animate', 'set', 'animateTransform'], FORBID_ATTR: ['style'], RETURN_TRUSTED_TYPE: false });
  const xml = new DOMParser().parseFromString(clean, 'image/svg+xml');
  if (xml.querySelector('parsererror') || xml.documentElement.localName !== 'svg') throw new Error('잘못된 페이지 이미지');
  // The renderer must never turn document links into network requests.
  for (const node of xml.querySelectorAll('*')) {
    for (const attr of [...node.attributes]) {
      if ((attr.localName === 'href' || attr.name === 'src') && !/^#[\w:.-]+$/.test(attr.value) && !/^data:image\/(?:png|jpeg|gif|webp|bmp);base64,[a-zA-Z0-9+/=\s]+$/.test(attr.value)) node.removeAttributeNode(attr);
      else if (/url\s*\(/i.test(attr.value) && !/^url\(#[\w:.-]+\)$/.test(attr.value)) node.removeAttributeNode(attr);
    }
  }
  return new XMLSerializer().serializeToString(xml.documentElement);
}

export function cropSvg(svg: string, figure: Figure): string {
  const xml = new DOMParser().parseFromString(svg, 'image/svg+xml');
  const root = xml.documentElement;
  const padding = 5;
  root.setAttribute('viewBox', `${figure.x - padding} ${figure.y - padding} ${figure.w + padding * 2} ${figure.h + padding * 2}`);
  root.setAttribute('width', `${figure.w + padding * 2}`);
  root.setAttribute('height', `${figure.h + padding * 2}`);
  return new XMLSerializer().serializeToString(root);
}

export function highlightSvg(svg: string, query: string): string {
  const needle = query.replace(/\s/g, '').toLocaleLowerCase();
  if (!needle) return svg;
  const xml = new DOMParser().parseFromString(svg, 'image/svg+xml');
  const nodes = [...xml.querySelectorAll('text')];
  let text = '';
  const ranges = nodes.map(node => { const start = text.length; text += (node.textContent || '').replace(/\s/g, '').toLocaleLowerCase(); return { node, start, end: text.length }; });
  let position = text.indexOf(needle);
  let count = 0;
  while (position !== -1 && count++ < 500) {
    for (const { node, start, end } of ranges) if (start < position + needle.length && end > position) {
      node.setAttribute('fill', '#8a4c00'); node.setAttribute('stroke', '#ffed9c'); node.setAttribute('stroke-width', '2'); node.setAttribute('paint-order', 'stroke');
    }
    position = text.indexOf(needle, position + needle.length);
  }
  return new XMLSerializer().serializeToString(xml);
}

export async function cropPng(svg: string, figure: Figure): Promise<Blob> {
  const input = URL.createObjectURL(new Blob([cropSvg(svg, figure)], { type: 'image/svg+xml' }));
  try {
    const image = new Image();
    await new Promise<void>((resolve, reject) => { image.onload = () => resolve(); image.onerror = () => reject(new Error('이미지를 생성하지 못했습니다.')); image.src = input; });
    const scale = Math.min(3, 2048 / (figure.w + 10), 2048 / (figure.h + 10));
    const canvas = document.createElement('canvas');
    canvas.width = Math.max(1, Math.round((figure.w + 10) * scale));
    canvas.height = Math.max(1, Math.round((figure.h + 10) * scale));
    const context = canvas.getContext('2d');
    if (!context) throw new Error('이미지를 생성하지 못했습니다.');
    context.fillStyle = '#fff'; context.fillRect(0, 0, canvas.width, canvas.height);
    context.drawImage(image, 0, 0, canvas.width, canvas.height);
    // Only painted pixels leave the viewer. Off-crop page text/resources must
    // not be included in a supposedly standalone downloaded figure.
    return await new Promise<Blob>((resolve, reject) => canvas.toBlob(blob => blob ? resolve(blob) : reject(new Error('이미지를 생성하지 못했습니다.')), 'image/png'));
  } finally { URL.revokeObjectURL(input); }
}
