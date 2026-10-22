import { privateStorageKey } from './privacy-context.ts';
export type PageLayout = {
  width_mm: number; height_mm: number; left_mm: number; right_mm: number; top_mm: number; bottom_mm: number;
  header_mm: number; footer_mm: number; columns: 1 | 2 | 3; gap_mm: number; separator: boolean;
  header_text: string | null; footer_text: string | null; page_numbers: 'keep' | 'none' | 'bottom-center' | 'bottom-right';
};
export type PagePreset = { id: string; name: string; layout: PageLayout };
export type BackgroundPage = { width_mm: number; height_mm: number; image: string };
export type TemplateExtraction = { name: string; source: 'hwpx' | 'pdf' | 'images' | 'json'; choices: { label: string; layout: PageLayout; notes: string[] }[]; background_pages: BackgroundPage[]; thumbnails: string[]; notes: string[] };

export function backgroundProblem(pages: unknown): string | null {
  if (!Array.isArray(pages) || pages.length > 10) return '배경 양식은 최대 10쪽까지 사용할 수 있습니다.';
  let total = 0;
  for (const page of pages) {
    if (!page || !Number.isFinite(page.width_mm) || !Number.isFinite(page.height_mm) || page.width_mm < 80 || page.width_mm > 600 || page.height_mm < 80 || page.height_mm > 800
      || typeof page.image !== 'string' || page.image.length > 6_000_000 || !/^data:image\/jpeg;base64,[A-Za-z0-9+/]+={0,2}$/.test(page.image)) return '배경 양식의 이미지 또는 용지 크기가 올바르지 않습니다.';
    total += page.image.length * 3 / 4;
  }
  return total > 12 * 1024 * 1024 ? '배경 양식의 이미지 합계는 12MB 이하여야 합니다.' : null;
}
export type LayoutSection = { section: number; layout: PageLayout; header_text: string; footer_text: string; page_number_position: string; column_regions: number; has_background?: boolean };
export type LayoutResponse = { revision: number; sections: LayoutSection[] };
export const defaultLayout: PageLayout = { width_mm: 210, height_mm: 297, left_mm: 20, right_mm: 20, top_mm: 20, bottom_mm: 20,
  header_mm: 10, footer_mm: 10, columns: 1, gap_mm: 8, separator: false, header_text: null, footer_text: null, page_numbers: 'keep' };
export const pagePresets: PagePreset[] = [
  { id: 'standard', name: '일반 문서 · 1단', layout: { ...defaultLayout } },
  { id: 'two-column', name: '2단 · 가운데 구분선', layout: { ...defaultLayout, columns: 2, gap_mm: 10, separator: true } },
  { id: 'letterhead', name: '머리말·꼬리말', layout: { ...defaultLayout, header_text: '문서 제목', footer_text: '작성자 · 작성일', page_numbers: 'bottom-right' } },
  { id: 'three-column', name: '가로 안내문 · 3단', layout: { ...defaultLayout, width_mm: 297, height_mm: 210, columns: 3, left_mm: 15, right_mm: 15, top_mm: 15, bottom_mm: 15 } },
];

export function layoutProblem(value: unknown): string | null {
  if (!value || typeof value !== 'object') return '쪽 양식 정보가 올바르지 않습니다.';
  const layout = value as PageLayout;
  if (Object.keys(layout).some(key => !Object.hasOwn(defaultLayout, key))) return '지원하지 않는 쪽 설정이 포함되어 있습니다.';
  const ranges = { width_mm: [80, 600], height_mm: [80, 800], left_mm: [0, 100], right_mm: [0, 100], top_mm: [0, 100], bottom_mm: [0, 100], header_mm: [0, 60], footer_mm: [0, 60], gap_mm: [0, 40] };
  for (const [key, [min, max]] of Object.entries(ranges)) {
    const number = layout[key as keyof typeof ranges];
    if (typeof number !== 'number' || !Number.isFinite(number) || number < min || number > max) return '용지 크기와 여백의 입력 범위를 확인해 주세요.';
  }
  if (![1, 2, 3].includes(layout.columns) || typeof layout.separator !== 'boolean' || !['keep', 'none', 'bottom-center', 'bottom-right'].includes(layout.page_numbers)) return '단 수와 쪽 번호 설정을 확인해 주세요.';
  if ((layout.width_mm - layout.left_mm - layout.right_mm - layout.gap_mm * (layout.columns - 1)) / layout.columns < 20) return '각 단의 본문 너비는 20mm 이상이어야 합니다.';
  if (layout.height_mm - layout.top_mm - layout.bottom_mm - layout.header_mm - layout.footer_mm < 40) return '본문 높이는 40mm 이상이어야 합니다.';
  if (layout.header_mm > layout.top_mm || layout.footer_mm > layout.bottom_mm) return '머리말·꼬리말 위치는 위·아래 여백 이내여야 합니다.';
  for (const text of [layout.header_text, layout.footer_text]) {
    if (text !== null && (typeof text !== 'string' || text.length > 200 || /[\u0000-\u001f\u007f\ud800-\udfff]/u.test(text))) return '머리말·꼬리말은 줄바꿈 없이 200자 이내로 입력해 주세요.';
  }
  return null;
}

const storageKey = 'jakal-page-presets-v1';
export function readPagePresets(): PagePreset[] {
  const raw = localStorage.getItem(privateStorageKey(storageKey));
  if (!raw) return [];
  if (raw.length > 100_000) throw new Error('보관된 쪽 양식의 크기가 올바르지 않습니다.');
  const items: unknown = JSON.parse(raw);
  if (!Array.isArray(items) || items.length > 20) throw new Error('보관된 쪽 양식을 읽지 못했습니다.');
  return items.filter((item): item is PagePreset => item && typeof item.id === 'string' && typeof item.name === 'string' && item.name.length <= 60 && !layoutProblem(item.layout));
}
export function savePagePreset(name: string, layout: PageLayout): PagePreset[] {
  const problem = layoutProblem(layout);
  if (problem) throw new Error(problem);
  if (!name.trim() || name.trim().length > 60) throw new Error('양식 이름을 1~60자로 입력해 주세요.');
  const current = readPagePresets();
  if (current.length >= 20) throw new Error('내 쪽 양식은 최대 20개까지 보관할 수 있습니다. 사용하지 않는 양식을 정리해 주세요.');
  const result = [...current, { id: crypto.randomUUID(), name: name.trim(), layout: structuredClone(layout) }];
  localStorage.setItem(privateStorageKey(storageKey), JSON.stringify(result));
  return result;
}
export function deletePagePreset(id: string): PagePreset[] {
  const result = readPagePresets().filter(item => item.id !== id);
  localStorage.setItem(privateStorageKey(storageKey), JSON.stringify(result));
  return result;
}
