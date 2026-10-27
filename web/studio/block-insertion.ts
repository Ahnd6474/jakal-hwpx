type Position = { sectionIndex: number; paragraphIndex: number; charOffset: number };
type TextEngine = {
  insertText: (section: number, paragraph: number, offset: number, text: string) => string;
  splitParagraph: (section: number, paragraph: number, offset: number) => string;
};

/** Called inside the editor's snapshot transaction: one undo and one refresh. */
export function insertBlock(engine: TextEngine, position: Position, text: unknown): Position {
  if (typeof text !== 'string' || !text.trim() || text.length > 12001 || /[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(text)) throw new Error('올바른 블록 내용을 입력하세요.');
  let at = { ...position };
  const lines = text.replace(/\r\n?/g, '\n').split('\n');
  if (lines.length > 121) throw new Error('블록 내용을 120줄 이내로 입력해 주세요.');
  for (let index = 0; index < lines.length; index++) {
    if (lines[index]) {
      const result = JSON.parse(engine.insertText(at.sectionIndex, at.paragraphIndex, at.charOffset, lines[index]));
      if (!result.ok || !Number.isInteger(result.charOffset)) throw new Error('블록의 텍스트를 삽입하지 못했습니다.');
      at = { ...at, charOffset: result.charOffset };
    }
    if (index < lines.length - 1) {
      const result = JSON.parse(engine.splitParagraph(at.sectionIndex, at.paragraphIndex, at.charOffset));
      if (!result.ok || !Number.isInteger(result.paraIdx)) throw new Error('블록의 문단을 만들지 못했습니다.');
      at = { ...at, paragraphIndex: result.paraIdx, charOffset: 0 };
    }
  }
  return at;
}
