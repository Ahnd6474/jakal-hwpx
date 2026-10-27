import type { Document, Operation } from './api';

type Steps = {
  save: (doc: Document, operations: Operation[]) => Promise<Document>;
  validate: (doc: Document) => Promise<Document>;
  export: (doc: Document) => Promise<string>;
};

// Always use the revision returned by the preceding step. A failed save or
// validation stops the chain before an export URL can be requested.
export async function prepareDownload(doc: Document, operations: Operation[], steps: Steps): Promise<string> {
  let current = doc;
  if (operations.length) {
    const saved = await steps.save(current, operations);
    if (saved.id !== current.id || saved.revision !== current.revision + 1) throw new Error('저장 중 문서 버전이 바뀌었습니다. 수정 내용을 다시 확인해 주세요.');
    current = saved;
  }
  if (current.validated !== current.revision) {
    const checked = await steps.validate(current);
    if (checked.id !== current.id || checked.revision !== current.revision) throw new Error('검증 중 문서 버전이 바뀌었습니다. 현재 문서를 다시 확인해 주세요.');
    current = checked;
  }
  if (current.validated !== current.revision) throw new Error('현재 수정 버전의 검증을 완료하지 못했습니다.');
  return steps.export(current);
}
