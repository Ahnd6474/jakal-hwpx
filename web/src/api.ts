export type Issue = { kind: string; message: string; location?: string; code?: string };
export type Report = { status: 'passed' | 'warning'; issues: Issue[]; new_issues: Issue[]; baseline_count: number; resolved_count: number; repairs: Issue[]; checks: string[]; checked_revision: number };
export type Paragraph = { id: string; section: number; index: number; text: string; editable: boolean };
export type Cell = { row: number; column: number; row_span: number; col_span: number; text: string; editable: boolean };
export type Table = { id: string; section: number; index: number; rows: number; columns: number; cells: Cell[] };
export type Section = { index: number; paragraphs: Paragraph[]; tables: Table[] };
export type DocumentSummary = { id: string; name: string; status: string; revision: number; created: number };
export type Settings = { instructions: string; protected_terms: string[]; strict_validation: boolean };
export type Document = DocumentSummary & { format: 'hwp' | 'hwpx'; format_note?: string; title: string; sections: Section[]; validated: number | null; expires: number; report: Report; error: string | null; settings: Settings; settings_version: number };
export type Operation = { type: 'paragraph'; section: number; index: number; text: string } | { type: 'cell'; section: number; table: number; row: number; column: number; text: string };
export type Job = { id: string; document_id: string; kind: string; status: 'queued' | 'running' | 'completed' | 'failed'; error: string | null };

export async function api<T>(path: string, init?: RequestInit): Promise<T> {
  const response = await fetch(`/api${path}`, init);
  if (!response.ok) {
    const data = await response.json().catch(() => ({}));
    throw new Error(typeof data.detail === 'string' ? data.detail : '요청을 처리하지 못했습니다. 입력 내용을 확인해 주세요.');
  }
  return response.json();
}

export function post<T>(path: string, data: unknown): Promise<T> {
  return api<T>(path, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(data) });
}

export async function waitForJob(id: string, onStage: (job: Job) => void): Promise<Job> {
  const started = Date.now();
  while (Date.now() - started < 135_000) {
    const job = await api<Job>(`/jobs/${id}`);
    onStage(job);
    if (job.status === 'failed') throw new Error(job.error || '작업에 실패했습니다.');
    if (job.status === 'completed') return job;
    await new Promise(resolve => setTimeout(resolve, 400));
  }
  throw new Error('응답 대기 시간이 지났습니다. 잠시 후 문서를 다시 열어 주세요.');
}
