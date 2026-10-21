import type { Document, Job, Operation } from './api';
import { ApiError, isConnectionFailure, requestJson, type RequestOptions } from './request.ts';

type Base = Pick<Document, 'id' | 'revision' | 'settings_version'>;
type Command = { kind: 'edit'; operations: Operation[] } | { kind: 'validate' }
  | { kind: 'restore'; source_revision: number } | { kind: 'studio'; bytes: Uint8Array };
type Accepted = { job_id: string; document_id: string; kind: string; base_revision: number };
type Progress = { signal?: AbortSignal; onStage?: (job: Job) => void; onReconnect?: () => void };
type Json = <T>(url: string, options?: RequestOptions) => Promise<T>;

export class DocumentCommandError extends Error {
  readonly code: string;
  readonly jobId?: string;
  constructor(message: string, code: string, jobId?: string) {
    super(message);
    this.name = 'DocumentCommandError';
    this.code = code;
    this.jobId = jobId;
  }
}

export function pause(ms: number, signal?: AbortSignal): Promise<void> {
  return new Promise((resolve, reject) => {
    signal?.throwIfAborted();
    const cancel = () => { clearTimeout(timer); reject(signal?.reason); };
    const timer = setTimeout(() => { signal?.removeEventListener('abort', cancel); resolve(); }, ms);
    signal?.addEventListener('abort', cancel, { once: true });
  });
}

/** Site-side use cases. No global/window bridge or external tool transport. */
export class DocumentCommands {
  // Only hashes and request IDs are retained. Retrying unchanged local edits
  // after an ambiguous response rejoins the same durable job in this page.
  private pending = new Map<string, string>();
  private json: Json;
  private delay: typeof pause;
  constructor(json: Json = requestJson, delay = pause) { this.json = json; this.delay = delay; }

  async execute(base: Base, command: Command, progress: Progress = {}): Promise<Document> {
    progress.signal?.throwIfAborted();
    base = { id: base.id, revision: base.revision, settings_version: base.settings_version };
    const { kind } = command;
    const payload = { revision: base.revision, settings_version: base.settings_version,
      ...(kind === 'edit' ? { operations: command.operations } : {}),
      ...(kind === 'restore' ? { source_revision: command.source_revision } : {}) };
    const serialized = JSON.stringify(payload);
    const bytes = kind === 'studio' ? new Uint8Array(command.bytes) : new TextEncoder().encode(serialized);
    const digest = Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', bytes)), b => b.toString(16).padStart(2, '0')).join('');
    const intent = `${base.id}:${base.revision}:${base.settings_version}:${kind}:${digest}`;
    let key = this.pending.get(intent);
    if (!key) {
      key = crypto.randomUUID().replaceAll('-', '');
      if (this.pending.size >= 64) this.pending.delete(this.pending.keys().next().value!);
      this.pending.set(intent, key);
    }
    const paths = { edit: 'operations', validate: 'validations', restore: 'restorations', studio: 'studio' };
    let path = `/api/documents/${base.id}/${paths[kind]}`;
    let body: BodyInit = serialized;
    const headers: Record<string, string> = { 'Idempotency-Key': key };
    if (kind === 'studio') {
      path += `?revision=${base.revision}&settings_version=${base.settings_version}`;
      const form = new FormData();
      form.append('file', new Blob([bytes], { type: 'application/hwp+zip' }), 'edited.hwpx');
      body = form;
    } else headers['Content-Type'] = 'application/json';
    let accepted: Accepted | undefined;
    try {
      for (let attempt = 0; attempt < 2; attempt++) {
        try {
          accepted = await this.json<Accepted>(path, { method: 'POST', headers, body, signal: progress.signal });
          break;
        } catch (error) {
          // Only these four endpoints guarantee idempotency. Generic requests,
          // uploads and document copies are deliberately never retried here.
          if (attempt || !isConnectionFailure(error) || progress.signal?.aborted) throw error;
          progress.onReconnect?.();
          await this.delay(400, progress.signal);
        }
      }
      if (!accepted || typeof accepted.job_id !== 'string' || !accepted.job_id || accepted.document_id !== base.id
        || accepted.kind !== kind || accepted.base_revision !== base.revision) {
        throw new DocumentCommandError('작업 접수 정보를 확인하지 못했습니다. 편집 내용은 유지됩니다.', 'invalid_receipt');
      }
      const job = await this.wait(accepted.job_id, base.id, kind, progress);
      const revision = base.revision + (kind === 'validate' ? 0 : 1);
      if (job.result?.revision !== revision) throw new DocumentCommandError('작업 결과의 문서 버전이 올바르지 않습니다. 편집 내용은 유지됩니다.', 'invalid_result', job.id);
      const fresh = await this.json<Document>(`/api/documents/${base.id}`, { signal: progress.signal });
      if (!fresh || fresh.id !== base.id || fresh.status !== 'ready' || fresh.revision !== revision
        || fresh.settings_version !== base.settings_version || (kind === 'validate' && fresh.validated !== revision)) {
        throw new DocumentCommandError('다른 작업에서 문서나 설정을 변경했습니다. 편집 내용을 보관한 뒤 최신 문서를 다시 열어 주세요.', 'revision_conflict', job.id);
      }
      this.pending.delete(intent);
      return fresh;
    } catch (error) {
      // Failed jobs and definite request rejections can be submitted anew. On
      // a lost response keep the identity for the user's next Save/Validate.
      if ((error instanceof ApiError && error.status !== null && error.status < 500)
        || (error instanceof DocumentCommandError && error.code === 'job_failed')) this.pending.delete(intent);
      throw error;
    }
  }

  private async wait(id: string, documentId: string, kind: string, progress: Progress): Promise<Job> {
    const deadline = Date.now() + 135_000;
    let disconnected = 0;
    while (Date.now() < deadline) {
      progress.signal?.throwIfAborted();
      const started = Date.now();
      let job: Job;
      try {
        job = await this.json<Job>(`/api/jobs/${id}?wait_ms=1200`, {
          signal: progress.signal, timeoutMs: Math.min(10_000, Math.max(1, deadline - Date.now())) });
        disconnected = 0;
      } catch (error) {
        if (!isConnectionFailure(error) || ++disconnected > 2 || progress.signal?.aborted) throw error;
        progress.onReconnect?.();
        await this.delay(500 * disconnected, progress.signal);
        continue;
      }
      if (!job || job.id !== id || job.document_id !== documentId || job.kind !== kind
        || !['queued', 'running', 'completed', 'failed'].includes(job.status)) {
        throw new DocumentCommandError('다른 문서의 작업 응답을 받았습니다. 편집 내용은 유지됩니다.', 'invalid_job', id);
      }
      progress.onStage?.(job);
      if (job.status === 'failed') throw new DocumentCommandError(job.error || '작업에 실패했습니다.', 'job_failed', id);
      if (job.status === 'completed') return job;
      const remaining = Math.min(200 - (Date.now() - started), deadline - Date.now());
      if (remaining > 0) await this.delay(remaining, progress.signal);
    }
    throw new DocumentCommandError('처리 결과를 기다리는 시간이 지났습니다. 편집 내용은 유지됩니다. 같은 작업을 다시 누르면 기존 처리 결과를 확인합니다.', 'job_timeout', id);
  }
}

export const documentCommands = new DocumentCommands();
