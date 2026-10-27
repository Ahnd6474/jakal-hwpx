import { lockWorkspace, privacyLocked, trackPrivateRequest, workspaceContext } from './privacy-context.ts';

export type RequestOptions = RequestInit & { timeoutMs?: number };

export class ApiError extends Error {
  readonly code: string;
  readonly status: number | null;
  constructor(message: string, code: string, status: number | null = null) {
    super(message);
    this.name = 'ApiError';
    this.code = code;
    this.status = status;
  }
}

export function isConnectionFailure(error: unknown): boolean {
  return error instanceof ApiError && ['network_error', 'request_timeout'].includes(error.code);
}

// Keep the deadline active until the response body has been consumed too.
async function request<T>(url: string, read: (response: Response) => Promise<T>, options: RequestOptions = {}, defaultTimeout = 30_000): Promise<T> {
  const { signal, timeoutMs = defaultTimeout, ...init } = options;
  const controller = new AbortController();
  const privateRequest = url.startsWith('/api/') && !['/api/health', '/api/templates'].includes(url.split('?')[0]);
  const context = workspaceContext();
  const untrack = privateRequest ? trackPrivateRequest(controller) : () => {};
  const sessionChanged = () => privacyLocked() || (privateRequest && workspaceContext() !== context);
  const changedError = () => new ApiError('사용 세션이 변경되거나 종료되었습니다. 다시 시작해 주세요.', 'session_changed', 401);
  let timedOut = false;
  const cancel = () => controller.abort(signal?.reason);
  if (signal?.aborted) cancel();
  else signal?.addEventListener('abort', cancel, { once: true });
  const timer = setTimeout(() => { timedOut = true; controller.abort(); }, timeoutMs);
  try {
    if (privacyLocked()) throw new ApiError('사용 세션이 종료되었습니다. 다시 시작해 주세요.', 'session_changed', 401);
    controller.signal.throwIfAborted();
    const headers = new Headers(init.headers);
    if (context && privateRequest) headers.set('X-Jakal-Workspace', context);
    const response = await fetch(url, { ...init, ...(privateRequest ? { cache: 'no-store', credentials: 'same-origin', redirect: 'error' } as const : {}), headers, signal: controller.signal });
    if (sessionChanged()) throw changedError();
    if (!response.ok) {
      const data = await response.json().catch(() => ({}));
      if (sessionChanged()) throw changedError();
      if (data?.code === 'session_changed') lockWorkspace();
      throw new ApiError(typeof data?.detail === 'string' ? data.detail : response.status >= 500
        ? '서버에서 요청을 처리하지 못했습니다. 잠시 후 다시 확인해 주세요.'
        : '요청을 처리하지 못했습니다. 입력 내용을 확인해 주세요.',
      typeof data?.code === 'string' ? data.code : `http_${response.status}`, response.status);
    }
    if (privateRequest && context && response.headers.get('X-Jakal-Workspace') !== context) {
      lockWorkspace(); throw changedError();
    }
    const value = await read(response);
    if (sessionChanged()) throw changedError();
    return value;
  } catch (error) {
    if (signal?.aborted) throw signal.reason ?? new DOMException('요청이 취소되었습니다.', 'AbortError');
    if (sessionChanged()) throw changedError();
    if (timedOut) throw new ApiError('서버 응답 대기 시간이 지났습니다. 편집 내용은 유지됩니다. 저장·업로드 요청이었다면 처리 여부를 확인한 뒤 다시 시도해 주세요.', 'request_timeout');
    if (error instanceof TypeError) throw new ApiError('서버에 연결할 수 없습니다. 네트워크와 서버 실행 상태를 확인해 주세요.', 'network_error');
    if (error instanceof SyntaxError) throw new ApiError('서버 응답을 읽을 수 없습니다. 서버 실행 상태를 확인해 주세요.', 'invalid_response');
    throw error;
  } finally {
    untrack();
    clearTimeout(timer);
    signal?.removeEventListener('abort', cancel);
  }
}

export const requestJson = <T>(url: string, options?: RequestOptions) => request<T>(url, response => response.json(), options,
  options?.body instanceof FormData ? 90_000 : 30_000);
export const requestBytes = (url: string, options?: RequestOptions) => request(url, response => response.arrayBuffer(), options, 60_000);
