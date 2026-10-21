export type RequestOptions = RequestInit & { timeoutMs?: number };

// Keep the deadline active until the response body has been consumed too.
async function request<T>(url: string, read: (response: Response) => Promise<T>, options: RequestOptions = {}, defaultTimeout = 30_000): Promise<T> {
  const { signal, timeoutMs = defaultTimeout, ...init } = options;
  const controller = new AbortController();
  let timedOut = false;
  const cancel = () => controller.abort(signal?.reason);
  if (signal?.aborted) cancel();
  else signal?.addEventListener('abort', cancel, { once: true });
  const timer = setTimeout(() => { timedOut = true; controller.abort(); }, timeoutMs);
  try {
    controller.signal.throwIfAborted();
    const response = await fetch(url, { ...init, signal: controller.signal });
    if (!response.ok) {
      const data = await response.json().catch(() => ({}));
      throw new Error(typeof data?.detail === 'string' ? data.detail : response.status >= 500
        ? '서버에서 요청을 처리하지 못했습니다. 잠시 후 다시 확인해 주세요.'
        : '요청을 처리하지 못했습니다. 입력 내용을 확인해 주세요.');
    }
    return await read(response);
  } catch (error) {
    if (signal?.aborted) throw signal.reason ?? new DOMException('요청이 취소되었습니다.', 'AbortError');
    if (timedOut) throw new Error('서버 응답 대기 시간이 지났습니다. 편집 내용은 유지됩니다. 저장·업로드 요청이었다면 처리 여부를 확인한 뒤 다시 시도해 주세요.');
    if (error instanceof TypeError) throw new Error('서버에 연결할 수 없습니다. 네트워크와 서버 실행 상태를 확인해 주세요.');
    if (error instanceof SyntaxError) throw new Error('서버 응답을 읽을 수 없습니다. 서버 실행 상태를 확인해 주세요.');
    throw error;
  } finally {
    clearTimeout(timer);
    signal?.removeEventListener('abort', cancel);
  }
}

export const requestJson = <T>(url: string, options?: RequestOptions) => request<T>(url, response => response.json(), options,
  options?.body instanceof FormData ? 90_000 : 30_000);
export const requestBytes = (url: string, options?: RequestOptions) => request(url, response => response.arrayBuffer(), options, 60_000);
