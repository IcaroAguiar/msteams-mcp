/**
 * HTTP utilities with retry, timeout, and error handling.
 */

import { 
  ErrorCode, 
  createError, 
  classifyHttpError, 
  extractRetryAfter,
  type McpError,
} from '../types/errors.js';
import { type Result, ok, err } from '../types/result.js';

/** Options for HTTP requests. */
export interface HttpOptions extends Omit<RequestInit, 'signal'> {
  /** Timeout in milliseconds (default: 30000). */
  timeoutMs?: number;
  /** Maximum retry attempts (default: 3). */
  maxRetries?: number;
  /** Base delay for exponential backoff in ms (default: 1000). */
  retryBaseDelayMs?: number;
  /** Maximum delay between retries in ms (default: 10000). */
  retryMaxDelayMs?: number;
  /** Safe to repeat after the request may have reached the server (default: by HTTP method). Mark read-only POSTs. */
  idempotent?: boolean;
  /** Stop reading a text body after this many bytes and flag the response `truncated`. */
  maxResponseBytes?: number;
}

/** Successful HTTP response. */
export interface HttpResponse<T = unknown> {
  status: number;
  headers: Headers;
  data: T;
  /** The body was longer than `maxResponseBytes`; `data` holds only its start. */
  truncated?: boolean;
}

/** Retry-After deadlines per host: one service's 429 must not block the others. */
const rateLimitedUntil = new Map<string, number>();

function hostOf(url: string): string {
  try {
    return new URL(url).host;
  } catch {
    return '';
  }
}

/**
 * Makes an HTTP request with timeout, retry, and error handling.
 */
export async function httpRequest<T = unknown>(
  url: string,
  options: HttpOptions = {}
): Promise<Result<HttpResponse<T>>> {
  const {
    timeoutMs = 30000,
    maxRetries = 3,
    retryBaseDelayMs = 1000,
    retryMaxDelayMs = 10000,
    idempotent = isIdempotent(options.method),
    maxResponseBytes,
    ...fetchOptions
  } = options;

  const host = hostOf(url);
  const limitedUntil = rateLimitedUntil.get(host);
  if (limitedUntil !== undefined && Date.now() < limitedUntil) {
    const waitMs = limitedUntil - Date.now();
    return err(createError(
      ErrorCode.RATE_LIMITED,
      `Rate limited. Retry after ${Math.ceil(waitMs / 1000)}s`,
      { retryable: true, retryAfterMs: waitMs }
    ));
  }

  let lastError: McpError | null = null;

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      const result = await fetchWithTimeout<T>(url, fetchOptions, timeoutMs, idempotent, maxResponseBytes);
      
      if (result.ok) {
        return result;
      }

      lastError = result.error;

      const retryAfterMs = result.error.code === ErrorCode.RATE_LIMITED ? result.error.retryAfterMs : undefined;
      if (retryAfterMs) {
        rateLimitedUntil.set(host, Date.now() + retryAfterMs);
        // Retrying before Retry-After expires only earns another 429.
        if (retryAfterMs > retryMaxDelayMs) return result;
      }

      // A request that may have been applied is repeated only when repeating it is harmless;
      // 429 means the server refused it unprocessed.
      if (!result.error.retryable || (!idempotent && result.error.code !== ErrorCode.RATE_LIMITED)) {
        return result;
      }

      // Don't retry on last attempt
      if (attempt === maxRetries) {
        return result;
      }

      // Calculate backoff delay with jitter to avoid thundering herd
      const baseDelay = Math.min(
        retryBaseDelayMs * Math.pow(2, attempt - 1),
        retryMaxDelayMs
      );
      const delay = baseDelay * (0.5 + Math.random() * 0.5);
      
      await sleep(delay);
      
    } catch (error) {
      // Unexpected errors
      lastError = createError(
        ErrorCode.UNKNOWN,
        error instanceof Error ? error.message : String(error),
        { retryable: false }
      );
      
      if (attempt === maxRetries) {
        return err(lastError);
      }
    }
  }

  return err(lastError ?? createError(ErrorCode.UNKNOWN, 'Request failed'));
}

/**
 * Makes a single fetch request with timeout.
 */
async function fetchWithTimeout<T>(
  url: string,
  options: RequestInit,
  timeoutMs: number,
  idempotent: boolean,
  maxResponseBytes?: number,
): Promise<Result<HttpResponse<T>>> {
  const controller = new AbortController();
  // The deadline covers the body too: a server can send headers promptly and then stall.
  const timeoutId = setTimeout(() => controller.abort(), timeoutMs);

  try {
    const response = await fetch(url, {
      ...options,
      signal: controller.signal,
    });

    if (options.redirect === 'manual' && response.status >= 300 && response.status < 400) {
      await response.body?.cancel().catch(() => undefined);
      return ok({ status: response.status, headers: response.headers, data: '' as T });
    }

    if (!response.ok) {
      const retryAfterMs = extractRetryAfter(response.headers);
      const { text: errorText } = await readText(response, maxResponseBytes).catch(() => ({ text: '' }));
      
      return err(createError(
        classifyHttpError(response.status, errorText),
        `HTTP ${response.status}: ${errorText || response.statusText}`,
        { retryAfterMs }
      ));
    }

    const contentType = response.headers.get('content-type') || '';
    const { text, truncated } = await readText(response, maxResponseBytes);
    let data: T;

    if (contentType.includes('application/json')) {
      if (truncated) {
        return err(createError(ErrorCode.API_ERROR,
          `Response is larger than ${maxResponseBytes} bytes`, { retryable: false }));
      }
      data = text ? JSON.parse(text) as T : {} as T;
    } else {
      data = text as unknown as T;
    }

    return ok({
      status: response.status,
      headers: response.headers,
      data,
      ...(truncated ? { truncated } : {}),
    });

  } catch (error) {
    if (error instanceof Error) {
      if (error.name === 'AbortError') {
        return err(createError(
          ErrorCode.TIMEOUT,
          `Request timed out after ${timeoutMs}ms`,
          { retryable: idempotent }
        ));
      }

      const networkCode = networkErrorCode(error);
      if (networkCode) {
        return err(createError(
          ErrorCode.NETWORK_ERROR,
          networkCode === error.message ? error.message : `${error.message} (${networkCode})`,
          { retryable: idempotent }
        ));
      }
    }

    return err(createError(
      ErrorCode.UNKNOWN,
      error instanceof Error ? error.message : String(error),
      { retryable: false }
    ));
  } finally {
    clearTimeout(timeoutId);
  }
}

async function readText(response: Response, maxBytes: number | undefined): Promise<{ text: string; truncated: boolean }> {
  if (maxBytes === undefined || !response.body) return { text: await response.text(), truncated: false };
  const reader = response.body.getReader();
  const decoder = new TextDecoder();
  let text = '';
  let size = 0;
  for (;;) {
    const { done, value } = await reader.read();
    if (done) return { text: text + decoder.decode(), truncated: false };
    if (size + value.byteLength > maxBytes) {
      // Not flushing the decoder drops only a character the cut split.
      text += decoder.decode(value.subarray(0, maxBytes - size), { stream: true });
      await reader.cancel().catch(() => undefined);
      return { text, truncated: true };
    }
    text += decoder.decode(value, { stream: true });
    size += value.byteLength;
  }
}

const NETWORK_CODES = [
  'ECONNRESET', 'ECONNREFUSED', 'ETIMEDOUT', 'ENOTFOUND', 'EAI_AGAIN', 'EPIPE',
  'ENETUNREACH', 'EHOSTUNREACH', 'UND_ERR_SOCKET', 'UND_ERR_CONNECT_TIMEOUT',
  'UND_ERR_HEADERS_TIMEOUT', 'UND_ERR_BODY_TIMEOUT', 'UND_ERR_CLOSED',
];

// Node's fetch rejects with `TypeError: fetch failed` and keeps the socket error code in `cause`.
function networkErrorCode(error: Error): string | undefined {
  for (let current: unknown = error, depth = 0; current instanceof Error && depth < 4; current = current.cause, depth++) {
    const code = 'code' in current && typeof current.code === 'string' ? current.code : undefined;
    if (code && NETWORK_CODES.includes(code)) return code;
    const named = NETWORK_CODES.find(candidate => (current as Error).message.includes(candidate));
    if (named) return named;
  }
  return error instanceof TypeError && error.message === 'fetch failed' ? 'fetch failed' : undefined;
}

function isIdempotent(method: string | undefined): boolean {
  return ['GET', 'HEAD', 'OPTIONS', 'PUT', 'DELETE'].includes((method ?? 'GET').toUpperCase());
}

/**
 * Sleep for a specified duration.
 */
function sleep(ms: number): Promise<void> {
  return new Promise(resolve => setTimeout(resolve, ms));
}

/**
 * Clears the rate limit state (for testing).
 */
export function clearRateLimitState(): void {
  rateLimitedUntil.clear();
}
