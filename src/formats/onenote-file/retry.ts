import { OneNoteFormatError } from './errors';

export interface RetryOptions {
	attempts?: number;
	baseDelayMs?: number;
	onRetry?: (attempt: number, error: unknown) => void;
	wait?: (milliseconds: number) => Promise<void>;
}

export function isTransientOneNoteError(error: unknown): boolean {
	if (error instanceof OneNoteFormatError) return false;

	const code = typeof error === 'object' && error !== null && 'code' in error ? String(error.code) : '';
	if (['EBUSY', 'EAGAIN', 'EMFILE', 'ENFILE', 'EPERM', 'ETIMEDOUT'].includes(code)) return true;

	const message = error instanceof Error ? error.message : String(error);
	return /temporar|timed? out|busy|locked|try again|network|sharing violation/i.test(message);
}

export async function retryTransient<T>(operation: () => Promise<T>, options: RetryOptions = {}): Promise<T> {
	const attempts = Math.max(1, options.attempts ?? 3);
	const baseDelayMs = Math.max(0, options.baseDelayMs ?? 150);
	const wait = options.wait ?? (milliseconds => new Promise(resolve => window.setTimeout(resolve, milliseconds)));

	for (let attempt = 1; ; attempt++) {
		try {
			return await operation();
		}
		catch (error) {
			if (attempt >= attempts || !isTransientOneNoteError(error)) throw error;
			options.onRetry?.(attempt + 1, error);
			await wait(baseDelayMs * 2 ** (attempt - 1));
		}
	}
}
