import { config as loadEnv } from 'dotenv';
import { logger } from './logger.js';

// ---------------------------------------------------------------------------
// Dual-mode API key loading
// ---------------------------------------------------------------------------
// Mode 1 — .env file:
//   Create a .env file in the project root with:
//   LEXWARE_OFFICE_API_KEY=your_api_key_here
//   dotenv loads it automatically on startup.
//
// Mode 2 — Per-request (runtime override):
//   Pass an `apiKey` argument directly to any request helper.
//   This takes precedence over the env var and is useful when the key is
//   injected by middleware (e.g. from an HTTP header or request state).
// ---------------------------------------------------------------------------
loadEnv();

const LEXOFFICE_API_BASE = 'https://api.lexoffice.io';
const USER_AGENT = 'mcp-lexware-office/0.3.0';

function resolveApiKey(provided?: string): string | null {
	if (provided) {
		return provided;
	}
	const key = process.env.LEXWARE_OFFICE_API_KEY;
	if (!key) {
		logger.error(
			'LEXWARE_OFFICE_API_KEY is not set. ' +
			'Either create a .env file or pass the key explicitly via the apiKey parameter.',
		);
	}
	return key || null;
}

export async function makeLexwareOfficeRequest<T>(
	path: string,
	apiKey?: string,
): Promise<T | null> {
	const key = resolveApiKey(apiKey);
	if (!key) {
		return null;
	}

	const url = `${LEXOFFICE_API_BASE}${path}`;
	const headers = {
		'User-Agent': USER_AGENT,
		Accept: 'application/json',
		Authorization: `Bearer ${key}`,
	};

	logger.log('Making Lexware Office request', { url });

	try {
		const response = await fetch(url, { headers });
		if (!response.ok) {
			throw new Error(`HTTP error! status: ${response.status}`);
		}
		const json = await response.json();
		logger.log('Lexware Office response', { json });
		return json as T;
	} catch (error) {
		logger.error('Error making Lexware Office request', { error });
		return null;
	}
}

export async function makeLexwareOfficeFileRequest(
	path: string,
	accept: 'application/pdf' | 'application/xml',
	apiKey?: string,
): Promise<{ data: Buffer; mimeType: string } | null> {
	const key = resolveApiKey(apiKey);
	if (!key) {
		return null;
	}

	const url = `${LEXOFFICE_API_BASE}${path}`;
	const headers = {
		'User-Agent': USER_AGENT,
		Accept: accept,
		Authorization: `Bearer ${key}`,
	};

	logger.log('Making Lexware Office file request', { url });

	try {
		const response = await fetch(url, { headers });
		if (!response.ok) {
			throw new Error(`HTTP error! status: ${response.status}`);
		}
		const contentType = response.headers.get('Content-Type') ?? accept;
		const mimeType = contentType.split(';')[0].trim();
		const arrayBuffer = await response.arrayBuffer();
		const data = Buffer.from(arrayBuffer);
		logger.log('Lexware Office file response received', { mimeType, bytes: data.length });
		return { data, mimeType };
	} catch (error) {
		logger.error('Error making Lexware Office file request', { error });
		return null;
	}
}

export type WriteResult<T> =
	| { ok: true; data: T }
	| { ok: false; status: number; error: unknown };

export async function makeLexwareOfficeWriteRequest<T>(
	path: string,
	method: 'POST' | 'PUT' | 'DELETE',
	body?: unknown,
	apiKey?: string,
): Promise<WriteResult<T> | null> {
	const key = resolveApiKey(apiKey);
	if (!key) {
		return null;
	}

	const url = `${LEXOFFICE_API_BASE}${path}`;
	const headers = {
		'User-Agent': USER_AGENT,
		'Content-Type': 'application/json',
		Accept: 'application/json',
		Authorization: `Bearer ${key}`,
	};

	logger.log('Making Lexware Office write request', { url, method });

	try {
		const response = await fetch(url, {
			method,
			headers,
			...(body !== undefined ? { body: JSON.stringify(body) } : {}),
		});

		let responseBody: unknown;
		try {
			responseBody = await response.json();
		} catch {
			responseBody = null;
		}

		if (!response.ok) {
			logger.error('Lexware Office write request failed', {
				status: response.status,
				error: responseBody,
			});
			return { ok: false, status: response.status, error: responseBody };
		}

		logger.log('Lexware Office write response', { status: response.status });
		return { ok: true; data: responseBody as T };
	} catch (error) {
		logger.error('Error making Lexware Office write request', { error });
		return null;
	}
}

export async function makeLexwareOfficeMultipartRequest<T>(
	path: string,
	formData: FormData,
	apiKey?: string,
): Promise<WriteResult<T> | null> {
	const key = resolveApiKey(apiKey);
	if (!key) {
		return null;
	}

	const url = `${LEXOFFICE_API_BASE}${path}`;
	// Do NOT set Content-Type — fetch sets it automatically with the multipart boundary
	const headers = {
		'User-Agent': USER_AGENT,
		Accept: 'application/json',
		Authorization: `Bearer ${key}`,
	};

	logger.log('Making Lexware Office multipart request', { url });

	try {
		const response = await fetch(url, { method: 'POST', headers, body: formData });

		let responseBody: unknown;
		try {
			responseBody = await response.json();
		} catch {
			responseBody = null;
		}

		if (!response.ok) {
			logger.error('Lexware Office multipart request failed', { status: response.status, error: responseBody });
			return { ok: false, status: response.status, error: responseBody };
		}

		logger.log('Lexware Office multipart response', { status: response.status });
		return { ok: true; data: responseBody as T };
	} catch (error) {
		logger.error('Error making Lexware Office multipart request', { error });
		return null;
	}
}
