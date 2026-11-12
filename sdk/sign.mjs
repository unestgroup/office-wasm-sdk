// Server-side helper (Node): signs the address of a viewer site that is gated by an access token.
//
//   import { signViewerUrl } from 'office-wasm-sdk/server';
//   const url = signViewerUrl({ base: 'https://viewer.educora.vn', secret: process.env.OFFICE_WASM_TOKEN_SECRET, ttlSeconds: 3600 });
//   // -> https://viewer.educora.vn/s/<expires>.<signature>/   give it to the logged-in user's page: OfficeViewer.create(box, { url })
//
// The viewer site checks the same secret (tools/serve.cjs --config with "tokenScope", env OFFICE_WASM_TOKEN_SECRET) and answers
// 403 to a URL that has no valid, unexpired token. signature = base64url(HMAC-SHA256(secret, scope + "\n" + expires)), expires = unix seconds.
// Keep the secret on your server; never put it in a page. Needs Node (node:crypto); the browser part of the SDK does not use it.
// SPDX-License-Identifier: Apache-2.0
import { createHmac } from 'node:crypto';

/**
 * @param {{base: string, secret: string, scope?: string, ttlSeconds?: number, now?: number}} o
 *   base: address of the viewer site; scope: must equal "tokenScope" of the site (default "viewer"); now: ms, for tests
 * @returns {string} base + "/s/<expires>.<signature>/"
 */
export function signViewerUrl({ base, secret, scope = 'viewer', ttlSeconds = 3600, now = Date.now() }) {
  if (typeof secret !== 'string' || secret.length < 16) throw new Error('secret must be at least 16 characters');
  if (!base) throw new Error('base is required');
  const expires = Math.floor(now / 1000) + ttlSeconds;
  const signature = createHmac('sha256', secret).update(`${scope}\n${expires}`).digest('base64url');
  return `${String(base).replace(/\/+$/, '')}/s/${expires}.${signature}/`;
}
