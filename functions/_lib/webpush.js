// Web Push (RFC 8291/8292) 구현 — 표준 Web Crypto API만 사용, 외부 라이브러리 없음.

function b64urlToBytes(b64url) {
  const b64 = b64url.replace(/-/g, '+').replace(/_/g, '/').padEnd(Math.ceil(b64url.length / 4) * 4, '=');
  const bin = atob(b64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return bytes;
}
function bytesToB64url(bytes) {
  const arr = bytes instanceof ArrayBuffer ? new Uint8Array(bytes) : bytes;
  let bin = '';
  for (let i = 0; i < arr.length; i++) bin += String.fromCharCode(arr[i]);
  return btoa(bin).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}
function concatBytes(...parts) {
  const total = parts.reduce((n, p) => n + p.length, 0);
  const out = new Uint8Array(total);
  let off = 0;
  for (const p of parts) { out.set(p, off); off += p.length; }
  return out;
}

async function importVapidSigningKey(privateKeyB64url, publicKeyB64url) {
  const pub = b64urlToBytes(publicKeyB64url);
  const jwk = {
    kty: 'EC', crv: 'P-256',
    d: privateKeyB64url,
    x: bytesToB64url(pub.slice(1, 33)),
    y: bytesToB64url(pub.slice(33, 65)),
    ext: true
  };
  return crypto.subtle.importKey('jwk', jwk, { name: 'ECDSA', namedCurve: 'P-256' }, false, ['sign']);
}

async function buildVapidHeader(endpoint, vapidPublicKey, vapidPrivateKey, subject) {
  const aud = new URL(endpoint).origin;
  const exp = Math.floor(Date.now() / 1000) + 12 * 3600;
  const headerB64 = bytesToB64url(new TextEncoder().encode(JSON.stringify({ typ: 'JWT', alg: 'ES256' })));
  const payloadB64 = bytesToB64url(new TextEncoder().encode(JSON.stringify({ aud, exp, sub: subject })));
  const unsigned = `${headerB64}.${payloadB64}`;
  const key = await importVapidSigningKey(vapidPrivateKey, vapidPublicKey);
  const sig = await crypto.subtle.sign({ name: 'ECDSA', hash: 'SHA-256' }, key, new TextEncoder().encode(unsigned));
  return `vapid t=${unsigned}.${bytesToB64url(sig)}, k=${vapidPublicKey}`;
}

// RFC 8291: aes128gcm 콘텐츠 암호화
async function encryptPayload(payloadObj, p256dhB64url, authB64url) {
  const plaintext = new TextEncoder().encode(JSON.stringify(payloadObj));
  const uaPublic = b64urlToBytes(p256dhB64url);
  const authSecret = b64urlToBytes(authB64url);

  const ephemeral = await crypto.subtle.generateKey({ name: 'ECDH', namedCurve: 'P-256' }, true, ['deriveBits']);
  const asPublicRaw = new Uint8Array(await crypto.subtle.exportKey('raw', ephemeral.publicKey));

  const subscriberKey = await crypto.subtle.importKey('raw', uaPublic, { name: 'ECDH', namedCurve: 'P-256' }, false, []);
  const sharedSecret = new Uint8Array(await crypto.subtle.deriveBits({ name: 'ECDH', public: subscriberKey }, ephemeral.privateKey, 256));

  const sharedSecretKey = await crypto.subtle.importKey('raw', sharedSecret, 'HKDF', false, ['deriveBits']);
  const info1 = concatBytes(new TextEncoder().encode('WebPush: info\0'), uaPublic, asPublicRaw);
  const ikm = new Uint8Array(await crypto.subtle.deriveBits(
    { name: 'HKDF', hash: 'SHA-256', salt: authSecret, info: info1 }, sharedSecretKey, 256));

  const recordSalt = crypto.getRandomValues(new Uint8Array(16));
  const ikmKey = await crypto.subtle.importKey('raw', ikm, 'HKDF', false, ['deriveBits']);
  const cekBits = await crypto.subtle.deriveBits(
    { name: 'HKDF', hash: 'SHA-256', salt: recordSalt, info: new TextEncoder().encode('Content-Encoding: aes128gcm\0') }, ikmKey, 128);
  const nonceBits = await crypto.subtle.deriveBits(
    { name: 'HKDF', hash: 'SHA-256', salt: recordSalt, info: new TextEncoder().encode('Content-Encoding: nonce\0') }, ikmKey, 96);

  const cekKey = await crypto.subtle.importKey('raw', cekBits, 'AES-GCM', false, ['encrypt']);
  const padded = concatBytes(plaintext, new Uint8Array([2])); // RFC 8188 마지막 레코드 구분자
  const ciphertext = new Uint8Array(await crypto.subtle.encrypt({ name: 'AES-GCM', iv: new Uint8Array(nonceBits) }, cekKey, padded));

  const rs = new Uint8Array(4);
  new DataView(rs.buffer).setUint32(0, 4096);
  const header = concatBytes(recordSalt, rs, new Uint8Array([asPublicRaw.length]), asPublicRaw);
  return concatBytes(header, ciphertext);
}

export async function sendWebPush(subscription, payloadObj, vapidKeys) {
  const { endpoint, keys } = subscription;
  const body = await encryptPayload(payloadObj, keys.p256dh, keys.auth);
  const authHeader = await buildVapidHeader(endpoint, vapidKeys.publicKey, vapidKeys.privateKey, vapidKeys.subject);
  return fetch(endpoint, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/octet-stream',
      'Content-Encoding': 'aes128gcm',
      'TTL': '86400',
      'Authorization': authHeader
    },
    body
  });
}
