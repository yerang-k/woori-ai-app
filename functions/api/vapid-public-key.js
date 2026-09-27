export async function onRequestGet({ env }) {
  return new Response(JSON.stringify({ publicKey: env.VAPID_PUBLIC_KEY || '' }), {
    headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' }
  });
}
