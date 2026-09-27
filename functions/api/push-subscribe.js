function json(obj, status = 200) {
  return new Response(JSON.stringify(obj), { status, headers: { 'Content-Type': 'application/json' } });
}

export async function onRequestPost({ request, env }) {
  try {
    const { familyId, subscription } = await request.json();
    if (!familyId || !/^wooriai-[0-9a-f]{16}$/.test(familyId) || !subscription || !subscription.endpoint || !subscription.keys) {
      return json({ ok: false, error: 'bad_request' }, 400);
    }
    const key = 'sub:' + familyId;
    const list = (await env.PUSH_SUBS.get(key, { type: 'json' })) || [];
    const filtered = list.filter(s => s.endpoint !== subscription.endpoint);
    filtered.push({ endpoint: subscription.endpoint, keys: subscription.keys, addedAt: Date.now() });
    const trimmed = filtered.slice(-20); // 가족당 최대 20개 기기까지만 보관
    await env.PUSH_SUBS.put(key, JSON.stringify(trimmed));
    return json({ ok: true });
  } catch (e) {
    return json({ ok: false, error: 'server_error' }, 500);
  }
}
