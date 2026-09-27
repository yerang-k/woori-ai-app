import { sendWebPush } from '../_lib/webpush.js';

function json(obj, status = 200) {
  return new Response(JSON.stringify(obj), { status, headers: { 'Content-Type': 'application/json' } });
}

// 가족 동기화 서버(Google Apps Script)만 호출하는 내부 엔드포인트.
// PUSH_SHARED_SECRET을 아는 요청만 처리해 임의의 제3자가 가족에게 알림을 보내지 못하게 막는다.
export async function onRequestPost({ request, env }) {
  const auth = request.headers.get('Authorization') || '';
  if (!env.PUSH_SHARED_SECRET || auth !== `Bearer ${env.PUSH_SHARED_SECRET}`) {
    return json({ ok: false, error: 'unauthorized' }, 401);
  }
  let body;
  try { body = await request.json(); } catch (e) { return json({ ok: false, error: 'bad_json' }, 400); }
  const { familyId, title, body: msgBody } = body || {};
  if (!familyId || !/^wooriai-[0-9a-f]{16}$/.test(familyId)) {
    return json({ ok: false, error: 'bad_request' }, 400);
  }

  const key = 'sub:' + familyId;
  const list = (await env.PUSH_SUBS.get(key, { type: 'json' })) || [];
  if (!list.length) return json({ ok: true, sent: 0 });

  const vapidKeys = {
    publicKey: env.VAPID_PUBLIC_KEY,
    privateKey: env.VAPID_PRIVATE_KEY,
    subject: env.VAPID_SUBJECT || 'mailto:admin@example.com'
  };
  const payload = { title: title || '우리 아이', body: msgBody || '새 소식이 있어요.' };

  const survivors = [];
  let sent = 0;
  await Promise.all(list.map(async (sub) => {
    try {
      const res = await sendWebPush(sub, payload, vapidKeys);
      if (res.status === 404 || res.status === 410) return; // 만료된 구독 → 목록에서 제거
      survivors.push(sub);
      if (res.ok) sent++;
    } catch (e) {
      survivors.push(sub);
    }
  }));
  if (survivors.length !== list.length) {
    await env.PUSH_SUBS.put(key, JSON.stringify(survivors));
  }
  return json({ ok: true, sent });
}
