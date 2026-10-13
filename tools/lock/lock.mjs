#!/usr/bin/env node
/*
  قفل صفحة بكلمة مرور (تشفير فعلي، وليس مجرد إخفاء).

  تشفير صفحة:
    LOCK_PASSWORD='...' node tools/lock/lock.mjs encrypt tools/watermark.src.html tools/watermark.html "إزالة العلامة المائية"

  استرجاع الأصل من الصفحة المقفلة (للتعديل عليها):
    LOCK_PASSWORD='...' node tools/lock/lock.mjs decrypt tools/watermark.html tools/watermark.src.html

  ملفات *.src.html مستبعدة من Git حتى لا تُنشر النسخة غير المشفّرة.
  لتغيير كلمة المرور: فك التشفير بالقديمة ثم شفّر بالجديدة.
*/
import { readFileSync, writeFileSync } from 'node:fs';
import { webcrypto as wc } from 'node:crypto';
import { createInterface } from 'node:readline/promises';

const ITER = 600000;
const b64e = u => Buffer.from(u).toString('base64');
const b64d = s => new Uint8Array(Buffer.from(s, 'base64'));

async function password(){
  if(process.env.LOCK_PASSWORD) return process.env.LOCK_PASSWORD;
  const rl = createInterface({input: process.stdin, output: process.stdout});
  const pw = await rl.question('كلمة المرور: '); rl.close(); return pw;
}
async function deriveKey(pw, salt, iter){
  const base = await wc.subtle.importKey('raw', new TextEncoder().encode(pw), 'PBKDF2', false, ['deriveKey']);
  return wc.subtle.deriveKey({name:'PBKDF2', salt, iterations:iter, hash:'SHA-256'}, base, {name:'AES-GCM', length:256}, false, ['encrypt','decrypt']);
}

const [cmd, input, output, title = 'صفحة محمية'] = process.argv.slice(2);
if(!['encrypt','decrypt'].includes(cmd) || !input || !output){
  console.error('الاستخدام: node lock.mjs encrypt|decrypt <ملف الإدخال> <ملف الإخراج> [عنوان الصفحة]');
  process.exit(1);
}
const pw = await password();
if(!pw){ console.error('كلمة المرور فارغة'); process.exit(1); }

if(cmd === 'encrypt'){
  const salt = wc.getRandomValues(new Uint8Array(16)), iv = wc.getRandomValues(new Uint8Array(12));
  const key = await deriveKey(pw, salt, ITER);
  const ct = await wc.subtle.encrypt({name:'AES-GCM', iv}, key, new TextEncoder().encode(readFileSync(input, 'utf8')));
  const payload = JSON.stringify({v:1, iter:ITER, salt:b64e(salt), iv:b64e(iv), ct:b64e(ct)});
  const gate = readFileSync(new URL('./gate.html', import.meta.url), 'utf8');
  const esc = title.replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
  writeFileSync(output, gate.replaceAll('__TITLE__', esc).replace('__PAYLOAD__', () => payload));
  console.log('تم التشفير →', output);
} else {
  const m = readFileSync(input, 'utf8').match(/const P = (\{"v":1[^\n]*\});/);
  if(!m){ console.error('الملف ليس صفحة مقفلة'); process.exit(1); }
  const P = JSON.parse(m[1]);
  const key = await deriveKey(pw, b64d(P.salt), P.iter);
  try {
    const pt = await wc.subtle.decrypt({name:'AES-GCM', iv:b64d(P.iv)}, key, b64d(P.ct));
    writeFileSync(output, new TextDecoder().decode(pt));
    console.log('تم فك التشفير →', output);
  } catch(e){ console.error('كلمة المرور غير صحيحة'); process.exit(1); }
}
