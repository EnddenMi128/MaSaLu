var __defProp = Object.defineProperty;
var __name = (target, value) => __defProp(target, "name", { value, configurable: true });

// cloudflare/workers.js
var workers_default = {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);
    if (url.pathname === "/calendar-api") return handleCalendarApi(request, env);
    if (request.method === "GET" && url.pathname === "/health") {
      return new Response("ok", { status: 200, headers: { "content-type": "text/plain; charset=utf-8" } });
    }
    if (request.method !== "POST" || url.pathname !== "/webhook") {
      return new Response("Not found", { status: 404 });
    }
    if (!env.LINE_CHANNEL_SECRET || !env.LINE_CHANNEL_ACCESS_TOKEN || !env.GAS_SHARED_SECRET || !env.GAS_WEB_APP_URL) {
      console.error("Required Worker environment variable missing");
      return new Response("Service unavailable", { status: 503 });
    }
    const originalBody = await request.text();
    const signature = request.headers.get("x-line-signature");
    if (!await verifyLineSignature(originalBody, signature, env.LINE_CHANNEL_SECRET)) {
      return new Response("Invalid LINE signature", { status: 401 });
    }
    let envelope;
    try {
      envelope = JSON.parse(originalBody);
    } catch {
      return new Response("Invalid JSON", { status: 400 });
    }
    if (!Array.isArray(envelope.events)) return new Response("Invalid events", { status: 400 });
    const userQueues = /* @__PURE__ */ new Map();
    const results = await Promise.allSettled(envelope.events.map((event) => {
      const uid = event?.source?.userId;
      const previous = userQueues.get(uid);
      const next = previous ? previous.catch(() => {
      }).then(() => handleEvent(event, env, ctx)) : handleEvent(event, env, ctx);
      userQueues.set(uid, next);
      return next;
    }));
    if (results.some((result) => result.status === "rejected")) {
      for (const result of results) {
        if (result.status === "rejected") console.error("Webhook processing failed:", String(result.reason?.message || result.reason));
      }
      return new Response("Temporary processing error", { status: 502 });
    }
    return new Response("OK", { status: 200 });
  }
};
async function handleCalendarApi(request, env) {
  const allowedOrigin = env.CALENDAR_ALLOWED_ORIGIN || "https://enddenmi128.github.io";
  if (request.headers.get("origin") !== allowedOrigin) {
    return new Response("Forbidden origin", { status: 403 });
  }
  const headers = {
    "access-control-allow-origin": allowedOrigin,
    "access-control-allow-methods": "POST, OPTIONS",
    "access-control-allow-headers": "content-type, authorization",
    "cache-control": "no-store",
    "vary": "Origin",
    "content-type": "application/json; charset=utf-8"
  };
  const json = /* @__PURE__ */ __name((value, status = 200) => new Response(JSON.stringify(value), { status, headers }), "json");
  if (request.method === "OPTIONS") return new Response(null, { status: 204, headers });
  if (request.method !== "POST") return json({ ok: false, error: "Method not allowed" }, 405);
  if (!env.GAS_SHARED_SECRET || !env.GAS_WEB_APP_URL) return json({ ok: false, error: "Worker \u5C1A\u672A\u8A2D\u5B9A GAS \u9023\u7DDA" }, 503);
  const auth = request.headers.get("authorization") || "";
  const tokenMatch = /^Bearer\s+([^\s]+)$/i.exec(auth);
  if (!tokenMatch || tokenMatch[1].length > 4096) return json({ ok: false, code: "unauthorized", error: "\u8ACB\u900F\u904E LINE \u767B\u5165\u6392\u73ED\u6708\u66C6" }, 401);
  let input;
  try {
    const body = await request.text();
    if (body.length > 1e5) return json({ ok: false, error: "\u6392\u73ED\u5167\u5BB9\u904E\u5927" }, 413);
    input = JSON.parse(body);
  } catch {
    return json({ ok: false, error: "Invalid JSON" }, 400);
  }
  const actions = ["nfcBind", "bootstrap", "managerSession", "saveDay", "clearDay", "addStaff", "deleteStaff", "addShift", "updateShift", "deleteShift"];
  if (!input || !actions.includes(input.action) || !input.payload || typeof input.payload !== "object" || Array.isArray(input.payload)) return json({ ok: false, error: "\u4E0D\u652F\u63F4\u7684\u64CD\u4F5C" }, 400);
  try {
    const response = await fetch(env.GAS_WEB_APP_URL, {
      method: "POST",
      headers: { "content-type": "application/json; charset=UTF-8" },
      redirect: "follow",
      body: JSON.stringify({ kind: "calendar_api", action: input.action, payload: input.payload, lineAccessToken: tokenMatch[1], secret: env.GAS_SHARED_SECRET })
    });
    if (!response.ok) throw new Error("GAS upstream failed");
    const result = await response.json();
    if (!result || typeof result.ok !== "boolean") throw new Error("Invalid upstream response");
    return json(result, result.code === "unauthorized" ? 401 : 200);
  } catch {
    return json({ ok: false, error: "GAS \u56DE\u61C9\u5931\u6557\uFF0C\u8ACB\u91CD\u65B0\u6574\u7406\u78BA\u8A8D\u6392\u73ED\u7D50\u679C\u5F8C\u518D\u64CD\u4F5C" }, 502);
  }
}
__name(handleCalendarApi, "handleCalendarApi");
async function verifyLineSignature(rawBody, signature, channelSecret) {
  if (!rawBody || !signature || !channelSecret) return false;
  try {
    const expected = Uint8Array.from(atob(signature), (char) => char.charCodeAt(0));
    if (expected.length !== 32) return false;
    const encoder = new TextEncoder();
    const key = await crypto.subtle.importKey(
      "raw",
      encoder.encode(channelSecret),
      { name: "HMAC", hash: "SHA-256" },
      false,
      ["verify"]
    );
    return await crypto.subtle.verify("HMAC", key, expected, encoder.encode(rawBody));
  } catch {
    return false;
  }
}
__name(verifyLineSignature, "verifyLineSignature");
function parsePostbackData(data) {
  if (typeof data !== "string" || data.length > 300) return null;
  const qs = new URLSearchParams(data);
  const id = qs.get("id");
  const version = Number(qs.get("v"));
  const choice = qs.get("choice");
  if (qs.get("action") !== "shift_reply" || !/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(id || "")) return null;
  if (!Number.isSafeInteger(version) || version < 1 || !["agree", "disagree"].includes(choice)) return null;
  return { id, version, choice };
}
__name(parsePostbackData, "parsePostbackData");
function parseManagerPostbackData(data) {
  if (typeof data !== "string" || data.length > 300) return null;
  const qs = new URLSearchParams(data);
  const id = qs.get("id"), version = Number(qs.get("v")), decision = qs.get("decision");
  if (qs.get("action") !== "manager_shift" || !/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(id || "") || !Number.isSafeInteger(version) || version < 1 || !["skip", "replace", "select", "cancel", "resend"].includes(decision)) return null;
  if (decision === "select") {
    const staffId = qs.get("staff");
    if (!staffId) return null;
    return { id, version, decision, staffId };
  }
  if (decision === "replace" && qs.has("page")) {
    const page = Number(qs.get("page"));
    if (!Number.isSafeInteger(page) || page < 0) return null;
    return { id, version, decision, page };
  }
  return { id, version, decision };
}
__name(parseManagerPostbackData, "parseManagerPostbackData");
function parseShiftFollowupData(data, pickedDate) {
  if (typeof data !== "string" || data.length > 300) return null;
  const qs = new URLSearchParams(data);
  const id = qs.get("id"), version = Number(qs.get("v")), choice = qs.get("choice");
  const token = qs.get("proof") || "";
  if (qs.get("action") !== "shift_followup" || !/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(id || "") || !Number.isSafeInteger(version) || version < 1 || !/^(contact(?:_end)?|swap(?:_date|_shift|_target)?|leave(?:_(?:personal|sick|other)(?:_send)?)?)$/.test(choice || "") || token && !/^[A-Za-z0-9_-]{43}$/.test(token)) return null;
  if (choice === "swap" && qs.has("page")) {
    const page = Number(qs.get("page"));
    if (!Number.isSafeInteger(page) || page < 0) return null;
    return { id, version, choice, token, page };
  }
  if (["swap_date", "swap_shift", "swap_target"].includes(choice)) {
    const date = String(pickedDate || qs.get("date") || "");
    const target = qs.get("target") || "";
    const targetVersion = Number(qs.get("tv"));
    const page = Number(qs.get("page"));
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || !Number.isSafeInteger(page) || page < 0 || choice !== "swap_date" && !/^[0-9a-f-]{36}$/i.test(target) || choice === "swap_target" && (!Number.isSafeInteger(targetVersion) || targetVersion < 1)) return null;
    return { id, version, choice, token, date, target, targetVersion, page };
  }
  return { id, version, choice, token };
}
__name(parseShiftFollowupData, "parseShiftFollowupData");
function parseSwapResponse(data) {
  if (typeof data !== "string" || data.length > 300) return null;
  const qs = new URLSearchParams(data);
  const id = qs.get("id"), decision = qs.get("decision");
  return qs.get("action") === "swap_response" && /^[0-9a-f-]{36}$/i.test(id || "") && ["agree", "decline"].includes(decision) ? { id, decision } : null;
}
__name(parseSwapResponse, "parseSwapResponse");
async function handleEvent(event, env, ctx) {
  if (!event || event.source?.type !== "user" || !/^U[0-9a-f]{32}$/i.test(event.source.userId || "")) return;
  const lineUid = event.source.userId;
  const eventId = event.webhookEventId || "";
  if (event.type === "follow") {
    const result = await sendToGas({ kind: "account_link_follow", lineUid, eventId }, env);
    if (event.replyToken && result.message) await replyLineSafe(event.replyToken, result.message, env);
    return;
    return;
  }
  if (event.type === "message" && event.message?.type === "text") {
    const text = String(event.message.text || "").trim();
    if (text === "\u5E33\u865F\u7D81\u5B9A") {
      return handleFollowupEvent(event, { kind: "account_link_start", lineUid, eventId: eventId || event.message.id || "" }, env);
    }
    if (/^A\d{2,3}$/i.test(text)) {
      return handleFollowupEvent(event, {
        kind: "account_link_employee_id",
        lineUid,
        employeeId: text.toUpperCase(),
        eventId: eventId || event.message.id || ""
      }, env);
    }
    if (text === "\u555F\u7528\u6708\u66C6\u81EA\u52D5\u66F4\u65B0") {
      return handleFollowupEvent(event, { kind: "enable_weekly_calendar_menu", lineUid, eventId: eventId || event.message.id || "" }, env);
    }
    if (["\u500B\u4EBA\u73ED\u8868", "\u66F4\u65B0\u500B\u4EBA\u73ED\u8868", "\u66F4\u65B0\u500B\u4EBA\u6708\u66C6"].includes(text)) {
      return handleFollowupEvent(event, { kind: "personal_calendar_menu", lineUid, eventId: eventId || event.message.id || "" }, env);
    }
    const menuChoice = { "\u806F\u7D61\u5718\u54E1": "contact", "\u8ACB\u5047\u5BE9\u6838": "leave", "\u56DE\u8986\u78BA\u8A8D": "replies" }[text];
    const contact = /^團員留言\s+(\S{1,60})\s+([\s\S]{1,1500})$/.exec(text);
    const contactReply = /^團員回覆\s+([a-f0-9-]{36})\s+([\s\S]{1,1500})$/i.exec(text);
    if (typeof menuChoice === "string" || contact || contactReply) {
      return handleFollowupEvent(event, {
        kind: "calendar_menu_action",
        lineUid,
        eventId: eventId || event.message.id || "",
        choice: contactReply ? "reply_send" : contact ? "send" : menuChoice,
        value: (contactReply || contact)?.[1],
        text: (contactReply || contact)?.[2]
      }, env);
    }
    if (/^團員(?:留言|回覆)(?:\s|$)/.test(text))
      return replyLineSafe(event.replyToken, "\u8ACB\u5F9E\u5C0D\u8A71\u6309\u9215\u91CD\u65B0\u958B\u555F\u8F38\u5165\u6B04\uFF0C\u63A5\u8457\u8F38\u5165 1 \u81F3 1500 \u5B57\u3002", env);
    if (text === "\u7F3A\u73ED\u67E5\u8A62") {
      return handleFollowupEvent(event, { kind: "vacancy", choice: "list", lineUid, eventId: eventId || event.message.id || "" }, env);
    }
    const leaveReview = /^審核請假\s+(L-\d{6,})$/i.exec(text);
    if (leaveReview) {
      return handleFollowupEvent(event, {
        kind: "leave_review",
        lineUid,
        eventId: eventId || event.message.id || "",
        serial: leaveReview[1].toUpperCase()
      }, env);
    }
    if (/^新增替補(?:\s|$)/.test(text)) {
      const entry = /^新增替補\s+(L-\d{6,})\s+([^\r\n]{1,40}?)\s+(U[0-9a-f]{32})$/i.exec(text);
      if (!entry) return replyLineSafe(event.replyToken, "\u8ACB\u8F38\u5165\u300C\u65B0\u589E\u66FF\u88DC L-000001 \u59D3\u540D LINE UID\u300D\u3002", env);
      return handleFollowupEvent(event, {
        kind: "leave_replacement",
        lineUid,
        eventId: eventId || event.message.id || "",
        serial: entry[1].toUpperCase(),
        decision: "submit",
        name: entry[2],
        targetUid: entry[3]
      }, env);
    }
    if (text.normalize("NFKC").replace(/\s+/g, "").toLowerCase() === "我的uid") {
      await replyLine(event.replyToken, "\u60A8\u7684 LINE UID\uFF1A\n" + lineUid + "\n\n\u82E5\u8981\u63A5\u6536\u81E8\u6642\u6392\u73ED\uFF0C\u8ACB\u79C1\u4E0B\u63D0\u4F9B\u7D66\u5718\u9577\u3002\u8ACB\u52FF\u516C\u958B\u5F35\u8CBC\u3002", env);
      return;
    }
    if (/^排班留言(?:\s|$)/.test(text)) {
      await handleFollowupEvent(event, { kind: "shift_followup", lineUid, eventId: eventId || event.message.id || "", text }, env);
      return;
    }
    if (!text.startsWith("\u7D81\u5B9A") && !/^[^\r\n]+?\s*[A-Za-z0-9]{12}$/u.test(text)) {
      if (!text) return;
      await handleFollowupEvent(event, {
        kind: "contact_message",
        lineUid,
        text,
        timestamp: event.timestamp,
        eventId: eventId || event.message.id || ""
      }, env);
      return;
    }
    const confirmation = handleBindingMessage(event.replyToken, lineUid, text, eventId, env);
    ctx?.waitUntil(confirmation.catch(() => {
    }));
    await confirmation;
    return;
  }
  if (event.type === "postback") {
    const data = event.postback?.data;
    if (typeof data === "string" && data.startsWith("action=calendar_menu&") && data.length <= 300) {
      const params = new URLSearchParams(data);
      const choice = params.get("choice"), value = params.get("value") || "";
      if (choice === "keyboard") return;
      if (["contact", "leave", "replies", "leave_detail", "contact_target", "leave_pick", "contact_reply"].includes(choice)) {
        return handleFollowupEvent(event, { kind: "calendar_menu_action", lineUid, eventId, choice, value }, env);
      }
      return;
    }
    const vacancy = typeof data === "string" && /^action=vacancy&choice=(list|apply|review|approve|reject)&id=([a-f0-9:-]{1,60}|L-\d{6,})$/i.exec(data);
    if (vacancy) {
      return handleFollowupEvent(event, { kind: "vacancy", lineUid, eventId, choice: vacancy[1], id: vacancy[2] }, env);
    }
    const leaveReview = typeof data === "string" && /^action=leave_review&serial=(L-\d{6,})&decision=(approve|reject)$/.exec(data);
    if (leaveReview) {
      return handleFollowupEvent(event, { kind: "leave_review", lineUid, eventId, serial: leaveReview[1], decision: leaveReview[2] }, env);
    }
    if (typeof data === "string" && data.startsWith("action=leave_replacement&") && data.length <= 300) {
      const params = new URLSearchParams(data);
      const serial = params.get("serial"), decision = params.get("decision"), staffId = params.get("staff");
      const page = Number(params.get("page"));
      if (!/^L-\d{6,}$/.test(serial || "") || !["list", "skip", "select", "add", "retry"].includes(decision) || !Number.isSafeInteger(page) || page < 0 || decision === "select" && !staffId) return;
      return handleFollowupEvent(event, { kind: "leave_replacement", lineUid, eventId, serial, decision, staffId, page }, env);
    }
    const parsed = parsePostbackData(data);
    const followup = parsed ? null : parseShiftFollowupData(data, event.postback?.params?.date);
    const swap = parsed || followup ? null : parseSwapResponse(data);
    const manager = parsed || followup || swap ? null : parseManagerPostbackData(data);
    if (!parsed && !followup && !swap && !manager) return;
    const kind = parsed ? "shift_reply" : followup ? "shift_followup" : swap ? "swap_response" : "manager_decision";
    const payload = { kind, lineUid, eventId, timestamp: event.timestamp, ...parsed || followup || swap || manager };
    if (followup || swap) return handleFollowupEvent(event, payload, env);
    const result = await sendToGas(payload, env);
    const replied = await replyLineSafe(event.replyToken, result.message, env, result.quickReply, result.replyMessage);
    if (!replied && result.message) {
      if (result.replyMessage && await replyLineSafe(event.replyToken, result.message, env, result.quickReply)) return;
      throw new Error("LINE decision reply not delivered: " + result.code);
    }
  }
}
__name(handleEvent, "handleEvent");
async function handleFollowupEvent(event, payload, env) {
  if (payload.kind === "personal_calendar_menu" || payload.kind === "enable_weekly_calendar_menu") {
    await showCalendarLoading(payload.lineUid, env);
  }
  const result = await sendToGas(payload, env);
  if (result.delivered || !result.message) return;
  if (!await replyLineSafe(event.replyToken, result.message, env, result.quickReply)) {
    throw new Error("LINE follow-up reply not delivered: " + result.code);
  }
}
__name(handleFollowupEvent, "handleFollowupEvent");
// The existing rich menu remains linked until a replacement is uploaded successfully.
// Loading uses the originating OA token and never consumes the reply token.
async function showCalendarLoading(lineUid, env) {
  try {
    await fetch("https://api.line.me/v2/bot/chat/loading/start", {
      method: "POST",
      headers: {"content-type":"application/json", authorization:"Bearer "+env.LINE_CHANNEL_ACCESS_TOKEN},
      body: JSON.stringify({chatId:lineUid,loadingSeconds:60}),
      signal: AbortSignal.timeout(3000)
    });
  } catch (_) { /* Optional indicator must never block a menu refresh. */ }
}
async function handleBindingMessage(replyToken, lineUid, text, eventId, env) {
  const result = await sendToGas({ kind: "bind_or_message", lineUid, text, eventId }, env);
  const replied = await replyLineSafe(replyToken, result.message, env);
  if (result.code === "bound" && result.message && !replied) {
    await pushBindingConfirmation(lineUid, result.message, env);
  }
}
__name(handleBindingMessage, "handleBindingMessage");
async function pushBindingConfirmation(lineUid, text, env) {
  const headers = {
    "content-type": "application/json; charset=UTF-8",
    authorization: "Bearer " + env.LINE_CHANNEL_ACCESS_TOKEN,
    "x-line-retry-key": crypto.randomUUID()
  };
  const body = JSON.stringify({ to: lineUid, messages: [{ type: "text", text: String(text).slice(0, 1800) }] });
  for (let attempt = 0; attempt < 2; attempt++) {
    let response;
    try {
      response = await fetch("https://api.line.me/v2/bot/message/push", {
        method: "POST",
        headers,
        body,
        signal: AbortSignal.timeout(5e3)
      });
    } catch (error) {
      if (attempt === 0) continue;
      throw error;
    }
    if (response.ok || response.status === 409 && response.headers.has("x-line-accepted-request-id")) return;
    if (response.status < 500 || attempt === 1) {
      throw new Error("LINE binding confirmation push HTTP " + response.status);
    }
  }
}
__name(pushBindingConfirmation, "pushBindingConfirmation");
async function sendToGas(payload, env) {
  const response = await fetch(env.GAS_WEB_APP_URL, {
    method: "POST",
    headers: { "content-type": "application/json; charset=UTF-8" },
    body: JSON.stringify({
      ...payload,
      channel: env.MASALU_CHANNEL === "backup" ? "backup" : "primary",
      secret: env.GAS_SHARED_SECRET
    }),
    redirect: "follow"
  });
  if (!response.ok) throw new Error("GAS HTTP " + response.status);
  let result;
  try {
    result = await response.json();
  } catch {
    throw new Error("GAS response is not JSON");
  }
  if (!result || result.ok !== true) throw new Error("GAS rejected the request: " + String(result?.code || "unknown"));
  return result;
}
__name(sendToGas, "sendToGas");
async function replyLineSafe(replyToken, text, env, quickReply, replyMessage) {
  if (!replyToken || !text) return false;
  try {
    await replyLine(replyToken, text, env, quickReply, replyMessage);
    return true;
  } catch (error) {
    console.error("LINE confirmation reply failed:", String(error.message));
    return false;
  }
}
__name(replyLineSafe, "replyLineSafe");
async function replyLine(replyToken, text, env, quickReply, replyMessage) {
  if (!replyToken) return;
  const message = replyMessage || { type: "text", text: String(text).slice(0, 1800), ...quickReply ? { quickReply } : {} };
  const response = await fetch("https://api.line.me/v2/bot/message/reply", {
    method: "POST",
    headers: {
      "content-type": "application/json; charset=UTF-8",
      authorization: "Bearer " + env.LINE_CHANNEL_ACCESS_TOKEN
    },
    body: JSON.stringify({ replyToken, messages: [message] })
  });
  if (!response.ok) throw new Error("LINE reply HTTP " + response.status);
}
__name(replyLine, "replyLine");
export {
  workers_default as default,
  handleCalendarApi,
  parseManagerPostbackData,
  parsePostbackData,
  verifyLineSignature
};
//# sourceMappingURL=workers.js.map
