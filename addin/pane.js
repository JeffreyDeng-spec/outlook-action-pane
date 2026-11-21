"use strict";
const byId = id => document.getElementById(id);
function initializeOffice() {
  if (window.Office) void window.Office.onReady();
}
byId("office-runtime").addEventListener("load", initializeOffice, { once: true });
initializeOffice();
const snapshots = new Map();
const accountsWithErrors = new Set();
const labels = { verify: "查看验证邮件", interview: "查看面试安排", material: "查看提交要求", reply: "查看回复", action: "查看处理要求" };
let accountList = [];
let rendered = "";
let version = 0;
let loading = false;
let updating = false;

function el(tag, text, cls) {
  const node = document.createElement(tag);
  node.textContent = text;
  if (cls) node.className = cls;
  return node;
}
function outlookLink(text, raw) {
  if (!raw) return null;
  try {
    const url = new URL(raw);
    if (url.protocol !== "https:" || url.username || url.password || url.port ||
        !["outlook.live.com", "outlook.office.com", "outlook.office365.com", "outlook.cloud.microsoft"].includes(url.hostname)) return null;
    const a = el("a", text, "source-link");
    a.href = url.href;
    a.target = "_blank";
    a.rel = "noopener noreferrer";
    return a;
  } catch { return null; }
}
async function request(path, account, method = "GET") {
  const url = new URL(path, document.baseURI);
  if (account) url.searchParams.set("account", account);
  const options = { method, cache: "no-store", credentials: "same-origin", redirect: "error" };
  if (method === "POST") options.headers = { "X-Outlook-Triage": "1" };
  const response = await fetch(url, options);
  if (!response.ok) throw new Error("request failed");
  return response.json();
}
function formatSync(value) {
  if (!value) return "";
  const stamp = new Date(value);
  return Number.isNaN(stamp.getTime()) ? "" : stamp.toLocaleString(undefined, { month: "short", day: "numeric", hour: "2-digit", minute: "2-digit" });
}
function card(task) {
  const node = document.createElement("article");
  node.dataset.key = task.account + "/" + task.id;
  if (task.due) node.append(el("time", task.due));
  node.append(el("h2", task.title), el("small", [task.source, task.account].filter(Boolean).join(" · "), "meta"));
  const details = document.createElement("details");
  const summary = el("summary", labels[task.kind] || labels.action);
  details.append(summary);
  if (task.detail) details.append(el("p", task.detail, "detail"));
  const controls = el("div", "", "controls");
  for (const link of [outlookLink("打开原邮件", task.mail_url), outlookLink("打开草稿", task.draft_url)]) {
    if (link) controls.append(link);
  }
  for (const [command, text] of [["done", "标记已完成"], ["ignore", "无需处理"]]) {
    const button = el("button", text);
    button.onclick = () => mutate(task, command);
    controls.append(button);
  }
  details.append(controls);
  node.append(details);
  return node;
}
function render() {
  const tasks = [];
  for (const [account, data] of snapshots) for (const item of data.items) tasks.push({ ...item, account });
  tasks.sort((a,b) => Number(b.kind === "verify") - Number(a.kind === "verify") ||
    Number(Boolean(b.due)) - Number(Boolean(a.due)) ||
    (Date.parse(b.received) || 0) - (Date.parse(a.received) || 0) || a.id.localeCompare(b.id));
  const partial = !accountList.length || accountList.some(a =>
    accountsWithErrors.has(a) || !snapshots.get(a)?.sync_ok || snapshots.get(a)?.backfilling);
  byId("count").textContent = tasks.length ? tasks.length + " 项" : "";
  byId("empty").hidden = tasks.length > 0 || partial;
  const key = JSON.stringify(tasks);
  if (key !== rendered) {
    const open = new Set(Array.from(byId("items").querySelectorAll("details[open]"), d => d.parentNode.dataset.key));
    const nodes = tasks.map(card);
    for (const n of nodes) n.querySelector("details").open = open.has(n.dataset.key);
    byId("items").replaceChildren(...nodes);
    rendered = key;
  }
  const connection = byId("connection");
  connection.hidden = !partial;
  if (partial) {
    const messages = [];
    for (const account of accountList) {
      if (!accountsWithErrors.has(account) && snapshots.get(account)?.sync_ok) {
        if (snapshots.get(account)?.backfilling) messages.push(account + " 正在整理已有邮件");
        continue;
      }
      const last = formatSync(snapshots.get(account)?.last_sync);
      messages.push(account + " 同步中断" + (last ? "，上次成功同步 " + last : "，最新待办暂不可用"));
    }
    connection.textContent = messages.join("。") || "暂时无法连接分类服务，最新待办暂不可用。";
  }
}
async function refresh() {
  if (loading || updating) return;
  loading = true;
  const start = version;
  try {
    const { accounts } = await request("accounts");
    if (!Array.isArray(accounts) || !accounts.length) throw new Error("accounts unavailable");
    const results = await Promise.allSettled(accounts.map(a => request("actions", a)));
    if (start !== version) return;
    accountList = accounts;
    for (const a of snapshots.keys()) if (!accounts.includes(a)) snapshots.delete(a);
    results.forEach((result,i) => {
      const account = accounts[i];
      const data = result.status === "fulfilled" ? result.value : null;
      if (data?.account === account && Array.isArray(data.items)) {
        snapshots.set(account, data);
        accountsWithErrors.delete(account);
      } else {
        accountsWithErrors.add(account);
      }
    });
  } catch {
    if (start === version) for (const a of accountList) accountsWithErrors.add(a);
  } finally {
    loading = false;
    if (start === version) render();
    else if (!updating) void refresh();
  }
}
async function mutate(task, command) {
  if (updating) return;
  updating = true;
  version++;
  const feedback = byId("feedback");
  try {
    await request("actions/" + encodeURIComponent(task.id) + "/" + command, task.account, "POST");
    const data = snapshots.get(task.account);
    if (data && command !== "undo") data.items = data.items.filter(t => t.id !== task.id);
    render();
    feedback.replaceChildren(el("span", {done:"已标记完成", ignore:"已移出待办", undo:"已撤销"}[command]));
    if (command !== "undo") {
      const undo = el("button", "撤销");
      undo.onclick = () => mutate(task, "undo");
      feedback.append(undo);
    }
  } catch {
    feedback.textContent = "未能确认操作结果，请刷新后查看。";
  } finally {
    updating = false;
    void refresh();
  }
}
void refresh();
window.setInterval(() => { if (!document.hidden) void refresh(); }, 15000);
window.addEventListener("focus", refresh);
document.addEventListener("visibilitychange", () => { if (!document.hidden) void refresh(); });
