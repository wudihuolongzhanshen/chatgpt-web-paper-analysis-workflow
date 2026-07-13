const DEFAULT_BASE_URL = "http://127.0.0.1:8765";

async function settings() {
  return chrome.storage.local.get({
    baseUrl: DEFAULT_BASE_URL,
    token: "",
    taskIds: "001,002",
    running: false,
    currentTask: null,
    attemptedTaskIds: [],
    controlledTabId: null,
    lastMessage: "尚未启动"
  });
}

async function api(path, options = {}) {
  const config = await settings();
  if (!config.token) throw new Error("尚未设置本地服务令牌");
  const response = await fetch(`${config.baseUrl}${path}`, {
    ...options,
    headers: {
      "Content-Type": "application/json",
      "X-Workflow-Token": config.token,
      ...(options.headers || {})
    }
  });
  const payload = await response.json();
  if (!response.ok) throw new Error(payload.error || `HTTP ${response.status}`);
  return payload;
}

async function ensureContentScript(tabId) {
  try {
    await chrome.tabs.sendMessage(tabId, {type: "ping"});
  } catch (error) {
    await chrome.scripting.executeScript({target: {tabId}, files: ["content.js"]});
  }
}

async function sendToChatTab(message) {
  const tabs = await chrome.tabs.query({url: "https://chatgpt.com/*"});
  let tab = tabs.find(item => item.active) || tabs[0];
  if (!tab) {
    tab = await chrome.tabs.create({url: "https://chatgpt.com/"});
    await new Promise(resolve => setTimeout(resolve, 2000));
  }
  await chrome.storage.local.set({controlledTabId: tab.id});
  await ensureContentScript(tab.id);
  await chrome.tabs.sendMessage(tab.id, message);
  return tab.id;
}

async function trustedClick(tabId, x, y) {
  const tab = await chrome.tabs.get(tabId);
  if (typeof tab.windowId === "number") {
    const window = await chrome.windows.get(tab.windowId);
    if (window.state === "minimized") {
      await chrome.windows.update(tab.windowId, {state: "normal"});
    }
    await chrome.tabs.update(tabId, {active: true});
    await chrome.windows.update(tab.windowId, {focused: true});
    // Clipboard operations require a focused document. Give Chrome enough
    // time to transfer focus before dispatching the trusted mouse event.
    await new Promise(resolve => setTimeout(resolve, 500));
  }
  const target = {tabId};
  await chrome.debugger.attach(target, "1.3");
  try {
    await chrome.debugger.sendCommand(target, "Input.dispatchMouseEvent", {
      type: "mouseMoved", x, y
    });
    await chrome.debugger.sendCommand(target, "Input.dispatchMouseEvent", {
      type: "mousePressed", x, y, button: "left", clickCount: 1
    });
    await chrome.debugger.sendCommand(target, "Input.dispatchMouseEvent", {
      type: "mouseReleased", x, y, button: "left", clickCount: 1
    });
  } finally {
    await chrome.debugger.detach(target).catch(() => {});
  }
}

async function openDashboard() {
  const dashboardUrl = chrome.runtime.getURL("popup.html");
  const windows = await chrome.windows.getAll({populate: true});
  const existing = windows.find(window =>
    window.tabs?.some(tab => tab.url?.startsWith(dashboardUrl))
  );
  if (existing?.id) {
    await chrome.windows.update(existing.id, {focused: true});
    return;
  }
  const displays = await chrome.system.display.getInfo();
  const display = displays.find(item => item.isPrimary) || displays[0];
  const workArea = display?.workArea || {left: 0, top: 0, width: 720, height: 1000};
  const width = Math.min(720, workArea.width);
  const height = Math.max(700, workArea.height - 20);
  await chrome.windows.create({
    url: dashboardUrl,
    type: "popup",
    width,
    height,
    left: workArea.left + Math.max(0, workArea.width - width),
    top: workArea.top,
    focused: true
  });
}

chrome.action.onClicked.addListener(() => {
  openDashboard().catch(error => chrome.storage.local.set({lastMessage: `打开任务窗口失败：${error.message}`}));
});

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  (async () => {
    if (message.type === "save-settings") {
      await chrome.storage.local.set(message.settings);
      sendResponse({ok: true});
    } else if (message.type === "start") {
      const previous = await settings();
      if (previous.currentTask?.task_id) {
        await api(`/api/tasks/${previous.currentTask.task_id}/release`, {method: "POST", body: "{}"}).catch(() => {});
      }
      // Never resume request_text cached by an older extension/server version.
      // A new start must fetch the current shared prompt from local_server.py.
      await chrome.storage.local.set({running: true, currentTask: null, attemptedTaskIds: [], lastMessage: "正在读取最新共享提示词"});
      await sendToChatTab({type: "run-queue"});
      sendResponse({ok: true});
    } else if (message.type === "stop") {
      const previous = await settings();
      if (previous.currentTask?.task_id) {
        await api(`/api/tasks/${previous.currentTask.task_id}/release`, {method: "POST", body: "{}"}).catch(() => {});
      }
      await chrome.storage.local.set({running: false, currentTask: null, controlledTabId: null, lastMessage: "已停止"});
      sendResponse({ok: true});
    } else if (message.type === "status") {
      sendResponse({ok: true, settings: await settings()});
    } else if (message.type === "list-tasks") {
      sendResponse({ok: true, state: await api("/api/state")});
    } else if (message.type === "set-task-status") {
      const result = await api("/api/status", {
        method: "POST",
        body: JSON.stringify({task_ids: message.taskIds, status: message.status})
      });
      sendResponse({ok: true, ...result});
    } else if (message.type === "delete-tasks") {
      const result = await api("/api/delete-tasks", {
        method: "POST",
        body: JSON.stringify({task_ids: message.taskIds})
      });
      sendResponse({ok: true, ...result});
    } else if (message.type === "should-run-here") {
      const config = await settings();
      sendResponse({ok: true, shouldRun: Boolean(config.running && sender.tab?.id === config.controlledTabId)});
    } else if (message.type === "trusted-click") {
      if (!sender.tab?.id) throw new Error("无法确定 ChatGPT 标签页");
      await trustedClick(sender.tab.id, Number(message.x), Number(message.y));
      sendResponse({ok: true});
    } else if (message.type === "next-task") {
      const config = await settings();
      const ids = encodeURIComponent(config.taskIds || "");
      const excluded = encodeURIComponent((config.attemptedTaskIds || []).join(","));
      const task = await api(`/api/next?ids=${ids}&exclude=${excluded}`);
      const attemptedTaskIds = task.done
        ? config.attemptedTaskIds
        : [...new Set([...(config.attemptedTaskIds || []), task.task_id])];
      await chrome.storage.local.set({currentTask: task.done ? null : task, attemptedTaskIds});
      sendResponse({ok: true, task});
    } else if (message.type === "fetch-pdf") {
      const config = await settings();
      const response = await fetch(`${message.url}?token=${encodeURIComponent(config.token)}`);
      if (!response.ok) throw new Error(`PDF 获取失败：HTTP ${response.status}`);
      const buffer = await response.arrayBuffer();
      const bytes = new Uint8Array(buffer);
      let binary = "";
      const chunk = 0x8000;
      for (let index = 0; index < bytes.length; index += chunk) {
        binary += String.fromCharCode(...bytes.subarray(index, index + chunk));
      }
      sendResponse({ok: true, base64: btoa(binary)});
    } else if (message.type === "complete-task") {
      const result = await api(`/api/tasks/${message.taskId}/complete`, {
        method: "POST",
        body: JSON.stringify({
          response: message.response,
          response_source: message.responseSource,
          clipboard_verified: message.clipboardVerified === true,
          chat_url: message.chatUrl
        })
      });
      await chrome.storage.local.set({currentTask: null, lastMessage: `${message.taskId}: ${result.status}`});
      sendResponse({ok: true, result});
    } else if (message.type === "fail-task") {
      await api(`/api/tasks/${message.taskId}/fail`, {
        method: "POST",
        body: JSON.stringify({error: message.error, chat_url: message.chatUrl})
      });
      await chrome.storage.local.set({currentTask: null, lastMessage: `${message.taskId}: 失败`});
      sendResponse({ok: true});
    } else if (message.type === "set-message") {
      await chrome.storage.local.set({lastMessage: message.text});
      sendResponse({ok: true});
    }
  })().catch(error => sendResponse({ok: false, error: error.message}));
  return true;
});

chrome.tabs.onUpdated.addListener((tabId, changeInfo, tab) => {
  if (changeInfo.status !== "complete" || !tab.url?.startsWith("https://chatgpt.com/")) return;
  chrome.storage.local.get({running: false, controlledTabId: null}, async state => {
    if (state.running && tabId === state.controlledTabId) {
      try {
        await ensureContentScript(tabId);
        await chrome.tabs.sendMessage(tabId, {type: "resume-task"});
      } catch (error) {
        await chrome.storage.local.set({lastMessage: `页面恢复失败：${error.message}`});
      }
    }
  });
});
