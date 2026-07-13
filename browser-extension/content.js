const sleep = milliseconds => new Promise(resolve => setTimeout(resolve, milliseconds));

async function message(payload) {
  const response = await chrome.runtime.sendMessage(payload);
  if (!response?.ok) throw new Error(response?.error || "扩展后台请求失败");
  return response;
}

async function waitFor(selector, timeout = 30000) {
  const deadline = Date.now() + timeout;
  while (Date.now() < deadline) {
    const element = document.querySelector(selector);
    if (element) return element;
    await sleep(500);
  }
  throw new Error(`等待页面元素超时：${selector}`);
}

function setComposerText(composer, text) {
  composer.focus();
  const selection = window.getSelection();
  const range = document.createRange();
  range.selectNodeContents(composer);
  selection.removeAllRanges();
  selection.addRange(range);
  const inserted = document.execCommand("insertText", false, text);
  if (!inserted || composer.innerText.replace(/\r/g, "").trim() !== text.replace(/\r/g, "").trim()) {
    composer.replaceChildren();
    for (const line of text.replace(/\r\n/g, "\n").split("\n")) {
      const paragraph = document.createElement("p");
      if (line) paragraph.textContent = line;
      else paragraph.appendChild(document.createElement("br"));
      composer.appendChild(paragraph);
    }
    composer.dispatchEvent(new InputEvent("input", {bubbles: true, inputType: "insertText", data: text}));
  }
}

async function attachPdf(task) {
  const input = await waitFor("#upload-files");
  const response = await message({type: "fetch-pdf", url: task.pdf_url});
  const binary = atob(response.base64);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
  const file = new File([bytes], task.paper_name, {type: "application/pdf"});
  const transfer = new DataTransfer();
  transfer.items.add(file);
  input.files = transfer.files;
  input.dispatchEvent(new Event("change", {bubbles: true}));

  const deadline = Date.now() + 120000;
  while (Date.now() < deadline) {
    const pageText = document.body.innerText;
    if (pageText.includes(task.paper_name)) return;
    await sleep(1000);
  }
  throw new Error("PDF 已注入上传控件，但页面未显示附件完成状态");
}

async function sendPrompt(task) {
  const composer = await waitFor('[contenteditable="true"][data-virtualkeyboard="true"], #prompt-textarea');
  setComposerText(composer, task.request_text);
  await sleep(500);
  const sendButton = document.querySelector('[data-testid="send-button"]');
  if (!sendButton || sendButton.disabled) throw new Error("发送按钮不可用");
  sendButton.click();
}

async function waitForAnswer() {
  const deadline = Date.now() + 45 * 60 * 1000;
  let generationObserved = false;
  let stableSince = 0;
  let previous = "";
  while (Date.now() < deadline) {
    const stopButton = document.querySelector('[data-testid="stop-button"], button[aria-label*="停止"]');
    if (stopButton) generationObserved = true;
    const assistants = document.querySelectorAll('[data-message-author-role="assistant"]');
    const latest = assistants.length ? assistants[assistants.length - 1].innerText.trim() : "";
    if (latest && latest === previous && !stopButton && generationObserved) {
      if (!stableSince) stableSince = Date.now();
      if (Date.now() - stableSince > 5000) {
        const assistant = assistants[assistants.length - 1];
        const formatted = assistant.querySelector(".markdown, .prose, [class*='markdown']") || assistant;
        return {text: latest, html: formatted.innerHTML, assistant};
      }
    } else {
      stableSince = 0;
      previous = latest;
    }
    await sleep(1000);
  }
  throw new Error("等待 GPT 回答完成超时");
}

async function readCopiedAnswer(answer) {
  const assistant = answer.assistant;
  const turn = assistant.closest('article, [data-testid^="conversation-turn-"]') || assistant.parentElement;
  const scope = turn || assistant;
  let copyButton = scope.querySelector('[data-testid="copy-turn-action-button"]');
  if (!copyButton) {
    const allowedLabels = new Set(["copy", "copy response", "复制", "复制回答"]);
    const candidates = [...scope.querySelectorAll("button[aria-label]")].filter(button =>
      allowedLabels.has((button.getAttribute("aria-label") || "").trim().toLowerCase())
    );
    copyButton = candidates.length ? candidates[candidates.length - 1] : null;
  }
  if (!copyButton) {
    throw new Error("未找到当前 GPT 回答的复制按钮；已禁止 HTML 回退");
  }

  copyButton.scrollIntoView({block: "center", inline: "center"});
  await sleep(300);
  const rect = copyButton.getBoundingClientRect();
  if (!rect.width || !rect.height) {
    throw new Error("当前 GPT 回答的复制按钮不可见");
  }
  const canonical = text => text.normalize("NFKC").toLowerCase().replace(/[\p{P}\p{S}\s]+/gu, "");
  const expected = canonical(answer.text);
  const clipboardMatches = copiedText => {
    const actual = canonical(copiedText);
    if (expected.length < 100 || actual.length < expected.length * 0.55) return actual === expected;
    if (actual.includes(expected.slice(0, 100))) return true;
    const positions = [0, 0.25, 0.5, 0.75];
    const matched = positions.filter(position => {
      const start = Math.floor((expected.length - 80) * position);
      return actual.includes(expected.slice(start, start + 80));
    }).length;
    return matched >= 2;
  };

  let clipboardError = "";
  for (let clickAttempt = 0; clickAttempt < 3; clickAttempt += 1) {
    await message({type: "set-message", text: `${answer.assistant ? "正在切换到 ChatGPT 前台并复制回答" : "正在复制回答"}（${clickAttempt + 1}/3）`});
    await message({
      type: "trusted-click",
      x: rect.left + rect.width / 2,
      y: rect.top + rect.height / 2
    });
    for (let poll = 0; poll < 20; poll += 1) {
      await sleep(500);
      try {
        const copiedText = await navigator.clipboard.readText();
        if (clipboardMatches(copiedText)) {
          return {
            text: answer.text,
            source: "chatgpt_copy_button",
            clipboardVerified: true
          };
        }
      } catch (error) {
        clipboardError = error.message;
        break;
      }
    }
  }
  throw new Error(clipboardError
    ? `无法读取复制后的剪贴板：${clipboardError}`
    : "复制回答后三次校验均不匹配当前回答，已拒绝启动 Word");
}

async function processCurrentTask(task) {
  await message({type: "set-message", text: `${task.task_id}: 正在上传 PDF`});
  await attachPdf(task);
  await message({type: "set-message", text: `${task.task_id}: 正在发送共享提示词`});
  await sendPrompt(task);
  const answer = await waitForAnswer();
  const response = await readCopiedAnswer(answer);
  await message({
    type: "complete-task",
    taskId: task.task_id,
    response: response.text,
    responseSource: response.source,
    clipboardVerified: response.clipboardVerified === true,
    chatUrl: location.href
  });
}

let queueBusy = false;

async function runQueueCore() {
  const state = await chrome.storage.local.get({running: false, currentTask: null});
  if (!state.running) return;
  let task = state.currentTask;
  if (!task) task = (await message({type: "next-task"})).task;
  if (task.done) {
    await chrome.storage.local.set({running: false, currentTask: null, lastMessage: "全部指定任务已完成"});
    return;
  }
  if (location.pathname !== "/") {
    location.href = "https://chatgpt.com/";
    return;
  }
  try {
    await processCurrentTask(task);
    location.href = "https://chatgpt.com/";
  } catch (error) {
    await message({type: "fail-task", taskId: task.task_id, error: error.message, chatUrl: location.href});
    const current = await chrome.storage.local.get({running: false});
    if (current.running) location.href = "https://chatgpt.com/";
  }
}

async function runQueue() {
  if (queueBusy) return;
  queueBusy = true;
  try {
    await runQueueCore();
  } finally {
    queueBusy = false;
  }
}

chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.type === "ping") {
    sendResponse({ok: true});
    return false;
  }
  if (request.type === "run-queue" || request.type === "resume-task") {
    runQueue().then(() => sendResponse({ok: true})).catch(error => sendResponse({ok: false, error: error.message}));
    return true;
  }
  return false;
});

message({type: "should-run-here"})
  .then(result => { if (result.shouldRun) runQueue().catch(() => {}); })
  .catch(() => {});
