const byId = id => document.getElementById(id);
let tasks = [];
let polling = false;
let uploading = false;

async function call(message) {
  const response = await chrome.runtime.sendMessage(message);
  if (!response?.ok) throw new Error(response?.error || "操作失败");
  return response;
}

function selectedIds() {
  return new Set([...document.querySelectorAll('#tasks input[type="checkbox"]:checked')].map(input => input.value));
}

function renderTasks(items, selected = new Set()) {
  tasks = items;
  const container = byId("tasks");
  container.replaceChildren();
  for (const task of items) {
    const label = document.createElement("label");
    label.className = "task";
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.value = task.task_id;
    checkbox.checked = selected.has(task.task_id);
    const id = document.createElement("strong");
    id.textContent = task.task_id;
    const description = document.createElement("span");
    description.textContent = task.paper_name;
    const status = document.createElement("small");
    status.textContent = ` · ${task.status}`;
    description.appendChild(status);
    label.append(checkbox, id, description);
    container.appendChild(label);
  }
}

async function saveConnection() {
  const current = (await call({type: "status"})).settings;
  await call({type: "save-settings", settings: {
    baseUrl: byId("baseUrl").value.trim(),
    token: byId("token").value.trim(),
    taskIds: current.taskIds || ""
  }});
}

async function loadTasks({preserve = true} = {}) {
  const selected = preserve ? selectedIds() : new Set();
  const response = await call({type: "list-tasks"});
  renderTasks(response.state.tasks, selected);
}

async function refresh() {
  const response = await call({type: "status"});
  const settings = response.settings;
  byId("baseUrl").value = settings.baseUrl;
  byId("token").value = settings.token;
  byId("status").textContent = `${settings.running ? "运行中" : "已停止"} · ${settings.lastMessage}`;
  if (settings.token) await loadTasks({preserve: true});
}

async function uploadPdfs(fileList) {
  if (uploading) throw new Error("已有 PDF 正在上传");
  uploading = true;
  try {
    const files = [...fileList].filter(file => file.name.toLowerCase().endsWith(".pdf"));
    if (!files.length) throw new Error("请选择 PDF 文件");
    await saveConnection();
    const baseUrl = byId("baseUrl").value.trim();
    const token = byId("token").value.trim();
    const failures = [];
    let uploaded = 0;
    for (const file of files) {
      byId("status").textContent = `正在上传 ${uploaded + 1}/${files.length}：${file.name}`;
      try {
        const response = await fetch(`${baseUrl}/api/import-pdf?filename=${encodeURIComponent(file.name)}`, {
          method: "POST",
          headers: {"Content-Type": "application/pdf", "X-Workflow-Token": token},
          body: file
        });
        const payload = await response.json();
        if (!response.ok) throw new Error(payload.error || `HTTP ${response.status}`);
        uploaded += 1;
      } catch (error) {
        failures.push(`${file.name}：${error.message}`);
      }
    }
    await loadTasks({preserve: true});
    byId("status").textContent = failures.length
      ? `成功上传 ${uploaded} 个；失败 ${failures.length} 个：${failures.join("；")}`
      : `成功上传并新增 ${uploaded} 篇论文`;
  } finally {
    uploading = false;
  }
}

byId("loadTasks").addEventListener("click", async () => {
  try {
    await saveConnection();
    await loadTasks({preserve: true});
    byId("status").textContent = "任务列表已刷新";
  } catch (error) { byId("status").textContent = error.message; }
});

byId("openOutput").addEventListener("click", async () => {
  try {
    await saveConnection();
    const result = await call({type: "open-output"});
    byId("status").textContent = `已打开：${result.path}`;
  } catch (error) { byId("status").textContent = error.message; }
});

byId("copyAnswerToWord").addEventListener("click", async () => {
  try {
    const ids = [...selectedIds()];
    if (ids.length !== 1) throw new Error("请只勾选一个要归档回答的论文任务");
    await saveConnection();
    byId("status").textContent = `正在把当前 GPT 回答归档到任务 ${ids[0]}……`;
    const result = await call({type: "copy-answer-to-word", taskId: ids[0]});
    await loadTasks({preserve: false});
    byId("status").textContent = result.message;
  } catch (error) { byId("status").textContent = error.message; }
});

byId("selectIncomplete").addEventListener("click", () => {
  const wanted = new Set(tasks.filter(task => task.status !== "completed").map(task => task.task_id));
  const checkboxes = [...document.querySelectorAll('#tasks input[type="checkbox"]')];
  const shouldSelect = checkboxes.some(input => wanted.has(input.value) && !input.checked);
  for (const input of checkboxes) input.checked = shouldSelect && wanted.has(input.value);
  byId("selectIncomplete").textContent = shouldSelect ? "取消未完成" : "选择未完成";
});

byId("selectAll").addEventListener("click", () => {
  const checkboxes = [...document.querySelectorAll('#tasks input[type="checkbox"]')];
  const shouldSelect = checkboxes.some(input => !input.checked);
  for (const input of checkboxes) input.checked = shouldSelect;
  byId("selectAll").textContent = shouldSelect ? "取消全部" : "选择全部";
});

byId("applyStatus").addEventListener("click", async () => {
  const ids = [...selectedIds()];
  if (!ids.length) return void (byId("status").textContent = "请先勾选要修改状态的任务");
  const processingSelected = tasks.some(task => ids.includes(task.task_id) && task.status === "processing");
  if (processingSelected && !confirm("所选任务中包含正在处理的任务。修改后网页流程可能仍会继续，确定修改吗？")) return;
  try {
    const status = byId("manualStatus").value;
    await call({type: "set-task-status", taskIds: ids, status});
    await loadTasks({preserve: false});
    byId("status").textContent = `已将 ${ids.length} 个任务设为 ${status}`;
  } catch (error) { byId("status").textContent = error.message; }
});

byId("deleteTasks").addEventListener("click", async () => {
  const ids = [...selectedIds()];
  if (!ids.length) return void (byId("status").textContent = "请先勾选要删除的任务");
  const processingSelected = tasks.some(task => ids.includes(task.task_id) && task.status === "processing");
  const warning = processingSelected
    ? "所选任务包含正在处理的任务。删除后网页流程可能仍会继续。将永久删除 PDF、Word、输出目录和任务记录，确定继续吗？"
    : "将永久删除所选任务的 PDF、Word、输出目录和任务记录，确定继续吗？";
  if (!confirm(warning)) return;
  try {
    const result = await call({type: "delete-tasks", taskIds: ids});
    await loadTasks({preserve: false});
    const failed = Object.keys(result.failures || {});
    byId("status").textContent = failed.length
      ? `已删除 ${result.deleted.length} 个；${failed.length} 个失败：${failed.join(", ")}`
      : `已删除 ${result.deleted.length} 个任务及本地文件`;
  } catch (error) { byId("status").textContent = error.message; }
});

byId("start").addEventListener("click", async () => {
  try {
    const checked = selectedIds();
    const ids = tasks.filter(task => checked.has(task.task_id) && ["pending", "needs_review"].includes(task.status)).map(task => task.task_id);
    if (!ids.length) throw new Error("请至少选择一个可执行任务");
    await call({type: "save-settings", settings: {
      baseUrl: byId("baseUrl").value.trim(), token: byId("token").value.trim(), taskIds: ids.join(",")
    }});
    await call({type: "start"});
    await refresh();
  } catch (error) { byId("status").textContent = error.message; }
});

byId("stop").addEventListener("click", async () => {
  try { await call({type: "stop"}); await refresh(); }
  catch (error) { byId("status").textContent = error.message; }
});

const pdfPicker = byId("pdfPicker");
const dragOverlay = byId("dragOverlay");
let dragDepth = 0;
byId("choosePdfs").addEventListener("click", () => pdfPicker.click());
pdfPicker.addEventListener("change", () => {
  uploadPdfs(pdfPicker.files).catch(error => { byId("status").textContent = error.message; }).finally(() => { pdfPicker.value = ""; });
});
document.addEventListener("dragenter", event => {
  if (![...(event.dataTransfer?.types || [])].includes("Files")) return;
  event.preventDefault();
  dragDepth += 1;
  dragOverlay.classList.add("visible");
});
document.addEventListener("dragover", event => {
  if (![...(event.dataTransfer?.types || [])].includes("Files")) return;
  event.preventDefault();
  if (event.dataTransfer) event.dataTransfer.dropEffect = "copy";
});
document.addEventListener("dragleave", event => {
  event.preventDefault();
  dragDepth = Math.max(0, dragDepth - 1);
  if (!dragDepth) dragOverlay.classList.remove("visible");
});
document.addEventListener("drop", event => {
  event.preventDefault();
  dragDepth = 0;
  dragOverlay.classList.remove("visible");
  uploadPdfs(event.dataTransfer.files).catch(error => { byId("status").textContent = error.message; });
});

setInterval(async () => {
  if (polling || uploading || !byId("token").value.trim()) return;
  polling = true;
  try { await refresh(); } catch (_) {} finally { polling = false; }
}, 1500);

refresh().catch(error => { byId("status").textContent = error.message; });
