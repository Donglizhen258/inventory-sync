const elements = {
  form: document.querySelector("#inventoryForm"),
  material: document.querySelector("#material"),
  materialSuggestions: document.querySelector("#materialSuggestions"),
  materialHint: document.querySelector("#materialHint"),
  specification: document.querySelector("#specification"),
  location: document.querySelector("#location"),
  locationSuggestions: document.querySelector("#locationSuggestions"),
  shelfLetter: document.querySelector("#shelfLetter"),
  shelfNumber: document.querySelector("#shelfNumber"),
  requester: document.querySelector("#requester"),
  requesterSuggestions: document.querySelector("#requesterSuggestions"),
  project: document.querySelector("#project"),
  projectSuggestions: document.querySelector("#projectSuggestions"),
  po: document.querySelector("#po"),
  changedAt: document.querySelector("#changedAt"),
  quantity: document.querySelector("#quantity"),
  changeType: document.querySelector("#changeType"),
  note: document.querySelector("#note"),
  balanceBefore: document.querySelector("#balanceBefore"),
  balanceAfter: document.querySelector("#balanceAfter"),
  clearButton: document.querySelector("#clearButton"),
  submitButton: document.querySelector("#submitButton"),
  nowButton: document.querySelector("#nowButton"),
  refreshButton: document.querySelector("#refreshButton"),
  shutdownButton: document.querySelector("#shutdownButton"),
  recentBody: document.querySelector("#recentBody"),
  toast: document.querySelector("#toast"),
  dialogBackdrop: document.querySelector("#dialogBackdrop"),
  dialog: document.querySelector(".dialog"),
  dialogIcon: document.querySelector("#dialogIcon"),
  dialogTitle: document.querySelector("#dialogTitle"),
  dialogMessage: document.querySelector("#dialogMessage"),
  dialogClose: document.querySelector("#dialogClose"),
  loginOverlay: document.querySelector("#loginOverlay"),
  loginForm: document.querySelector("#loginForm"),
  loginPassword: document.querySelector("#loginPassword"),
  loginButton: document.querySelector("#loginButton"),
  loginHint: document.querySelector("#loginHint"),
  serverSettingLink: document.querySelector("#serverSettingLink"),
  serverSettingOverlay: document.querySelector("#serverSettingOverlay"),
  currentServerUrl: document.querySelector("#currentServerUrl"),
  discoverButton: document.querySelector("#discoverButton"),
  discoverResults: document.querySelector("#discoverResults"),
  serverUrlInput: document.querySelector("#serverUrlInput"),
  serverHistorySelect: document.querySelector("#serverHistorySelect"),
  saveServerUrlButton: document.querySelector("#saveServerUrlButton"),
  closeServerSettingButton: document.querySelector("#closeServerSettingButton"),
  roleBadge: document.querySelector("#roleBadge"),
  syncBadge: document.querySelector("#syncBadge"),
  syncSummary: document.querySelector("#syncSummary"),
  syncPullButton: document.querySelector("#syncPullButton"),
  syncPushButton: document.querySelector("#syncPushButton"),
  syncRefreshButton: document.querySelector("#syncRefreshButton"),
  versionsBody: document.querySelector("#versionsBody"),
  rollbackNote: document.querySelector("#rollbackNote"),
  adminCard: document.querySelector("#adminCard"),
  adminRefreshButton: document.querySelector("#adminRefreshButton"),
  addUserForm: document.querySelector("#addUserForm"),
  newUserPassword: document.querySelector("#newUserPassword"),
  newUserNote: document.querySelector("#newUserNote"),
  addUserButton: document.querySelector("#addUserButton"),
  usersBody: document.querySelector("#usersBody"),
  usersRefreshButton: document.querySelector("#usersRefreshButton"),
  addProjectForm: document.querySelector("#addProjectForm"),
  newProjectName: document.querySelector("#newProjectName"),
  addProjectButton: document.querySelector("#addProjectButton"),
  projectsBody: document.querySelector("#projectsBody"),
  changePasswordForm: document.querySelector("#changePasswordForm"),
  oldAdminPassword: document.querySelector("#oldAdminPassword"),
  newAdminPassword: document.querySelector("#newAdminPassword"),
  changePasswordButton: document.querySelector("#changePasswordButton"),
  downloadButton: document.querySelector("#downloadButton"),
  replaceButton: document.querySelector("#replaceButton"),
  replaceInput: document.querySelector("#replaceInput"),
};

const DEFAULT_REQUESTERS = [
  "张三",
  "李四",
  "王五",
  "赵六",
  "钱七",
  "孙八",
  "周九",
  "吴十",
  "郑十一",
  "王小明",
];
const DEFAULT_PROJECTS = ["项目A", "项目B", "项目C", "项目D"];
const SHELF_LETTERS = ["A", "B", "C", "D", "E", "F", "G", "H"];
const SHELF_NUMBER_LIMITS = { A: 4, B: 4, C: 4, D: 4, E: 4, F: 4, G: 4, H: 4 };
const DEFAULT_LOCATIONS = SHELF_LETTERS.flatMap((letter) =>
  Array.from(
    { length: SHELF_NUMBER_LIMITS[letter] },
    (_, index) => `货架${letter}${index + 1}`,
  ),
);

let selectedMaterial = null;
let materialCandidates = [];
let activeCandidateIndex = -1;
let materialSearchTimer;
let materialRequestController;
let materialSearchSequence = 0;
let balanceTimer;
let balanceRequestController;
let balanceRequestSequence = 0;
let toastTimer;

function localDate() {
  const now = new Date();
  const pad = (value) => String(value).padStart(2, "0");
  return `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}`;
}

function formatNumber(value) {
  const number = Number(value);
  if (!Number.isFinite(number)) return "—";
  return new Intl.NumberFormat("zh-CN", {
    minimumFractionDigits: 0,
    maximumFractionDigits: 0,
  }).format(number);
}

async function fetchJson(url, options = {}) {
  const response = await fetch(url, {
    ...options,
    headers: {
      "Content-Type": "application/json",
      ...(options.headers || {}),
    },
  });
  let payload;
  try {
    payload = await response.json();
  } catch {
    throw new Error("程序没有返回有效响应。");
  }
  if (!response.ok || payload.ok === false) {
    const error = new Error(payload.error || "操作失败。");
    error.conflict = Boolean(payload.conflict);
    throw error;
  }
  return payload;
}

function fillSelect(select, values) {
  const safeValues = Array.isArray(values) ? values : [];
  const previous = select.value;
  select.replaceChildren();
  const placeholder = document.createElement("option");
  placeholder.value = "";
  placeholder.textContent = "请选择";
  placeholder.disabled = true;
  select.append(placeholder);
  safeValues.forEach((value) => {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = value;
    select.append(option);
  });
  select.value = safeValues.includes(previous) ? previous : "";
}

function fillDatalist(input, datalist, values, defaults) {
  const safeValues = Array.isArray(values) ? values : [];
  const suggestions = [...new Set([...defaults, ...safeValues].filter(Boolean))];
  datalist.replaceChildren();
  suggestions.forEach((value) => {
    const option = document.createElement("option");
    option.value = value;
    datalist.append(option);
  });
}

function fillShelfNumbers(letter, selected = "") {
  const maximum = SHELF_NUMBER_LIMITS[letter] || 0;
  elements.shelfNumber.replaceChildren();
  const placeholder = document.createElement("option");
  placeholder.value = "";
  placeholder.textContent = letter ? "选编号" : "先选字母";
  placeholder.disabled = true;
  elements.shelfNumber.append(placeholder);
  for (let number = 1; number <= maximum; number += 1) {
    const option = document.createElement("option");
    option.value = String(number);
    option.textContent = String(number);
    elements.shelfNumber.append(option);
  }
  elements.shelfNumber.disabled = !letter;
  elements.shelfNumber.value =
    selected && Number(selected) <= maximum ? String(selected) : "";
}

function syncShelfFromLocation() {
  const match = elements.location.value.trim().match(/^货架([A-H])([1-4])$/);
  if (!match || Number(match[2]) > SHELF_NUMBER_LIMITS[match[1]]) {
    elements.shelfLetter.value = "";
    fillShelfNumbers("");
    return;
  }
  elements.shelfLetter.value = match[1];
  fillShelfNumbers(match[1], match[2]);
}

function syncLocationFromShelf() {
  const letter = elements.shelfLetter.value;
  const number = elements.shelfNumber.value;
  elements.location.value = letter && number ? `货架${letter}${number}` : "";
}

function showToast(message) {
  clearTimeout(toastTimer);
  elements.toast.textContent = message;
  elements.toast.classList.add("show");
  toastTimer = setTimeout(() => elements.toast.classList.remove("show"), 2800);
}

function showDialog(title, message, error = false, html = false) {
  elements.dialogTitle.textContent = title;
  elements.dialogMessage[html ? "innerHTML" : "textContent"] = message;
  elements.dialogIcon.textContent = error ? "!" : "✓";
  elements.dialog.classList.toggle("error", error);
  elements.dialogBackdrop.hidden = false;
  elements.dialogClose.focus();
}

function closeDialog() {
  elements.dialogBackdrop.hidden = true;
}

function promptSavePath(defaultPath) {
  // 纯 HTML 输入框弹窗：返回用户输入的路径；取消返回 null
  return new Promise((resolve) => {
    const wrap = document.createElement("div");
    wrap.style.cssText =
      "position:fixed;inset:0;background:rgba(0,0,0,.45);display:flex;" +
      "align-items:center;justify-content:center;z-index:9999;";
    const box = document.createElement("div");
    box.style.cssText =
      "background:#f4f8fc;border:1px solid #8fb7dd;border-radius:6px;" +
      "padding:20px 22px;width:460px;box-shadow:0 4px 16px rgba(0,0,0,.3);";
    box.innerHTML =
      '<div style="font-size:15px;font-weight:bold;color:#1d66b8;margin-bottom:10px;">选择保存位置</div>' +
      '<div style="font-size:12px;color:#555;margin-bottom:6px;">文件夹路径（留空保存到桌面）：</div>' +
      '<input id="savePathInput" type="text" style="width:100%;box-sizing:border-box;padding:6px 8px;' +
      'border:1px solid #b8cfe0;border-radius:3px;font-size:13px;" />' +
      '<div style="margin-top:14px;text-align:right;">' +
      '<button id="savePathCancel" style="padding:6px 16px;border:1px solid #b8cfe0;background:#eef4fa;' +
      'border-radius:3px;cursor:pointer;font-size:13px;">取消</button> ' +
      '<button id="savePathOk" style="padding:6px 16px;border:1px solid #1d66b8;background:#2d7dd8;color:#fff;' +
      'border-radius:3px;cursor:pointer;font-size:13px;">保存到此位置</button>' +
      '</div>';
    wrap.append(box);
    document.body.append(wrap);
    const input = box.querySelector("#savePathInput");
    input.value = defaultPath || "";
    const done = (value) => {
      wrap.remove();
      resolve(value);
    };
    box.querySelector("#savePathOk").addEventListener("click", () => done(input.value.trim()));
    box.querySelector("#savePathCancel").addEventListener("click", () => done(null));
    input.addEventListener("keydown", (event) => {
      if (event.key === "Enter") done(input.value.trim());
      if (event.key === "Escape") done(null);
    });
    input.focus();
    input.select();
  });
}

function setBusy(button, busy, busyText) {
  if (!button.dataset.label) button.dataset.label = button.textContent;
  button.disabled = busy;
  button.textContent = busy ? busyText : button.dataset.label;
}

function normalizeMaterial(raw) {
  return {
    materialCode: String(
      raw.materialCode ?? raw.material_code ?? raw.code ?? "",
    ).trim(),
    material: String(
      raw.material ?? raw.material_name ?? raw.name ?? "",
    ).trim(),
    specification: String(
      raw.specification ?? raw.spec ?? raw.model ?? "",
    ).trim(),
    balance:
      raw.balance ??
      raw.currentBalance ??
      raw.current_balance ??
      raw.stock ??
      0,
    requester: String(
      raw.requester ?? raw.lastRequester ?? raw.last_requester ?? "",
    ).trim(),
    project: String(
      raw.project ?? raw.lastProject ?? raw.last_project ?? "",
    ).trim(),
    po: String(raw.po ?? raw.lastPo ?? raw.last_po ?? "").trim(),
    location: String(
      raw.location ?? raw.storageLocation ?? raw.storage_location ?? "",
    ).trim(),
  };
}

function setMaterialHint(text, selected = false) {
  elements.materialHint.textContent = text;
  elements.materialHint.classList.toggle("selected", selected);
}

function closeMaterialSuggestions() {
  elements.materialSuggestions.hidden = true;
  elements.material.setAttribute("aria-expanded", "false");
  elements.material.removeAttribute("aria-activedescendant");
  activeCandidateIndex = -1;
}

function setActiveCandidate(index) {
  if (!materialCandidates.length) return;
  activeCandidateIndex =
    (index + materialCandidates.length) % materialCandidates.length;
  const options = [...elements.materialSuggestions.querySelectorAll(".material-option")];
  options.forEach((option, optionIndex) => {
    const active = optionIndex === activeCandidateIndex;
    option.classList.toggle("active", active);
    option.setAttribute("aria-selected", String(active));
    if (active) {
      elements.material.setAttribute("aria-activedescendant", option.id);
      option.scrollIntoView({ block: "nearest" });
    }
  });
}

function renderMaterialSuggestions(candidates, query) {
  materialCandidates = candidates;
  activeCandidateIndex = -1;
  elements.materialSuggestions.replaceChildren();

  if (!query) {
    closeMaterialSuggestions();
    return;
  }

  if (!candidates.length) {
    const empty = document.createElement("div");
    empty.className = "autocomplete-empty";
    empty.textContent = "没有匹配物料；提交后将作为新材料自动生成物料编号。";
    elements.materialSuggestions.append(empty);
    setMaterialHint("没有匹配物料，将作为新材料自动编号。");
  } else {
    candidates.forEach((candidate, index) => {
      const option = document.createElement("button");
      option.type = "button";
      option.id = `material-option-${index}`;
      option.className = "material-option";
      option.setAttribute("role", "option");
      option.setAttribute("aria-selected", "false");
      option.textContent = `${candidate.materialCode || "待编号"}｜${
        candidate.material || "未命名"
      }｜${candidate.specification || "—"}｜项目 ${
        candidate.project || "—"
      }｜位置 ${candidate.location || "—"}｜库存 ${formatNumber(candidate.balance)}`;
      option.title = option.textContent;
      option.addEventListener("mouseenter", () => setActiveCandidate(index));
      option.addEventListener("mousedown", (event) => {
        event.preventDefault();
        selectMaterial(candidate);
      });
      elements.materialSuggestions.append(option);
    });
    setMaterialHint(`找到 ${candidates.length} 项，请选择准确的物料。`);
  }

  elements.materialSuggestions.hidden = false;
  elements.material.setAttribute("aria-expanded", "true");
}

function clearSelectedMaterial(keepValues = true) {
  selectedMaterial = null;
  if (!keepValues) {
    elements.material.value = "";
    elements.specification.value = "";
  }
  setMaterialHint("输入关键词可搜索现有物料；无匹配时将自动生成编号。");
}

function clearForm({ focus = true, notify = true } = {}) {
  clearTimeout(materialSearchTimer);
  clearTimeout(balanceTimer);
  invalidateMaterialSearch();
  invalidateBalanceRequest();
  elements.form.reset();
  clearSelectedMaterial(false);
  materialCandidates = [];
  closeMaterialSuggestions();
  elements.requester.value = "";
  elements.project.value = "";
  elements.location.value = "";
  elements.shelfLetter.value = "";
  fillShelfNumbers("");
  elements.po.value = "";
  elements.changedAt.value = "";
  elements.quantity.value = "";
  elements.changeType.value = "";
  elements.note.value = "";
  elements.balanceBefore.textContent = "—";
  elements.balanceAfter.textContent = "—";
  if (notify) showToast("已清空当前输入内容");
  if (focus) elements.material.focus();
}

function selectMaterial(candidate) {
  selectedMaterial = candidate;
  elements.material.value = candidate.material;
  elements.specification.value = candidate.specification;
  elements.location.value = candidate.location || "";
  syncShelfFromLocation();
  elements.requester.value = candidate.requester || "";
  elements.project.value = candidate.project || "";
  elements.po.value = candidate.po || "";
  const codeLabel = candidate.materialCode || "待升级物料";
  setMaterialHint(`已选择 ${codeLabel}；提交时将按物料身份更新库存。`, true);
  closeMaterialSuggestions();
  elements.balanceBefore.textContent = formatNumber(candidate.balance);
  updateBalanceAfter(candidate.balance);
  refreshBalance();
}

function materialSearchQuery() {
  return [elements.material.value.trim(), elements.specification.value.trim()]
    .filter(Boolean)
    .join(" ");
}

function invalidateMaterialSearch() {
  materialSearchSequence += 1;
  if (materialRequestController) materialRequestController.abort();
  return materialSearchSequence;
}

async function runMaterialSearch(sequence = invalidateMaterialSearch()) {
  clearTimeout(materialSearchTimer);
  const query = materialSearchQuery();
  if (!query) {
    materialCandidates = [];
    closeMaterialSuggestions();
    return;
  }

  materialRequestController = new AbortController();
  try {
    const params = new URLSearchParams({ q: query, limit: "12" });
    const payload = await fetchJson(`/api/materials?${params}`, {
      signal: materialRequestController.signal,
    });
    if (sequence !== materialSearchSequence) return;
    const rawMaterials = Array.isArray(payload)
      ? payload
      : payload.materials ?? payload.records ?? [];
    const candidates = rawMaterials
      .map(normalizeMaterial)
      .filter((candidate) => candidate.material);
    renderMaterialSuggestions(candidates, query);
  } catch (error) {
    if (error.name === "AbortError") return;
    closeMaterialSuggestions();
    showToast(error.message);
  }
}

function scheduleMaterialSearch() {
  clearTimeout(materialSearchTimer);
  const sequence = invalidateMaterialSearch();
  materialSearchTimer = setTimeout(() => runMaterialSearch(sequence), 180);
}

function updateBalanceAfter(beforeValue = elements.balanceBefore.textContent) {
  const before =
    typeof beforeValue === "number"
      ? beforeValue
      : Number(String(beforeValue).replace(/,/g, ""));
  const quantity = Number(elements.quantity.value.replace(/,/g, ""));
  elements.balanceAfter.textContent = Number.isFinite(before)
    ? formatNumber(before + (Number.isFinite(quantity) ? quantity : 0))
    : "—";
}

async function loadState() {
  const params = new URLSearchParams({
    materialCode: selectedMaterial?.materialCode || "",
    material: elements.material.value.trim(),
    specification: elements.specification.value.trim(),
    quantity: elements.quantity.value.trim(),
  });
  const state = await fetchJson(`/api/state?${params}`);
  fillDatalist(
    elements.requester,
    elements.requesterSuggestions,
    state.requesters,
    DEFAULT_REQUESTERS,
  );
  fillSelect(elements.project, state.projects);
  fillDatalist(
    elements.location,
    elements.locationSuggestions,
    state.locations,
    DEFAULT_LOCATIONS,
  );
  fillSelect(elements.shelfLetter, SHELF_LETTERS);
  syncShelfFromLocation();
  fillSelect(elements.changeType, state.changeTypes);
  if (elements.material.value.trim()) {
    elements.balanceBefore.textContent = formatNumber(state.balanceBefore);
    elements.balanceAfter.textContent = formatNumber(state.balanceAfter);
  } else {
    elements.balanceBefore.textContent = "—";
    elements.balanceAfter.textContent = "—";
  }
}

function invalidateBalanceRequest() {
  balanceRequestSequence += 1;
  if (balanceRequestController) balanceRequestController.abort();
  return balanceRequestSequence;
}

async function refreshBalance() {
  clearTimeout(balanceTimer);
  const sequence = invalidateBalanceRequest();
  const material = elements.material.value.trim();
  if (!material) {
    elements.balanceBefore.textContent = "—";
    elements.balanceAfter.textContent = "—";
    return;
  }
  balanceTimer = setTimeout(async () => {
    balanceRequestController = new AbortController();
    try {
      const params = new URLSearchParams({
        materialCode: selectedMaterial?.materialCode || "",
        material,
        specification: elements.specification.value.trim(),
        quantity: elements.quantity.value.trim(),
      });
      const state = await fetchJson(`/api/state?${params}`, {
        signal: balanceRequestController.signal,
      });
      if (sequence !== balanceRequestSequence) return;
      elements.balanceBefore.textContent = formatNumber(state.balanceBefore);
      elements.balanceAfter.textContent = formatNumber(state.balanceAfter);
    } catch (error) {
      if (error.name === "AbortError") return;
      showToast(error.message);
    }
  }, 220);
}

function tableCell(text, className = "") {
  const cell = document.createElement("td");
  cell.textContent = text ?? "";
  if (className) cell.className = className;
  return cell;
}

async function refreshRecent() {
  elements.recentBody.innerHTML =
    '<tr><td colspan="10" class="empty-cell">正在读取记录…</td></tr>';
  try {
    const payload = await fetchJson("/api/recent");
    elements.recentBody.replaceChildren();
    if (!payload.records.length) {
      elements.recentBody.innerHTML =
        '<tr><td colspan="10" class="empty-cell">还没有库存变动记录</td></tr>';
      return;
    }
    payload.records.forEach((record) => {
      const row = document.createElement("tr");
      const quantityClass = Number(record.quantity) < 0 ? "negative" : "positive";
      row.append(
        tableCell(record.material_code ?? record.materialCode ?? "—"),
        tableCell(record.material ?? record.material_name ?? ""),
        tableCell(record.specification ?? record.spec ?? "—"),
        tableCell(record.location || "—"),
        tableCell(record.changed_at ?? record.changedAt ?? ""),
        tableCell(formatNumber(record.quantity), `number-cell ${quantityClass}`),
        tableCell(
          formatNumber(record.balance ?? record.balance_after),
          "number-cell",
        ),
        tableCell(record.change_type ?? record.changeType ?? ""),
        tableCell(record.requester),
        tableCell(record.project),
      );
      elements.recentBody.append(row);
    });
  } catch (error) {
    elements.recentBody.innerHTML =
      '<tr><td colspan="10" class="empty-cell">无法读取记录</td></tr>';
    showToast(error.message);
  }
}

async function submitRecord(event) {
  event.preventDefault();
  if (!elements.form.checkValidity()) {
    elements.form.reportValidity();
    return;
  }
  const payload = {
    materialCode: selectedMaterial?.materialCode || "",
    material: elements.material.value.trim(),
    specification: elements.specification.value.trim(),
    location: elements.location.value.trim(),
    requester: elements.requester.value.trim(),
    project: elements.project.value.trim(),
    po: elements.po.value.trim(),
    changedAt: elements.changedAt.value.trim(),
    quantity: elements.quantity.value.trim(),
    changeType: elements.changeType.value,
    note: elements.note.value.trim(),
  };
  closeMaterialSuggestions();
  setBusy(elements.submitButton, true, "正在写入并上传…");
  try {
    const result = await fetchJson("/api/submit", {
      method: "POST",
      body: JSON.stringify(payload),
    });
    const resultCode = String(
      result.materialCode ?? result.material_code ?? payload.materialCode,
    ).trim();
    if (resultCode) {
      selectedMaterial = {
        ...(selectedMaterial || {}),
        materialCode: resultCode,
        material: payload.material,
        specification: payload.specification,
        location: result.location ?? payload.location,
        balance: result.balanceAfter,
        requester: payload.requester,
        project: payload.project,
        po: payload.po,
      };
      setMaterialHint(`已选择 ${resultCode}；提交时将按此编码更新库存。`, true);
    }
    if (result.uploaded) {
      showDialog(
        "写入成功",
        `${result.message}\n\n已同步上传至云端。`,
      );
    } else {
      showDialog(
        "写入成功",
        `${result.message}\n\n已保存到本地，但云端上传失败：${
          result.uploadError || "网络异常"
        }\n系统将在后续自动重试同步。`,
      );
    }
    elements.po.value = "";
    elements.quantity.value = "";
    elements.note.value = "";
    elements.changedAt.value = localDate();
    await Promise.all([refreshRecent(), refreshBalance()]);
  } catch (error) {
    if (error.conflict) {
      showDialog(
        "无法写入",
        `${error.message}\n\n请先点击「同步最新版本」下载云端数据后再操作。`,
        true,
      );
    } else {
      showDialog("无法写入", error.message, true);
    }
  } finally {
    setBusy(elements.submitButton, false, "");
  }
}

async function openWorkbook() {
  try {
    await fetchJson("/api/open-excel", {
      method: "POST",
      body: "{}",
    });
    showToast("已打开 Excel；再次写入前请先关闭文件");
  } catch (error) {
    showDialog("无法打开 Excel", error.message, true);
  }
}

async function shutdownApp() {
  const confirmed = window.confirm("确定退出库存录入程序吗？");
  if (!confirmed) return;
  try {
    await fetchJson("/api/shutdown", {
      method: "POST",
      body: "{}",
    });
  } catch {
    // The server may close before the browser receives the last packet.
  }
  document.body.innerHTML = `
    <main class="closed-page">
      <section class="dialog">
        <span class="dialog-icon">✓</span>
        <h3>程序已退出</h3>
        <p>窗口即将自动关闭，托盘图标将消失，下次打开需要重新登录。</p>
      </section>
    </main>
  `;
}

function handleMaterialInput() {
  if (selectedMaterial) clearSelectedMaterial(true);
  materialCandidates = [];
  closeMaterialSuggestions();
  scheduleMaterialSearch();
  refreshBalance();
}

let syncState = null;

function applyRoleUI(role) {
  // 普通用户只保留「下载 Excel」；录入/同步/历史版本/最近记录仅管理员可见
  const isAdmin = role === "admin";
  document
    .querySelectorAll(".form-card, .sync-card, .recent-card")
    .forEach((el) => {
      el.hidden = !isAdmin;
    });
  if (elements.adminCard) elements.adminCard.hidden = !isAdmin;
  if (elements.replaceButton) elements.replaceButton.hidden = !isAdmin;
}

let lastOnlineState = null; // 记录上一次云端连接状态，用于断开提示
let cloudCheckedAt = 0; // 定时检测时间戳

function showCloudInfo() {
  // 云端连接信息弹窗（HTML 弹窗，含状态、地址、复制按钮）
  const state = syncState || {};
  const wrap = document.createElement("div");
  wrap.style.cssText =
    "position:fixed;inset:0;background:rgba(0,0,0,.45);display:flex;" +
    "align-items:center;justify-content:center;z-index:9999;";
  const box = document.createElement("div");
  box.style.cssText =
    "background:#f4f8fc;border:1px solid #8fb7dd;border-radius:6px;" +
    "padding:20px 24px;width:480px;box-shadow:0 4px 16px rgba(0,0,0,.3);";
  const online = !!state.online && !!state.loggedIn;
  const localUrl = state.serverUrl || "未连接";
  const publicUrl = state.publicUrl || "未配置";
  const dotColor = online ? "#3bb94a" : "#e34f4f";
  const statusText = online ? "当前已连接" : "当前已断开";
  box.innerHTML =
    '<div style="font-size:15px;font-weight:bold;color:#1d66b8;margin-bottom:12px;">云端连接信息</div>' +
    '<div style="font-size:13px;color:#333;margin-bottom:10px;">' +
    '<span style="display:inline-block;width:10px;height:10px;border-radius:50%;' +
    'background:' + dotColor + ';margin-right:6px;"></span><b>' + statusText + '</b></div>' +
    '<div style="font-size:13px;color:#333;margin-bottom:6px;">局域网地址（同网络同事用）：</div>' +
    '<div style="display:flex;gap:8px;margin-bottom:12px;">' +
    '<input id="ciLocal" type="text" readonly value="' + localUrl + '" style="flex:1;padding:6px 8px;' +
    'border:1px solid #b8cfe0;border-radius:3px;font-size:13px;background:#fff;" />' +
    '<button id="ciCopyLocal" style="padding:6px 12px;border:1px solid #1d66b8;background:#2d7dd8;color:#fff;' +
    'border-radius:3px;cursor:pointer;font-size:13px;">复制</button></div>' +
    '<div style="font-size:13px;color:#333;margin-bottom:6px;">公网地址（远程同事用）：</div>' +
    '<div style="display:flex;gap:8px;margin-bottom:16px;">' +
    '<input id="ciPublic" type="text" readonly value="' + publicUrl + '" style="flex:1;padding:6px 8px;' +
    'border:1px solid #b8cfe0;border-radius:3px;font-size:13px;background:#fff;" />' +
    '<button id="ciCopyPublic" style="padding:6px 12px;border:1px solid #1d66b8;background:#2d7dd8;color:#fff;' +
    'border-radius:3px;cursor:pointer;font-size:13px;">复制</button></div>' +
    '<div style="text-align:right;">' +
    '<button id="ciClose" style="padding:6px 16px;border:1px solid #b8cfe0;background:#eef4fa;' +
    'border-radius:3px;cursor:pointer;font-size:13px;">关闭</button></div>';
  wrap.append(box);
  document.body.append(wrap);
  const close = () => wrap.remove();
  box.querySelector("#ciClose").addEventListener("click", close);
  box.querySelector("#ciCopyLocal").addEventListener("click", () => {
    navigator.clipboard.writeText(localUrl).then(() => showToast("局域网地址已复制")).catch(() => showToast("复制失败"));
  });
  box.querySelector("#ciCopyPublic").addEventListener("click", () => {
    navigator.clipboard.writeText(publicUrl).then(() => showToast("公网地址已复制")).catch(() => showToast("复制失败"));
  });
  wrap.addEventListener("click", (event) => { if (event.target === wrap) close(); });
}

async function checkCloudConnection() {
  // 定时检测云端连接（20 秒一次）；从已连接变为断开时弹窗提示
  if (!syncState || !syncState.loggedIn) return;
  const wasOnline = lastOnlineState !== null ? lastOnlineState : !!syncState.online;
  try {
    const fresh = await fetchJson("/api/sync-status");
    const nowOnline = !!fresh.online;
    // 同步最新连接状态（含公网地址，供云端信息弹窗使用）
    syncState = { ...syncState, online: fresh.online, serverUrl: fresh.serverUrl, publicUrl: fresh.publicUrl || syncState.publicUrl };
    lastOnlineState = nowOnline;
    updateSyncUI();
    if (wasOnline && !nowOnline) {
      showDialog(
        "云端链接已断开",
        "云端链接已断开，您当前进行的更改不会被上传。\n请检查网络连接，恢复后数据会自动同步。",
        true,
      );
    } else if (!wasOnline && nowOnline) {
      // 恢复连接：自动拉取云端最新数据并提示
      showToast("云端已重新连接，正在同步最新数据…");
      try {
        const pull = await fetchJson("/api/sync-pull", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: "{}",
        });
        if (pull && typeof pull === "object") {
          syncState = { ...syncState, ...pull };
        }
        updateSyncUI();
        showToast("已同步云端最新数据");
      } catch (error) {}
    }
  } catch (error) {
    if (wasOnline) {
      showDialog(
        "云端链接已断开",
        "云端链接已断开，您当前进行的更改不会被上传。\n请检查网络连接，恢复后数据会自动同步。",
        true,
      );
    }
    lastOnlineState = false;
  }
}

function updateSyncUI() {
  const state = syncState || {};
  if (state.loggedIn) {
    elements.roleBadge.hidden = false;
    elements.roleBadge.textContent = state.role === "admin" ? "管理员" : "普通用户";
    elements.replaceButton.hidden = state.role !== "admin";
    applyRoleUI(state.role);
    if (state.online) {
      elements.syncBadge.innerHTML = '<span class="dot"></span>云端已连接';
      elements.syncBadge.classList.remove("offline");
    } else {
      elements.syncBadge.innerHTML = '<span class="dot"></span>云端已断开';
      elements.syncBadge.classList.add("offline");
    }
    const parts = [`云端：${state.serverUrl}`];
    if (state.online && state.latestVersion) {
      parts.push(`云端最新：${state.latestVersion.uploadedAt}（${state.latestVersion.versionId}）`);
    }
    if (state.baselineVersion) parts.push(`本地基线：${state.baselineVersion}`);
    if (state.pendingPush) parts.push("本地有未上传的更改");
    if (state.lastSyncAt) parts.push(`上次同步：${state.lastSyncAt}`);
    elements.syncSummary.textContent = parts.join("；");
  } else {
    elements.roleBadge.hidden = true;
    elements.replaceButton.hidden = true;
    elements.syncBadge.innerHTML = '<span class="dot"></span>未登录';
    elements.syncBadge.classList.remove("offline");
    elements.syncSummary.textContent = "登录后可同步到云端。";
  }
}

async function handleLogin(event) {
  event.preventDefault();
  const password = elements.loginPassword.value.trim();
  if (!password) return;
  elements.loginHint.textContent = "";
  setBusy(elements.loginButton, true, "正在登录…");
  try {
    const result = await fetchJson("/api/login", {
      method: "POST",
      body: JSON.stringify({ password }),
    });
    syncState = result;
    elements.loginOverlay.hidden = true;
    elements.loginPassword.value = "";
    updateSyncUI();
    await Promise.all([
      loadState(),
      refreshRecent(),
      refreshVersions(),
      refreshUsers(),
      refreshProjects(),
    ]);
    checkRemoteUpdate();
  } catch (error) {
    elements.loginHint.textContent = error.message;
  } finally {
    setBusy(elements.loginButton, false, "");
  }
}

function checkRemoteUpdate() {
  const state = syncState;
  if (!state || !state.loggedIn || !state.online || !state.latestVersion) return;
  if (state.baselineVersion === state.latestVersion.versionId) return;
  // Excel 以主机端为准：检测到云端更新后自动强制同步本地副本，不再询问
  syncPull(true);
}

async function syncPull(auto) {
  if (!syncState?.loggedIn) {
    showDialog("未登录", "请先登录后再同步。", true);
    return;
  }
  setBusy(elements.syncPullButton, true, "正在同步…");
  try {
    const result = await fetchJson("/api/sync-pull", {
      method: "POST",
      body: "{}",
    });
    syncState = result;
    updateSyncUI();
    await Promise.all([loadState(), refreshRecent(), refreshVersions()]);
    if (auto) {
      showToast("已自动同步云端最新数据");
    } else {
      showDialog("同步完成", "已下载云端最新版本并覆盖本地数据，覆盖前已自动备份。");
    }
  } catch (error) {
    showDialog("同步失败", error.message, true);
  } finally {
    setBusy(elements.syncPullButton, false, "");
  }
}

async function syncPush() {
  if (!syncState?.loggedIn) {
    showDialog("未登录", "请先登录后再上传。", true);
    return;
  }
  setBusy(elements.syncPushButton, true, "正在上传…");
  try {
    const result = await fetchJson("/api/sync-push", {
      method: "POST",
      body: "{}",
    });
    syncState = result;
    updateSyncUI();
    showDialog("上传完成", "本地版本已上传到云端。");
  } catch (error) {
    showDialog("上传失败", error.message, true);
  } finally {
    setBusy(elements.syncPushButton, false, "");
  }
}

async function refreshVersions() {
  try {
    const payload = await fetchJson("/api/versions");
    const versions = Array.isArray(payload.versions) ? payload.versions : [];
    elements.versionsBody.replaceChildren();
    if (!versions.length) {
      elements.versionsBody.innerHTML =
        '<tr><td colspan="5" class="empty-cell">云端还没有任何版本</td></tr>';
      return;
    }
    const isAdmin = syncState?.role === "admin";
    versions.forEach((version) => {
      const row = document.createElement("tr");
      const size = Number(version.size) || 0;
      const sizeText = size >= 1024 ? `${(size / 1024).toFixed(1)} KB` : `${size} B`;
      const actionCell = document.createElement("td");
      actionCell.className = "number-cell";
      if (isAdmin) {
        const replaceBtn = document.createElement("button");
        replaceBtn.type = "button";
        replaceBtn.className = "rollback-link";
        replaceBtn.textContent = "用此版本替换";
        replaceBtn.addEventListener("click", () => rollbackVersion(version.versionId));
        actionCell.append(replaceBtn);
      }
      const downloadBtn = document.createElement("button");
      downloadBtn.type = "button";
      downloadBtn.className = "rollback-link";
      downloadBtn.textContent = "下载";
      downloadBtn.title = "下载该版本到桌面";
      downloadBtn.addEventListener("click", () => downloadVersion(version.versionId));
      actionCell.append(downloadBtn);
      row.append(
        tableCell(version.uploadedAt),
        tableCell(version.versionId),
        tableCell(sizeText),
        tableCell(version.uploader === "admin" ? "管理员" : "普通用户"),
        actionCell,
      );
      elements.versionsBody.append(row);
    });
  } catch (error) {
    elements.versionsBody.innerHTML =
      '<tr><td colspan="5" class="empty-cell">无法读取云端版本</td></tr>';
    showToast(error.message);
  }
}

async function rollbackVersion(versionId) {
  const confirmed = window.confirm(
    `确定用该版本覆盖当前库存数据吗？\n覆盖前会自动备份本地文件。`,
  );
  if (!confirmed) return;
  try {
    const result = await fetchJson("/api/rollback", {
      method: "POST",
      body: JSON.stringify({ versionId }),
    });
    syncState = result;
    updateSyncUI();
    await Promise.all([loadState(), refreshRecent(), refreshVersions()]);
    showDialog("回溯完成", "已恢复到所选版本，本地已自动备份。");
  } catch (error) {
    showDialog("回溯失败", error.message, true);
  }
}

async function downloadVersion(versionId) {
  try {
    let defaultPath = "";
    try {
      defaultPath = (await fetchJson("/api/desktop-path")).path || "";
    } catch (error) {}
    const saveDir = await promptSavePath(defaultPath);
    if (saveDir === null) {
      showToast("已取消下载");
      return;
    }
    const payload = await fetchJson("/api/download-version", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ versionId, saveDir }),
    });
    showDialog("下载完成", payload.message || "已保存。");
  } catch (error) {
    showDialog("下载失败", error.message, true);
  }
}

async function refreshUsers() {
  if (syncState?.role !== "admin") {
    elements.adminCard.hidden = true;
    return;
  }
  elements.adminCard.hidden = false;
  try {
    const payload = await fetchJson("/api/users");
    const users = Array.isArray(payload.users) ? payload.users : [];
    elements.usersBody.replaceChildren();
    if (!users.length) {
      elements.usersBody.innerHTML =
        '<tr><td colspan="4" class="empty-cell">还没有普通用户</td></tr>';
      return;
    }
    users.forEach((user) => {
      const row = document.createElement("tr");
      const actionCell = document.createElement("td");
      actionCell.className = "number-cell";
      const button = document.createElement("button");
      button.type = "button";
      button.className = "delete-user-link";
      button.textContent = "删除";
      button.addEventListener("click", () => removeUser(user.id));
      actionCell.append(button);
      row.append(
        tableCell(user.password || user.masked || "已加密"),
        tableCell(user.note || "—"),
        tableCell(user.createdAt),
        actionCell,
      );
      elements.usersBody.append(row);
    });
  } catch (error) {
    elements.usersBody.innerHTML =
      '<tr><td colspan="4" class="empty-cell">无法读取用户</td></tr>';
    showToast(error.message);
  }
}

async function addUser(event) {
  event.preventDefault();
  const password = elements.newUserPassword.value.trim();
  const note = elements.newUserNote.value.trim();
  if (!password) {
    showToast("请输入数字密码");
    return;
  }
  setBusy(elements.addUserButton, true, "添加中…");
  try {
    const result = await fetchJson("/api/users", {
      method: "POST",
      body: JSON.stringify({ action: "add", password, note }),
    });
    elements.newUserPassword.value = "";
    elements.newUserNote.value = "";
    showDialog(
      "添加成功",
      "已添加普通用户，密码为：" +
        (result.password || password) +
        "（仅此一次显示，请立即告知对方）。",
    );
    await refreshUsers();
  } catch (error) {
    showDialog("添加失败", error.message, true);
  } finally {
    setBusy(elements.addUserButton, false, "");
  }
}

async function removeUser(password) {
  const confirmed = window.confirm(
    `确定删除普通用户 ${password} 吗？\n删除后该用户将无法再登录。`,
  );
  if (!confirmed) return;
  try {
    await fetchJson("/api/users", {
      method: "POST",
      body: JSON.stringify({ action: "remove", password }),
    });
    showToast(`已删除用户 ${password}`);
    await refreshUsers();
  } catch (error) {
    showDialog("删除失败", error.message, true);
  }
}

async function refreshProjects() {
  try {
    const payload = await fetchJson("/api/projects");
    const projects = Array.isArray(payload.projects) ? payload.projects : [];
    fillSelect(elements.project, projects);
    if (syncState?.role === "admin") {
      elements.projectsBody.replaceChildren();
      if (!projects.length) {
        elements.projectsBody.innerHTML =
          '<tr><td colspan="2" class="empty-cell">暂无项目</td></tr>';
        return;
      }
      projects.forEach((project) => {
        const row = document.createElement("tr");
        const actionCell = document.createElement("td");
        actionCell.className = "number-cell";
        const button = document.createElement("button");
        button.type = "button";
        button.className = "delete-user-link";
        button.textContent = "删除";
        button.addEventListener("click", () => removeProject(project));
        actionCell.append(button);
        row.append(tableCell(project), actionCell);
        elements.projectsBody.append(row);
      });
    }
  } catch (error) {
    showToast(error.message);
  }
}

async function addProject(event) {
  event.preventDefault();
  const project = elements.newProjectName.value.trim();
  if (!project) {
    showToast("请输入项目名称");
    return;
  }
  setBusy(elements.addProjectButton, true, "创建中…");
  try {
    const payload = await fetchJson("/api/projects", {
      method: "POST",
      body: JSON.stringify({ action: "add", project }),
    });
    elements.newProjectName.value = "";
    fillSelect(elements.project, payload.projects);
    showDialog("新增成功", `已创建项目 ${project}，并同步至云端。`);
    await Promise.all([refreshProjects(), refreshSyncAll()]);
  } catch (error) {
    showDialog("新增失败", error.message, true);
  } finally {
    setBusy(elements.addProjectButton, false, "");
  }
}

async function removeProject(project) {
  const confirmed = window.confirm(
    `确定删除项目 ${project} 吗？\n该项目下的全部物料数据将被删除（操作前会自动备份整个文件）。`,
  );
  if (!confirmed) return;
  try {
    const payload = await fetchJson("/api/projects", {
      method: "POST",
      body: JSON.stringify({ action: "remove", project }),
    });
    fillSelect(elements.project, payload.projects);
    showDialog("删除完成", `项目 ${project} 已删除并同步至云端。`);
    await Promise.all([refreshProjects(), refreshSyncAll()]);
  } catch (error) {
    showDialog("删除失败", error.message, true);
  }
}

async function changePassword(event) {
  event.preventDefault();
  const oldPassword = elements.oldAdminPassword.value.trim();
  const newPassword = elements.newAdminPassword.value.trim();
  if (!oldPassword || !newPassword) {
    showToast("请填写原密码和新密码");
    return;
  }
  setBusy(elements.changePasswordButton, true, "修改中…");
  try {
    await fetchJson("/api/admin/password", {
      method: "POST",
      body: JSON.stringify({ oldPassword, newPassword }),
    });
    showDialog("密码已修改", "修改成功后云端已使所有登录失效，请使用新密码重新登录。");
    syncState = null;
    elements.loginOverlay.hidden = false;
    elements.loginPassword.value = "";
    elements.loginPassword.focus();
    updateSyncUI();
  } catch (error) {
    showDialog("修改失败", error.message, true);
  } finally {
    setBusy(elements.changePasswordButton, false, "");
  }
}

async function openServerSetting() {
  elements.discoverResults.replaceChildren();
  elements.serverUrlInput.value = "";
  try {
    const payload = await fetchJson("/api/sync-status");
    elements.currentServerUrl.textContent =
      payload.serverUrl || "未配置";
  } catch (error) {
    elements.currentServerUrl.textContent = "—";
  }
  try {
    const payload = await fetchJson("/api/server-history");
    const history = Array.isArray(payload.history) ? payload.history : [];
    const select = elements.serverHistorySelect;
    select.replaceChildren();
    const placeholder = document.createElement("option");
    placeholder.value = "";
    placeholder.textContent = "— 选择历史地址 —";
    select.append(placeholder);
    history.forEach((url) => {
      const option = document.createElement("option");
      option.value = url;
      option.textContent = url;
      select.append(option);
    });
  } catch (error) {
    /* 历史加载失败不影响使用 */
  }
  elements.serverSettingOverlay.hidden = false;
}

function closeServerSetting() {
  elements.serverSettingOverlay.hidden = true;
}

async function discoverServers() {
  elements.discoverResults.innerHTML =
    '<span class="field-hint">正在搜索局域网服务器…</span>';
  elements.discoverButton.disabled = true;
  try {
    const payload = await fetchJson("/api/discover-servers");
    const servers = Array.isArray(payload.servers) ? payload.servers : [];
    elements.discoverResults.replaceChildren();
    if (!servers.length) {
      elements.discoverResults.innerHTML =
        '<span class="field-hint">未发现服务器。请确认服务器电脑已启动且与本机在同一网络，或手动输入地址。</span>';
      return;
    }
    servers.forEach((server) => {
      const row = document.createElement("div");
      row.className = "discover-row";
      const label = document.createElement("span");
      label.textContent = `${server.hostname || "未知主机"}（${server.ip}）`;
      const button = document.createElement("button");
      button.type = "button";
      button.className = "secondary-button";
      button.textContent = "使用此服务器";
      button.addEventListener("click", () => {
        elements.serverUrlInput.value = server.url;
      });
      row.append(label, button);
      elements.discoverResults.append(row);
    });
  } catch (error) {
    elements.discoverResults.innerHTML =
      `<span class="field-hint">搜索失败：${error.message}</span>`;
  } finally {
    elements.discoverButton.disabled = false;
  }
}

async function saveServerUrl() {
  const url = elements.serverUrlInput.value.trim();
  if (!url) {
    showToast("请输入服务器地址");
    return;
  }
  setBusy(elements.saveServerUrlButton, true, "保存中…");
  try {
    await fetchJson("/api/set-server-url", {
      method: "POST",
      body: JSON.stringify({ url }),
    });
    closeServerSetting();
    syncState = null;
    elements.loginOverlay.hidden = false;
    elements.loginPassword.value = "";
    elements.loginPassword.focus();
    updateSyncUI();
    showDialog("设置已保存", `服务器地址已更新为 ${url}，请重新登录。`);
  } catch (error) {
    showDialog("保存失败", error.message, true);
  } finally {
    setBusy(elements.saveServerUrlButton, false, "");
  }
}

async function downloadExcel() {
  try {
    let defaultPath = "";
    try {
      defaultPath = (await fetchJson("/api/desktop-path")).path || "";
    } catch (error) {}
    const saveDir = await promptSavePath(defaultPath);
    if (saveDir === null) {
      showToast("已取消下载");
      return;
    }
    const payload = await fetchJson("/api/download-excel", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ saveDir }),
    });
    showDialog("下载完成", payload.message || "已保存。");
  } catch (error) {
    showDialog("下载失败", error.message, true);
  }
}

async function replaceWorkbookFile(event) {
  const file = event.target.files?.[0];
  event.target.value = "";
  if (!file) return;
  if (
    !window.confirm(
      "确定用「" +
        file.name +
        "」替换当前库存表吗？\n主机将检查格式并反馈结果，替换前会自动备份旧表。",
    )
  ) {
    return;
  }
  const reader = new FileReader();
  const dataUrl = await new Promise((resolve, reject) => {
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(new Error("读取文件失败。"));
    reader.readAsDataURL(file);
  });
  const base64 = String(dataUrl).split(",")[1] || "";
  try {
    const payload = await fetchJson("/api/upload-replace", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ fileName: file.name, data: base64 }),
    });
    showDialog("替换成功", payload.message || "库存表已替换。");
    if (Array.isArray(payload.projects)) {
      fillSelect(elements.project, payload.projects);
    }
  } catch (error) {
    showDialog("替换失败", error.message, true);
  }
}

async function refreshSyncAll() {
  try {
    syncState = await fetchJson("/api/sync-status");
    updateSyncUI();
    await Promise.all([
      loadState(),
      refreshRecent(),
      refreshVersions(),
      refreshUsers(),
      refreshProjects(),
    ]);
  } catch (error) {
    showToast(error.message);
  }
}

elements.form.addEventListener("submit", submitRecord);

// 底部公司 logo：点击显示制作者信息
const appLogo = document.querySelector("#appLogo");
if (appLogo) {
  appLogo.addEventListener("click", () => {
    showDialog(
      "制作者信息",
      "董理臻<br/>2026年8月17日完成<br/><br/>" +
        "联系方式：<br/>微信：18551780019<br/>QQ：1301535058<br/><br/>" +
        "<span class='thanks-line'>感谢 <b class='thanks-ds'>DeepSeek-V4 Flash</b> 和 " +
        "<b class='thanks-tf'>永雏塔菲</b> 的协助</span>",
      false,
      true,
    );
  });
}
elements.loginForm.addEventListener("submit", handleLogin);
elements.addUserForm.addEventListener("submit", addUser);
elements.syncPullButton.addEventListener("click", syncPull);
elements.syncPushButton.addEventListener("click", syncPush);
elements.syncRefreshButton.addEventListener("click", refreshSyncAll);
elements.usersRefreshButton.addEventListener("click", refreshUsers);
elements.addProjectForm.addEventListener("submit", addProject);
elements.changePasswordForm.addEventListener("submit", changePassword);
elements.serverSettingLink.addEventListener("click", openServerSetting);
elements.closeServerSettingButton.addEventListener("click", closeServerSetting);
elements.discoverButton.addEventListener("click", discoverServers);
elements.saveServerUrlButton.addEventListener("click", saveServerUrl);
elements.serverHistorySelect.addEventListener("change", () => {
  elements.serverUrlInput.value = elements.serverHistorySelect.value;
});
elements.downloadButton.addEventListener("click", downloadExcel);
elements.replaceButton.addEventListener("click", () => elements.replaceInput.click());
elements.replaceInput.addEventListener("change", replaceWorkbookFile);
elements.syncBadge.addEventListener("click", showCloudInfo);
document.querySelector("#cloudInfoBtn").addEventListener("click", showCloudInfo);
setInterval(checkCloudConnection, 10000);
elements.adminRefreshButton.addEventListener("click", refreshSyncAll);
elements.material.addEventListener("input", handleMaterialInput);
elements.material.addEventListener("focus", () => {
  if (elements.material.value.trim()) runMaterialSearch();
});
elements.material.addEventListener("blur", () => {
  setTimeout(closeMaterialSuggestions, 120);
});
elements.material.addEventListener("keydown", (event) => {
  if (event.key === "Escape" && !elements.materialSuggestions.hidden) {
    event.preventDefault();
    closeMaterialSuggestions();
    return;
  }
  if (event.key === "ArrowDown" || event.key === "ArrowUp") {
    if (!materialCandidates.length) {
      runMaterialSearch();
      return;
    }
    event.preventDefault();
    setActiveCandidate(
      activeCandidateIndex +
        (event.key === "ArrowDown" ? 1 : -1),
    );
    return;
  }
  if (
    event.key === "Enter" &&
    !elements.materialSuggestions.hidden &&
    materialCandidates.length
  ) {
    event.preventDefault();
    selectMaterial(
      materialCandidates[activeCandidateIndex < 0 ? 0 : activeCandidateIndex],
    );
  }
});
elements.specification.addEventListener("input", () => {
  if (selectedMaterial) clearSelectedMaterial(true);
  scheduleMaterialSearch();
  refreshBalance();
});
elements.shelfLetter.addEventListener("change", () => {
  fillShelfNumbers(elements.shelfLetter.value);
  syncLocationFromShelf();
});
elements.shelfNumber.addEventListener("change", syncLocationFromShelf);
elements.location.addEventListener("input", syncShelfFromLocation);
elements.quantity.addEventListener("input", () => {
  updateBalanceAfter();
  refreshBalance();
});
elements.clearButton.addEventListener("click", () => clearForm());
elements.nowButton.addEventListener("click", () => {
  elements.changedAt.value = localDate();
});
elements.refreshButton.addEventListener("click", refreshRecent);
elements.shutdownButton.addEventListener("click", shutdownApp);
elements.dialogClose.addEventListener("click", closeDialog);
elements.dialogBackdrop.addEventListener("click", (event) => {
  if (event.target === elements.dialogBackdrop) closeDialog();
});
document.addEventListener("mousedown", (event) => {
  if (!event.target.closest(".material-search")) closeMaterialSuggestions();
});
document.addEventListener("keydown", (event) => {
  if (event.key === "Escape" && !elements.dialogBackdrop.hidden) closeDialog();
});

async function startup() {
  try {
    syncState = await fetchJson("/api/sync-status");
    updateSyncUI();
    if (!syncState.loggedIn) {
      elements.loginOverlay.hidden = false;
      elements.loginPassword.focus();
      return;
    }
    await Promise.all([
      loadState(),
      refreshRecent(),
      refreshVersions(),
      refreshUsers(),
      refreshProjects(),
    ]);
    checkRemoteUpdate();
  } catch (error) {
    showDialog("程序初始化失败", error.message, true);
  }
}

startup();
