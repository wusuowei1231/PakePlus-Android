//全局变量
let fileInput = null;
let detailFileInput = null;
let importModal = null;
let detailImportModal = null;
let detailKeyColSelect = null;
let detailDoAppendCheck = null;
let detailPreviewBox = null;
let importConfirmModal = null;
let keyColSelect = null;
let dedupeCheck = null;
let previewBox = null;
let noteMask = null;
let noteTextarea = null;
let synMask = null;

let currentNoteRowIdx = null;
let noteTarget = "detail"; // 备注弹窗来源：detail=详情分组 / lib=产品库
let currentViewRowIdx = null;
let viewRowSource = "detail"; // 查看弹窗来源：detail / lib
let tempImportDetailRows = [];
let tempDetailExcelRows = null;
let detailSubTabs = [];
let productLib = [];
let quizList = [];
let wrongPool = [];
let pendingRows = null;
let currentQ = 0;
let timer = null;
let sortLibTbody = null;
let sortDetailTbody = null;
let detailSearchText = "";
let libSearchText = "";

// ===== 移动端自适应工具（仅布局判断，不改任何业务逻辑/数据）=====
function isMobile() {
	return window.innerWidth <= 768;
}

// ===== PC端左侧抽屉工具栏 =====
function openDrawer(btn) {
	const panel = btn.closest('.panel');
	const left = panel ? panel.querySelector('.prod-left') : document.querySelector('.prod-left');
	if (left) left.classList.add('drawer-open');
	const mask = document.querySelector('.drawer-mask');
	if (mask) mask.classList.add('show');
}

function closeDrawer() {
	document.querySelectorAll('.prod-left').forEach(l => {
		if (!l.classList.contains('drawer-open')) return;
		l.classList.remove('drawer-open');
		l.classList.add('closing');
		setTimeout(() => l.classList.remove('closing'), 280);
	});
	const mask = document.querySelector('.drawer-mask');
	if (mask) mask.classList.remove('show');
}
window.openDrawer = openDrawer;
window.closeDrawer = closeDrawer;

function refreshLayout() {
	// 断点切换（F12 切设备/旋转屏幕）重渲染期间先隐藏内容，避免上一态结构套新样式造成错位闪现
	const boxes = [document.getElementById("detailTableBox"), document.getElementById("libTableBox")].filter(Boolean);
	boxes.forEach(b => {
		b.style.visibility = "hidden";
	});
	if (typeof renderDetailTable === "function") renderDetailTable();
	if (typeof renderLibTable === "function") renderLibTable();
	const restore = () => {
		boxes.forEach(b => {
			b.style.visibility = "";
		});
	};
	requestAnimationFrame(restore);
	setTimeout(restore, 80); // 兜底，防止 rAF 未触发导致内容被隐藏
}
let _resizeTimer = null;
window.addEventListener("resize", function() {
	clearTimeout(_resizeTimer);
	_resizeTimer = setTimeout(refreshLayout, 150);
});
// DevTools切换设备/旋转屏幕时，媒体查询变化稳定触发，补充resize可能不触发的情况
if (window.matchMedia) {
	window.matchMedia("(max-width:768px)").addEventListener("change", refreshLayout);
}

//挂载到window，供onclick调用
window.switchTab = switchTab;
window.closeNoteMask = closeNoteMask;
window.openNoteMask = openNoteMask;
window.openDetailView = openDetailView;
window.closeDetailView = closeDetailView;
window.copyCellText = copyCellText;
window.copyCurrentNote = copyCurrentNote;
window.copyCurrentViewRow = copyCurrentViewRow;
window.saveNote = saveNote;
window.closeImportConfirmModal = closeImportConfirmModal;
window.closeDetailImportModal = closeDetailImportModal;
window.renderSubSelect = renderSubSelect;
window.onSubSelectChange = onSubSelectChange;
window.addNewSubTab = addNewSubTab;
window.renameCurrentSubTab = renameCurrentSubTab;
window.deleteCurrentSubTab = deleteCurrentSubTab;
window.downloadDetailGroupTemplate = downloadDetailGroupTemplate;
window.openSynModal = openSynModal;
window.closeSynModal = closeSynModal;
window.renderSynList = renderSynList;
window.addSynLine = addSynLine;
window.delSyn = delSyn;
window.saveSynGroups = saveSynGroups;
window.onDetailSearch = onDetailSearch;
window.clearDetailSearch = clearDetailSearch;
window.onLibSearch = onLibSearch;
window.clearLibSearch = clearLibSearch;
window.readDetailExcel = readDetailExcel;
window.renderDetailImportPreview = renderDetailImportPreview;
window.confirmDetailImport = confirmDetailImport;
window.closeModal = closeModal;
window.readCustomExcel = readCustomExcel;
window.renderPreview = renderPreview;
window.confirmImport = confirmImport;
window.downloadTemplate = downloadTemplate;
window.exportAll = exportAll;
window.clearAllLib = clearAllLib;
window.addDetailEmptyRow = addDetailEmptyRow;
window.renderDetailTable = renderDetailTable;
window.detailCellChange = detailCellChange;
window.delDetailRow = delDetailRow;
window.renderLibTable = renderLibTable;
window.openLibNoteMask = openLibNoteMask;
window.libCellChange = libCellChange;
window.delLibRow = delLibRow;
window.singleImportToLib = singleImportToLib;
window.batchImportDetailToLib = batchImportDetailToLib;
window.doImportFromDetail = doImportFromDetail;
window.startQuiz = startQuiz;
window.markQuestion = markQuestion;
window.startWrongQuiz = startWrongQuiz;
window.resetQuiz = resetQuiz;
window.saveCurrentAns = saveCurrentAns;
window.prev = prev;
window.next = next;
window.submitAll = submitAll;
window.toggleWrong = toggleWrong;
window.backQuiz = backQuiz;
window.exportResult = exportResult;
window.exportWrong = exportWrong;
window.backupAllData = backupAllData;
window.restoreAllData = restoreAllData;
window.pickRestoreFile = pickRestoreFile;

//DOM全部加载完毕再执行
document.addEventListener('DOMContentLoaded', function() {
	fileInput = document.getElementById('fileInput');
	detailFileInput = document.getElementById('detailFileInput');
	importModal = document.getElementById('importModal');
	detailImportModal = document.getElementById('detailImportModal');
	detailKeyColSelect = document.getElementById('detailKeyCol');
	detailDoAppendCheck = document.getElementById('detailDoAppend');
	detailPreviewBox = document.getElementById('detailPreviewBox');
	importConfirmModal = document.getElementById('importConfirmModal');
	keyColSelect = document.getElementById('keyCol');
	dedupeCheck = document.getElementById('doDedupe');
	previewBox = document.getElementById('previewBox');
	noteMask = document.getElementById('noteMask');
	noteTextarea = document.getElementById('noteTextarea');
	synMask = document.getElementById('synMask');

	if (noteMask) {
		noteMask.addEventListener('click', function() {
			closeNoteMask();
		});
	}
	const detailViewMask = document.getElementById('detailViewMask');
	if (detailViewMask) {
		detailViewMask.addEventListener('click', function() {
			closeDetailView();
		});
	}

	//本地存储读取
	const oldDetail = localStorage.getItem("productDetailStorage");
	const subStr = localStorage.getItem("detailSubTabsStorage");
	if (subStr) {
		detailSubTabs = JSON.parse(subStr);
		if (!Array.isArray(detailSubTabs) || !detailSubTabs.length) {
			detailSubTabs = [{
				id: "sub_1",
				name: "分组1",
				data: []
			}];
			saveSubTabsStorage();
		}
	} else {
		if (oldDetail) {
			const oldData = JSON.parse(oldDetail);
			detailSubTabs = [{
				id: "sub_1",
				name: "旧数据迁移",
				data: oldData
			}];
		} else {
			detailSubTabs = [{
				id: "sub_1",
				name: "分组1",
				data: []
			}];
		}
		saveSubTabsStorage();
		localStorage.removeItem("productDetailStorage");
	}
	detailSubTabs.forEach(tab => {
		if (!Array.isArray(tab.data)) tab.data = [];
		if (tab.locked === undefined) tab.locked = false;
		tab.data.forEach(row => {
			if (row["备注"] === undefined) row["备注"] = "";
			if (!Array.isArray(row.imgs)) row.imgs = [];
		});
	});
	document.getElementById("activeSubTabId").value = detailSubTabs[0].id;

	const lib = localStorage.getItem("productLibStorage");
	const w = localStorage.getItem("wrongPool");
	if (lib) productLib = JSON.parse(lib);
	if (w) wrongPool = JSON.parse(w);
	productLib.forEach(item => {
		if (item["备注"] === undefined) item["备注"] = "";
	});
	saveLibData();

	renderSubSelect();
	renderDetailTable();
	// 初始化渲染完成，解除手机端首帧隐藏，一次性显示（PC不受影响）
	document.body.classList.add('app-ready');
});

function switchTab(type) {
	document.querySelectorAll('.tab-item').forEach(t => t.classList.remove('active'));
	document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
	if (type === 'detail') {
		document.querySelector('.tab-detail').classList.add('active');
		document.getElementById('panel-detail').classList.add('active');
		renderSubSelect();
		renderDetailTable();
	} else if (type === 'prod') {
		document.querySelector('.tab-prod').classList.add('active');
		document.getElementById('panel-prod').classList.add('active');
		renderLibTable();
	} else if (type === 'quiz') {
		document.querySelector('.tab-quiz').classList.add('active');
		document.getElementById('panel-quiz').classList.add('active');
	}
	// 同步移动端底部Tab高亮
	document.querySelectorAll('.m-tab').forEach(t => t.classList.toggle('active', t.dataset.tab === type));
	closeDrawer();
}

function closeNoteMask() {
	if (!noteMask) return;
	noteMask.style.display = "none";
	currentNoteRowIdx = null;
}

function escapeHtml(s) {
	return String(s ?? "").replace(/[&<>"']/g, c => ({
		"&": "&amp;",
		"<": "&lt;",
		">": "&gt;",
		'"': "&quot;",
		"'": "&#39;"
	} [c]));
}

function copyCellText(text) {
	const val = String(text ?? "").trim();
	if (!val) return;
	copyRawText(val, "已复制：" + val);
}
// 原样复制（保留换行，不做trim）
function copyRawText(raw, tip) {
	const val = String(raw ?? "");
	if (!val.trim()) return;
	const done = () => {
		try {
			alert(tip || "已复制");
		} catch (e) {}
	};
	if (navigator.clipboard && navigator.clipboard.writeText) {
		navigator.clipboard.writeText(val).then(done).catch(() => fallbackCopy(val, done));
	} else {
		fallbackCopy(val, done);
	}
}
// 复制当前备注弹窗里的完整内容（保留原始换行）
function copyCurrentNote() {
	if (!noteTextarea) {
		return;
	}
	copyRawText(noteTextarea.value, "备注已复制（换行已保留）");
}
// 复制查看弹窗里这一整条数据（每条字段一行，保留换行）
function copyCurrentViewRow() {
	if (currentViewRowIdx === null) return;
	const row = viewRowSource === "lib" ? productLib[currentViewRowIdx] : getActiveSubTab().data[currentViewRowIdx];
	if (!row) return;
	const text = Object.keys(row).filter(k => k !== "imgs").map(k => k + "：" + String(row[k] ?? "")).join("\n");
	copyRawText(text, "本条数据已复制");
}

function fallbackCopy(val, done) {
	const ta = document.createElement("textarea");
	ta.value = val;
	ta.style.position = "fixed";
	ta.style.opacity = "0";
	document.body.appendChild(ta);
	ta.select();
	try {
		document.execCommand("copy");
		done();
	} catch (e) {}
	document.body.removeChild(ta);
}

function openDetailView(idx) {
	const tab = getActiveSubTab();
	const row = tab.data[idx];
	if (!row) return;
	currentViewRowIdx = idx;
	viewRowSource = "detail";
	const box = document.getElementById("detailViewBox");
	const rows = Object.keys(row).filter(k => k !== "imgs").map(k => {
		const v = String(row[k] ?? "").trim();
		return `<div class="dv-row"><span class="dv-key">${escapeHtml(k)}</span><span class="dv-val">${escapeHtml(v)||"(空)"}</span></div>`;
	}).join("");
	box.innerHTML = rows;
	renderViewImgs();
	document.getElementById("detailViewMask").style.display = "flex";
}
let currentPreviewIdx = null;
let imgReplaceIdx = null;
let imgMode = 'add';

function renderViewImgs() {
	const tab = getActiveSubTab();
	const row = tab.data[currentViewRowIdx];
	const box = document.getElementById("viewImgsBox");
	if (!row || !box) return;
	if (!Array.isArray(row.imgs)) row.imgs = [];
	const locked = tab.locked === true;
	let html = '<div class="imgs-title">图片附件（' + row.imgs.length + '）</div><div class="imgs-list">';
	row.imgs.forEach((src, i) => {
		html += `<div class="img-thumb"><img src="${src}" onclick="previewImg(${i})" alt="附件图"></div>`;
	});
	if (!row.imgs.length) html += '<div class="imgs-empty">暂无图片</div>';
	html += '</div>';
	if (locked) {
		html += `<button class="btn-primary btn-disabled" disabled>📷 添加图片</button>`;
	} else {
		html += `<button class="btn-primary" onclick="triggerAddImg()">📷 添加图片</button>`;
	}
	box.innerHTML = html;
}

function compressImage(file) {
	return new Promise((resolve, reject) => {
		const reader = new FileReader();
		reader.onload = ev => {
			const img = new Image();
			img.onload = () => {
				const MAX = 1200;
				let w = img.width,
					h = img.height;
				if (Math.max(w, h) > MAX) {
					const r = MAX / Math.max(w, h);
					w = Math.round(w * r);
					h = Math.round(h * r);
				}
				const cv = document.createElement("canvas");
				cv.width = w;
				cv.height = h;
				cv.getContext("2d").drawImage(img, 0, 0, w, h);
				resolve(cv.toDataURL("image/jpeg", 0.82));
			};
			img.onerror = reject;
			img.src = ev.target.result;
		};
		reader.onerror = reject;
		reader.readAsDataURL(file);
	});
}

function triggerAddImg() {
	if (isActiveLocked()) {
		alert("🔒 当前分组已锁定，禁止上传图片。");
		return;
	}
	imgMode = 'add';
	imgReplaceIdx = null;
	const inp = document.getElementById('viewImgInput');
	inp.value = "";
	inp.click();
}

function onViewImgChange(e) {
	if (isActiveLocked()) {
		alert("🔒 当前分组已锁定，禁止修改图片。");
		e.target.value = "";
		return;
	}
	const files = [...(e.target.files || [])];
	if (!files.length) return;
	const tab = getActiveSubTab();
	const row = tab.data[currentViewRowIdx];
	if (!row) return;
	if (!Array.isArray(row.imgs)) row.imgs = [];
	(async () => {
		const compressed = [];
		for (const f of files) {
			try {
				compressed.push(await compressImage(f));
			} catch (err) {
				alert("图片读取失败：" + f.name);
			}
		}
		if (!compressed.length) {
			e.target.value = "";
			return;
		}
		if (imgMode === 'replace' && imgReplaceIdx !== null && row.imgs[imgReplaceIdx] !== undefined) {
			row.imgs.splice(imgReplaceIdx, 1, ...compressed);
		} else {
			row.imgs.push(...compressed);
		}
		saveSubTabsStorage();
		renderViewImgs();
		// 替换后若大图预览开着，刷新预览
		if (currentPreviewIdx !== null && document.getElementById('imgPreviewMask').style.display === 'flex') {
			const ni = imgMode === 'replace' ? imgReplaceIdx : row.imgs.length - 1;
			if (row.imgs[ni]) {
				document.getElementById('imgPreviewImg').src = row.imgs[ni];
				currentPreviewIdx = ni;
			}
		}
		e.target.value = "";
	})();
}

function removeViewImg(i) {
	if (isActiveLocked()) {
		alert("🔒 当前分组已锁定，禁止删除图片。");
		return;
	}
	const tab = getActiveSubTab();
	const row = tab.data[currentViewRowIdx];
	if (!row || !row.imgs) return;
	row.imgs.splice(i, 1);
	saveSubTabsStorage();
	renderViewImgs();
}

function previewImg(i) {
	const tab = getActiveSubTab();
	const row = tab.data[currentViewRowIdx];
	if (!row || !row.imgs || !row.imgs[i]) return;
	currentPreviewIdx = i;
	document.getElementById("imgPreviewImg").src = row.imgs[i];
	const locked = tab.locked === true;
	const rep = document.getElementById('replaceImgBtn'),
		del = document.getElementById('delPreviewImgBtn');
	if (rep) {
		rep.disabled = locked;
		rep.classList.toggle('btn-disabled', locked);
	}
	if (del) {
		del.disabled = locked;
		del.classList.toggle('btn-disabled', locked);
	}
	document.getElementById("imgPreviewMask").style.display = "flex";
}

function replacePreviewImg() {
	if (isActiveLocked()) {
		alert("🔒 当前分组已锁定，禁止替换图片。");
		return;
	}
	if (currentPreviewIdx === null) return;
	imgMode = 'replace';
	imgReplaceIdx = currentPreviewIdx;
	const inp = document.getElementById('viewImgInput');
	inp.value = "";
	inp.click();
}

function removePreviewImg() {
	if (isActiveLocked()) {
		alert("🔒 当前分组已锁定，禁止删除图片。");
		return;
	}
	if (currentPreviewIdx === null) return;
	removeViewImg(currentPreviewIdx);
	closeImgPreview();
}

function closeImgPreview() {
	document.getElementById("imgPreviewMask").style.display = "none";
	currentPreviewIdx = null;
}

function closeDetailView() {
	const m = document.getElementById("detailViewMask");
	if (m) m.style.display = "none";
}

function openNoteMask(idx) {
	const tab = getActiveSubTab();
	if (!tab.data[idx]) return;
	noteTarget = "detail";
	currentNoteRowIdx = idx;
	noteTextarea.value = tab.data[idx]["备注"] ?? "";
	noteTextarea.readOnly = tab.locked === true;
	noteMask.style.display = "flex";
}

function saveNote() {
	if (currentNoteRowIdx === null) return;
	if (noteTarget === "lib") {
		if (!productLib[currentNoteRowIdx]) return;
		productLib[currentNoteRowIdx]["备注"] = noteTextarea.value;
		saveLibData();
		renderLibTable();
		closeNoteMask();
		return;
	}
	if (isActiveLocked()) {
		alert("🔒 当前分组已锁定，禁止修改备注。");
		return;
	}
	const tab = getActiveSubTab();
	if (!tab.data[currentNoteRowIdx]) return;
	tab.data[currentNoteRowIdx]["备注"] = noteTextarea.value;
	saveSubTabsStorage();
	renderDetailTable();
	closeNoteMask();
}

function closeImportConfirmModal() {
	if (importConfirmModal) importConfirmModal.style.display = "none";
	tempImportDetailRows = [];
}

function closeDetailImportModal() {
	if (detailImportModal) detailImportModal.style.display = "none";
	tempDetailExcelRows = null;
}

function getActiveSubTab() {
	const aid = document.getElementById("activeSubTabId").value;
	for (let i = 0; i < detailSubTabs.length; i++) {
		if (detailSubTabs[i].id === aid) return detailSubTabs[i];
	}
	return detailSubTabs[0];
}

function saveSubTabsStorage() {
	localStorage.setItem("detailSubTabsStorage", JSON.stringify(detailSubTabs));
}

function renderSubSelect() {
	const sel = document.getElementById("subTabSelect");
	const activeId = document.getElementById("activeSubTabId").value;
	sel.innerHTML = "";
	detailSubTabs.forEach(tab => {
		const opt = document.createElement("option");
		opt.value = tab.id;
		opt.textContent = (tab.locked ? "🔒 " : "") + tab.name;
		if (tab.id === activeId) opt.selected = true;
		sel.appendChild(opt);
	});
	refreshLockBtn();
}

function isActiveLocked() {
	const tab = getActiveSubTab();
	return !!(tab && tab.locked);
}

function refreshLockBtn() {
	const btn = document.getElementById("lockGroupBtn");
	if (!btn) return;
	const locked = isActiveLocked();
	btn.innerHTML = locked ? "🔒 已锁定" : "🔓 未锁定";
	btn.classList.toggle("is-locked", locked);
	refreshLockButtons();
}

function refreshLockButtons() {
	const locked = isActiveLocked();
	document.querySelectorAll('#panel-detail button').forEach(b => {
		const oc = b.getAttribute('onclick') || "";
		const editOn = /renameCurrentSubTab|deleteCurrentSubTab|addDetailEmptyRow|readDetailExcel/.test(oc);
		if (editOn) {
			b.disabled = locked;
			b.classList.toggle('btn-disabled', locked);
		}
	});
}

function toggleLockSubTab() {
	const tab = getActiveSubTab();
	if (!tab) return;
	if (!tab.locked) {
		if (!confirm("确定锁定分组【" + tab.name + "】？锁定后禁止新增/修改/删除/导入，仅可查看、搜索与导出。")) return;
		tab.locked = true;
	} else {
		if (!confirm("确定解锁分组【" + tab.name + "】？解锁后可编辑。")) return;
		tab.locked = false;
	}
	saveSubTabsStorage();
	renderSubSelect();
	renderDetailTable();
}

function onSubSelectChange() {
	const sel = document.getElementById("subTabSelect");
	document.getElementById("activeSubTabId").value = sel.value;
	clearDetailSearch();
	renderDetailTable();
}

function addNewSubTab() {
	const newId = "sub_" + Date.now();
	const name = prompt("请输入新分组名称：", "新分组");
	if (!name || !name.trim()) return;
	detailSubTabs.push({
		id: newId,
		name: name.trim(),
		data: []
	});
	saveSubTabsStorage();
	document.getElementById("activeSubTabId").value = newId;
	renderSubSelect();
	renderDetailTable();
}

function renameCurrentSubTab() {
	if (isActiveLocked()) {
		alert("🔒 当前分组已锁定，禁止重命名，请先解锁。");
		return;
	}
	const tab = getActiveSubTab();
	const newName = prompt("重命名分组", tab.name);
	if (!newName || !newName.trim()) return;
	tab.name = newName.trim();
	saveSubTabsStorage();
	renderSubSelect();
}

function deleteCurrentSubTab() {
	if (isActiveLocked()) {
		alert("🔒 当前分组已锁定，禁止删除，请先解锁。");
		return;
	}
	if (detailSubTabs.length <= 1) {
		alert("⚠️至少保留一个分组，不允许全部删除");
		return;
	}
	const tab = getActiveSubTab();
	if (!confirm("确定删除分组【" + tab.name + "】？分组内全部数据将丢失！")) return;
	const idx = detailSubTabs.findIndex(t => t.id === tab.id);
	if (idx > -1) detailSubTabs.splice(idx, 1);
	saveSubTabsStorage();
	document.getElementById("activeSubTabId").value = detailSubTabs[0].id;
	renderSubSelect();
	renderDetailTable();
}

function downloadDetailGroupTemplate() {
	if (typeof XLSX === 'undefined') return alert("缺少Excel库(xlsx.full.min.js)，请把库文件放到js目录");
	const tab = getActiveSubTab();
	let headers;
	if (tab.data.length > 0) headers = Object.keys(tab.data[0]).filter(k => k !== "imgs");
	else if (productLib.length > 0) headers = Object.keys(productLib[0]).filter(k => k !== "imgs");
	else headers = ["型号", "备注"];
	const ws = XLSX.utils.aoa_to_sheet([headers]);
	const wb = XLSX.utils.book_new();
	XLSX.utils.book_append_sheet(wb, ws, tab.name + "模板");
	XLSX.writeFile(wb, tab.name + "_导入模板.xlsx");
}

//同义组
let synonymGroups = [];
(function loadSyn() {
	const s = localStorage.getItem("synonymGroups");
	if (s) {
		synonymGroups = JSON.parse(s);
	} else {
		synonymGroups = [
			["1920*1080", "1080p", "1k", "fhd", "1920×1080"],
			["2560*1440", "1440p", "2k", "qhd", "2560×1440"],
			["3840*2160", "2160p", "4k", "uhd", "3840×2160"],
			["210hz", "210"],
			["144hz", "144"],
			["60hz", "60"],
			["0.5ms", "0.5"],
			["1ms", "1"]
		];
		saveSyn();
	}
})();

function saveSyn() {
	localStorage.setItem("synonymGroups", JSON.stringify(synonymGroups));
}

function openSynModal() {
	renderSynList();
	if (synMask) synMask.style.display = "flex";
}

function closeSynModal() {
	if (synMask) synMask.style.display = "none";
}

function renderSynList() {
	const box = document.getElementById("synListBox");
	let html = "";
	synonymGroups.forEach((g, i) => {
		html +=
			`<div class="syn-item"><input value="${g.join(',')}" data-idx="${i}"><button class="btn-danger" style="width:auto;padding:4px 6px" onclick="delSyn(${i})">删除</button></div>`;
	});
	box.innerHTML = html;
}

function addSynLine() {
	synonymGroups.push([""]);
	renderSynList();
}

function delSyn(i) {
	synonymGroups.splice(i, 1);
	renderSynList();
}

function saveSynGroups() {
	const inputs = document.querySelectorAll("#synListBox .syn-item input");
	const newArr = [];
	inputs.forEach(inp => {
		const arr = inp.value.split(",").map(x => x.trim().toLowerCase()).filter(x => x);
		if (arr.length) newArr.push(arr);
	});
	synonymGroups = newArr;
	saveSyn();
	alert("同义组保存成功");
	closeSynModal();
}

function parseToken(text) {
	if (text === null || text === undefined || text === "") return [];
	let str = String(text).toLowerCase().replace(/×/g, "*").replace(/[^\w\*.\+\-\d%]/g, " ").replace(/\s+/g, " ")
.trim();
	const reg = /\d+\.?\d*(?:hz|k|p|nits|ms|bit|s|inch|w|v|a|%)?/g;
	const rawTokens = str.match(reg) || [];
	const tokens = [];
	rawTokens.forEach(t => {
		if (!(/^\d+$/.test(t) && t.length === 1)) tokens.push(t);
	});
	return [...new Set(tokens)];
}

function findSynGroup(token) {
	for (let g of synonymGroups) {
		if (g.includes(token)) return g;
	}
	return null;
}

function checkAnsSynonym(stdText, userText) {
	const stdTokens = parseToken(stdText);
	const userTokens = parseToken(userText);
	if (stdTokens.length === 0 || userTokens.length === 0) return false;
	for (let st of stdTokens) {
		const group = findSynGroup(st);
		let ok = false;
		if (group) {
			for (let ut of userTokens) {
				if (group.includes(ut)) {
					ok = true;
					break;
				}
			}
		} else {
			if (userTokens.includes(st)) ok = true;
		}
		if (!ok) return false;
	}
	return true;
}

function checkAnsBorder(stdText, userText) {
	const us = userText.trim().toLowerCase();
	const ss = stdText.trim().toLowerCase();
	if (us.length < 2) return false;
	const r1 = new RegExp("\\b" + us.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + "\\b");
	const r2 = new RegExp("\\b" + ss.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + "\\b");
	return r1.test(ss) || r2.test(us);
}

function checkAns(u, s) {
	const mode = document.querySelector('input[name="matchMode"]:checked').value;
	const us = String(u).trim().toLowerCase();
	const ss = String(s).trim().toLowerCase();
	if (mode === "exact") return us === ss;
	if (mode === "syn") return checkAnsSynonym(s, u);
	if (mode === "border") return checkAnsBorder(s, u);
	return false;
}

function saveLibData() {
	localStorage.setItem("productLibStorage", JSON.stringify(productLib));
}

function saveWrong() {
	localStorage.setItem("wrongPool", JSON.stringify(wrongPool));
}

//搜索
function onDetailSearch() {
	detailSearchText = document.getElementById("detailSearchInput").value.trim().toLowerCase();
	renderDetailTable();
}

function clearDetailSearch() {
	document.getElementById("detailSearchInput").value = "";
	detailSearchText = "";
}

function onLibSearch() {
	libSearchText = document.getElementById("libSearchInput").value.trim().toLowerCase();
	renderLibTable();
}

function clearLibSearch() {
	document.getElementById("libSearchInput").value = "";
	libSearchText = "";
	renderLibTable();
}

//产品详情导入
function readDetailExcel() {
	if (typeof XLSX === 'undefined') return alert("缺少Excel解析库(xlsx.full.min.js)，请把库文件放到js目录");
	if (isActiveLocked()) {
		alert("🔒 当前分组已锁定，禁止导入覆盖，请先解锁。");
		return;
	}
	if (!detailFileInput) {
		alert("文件控件未初始化，请刷新页面");
		return;
	}
	const f = detailFileInput.files[0];
	if (!f) return alert("请选择Excel文件");
	const r = new FileReader();
	r.onload = function(e) {
		const wb = XLSX.read(e.target.result, {
			type: 'binary'
		});
		const sheet = wb.Sheets[wb.SheetNames[0]];
		const rows = XLSX.utils.sheet_to_json(sheet);
		if (!rows.length) return alert("表格为空");
		tempDetailExcelRows = rows;
		const headers = Object.keys(rows[0]);
		detailKeyColSelect.innerHTML = "";
		headers.forEach(h => {
			const opt = document.createElement("option");
			opt.value = h;
			opt.innerText = h;
			if (h.includes("型号") || h.includes("SKU") || h.includes("物料")) opt.selected = true;
			detailKeyColSelect.appendChild(opt);
		});
		renderDetailImportPreview(rows.slice(0, 8));
		detailImportModal.style.display = "flex";
	};
	r.readAsBinaryString(f);
}

function renderDetailImportPreview(rows) {
	if (!rows.length) {
		detailPreviewBox.innerHTML = "无数据";
		return;
	}
	const h = Object.keys(rows[0]);
	let html = "<table><tr>" + h.map(x => `<th>${escapeHtml(x)}</th>`).join("") + "</tr>";
	rows.forEach(r => {
		html += "<tr>" + h.map(k => `<td>${escapeHtml(r[k]??"")}</td>`).join("") + "</tr>";
	});
	html += "</table>";
	detailPreviewBox.innerHTML = html;
}

function confirmDetailImport() {
	const tab = getActiveSubTab();
	const keyName = detailKeyColSelect.value.trim();
	if (!tempDetailExcelRows || !tempDetailExcelRows.length) {
		alert("没有待导入数据");
		return;
	}
	if (!keyName) {
		alert("请选择型号/SKU主键");
		return;
	}
	const targetMap = new Map();
	tempDetailExcelRows.forEach(r => {
		const sku = String(r[keyName] ?? "").trim();
		if (!sku) return;
		const obj = {};
		Object.keys(r).forEach(k => obj[k] = String(r[k] ?? "").trim());
		if (obj["备注"] === undefined) obj["备注"] = "";
		targetMap.set(sku, obj);
	});
	const newArr = Array.from(targetMap.values());
	if (detailDoAppendCheck.checked) {
		tab.data = tab.data.concat(newArr);
	} else {
		tab.data = newArr;
	}
	saveSubTabsStorage();
	renderDetailTable();
	closeDetailImportModal();
	alert("✅导入完成，有效：" + newArr.length + "条");
}

function closeModal() {
	if (importModal) importModal.style.display = "none";
	pendingRows = null;
}

function readCustomExcel() {
	if (typeof XLSX === 'undefined') return alert("缺少Excel解析库(xlsx.full.min.js)，请把库文件放到js目录");
	if (!fileInput) {
		alert("文件控件未初始化，请刷新页面");
		return;
	}
	const f = fileInput.files[0];
	if (!f) return alert("请选择Excel文件");
	const r = new FileReader();
	r.onload = function(e) {
		const wb = XLSX.read(e.target.result, {
			type: 'binary'
		});
		const sheet = wb.Sheets[wb.SheetNames[0]];
		const rows = XLSX.utils.sheet_to_json(sheet);
		if (!rows.length) return alert("表格为空");
		// 自动去除每个单元格首尾空格，避免参数带空格导致判分不一致
		const cleaned = rows.map(r => {
			const o = {};
			Object.keys(r).forEach(k => {
				o[k] = String(r[k] ?? "").trim();
			});
			return o;
		});
		pendingRows = cleaned;
		const headers = Object.keys(cleaned[0]);
		keyColSelect.innerHTML = "";
		headers.forEach(h => {
			const opt = document.createElement("option");
			opt.value = h;
			opt.innerText = h;
			if (h.includes("型号") || h.includes("SKU") || h.includes("物料")) opt.selected = true;
			keyColSelect.appendChild(opt);
		});
		renderPreview(cleaned.slice(0, 8));
		importModal.style.display = "flex";
	};
	r.readAsBinaryString(f);
}

function renderPreview(rows) {
	if (!rows.length) {
		previewBox.innerHTML = "无数据";
		return;
	}
	const h = Object.keys(rows[0]);
	let html = "<table><tr>" + h.map(x => `<th>${escapeHtml(x)}</th>`).join("") + "</tr>";
	rows.forEach(r => {
		html += "<tr>" + h.map(k => `<td>${escapeHtml(r[k]??"")}</td>`).join("") + "</tr>";
	});
	html += "</table>";
	previewBox.innerHTML = html;
}

function confirmImport() {
	const keyName = keyColSelect.value.trim();
	if (!pendingRows || !pendingRows.length) {
		alert("没有待导入数据");
		return;
	}
	if (!keyName) {
		alert("请选择型号/SKU主键");
		return;
	}
	const targetMap = new Map();
	pendingRows.forEach(r => {
		const sku = String(r[keyName] ?? "").trim();
		if (!sku) return;
		const obj = {};
		Object.keys(r).forEach(k => obj[k] = String(r[k] ?? "").trim());
		if (obj["备注"] === undefined) obj["备注"] = "";
		targetMap.set(sku, obj);
	});
	const newArr = Array.from(targetMap.values());
	if (dedupeCheck.checked) {
		const keep = new Map();
		productLib.forEach(p => keep.set(p[keyName], p));
		newArr.forEach(p => keep.set(p[keyName], p));
		productLib = Array.from(keep.values());
	} else {
		productLib = productLib.concat(newArr);
	}
	saveLibData();
	renderLibTable();
	closeModal();
	alert("✅产品库导入完成，有效：" + newArr.length + "条");
}

function downloadTemplate() {
	if (typeof XLSX === 'undefined') return alert("缺少Excel库(xlsx.full.min.js)，请把库文件放到js目录");
	let headers;
	if (productLib.length > 0) headers = Object.keys(productLib[0]);
	else headers = ["型号", "备注"];
	const ws = XLSX.utils.aoa_to_sheet([headers]);
	const wb = XLSX.utils.book_new();
	XLSX.utils.book_append_sheet(wb, ws, "模板");
	XLSX.writeFile(wb, "产品导入模板.xlsx");
}

function exportAll() {
	if (typeof XLSX === 'undefined') return alert("缺少Excel库(xlsx.full.min.js)，请把库文件放到js目录");
	if (!productLib.length) return alert("暂无产品库数据");
	const ws = XLSX.utils.json_to_sheet(productLib);
	const wb = XLSX.utils.book_new();
	XLSX.utils.book_append_sheet(wb, ws, "全部产品库");
	XLSX.writeFile(wb, "全部产品库.xlsx");
}

function clearAllLib() {
	if (!confirm("确定清空产品库？不可恢复！")) return;
	productLib = [];
	saveLibData();
	renderLibTable();
}

function addDetailEmptyRow() {
	if (isActiveLocked()) {
		alert("🔒 当前分组已锁定，禁止新增行，请先解锁。");
		return;
	}
	const tab = getActiveSubTab();
	let headers;
	if (tab.data.length > 0) headers = Object.keys(tab.data[0]).filter(k => k !== "imgs");
	else if (productLib.length > 0) headers = Object.keys(productLib[0]).filter(k => k !== "imgs");
	else headers = ["型号", "备注"];
	if (!headers.includes("备注")) headers.push("备注");
	const obj = {};
	headers.forEach(k => obj[k] = "");
	tab.data.push(obj);
	saveSubTabsStorage();
	renderDetailTable();
}
// 559行 修复：安全销毁拖拽实例
function destroySortable(sortableIns) {
	// 校验实例存在 + DOM节点还挂载在页面上
	if (sortableIns && sortableIns.el && document.body.contains(sortableIns.el)) {
		sortableIns.destroy();
	}
}

function renderDetailTable() {
	const tab = getActiveSubTab();
	const box = document.getElementById("detailTableBox");
	const locked = tab.locked === true;
	refreshLockButtons();

	// ===== 移动端：卡片只读列表（PC走下方原表格代码，完全不动）=====
	if (isMobile()) {
		const arr = tab.data;
		if (!arr.length) {
			box.innerHTML = "<div class='tip'>当前分组暂无数据</div>";
			return;
		}
		let showList = [...arr];
		if (detailSearchText) {
			showList = showList.filter(row => Object.values(row).some(v => String(v).toLowerCase().includes(
				detailSearchText)));
		}
		if (!showList.length) {
			box.innerHTML = "<div class='tip'>🔍无匹配数据</div>";
			return;
		}
		let html = '<div class="mobile-cards">';
		showList.forEach(row => {
			const realIdx = arr.indexOf(row);
			html += '<div class="m-card">';
			Object.keys(row).filter(k => k !== "imgs").forEach(key => {
				if (key === "备注") {
					const txt = row[key] || "";
					html +=
						`<div class="m-note" onclick="openNoteMask(${realIdx})">📝 ${escapeHtml(txt)||"(点击查看备注)"}</div>`;
				} else {
					html +=
						`<div class="m-field"><b>${escapeHtml(key)}：</b>${escapeHtml(row[key]??"")}</div>`;
				}
			});
			const delBtn = locked ? "" :
				`<button class="btn-danger" onclick="delDetailRow(${realIdx})">删除</button>`;
			html += `<div class="m-btn-group">
                <button class="btn-gray" onclick="openDetailView(${realIdx})">查看</button>
                <button class="btn-import" onclick="singleImportToLib(${realIdx})">导入库</button>
                ${delBtn}
            </div>`;
			html += '</div>';
		});
		html += '</div>';
		box.innerHTML = html;
		return;
	}

	if (sortDetailTbody) {
		destroySortable(sortDetailTbody);
		sortDetailTbody = null;
	}
	const arr = tab.data;
	if (!arr.length) {
		box.innerHTML = "<div class='tip'>当前分组暂无数据</div>";
		return;
	}
	let showList = [...arr];
	if (detailSearchText) {
		showList = showList.filter(row => {
			return Object.values(row).some(v => String(v).toLowerCase().includes(detailSearchText));
		});
	}
	if (!showList.length) {
		box.innerHTML = "<div class='tip'>🔍无匹配数据</div>";
		return;
	}
	let headers = Object.keys(showList[0]).filter(k => k !== "imgs");
	if (!headers.includes("备注")) {
		headers.push("备注");
		arr.forEach(i => {
			if (i["备注"] === undefined) i["备注"] = "";
		});
		saveSubTabsStorage();
	}
	let html =
		`<table class="prod-table"><thead><tr>${headers.map(h=>`<th>${escapeHtml(h)}</th>`).join("")}<th>操作</th></tr></thead><tbody id="detailTbody">`;
	showList.forEach(row => {
		const realIdx = arr.indexOf(row);
		html += `<tr data-real-idx="${realIdx}">`;
		headers.forEach(key => {
			if (key === "备注") {
				const txt = row[key] || "";
				html +=
					`<td class="note-cell" onclick="openNoteMask(${realIdx})"><div class="note-preview">${escapeHtml(txt)||"(点击查看备注)"}</div></td>`;
			} else {
				const ro = locked ? "readonly" : "";
				html +=
					`<td><input value="${escapeHtml(row[key]??"")}" ${ro} ondblclick="copyCellText(this.value)" onchange="detailCellChange(${realIdx},'${key}',this.value.trim())" onkeydown="if(event.key==='Enter')detailCellChange(${realIdx},'${key}',this.value.trim())"></td>`;
			}
		});
		const delBtn = locked ? "" :
			`<button class="btn-danger" style="width:auto;padding:3px 6px;font-size:12px" onclick="delDetailRow(${realIdx})">删除</button>`;
		html +=
			`<td style="white-space:nowrap;"><button class="btn-gray" style="width:auto;padding:3px 6px;font-size:12px" onclick="openDetailView(${realIdx})">查看</button> <button class="btn-import" style="width:auto;padding:3px 6px;font-size:12px" onclick="singleImportToLib(${realIdx})">导入库</button> ${delBtn}</td></tr>`;
	});
	html += "</tbody></table>";
	box.innerHTML = html;
	const tbody = document.getElementById("detailTbody");
	if (!locked && !detailSearchText && typeof Sortable !== 'undefined') {
		sortDetailTbody = new Sortable(tbody, {
			animation: 150,
			ghostClass: "sortable-ghost",
			chosenClass: "sortable-chosen",
			filter: "input,button",
			preventOnFilter: false,
			delayOnTouchOnly: 100,
			onEnd: function(evt) {
				const oldIndex = evt.oldIndex;
				const newIndex = evt.newIndex;
				if (oldIndex === newIndex) return;
				const item = arr.splice(oldIndex, 1)[0];
				arr.splice(newIndex, 0, item);
				saveSubTabsStorage();
				renderDetailTable();
			}
		});
	}
}

function detailCellChange(idx, key, val) {
	if (isActiveLocked()) {
		alert("🔒 当前分组已锁定，禁止修改。");
		return;
	}
	const tab = getActiveSubTab();
	if (!tab.data[idx]) return;
	tab.data[idx][key] = val;
	saveSubTabsStorage();
}

function delDetailRow(idx) {
	if (isActiveLocked()) {
		alert("🔒 当前分组已锁定，禁止删除行。");
		return;
	}
	if (!confirm("删除这条记录？")) return;
	const tab = getActiveSubTab();
	if (!tab.data[idx]) return;
	tab.data.splice(idx, 1);
	saveSubTabsStorage();
	renderDetailTable();
}

function renderLibTable() {
	const box = document.getElementById("libTableBox");

	// ===== 移动端：卡片只读列表（PC走下方原表格代码，完全不动）=====
	if (isMobile()) {
		if (!productLib.length) {
			box.innerHTML = "<div class='tip'>产品库暂无数据</div>";
			return;
		}
		let showList = [...productLib];
		if (libSearchText) {
			showList = showList.filter(row => Object.values(row).some(v => String(v).toLowerCase().includes(
				libSearchText)));
		}
		if (!showList.length) {
			box.innerHTML = "<div class='tip'>🔍无匹配数据</div>";
			return;
		}
		let html = '<div class="mobile-cards">';
		showList.forEach(row => {
			const realIdx = productLib.indexOf(row);
			html += '<div class="m-card">';
			Object.keys(row).filter(k => k !== "imgs").forEach(key => {
				if (key === "备注") {
					const txt = row[key] || "";
					html +=
						`<div class="m-note" onclick="openLibNoteMask(${realIdx})">📝 ${escapeHtml(txt)||"(点击填写备注)"}</div>`;
				} else {
					html +=
						`<div class="m-field"><b>${escapeHtml(key)}：</b>${escapeHtml(row[key]??"")}</div>`;
				}
			});
			html += `<div class="m-btn-group">
                <button class="btn-danger" onclick="delLibRow(${realIdx})">删除</button>
            </div>`;
			html += '</div>';
		});
		html += '</div>';
		box.innerHTML = html;
		return;
	}

	if (sortLibTbody) {
		destroySortable(sortLibTbody);
		sortLibTbody = null;
	}
	if (!productLib.length) {
		box.innerHTML = "<div class='tip'>产品库暂无数据</div>";
		return;
	}
	let showList = [...productLib];
	if (libSearchText) {
		showList = showList.filter(row => {
			return Object.values(row).some(v => String(v).toLowerCase().includes(libSearchText));
		});
	}
	if (!showList.length) {
		box.innerHTML = "<div class='tip'>🔍无匹配数据</div>";
		return;
	}
	const headers = Object.keys(showList[0]).filter(k => k !== "imgs");
	let html =
		`<table class="prod-table"><thead><tr>${headers.map(h=>`<th>${escapeHtml(h)}</th>`).join("")}<th>操作</th></tr></thead><tbody id="libTbody">`;
	showList.forEach(row => {
		const realIdx = productLib.indexOf(row);
		html += `<tr data-real-idx="${realIdx}">`;
		headers.forEach(key => {
			if (key === "备注") {
				const txt = row[key] || "";
				html +=
					`<td class="note-cell" onclick="openLibNoteMask(${realIdx})"><div class="note-preview">${escapeHtml(txt)||"(点击填写备注)"}</div></td>`;
			} else {
				html +=
					`<td><input value="${escapeHtml(row[key]??"")}" ondblclick="copyCellText(this.value)" onchange="libCellChange(${realIdx},'${key}',this.value.trim())" onkeydown="if(event.key==='Enter')libCellChange(${realIdx},'${key}',this.value.trim())"></td>`;
			}
		});
		html +=
			`<td><button class="btn-danger" style="width:auto;padding:4px 6px" onclick="delLibRow(${realIdx})">删除</button></td></tr>`;
	});
	html += "</tbody></table>";
	box.innerHTML = html;
	const tbody = document.getElementById("libTbody");
	if (!libSearchText && typeof Sortable !== 'undefined') {
		sortLibTbody = new Sortable(tbody, {
			animation: 150,
			ghostClass: "sortable-ghost",
			chosenClass: "sortable-chosen",
			filter: "input,button",
			preventOnFilter: false,
			delayOnTouchOnly: 100,
			onEnd: function(evt) {
				const oldIndex = evt.oldIndex;
				const newIndex = evt.newIndex;
				if (oldIndex === newIndex) return;
				const item = productLib.splice(oldIndex, 1)[0];
				productLib.splice(newIndex, 0, item);
				saveLibData();
				renderLibTable();
			}
		});
	}
}

function openLibNoteMask(idx) {
	if (!productLib[idx]) return;
	noteTarget = "lib";
	currentNoteRowIdx = idx;
	noteTextarea.value = productLib[idx]["备注"] ?? "";
	noteMask.style.display = "flex";
}

function libCellChange(idx, key, val) {
	if (!productLib[idx]) return;
	productLib[idx][key] = val;
	saveLibData();
}

function delLibRow(idx) {
	if (!confirm("确定删除？")) return;
	if (!productLib[idx]) return;
	productLib.splice(idx, 1);
	saveLibData();
	renderLibTable();
}

function singleImportToLib(detailIdx) {
	const tab = getActiveSubTab();
	tempImportDetailRows = [JSON.parse(JSON.stringify(tab.data[detailIdx]))];
	importConfirmModal.style.display = "flex";
}

function batchImportDetailToLib() {
	const tab = getActiveSubTab();
	if (!tab.data.length) return alert("当前分组无数据");
	tempImportDetailRows = JSON.parse(JSON.stringify(tab.data));
	importConfirmModal.style.display = "flex";
}

function doImportFromDetail() {
	const mode = document.querySelector('input[name="importMode"]:checked').value;
	let keyName = null;
	if (tempImportDetailRows.length > 0) {
		const keys = Object.keys(tempImportDetailRows[0]);
		keyName = keys.find(k => k.toLowerCase().includes("型号") || k.toLowerCase().includes("sku"));
	}
	if (!keyName) {
		alert("找不到型号/SKU字段");
		closeImportConfirmModal();
		return;
	}
	const existMap = new Map();
	productLib.forEach(p => existMap.set(String(p[keyName]).trim(), p));
	let addCnt = 0,
		overCnt = 0,
		skipCnt = 0;
	tempImportDetailRows.forEach(item => {
		delete item.imgs; // 产品库不携带图片
		const sku = String(item[keyName] ?? "").trim();
		if (!sku) return;
		if (existMap.has(sku)) {
			if (mode === "over") {
				existMap.set(sku, item);
				overCnt++;
			} else skipCnt++;
		} else {
			existMap.set(sku, item);
			addCnt++;
		}
	});
	productLib = Array.from(existMap.values());
	saveLibData();
	renderLibTable();
	closeImportConfirmModal();
	alert(`导入完成：新增${addCnt}，覆盖${overCnt}，跳过${skipCnt}`);
}

function shuffle(arr) {
	const a = [...arr];
	for (let i = a.length - 1; i > 0; i--) {
		const j = Math.floor(Math.random() * (i + 1));
		[a[i], a[j]] = [a[j], a[i]];
	}
	return a;
}

function startQuiz() {
	const mode = document.querySelector('input[name="quizMode"]:checked').value;
	if (!productLib.length) return alert("产品库为空，请导入数据");
	const firstProd = productLib[0];
	const modelKey = Object.keys(firstProd).find(k => k.toLowerCase().includes("型号") || k.toLowerCase().includes(
	"sku"));
	if (!modelKey) return alert("未识别型号列");
	const need = Number(document.getElementById("qty").value);
	if (!need || need <= 0) {
		alert("题目数量必须大于0");
		return;
	}

	// 指定型号过滤：留空=全部；逗号/换行/分号分隔，忽略多余空格与空行
	const filterText = (document.getElementById('modelFilter')?.value || "").trim();
	let wanted = null;
	if (filterText) {
		wanted = new Set(filterText.split(/[,，\n;；]/).map(s => s.replace(/\s+/g, "").trim()).filter(Boolean));
	}
	const libPool = productLib.filter(p => {
		if (!wanted) return true;
		return wanted.has(String(p[modelKey] ?? "").replace(/\s+/g, "").trim());
	});
	if (!libPool.length) return alert("指定型号在产品库中未找到，或留空以使用全部产品");

	if (mode === "param") {
		const pool = [];
		libPool.forEach(p => {
			const keys = Object.keys(p).filter(k => k !== modelKey && k !== '备注');
			keys.forEach(f => {
				const ans = String(p[f] ?? "").trim();
				if (ans) {
					pool.push({
						type: "param",
						prod: p,
						modelKey,
						field: f,
						stdAns: ans,
						qText: `【${p[modelKey]}】的${f}是多少？`,
						marked: false,
						userAns: ""
					});
				}
			});
		});
		if (!pool.length) return alert("无出题数据");
		if (need > pool.length) {
			alert(`可出题量不足：当前可用 ${pool.length} 题，已按 ${pool.length} 出题`);
		}
		quizList = shuffle(pool).slice(0, Math.min(need, pool.length));
	} else {
		const pool = [];
		libPool.forEach(p => {
			const fields = Object.keys(p).filter(k => k !== modelKey && k !== '备注' && String(p[k]).trim());
			if (fields.length < 2) return;
			const pickCnt = Math.min(fields.length, Math.floor(Math.random() * 2) + 2);
			const picked = shuffle(fields).slice(0, pickCnt);
			const desc = picked.map(f => `${f}：${p[f]}`).join("，");
			pool.push({
				type: "model",
				prod: p,
				modelKey,
				pickedFields: picked,
				stdAns: p[modelKey],
				qText: `已知：${desc}，请问对应的型号是？`,
				marked: false,
				userAns: ""
			});
		});
		if (!pool.length) return alert("满足条件产品不足");
		if (need > pool.length) {
			alert(`可出题量不足：当前可用 ${pool.length} 题，已按 ${pool.length} 出题`);
		}
		quizList = shuffle(pool).slice(0, Math.min(need, pool.length));
	}
	currentQ = 0;
	renderQuiz();
	if (document.getElementById('useCountdown').checked) {
		startTimer();
	} else {
		document.getElementById("timerBox").style.display = "none";
	}
}

function markQuestion() {
	if (!quizList.length) return;
	quizList[currentQ].marked = !quizList[currentQ].marked;
	renderQuiz();
}

function startWrongQuiz() {
	if (!wrongPool.length) return alert("错题库为空");
	const need = Number(document.getElementById("qty").value);
	if (!need || need <= 0) {
		alert("题目数量必须大于0");
		return;
	}
	const map = new Map();
	wrongPool.forEach(item => {
		if (!map.has(item.qText)) map.set(item.qText, item);
	});
	const unique = Array.from(map.values());
	if (need > unique.length) {
		alert(`错题库去重后仅 ${unique.length} 题，已按 ${unique.length} 出题`);
	}
	quizList = shuffle(unique).slice(0, Math.min(need, unique.length));
	currentQ = 0;
	renderQuiz();
	if (document.getElementById('useCountdown').checked) {
		startTimer();
	} else {
		document.getElementById("timerBox").style.display = "none";
	}
}

function resetQuiz() {
	clearInterval(timer);
	timer = null;
	quizList = [];
	currentQ = 0;
	document.getElementById("quizBox").innerHTML = "<div class='tip'>先导入【产品库管理】再来答题</div>";
	document.getElementById("qIndexBox").innerHTML = "";
	document.getElementById("resultBox").classList.add("hide");
	document.getElementById("inlineResult").classList.add("hide");
	document.getElementById("quizBox").classList.remove("hide");
	document.querySelectorAll("#quizPanel .m-submit-btn").forEach(b => b.classList.remove("hide"));
	document.getElementById("quizPanel").classList.remove("hide");
	document.getElementById("quizPanel").classList.remove("ready");
	document.getElementById("timerBox").style.display = "none";
	document.getElementById("timerBox").classList.remove("warn");
}

function renderQuiz() {
	// ready=已生成题目；PC始终显示空白框，仅手机端未出题时隐藏
	const hasQ = quizList && quizList.length > 0;
	document.getElementById('quizPanel').classList.toggle('ready', hasQ);
	// 每次渲染题目都回到答题视图：隐藏结果区，恢复题目框与提交按钮（修复交卷后重抽题不显示）
	document.getElementById("inlineResult").classList.add("hide");
	document.getElementById("resultBox").classList.add("hide");
	document.getElementById("quizBox").classList.remove("hide");
	document.querySelectorAll("#quizPanel .m-submit-btn").forEach(b => b.classList.remove("hide"));
	document.getElementById("quizPanel").classList.remove("hide");
	const idxBox = document.getElementById("qIndexBox");
	idxBox.innerHTML = "";
	quizList.forEach(it => {
		if (it.userAns === undefined) it.userAns = "";
	});
	quizList.forEach((it, i) => {
		const b = document.createElement("button");
		b.className = "index-btn";
		if (i === currentQ) b.classList.add("active");
		if (it.userAns.trim()) b.classList.add("answered");
		if (it.marked) b.classList.add("mark");
		b.innerText = String(i + 1);
		b.onclick = () => {
			saveCurrentAns();
			currentQ = i;
			renderQuiz();
		};
		idxBox.appendChild(b);
	});
	const q = quizList[currentQ];
	document.getElementById("quizBox").innerHTML = `
<div class="q-card">
<div class="q-text">第${currentQ+1}/${quizList.length} ${q.marked?" ⚠️【已标记存疑】":""}<br>${q.qText}</div>
<input class="ans-input" id="ansInput" value="${q.userAns}" placeholder="填写答案">
<div class="ctrl-row">
<button class="btn-primary" onclick="saveCurrentAns();prev()">上一题</button>
<button class="btn-success" onclick="saveCurrentAns();next()">下一题</button>
</div>
</div>`;
	document.getElementById("ansInput").focus();
}

function saveCurrentAns() {
	const inp = document.getElementById("ansInput");
	if (!inp) return;
	quizList[currentQ].userAns = inp.value;
}

function prev() {
	if (currentQ > 0) {
		currentQ--;
		renderQuiz();
	}
}

function next() {
	if (currentQ < quizList.length - 1) {
		currentQ++;
		renderQuiz();
	}
}

function startTimer() {
	clearInterval(timer);
	const min = Number(document.getElementById("timeLimit").value);
	if (!min || min <= 0) {
		alert("答题时长必须大于0，已取消倒计时");
		document.getElementById("timerBox").style.display = "none";
		return;
	}
	const limit = min * 60;
	let left = limit;
	const txt = document.getElementById("timerText");
	const timerDom = document.getElementById("timerBox");
	timerDom.style.display = "block";
	timerDom.classList.remove("warn");
	timer = setInterval(() => {
		left--;
		if (left <= 30) timerDom.classList.add("warn");
		if (left <= 0) {
			clearInterval(timer);
			alert("时间到！自动交卷");
			submitAll(true);
			return;
		}
		const m = String(Math.floor(left / 60)).padStart(2, "0");
		const s = String(left % 60).padStart(2, "0");
		txt.innerText = `${m}:${s}`;
	}, 1000);
}

function findFirstUnAnswered() {
	return quizList.findIndex(it => (it.userAns || "").trim() === "");
}

function submitAll(force) {
	clearInterval(timer);
	timer = null;
	saveCurrentAns();
	quizList.forEach(it => {
		if (it.userAns === undefined) it.userAns = "";
	});
	const firstUn = findFirstUnAnswered();
	if (!force && firstUn !== -1) {
		alert(`还有未作答，跳转到${firstUn+1}题`);
		currentQ = firstUn;
		renderQuiz();
		return;
	}
	let ok = 0,
		no = 0,
		un = 0;
	let detailHtml = "";
	quizList.forEach(it => {
		const u = it.userAns.trim();
		if (u === "") {
			un++;
			detailHtml +=
				`<div class="detail-item unanswered"><div><strong>${it.qText}</strong>${it.marked?" ⚠️标记":""}</div><div>你的答案：(未作答)</div><div>标准答案：${it.stdAns}</div><div>🟡未作答</div></div>`;
		} else {
			const isOk = checkAns(u, it.stdAns);
			if (isOk) ok++;
			else no++;
			detailHtml += `<div class="detail-item ${isOk?'right':'wrong'}">
<div><strong>${it.qText}</strong>${it.marked?" ⚠️标记":""}</div>
<div>你的答案：${it.userAns}</div>
<div>标准答案：${it.stdAns}</div>
<div>${isOk?"✅正确":"❌错误"}</div>
${!isOk?`<button class="btn-gray" style="width:auto;margin-top:4px;padding:4px 8px" onclick="toggleWrong('${it.qText}')">从错题本移除</button>`:""}
</div>`;
			if (!isOk) {
				const has = wrongPool.some(x => x.qText === it.qText);
				if (!has) wrongPool.push(it);
			}
		}
	});
	saveWrong();
	document.getElementById("rTotal").innerText = quizList.length;
	document.getElementById("rOk").innerText = ok;
	document.getElementById("rNo").innerText = no;
	document.getElementById("rUn").innerText = un;
	document.getElementById("rRate").innerText = quizList.length ? ((ok / quizList.length) * 100).toFixed(1) + "%" :
		"0%";
	document.getElementById("resultDetail").innerHTML = detailHtml;
	// PC端：结果直接显示在答题区位置；手机端：隐藏答题区，用独立结果面板
	const isM = window.innerWidth <= 768;
	document.getElementById("rTotal2").innerText = quizList.length;
	document.getElementById("rOk2").innerText = ok;
	document.getElementById("rNo2").innerText = no;
	document.getElementById("rUn2").innerText = un;
	document.getElementById("rRate2").innerText = quizList.length ? ((ok / quizList.length) * 100).toFixed(1) + "%" :
		"0%";
	document.getElementById("resultDetail2").innerHTML = detailHtml;
	const panel = document.getElementById("quizPanel");
	if (isM) {
		panel.classList.add("hide");
		document.getElementById("inlineResult").classList.add("hide");
		document.getElementById("resultBox").classList.remove("hide");
	} else {
		document.getElementById("resultBox").classList.add("hide");
		document.getElementById("quizBox").classList.add("hide");
		document.querySelectorAll("#quizPanel .m-submit-btn").forEach(b => b.classList.add("hide"));
		document.getElementById("inlineResult").classList.remove("hide");
	}
}

function toggleWrong(qText) {
	wrongPool = wrongPool.filter(x => x.qText !== qText);
	saveWrong();
	submitAll();
}

function backQuiz() {
	document.getElementById("resultBox").classList.add("hide");
	document.getElementById("inlineResult").classList.add("hide");
	document.getElementById("quizBox").classList.remove("hide");
	document.querySelectorAll("#quizPanel .m-submit-btn").forEach(b => b.classList.remove("hide"));
	document.getElementById("quizPanel").classList.remove("hide");
}

function exportResult() {
	if (typeof XLSX === 'undefined') return alert("缺少Excel库(xlsx.full.min.js)，请把库文件放到js目录");
	const exp = quizList.map((it, i) => {
		const u = (it.userAns || "").trim();
		let stat;
		if (u === "") stat = "未作答";
		else stat = checkAns(u, it.stdAns) ? "正确" : "错误";
		return {
			题号: i + 1,
			题目: it.qText,
			你的答案: it.userAns || "",
			标准答案: it.stdAns,
			是否正确: stat,
			标记存疑: it.marked ? "是" : "否"
		};
	});
	const ws = XLSX.utils.json_to_sheet(exp);
	const wb = XLSX.utils.book_new();
	XLSX.utils.book_append_sheet(wb, ws, "答题结果");
	XLSX.writeFile(wb, "答题结果.xlsx");
}

function exportWrong() {
	if (typeof XLSX === 'undefined') return alert("缺少Excel库(xlsx.full.min.js)，请把库文件放到js目录");
	if (!wrongPool.length) return alert("暂无错题");
	const exp = wrongPool.map(it => ({
		题目: it.qText,
		标准答案: it.stdAns
	}));
	const ws = XLSX.utils.json_to_sheet(exp);
	const wb = XLSX.utils.book_new();
	XLSX.utils.book_append_sheet(wb, ws, "错题本");
	XLSX.writeFile(wb, "错题本.xlsx");
}

// ====================== 全量JSON备份 / 恢复（新增） ======================
function backupAllData() {
	const all = {
		detailSubTabs,
		productLib,
		synonymGroups,
		wrongPool,
		exportTime: new Date().toLocaleString()
	};
	const blob = new Blob([JSON.stringify(all, null, 2)], {
		type: "application/json"
	});
	const a = document.createElement('a');
	a.href = URL.createObjectURL(blob);
	a.download = "产品考核全量备份_" + stampNow() + ".json";
	a.click();
	URL.revokeObjectURL(a.href);
}

function stampNow() {
	const d = new Date();
	const p = n => String(n).padStart(2, "0");
	return `${d.getFullYear()}${p(d.getMonth()+1)}${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}`;
}

function exportCurrentSubTabJson() {
	const tab = getActiveSubTab();
	if (!tab) return alert("没有可导出的分组");
	if (!tab.data || !tab.data.length) return alert("当前分组无数据");
	const payload = {
		type: "singleSubTab",
		subTab: {
			id: tab.id,
			name: tab.name,
			data: tab.data
		},
		exportTime: new Date().toLocaleString()
	};
	const blob = new Blob([JSON.stringify(payload, null, 2)], {
		type: "application/json"
	});
	const a = document.createElement('a');
	a.href = URL.createObjectURL(blob);
	a.download = (tab.name || "分组") + "_" + stampNow() + ".json";
	a.click();
	URL.revokeObjectURL(a.href);
}

function findModelKey(list) {
	if (!Array.isArray(list) || !list.length) return null;
	return Object.keys(list[0]).find(k => k.toLowerCase().includes("型号") || k.toLowerCase().includes("sku")) || null;
}

function restoreAllData(file, mode) {
	const reader = new FileReader();
	reader.onload = function(e) {
		try {
			const d = JSON.parse(e.target.result);
			if (!d || typeof d !== 'object') throw new Error("格式");
			const incomingTabs = Array.isArray(d.detailSubTabs) ? d.detailSubTabs : [];
			const incomingLib = Array.isArray(d.productLib) ? d.productLib : [];
			const incomingSyn = Array.isArray(d.synonymGroups) ? d.synonymGroups : [];
			const incomingWrong = Array.isArray(d.wrongPool) ? d.wrongPool : [];

			if (mode === 'merge') {
				// 追加合并：产品库按型号主键覆盖/新增，现有保留
				const mKey = findModelKey(incomingLib) || findModelKey(productLib);
				if (mKey) {
					const existMap = new Map();
					productLib.forEach(p => existMap.set(String(p[mKey] ?? "").trim(), p));
					incomingLib.forEach(item => {
						const sku = String(item[mKey] ?? "").trim();
						if (!sku) return;
						existMap.set(sku, item);
					});
					productLib = Array.from(existMap.values());
				} else {
					productLib = productLib.concat(incomingLib);
				}
				// 详情分组：同名分组追加数据（组内按型号去重，本地优先），不同名整体新增
				incomingTabs.forEach(itab => {
					const same = detailSubTabs.find(t => t.name === itab.name);
					const inRows = Array.isArray(itab.data) ? itab.data : [];
					if (same) {
						const mk = findModelKey(same.data) || findModelKey(inRows);
						if (mk) {
							const exist = new Set(same.data.map(r => String(r[mk] ?? "").trim()));
							inRows.forEach(r => {
								const sku = String(r[mk] ?? "").trim();
								if (!sku) return;
								if (exist.has(sku)) return; // 本地已有该型号 → 跳过，本地优先
								exist.add(sku);
								same.data.push(r);
							});
						} else {
							same.data = same.data.concat(inRows);
						}
					} else {
						detailSubTabs.push({
							id: itab.id || ("sub_" + Date.now() + Math.random()),
							name: itab.name || "导入分组",
							data: inRows
						});
					}
				});
				// 同义组、错题追加去重
				const synSet = new Set(synonymGroups.map(g => JSON.stringify(g)));
				incomingSyn.forEach(g => {
					const s = JSON.stringify(g);
					if (!synSet.has(s)) {
						synSet.add(s);
						synonymGroups.push(g);
					}
				});
				const wrongSet = new Set(wrongPool.map(w => w.qText + "|" + w.stdAns));
				incomingWrong.forEach(w => {
					const k = w.qText + "|" + w.stdAns;
					if (!wrongSet.has(k)) {
						wrongSet.add(k);
						wrongPool.push(w);
					}
				});
			} else {
				detailSubTabs = incomingTabs.length ? incomingTabs : [{
					id: "sub_" + Date.now(),
					name: "分组1",
					data: []
				}];
				productLib = incomingLib;
				synonymGroups = incomingSyn;
				wrongPool = incomingWrong;
			}

			detailSubTabs.forEach(tab => {
				if (!Array.isArray(tab.data)) tab.data = [];
				tab.data.forEach(row => {
					if (row["备注"] === undefined) row["备注"] = "";
					if (!Array.isArray(row.imgs)) row.imgs = [];
				});
			});
			productLib = productLib.map(item => {
				if (item && item["备注"] === undefined) item["备注"] = "";
				return item;
			});
			saveSubTabsStorage();
			saveLibData();
			saveSyn();
			saveWrong();
			document.getElementById("activeSubTabId").value = detailSubTabs[0].id;
			renderSubSelect();
			renderDetailTable();
			renderLibTable();
			renderSynList();
			alert(mode === 'merge' ? "追加合并完成" : "恢复成功");
		} catch (err) {
			alert("备份文件格式错误，无法恢复");
		}
	};
	reader.readAsText(file);
}

function pickRestoreFile() {
	const input = document.createElement('input');
	input.type = 'file';
	input.accept = '.json';
	input.onchange = e => {
		if (!e.target.files[0]) return;
		const mode = confirm("恢复方式：\n【确定】= 完全覆盖当前所有数据\n【取消】= 追加合并到现有数据（不删除已有内容）");
		const ok = confirm(mode ? "确定完全覆盖？当前数据将被替换" : "确定追加合并？备份内容将并入现有数据");
		if (!ok) {
			e.target.value = "";
			return;
		}
		restoreAllData(e.target.files[0], mode ? 'over' : 'merge');
		e.target.value = "";
	};
	input.click();
}