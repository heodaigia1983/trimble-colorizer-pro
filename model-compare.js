/* IFC comparison overlay. The large IFC files are processed by the local IFC Diff app. */
(function () {
  "use strict";
  var report = null;
  var mapped = new Map();
  var painted = null;
  var busy = false;
  var COLORS = { A: "#00D46A", M: "#FF8A00", N: "#1E90FF", T: "#A33BFF", R: "#F12448", B: "#00C5E8" };
  var LABELS = { A: "Chỉ có mốc mới", M: "Thay đổi", N: "Đổi mã", T: "Tạo lại / dựng lại", R: "Chỉ có mốc cũ", GOP: "Gộp vào assembly", B: "Bu lông mới" };
  var $ = function (id) { return document.getElementById(id); };

  function show(message, bad) {
    $("compareStatus").textContent = message;
    $("compareStatus").classList.toggle("compare-error", !!bad);
    if (typeof log === "function") log((bad ? "✗ " : "✓ ") + message, bad ? "warn" : "info");
  }
  function setBusy(value) {
    busy = value;
    ["compareRefresh", "comparePaint", "compareClear", "compareFile"].forEach(function (id) { $(id).disabled = value; });
  }
  function number(value) { return Number(value || 0).toLocaleString("vi-VN"); }
  function tons(value) { return Number(value || 0).toLocaleString("vi-VN", { minimumFractionDigits: 3, maximumFractionDigits: 3 }); }
  function sameName(a, b) { return String(a || "").trim().toLowerCase() === String(b || "").trim().toLowerCase(); }
  function modelId(model) { return model.versionId || model.id; }
  function selectedModels() {
    var oldId = $("compareOldModel").value, newId = $("compareNewModel").value;
    if (!oldId || !newId || oldId === newId) throw new Error("Chọn hai model cũ và mới khác nhau đang mở trong Viewer.");
    var oldModel = models.find(function (model) { return modelId(model) === oldId; });
    var newModel = models.find(function (model) { return modelId(model) === newId; });
    if (!oldModel || !newModel) throw new Error("Model đã đổi; bấm Làm mới danh sách rồi chọn lại.");
    return { old: oldModel, newer: newModel, oldId: oldId, newId: newId };
  }
  var models = [];

  async function refreshModels() {
    if (busy) return;
    try {
      var api = await getAPI();
      var oldValue = $("compareOldModel").value, newValue = $("compareNewModel").value;
      var found;
      try { found = await api.viewer.getModels("loaded"); }
      catch (_) { found = null; }
      if (!Array.isArray(found) || !found.length) found = (await api.viewer.getModels()).filter(function (model) { return model.state === "loaded"; });
      models = found.filter(function (model) { return model && modelId(model); });
      ["compareOldModel", "compareNewModel"].forEach(function (id) {
        var select = $(id); select.replaceChildren();
        var prompt = new Option("Chọn model đang mở", ""); select.add(prompt);
        models.forEach(function (model) { select.add(new Option(model.name + " · " + String(modelId(model)).slice(0, 8), modelId(model))); });
      });
      $("compareOldModel").value = oldValue;
      $("compareNewModel").value = newValue;
      if (report && !sameName(report.meta.old_file, report.meta.new_file)) {
        models.forEach(function (model) {
          if (sameName(model.name, report.meta.old_file)) $("compareOldModel").value = modelId(model);
          if (sameName(model.name, report.meta.new_file)) $("compareNewModel").value = modelId(model);
        });
      }
      show("Viewer đang mở " + number(models.length) + " model. Chọn đúng hai mốc IFC.");
    } catch (error) { show("Không đọc được model Trimble: " + error.message, true); }
  }

  function validReport(data) {
    return data && data.meta && typeof data.meta.old_file === "string" && typeof data.meta.new_file === "string"
      && Array.isArray(data.items) && Array.isArray(data.bolts) && data.summary && data.summary.counts
      && data.summary.selfcheck && data.summary.selfcheck.pass === true;
  }
  async function loadReport(event) {
    var file = event.target.files && event.target.files[0];
    if (!file) return;
    try {
      if (!/\.json$/i.test(file.name) || file.size > 100 * 1024 * 1024) throw new Error("Chọn file diff_*.json (tối đa 100 MB) từ công cụ So sánh IFC.");
      var data = JSON.parse(await file.text());
      if (!validReport(data)) throw new Error("File kết quả thiếu nguồn IFC, danh sách thay đổi hoặc kiểm tra khối lượng không đạt.");
      if (painted) await clearOverlay(await getAPI());
      report = data; mapped.clear();
      $("compareConfirmNames").checked = false;
      $("compareSource").textContent = "Cũ: " + data.meta.old_file + "  →  Mới: " + data.meta.new_file;
      $("compareWarning").textContent = "Phạm vi: " + (data.meta.scope_status || "Chưa xác nhận")
        + ". " + ((data.meta.export_warnings || []).map(function (w) { return w.option + " " + w.old + "→" + w.new; }).join("; ") || "Không thấy khác cấu hình trong báo cáo.")
        + " Số tấn chỉ là kết quả đối chiếu, chưa dùng đánh giá nhân viên khi phạm vi xuất chưa xác nhận. Khôi phục màu trả màu gốc model, không lấy lại màu tô tay trước đó.";
      var s = data.summary, c = s.counts;
      $("compareKpis").innerHTML = '<div><b>' + number(c.A) + '</b><span>Chỉ có mốc mới</span></div>'
        + '<div><b>' + number(c.M) + '</b><span>Thay đổi</span></div>'
        + '<div><b>' + number(c.R) + '</b><span>Chỉ có mốc cũ</span></div>'
        + '<div><b>' + tons(s.them_t) + '</b><span>Tấn ứng viên thêm</span></div>'
        + '<div><b>' + tons(s.bo_t) + '</b><span>Tấn ứng viên bỏ</span></div>'
        + '<div><b>' + tons(s.rong_t) + '</b><span>Tấn ròng tham khảo</span></div>';
      renderRows();
      await refreshModels();
      show("Đã đọc " + number(data.items.length) + " dòng thay đổi; chọn hai model rồi bấm Tô so sánh.");
    } catch (error) { show(error.message, true); }
  }

  function renderRows() {
    if (!report) return;
    var filter = $("compareFilter").value;
    var rows = report.items.filter(function (item) { return filter === "all" || item.status === filter; });
    $("compareRowsCount").textContent = "Hiện " + number(Math.min(rows.length, 200)) + " / " + number(rows.length) + " dòng; bấm dòng để chọn trên 3D.";
    var tbody = $("compareRows"); tbody.replaceChildren();
    rows.slice(0, 200).forEach(function (item) {
      var tr = document.createElement("tr");
      [LABELS[item.status] || item.status, item.name || "—", item.tag || "—", item.profile || "—", item.weight_kg == null ? "—" : tons(item.weight_kg / 1000)].forEach(function (value) {
        var td = document.createElement("td"); td.textContent = value; tr.appendChild(td);
      });
      tr.title = (item.changes || []).map(function (change) { return change.f + ": " + change.old + " → " + change.new; }).join("\n") || item.k || "";
      tr.addEventListener("click", function () { selectItem(item); });
      tbody.appendChild(tr);
    });
  }

  async function convert(api, modelIdValue, entries, target) {
    var unique = Array.from(new Set(entries.map(function (entry) { return entry.guid; }).filter(function (guid) { return typeof guid === "string" && guid.length === 22; })));
    var hits = 0;
    for (var i = 0; i < unique.length; i += 250) {
      var guids = unique.slice(i, i + 250);
      var uuids = guids.map(ifc2uuid);
      var ids = await api.viewer.convertToObjectRuntimeIds(modelIdValue, uuids);
      if (!Array.isArray(ids) || ids.length !== uuids.length) throw new Error("Viewer trả danh sách ID không khớp cho model " + modelIdValue.slice(0, 8));
      ids.forEach(function (id, n) { if (Number.isFinite(Number(id)) && id != null) { target.set(guids[n], Number(id)); hits++; } });
      $("compareProgress").textContent = "Đang ghép GUID: " + number(Math.min(i + 250, unique.length)) + "/" + number(unique.length) + " (khớp " + number(hits) + ")";
    }
    return { total: unique.length, hits: hits };
  }
  async function paintIds(api, model, ids, color) {
    var array = Array.from(new Set(ids));
    for (var i = 0; i < array.length; i += 300) {
      await api.viewer.setObjectState({ modelObjectIds: [{ modelId: model, objectRuntimeIds: array.slice(i, i + 300) }] }, { color: color });
    }
  }
  async function clearOverlay(api) {
    if (!painted) return;
    for (var entry of painted.newIds) await paintIds(api, painted.newId, entry[1], "reset");
    await api.viewer.setObjectState({ modelObjectIds: [{ modelId: painted.oldId }] }, { color: "reset" });
    painted = null; mapped.clear();
  }
  async function clearComparison() {
    if (busy) return;
    setBusy(true);
    try { await clearOverlay(await getAPI()); $("compareProgress").textContent = ""; show("Đã khôi phục màu gốc của hai model trong chế độ so sánh."); }
    catch (error) { show("Không khôi phục được toàn bộ màu: " + error.message, true); }
    finally { setBusy(false); }
  }

  async function paintComparison() {
    if (busy) return;
    if (!report) { show("Chọn file diff_*.json trước khi tô so sánh.", true); return; }
    setBusy(true);
    try {
      var pair = selectedModels();
      if ((!sameName(pair.old.name, report.meta.old_file) || !sameName(pair.newer.name, report.meta.new_file)) && !$("compareConfirmNames").checked) {
        throw new Error("Tên model không khớp báo cáo. Kiểm tra hai phiên bản đã chọn; nếu đúng, tích ô xác nhận tên khác rồi tô lại.");
      }
      var api = await getAPI();
      await clearOverlay(api);
      var onlyOld = report.items.filter(function (item) { return item.status === "R"; });
      var onlyNew = report.items.filter(function (item) { return item.status === "A"; });
      var oldSample = onlyOld.slice(0, 40), newSample = onlyNew.slice(0, 40);
      if (oldSample.length && newSample.length) {
        var oldCorrect = await convert(api, pair.oldId, oldSample, new Map());
        var oldWrong = await convert(api, pair.newId, oldSample, new Map());
        var newCorrect = await convert(api, pair.newId, newSample, new Map());
        var newWrong = await convert(api, pair.oldId, newSample, new Map());
        if (oldWrong.hits > oldCorrect.hits || newWrong.hits > newCorrect.hits) {
          throw new Error("Hai mốc có vẻ bị chọn ngược: GUID chỉ có ở một phía khớp nhiều hơn với model đối diện. Kiểm tra IFC cũ/mới.");
        }
      }
      var oldItems = report.items.filter(function (item) { return item.status === "R" || item.status === "GOP" || item.status === "T"; });
      var newItems = report.items.filter(function (item) { return ["A", "M", "N", "T"].includes(item.status); });
      var newBolts = report.bolts.filter(function (item) { return item.status === "A"; });
      var oldMap = new Map(), newMap = new Map();
      var oldResult = await convert(api, pair.oldId, oldItems, oldMap);
      var newResult = await convert(api, pair.newId, newItems.concat(newBolts), newMap);
      if ((oldResult.total && oldResult.hits / oldResult.total < 0.1) || (newResult.total && newResult.hits / newResult.total < 0.1)) {
        throw new Error("GUID khớp quá ít (cũ " + number(oldResult.hits) + "/" + number(oldResult.total) + ", mới " + number(newResult.hits) + "/" + number(newResult.total) + "). Kiểm tra model và file diff.");
      }
      mapped = new Map();
      oldMap.forEach(function (id, guid) { mapped.set("old:" + guid, id); });
      newMap.forEach(function (id, guid) { mapped.set("new:" + guid, id); });
      await api.viewer.setObjectState({ modelObjectIds: [{ modelId: pair.oldId }] }, { color: "#9AA2AE" });
      painted = { oldId: pair.oldId, newId: pair.newId, newIds: [] };
      var group = [
        ["R", pair.oldId, report.items.filter(function (item) { return item.status === "R"; }), oldMap],
        ["A", pair.newId, report.items.filter(function (item) { return item.status === "A"; }), newMap],
        ["M", pair.newId, report.items.filter(function (item) { return item.status === "M"; }), newMap],
        ["N", pair.newId, report.items.filter(function (item) { return item.status === "N"; }), newMap],
        ["T", pair.newId, report.items.filter(function (item) { return item.status === "T"; }), newMap],
        ["B", pair.newId, newBolts, newMap]
      ];
      for (var row of group) {
        var ids = row[2].map(function (item) { return row[3].get(item.guid); }).filter(function (id) { return id != null; });
        await paintIds(api, row[1], ids, COLORS[row[0]]);
        if (row[1] === pair.newId) painted.newIds.push([row[0], ids]);
        $("compareProgress").textContent = "Đã tô " + LABELS[row[0]] + ": " + number(ids.length) + " cấu kiện";
      }
      show("Đã tô so sánh. GUID khớp: cũ " + number(oldResult.hits) + "/" + number(oldResult.total) + ", mới " + number(newResult.hits) + "/" + number(newResult.total) + ". Số liệu tấn vẫn cần xác nhận cùng phạm vi xuất.");
    } catch (error) { show(error.message, true); }
    finally { setBusy(false); }
  }

  async function selectItem(item) {
    if (!painted || busy) { show("Bấm Tô so sánh trước khi chọn cấu kiện.", true); return; }
    var side = item.status === "R" || item.status === "GOP" || (item.status === "T" && !mapped.has("new:" + item.guid)) ? "old" : "new";
    var id = mapped.get(side + ":" + item.guid);
    if (id == null) { show("Cấu kiện này chưa ghép được với Viewer; xem tên model và GUID.", true); return; }
    try {
      var api = await getAPI(), selector = { modelObjectIds: [{ modelId: side === "old" ? painted.oldId : painted.newId, objectRuntimeIds: [id] }] };
      await api.viewer.setSelection(selector, "set");
      await api.viewer.setCamera(selector);
      show("Đã chọn " + (item.tag || item.name || item.guid) + " trên model " + (side === "old" ? "cũ" : "mới") + ".");
    } catch (error) { show("Không chọn được cấu kiện: " + error.message, true); }
  }

  document.addEventListener("DOMContentLoaded", function () {
    $("compareRefresh").addEventListener("click", refreshModels);
    $("compareFile").addEventListener("change", loadReport);
    $("compareFilter").addEventListener("change", renderRows);
    $("comparePaint").addEventListener("click", paintComparison);
    $("compareClear").addEventListener("click", clearComparison);
  });
})();
