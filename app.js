let map;
let currentRow = null;
let allData = [];
let hoods = [];
let markers = [];
let activeFilters = {};

// ── Add-Point mode ──────────────────────────────────────────
let addPointMode = false;
let addPointMarker = null;
let hoodLayers = [];

// ── Extra layer data & markers ──────────────────────────────
let hotspotData      = [];
let demandData       = [];
let idleData         = [];
let centroidData     = [];

let hotspotMarkers   = [];
let demandMarkers    = [];
let idleMarkers      = [];
let centroidMarkers  = [];

// ── Layer visibility state (legend toggles) ─────────────────
const layerVisible = {
  hoods:      true,
  properties: true,
  hotspots:   true,
  demand:     true,
  idle:       true,
  centroids:  true
};

console.log("🚀 App initializing...");
init();

async function init() {
  map = L.map('map').setView([12.9, 77.65], 12);
  L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png').addTo(map);
  setTimeout(() => map.invalidateSize(), 300);

  hoods = await fetch("hoods.json").then(r => r.json());
  console.log(`📦 hoods.json loaded — ${hoods.length} hoods`);

  drawHoods();
  buildLegend();

  await loadData();
  await loadExtraLayers();

  // ── NO auto-refresh intervals — manual only ──────────────

  initMapSearch();

  map.on("click", function (e) {
    if (addPointMode) {
      openAddPointModal(e.latlng.lat, e.latlng.lng);
      return;
    }

    // Show a copyable lat/long popup on every normal click
    const { lat, lng } = e.latlng;
    const coordStr = `${lat.toFixed(7)}, ${lng.toFixed(7)}`;
    L.popup({ closeButton: true, className: "latlng-popup" })
      .setLatLng(e.latlng)
      .setContent(`
        <div style="font-size:13px;line-height:1.6">
          <div style="font-weight:600;margin-bottom:4px">📌 Coordinates</div>
          <code style="background:#f4f4f4;padding:3px 7px;border-radius:4px;font-size:12px;display:block;margin-bottom:8px">${lat.toFixed(7)}, ${lng.toFixed(7)}</code>
          <button onclick="
            navigator.clipboard.writeText('${coordStr}')
              .then(() => { this.textContent='✅ Copied!'; setTimeout(()=>this.textContent='📋 Copy',1500); })
              .catch(() => { this.textContent='❌ Failed'; setTimeout(()=>this.textContent='📋 Copy',1500); });
          " style="
            width:100%;padding:5px 0;border:none;border-radius:5px;
            background:#2980b9;color:#fff;font-size:12px;font-weight:600;cursor:pointer
          ">📋 Copy</button>
        </div>
      `)
      .openOn(map);
  });
}

// ============================================================
// LEGEND
// ============================================================
function buildLegend() {
  const legend = document.getElementById("mapLegend");
  if (!legend) return;

  const items = [
    { key: "hoods",      color: "#4da6ff", symbol: "■", label: "Hood Polygons"   },
    { key: "properties", color: "#e74c3c", symbol: "📍", label: "Properties"     },
    { key: "hotspots",   color: "#f39c12", symbol: "H",  label: "Hotspots"       },
    { key: "demand",     color: "#2980b9", symbol: "●",  label: "Demand (size = point count)"  },
    { key: "idle",       color: "#c0392b", symbol: "●",  label: "Idle (size = idle minutes)"    },
    { key: "centroids",  color: "#27ae60", symbol: "C",  label: "Demand Centroids"}
  ];

  legend.innerHTML = `<div class="legend-title">Layers</div>` +
    items.map(item => `
      <div class="legend-item" id="legend_${item.key}" onclick="toggleLayer('${item.key}')" style="cursor:pointer">
        <span class="legend-symbol" style="color:${item.color};font-weight:bold">${item.symbol}</span>
        <span class="legend-label">${item.label}</span>
        <span class="legend-eye" id="eye_${item.key}">👁</span>
      </div>
    `).join("");
}

function toggleLayer(key) {
  layerVisible[key] = !layerVisible[key];
  const eye  = document.getElementById("eye_" + key);
  const item = document.getElementById("legend_" + key);
  if (eye)  eye.textContent  = layerVisible[key] ? "👁" : "🚫";
  if (item) item.style.opacity = layerVisible[key] ? "1" : "0.4";

  if (key === "hoods") {
    hoodLayers.forEach(({ layer }) => {
      layerVisible.hoods ? map.addLayer(layer) : map.removeLayer(layer);
    });
  }
  if (key === "properties") {
    markers.forEach(m => layerVisible.properties ? map.addLayer(m) : map.removeLayer(m));
  }
  if (key === "hotspots") {
    hotspotMarkers.forEach(m => layerVisible.hotspots ? map.addLayer(m) : map.removeLayer(m));
  }
  if (key === "demand") {
    demandMarkers.forEach(m => layerVisible.demand ? map.addLayer(m) : map.removeLayer(m));
  }
  if (key === "idle") {
    idleMarkers.forEach(m => layerVisible.idle ? map.addLayer(m) : map.removeLayer(m));
  }
  if (key === "centroids") {
    centroidMarkers.forEach(m => layerVisible.centroids ? map.addLayer(m) : map.removeLayer(m));
  }
}

// ============================================================
// EXTRA LAYERS LOADING
// ============================================================
async function loadExtraLayers() {
  console.log("📡 loadExtraLayers()");
  await Promise.all([
    loadLayer(CONFIG.HOTSPOT_URL,  "hotspots",  renderHotspots),
    loadLayer(CONFIG.DEMAND_URL,   "demand",    renderDemand),
    loadLayer(CONFIG.IDLE_URL,     "idle",      renderIdle),
    loadLayer(CONFIG.CENTROID_URL, "centroids", renderCentroids),
  ]);
}

async function loadLayer(url, name, renderFn) {
  if (!url) { console.warn(`⚠️ No URL configured for ${name}`); return; }
  try {
    const res  = await fetch(url + "?t=" + Date.now());
    const data = await res.json();
    console.log(`✅ ${name}: ${data.length} rows`);
    renderFn(data);
  } catch (err) {
    console.error(`❌ Failed to load ${name}:`, err);
  }
}

// ── Icon factories ───────────────────────────────────────────
function letterIcon(letter, bg, textColor = "#fff") {
  return L.divIcon({
    className: "",
    html: `<div style="
      background:${bg};color:${textColor};border-radius:50%;
      width:24px;height:24px;display:flex;align-items:center;
      justify-content:center;font-weight:700;font-size:13px;
      border:2px solid rgba(0,0,0,0.25);box-shadow:0 1px 3px rgba(0,0,0,0.3)
    ">${letter}</div>`,
    iconSize: [24, 24],
    iconAnchor: [12, 12]
  });
}

function dotIcon(color) {
  return L.divIcon({
    className: "",
    html: `<div style="
      background:${color};border-radius:50%;
      width:10px;height:10px;
      border:1.5px solid rgba(0,0,0,0.3);
      box-shadow:0 1px 3px rgba(0,0,0,0.3)
    "></div>`,
    iconSize: [10, 10],
    iconAnchor: [5, 5]
  });
}

// Scaled dot icon — size and color intensity vary with value.
function scaledDotIcon(value, min, max, hLow, hHigh) {
  const MIN_R = 5, MAX_R = 22;
  const t = (max > min) ? Math.max(0, Math.min(1, (value - min) / (max - min))) : 0.5;
  const r = Math.round(MIN_R + t * (MAX_R - MIN_R));
  const hue  = Math.round(hLow + t * (hHigh - hLow));
  const sat  = 85;
  const lite = Math.round(72 - t * 47);
  const color = `hsl(${hue},${sat}%,${lite}%)`;
  const borderAlpha = (0.2 + t * 0.5).toFixed(2);

  return L.divIcon({
    className: "",
    html: `<div style="
      background:${color};
      border-radius:50%;
      width:${r*2}px;height:${r*2}px;
      border:1.5px solid rgba(0,0,0,${borderAlpha});
      box-shadow:0 1px 4px rgba(0,0,0,${(0.2+t*0.3).toFixed(2)});
      opacity:0.88;
    "></div>`,
    iconSize: [r*2, r*2],
    iconAnchor: [r, r]
  });
}

// ── NM/MM lookup for extra-layer rows ────────────────────────
function stampHoodInfo(rows, latKey, lngKey) {
  rows.forEach(row => {
    if (row._nm) return;
    const lat = parseFloat(row[latKey]), lng = parseFloat(row[lngKey]);
    if (isNaN(lat) || isNaN(lng)) return;
    const hood = assignHood({ lat, lng });
    if (hood) {
      row._nm = hood.nano_market;
      row._mm = hood.micro_market;
    }
  });
}

// Returns true if the row should be visible given current NM/MM filters
function passesNMMFilter(row) {
  const filterNM = activeFilters.NM || "";
  const filterMM = activeFilters.MM || "";
  if (!filterNM && !filterMM) return true;
  if (filterNM && row._nm !== filterNM) return false;
  if (filterMM && row._mm !== filterMM) return false;
  return true;
}

// ── Render functions (filter-aware) ──────────────────────────
function renderHotspots(data) {
  hotspotMarkers.forEach(m => map.removeLayer(m));
  hotspotMarkers = [];
  hotspotData = data;
  stampHoodInfo(data, "lat", "lng");
  data.forEach(row => {
    const lat = parseFloat(row.lat), lng = parseFloat(row.lng);
    if (isNaN(lat) || isNaN(lng)) return;
    const m = L.marker([lat, lng], { icon: letterIcon("H", "#f39c12") })
      .bindPopup(`<b>🔥 ${row.name || "Hotspot"}</b><br>Hood: ${row.hood || "-"}<br>Cluster: ${row.cluster || "-"}<br>NM: ${row._nm || "-"}<br>MM: ${row._mm || "-"}`);
    m._extraRow = row;
    if (layerVisible.hotspots && passesNMMFilter(row)) m.addTo(map);
    hotspotMarkers.push(m);
  });
  console.log(`📍 ${hotspotMarkers.length} hotspot markers`);
}

function renderDemand(data) {
  demandMarkers.forEach(m => map.removeLayer(m));
  demandMarkers = [];
  demandData = data;
  stampHoodInfo(data, "lat", "lng");

  const vals = data.map(r => parseFloat(r.num_points)).filter(v => !isNaN(v));
  const minV = vals.length ? Math.min(...vals) : 0;
  const maxV = vals.length ? Math.max(...vals) : 1;

  data.forEach(row => {
    const lat = parseFloat(row.lat), lng = parseFloat(row.lng);
    if (isNaN(lat) || isNaN(lng)) return;
    const numPts = parseFloat(row.num_points) || 0;
    const icon = scaledDotIcon(numPts, minV, maxV, 200, 220);
    const m = L.marker([lat, lng], { icon })
      .bindPopup(`
        <b>📦 Demand</b><br>
        Cluster: ${row.cluster || "-"}<br>
        Points: ${row.num_points || "-"}<br>
        NM: ${row._nm || "-"}<br>
        MM: ${row._mm || "-"}
      `);
    m._extraRow = row;
    if (layerVisible.demand && passesNMMFilter(row)) m.addTo(map);
    demandMarkers.push(m);
  });
  console.log(`📍 ${demandMarkers.length} demand markers (scaled by num_points, min:${minV} max:${maxV})`);
}

function renderIdle(data) {
  idleMarkers.forEach(m => map.removeLayer(m));
  idleMarkers = [];
  idleData = data;
  stampHoodInfo(data, "lat", "lng");

  const vals = data.map(r => parseFloat(r.idle_min)).filter(v => !isNaN(v));
  const minV = vals.length ? Math.min(...vals) : 0;
  const maxV = vals.length ? Math.max(...vals) : 1;

  data.forEach(row => {
    const lat = parseFloat(row.lat), lng = parseFloat(row.lng);
    if (isNaN(lat) || isNaN(lng)) return;
    const idleMin = parseFloat(row.idle_min) || 0;
    const icon = scaledDotIcon(idleMin, minV, maxV, 5, 0);
    const m = L.marker([lat, lng], { icon })
      .bindPopup(`
        <b>🚗 Idle</b><br>
        Cluster: ${row.cluster || "-"}<br>
        Hood: ${row.hood || "-"}<br>
        Idle min: ${row.idle_min || "-"}<br>
        NM: ${row._nm || "-"}<br>
        MM: ${row._mm || "-"}
      `);
    m._extraRow = row;
    if (layerVisible.idle && passesNMMFilter(row)) m.addTo(map);
    idleMarkers.push(m);
  });
  console.log(`📍 ${idleMarkers.length} idle markers (scaled by idle_min, min:${minV} max:${maxV})`);
}

function renderCentroids(data) {
  centroidMarkers.forEach(m => map.removeLayer(m));
  centroidMarkers = [];
  centroidData = data;
  stampHoodInfo(data, "centroid_lat", "centroid_lng");
  data.forEach(row => {
    const lat = parseFloat(row.centroid_lat), lng = parseFloat(row.centroid_lng);
    if (isNaN(lat) || isNaN(lng)) return;
    const m = L.marker([lat, lng], { icon: letterIcon("C", "#27ae60") })
      .bindPopup(`<b>📊 ${row.hood_name || "Centroid"}</b><br>Cluster ID: ${row.cluster_id || "-"}<br>NM: ${row._nm || "-"}<br>MM: ${row._mm || "-"}`);
    m._extraRow = row;
    if (layerVisible.centroids && passesNMMFilter(row)) m.addTo(map);
    centroidMarkers.push(m);
  });
  console.log(`📍 ${centroidMarkers.length} centroid markers`);
}

// ── Filter extra layers by NM/MM ─────────────────────────────
function filterExtraLayers() {
  const sets = [
    { markerList: hotspotMarkers,  visKey: "hotspots"  },
    { markerList: demandMarkers,   visKey: "demand"    },
    { markerList: idleMarkers,     visKey: "idle"      },
    { markerList: centroidMarkers, visKey: "centroids" },
  ];
  sets.forEach(({ markerList, visKey }) => {
    markerList.forEach(m => {
      const show = layerVisible[visKey] && passesNMMFilter(m._extraRow || {});
      show ? (map.hasLayer(m) || m.addTo(map)) : (map.hasLayer(m) && map.removeLayer(m));
    });
  });
}

// ── Filtered data getters for downloads ──────────────────────
function getFilteredHotspots()  { return hotspotData.filter(passesNMMFilter);  }
function getFilteredDemand()    { return demandData.filter(passesNMMFilter);    }
function getFilteredIdle()      { return idleData.filter(passesNMMFilter);      }
function getFilteredCentroids() { return centroidData.filter(passesNMMFilter);  }

// ============================================================
// MAIN DATA LOADING
// ============================================================
async function loadData() {
  const url = CONFIG.API_URL + "?t=" + Date.now();
  console.log("📡 loadData() —", url);
  try {
    const res  = await fetch(url);
    const text = await res.text();
    const data = JSON.parse(text);
    allData = data;
    console.log(`✅ ${allData.length} rows loaded`);

    // Render markers — preserve current zoom/pan, no fitBounds
    if (Object.values(activeFilters).some(v => v)) {
      filterAndRender();
    } else {
      renderMarkers();
    }

    // Sheet preview — re-apply sheet filters if active
    const sfActive = Object.values(getSheetFilters()).some(v => v);
    renderSheetPreview(sfActive ? getSheetFilteredData() : allData);

    renderSummaryTables();
  } catch (err) {
    console.error("❌ Fetch failed:", err);
  }
  populateFilters();
}

// ============================================================
// RENDER PROPERTY MARKERS
// ============================================================
function getPropertyName(row) {
  return row["Name of the property"] || row["Name"] || "No Name";
}

function renderMarkers() {
  markers.forEach(m => map.removeLayer(m));
  markers = [];
  let skipped = 0;
  allData.forEach(row => {
    const lat = parseFloat(row.Lat), lng = parseFloat(row.Long);
    if (!isNaN(lat) && !isNaN(lng)) {
      if (!row.NM || !row.MM) {
        const hood = assignHood({ lat, lng });
        if (hood) {
          row.NM = hood.nano_market;
          row.MM = hood.micro_market;
          row["NM Id"] = hood.hood_id;
          if (row._rowIndex) {
            fetch(CONFIG.API_URL, { method: "POST", body: JSON.stringify(row) })
              .catch(err => console.error(`❌ Auto-save failed`, err));
          }
        }
      }
      const name   = getPropertyName(row);
      const marker = L.marker([lat, lng], { icon: getCategoryIcon(row.Category) })
        .bindPopup(`<b>${name}</b><br>${row.Category || ""}<br>NM: ${row.NM || "-"}<br>MM: ${row.MM || "-"}`)
        .on('click', () => showDetails(row));
      if (layerVisible.properties) marker.addTo(map);
      markers.push(marker);
    } else {
      skipped++;
    }
  });
  console.log(`📍 ${markers.length} markers, ${skipped} skipped`);
}

// ============================================================
// FILTERS
// ============================================================
function applyFilters() {
  activeFilters = {
    Category:       document.getElementById("filterCategory").value,
    Property:       document.getElementById("filterProperty").value,
    "App status":   document.getElementById("filterAppStatus").value,
    "Lead Status":  document.getElementById("filterLeadStatus").value,
    "Final Status": document.getElementById("filterFinalStatus").value,
    NM:             document.getElementById("filterNM").value,
    MM:             document.getElementById("filterMM").value,
    dateFrom:       document.getElementById("filterDateFrom")?.value || "",
    dateTo:         document.getElementById("filterDateTo")?.value   || "",
  };
  filterAndRender();
}

function filterAndRender() {
  const from = activeFilters.dateFrom ? new Date(activeFilters.dateFrom + "T00:00:00") : null;
  const to   = activeFilters.dateTo   ? new Date(activeFilters.dateTo   + "T23:59:59") : null;

  const filtered = allData.filter(row => {
    // Standard column filters
    const colKeys = ["Category", "Property", "App status", "Lead Status", "Final Status", "NM", "MM"];
    for (const key of colKeys) {
      if (activeFilters[key] && row[key] !== activeFilters[key]) return false;
    }
    // Date range filter on Timestamp
    if (from || to) {
      const ts = parseTimestamp(row["Timestamp"]);
      if (!ts) return false;
      if (from && ts < from) return false;
      if (to   && ts > to)   return false;
    }
    return true;
  });

  console.log(`🔽 ${filtered.length}/${allData.length} rows match filters`);
  updateHoodVisibility();
  renderFilteredMarkers(filtered, true); // fitView=true when user explicitly filters
  filterExtraLayers();
}

function renderFilteredMarkers(data, fitView = false) {
  markers.forEach(m => map.removeLayer(m));
  markers = [];
  const bounds = [];
  data.forEach(row => {
    const lat = parseFloat(row.Lat), lng = parseFloat(row.Long);
    if (!isNaN(lat) && !isNaN(lng)) {
      const name   = getPropertyName(row);
      const marker = L.marker([lat, lng], { icon: getCategoryIcon(row.Category) })
        .bindPopup(`<b>${name}</b><br>${row.Category || ""}<br>NM: ${row.NM || "-"}<br>MM: ${row.MM || "-"}`)
        .on('click', () => showDetails(row));
      if (layerVisible.properties) marker.addTo(map);
      markers.push(marker);
      bounds.push([lat, lng]);
    }
  });
  // Only zoom/pan to fit when user explicitly requests it (fitView flag)
  // Never on background reloads — preserves the user's current map position
  if (fitView && bounds.length) map.fitBounds(bounds);
}

function clearFilters() {
  activeFilters = {};
  // Scope to map filter bar selects only — don't touch sheet filter selects
  document.querySelectorAll(".map-filter-bar select").forEach(s => s.value = "");
  const fd = document.getElementById("filterDateFrom");
  const td = document.getElementById("filterDateTo");
  if (fd) fd.value = "";
  if (td) td.value = "";
  updateHoodVisibility();
  renderMarkers();
  filterExtraLayers();
}

function populateFilters() {
  const fields = [
    { key: "Category",      id: "filterCategory"    },
    { key: "Property",      id: "filterProperty"    },
    { key: "App status",    id: "filterAppStatus"   },
    { key: "Lead Status",   id: "filterLeadStatus"  },
    { key: "Final Status",  id: "filterFinalStatus" },
    { key: "NM",            id: "filterNM"          },
    { key: "MM",            id: "filterMM"          }
  ];
  fields.forEach(f => {
    const select = document.getElementById(f.id);
    if (!select) return;
    const current = select.value;
    const values  = [...new Set(allData.map(r => r[f.key]).filter(Boolean))].sort();
    select.innerHTML = `<option value="">${f.key}</option>` +
      values.map(v => `<option value="${v}">${v}</option>`).join("");
    select.value = current;
  });
}

// ============================================================
// ADD POINT BY CLICKING MAP
// ============================================================
const ADD_POINT_FIELDS = [
  { key: "Name of the property",                  label: "Property Name",         type: "text" },
  { key: "Category",                              label: "Category",              type: "text" },
  { key: "Sub Category",                          label: "Sub Category",          type: "text" },
  { key: "Road",                                  label: "Road",                  type: "text" },
  { key: "Property",                              label: "Property Type",         type: "text" },
  { key: "App status",                            label: "App Status",            type: "text" },
  { key: "Lead Status",                           label: "Lead Status",           type: "text" },
  { key: "Final Status",                          label: "Final Status",          type: "text" },
  { key: "Location (Google Maps URL) / Map Code", label: "Maps Link / Plus Code", type: "url"  },
  { key: "Contact Name",                          label: "Contact Name",          type: "text" },
  { key: "Contact number",                        label: "Contact Number",        type: "tel"  },
  { key: "Comment",                               label: "Comment",               type: "text" },
  { key: "Restroom ID",                           label: "Restroom ID",           type: "text" },
];

function toggleAddPointMode() {
  addPointMode = !addPointMode;
  const btn = document.getElementById("btnAddPoint");
  if (addPointMode) {
    btn.textContent = "❌ Cancel Add Point";
    btn.style.background = "#c0392b";
    map.getContainer().style.cursor = "crosshair";
  } else {
    btn.textContent = "➕ Add Point";
    btn.style.background = "";
    map.getContainer().style.cursor = "";
    if (addPointMarker) { map.removeLayer(addPointMarker); addPointMarker = null; }
  }
}

function openAddPointModal(lat, lng) {
  if (addPointMarker) map.removeLayer(addPointMarker);
  addPointMarker = L.marker([lat, lng], {
    icon: L.divIcon({ className: "custom-icon", html: `<div style="font-size:24px">📌</div>` })
  }).addTo(map);

  const container = document.getElementById("addPointFields");
  if (!container) return;

  container.innerHTML = `
    <div class="modal-field">
      <label>Lat (auto)</label>
      <input type="number" id="ap_Lat" value="${lat.toFixed(7)}" readonly style="background:#f0f0f0" />
    </div>
    <div class="modal-field">
      <label>Long (auto)</label>
      <input type="number" id="ap_Long" value="${lng.toFixed(7)}" readonly style="background:#f0f0f0" />
    </div>
  ` + ADD_POINT_FIELDS.map(f => `
    <div class="modal-field">
      <label>${f.label}</label>
      <input type="${f.type}" id="ap_${f.key.replace(/[\s.()/]/g,'_')}" placeholder="${f.label} (optional)" />
    </div>
  `).join("");

  document.getElementById("addPointModal").style.display = "flex";
}

function closeAddPointModal() {
  document.getElementById("addPointModal").style.display = "none";
  if (addPointMarker) { map.removeLayer(addPointMarker); addPointMarker = null; }
  if (addPointMode) toggleAddPointMode();
}

async function submitAddPoint() {
  const lat = document.getElementById("ap_Lat").value;
  const lng = document.getElementById("ap_Long").value;
  const newRow = {
    "Lat":  parseFloat(lat),
    "Long": parseFloat(lng)
  };

  ADD_POINT_FIELDS.forEach(f => {
    const el = document.getElementById("ap_" + f.key.replace(/[\s.()/]/g,'_'));
    if (el && el.value.trim()) {
      newRow[f.key] = f.type === "number" ? parseFloat(el.value) : el.value.trim();
    }
  });

  const hood = assignHood({ lat: parseFloat(lat), lng: parseFloat(lng) });
  if (hood) {
    newRow.NM = hood.nano_market;
    newRow.MM = hood.micro_market;
    newRow["NM Id"] = hood.hood_id;
  }

  try {
    const res  = await fetch(CONFIG.API_URL, { method: "POST", body: JSON.stringify(newRow) });
    const json = await res.json();
    if (json.success) {
      alert(`✅ Point added\nNM: ${newRow.NM || "-"}, MM: ${newRow.MM || "-"}`);
      closeAddPointModal();
      loadData();
    } else {
      alert("❌ Failed: " + JSON.stringify(json));
    }
  } catch (err) {
    alert("❌ Error: " + err.message);
  }
}

// ============================================================
// DOWNLOAD KML / CSV (WKT)
// ============================================================
function getFilteredData() {
  const hasFilter = Object.values(activeFilters).some(v => v);
  if (!hasFilter) return allData;
  const from = activeFilters.dateFrom ? new Date(activeFilters.dateFrom + "T00:00:00") : null;
  const to   = activeFilters.dateTo   ? new Date(activeFilters.dateTo   + "T23:59:59") : null;
  return allData.filter(row => {
    const colKeys = ["Category", "Property", "App status", "Lead Status", "Final Status", "NM", "MM"];
    for (const key of colKeys) {
      if (activeFilters[key] && row[key] !== activeFilters[key]) return false;
    }
    if (from || to) {
      const ts = parseTimestamp(row["Timestamp"]);
      if (!ts) return false;
      if (from && ts < from) return false;
      if (to   && ts > to)   return false;
    }
    return true;
  });
}

function getFilteredNMs() { return [...new Set(getFilteredData().map(r => r.NM).filter(Boolean))]; }
function getFilteredMMs() { return [...new Set(getFilteredData().map(r => r.MM).filter(Boolean))]; }
function hoodsByNM(nmList) { return hoods.filter(h => nmList.includes(h.nano_market)); }
function hoodsByMM(mmList) { return hoods.filter(h => mmList.includes(h.micro_market)); }

function geojsonCoordToKmlRing(coords) { return coords.map(c => `${c[0]},${c[1]},0`).join(" "); }

function geometryToKmlGeometry(geometry) {
  if (!geometry) return "";
  if (geometry.type === "Polygon") {
    const outer = geometry.coordinates[0];
    const inner = geometry.coordinates.slice(1);
    return `<Polygon>
      <outerBoundaryIs><LinearRing><coordinates>${geojsonCoordToKmlRing(outer)}</coordinates></LinearRing></outerBoundaryIs>
      ${inner.map(r => `<innerBoundaryIs><LinearRing><coordinates>${geojsonCoordToKmlRing(r)}</coordinates></LinearRing></innerBoundaryIs>`).join("")}
    </Polygon>`;
  }
  if (geometry.type === "MultiPolygon") {
    return `<MultiGeometry>${geometry.coordinates.map(poly => {
      const outer = poly[0], inner = poly.slice(1);
      return `<Polygon>
        <outerBoundaryIs><LinearRing><coordinates>${geojsonCoordToKmlRing(outer)}</coordinates></LinearRing></outerBoundaryIs>
        ${inner.map(r => `<innerBoundaryIs><LinearRing><coordinates>${geojsonCoordToKmlRing(r)}</coordinates></LinearRing></innerBoundaryIs>`).join("")}
      </Polygon>`;
    }).join("")}</MultiGeometry>`;
  }
  return "";
}

function escXml(str) {
  return String(str || "").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
}

function hoodsToKml(hoodList, layerName, color = "7f0000ff") {
  const placemarks = hoodList.map(h => `
  <Placemark>
    <name>${escXml(h.nano_market || h.micro_market || h.hood_id)}</name>
    <description><![CDATA[NM: ${h.nano_market || ""}<br>MM: ${h.micro_market || ""}<br>ID: ${h.hood_id || ""}]]></description>
    <Style><PolyStyle><color>${color}</color><outline>1</outline></PolyStyle></Style>
    ${geometryToKmlGeometry(h.geometry)}
  </Placemark>`).join("\n");
  return `<?xml version="1.0" encoding="UTF-8"?>\n<kml xmlns="http://www.opengis.net/kml/2.2"><Document><name>${escXml(layerName)}</name>\n${placemarks}\n</Document></kml>`;
}

function pointsToKml(data, layerName) {
  const placemarks = data
    .filter(row => !isNaN(parseFloat(row.Lat)) && !isNaN(parseFloat(row.Long)))
    .map(row => `
  <Placemark>
    <name>${escXml(getPropertyName(row))}</name>
    <description><![CDATA[Category: ${row.Category || ""}<br>NM: ${row.NM || ""}<br>MM: ${row.MM || ""}<br>Road: ${row.Road || ""}<br>Status: ${row["Final Status"] || ""}]]></description>
    <Point><coordinates>${parseFloat(row.Long)},${parseFloat(row.Lat)},0</coordinates></Point>
  </Placemark>`).join("\n");
  return `<?xml version="1.0" encoding="UTF-8"?>\n<kml xmlns="http://www.opengis.net/kml/2.2"><Document><name>${escXml(layerName)}</name>\n${placemarks}\n</Document></kml>`;
}

function geometryToWkt(geometry) {
  if (!geometry) return "";
  if (geometry.type === "Polygon") {
    const ring = geometry.coordinates[0].map(c => `${c[0]} ${c[1]}`).join(", ");
    return `POLYGON((${ring}))`;
  }
  if (geometry.type === "MultiPolygon") {
    return `MULTIPOLYGON(${geometry.coordinates.map(poly => `((${poly[0].map(c=>`${c[0]} ${c[1]}`).join(", ")}))`).join(", ")})`;
  }
  return "";
}

function hoodsToCsvWkt(hoodList, nameField) {
  const rows = hoodList.map(h => [
    `"${geometryToWkt(h.geometry)}"`,
    `"${h[nameField] || h.nano_market || h.micro_market || ""}"`,
    `"${h.nano_market  || ""}"`,
    `"${h.micro_market || ""}"`,
    `"${h.hood_id      || ""}"`
  ].join(","));
  return ["WKT,name,nm,mm,hood_id", ...rows].join("\n");
}

function pointsToCsvWkt(data) {
  const rows = data
    .filter(row => !isNaN(parseFloat(row.Lat)) && !isNaN(parseFloat(row.Long)))
    .map(row => {
      const lat = parseFloat(row.Lat), lng = parseFloat(row.Long);
      return [
        `"POINT(${lng} ${lat})"`,
        `"${getPropertyName(row).replace(/"/g,'""')}"`,
        `"${row.Category        || ""}"`,
        `"${row.NM              || ""}"`,
        `"${row.MM              || ""}"`,
        `"${(row.Road || "").replace(/"/g,'""')}"`,
        `"${row["Final Status"] || ""}"`,
        lat, lng
      ].join(",");
    });
  return ["WKT,name,category,nm,mm,road,final_status,lat,long", ...rows].join("\n");
}

function downloadBlob(content, filename, mimeType) {
  const blob = new Blob([content], { type: mimeType });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement("a");
  a.href = url; a.download = filename; a.click();
  URL.revokeObjectURL(url);
}

function downloadLayerKML(type) {
  const filteredData = getFilteredData();
  const label = activeFilters.NM || activeFilters.MM || "filtered";
  if (type === "nm") {
    const hoodList = hoodsByNM(getFilteredNMs());
    if (!hoodList.length) { alert("No NM hoods found."); return; }
    downloadBlob(hoodsToKml(hoodList, `NM Layer — ${label}`, "7f0000ff"), `nm_layer_${label}.kml`, "application/vnd.google-earth.kml+xml");
  } else if (type === "mm") {
    const seen = new Set();
    const hoodList = hoods.filter(h => {
      if (!getFilteredMMs().includes(h.micro_market) || seen.has(h.micro_market)) return false;
      seen.add(h.micro_market); return true;
    });
    if (!hoodList.length) { alert("No MM hoods found."); return; }
    downloadBlob(hoodsToKml(hoodList, `MM Layer — ${label}`, "7fff0000"), `mm_layer_${label}.kml`, "application/vnd.google-earth.kml+xml");
  } else if (type === "points") {
    if (!filteredData.length) { alert("No data points."); return; }
    downloadBlob(pointsToKml(filteredData, `Data Points — ${label}`), `points_${label}.kml`, "application/vnd.google-earth.kml+xml");
  }
}

function downloadLayerCSV(type) {
  const filteredData = getFilteredData();
  const label = activeFilters.NM || activeFilters.MM || "filtered";
  if (type === "nm") {
    const hoodList = hoodsByNM(getFilteredNMs());
    if (!hoodList.length) { alert("No NM hoods found."); return; }
    downloadBlob(hoodsToCsvWkt(hoodList, "nano_market"), `nm_layer_${label}.csv`, "text/csv");
  } else if (type === "mm") {
    const seen = new Set();
    const hoodList = hoods.filter(h => {
      if (!getFilteredMMs().includes(h.micro_market) || seen.has(h.micro_market)) return false;
      seen.add(h.micro_market); return true;
    });
    if (!hoodList.length) { alert("No MM hoods found."); return; }
    downloadBlob(hoodsToCsvWkt(hoodList, "micro_market"), `mm_layer_${label}.csv`, "text/csv");
  } else if (type === "points") {
    if (!filteredData.length) { alert("No data points."); return; }
    downloadBlob(pointsToCsvWkt(filteredData), `points_${label}.csv`, "text/csv");
  }
}

// ── Extra layer downloads ────────────────────────────────────
function extraLayerToCsvWkt(rows, latKey, lngKey, extraCols) {
  const headers = ["WKT", "nm", "mm", ...extraCols];
  const csvRows = rows
    .filter(r => !isNaN(parseFloat(r[latKey])) && !isNaN(parseFloat(r[lngKey])))
    .map(r => {
      const lat = parseFloat(r[latKey]), lng = parseFloat(r[lngKey]);
      return [
        `"POINT(${lng} ${lat})"`,
        `"${r._nm || ""}"`,
        `"${r._mm || ""}"`,
        ...extraCols.map(c => `"${String(r[c] || "").replace(/"/g, '""')}"`)
      ].join(",");
    });
  return [headers.join(","), ...csvRows].join("\n");
}

function extraLayerToKml(rows, latKey, lngKey, layerName, popupFn) {
  const placemarks = rows
    .filter(r => !isNaN(parseFloat(r[latKey])) && !isNaN(parseFloat(r[lngKey])))
    .map(r => `
  <Placemark>
    <n>${escXml(r.name || r.hood_name || "Point")}</n>
    <description><![CDATA[${popupFn(r)}]]></description>
    <Point><coordinates>${parseFloat(r[lngKey])},${parseFloat(r[latKey])},0</coordinates></Point>
  </Placemark>`).join("\n");
  return `<?xml version="1.0" encoding="UTF-8"?>\n<kml xmlns="http://www.opengis.net/kml/2.2"><Document><n>${escXml(layerName)}</n>\n${placemarks}\n</Document></kml>`;
}

function downloadExtraLayer(layerType, format) {
  const label = activeFilters.NM || activeFilters.MM || "all";

  if (layerType === "hotspots") {
    const data = getFilteredHotspots();
    if (!data.length) { alert("No hotspot data for current filter."); return; }
    if (format === "csv") downloadBlob(extraLayerToCsvWkt(data, "lat", "lng", ["name", "hood", "cluster"]), `hotspots_${label}.csv`, "text/csv");
    else downloadBlob(extraLayerToKml(data, "lat", "lng", `Hotspots — ${label}`, r => `Hood: ${r.hood || "-"}<br>Cluster: ${r.cluster || "-"}<br>NM: ${r._nm || "-"}<br>MM: ${r._mm || "-"}`), `hotspots_${label}.kml`, "application/vnd.google-earth.kml+xml");

  } else if (layerType === "demand") {
    const data = getFilteredDemand();
    if (!data.length) { alert("No demand data for current filter."); return; }
    if (format === "csv") downloadBlob(extraLayerToCsvWkt(data, "lat", "lng", ["cluster", "orders"]), `demand_${label}.csv`, "text/csv");
    else downloadBlob(extraLayerToKml(data, "lat", "lng", `Demand — ${label}`, r => `Cluster: ${r.cluster || "-"}<br>Orders: ${r.orders || "-"}<br>NM: ${r._nm || "-"}<br>MM: ${r._mm || "-"}`), `demand_${label}.kml`, "application/vnd.google-earth.kml+xml");

  } else if (layerType === "idle") {
    const data = getFilteredIdle();
    if (!data.length) { alert("No idle data for current filter."); return; }
    if (format === "csv") downloadBlob(extraLayerToCsvWkt(data, "lat", "lng", ["cluster", "hood", "idle_min", "w", "hood_pings"]), `idle_${label}.csv`, "text/csv");
    else downloadBlob(extraLayerToKml(data, "lat", "lng", `Idle — ${label}`, r => `Cluster: ${r.cluster || "-"}<br>Hood: ${r.hood || "-"}<br>Idle min: ${r.idle_min || "-"}<br>NM: ${r._nm || "-"}<br>MM: ${r._mm || "-"}`), `idle_${label}.kml`, "application/vnd.google-earth.kml+xml");

  } else if (layerType === "centroids") {
    const data = getFilteredCentroids();
    if (!data.length) { alert("No centroid data for current filter."); return; }
    if (format === "csv") downloadBlob(extraLayerToCsvWkt(data, "centroid_lat", "centroid_lng", ["hood_name", "cluster_id"]), `centroids_${label}.csv`, "text/csv");
    else downloadBlob(extraLayerToKml(data, "centroid_lat", "centroid_lng", `Centroids — ${label}`, r => `Hood: ${r.hood_name || "-"}<br>Cluster ID: ${r.cluster_id || "-"}<br>NM: ${r._nm || "-"}<br>MM: ${r._mm || "-"}`), `centroids_${label}.kml`, "application/vnd.google-earth.kml+xml");
  }
}

// ============================================================
// URL / COORD RESOLUTION
// ============================================================
const LOCATION_COL = "Location (Google Maps URL) / Map Code";

async function resolveCoords(input) {
  if (!input || !input.toString().trim()) throw new Error("Empty input");
  const raw = input.toString().trim();
  console.log(`🔍 resolveCoords() — raw input: "${raw}"`);

  const directMatch = raw.match(/@(-?\d+\.\d+),(-?\d+\.\d+)/);
  if (directMatch) {
    return { lat: parseFloat(directMatch[1]), lng: parseFloat(directMatch[2]) };
  }

  const backendUrl = CONFIG.API_URL + "?action=resolveUrl&url=" + encodeURIComponent(raw);
  const res  = await fetch(backendUrl);
  if (!res.ok) throw new Error(`Backend resolve failed: ${res.status}`);

  const text = await res.text();
  const json = JSON.parse(text);

  if (json.error) throw new Error(`Backend error: ${json.error}`);
  if (json.lat != null && json.lng != null) {
    return { lat: parseFloat(json.lat), lng: parseFloat(json.lng) };
  }

  const expandedUrl = json.url || "";
  const d3Match = expandedUrl.match(/!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)/);
  if (d3Match) return { lat: parseFloat(d3Match[1]), lng: parseFloat(d3Match[2]) };
  const atMatch = expandedUrl.match(/@(-?\d+\.\d+),(-?\d+\.\d+)/);
  if (atMatch) return { lat: parseFloat(atMatch[1]), lng: parseFloat(atMatch[2]) };

  throw new Error(`Could not extract coords from: "${expandedUrl || raw}"`);
}

async function fillLatLong() {
  let updated = [], skippedCount = 0, failedCount = 0;
  console.log(`🌍 fillLatLong() — starting, total rows: ${allData.length}`);

  for (let row of allData) {
    const locationInput = row[LOCATION_COL];
    if (!locationInput || !locationInput.toString().trim()) { skippedCount++; continue; }
    const latOk = row.Lat  && !isNaN(parseFloat(row.Lat))  && parseFloat(row.Lat)  !== 0;
    const lngOk = row.Long && !isNaN(parseFloat(row.Long)) && parseFloat(row.Long) !== 0;
    if (latOk && lngOk) { skippedCount++; continue; }
    try {
      const coords = await resolveCoords(locationInput);
      row.Lat = coords.lat; row.Long = coords.lng;
      updated.push(row);
    } catch (e) {
      failedCount++;
      console.error(`❌ _rowIndex:${row._rowIndex} — failed for "${locationInput}": ${e.message}`);
    }
  }

  console.log(`🌍 fillLatLong() — done. Updated: ${updated.length}, Skipped: ${skippedCount}, Failed: ${failedCount}`);
  // updated.forEach(r => fetch(CONFIG.API_URL, { method: "POST", body: JSON.stringify(r) }));
  updated.forEach(r => {
  const payload = {
    _rowIndex: r._rowIndex,
    Lat: r.Lat,
    Long: r.Long
  };

  fetch(CONFIG.API_URL, {
      method: "POST",
      body: JSON.stringify(payload)
    });
  });
  alert(`✅ Updated ${updated.length} rows\n⏭️ Skipped: ${skippedCount}\n❌ Failed: ${failedCount}`);
  if (updated.length > 0) loadData();
}

// ============================================================
// FIX MISSING NM
// ============================================================
function fixMissingNM() {
  let updated = [];
  allData.forEach(row => {
    const lat = parseFloat(row.Lat), lng = parseFloat(row.Long);
    if (isNaN(lat) || isNaN(lng)) return;
    if (isEmpty(row.NM) || isEmpty(row.MM)) {
      const hood = assignHood({ lat, lng });
      if (hood) {
        row.NM = hood.nano_market;
        row.MM = hood.micro_market;
        row["NM Id"] = hood.hood_id;
        updated.push(row);
      }
    }
  });
  updated.forEach(r => {
    const payload = {
      _rowIndex: r._rowIndex,
      NM:        r.NM,
      MM:        r.MM,
      "NM Id":   r["NM Id"]
    };
    fetch(CONFIG.API_URL, { method: "POST", body: JSON.stringify(payload) });
  });
  alert(`✅ Updated ${updated.length} rows`);
}

// ============================================================
// HOOD UTILITIES
// ============================================================
function drawHoods() {
  hoodLayers = [];
  hoods.forEach(h => {
    if (!h.geometry) return;
    const layer = L.geoJSON(h.geometry, {
      style: { color: "blue", weight: 1, fillColor: "#4da6ff", fillOpacity: 0.15 }
    }).addTo(map);
    layer.on("click", () => {
      layer.setStyle({ fillColor: "orange", fillOpacity: 0.4 });
      showHoodDetails(h);
    });
    hoodLayers.push({ layer, nm: h.nano_market, mm: h.micro_market });
  });
}

function updateHoodVisibility() {
  const filterNM = activeFilters.NM || "";
  const filterMM = activeFilters.MM || "";
  const noFilter = !filterNM && !filterMM;
  hoodLayers.forEach(({ layer, nm, mm }) => {
    const visible = layerVisible.hoods && (noFilter || ((!filterNM || nm === filterNM) && (!filterMM || mm === filterMM)));
    visible ? (map.hasLayer(layer) || map.addLayer(layer)) : (map.hasLayer(layer) && map.removeLayer(layer));
    if (visible) layer.setStyle({ color: "blue", weight: 1, fillColor: "#4da6ff", fillOpacity: 0.15 });
  });
}

function assignHood(coords) {
  const pt = turf.point([coords.lng, coords.lat]);
  let nearest = null, minDist = Infinity;
  for (let h of hoods) {
    const polygon = { type: "Feature", geometry: h.geometry };
    try {
      if (turf.booleanPointInPolygon(pt, polygon)) return h;
      const dist = turf.distance(pt, turf.centroid(polygon));
      if (dist < minDist) { minDist = dist; nearest = h; }
    } catch (e) {}
  }
  return nearest;
}

function isEmpty(val) { return !val || val.toString().trim() === "" || val === "NA"; }

function escHtml(str) {
  return String(str || "").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
}

function getCategoryIcon(category) {
  return L.divIcon({
    className: "custom-icon",
    html: `<div style="font-size:18px">📍</div>`
  });
}

// ============================================================
// DETAILS + SAVE
// ============================================================
function showHoodDetails(h) {
  document.getElementById("detailsTable").innerHTML = `
    <tr><td><b>NM</b></td><td>${h.nano_market}</td></tr>
    <tr><td><b>MM</b></td><td>${h.micro_market}</td></tr>
    <tr><td><b>Region</b></td><td>${h.region}</td></tr>
    <tr><td><b>Hood ID</b></td><td>${h.hood_id}</td></tr>
  `;
}

const DETAIL_DROPDOWNS = {
  "Property": [
    "Private", "Public"
  ],
  "App status": [
    "Active", "Inactive"
  ],
  "Lead Status": [
    "2. Owner conversation pending",
    "3. Owner's confirmation pending",
    "4. Confirmed",
    "5. Follow up required",
    "6. Dropped"
  ],
  "Final Status": [
    "Dropped off", "Active", "Cold", "Deal closed",
    "Deal closed - sign pending", "No deal required",
    "To be reactivated", "Deal - closed - Chairs pending",
    "Dropped off after launch"
  ],
  "Closure type": [
    "Resting + Washroom", "Resting", "NA"
  ],
  "Set up": [
    "Chairs to be set", "Owner will setup chairs", "Chairs available", "NA"
  ],
  "Category": [
    "Ladies PG", "Shop", "Restaurant", "Apartment", "Gated community", "Independent Builder floor",
    "Bus Stop", "Park", "Petrol Pump", "Public Washroom", "Other"
  ]
};

function showDetails(row) {
  currentRow = row;
  const table = document.getElementById("detailsTable");
  table.innerHTML = "";

  Object.keys(row).forEach(key => {
    if (key === "_rowIndex") return;
    const val = row[key] != null ? row[key] : "";

    if (DETAIL_DROPDOWNS[key]) {
      const options = DETAIL_DROPDOWNS[key].map(opt =>
        `<option value="${escHtml(opt)}" ${opt === String(val) ? "selected" : ""}>${escHtml(opt)}</option>`
      ).join("");
      const extraOption = val && !DETAIL_DROPDOWNS[key].includes(String(val))
        ? `<option value="${escHtml(val)}" selected>${escHtml(val)}</option>`
        : "";
      table.innerHTML += `
        <tr>
          <td>${escHtml(key)}</td>
          <td>
            <select class="detail-select" data-key="${escHtml(key)}">
              <option value="">-- select --</option>
              ${extraOption}${options}
            </select>
          </td>
        </tr>`;
    } else if (key === "Timestamp") {
      // Timestamp is the original Google Form submission time — never editable
      const displayVal = formatTsDisplay(val) || escHtml(String(val));
      table.innerHTML += `
        <tr>
          <td>${escHtml(key)}</td>
          <td style="background:#f5f5f5;color:#888;font-style:italic" title="Original form submission time — read only">${displayVal}</td>
        </tr>`;
    } else {
      table.innerHTML += `
        <tr>
          <td>${escHtml(key)}</td>
          <td contenteditable="true" data-key="${escHtml(key)}">${escHtml(String(val))}</td>
        </tr>`;
    }
  });

  table.innerHTML += `
    <tr>
      <td colspan="2" style="padding-top:10px;border-top:1px solid #eee">
        <button class="primary" onclick="saveCurrent()" style="width:100%;padding:9px;font-size:14px">
          💾 Save Changes
        </button>
      </td>
    </tr>`;
}

// Returns timestamp string in "M/D/YYYY HH:MM:SS" format, always in IST (UTC+5:30).
// Uses explicit offset arithmetic so it works correctly regardless of the browser's
// local timezone — prevents UTC timestamps appearing when the user's system is UTC.
function formatTimestamp(date) {
  // Shift to IST wall-clock time
  const ist = new Date(date.getTime() + 5.5 * 60 * 60 * 1000);
  // Use UTC getters on the shifted date to read IST values
  return `${ist.getUTCMonth()+1}/${ist.getUTCDate()}/${ist.getUTCFullYear()} ` +
    `${String(ist.getUTCHours()).padStart(2,'0')}:${String(ist.getUTCMinutes()).padStart(2,'0')}:${String(ist.getUTCSeconds()).padStart(2,'0')}`;
}

function saveCurrent() {
  if (!currentRow) return;

  // ── Snapshot old App status BEFORE overwriting with new values ───────────
  const prevAppStatus = String(currentRow["App status"] || "").trim();
  const prevFinalStatus = String(currentRow["Final Status"] || "").trim();
  // ── Read all editable cells into currentRow ──────────────────────────────
  document.querySelectorAll("[contenteditable]").forEach(cell => {
    const key = cell.dataset.key;
    if (key) currentRow[key] = cell.innerText.trim();
  });

  document.querySelectorAll(".detail-select").forEach(sel => {
    const key = sel.dataset.key;
    if (key) currentRow[key] = sel.value;
  });

  if (!currentRow._rowIndex) { alert("\u274c Cannot save \u2014 row index missing, try refreshing"); return; }

  const nowIST = formatTimestamp(new Date());

  // // ── Rule 1: Signage date ─────────────────────────────────────────────────
  // // Fill ONLY when: Signage date is currently empty AND Lead Status is a deal-closed state
  // const SIGNAGE_STATUSES = ["Deal closed", "Deal - closed - Chairs pending"];
  // const leadStatus   = String(currentRow["Lead Status"] || "").trim();
  // const signageEmpty = !currentRow["Signage date"] || String(currentRow["Signage date"]).trim() === "";
  // if (signageEmpty && SIGNAGE_STATUSES.includes(leadStatus)) {
  //   currentRow["Signage date"] = nowIST;
  //   console.log("\ud83d\udcc5 Signage date auto-filled:", nowIST, "(Lead Status:", leadStatus + ")");
  // }

  // ── Rule 1: Launch date ──────────────────────────────────────────────────
  // Fill when App status is being changed TO "Active" (was not "Active" before)
  const newAppStatus = String(currentRow["App status"] || "").trim();
  if (newAppStatus === "Active" && prevAppStatus !== "Active") {
    currentRow["Launch date"] = nowIST;
    console.log("\ud83d\ude80 Launch date auto-filled:", nowIST, "(App status changed to Active)");
  }

  // ── Rule 2: Signage date ──────────────────────────────────────────────────
  // Fill when Final Status changes to any "closed" state

  const CLOSED_STATUSES = [
    "Deal closed",
    "Deal - closed - Chairs pending"
  ];

  const newFinalStatus = String(currentRow["Final Status"] || "").trim();

  if (
    CLOSED_STATUSES.includes(newFinalStatus) &&
    !CLOSED_STATUSES.includes(prevFinalStatus)
  ) {
    currentRow["Signage date"] = nowIST;
    console.log("📅 Signage date auto-filled:", nowIST, "(Final Status changed to closed state)");
  }

  // ── Timestamp column is NEVER written back ───────────────────────────────
  // It reflects the original Google Form submission time.
  // Build a clean payload that explicitly excludes Timestamp.
  const savePayload = Object.assign({}, currentRow);
  delete savePayload["Timestamp"];

  fetch(CONFIG.API_URL, { method: "POST", body: JSON.stringify(savePayload) })
    .then(() => {
      alert("\u2705 Saved");
      showDetails(currentRow);
    })
    .catch(err => alert("\u274c Save failed: " + err.message));
}

// ============================================================
// TIMESTAMP PARSING — Sheet stores IST directly as M/D/YYYY H:MM:SS
// No UTC shifting needed. Parse as local time only.
// ============================================================
function parseTimestamp(val) {
  if (!val || val === "") return null;

  // Already a Date object
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val;

  // Google Sheets serial number (days since 1899-12-30)
  // Rare — only if cell is formatted as Number instead of Date/Text
  if (typeof val === "number") {
    const d = new Date((val - 25569) * 86400000);
    return isNaN(d.getTime()) ? null : d;
  }

  const str = val.toString().trim();

  // ── ISO 8601 / UTC strings ──
  // e.g. "2026-04-14T12:26:08.000Z" — treat as UTC, display in IST
  if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/.test(str)) {
    const d = new Date(str);
    return isNaN(d.getTime()) ? null : d;
  }

  // ── PRIMARY FORMAT ──
  // "M/D/YYYY H:MM:SS" written by formatTimestamp() — always IST, parse as local
  const slashMatch = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2}):(\d{2})$/);
  if (slashMatch) {
    const [, m, d, y, hr, min, sec] = slashMatch;
    return new Date(+y, +m - 1, +d, +hr, +min, +sec);
  }

  // ── FALLBACK ──
  // Plain date string without time e.g. "2026-04-09"
  const fallback = new Date(str);
  return isNaN(fallback.getTime()) ? null : fallback;
}

// Format a timestamp value for display — always shows IST time
function formatTsDisplay(val) {
  if (!val || val === "") return "";
  const d = parseTimestamp(val);
  if (!d) return String(val); // unparseable — show raw

  // For the primary M/D/YYYY format (already IST), use local getters directly
  const str = val.toString().trim();
  const isSlashFormat = /^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2}):(\d{2})$/.test(str);
  if (isSlashFormat) {
    return `${d.getDate()}/${d.getMonth()+1}/${d.getFullYear()} ` +
      `${String(d.getHours()).padStart(2,'0')}:${String(d.getMinutes()).padStart(2,'0')}:${String(d.getSeconds()).padStart(2,'0')}`;
  }

  // For ISO/UTC strings, shift to IST (UTC+5:30) before displaying
  const ist = new Date(d.getTime() + 5.5 * 60 * 60 * 1000);
  return `${ist.getUTCDate()}/${ist.getUTCMonth()+1}/${ist.getUTCFullYear()} ` +
    `${String(ist.getUTCHours()).padStart(2,'0')}:${String(ist.getUTCMinutes()).padStart(2,'0')}:${String(ist.getUTCSeconds()).padStart(2,'0')}`;
}

// ============================================================
// SHEET PREVIEW — independent filters + totals row
// ============================================================
function getSheetFilters() {
  return {
    Category:       document.getElementById("sfCategory")?.value       || "",
    Property:       document.getElementById("sfProperty")?.value       || "",
    "App status":   document.getElementById("sfAppStatus")?.value      || "",
    "Lead Status":  document.getElementById("sfLeadStatus")?.value     || "",
    "Final Status": document.getElementById("sfFinalStatus")?.value    || "",
    NM:             document.getElementById("sfNM")?.value             || "",
    MM:             document.getElementById("sfMM")?.value             || "",
    dateFrom:       document.getElementById("sfDateFrom")?.value       || "",
    dateTo:         document.getElementById("sfDateTo")?.value         || "",
    signageFrom:    document.getElementById("sfSignageFrom")?.value    || "",
    signateTo:      document.getElementById("sfSignageTo")?.value      || "",
    launchFrom:     document.getElementById("sfLaunchFrom")?.value     || "",
    launchTo:       document.getElementById("sfLaunchTo")?.value       || "",
  };
}

function getSheetFilteredData() {
  const sf   = getSheetFilters();
  const from        = sf.dateFrom    ? new Date(sf.dateFrom    + "T00:00:00") : null;
  const to          = sf.dateTo      ? new Date(sf.dateTo      + "T23:59:59") : null;
  const signageFrom = sf.signageFrom ? new Date(sf.signageFrom + "T00:00:00") : null;
  const signateTo   = sf.signateTo   ? new Date(sf.signateTo   + "T23:59:59") : null;
  const launchFrom  = sf.launchFrom  ? new Date(sf.launchFrom  + "T00:00:00") : null;
  const launchTo    = sf.launchTo    ? new Date(sf.launchTo    + "T23:59:59") : null;

  return allData.filter(row => {
    const colKeys = ["Category", "Property", "App status", "Lead Status", "Final Status", "NM", "MM"];
    for (const key of colKeys) {
      if (sf[key] && row[key] !== sf[key]) return false;
    }
    // Timestamp range
    if (from || to) {
      const ts = parseTimestamp(row["Timestamp"]);
      if (!ts) return false;
      if (from && ts < from) return false;
      if (to   && ts > to)   return false;
    }
    // Signage date range
    if (signageFrom || signateTo) {
      const ts = parseTimestamp(row["Signage date"]);
      if (!ts) return false;
      if (signageFrom && ts < signageFrom) return false;
      if (signateTo   && ts > signateTo)   return false;
    }
    // Launch date range
    if (launchFrom || launchTo) {
      const ts = parseTimestamp(row["Launch date"]);
      if (!ts) return false;
      if (launchFrom && ts < launchFrom) return false;
      if (launchTo   && ts > launchTo)   return false;
    }
    return true;
  });
}

function applySheetFilters() {
  renderSheetPreview(getSheetFilteredData());
}

function clearSheetFilters() {
  ["sfCategory","sfProperty","sfAppStatus","sfLeadStatus","sfFinalStatus","sfNM","sfMM"].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = "";
  });
  ["sfDateFrom","sfDateTo","sfSignageFrom","sfSignageTo","sfLaunchFrom","sfLaunchTo"].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = "";
  });
  renderSheetPreview(allData);
}

// Backward-compat aliases
function applyDateFilter() { applySheetFilters(); }
function clearDateFilter()  { clearSheetFilters(); }

function populateSheetFilters() {
  const fields = [
    { key: "Category",      id: "sfCategory"    },
    { key: "Property",      id: "sfProperty"    },
    { key: "App status",    id: "sfAppStatus"   },
    { key: "Lead Status",   id: "sfLeadStatus"  },
    { key: "Final Status",  id: "sfFinalStatus" },
    { key: "NM",            id: "sfNM"          },
    { key: "MM",            id: "sfMM"          },
  ];
  fields.forEach(f => {
    const el = document.getElementById(f.id);
    if (!el) return;
    const current = el.value;
    const vals    = [...new Set(allData.map(r => r[f.key]).filter(Boolean))].sort();
    el.innerHTML  = `<option value="">${f.key}</option>` +
      vals.map(v => `<option value="${v}">${escHtml(v)}</option>`).join("");
    el.value = current;
  });
}

function renderSheetPreview(data) {
  const table   = document.getElementById("sheetPreviewTable");
  const countEl = document.getElementById("sheetRowCount");
  if (!table) return;

  populateSheetFilters();

  if (!data.length) {
    table.innerHTML = "<tr><td colspan='99' style='text-align:center;color:#aaa;padding:16px'>No rows match filters</td></tr>";
    if (countEl) countEl.textContent = "0 rows";
    return;
  }
  if (countEl) countEl.textContent = `${data.length} rows`;

  // Columns shown in the sheet preview table (in display order)
  const previewCols = [
    "MM", "NM", "NM Id", "Name of the property", "Name", "Category",
    "Closure type", "Lat", "Long", "Location (Google Maps URL) / Map Code",
    "Owner Contact Name", "Owner Contact Number",
    "Contact Name", "Contact number",
    "Property", "App status", "Lead Status", "Final Status",
    "Signage date", "Launch date"
  ];
  const availableCols = previewCols.filter(c => data[0].hasOwnProperty(c));

  const bodyRows = data.map((row, idx) => `
    <tr data-idx="${idx}" style="cursor:pointer">
      ${availableCols.map(c => {
        const raw     = row[c] != null ? row[c] : "";
        const TS_COLS = ["Timestamp", "Signage date", "Launch date"];
        const display = TS_COLS.includes(c) ? formatTsDisplay(row[c]) : escHtml(raw);
        return `<td title="${escHtml(String(raw))}">${display}</td>`;
      }).join("")}
    </tr>`).join("");

  const totalsRow = `<tr class="summary-total-row" style="position:sticky;bottom:0">
    ${availableCols.map((c, i) => {
      if (i === 0) return `<td><b>Total: ${data.length}</b></td>`;
      const nonEmpty = data.filter(r => r[c] != null && r[c] !== "").length;
      return `<td><b>${nonEmpty}</b></td>`;
    }).join("")}
  </tr>`;

  table.innerHTML =
    `<thead><tr>${availableCols.map(c => `<th>${escHtml(c)}</th>`).join("")}</tr></thead>` +
    `<tbody>${bodyRows}</tbody>` +
    `<tfoot>${totalsRow}</tfoot>`;

  table.querySelectorAll("tbody tr").forEach(tr => {
    tr.addEventListener("click", () => {
      const idx = parseInt(tr.dataset.idx);
      showDetails(data[idx]);
      table.querySelectorAll("tbody tr").forEach(r => r.classList.remove("active-row"));
      tr.classList.add("active-row");
      document.querySelector(".details-card").scrollIntoView({ behavior: "smooth", block: "start" });
    });
  });
}

function downloadSheetPreviewCSV() {
  const table = document.getElementById("sheetPreviewTable");
  if (!table) return;
  const rows = [...table.querySelectorAll("thead tr, tbody tr, tfoot tr")].map(tr =>
    [...tr.querySelectorAll("th,td")].map(td => `"${td.innerText.replace(/"/g,'""')}"`).join(",")
  );
  const dateTag = new Date().toISOString().slice(0,10);
  downloadBlob(rows.join("\n"), `sheet_preview_${dateTag}.csv`, "text/csv");
}

// ============================================================
// CROSS-TAB SUMMARY TABLE
// ============================================================
function getSummaryFilters() {
  return {
    NM:       document.getElementById("stNM")?.value       || "",
    MM:       document.getElementById("stMM")?.value       || "",
    dateFrom: document.getElementById("stDateFrom")?.value || "",
    dateTo:   document.getElementById("stDateTo")?.value   || "",
  };
}

function getSummaryFilteredData() {
  const sf   = getSummaryFilters();
  const from = sf.dateFrom ? new Date(sf.dateFrom + "T00:00:00") : null;
  const to   = sf.dateTo   ? new Date(sf.dateTo   + "T23:59:59") : null;

  return allData.filter(row => {
    if (sf.NM && row.NM !== sf.NM) return false;
    if (sf.MM && row.MM !== sf.MM) return false;
    if (from || to) {
      const ts = parseTimestamp(row["Timestamp"]);
      if (!ts) return false;
      if (from && ts < from) return false;
      if (to   && ts > to)   return false;
    }
    return true;
  });
}

function populateSummaryFilters() {
  [{ key: "NM", id: "stNM" }, { key: "MM", id: "stMM" }].forEach(f => {
    const el = document.getElementById(f.id);
    if (!el) return;
    const current = el.value;
    const vals    = [...new Set(allData.map(r => r[f.key]).filter(Boolean))].sort();
    el.innerHTML  = `<option value="">${f.key}</option>` +
      vals.map(v => `<option value="${v}">${escHtml(v)}</option>`).join("");
    el.value = current;
  });
}

function applySummaryFilters() { renderSummaryTables(); }

function clearSummaryFilters() {
  ["stNM","stMM","stDateFrom","stDateTo"].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = "";
  });
  renderSummaryTables();
}

function renderSummaryTables() {
  const container = document.getElementById("summaryTablesContainer");
  if (!container) return;

  populateSummaryFilters();

  const data = getSummaryFilteredData();
  if (!data.length) {
    container.innerHTML = `<p style="color:#aaa;padding:12px">No data for selected filters.</p>`;
    return;
  }

  const FIXED_CATEGORIES = [
    "Ladies PG", "Shop", "Restaurant", "Gated community",
    "Independent Builder floor", "Bus Stop", "Park",
    "Petrol Pump", "Public Washroom", "Other"
  ];

  const FIXED_FINAL_STATUSES = [
    "Total in funnel", "To be reactivated", "Cold", "Dropped off",
    "Places Finalised", "Deal closed - sign pending",
    "Deal - closed - Chairs pending", "Deal closed"
  ];

  const FIXED_COMMERCIAL_BUCKETS = [
    "2000", "2500", "3000", "3500", "4000", "others", "NA"
  ];

  const allCats     = FIXED_CATEGORIES;
  const allStatuses = FIXED_FINAL_STATUSES;

  const normalizeCategory = (cat) => {
    if (cat === "Apartment") return "Independent Builder floor";
    return cat;
  };

  const normalizeCommercial = (val) => {
    if (!val) return "NA";
    val = String(val).trim();
    if (["2000","2500","3000","3500","4000"].includes(val)) return val;
    if (val.toLowerCase() === "na") return "NA";
    return "others";
  };

  const countByCat = (rows) => {
    const c = {};
    allCats.forEach(cat => {
      c[cat] = rows.filter(r => normalizeCategory(r.Category) === cat).length || 0;
    });
    return c;
  };

  const headerCols = `<th>Final Status</th><th>Total</th>${
    allCats.map(c => `<th>${escHtml(c)}</th>`).join("")
  }`;

  const statusRows = allStatuses.map(status => {
    let rows;

    if (status === "Total in funnel") {
      rows = data;
    } else if (status === "Places Finalised") {
      const finalisedRows = data.filter(r =>
        ["Deal closed", "Deal closed - sign pending", "Deal - closed - Chairs pending"].includes(r["Final Status"])
      );
      const totalRows = data.length;
      const pct = (num, denom) => !denom ? "0 (0%)" : `${num} (${(num / denom * 100).toFixed(1)}%)`;
      const cats = countByCat(finalisedRows);
      return `<tr>
        <td>${escHtml(status)}</td>
        <td><b>${pct(finalisedRows.length, totalRows)}</b></td>
        ${allCats.map(c => {
          const catTotal = data.filter(r => normalizeCategory(r.Category) === c).length;
          return `<td>${pct(cats[c], catTotal)}</td>`;
        }).join("")}
      </tr>`;
    } else {
      rows = data.filter(r => r["Final Status"] === status);
    }

    const cats = countByCat(rows);
    return `<tr>
      <td>${escHtml(status)}</td>
      <td><b>${rows.length}</b></td>
      ${allCats.map(c => `<td>${cats[c]}</td>`).join("")}
    </tr>`;
  }).join("");

  const section1 = `
    <div class="summary-block">
      <div class="summary-block-header">
        <h4 class="summary-block-title">Final Status × Category</h4>
        <button class="summary-dl-btn" onclick="downloadSummaryTable('stMainTable','final_status_x_category')">⬇ CSV</button>
      </div>
      <div class="summary-table-wrapper">
        <table class="summary-table" id="stMainTable">
          <thead><tr>${headerCols}</tr></thead>
          <tbody>${statusRows}</tbody>
        </table>
      </div>
    </div>`;

  const closureTypes    = ["Resting + Washroom", "Resting"];
  const commercialRows  = closureTypes.map(closureType => {
    const closureRows = data.filter(r => r["Closure type"] === closureType);

    const headerRow = `<tr class="summary-closure-header">
      <td colspan="${2 + allCats.length}">
        <b>${escHtml(closureType)}</b>
        <span style="color:#888;margin-left:8px">(${closureRows.length} total)</span>
      </td>
    </tr>`;

    const valueRows = FIXED_COMMERCIAL_BUCKETS.map(val => {
      const valRows = closureRows.filter(r => {
        const raw = r["Closure commercial (Ex. 2000, 4000 etc)"] || r["Closure commercial"];
        return normalizeCommercial(raw) === val;
      });
      const cats = countByCat(valRows);
      return `<tr>
        <td style="padding-left:16px">${val}</td>
        <td>${valRows.length}</td>
        ${allCats.map(c => `<td>${cats[c]}</td>`).join("")}
      </tr>`;
    }).join("");

    const subtotalCats = countByCat(closureRows);
    const subtotalRow  = `<tr style="background:#f5f5f5">
      <td><i>Subtotal</i></td>
      <td>${closureRows.length}</td>
      ${allCats.map(c => `<td>${subtotalCats[c]}</td>`).join("")}
    </tr>`;

    return headerRow + valueRows + subtotalRow;
  }).join("");

  const section2 = `
    <div class="summary-block" style="margin-top:24px">
      <div class="summary-block-header">
        <h4 class="summary-block-title">Commercials × Property Type</h4>
        <button class="summary-dl-btn" onclick="downloadSummaryTable('stCommTable','commercials_x_property_type')">⬇ CSV</button>
      </div>
      <div class="summary-table-wrapper">
        <table class="summary-table" id="stCommTable">
          <thead>
            <tr>${headerCols.replace("Final Status","Closure Type / Value")}</tr>
          </thead>
          <tbody>${commercialRows}</tbody>
        </table>
      </div>
    </div>`;

  container.innerHTML = section1 + section2;
}

function downloadSummaryTable(tableId, filename) {
  const table = document.getElementById(tableId);
  if (!table) { alert("Table not found"); return; }
  const rows = [...table.querySelectorAll("tr")].map(tr =>
    [...tr.querySelectorAll("th,td")].map(td => `"${td.innerText.replace(/"/g,'""')}"`).join(",")
  );
  const dateTag = new Date().toISOString().slice(0,10);
  downloadBlob(rows.join("\n"), `${filename}_${dateTag}.csv`, "text/csv");
}

function refreshData() { loadData(); loadExtraLayers(); }

// ============================================================
// MAP SEARCH
// ============================================================
let searchMarker = null;
let searchDebounceTimer = null;

function initMapSearch() {
  const input   = document.getElementById("mapSearchInput");
  const results = document.getElementById("mapSearchResults");
  input.addEventListener("input", () => {
    clearTimeout(searchDebounceTimer);
    const q = input.value.trim();
    if (q.length < 3) { results.classList.remove("open"); return; }
    searchDebounceTimer = setTimeout(() => searchLocation(q), 350);
  });
  input.addEventListener("keydown", e => {
    if (e.key === "Escape") { results.classList.remove("open"); input.blur(); }
  });
  document.addEventListener("click", e => {
    if (!e.target.closest(".map-search-wrapper")) results.classList.remove("open");
  });
}

async function searchLocation(query) {
  const results = document.getElementById("mapSearchResults");
  results.innerHTML = `<div class="search-result-item" style="color:#888">Searching...</div>`;
  results.classList.add("open");
  try {
    const url  = `https://nominatim.openstreetmap.org/search?q=${encodeURIComponent(query)}&format=json&limit=6&addressdetails=1`;
    const res  = await fetch(url, { headers: { "Accept-Language": "en" } });
    const data = await res.json();
    if (!data.length) { results.innerHTML = `<div class="search-result-item" style="color:#888">No results found</div>`; return; }
    results.innerHTML = data.map(item => {
      const name = item.name || item.display_name.split(",")[0];
      return `<div class="search-result-item" onclick="selectSearchResult(${item.lat}, ${item.lon}, '${name.replace(/'/g,"\\'")}')">
        <div class="result-name">${name}</div>
        <div class="result-addr">${item.display_name}</div>
      </div>`;
    }).join("");
  } catch (err) {
    results.innerHTML = `<div class="search-result-item" style="color:#c00">Search failed</div>`;
  }
}

function selectSearchResult(lat, lng, name) {
  if (searchMarker) { map.removeLayer(searchMarker); searchMarker = null; }
  const latlng = [parseFloat(lat), parseFloat(lng)];
  searchMarker = L.marker(latlng, {
    icon: L.divIcon({ className: "custom-icon", html: `<div style="font-size:28px">📌</div>` })
  }).addTo(map).bindPopup(`<b>${name}</b><br><small>${lat}, ${lng}</small>`).openPopup();
  map.setView(latlng, 16);
  document.getElementById("mapSearchInput").value = name;
  document.getElementById("mapSearchResults").classList.remove("open");
}