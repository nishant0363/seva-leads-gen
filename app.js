// ============================================================
// GLOBALS
// ============================================================
let map;
let currentRow = null;
let allData = [];
let hoods = [];

// MapLibre uses source/layer IDs instead of Leaflet marker arrays.
// We keep arrays of popup/marker objects for property pins (DOM markers).
let propertyMarkers   = [];   // maplibregl.Marker[]
let hotspotMarkers    = [];
let demandMarkers     = [];
let idleMarkers       = [];
let centroidMarkers   = [];

let activeFilters = {};

// Add-Point mode
let addPointMode = false;
let addPointMarker = null;  // maplibregl.Marker

// Extra layer data
let hotspotData   = [];
let demandData    = [];
let idleData      = [];
let centroidData  = [];

// Layer visibility (legend toggles)
const layerVisible = {
  hoods:      true,
  properties: true,
  hotspots:   true,
  demand:     false,
  idle:       false,
  centroids:  true
};

const MAP_STYLES = {
  street: "https://tiles.openfreemap.org/styles/liberty",
  light:  "https://basemaps.cartocdn.com/gl/positron-gl-style/style.json",
  dark:   "https://basemaps.cartocdn.com/gl/dark-matter-gl-style/style.json",
  basic:  "https://demotiles.maplibre.org/style.json"
};
// Search
let searchMarker = null;
let searchDebounceTimer = null;

// MapLibre source/layer IDs for heatmaps
const DEMAND_SOURCE = "demand-source";
const DEMAND_LAYER  = "demand-heat";
const IDLE_SOURCE   = "idle-source";
const IDLE_LAYER    = "idle-heat";
const HOODS_SOURCE  = "hoods-source";
const HOODS_FILL    = "hoods-fill";
const HOODS_LINE    = "hoods-line";

console.log("🚀 App initializing...");
init();

// ============================================================
// INIT
// ============================================================
async function init() {
  map = new maplibregl.Map({
    container: "map",
    style: "https://basemaps.cartocdn.com/gl/positron-gl-style/style.json",
    center: [77.65, 12.9],
    zoom: 12,
    attributionControl: true
  });

  // Wait for map to finish loading before adding sources/layers
  await new Promise(resolve => map.on("load", resolve));

  // Load hoods
  hoods = await fetch(CONFIG.API_URL + "?action=getHoods&t=" + Date.now(), {
    credentials: "omit"
  }).then(r => r.json()).catch(() => ([]));

  if (!Array.isArray(hoods) || hoods.error) {
    console.error("❌ Failed to load hoods from sheet, trying hoods.json");
    hoods = await fetch("hoods.json").then(r => r.json()).catch(() => []);
    console.log(`📦 hoods.json loaded — ${hoods.length} hoods`);
  } else {
    console.log(`📦 hoods loaded from sheet — ${hoods.length} hoods`);
  }

  drawHoods();
  buildLegend();

  await loadData();
  await loadExtraLayers();

  initMapSearch();

  // Map click handler
  map.on("click", function (e) {
    if (addPointMode) {
      openAddPointModal(e.lngLat.lat, e.lngLat.lng);
      return;
    }

    // Show a copyable lat/long popup on every normal click
    const lat = e.lngLat.lat, lng = e.lngLat.lng;
    const coordStr = `${lat.toFixed(7)}, ${lng.toFixed(7)}`;
    const popupEl = document.createElement("div");
    popupEl.style.cssText = "font-size:13px;line-height:1.6";
    popupEl.innerHTML = `
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
    `;
    new maplibregl.Popup({ closeButton: true })
      .setLngLat([lng, lat])
      .setDOMContent(popupEl)
      .addTo(map);
  });

  // Hood polygon clicks
  map.on("click", HOODS_FILL, function (e) {
    const hoodId = e.features[0]?.properties?.hood_id;
    const hood = hoods.find(h => h.hood_id === hoodId);
    if (hood) showHoodDetails(hood);
  });
  map.on("mouseenter", HOODS_FILL, () => map.getCanvas().style.cursor = "pointer");
  map.on("mouseleave", HOODS_FILL, () => map.getCanvas().style.cursor = "");
}

function switchBaseMap(styleKey, event) {
  document.querySelectorAll(".style-btn").forEach(b => b.classList.remove("active"));
  if (event) event.target.classList.add("active");

  const styleUrl = MAP_STYLES[styleKey];
  if (!styleUrl) return;

  map.setStyle(styleUrl);

  map.once("styledata", async () => {
    drawHoods();
    await loadExtraLayers();
    renderMarkers();
    buildLegend();
  });
}
// ============================================================
// HOODS — drawn as MapLibre fill + line layers
// ============================================================
function drawHoods() {
  const features = hoods
    .filter(h => h.geometry)
    .map(h => ({
      type: "Feature",
      geometry: h.geometry,
      properties: {
        hood_id:     h.hood_id     || "",
        nano_market: h.nano_market || "",
        micro_market: h.micro_market || "",
        region:      h.region      || ""
      }
    }));

  const geojson = { type: "FeatureCollection", features };

  if (map.getSource(HOODS_SOURCE)) {
    map.getSource(HOODS_SOURCE).setData(geojson);
    return;
  }

  map.addSource(HOODS_SOURCE, { type: "geojson", data: geojson });

  map.addLayer({
    id: HOODS_FILL,
    type: "fill",
    source: HOODS_SOURCE,
    paint: {
      "fill-color": "#4da6ff",
      "fill-opacity": 0.15
    }
  });

  map.addLayer({
    id: HOODS_LINE,
    type: "line",
    source: HOODS_SOURCE,
    paint: {
      "line-color": "#0055cc",
      "line-width": 1
    }
  });
}

function updateHoodVisibility() {
  const filterNM = activeFilters.NM || "";
  const filterMM = activeFilters.MM || "";
  const noFilter = !filterNM && !filterMM;

  if (!map.getLayer(HOODS_FILL)) return;

  if (!layerVisible.hoods) {
    map.setLayoutProperty(HOODS_FILL, "visibility", "none");
    map.setLayoutProperty(HOODS_LINE, "visibility", "none");
    return;
  }

  map.setLayoutProperty(HOODS_FILL, "visibility", "visible");
  map.setLayoutProperty(HOODS_LINE, "visibility", "visible");

  if (!noFilter) {
    // Filter by NM and/or MM using MapLibre expression
    const expr = ["all"];
    if (filterNM) expr.push(["==", ["get", "nano_market"],  filterNM]);
    if (filterMM) expr.push(["==", ["get", "micro_market"], filterMM]);
    map.setFilter(HOODS_FILL, expr);
    map.setFilter(HOODS_LINE, expr);
  } else {
    map.setFilter(HOODS_FILL, null);
    map.setFilter(HOODS_LINE, null);
  }
}

function assignHood(coords) {
  const pt = turf.point([coords.lng, coords.lat]);
  let nearest = null, minDist = Infinity;

  for (let h of hoods) {
    const polygon = { type: "Feature", geometry: h.geometry };
    try {
      // Step 1: exact containment — return immediately if inside
      if (turf.booleanPointInPolygon(pt, polygon)) return h;

      // Step 2: distance to nearest boundary edge (not centroid)
      // Convert polygon rings to line strings and measure point-to-line distance
      const geom = h.geometry;
      const rings = geom.type === "Polygon"
        ? geom.coordinates
        : geom.coordinates.flat(); // MultiPolygon — flatten to all rings

      let hoodMinDist = Infinity;
      for (const ring of rings) {
        if (ring.length < 2) continue;
        const line = turf.lineString(ring);
        const d    = turf.pointToLineDistance(pt, line, { units: "kilometers" });
        if (d < hoodMinDist) hoodMinDist = d;
      }

      if (hoodMinDist < minDist) { minDist = hoodMinDist; nearest = h; }
    } catch (e) {}
  }

  return nearest;
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
    { key: "demand",     color: "#2980b9", symbol: "🌊", label: "Demand Heatmap" },
    { key: "idle",       color: "#c0392b", symbol: "🔥", label: "Idle Heatmap"   },
    { key: "centroids",  color: "#27ae60", symbol: "C",  label: "Demand Centroids"}
  ];

  legend.innerHTML = `<div class="legend-title">Layers</div>` +
    items.map(item => `
      <div class="legend-item" id="legend_${item.key}" onclick="toggleLayer('${item.key}')" style="cursor:pointer">
        <span class="legend-symbol" style="color:${item.color};font-weight:bold">${item.symbol}</span>
        <span class="legend-label">${item.label}</span>
        <span class="legend-eye" id="eye_${item.key}">${layerVisible[item.key] ? "👁" : "🚫"}</span>
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
    updateHoodVisibility();
  }
  if (key === "properties") {
    propertyMarkers.forEach(m => {
      const el = m.getElement();
      el.style.display = layerVisible.properties ? "" : "none";
    });
  }
  if (key === "hotspots") {
    hotspotMarkers.forEach(m => {
      m.getElement().style.display = layerVisible.hotspots ? "" : "none";
    });
  }
  if (key === "demand") {
    if (map.getLayer(DEMAND_LAYER)) {
      map.setLayoutProperty(DEMAND_LAYER, "visibility", layerVisible.demand ? "visible" : "none");
    }
  }
  if (key === "idle") {
    if (map.getLayer(IDLE_LAYER)) {
      map.setLayoutProperty(IDLE_LAYER, "visibility", layerVisible.idle ? "visible" : "none");
    }
  }
  if (key === "centroids") {
    centroidMarkers.forEach(m => {
      m.getElement().style.display = layerVisible.centroids ? "" : "none";
    });
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

// ── NM/MM lookup for extra-layer rows ───────────────────────
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

function passesNMMFilter(row) {
  const filterNM = activeFilters.NM || "";
  const filterMM = activeFilters.MM || "";
  if (!filterNM && !filterMM) return true;
  if (filterNM && row._nm !== filterNM) return false;
  if (filterMM && row._mm !== filterMM) return false;
  return true;
}

// ── Helper: create a letter-badge DOM marker ────────────────
function createLetterMarker(lat, lng, letter, bg, popupHtml, extraRow) {
  const el = document.createElement("div");
  el.style.cssText = `
    background:${bg};color:#fff;border-radius:50%;
    width:24px;height:24px;display:flex;align-items:center;
    justify-content:center;font-weight:700;font-size:13px;
    border:2px solid rgba(0,0,0,0.25);box-shadow:0 1px 3px rgba(0,0,0,0.3);
    cursor:pointer;
  `;
  el.textContent = letter;
  el._extraRow = extraRow;

  const popup = new maplibregl.Popup({ offset: 14, closeButton: true })
    .setHTML(popupHtml);

  const marker = new maplibregl.Marker({ element: el })
    .setLngLat([lng, lat])
    .setPopup(popup);

  return marker;
}

// ── Render Hotspots ──────────────────────────────────────────
function renderHotspots(data) {
  hotspotMarkers.forEach(m => m.remove());
  hotspotMarkers = [];
  hotspotData = data;
  stampHoodInfo(data, "lat", "lng");

  data.forEach(row => {
    const lat = parseFloat(row.lat), lng = parseFloat(row.lng);
    if (isNaN(lat) || isNaN(lng)) return;
    const html = `<b>🔥 ${row.name || "Hotspot"}</b><br>Hood: ${row.hood || "-"}<br>Cluster: ${row.cluster || "-"}<br>NM: ${row._nm || "-"}<br>MM: ${row._mm || "-"}`;
    const m = createLetterMarker(lat, lng, "H", "#f39c12", html, row);
    if (layerVisible.hotspots && passesNMMFilter(row)) m.addTo(map);
    hotspotMarkers.push(m);
  });
  console.log(`📍 ${hotspotMarkers.length} hotspot markers`);
}

// ── Render Demand as real heatmap ────────────────────────────
function renderDemand(data) {
  demandData = data;
  stampHoodInfo(data, "lat", "lng");

  const geo = toGeoJSON(data, "num_points", "lat", "lng");

  if (map.getSource(DEMAND_SOURCE)) {
    map.getSource(DEMAND_SOURCE).setData(geo);
    return;
  }

  map.addSource(DEMAND_SOURCE, { type: "geojson", data: geo });

  map.addLayer({
    id: DEMAND_LAYER,
    type: "heatmap",
    source: DEMAND_SOURCE,
    layout: { visibility: layerVisible.demand ? "visible" : "none" },
    paint: {
    "heatmap-weight": 1,
    "heatmap-intensity": 1,
    "heatmap-radius": 20,

    "heatmap-color": [
      "interpolate", ["linear"], ["heatmap-density"],
      0,   "rgba(0,0,255,0)",
      0.2, "rgba(0,0,255,0.3)",
      0.4, "rgba(0,0,255,0.6)",
      0.7, "rgba(0,0,180,0.8)",
      1,   "rgba(0,0,100,1)"
    ],

    "heatmap-opacity": 0.8
  }
  });
  console.log(`🌊 Demand heatmap rendered — ${data.length} points`);
}

// ── Render Idle as real heatmap ──────────────────────────────
function renderIdle(data) {
  idleData = data;
  stampHoodInfo(data, "lat", "lng");

  const geo = toGeoJSON(data, "idle_min", "lat", "lng");

  if (map.getSource(IDLE_SOURCE)) {
    map.getSource(IDLE_SOURCE).setData(geo);
    return;
  }

  map.addSource(IDLE_SOURCE, { type: "geojson", data: geo });

  map.addLayer({
    id: IDLE_LAYER,
    type: "heatmap",
    source: IDLE_SOURCE,
    layout: { visibility: layerVisible.idle ? "visible" : "none" },
    paint: {
    "heatmap-weight": 1,
    "heatmap-intensity": 1,
    "heatmap-radius": 20,

    "heatmap-color": [
      "interpolate", ["linear"], ["heatmap-density"],
      0,   "rgba(255,0,0,0)",
      0.3, "rgba(172, 7, 7, 0.4)",
      0.6, "rgba(113, 3, 3, 0.7)",
      1,   "rgb(69, 3, 3)"
    ],

    "heatmap-opacity": 0.8
  }
  });
  console.log(`🔥 Idle heatmap rendered — ${data.length} points`);
}

// ── Render Centroids ─────────────────────────────────────────
function renderCentroids(data) {
  centroidMarkers.forEach(m => m.remove());
  centroidMarkers = [];
  centroidData = data;
  stampHoodInfo(data, "centroid_lat", "centroid_lng");

  data.forEach(row => {
    const lat = parseFloat(row.centroid_lat), lng = parseFloat(row.centroid_lng);
    if (isNaN(lat) || isNaN(lng)) return;
    const html = `<b>📊 ${row.hood_name || "Centroid"}</b><br>Cluster ID: ${row.cluster_id || "-"}<br>NM: ${row._nm || "-"}<br>MM: ${row._mm || "-"}`;
    const m = createLetterMarker(lat, lng, "C", "#27ae60", html, row);
    if (layerVisible.centroids && passesNMMFilter(row)) m.addTo(map);
    centroidMarkers.push(m);
  });
  console.log(`📍 ${centroidMarkers.length} centroid markers`);
}

// Convert flat array to GeoJSON FeatureCollection
function toGeoJSON(data, valueKey, latKey, lngKey) {
  return {
    type: "FeatureCollection",
    features: data
      .filter(r => !isNaN(parseFloat(r[latKey])) && !isNaN(parseFloat(r[lngKey])))
      .map(r => ({
        type: "Feature",
        geometry: {
          type: "Point",
          coordinates: [parseFloat(r[lngKey]), parseFloat(r[latKey])]
        },
        properties: {
          value:   parseFloat(r[valueKey]) || 0,
          cluster: r.cluster || "",
          hood:    r.hood    || ""
        }
      }))
  };
}

// Filter extra layer markers by NM/MM
function filterExtraLayers() {
  hotspotMarkers.forEach(m => {
    const show = layerVisible.hotspots && passesNMMFilter(m.getElement()._extraRow || {});
    show ? m.addTo(map) : m.remove();
  });
  centroidMarkers.forEach(m => {
    const show = layerVisible.centroids && passesNMMFilter(m.getElement()._extraRow || {});
    show ? m.addTo(map) : m.remove();
  });
  // Demand/idle heatmaps: re-render with filtered data
  if (map.getSource(DEMAND_SOURCE)) {
    const filtered = demandData.filter(passesNMMFilter);
    map.getSource(DEMAND_SOURCE).setData(toGeoJSON(filtered, "num_points", "lat", "lng"));
  }
  if (map.getSource(IDLE_SOURCE)) {
    const filtered = idleData.filter(passesNMMFilter);
    map.getSource(IDLE_SOURCE).setData(toGeoJSON(filtered, "idle_min", "lat", "lng"));
  }
}

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

    if (Object.values(activeFilters).some(v => v)) {
      filterAndRender();
    } else {
      renderMarkers();
    }

    const sfActive = Object.values(getSheetFilters()).some(v => v);
    renderSheetPreview(sfActive ? getSheetFilteredData() : allData);
    renderSummaryTables();
    renderReminderTable();
    renderBangaloreOverview();
  } catch (err) {
    console.error("❌ Fetch failed:", err);
  }
  populateFilters();
}

// ============================================================
// RECALCULATE HOTSPOT DATA
// ============================================================
async function recalculateHotspots() {
  const btn = document.getElementById("btnRecalcHotspots");
  if (btn) { btn.disabled = true; btn.textContent = "⏳ Recalculating…"; }
  const statusEl = document.getElementById("recalcStatus");
  if (statusEl) { statusEl.textContent = "Fetching & enriching hotspot data from server…"; statusEl.style.color = "#555"; }
  try {
    const url = CONFIG.API_URL + "?recalc=1&t=" + Date.now();
    const res = await fetch(url, { method: "GET", redirect: "follow", credentials: "omit" });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const data = await res.json();
    if (!Array.isArray(data)) throw new Error(data.error || "Unexpected response");

    if (statusEl) statusEl.textContent = `⏳ Writing ${data.length} enriched rows back to sheet…`;

    const ENRICHED_COLS = ["nearest hotspot","displacement to nearest hotspot","demand count","idle count","launch feasibility"];
    const writes = data.filter(row => row._rowIndex).map(row => {
      const payload = { _rowIndex: row._rowIndex };
      ENRICHED_COLS.forEach(col => { if (row[col] !== undefined) payload[col] = row[col]; });
      return fetch(CONFIG.API_URL, { method: "POST", mode: "no-cors", credentials: "omit", body: JSON.stringify(payload) });
    });
    await Promise.all(writes);

    allData = data;
    if (statusEl) { statusEl.textContent = `✅ Done — ${data.length} rows enriched and saved.`; statusEl.style.color = "#27ae60"; }
    renderMarkers();
    renderSheetPreview(allData);
    renderSummaryTables();
  } catch (err) {
    if (statusEl) { statusEl.textContent = "❌ Error: " + err.message; statusEl.style.color = "#c0392b"; }
  } finally {
    if (btn) { btn.disabled = false; btn.textContent = "🔄 Recalculate Hotspot Data"; }
  }
}

function getPropertyName(row) {
  return row["Name of the property"] || row["Name"] || "No Name";
}

// ── Create a DOM element for a property pin ──────────────────
function createPropertyMarkerEl(row) {
  const el = document.createElement("div");
  el.style.cssText = "font-size:18px;cursor:pointer;user-select:none;line-height:1";
  el.textContent = "📍";
  return el;
}

// ============================================================
// RENDER MARKERS (all properties)
// ============================================================
function renderMarkers() {
  propertyMarkers.forEach(m => m.remove());
  propertyMarkers = [];
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
              .catch(err => console.error("❌ Auto-save failed", err));
          }
        }
      }
      const name = getPropertyName(row);
      const el = createPropertyMarkerEl(row);
      el.addEventListener("click", () => showDetails(row));

      const popup = new maplibregl.Popup({ offset: 14 })
        .setHTML(`<b>${escHtml(name)}</b><br>${row.Category || ""}<br>NM: ${row.NM || "-"}<br>MM: ${row.MM || "-"}`);

      const marker = new maplibregl.Marker({ element: el })
        .setLngLat([lng, lat])
        .setPopup(popup);

      if (layerVisible.properties) marker.addTo(map);
      propertyMarkers.push(marker);
    } else {
      skipped++;
    }
  });
  console.log(`📍 ${propertyMarkers.length} markers, ${skipped} skipped`);
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

  console.log(`🔽 ${filtered.length}/${allData.length} rows match filters`);
  updateHoodVisibility();
  renderFilteredMarkers(filtered, true);
  filterExtraLayers();
}

function renderFilteredMarkers(data, fitView = false) {
  propertyMarkers.forEach(m => m.remove());
  propertyMarkers = [];
  const bounds = new maplibregl.LngLatBounds();
  let hasPoints = false;

  data.forEach(row => {
    const lat = parseFloat(row.Lat), lng = parseFloat(row.Long);
    if (!isNaN(lat) && !isNaN(lng)) {
      const name = getPropertyName(row);
      const el = createPropertyMarkerEl(row);
      el.addEventListener("click", () => showDetails(row));

      const popup = new maplibregl.Popup({ offset: 14 })
        .setHTML(`<b>${escHtml(name)}</b><br>${row.Category || ""}<br>NM: ${row.NM || "-"}<br>MM: ${row.MM || "-"}`);

      const marker = new maplibregl.Marker({ element: el })
        .setLngLat([lng, lat])
        .setPopup(popup);

      if (layerVisible.properties) marker.addTo(map);
      propertyMarkers.push(marker);
      bounds.extend([lng, lat]);
      hasPoints = true;
    }
  });

  if (fitView && hasPoints) {
    map.fitBounds(bounds, { padding: 60, maxZoom: 16 });
  }
}

function clearFilters() {
  activeFilters = {};
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
// ADD POINT MODE
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
    map.getContainer().classList.add("map-crosshair");
  } else {
    btn.textContent = "➕ Add Point";
    btn.style.background = "";
    map.getContainer().classList.remove("map-crosshair");
    if (addPointMarker) { addPointMarker.remove(); addPointMarker = null; }
  }
}

function openAddPointModal(lat, lng) {
  if (addPointMarker) addPointMarker.remove();
  const el = document.createElement("div");
  el.style.cssText = "font-size:24px;cursor:pointer";
  el.textContent = "📌";
  addPointMarker = new maplibregl.Marker({ element: el })
    .setLngLat([lng, lat])
    .addTo(map);

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
  if (addPointMarker) { addPointMarker.remove(); addPointMarker = null; }
  if (addPointMode) toggleAddPointMode();
}

async function submitAddPoint() {
  const lat = document.getElementById("ap_Lat").value;
  const lng = document.getElementById("ap_Long").value;
  const newRow = { "Lat": parseFloat(lat), "Long": parseFloat(lng) };

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
// DOWNLOAD KML / CSV
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

function getFilteredNMs() {
  // Use hoods as source of truth — all NMs from polygon sheet
  // If a specific NM filter is active, return just that; if MM filter, filter by MM
  const filterNM = activeFilters.NM || "";
  const filterMM = activeFilters.MM || "";
  if (filterNM) return [filterNM];
  if (filterMM) return [...new Set(hoods.filter(h => h.micro_market === filterMM).map(h => h.nano_market).filter(Boolean))];
  return [...new Set(hoods.map(h => h.nano_market).filter(Boolean))];
}

function getFilteredMMs() {
  // Use hoods as source of truth — all MMs from polygon sheet
  const filterNM = activeFilters.NM || "";
  const filterMM = activeFilters.MM || "";
  if (filterMM) return [filterMM];
  if (filterNM) return [...new Set(hoods.filter(h => h.nano_market === filterNM).map(h => h.micro_market).filter(Boolean))];
  return [...new Set(hoods.map(h => h.micro_market).filter(Boolean))];
}
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
    <n>${escXml(h.nano_market || h.micro_market || h.hood_id)}</n>
    <description><![CDATA[NM: ${h.nano_market || ""}<br>MM: ${h.micro_market || ""}<br>ID: ${h.hood_id || ""}]]></description>
    <Style><PolyStyle><color>${color}</color><outline>1</outline></PolyStyle></Style>
    ${geometryToKmlGeometry(h.geometry)}
  </Placemark>`).join("\n");
  return `<?xml version="1.0" encoding="UTF-8"?>\n<kml xmlns="http://www.opengis.net/kml/2.2"><Document><n>${escXml(layerName)}</n>\n${placemarks}\n</Document></kml>`;
}

function pointsToKml(data, layerName) {
  const placemarks = data
    .filter(row => !isNaN(parseFloat(row.Lat)) && !isNaN(parseFloat(row.Long)))
    .map(row => `
  <Placemark>
    <n>${escXml(getPropertyName(row))}</n>
    <description><![CDATA[Category: ${row.Category || ""}<br>NM: ${row.NM || ""}<br>MM: ${row.MM || ""}<br>Road: ${row.Road || ""}<br>Status: ${row["Final Status"] || ""}]]></description>
    <Point><coordinates>${parseFloat(row.Long)},${parseFloat(row.Lat)},0</coordinates></Point>
  </Placemark>`).join("\n");
  return `<?xml version="1.0" encoding="UTF-8"?>\n<kml xmlns="http://www.opengis.net/kml/2.2"><Document><n>${escXml(layerName)}</n>\n${placemarks}\n</Document></kml>`;
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
    const hoodList = hoodsByMM(getFilteredMMs());
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
    const hoodList = hoodsByMM(getFilteredMMs());
    if (!hoodList.length) { alert("No MM hoods found."); return; }
    downloadBlob(hoodsToCsvWkt(hoodList, "micro_market"), `mm_layer_${label}.csv`, "text/csv");
  } else if (type === "points") {
    if (!filteredData.length) { alert("No data points."); return; }
    downloadBlob(pointsToCsvWkt(filteredData), `points_${label}.csv`, "text/csv");
  }
}

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
  const directMatch = raw.match(/@(-?\d+\.\d+),(-?\d+\.\d+)/);
  if (directMatch) return { lat: parseFloat(directMatch[1]), lng: parseFloat(directMatch[2]) };
  const backendUrl = CONFIG.API_URL + "?action=resolveUrl&url=" + encodeURIComponent(raw);
  const res  = await fetch(backendUrl);
  if (!res.ok) throw new Error(`Backend resolve failed: ${res.status}`);
  const text = await res.text();
  const json = JSON.parse(text);
  if (json.error) throw new Error(`Backend error: ${json.error}`);
  if (json.lat != null && json.lng != null) return { lat: parseFloat(json.lat), lng: parseFloat(json.lng) };
  const expandedUrl = json.url || "";
  const d3Match = expandedUrl.match(/!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)/);
  if (d3Match) return { lat: parseFloat(d3Match[1]), lng: parseFloat(d3Match[2]) };
  const atMatch = expandedUrl.match(/@(-?\d+\.\d+),(-?\d+\.\d+)/);
  if (atMatch) return { lat: parseFloat(atMatch[1]), lng: parseFloat(atMatch[2]) };
  throw new Error(`Could not extract coords from: "${expandedUrl || raw}"`);
}

async function fillLatLong() {
  let updated = [], skippedCount = 0, failedCount = 0;
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
  updated.forEach(r => {
    const payload = { _rowIndex: r._rowIndex, Lat: r.Lat, Long: r.Long };
    fetch(CONFIG.API_URL, { method: "POST", body: JSON.stringify(payload) });
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
    const payload = { _rowIndex: r._rowIndex, NM: r.NM, MM: r.MM, "NM Id": r["NM Id"] };
    fetch(CONFIG.API_URL, { method: "POST", body: JSON.stringify(payload) });
  });
  alert(`✅ Updated ${updated.length} rows`);
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
  "Property": ["Private", "Public"],
  "App status": ["Active", "Inactive"],
  "Lead Status": ["2. Owner conversation pending","3. Owner's confirmation pending","4. Confirmed","5. Follow up required","6. Dropped"],
  "Final Status": ["Dropped off", "Active", "Cold", "Deal closed","Deal closed - sign pending", "No deal required","To be reactivated", "Deal - closed - Chairs pending","Dropped off after launch"],
  "Closure type": ["Resting + Washroom", "Resting", "NA"],
  "Set up": ["Chairs to be set", "Owner will setup chairs", "Chairs available", "NA"],
  "Category": ["Ladies PG", "Shop", "Restaurant", "Apartment", "Gated community", "Independent Builder floor","Bus Stop", "Park", "Petrol Pump", "Public Washroom", "Other"]
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
          <td><select class="detail-select" data-key="${escHtml(key)}">
            <option value="">-- select --</option>
            ${extraOption}${options}
          </select></td>
        </tr>`;
    } else if (key === "Timestamp") {
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

function formatTimestamp(date) {
  const ist = new Date(date.getTime() + 5.5 * 60 * 60 * 1000);
  return `${ist.getUTCMonth()+1}/${ist.getUTCDate()}/${ist.getUTCFullYear()} ` +
    `${String(ist.getUTCHours()).padStart(2,'0')}:${String(ist.getUTCMinutes()).padStart(2,'0')}:${String(ist.getUTCSeconds()).padStart(2,'0')}`;
}

function saveCurrent() {
  if (!currentRow) return;

  const prevAppStatus   = String(currentRow["App status"]   || "").trim();
  const prevFinalStatus = String(currentRow["Final Status"] || "").trim();

  document.querySelectorAll("[contenteditable]").forEach(cell => {
    const key = cell.dataset.key;
    if (key) currentRow[key] = cell.innerText.trim();
  });

  document.querySelectorAll(".detail-select").forEach(sel => {
    const key = sel.dataset.key;
    if (key) currentRow[key] = sel.value;
  });

  if (!currentRow._rowIndex) { alert("❌ Cannot save — row index missing, try refreshing"); return; }

  const nowIST = formatTimestamp(new Date());

  const newAppStatus = String(currentRow["App status"] || "").trim();
  if (newAppStatus === "Active" && prevAppStatus !== "Active") {
    currentRow["Launch date"] = nowIST;
    console.log("🚀 Launch date auto-filled:", nowIST);
  }

  const CLOSED_STATUSES = ["Deal closed", "Deal - closed - Chairs pending"];
  const newFinalStatus = String(currentRow["Final Status"] || "").trim();
  if (CLOSED_STATUSES.includes(newFinalStatus) && !CLOSED_STATUSES.includes(prevFinalStatus)) {
    currentRow["Signage date"] = nowIST;
    console.log("📅 Signage date auto-filled:", nowIST);
  }

  const savePayload = Object.assign({}, currentRow);
  delete savePayload["Timestamp"];

  fetch(CONFIG.API_URL, { method: "POST", body: JSON.stringify(savePayload) })
    .then(() => { alert("✅ Saved"); showDetails(currentRow); })
    .catch(err => alert("❌ Save failed: " + err.message));
}

// ============================================================
// TIMESTAMP UTILITIES
// ============================================================
function parseTimestamp(val) {
  if (!val || val === "") return null;
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val;
  if (typeof val === "number") {
    const d = new Date((val - 25569) * 86400000);
    return isNaN(d.getTime()) ? null : d;
  }
  const str = val.toString().trim();
  if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/.test(str)) {
    const d = new Date(str);
    return isNaN(d.getTime()) ? null : d;
  }
  const slashMatch = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2}):(\d{2})$/);
  if (slashMatch) {
    const [, m, d, y, hr, min, sec] = slashMatch;
    return new Date(+y, +m - 1, +d, +hr, +min, +sec);
  }
  const fallback = new Date(str);
  return isNaN(fallback.getTime()) ? null : fallback;
}

function formatTsDisplay(val) {
  if (!val || val === "") return "";
  const d = parseTimestamp(val);
  if (!d) return String(val);
  const str = val.toString().trim();
  const isSlashFormat = /^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2}):(\d{2})$/.test(str);
  if (isSlashFormat) {
    return `${d.getDate()}/${d.getMonth()+1}/${d.getFullYear()} ` +
      `${String(d.getHours()).padStart(2,'0')}:${String(d.getMinutes()).padStart(2,'0')}:${String(d.getSeconds()).padStart(2,'0')}`;
  }
  const ist = new Date(d.getTime() + 5.5 * 60 * 60 * 1000);
  return `${ist.getUTCDate()}/${ist.getUTCMonth()+1}/${ist.getUTCFullYear()} ` +
    `${String(ist.getUTCHours()).padStart(2,'0')}:${String(ist.getUTCMinutes()).padStart(2,'0')}:${String(ist.getUTCSeconds()).padStart(2,'0')}`;
}

// ============================================================
// SHEET PREVIEW
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
  const sf = getSheetFilters();
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
    if (from || to) {
      const ts = parseTimestamp(row["Timestamp"]);
      if (!ts) return false;
      if (from && ts < from) return false;
      if (to   && ts > to)   return false;
    }
    if (signageFrom || signateTo) {
      const ts = parseTimestamp(row["Signage date"]);
      if (!ts) return false;
      if (signageFrom && ts < signageFrom) return false;
      if (signateTo   && ts > signateTo)   return false;
    }
    if (launchFrom || launchTo) {
      const ts = parseTimestamp(row["Launch date"]);
      if (!ts) return false;
      if (launchFrom && ts < launchFrom) return false;
      if (launchTo   && ts > launchTo)   return false;
    }
    return true;
  });
}

function applySheetFilters() { renderSheetPreview(getSheetFilteredData()); }

function clearSheetFilters() {
  ["sfCategory","sfProperty","sfAppStatus","sfLeadStatus","sfFinalStatus","sfNM","sfMM"].forEach(id => {
    const el = document.getElementById(id); if (el) el.value = "";
  });
  ["sfDateFrom","sfDateTo","sfSignageFrom","sfSignageTo","sfLaunchFrom","sfLaunchTo"].forEach(id => {
    const el = document.getElementById(id); if (el) el.value = "";
  });
  renderSheetPreview(allData);
}

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

  const previewCols = [
    "MM", "NM", "NM Id", "Name of the property", "App status", "Category",
    "Closure type", "Lat", "Long", "Location (Google Maps URL) / Map Code",
    "Owner Contact Name", "Owner Contact Number",
    "Contact Name", "Contact number", "Owner Designation",
    "Property", "Signage date", "Launch date",
    "Photo 1 (Image Upload) (From Road)",
    "Photo 2 (Image Upload) (Sitting Area)",
    "Photo 3 (Image Upload)", "Agreement Photo (Image Upload)",
    "Lead Status", "Final Status"
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
    Property: document.getElementById("stProperty")?.value || "",
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
    if (sf.Property && row.Property !== sf.Property) return false;
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
  ["stNM","stMM","stProperty","stDateFrom","stDateTo"].forEach(id => {
    const el = document.getElementById(id); if (el) el.value = "";
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
    "Deal closed - sign pending", "Places Finalised", "Launched",
    "Deal - closed - Chairs pending", "Deal closed", "No deal required"
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
        ["Deal closed", "Deal - closed - Chairs pending", "No deal required"].includes(r["Final Status"])
      );
      const totalRows = data.length;
      const pct = (num, denom) => !denom ? "0 (0%)" : `${num} (${(num / denom * 100).toFixed(1)}%)`;
      const cats = countByCat(finalisedRows);
      return `<tr class="summary-finalised-row">
        <td>Places Finalised</td>
        <td>${pct(finalisedRows.length, totalRows)}</td>
        ${allCats.map(c => `<td>${pct(cats[c] || 0, data.filter(r => normalizeCategory(r.Category) === c).length)}</td>`).join("")}
      </tr>`;
    } else if (status === "Launched") {
      const launchedRows = data.filter(r => (r["App status"] || "").trim() === "Active");
      const totalRows = data.length;
      const pct = (num, denom) => !denom ? "0 (0%)" : `${num} (${(num / denom * 100).toFixed(1)}%)`;
      const cats = countByCat(launchedRows);
      return `<tr class="summary-finalised-row" style="background:#e8f0ff!important;color:#1a3a7a">
        <td>Launched</td>
        <td>${pct(launchedRows.length, totalRows)}</td>
        ${allCats.map(c => `<td>${pct(cats[c] || 0, data.filter(r => normalizeCategory(r.Category) === c).length)}</td>`).join("")}
      </tr>`;
    } else {
      rows = data.filter(r => r["Final Status"] === status);
    }

    const cats = countByCat(rows);
    const total = rows.length;
    return `<tr>
      <td>${escHtml(status)}</td>
      <td>${total}</td>
      ${allCats.map(c => `<td>${cats[c] || 0}</td>`).join("")}
    </tr>`;
  }).join("");

  const section1 = `
    <div class="summary-block">
      <div class="summary-block-header">
        <h4 class="summary-block-title">Final Status × Category</h4>
        <button class="summary-dl-btn" onclick="downloadSummaryTable('stMainTable','status_x_category')">⬇ CSV</button>
      </div>
      <div class="summary-table-wrapper">
        <table class="summary-table" id="stMainTable">
          <thead><tr>${headerCols}</tr></thead>
          <tbody>${statusRows}</tbody>
        </table>
      </div>
    </div>`;

  // Commercials breakdown
  const FIXED_COMMERCIAL_BUCKETS = ["2000", "2500", "3000", "3500", "4000", "others", "NA"];
  const allCommercials = FIXED_COMMERCIAL_BUCKETS;

  const headerCols2 = `<th>Closure Type / Value</th><th>Total</th>${
    allCats.map(c => `<th>${escHtml(c)}</th>`).join("")
  }`;

  const closureTypes = [...new Set(data.map(r => r["Closure type"] || "NA").filter(Boolean))].sort();
  const commercialRows = closureTypes.map(closureType => {
    const closureRows = data.filter(r => (r["Closure type"] || "NA") === closureType);
    const headerRow = `<tr class="summary-closure-header">
      <td colspan="${allCats.length + 2}"><b>${escHtml(closureType)}</b></td>
    </tr>`;

    const valueRows = allCommercials.map(bucket => {
      const bucketRows = closureRows.filter(r => normalizeCommercial(r["Commercials"]) === bucket);
      const cats = countByCat(bucketRows);
      return `<tr>
        <td style="padding-left:16px">${bucket}</td>
        <td>${bucketRows.length}</td>
        ${allCats.map(c => `<td>${cats[c] || 0}</td>`).join("")}
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
          <thead><tr>${headerCols2}</tr></thead>
          <tbody>${commercialRows}</tbody>
        </table>
      </div>
    </div>`;

  container.innerHTML = section1 + section2;
  renderNmMmSummary();
  renderBangaloreOverview();
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
// MAP SEARCH (Nominatim — unchanged logic, updated for MapLibre)
// ============================================================
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
  if (searchMarker) { searchMarker.remove(); searchMarker = null; }
  const el = document.createElement("div");
  el.style.cssText = "font-size:28px;cursor:pointer;animation:pulse 1.5s ease infinite";
  el.textContent = "📌";

  const popup = new maplibregl.Popup({ offset: 14 })
    .setHTML(`<b>${escHtml(name)}</b><br><small>${lat}, ${lng}</small>`);

  searchMarker = new maplibregl.Marker({ element: el })
    .setLngLat([parseFloat(lng), parseFloat(lat)])
    .setPopup(popup)
    .addTo(map);

  searchMarker.togglePopup();
  map.flyTo({ center: [parseFloat(lng), parseFloat(lat)], zoom: 16 });
  document.getElementById("mapSearchInput").value = name;
  document.getElementById("mapSearchResults").classList.remove("open");
}

// ============================================================
// INCENTIVE TRACKER
// ============================================================
function getIncentiveFilters() {
  return {
    dateFrom:   document.getElementById("incDateFrom")?.value   || "",
    dateTo:     document.getElementById("incDateTo")?.value     || "",
    proximity:  document.getElementById("incProximity")?.value  || "all",
    duplicates: document.getElementById("incDuplicates")?.value || "all",
  };
}

const LAUNCH_MILESTONES = [0, 100, 200, 400, 600, 800, 1000];

function calcPrivateIncentive(n, washroomCount) {
  if (!n || n <= 0) return 0;
  let base = 0;
  if (n < LAUNCH_MILESTONES.length) {
    base = LAUNCH_MILESTONES[n];
  } else {
    base = 1000 + (n - 6) * 200;
  }
  const washroomBonus = (washroomCount || 0) * 100;
  return base + washroomBonus;
}

function renderIncentiveTables() {
  const container = document.getElementById("incentiveContainer");
  if (!container) return;

  const { dateFrom, dateTo, proximity, duplicates } = getIncentiveFilters();
  const from = dateFrom ? new Date(dateFrom + "T00:00:00") : null;
  const to   = dateTo   ? new Date(dateTo   + "T23:59:59") : null;

  let base = allData;
  if (proximity === "250") {
    base = allData.filter(r => {
      const d = parseFloat(r["displacement to nearest hotspot"]);
      return !isNaN(d) && d <= 250;
    });
  }

  if (duplicates === "exclude") {
    base = base.filter(r => (r["Duplicate"] || "").trim().toLowerCase() !== "duplicate");
  } else if (duplicates === "only") {
    base = base.filter(r => (r["Duplicate"] || "").trim().toLowerCase() === "duplicate");
  }

  const srNoToName = {};
  allData.forEach(r => {
    const srNo = String(r["Sr No"] || "").trim();
    if (srNo && (r["Duplicate"] || "").trim().toLowerCase() !== "duplicate") {
      srNoToName[srNo] = (r["Name of the property"] || "").trim();
    }
  });

  const inRange = (val) => {
    if (!val) return false;
    const ts = parseTimestamp(val);
    if (!ts) return false;
    if (from && ts < from) return false;
    if (to   && ts > to)   return false;
    return true;
  };

  const emailGroups = {};
  base.forEach(r => {
    const email = (r["Email"] || r["email"] || "").trim().toLowerCase();
    if (!email) return;
    if (!emailGroups[email]) emailGroups[email] = [];
    emailGroups[email].push(r);
  });

  function buildTableData(propType) {
    return Object.entries(emailGroups).map(([email, rows]) => {
      const name = rows.find(r => r["Name"] || r["name"])?.["Name"] || rows.find(r => r["name"])?.["name"] || email;

      const launchRows = rows.filter(r => r["Property"] === propType && inRange(r["Launch date"]));
      const launchCount = launchRows.length;
      const launchNames = launchRows.map(r => r["Name of the property"] || "—").join(", ");
      const washroomCount = launchRows.filter(r => (r["Closure type"] || "").toLowerCase().includes("washroom")).length;

      const leadRowsInRange = rows.filter(r => r["Property"] === propType && inRange(r["Timestamp"]));
      const leadCount = leadRowsInRange.length;
      const nonDupLeadCount = leadRowsInRange.filter(r =>
        (r["Duplicate"] || "").trim().toLowerCase() !== "duplicate"
      ).length;

      const dupPairs = [];
      leadRowsInRange.forEach(r => {
        if ((r["Duplicate"] || "").trim().toLowerCase() !== "duplicate") return;
        const thisProp = (r["Name of the property"] || "—").trim();
        const srNo     = String(r["Sr No"] || "").trim();
        const original = allData.find(o =>
          String(o["Sr No"] || "").trim() === srNo &&
          (o["Duplicate"] || "").trim().toLowerCase() !== "duplicate" &&
          (o["Name of the property"] || "").trim() !== thisProp
        );
        const origName = original
          ? (original["Name of the property"] || "").trim()
          : (srNoToName[srNo] || `Sr No ${srNo}`);
        dupPairs.push({ thisProp, origName, srNo });
      });

      const dealRows  = rows.filter(r => r["Property"] === propType && inRange(r["Signage date"]));
      const dealCount = dealRows.length;
      const dealNames = dealRows.map(r => r["Name of the property"] || "—").join(", ");

      const incentive = propType === "Private"
        ? calcPrivateIncentive(launchCount, washroomCount)
        : launchCount * 20;

      return { email, name, launchCount, launchNames, washroomCount, leadCount, nonDupLeadCount, dupPairs, dealCount, dealNames, incentive };
    }).filter(r => r.launchCount > 0 || r.leadCount > 0 || r.dealCount > 0);
  }

  function renderTable(tableData, propType, tableId) {
    if (!tableData.length) {
      return `<p style="color:#aaa;font-size:13px;padding:8px 0">No data for ${propType} in this range.</p>`;
    }

    const totalLaunches    = tableData.reduce((s, r) => s + r.launchCount,    0);
    const totalLeads       = tableData.reduce((s, r) => s + r.leadCount,      0);
    const totalNonDupLeads = tableData.reduce((s, r) => s + r.nonDupLeadCount, 0);
    const totalDeals       = tableData.reduce((s, r) => s + r.dealCount,      0);
    const totalIncentive   = tableData.reduce((s, r) => s + r.incentive,      0);
    const totalWashrooms   = propType === "Private"
      ? tableData.reduce((s, r) => s + (r.washroomCount || 0), 0)
      : null;

    const rows = tableData.map(r => {
      const dupCell = r.dupPairs.length
        ? r.dupPairs.map(p =>
            `<span class="dup-pair">
              <span class="dup-this">${escHtml(p.thisProp)}</span>
              <span class="dup-arrow">→</span>
              <span class="dup-orig">${escHtml(p.origName)}</span>
            </span>`
          ).join("")
        : `<span style="color:#ccc">—</span>`;

      return `
        <tr>
          <td style="font-size:12px;white-space:nowrap">${escHtml(r.name || "—")}</td>
          <td style="font-size:11px;color:#555">${escHtml(r.email)}</td>
          <td style="text-align:center"><b>${r.launchCount}</b></td>
          <td class="inc-names-cell">${escHtml(r.launchNames || "—")}</td>
          ${propType === "Private"
            ? `<td style="text-align:center;color:${r.washroomCount > 0 ? '#0077b6' : '#aaa'}">${r.washroomCount > 0 ? '+₹' + (r.washroomCount * 100) : '—'}</td>`
            : ""}
          <td style="text-align:center">${r.leadCount}</td>
          <td style="text-align:center;font-weight:600;color:${r.nonDupLeadCount < r.leadCount ? '#e67e22' : '#27ae60'}">${r.nonDupLeadCount}</td>
          <td class="dup-pairs-cell">${dupCell}</td>
          <td style="text-align:center"><b>${r.dealCount}</b></td>
          <td class="inc-names-cell">${escHtml(r.dealNames || "—")}</td>
          <td style="text-align:right;font-weight:700;color:${r.incentive > 0 ? '#27ae60' : '#aaa'}">₹${r.incentive}</td>
        </tr>`;
    }).join("");

    const privateExtraTh = propType === "Private"
      ? `<th style="text-align:center">🚿 Washroom<br>Bonus</th>`
      : "";
    const privateExtraTd = propType === "Private"
      ? `<td style="text-align:center;color:#0077b6;font-weight:700">${totalWashrooms > 0 ? '+₹' + (totalWashrooms * 100) : '—'}</td>`
      : "";

    return `
      <div style="overflow-x:auto">
        <table class="summary-table incentive-table" id="${tableId}"
               style="min-width:${propType === "Private" ? "1100px" : "980px"}">
          <thead>
            <tr>
              <th>Name</th>
              <th>Email</th>
              <th style="text-align:center">Launched</th>
              <th>Launched Properties</th>
              ${privateExtraTh}
              <th style="text-align:center">Leads</th>
              <th style="text-align:center">Non-Dup<br>Leads</th>
              <th>Duplicate Map</th>
              <th style="text-align:center">Deals<br>Closed</th>
              <th>Closed Properties</th>
              <th style="text-align:right">Incentive</th>
            </tr>
          </thead>
          <tbody>${rows}</tbody>
          <tfoot>
            <tr class="summary-total-row">
              <td></td>
              <td><b>Total</b></td>
              <td style="text-align:center"><b>${totalLaunches}</b></td>
              <td></td>
              ${privateExtraTd}
              <td style="text-align:center"><b>${totalLeads}</b></td>
              <td style="text-align:center"><b>${totalNonDupLeads}</b></td>
              <td></td>
              <td style="text-align:center"><b>${totalDeals}</b></td>
              <td></td>
              <td style="text-align:right;font-weight:700;color:#27ae60">₹${totalIncentive}</td>
            </tr>
          </tfoot>
        </table>
      </div>
      <button class="summary-dl-btn" style="margin-top:6px"
              onclick="downloadSummaryTable('${tableId}','incentive_${propType.toLowerCase()}')">⬇ CSV</button>`;
  }

  const privateData = buildTableData("Private");
  const publicData  = buildTableData("Public");

  const proximityLabel = proximity === "250"
    ? `<span style="background:#e74c3c;color:#fff;font-size:11px;padding:2px 8px;border-radius:10px;margin-left:8px">≤250m from hotspot</span>`
    : `<span style="background:#2980b9;color:#fff;font-size:11px;padding:2px 8px;border-radius:10px;margin-left:8px">All properties</span>`;

  const duplicatesLabel = duplicates === "exclude"
    ? `<span style="background:#e67e22;color:#fff;font-size:11px;padding:2px 8px;border-radius:10px;margin-left:8px">Duplicates excluded</span>`
    : duplicates === "only"
    ? `<span style="background:#8e44ad;color:#fff;font-size:11px;padding:2px 8px;border-radius:10px;margin-left:8px">Duplicates only</span>`
    : "";

  const incentiveNote = `
    <div style="background:#fffbe6;border:1px solid #f0d060;border-radius:6px;padding:10px 14px;font-size:12px;color:#7a6000;margin-bottom:16px;line-height:1.8">
      <b>Incentive formula —</b><br>
      <b>Private Launches:</b> 1=₹100 &nbsp;·&nbsp; 2=₹200 &nbsp;·&nbsp; 3=₹400 &nbsp;·&nbsp; 4=₹600 &nbsp;·&nbsp; 5=₹800 &nbsp;·&nbsp; 6=₹1000 (each extra adds ₹200 more)
      &nbsp;&nbsp;+&nbsp;&nbsp;
      <b>🚿 Washroom Bonus: ₹100 per property</b> with "Washroom" in Closure type
      &nbsp;&nbsp;|&nbsp;&nbsp;
      <b>Public:</b> ₹20 per launched property
    </div>`;

  container.innerHTML = incentiveNote + `
    <div class="summary-block">
      <div class="summary-block-header">
        <h4 class="summary-block-title">🏠 Private Properties ${proximityLabel}${duplicatesLabel}</h4>
      </div>
      ${renderTable(privateData, "Private", "incPrivateTable")}
    </div>
    <div class="summary-block" style="margin-top:24px">
      <div class="summary-block-header">
        <h4 class="summary-block-title">🏢 Public Properties ${proximityLabel}${duplicatesLabel}</h4>
      </div>
      ${renderTable(publicData, "Public", "incPublicTable")}
    </div>`;
}

function applyIncentiveFilters() { renderIncentiveTables(); }

function clearIncentiveFilters() {
  ["incDateFrom", "incDateTo"].forEach(id => {
    const el = document.getElementById(id); if (el) el.value = "";
  });
  const prox = document.getElementById("incProximity");
  if (prox) prox.value = "all";
  const dup = document.getElementById("incDuplicates");
  if (dup) dup.value = "all";
  renderIncentiveTables();
}

// ============================================================
// BANGALORE OVERVIEW DASHBOARD
// ============================================================

// Returns 0–1 similarity score between two strings (LCS ratio after normalisation)
function fuzzyScore(str, query) {
  if (!str || !query) return 0;
  const s = str.toLowerCase().replace(/[^a-z0-9]/g, "");
  const q = query.toLowerCase().replace(/[^a-z0-9]/g, "");
  if (s === q) return 1;
  if (s.includes(q) || q.includes(s)) return 0.85;
  // LCS-based similarity
  const m = s.length, n = q.length;
  const dp = Array.from({length: m+1}, () => new Array(n+1).fill(0));
  for (let i = 1; i <= m; i++)
    for (let j = 1; j <= n; j++)
      dp[i][j] = s[i-1] === q[j-1] ? dp[i-1][j-1]+1 : Math.max(dp[i-1][j], dp[i][j-1]);
  return dp[m][n] / Math.max(m, n);
}

function fuzzyMatch(str, query) { return fuzzyScore(str, query) >= 0.5; }

// Region names as provided — these are MM names used as fallback
const BANGALORE_REGION_MMS = {
  "Mid Belt": [
    "Bellandur","Brookefield","Hoodi","Kudlu","Mahadevapura","Marathahalli",
    "Munnekollal","Sarjapur","Whitefield 1","Whitefield 2","Whitefield 3",
    "HSR","Indiranagar","Koramangala","Varthur"
  ],
  "South": [
    "Hongasandra","Hulimavu","Nagasandra","Singasandra","Tejaswini Nagar",
    "Rayasandra","Electronic City 1","Electronic City 2"
  ],
  "North": [
    "Hebbal","Thannisandra","Yelahanka","Yeswanthpur","Segehalli"
  ]
};

// Build region→MM mapping via fuzzy matching of hood micro_market names
// against the known MM lists, since the region field is not subdivided.
function buildRegionMaps() {
  const allHoodMMs = [...new Set(hoods.map(h => h.micro_market).filter(Boolean))];

  // Debug: log actual MM values from hoods so mismatches can be spotted
  console.log("🗺️ Hood MMs available:", allHoodMMs.sort().join(", "));

  const result = { source: "fuzzy_mm", "Mid Belt": [], "North": [], "South": [] };
  const unmatched = [];

  allHoodMMs.forEach(mm => {
    let bestRegion = null, bestScore = 0;

    for (const [region, knownMMs] of Object.entries(BANGALORE_REGION_MMS)) {
      for (const known of knownMMs) {
        const score = fuzzyScore(mm, known);
        if (score > bestScore) { bestScore = score; bestRegion = region; }
      }
    }

    if (bestScore >= 0.5) {
      result[bestRegion].push(mm);
      console.log(`✅ "${mm}" → ${bestRegion} (score ${bestScore.toFixed(2)})`);
    } else {
      unmatched.push(mm);
      console.warn(`❌ "${mm}" unmatched (best score ${bestScore.toFixed(2)})`);
    }
  });

  if (unmatched.length) {
    console.warn("⚠️ Unmatched MMs (not assigned to any region):", unmatched);
  }

  return result;
}

function renderBangaloreOverview() {
  const containerPrivate = document.getElementById("bangaloreOverviewContainer");
  const containerAll     = document.getElementById("bangaloreOverviewAllContainer");
  const containerPublic  = document.getElementById("bangaloreOverviewPublicContainer");

  const regionMaps = buildRegionMaps();

  if (containerPrivate) containerPrivate.innerHTML = buildBangaloreOverviewHTML("Private",  regionMaps);
  if (containerAll)     containerAll.innerHTML     = buildBangaloreOverviewHTML("All",      regionMaps);
  if (containerPublic)  containerPublic.innerHTML  = buildBangaloreOverviewHTML("Public",   regionMaps);
}

function buildBangaloreOverviewHTML(propertyType, regionMaps) {
  const activeData = allData.filter(r => (r["App status"] || "").trim() === "Active");
  const publicActive  = activeData.filter(r => (r["Property"] || "") === "Public");
  const privateActive = activeData.filter(r => (r["Property"] || "") === "Private");

  const WASHROOM_TYPES = ["Resting + Washroom", "Washroom"];
  const RESTING_TYPES  = ["Resting + Washroom", "Resting"];

  const allHoodNMs = [...new Set(hoods.map(h => h.nano_market).filter(Boolean))];
  const allHoodMMs = [...new Set(hoods.map(h => h.micro_market).filter(Boolean))];

  function uniqueByName(rows) {
    const seen = new Set();
    return rows.filter(r => {
      const name = (r["Name of the property"] || "").trim().toLowerCase();
      if (!name || seen.has(name)) return false;
      seen.add(name);
      return true;
    });
  }

  // Select base dataset by property type
  const scopedActive = propertyType === "Private" ? privateActive
                     : propertyType === "Public"  ? publicActive
                     : activeData;

  const scopedUniq    = uniqueByName(scopedActive);
  const allDataUniq   = uniqueByName(allData);
  const privUniq      = uniqueByName(privateActive);
  const pubUniq       = uniqueByName(publicActive);

  function nmsForMMs(mmList) {
    return [...new Set(
      hoods.filter(h => mmList.includes(h.micro_market)).map(h => h.nano_market).filter(Boolean)
    )];
  }

  // NM/MM helpers scoped to the chosen property type
  function nmHasActive(nm)        { return scopedUniq.some(r => r.NM === nm); }
  function nmHasWashroom(nm)      { return scopedUniq.some(r => r.NM === nm && WASHROOM_TYPES.includes(r["Closure type"] || "")); }
  function nmHasResting(nm)       { return scopedUniq.some(r => r.NM === nm && RESTING_TYPES.includes(r["Closure type"] || "")); }
  function nmHasPrivateActive(nm) { return privUniq.some(r => r.NM === nm); }
  function nmHasAnyLead(nm)       { return allDataUniq.some(r => r.NM === nm); }
  function mmHasPrivateActive(mm) { return privUniq.some(r => r.MM === mm); }
  function mmHasPublicActive(mm)  { return pubUniq.some(r => r.MM === mm); }
  function mmHasAnyLead(mm)       { return allDataUniq.some(r => r.MM === mm); }

  const pct = (n, d) => d ? `${(n / d * 100).toFixed(1)}%` : "0%";
  const fmt = (n, d) => `<b>${n}</b> <span style="color:#888;font-size:11px">(${pct(n,d)})</span>`;

  function isToday(val) {
    if (!val) return false;
    const d = parseTimestamp(val);
    if (!d) return false;
    const now = new Date();
    return d.getFullYear() === now.getFullYear() &&
           d.getMonth()    === now.getMonth()    &&
           d.getDate()     === now.getDate();
  }

  // Launches — always Private
  const launchedToday = privUniq.filter(r => isToday(r["Launch date"])).length;
  const launchedTotal = privUniq.length;

  // Region stats
  const blrMMs     = [...new Set([...regionMaps["Mid Belt"], ...regionMaps["North"], ...regionMaps["South"]])];
  const blrNMs     = nmsForMMs(blrMMs);
  const blrTotal   = blrNMs.length;

  const blrWithActive   = blrNMs.filter(nmHasActive);
  const blrWithWashroom = blrNMs.filter(nmHasWashroom);
  const blrWithResting  = blrNMs.filter(nmHasResting);
  const blrWithBoth     = blrNMs.filter(nm => nmHasWashroom(nm) && nmHasResting(nm));

  function regionRow(regionName) {
    const mms       = regionMaps[regionName] || [];
    const regionNMs = nmsForMMs(mms);
    const withActive = regionNMs.filter(nmHasActive);
    return `
      <tr>
        <td style="font-weight:600;color:#34495e">${escHtml(regionName)} Bangalore</td>
        <td style="text-align:center">${mms.length} MMs → <b>${regionNMs.length}</b> NMs</td>
        <td style="text-align:center">${fmt(withActive.length, regionNMs.length)}</td>
      </tr>`;
  }

  // Pipeline penetration — always Private vs any lead
  const nmsWithPrivateActive = allHoodNMs.filter(nmHasPrivateActive).length;
  const nmsWithAnyLead       = allHoodNMs.filter(nmHasAnyLead).length;
  const mmsWithPrivateActive = allHoodMMs.filter(mmHasPrivateActive).length;
  const mmsWithAnyLead       = allHoodMMs.filter(mmHasAnyLead).length;

  const sourceNote = `<span style="background:#fff3e0;color:#e65100;font-size:10px;padding:2px 7px;border-radius:10px;font-weight:600;margin-left:8px">fuzzy MM match — check console for details</span>`;

  const coverageLabel = propertyType === "All" ? "All Active" : `${propertyType} Active`;
  const nmColLabel    = `NMs with ${coverageLabel} Property`;

  const tile = (icon, label, value, sub, color = "#2c3e50") => `
    <div style="background:#fff;border:1px solid #e8e8e8;border-radius:12px;padding:14px 18px;
                border-left:4px solid ${color};min-width:160px;flex:1">
      <div style="font-size:11px;font-weight:700;color:#888;text-transform:uppercase;letter-spacing:0.4px;margin-bottom:6px">${icon} ${escHtml(label)}</div>
      <div style="font-size:22px;font-weight:800;color:${color}">${value}</div>
      ${sub ? `<div style="font-size:11px;color:#999;margin-top:3px">${sub}</div>` : ""}
    </div>`;

  return `
    <div style="margin-bottom:20px">
      <div style="font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:0.5px;color:#555;margin-bottom:10px;padding-bottom:6px;border-bottom:2px solid #eee">
        📍 NM Coverage — ${coverageLabel} Properties ${sourceNote}
      </div>
      <div style="overflow-x:auto">
        <table class="summary-table" style="max-width:620px">
          <thead><tr>
            <th style="text-align:left">Region</th>
            <th>MMs / NMs</th>
            <th>${escHtml(nmColLabel)}</th>
          </tr></thead>
          <tbody>
            <tr style="background:#f0f4ff">
              <td style="font-weight:700;color:#1a237e">🏙️ Bangalore (all)</td>
              <td style="text-align:center">${blrMMs.length} MMs → <b>${blrTotal}</b> NMs</td>
              <td style="text-align:center">${fmt(blrWithActive.length, blrTotal)}</td>
            </tr>
            ${["Mid Belt","North","South"].map(r => regionRow(r)).join("")}
          </tbody>
        </table>
      </div>
    </div>

    <div style="margin-bottom:20px">
      <div style="font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:0.5px;color:#555;margin-bottom:10px;padding-bottom:6px;border-bottom:2px solid #eee">
        🚿 Bangalore NM — Closure Coverage · ${coverageLabel} · (${blrTotal} NMs)
      </div>
      <div style="display:flex;gap:12px;flex-wrap:wrap">
        ${tile("🚿", "NMs with Washroom",  `${blrWithWashroom.length} / ${blrTotal}`, pct(blrWithWashroom.length, blrTotal), "#2196f3")}
        ${tile("🛋️", "NMs with Resting",   `${blrWithResting.length} / ${blrTotal}`,  pct(blrWithResting.length, blrTotal),  "#4caf50")}
        ${tile("✅", "NMs with Both",      `${blrWithBoth.length} / ${blrTotal}`,      pct(blrWithBoth.length, blrTotal),     "#9c27b0")}
      </div>
    </div>

    <div style="margin-bottom:20px">
      <div style="font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:0.5px;color:#555;margin-bottom:10px;padding-bottom:6px;border-bottom:2px solid #eee">
        🚀 Private Property Launches
      </div>
      <div style="display:flex;gap:12px;flex-wrap:wrap">
        ${tile("📅", "Launched Today", launchedToday, "Private · App Status = Active · Launch date = today", "#e67e22")}
        ${tile("📦", "Launched Total", launchedTotal, "All Private · App Status = Active",                   "#27ae60")}
      </div>
    </div>

    <div>
      <div style="font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:0.5px;color:#555;margin-bottom:10px;padding-bottom:6px;border-bottom:2px solid #eee">
        📊 Pipeline Penetration (Private Active vs Any Lead)
      </div>
      <div style="display:flex;gap:12px;flex-wrap:wrap">
        ${tile("🗺️", "NMs: Private Active / Any Lead",
          `${nmsWithPrivateActive} / ${nmsWithAnyLead}`,
          pct(nmsWithPrivateActive, nmsWithAnyLead), "#0288d1")}
        ${tile("🗺️", "MMs: Private Active / Any Lead",
          `${mmsWithPrivateActive} / ${mmsWithAnyLead}`,
          pct(mmsWithPrivateActive, mmsWithAnyLead), "#00838f")}
      </div>
    </div>`;
}

// ============================================================
// NM / MM LEVEL SUMMARY (Tasks 3 & 4)
// ============================================================
function renderNmMmSummary() {
  const container = document.getElementById("nmMmSummaryContainer");
  if (!container) return;

  // Read Property filter (Public / Private / all)
  const propFilter = document.getElementById("nmMmPropertyFilter")?.value || "";

  const FIXED_CATEGORIES = [
    "Ladies PG", "Shop", "Restaurant", "Gated community",
    "Independent Builder floor", "Bus Stop", "Park",
    "Petrol Pump", "Public Washroom", "Other"
  ];

  const normalizeCategory = (cat) => cat === "Apartment" ? "Independent Builder floor" : cat;

  // Base active data — optionally filtered by Property type
  let activeData = allData.filter(r => (r["App status"] || "").trim() === "Active");
  if (propFilter) activeData = activeData.filter(r => (r["Property"] || "") === propFilter);

  // Helper: build category count table per groupKey — uses all hoods as source of truth
  function buildGroupTable(groupKey, tableId) {
    const hoodField = groupKey === "NM" ? "nano_market" : "micro_market";
    const groups = [...new Set(hoods.map(h => h[hoodField]).filter(Boolean))].sort();
    if (!groups.length) return `<p style="color:#aaa;font-size:12px">No hood data.</p>`;

    const rows = groups.map(g => {
      const gRows = activeData.filter(r => r[groupKey] === g);
      const cats = {};
      FIXED_CATEGORIES.forEach(c => {
        cats[c] = gRows.filter(r => normalizeCategory(r.Category) === c).length;
      });
      return `<tr data-group="${escHtml(g)}">
        <td>${escHtml(g)}</td>
        <td style="font-weight:700">${gRows.length}</td>
        ${FIXED_CATEGORIES.map(c => `<td>${cats[c] || 0}</td>`).join("")}
      </tr>`;
    }).join("");

    return `
      <div class="nm-table-wrap" id="${tableId}_wrap">
        <div style="overflow-x:auto;max-height:400px;overflow-y:auto" id="${tableId}_scroll">
          <table class="summary-table" id="${tableId}" style="min-width:700px">
            <thead><tr>
              <th>${groupKey}</th><th>Total Active</th>
              ${FIXED_CATEGORIES.map(c => `<th>${escHtml(c)}</th>`).join("")}
            </tr></thead>
            <tbody>${rows}</tbody>
          </table>
        </div>
      </div>
      <div style="display:flex;gap:8px;margin-top:6px;align-items:center">
        <button class="summary-dl-btn" onclick="downloadSummaryTable('${tableId}','active_by_${groupKey.toLowerCase()}')">⬇ CSV</button>
        <button class="nm-fullscreen-btn" onclick="openNmMmFullscreen('${tableId}','${groupKey} Level — Active Properties by Category${propFilter ? " (" + propFilter + ")" : ""}')">⛶ Fullscreen</button>
      </div>`;
  }

  // Washroom/Resting stats
  const WASHROOM_TYPES = ["Resting + Washroom", "Washroom"];
  const RESTING_TYPES  = ["Resting + Washroom", "Resting"];

  function buildWrStats(groupKey, label) {
    const hoodField = groupKey === "NM" ? "nano_market" : "micro_market";
    const groups = [...new Set(hoods.map(h => h[hoodField]).filter(Boolean))].sort();
    const total  = groups.length;
    if (!total) return { statsHtml: `<p style="color:#aaa;font-size:12px">No data.</p>`, groups: [] };

    const withWashroom  = groups.filter(g => activeData.some(r => r[groupKey] === g && WASHROOM_TYPES.includes(r["Closure type"] || "")));
    const withResting   = groups.filter(g => activeData.some(r => r[groupKey] === g && RESTING_TYPES.includes(r["Closure type"] || "")));
    const withBoth      = groups.filter(g => withWashroom.includes(g) && withResting.includes(g));
    const noWashroom    = groups.filter(g => !withWashroom.includes(g));
    const noResting     = groups.filter(g => !withResting.includes(g));
    const withNeither   = groups.filter(g => !withWashroom.includes(g) && !withResting.includes(g));

    const pct = (n) => total ? `(${(n / total * 100).toFixed(1)}%)` : "(0%)";

    const cardHtml = (title, count, pctStr, borderColor, bgColor, kind, gKey) =>
      `<div class="wr-card" onclick="showWrHighlight('${kind}','${gKey}',this)" style="border-left:4px solid ${borderColor};background:${bgColor}">
        <div class="wr-card-title">${escHtml(title)}</div>
        <div class="wr-card-value">${count} <span class="wr-card-pct">${pctStr}</span></div>
        <div class="wr-card-label">of ${total} total ${label}s</div>
      </div>`;

    const statsHtml = `
      <div class="washroom-resting-grid">
        ${cardHtml(`${label}s with Washroom`,              withWashroom.length, pct(withWashroom.length), "#2196f3", "#f0f8ff", "withWashroom", groupKey)}
        ${cardHtml(`${label}s with Resting`,               withResting.length,  pct(withResting.length),  "#4caf50", "#f0fff4", "withResting",  groupKey)}
        ${cardHtml(`${label}s with Washroom & Resting`,    withBoth.length,     pct(withBoth.length),     "#9c27b0", "#faf0ff", "withBoth",     groupKey)}
      </div>
      <hr class="wr-divider"/>
      <div class="washroom-resting-grid">
        ${cardHtml(`${label}s without Washroom`,           noWashroom.length,   pct(noWashroom.length),   "#e53935", "#fff5f5", "noWashroom",   groupKey)}
        ${cardHtml(`${label}s without Resting`,            noResting.length,    pct(noResting.length),    "#ff9800", "#fffbf0", "noResting",    groupKey)}
        ${cardHtml(`${label}s without Washroom & Resting`, withNeither.length,  pct(withNeither.length),  "#546e7a", "#f4f6f7", "withNeither",  groupKey)}
      </div>`;

    return { statsHtml, withWashroom, withResting, withBoth, noWashroom, noResting, withNeither, total, groups };
  }

  const nmStats = buildWrStats("NM", "NM");
  const mmStats = buildWrStats("MM", "MM");

  const totalNMs = [...new Set(hoods.map(h => h.nano_market).filter(Boolean))].length;
  const totalMMs = [...new Set(hoods.map(h => h.micro_market).filter(Boolean))].length;

  container.innerHTML = `
    <div style="display:flex;align-items:center;gap:10px;margin-bottom:14px;flex-wrap:wrap">
      <span style="font-size:12px;font-weight:600;color:#555">🔍 Filter by Property:</span>
      <select id="nmMmPropertyFilter" onchange="renderNmMmSummary()"
              style="padding:6px 10px;border-radius:8px;border:1px solid #ddd;font-size:13px;min-width:160px">
        <option value="" ${!propFilter ? "selected" : ""}>All Properties</option>
        <option value="Public"  ${propFilter === "Public"  ? "selected" : ""}>Public</option>
        <option value="Private" ${propFilter === "Private" ? "selected" : ""}>Private</option>
      </select>
      ${propFilter ? `<span style="background:#e8f0fe;color:#3b5bdb;font-size:11px;padding:3px 10px;border-radius:20px;font-weight:600">Showing: ${propFilter}</span>` : ""}
    </div>

    <div class="nm-mm-grid" style="margin-bottom:24px">
      <div class="nm-mm-panel">
        <h4>NM Level — Active Properties by Category</h4>
        ${buildGroupTable("NM", "nmActiveTable")}
      </div>
      <div class="nm-mm-panel">
        <h4>MM Level — Active Properties by Category</h4>
        ${buildGroupTable("MM", "mmActiveTable")}
      </div>
    </div>

    <div class="nm-mm-grid">
      <div class="nm-mm-panel">
        <h4>NM Washroom &amp; Resting Coverage</h4>
        <div style="font-size:11px;color:#888;margin-bottom:10px">Total NMs (from hoods): <b>${totalNMs}</b></div>
        ${nmStats.statsHtml}
      </div>
      <div class="nm-mm-panel">
        <h4>MM Washroom &amp; Resting Coverage</h4>
        <div style="font-size:11px;color:#888;margin-bottom:10px">Total MMs (from hoods): <b>${totalMMs}</b></div>
        ${mmStats.statsHtml}
      </div>
    </div>

    <div id="wrHighlightCard" style="display:none;margin-top:16px" class="wr-highlight-list">
      <h5 id="wrHighlightTitle"></h5>
      <div id="wrHighlightItems"></div>
    </div>`;

  // Store stats for click handler
  window._wrStatsNM = nmStats;
  window._wrStatsMM = mmStats;
}

function showWrHighlight(kind, groupKey, clickedEl) {
  clickedEl.closest(".nm-mm-panel").querySelectorAll(".wr-card").forEach(c => c.classList.remove("active"));
  clickedEl.classList.add("active");

  const stats = groupKey === "NM" ? window._wrStatsNM : window._wrStatsMM;
  const groups = stats[kind] || [];
  const label = groupKey === "NM" ? "NM" : "MM";

  const kindLabel = {
    withWashroom: `${label}s with Washroom`,
    withResting:  `${label}s with Resting`,
    withBoth:     `${label}s with Washroom & Resting`,
    noWashroom:   `${label}s without Washroom`,
    noResting:    `${label}s without Resting`,
    withNeither:  `${label}s without any Washroom or Resting`
  }[kind] || kind;

  const card = document.getElementById("wrHighlightCard");
  document.getElementById("wrHighlightTitle").textContent = `${kindLabel} (${groups.length})`;
  document.getElementById("wrHighlightItems").innerHTML = groups.length
    ? groups.map(g => `<div class="wr-highlight-item">📍 ${escHtml(g)}</div>`).join("")
    : `<div style="color:#aaa">None</div>`;
  card.style.display = "block";
  card.scrollIntoView({ behavior: "smooth", block: "nearest" });
}

function openNmMmFullscreen(tableId, title) {
  const srcTable = document.getElementById(tableId);
  if (!srcTable) return;

  // Create overlay
  const overlay = document.createElement("div");
  overlay.id = "nmMmFullscreenOverlay";
  overlay.style.cssText = `
    position:fixed;inset:0;background:rgba(0,0,0,0.6);z-index:99999;
    display:flex;align-items:flex-start;justify-content:center;padding:24px;box-sizing:border-box;
  `;

  overlay.innerHTML = `
    <div style="background:#fff;border-radius:16px;width:100%;max-width:1400px;max-height:calc(100vh - 48px);
                display:flex;flex-direction:column;box-shadow:0 8px 40px rgba(0,0,0,0.25)">
      <div style="display:flex;justify-content:space-between;align-items:center;
                  padding:14px 20px;border-bottom:1px solid #eee;flex-shrink:0">
        <h3 style="margin:0;font-size:15px;color:#333">${escHtml(title)}</h3>
        <div style="display:flex;gap:8px">
          <button class="summary-dl-btn" onclick="downloadSummaryTable('${tableId}_fs','${tableId}_fullscreen')">⬇ CSV</button>
          <button onclick="document.getElementById('nmMmFullscreenOverlay').remove()"
                  style="background:#e74c3c;color:#fff;border:none;border-radius:8px;
                         padding:6px 14px;cursor:pointer;font-size:13px;font-weight:600">✕ Close</button>
        </div>
      </div>
      <div style="overflow:auto;flex:1;padding:16px">
        <div id="nmMmFsTableWrap"></div>
      </div>
    </div>`;

  document.body.appendChild(overlay);

  // Clone table into fullscreen with new id
  const clone = srcTable.cloneNode(true);
  clone.id = tableId + "_fs";
  clone.style.minWidth = srcTable.style.minWidth;
  document.getElementById("nmMmFsTableWrap").appendChild(clone);

  // Close on backdrop click
  overlay.addEventListener("click", e => { if (e.target === overlay) overlay.remove(); });
  // Close on Escape
  const escHandler = (e) => { if (e.key === "Escape") { overlay.remove(); document.removeEventListener("keydown", escHandler); } };
  document.addEventListener("keydown", escHandler);
}

// ============================================================
// REMINDER FOR PROPERTY CLOSURES (Task 1)
// ============================================================
function getReminderStatus(row) {
  const leadStatus  = (row["Lead Status"]  || "").trim();
  const appStatus   = (row["App status"]   || "").trim();
  const finalStatus = (row["Final Status"] || "").trim();

  if (finalStatus === "Dropped off") return null;

  const signPendingLeadStatuses = [
    "5. Follow up required",
    "3. Owner's confirmation pending",
    "2. Owner conversation pending"
  ];

  if (signPendingLeadStatuses.includes(leadStatus)) return "Sign Pending";
  if (appStatus === "Inactive") return "Chairs and Poster Pending";
  return null;
}

function getReminderData() {
  const from = document.getElementById("reminderDateFrom")?.value;
  const to   = document.getElementById("reminderDateTo")?.value;
  const statusFilter = document.getElementById("reminderStatusFilter")?.value || "";

  const fromDate = from ? new Date(from + "T00:00:00") : null;
  const toDate   = to   ? new Date(to   + "T23:59:59") : null;

  return allData.filter(row => {
    const finalStatus = (row["Final Status"] || "").trim();
    if (finalStatus === "Dropped off") return false;

    const reminderStatus = getReminderStatus(row);
    if (!reminderStatus) return false;
    if (statusFilter && reminderStatus !== statusFilter) return false;

    if (fromDate || toDate) {
      const ts = parseTimestamp(row["Timestamp"]);
      if (!ts) return false;
      if (fromDate && ts < fromDate) return false;
      if (toDate   && ts > toDate)   return false;
    }

    return true;
  }).map(row => ({ ...row, _reminderStatus: getReminderStatus(row) }));
}

function renderReminderTable() {
  const table   = document.getElementById("reminderTable");
  const countEl = document.getElementById("reminderRowCount");
  if (!table) return;

  const data = getReminderData();
  if (countEl) countEl.textContent = `${data.length} rows`;

  const COLS = [
    "_reminderStatus", "Lead From", "Timestamp", "Email Address",
    "MM", "NM", "Name of the property",
    "Owner Contact Name", "Owner Contact Number", "Owner Designation",
    "Lat", "Long"
  ];

  if (!data.length) {
    table.innerHTML = `<tr><td colspan="99" style="text-align:center;color:#aaa;padding:16px">No reminders found for the selected filters.</td></tr>`;
    return;
  }

  const availableCols = COLS.filter(c => c === "_reminderStatus" || data[0].hasOwnProperty(c));

  const headerRow = availableCols.map(c => {
    const label = c === "_reminderStatus" ? "Reminder Status" : c;
    return `<th>${escHtml(label)}</th>`;
  }).join("");

  const bodyRows = data.map(row => {
    return `<tr>${availableCols.map(c => {
      if (c === "_reminderStatus") {
        const cls = row._reminderStatus === "Sign Pending" ? "reminder-sign" : "reminder-chairs";
        return `<td><span class="${cls}">${escHtml(row._reminderStatus)}</span></td>`;
      }
      const val = row[c] != null ? row[c] : "";
      const display = c === "Timestamp" ? (formatTsDisplay(val) || escHtml(String(val))) : escHtml(String(val));
      return `<td>${display}</td>`;
    }).join("")}</tr>`;
  }).join("");

  table.innerHTML =
    `<thead><tr>${headerRow}</tr></thead>` +
    `<tbody>${bodyRows}</tbody>`;
}

function applyReminderFilters() { renderReminderTable(); }

function clearReminderFilters() {
  ["reminderDateFrom","reminderDateTo","reminderStatusFilter"].forEach(id => {
    const el = document.getElementById(id); if (el) el.value = "";
  });
  renderReminderTable();
}

function downloadReminderCSV() {
  const table = document.getElementById("reminderTable");
  if (!table) return;
  const rows = [...table.querySelectorAll("thead tr, tbody tr")].map(tr =>
    [...tr.querySelectorAll("th,td")].map(td => `"${td.innerText.replace(/"/g,'""')}"`).join(",")
  );
  const dateTag = new Date().toISOString().slice(0,10);
  downloadBlob(rows.join("\n"), `reminder_closures_${dateTag}.csv`, "text/csv");
}

// ============================================================
// UTILITIES
// ============================================================
function isEmpty(val) { return !val || val.toString().trim() === "" || val === "NA"; }

function escHtml(str) {
  return String(str || "").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
}