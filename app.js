let map;
let currentRow = null;
let allData = [];
let hoods = [];
let markers = [];
let activeFilters = {};

// ── Add-Point mode ──────────────────────────────────────────
let addPointMode = false;
let addPointMarker = null;   // temporary crosshair marker while picking
let hoodLayers = [];         // { layer, nm, mm } — all drawn hood polygons

console.log("🚀 App initializing...");
init();

async function init() {
  console.log("🗺️ init() called");

  map = L.map('map').setView([12.9, 77.65], 12);
  console.log("🗺️ Map created");

  L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png')
    .addTo(map);

  setTimeout(() => map.invalidateSize(), 300);

  console.log("📦 Fetching hoods.json...");
  hoods = await fetch("hoods.json").then(r => r.json());
  console.log(`📦 hoods.json loaded — ${hoods.length} hoods`);

  drawHoods();

  await loadData();

  setInterval(loadData, 30000);

  initMapSearch();

  // ── Map-click handler for "Add Point" mode ──
  map.on("click", function (e) {
    if (!addPointMode) return;
    const { lat, lng } = e.latlng;
    console.log(`📍 Map clicked in Add-Point mode: [${lat}, ${lng}]`);
    openAddPointModal(lat, lng);
  });
}

// ============================================================
// DATA LOADING
// ============================================================
async function loadData() {
  const url = CONFIG.API_URL + "?t=" + Date.now();
  console.log("📡 loadData() —", url);
  try {
    const res = await fetch(url);
    const text = await res.text();
    const data = JSON.parse(text);
    allData = data;
    console.log(`✅ ${allData.length} rows loaded`);

    if (Object.values(activeFilters).some(v => v)) {
      filterAndRender();
    } else {
      renderMarkers();
    }
    renderSheetPreview(allData);
  } catch (err) {
    console.error("❌ Fetch failed:", err);
  }
  populateFilters();
}

// ============================================================
// RENDER MARKERS
// ============================================================
function renderMarkers() {
  console.log("📍 renderMarkers()");
  markers.forEach(m => map.removeLayer(m));
  markers = [];
  let skipped = 0;

  allData.forEach((row, i) => {
    const lat = parseFloat(row.Lat);
    const lng = parseFloat(row.Long);
    if (!isNaN(lat) && !isNaN(lng)) {
      if (!row.NM || !row.MM) {
        const hood = assignHood({ lat, lng });
        if (hood) {
          row.NM = hood.nano_market;
          row.MM = hood.micro_market;
          row["NM Id"] = hood.hood_id;
          if (row._rowIndex) {
            fetch(CONFIG.API_URL, { method: "POST", body: JSON.stringify(row) })
              .then(() => console.log(`✅ Auto-saved _rowIndex: ${row._rowIndex}`))
              .catch(err => console.error(`❌ Auto-save failed`, err));
          }
        }
      }
      const marker = L.marker([lat, lng], { icon: getCategoryIcon(row.Category) })
        .addTo(map)
        .bindPopup(`<b>${row.Name || "No Name"}</b><br>${row.Category || ""}<br>NM: ${row.NM || "-"}<br>MM: ${row.MM || "-"}`)
        .on('click', () => showDetails(row));
      markers.push(marker);
    } else {
      skipped++;
    }
  });
  console.log(`📍 ${markers.length} markers, ${skipped} skipped`);
}

// ============================================================
// FILTERS — only matching rows stay on map
// ============================================================
function applyFilters() {
  console.log("🔽 applyFilters()");
  activeFilters = {
    Category:      document.getElementById("filterCategory").value,
    Property:      document.getElementById("filterProperty").value,
    "App status":  document.getElementById("filterAppStatus").value,
    "Lead Status": document.getElementById("filterLeadStatus").value,
    NM:            document.getElementById("filterNM").value,
    MM:            document.getElementById("filterMM").value
  };
  console.log("🔽 Filters:", activeFilters);
  filterAndRender();
}

function filterAndRender() {
  const filtered = allData.filter(row =>
    Object.keys(activeFilters).every(key => !activeFilters[key] || row[key] === activeFilters[key])
  );
  console.log(`🔽 ${filtered.length}/${allData.length} rows match filters`);
  updateHoodVisibility();
  renderFilteredMarkers(filtered);
}

function renderFilteredMarkers(data) {
  markers.forEach(m => map.removeLayer(m));
  markers = [];
  const bounds = [];

  data.forEach((row, i) => {
    const lat = parseFloat(row.Lat);
    const lng = parseFloat(row.Long);
    if (!isNaN(lat) && !isNaN(lng)) {
      const marker = L.marker([lat, lng], { icon: getCategoryIcon(row.Category) })
        .addTo(map)
        .bindPopup(`<b>${row.Name || "No Name"}</b><br>${row.Category || ""}<br>NM: ${row.NM || "-"}<br>MM: ${row.MM || "-"}`)
        .on('click', () => showDetails(row));
      markers.push(marker);
      bounds.push([lat, lng]);
    }
  });

  if (bounds.length) map.fitBounds(bounds);
  console.log(`📍 renderFilteredMarkers: ${markers.length} markers`);
}

function clearFilters() {
  activeFilters = {};
  document.querySelectorAll("select").forEach(s => s.value = "");
  updateHoodVisibility();
  renderMarkers();
}

function populateFilters() {
  const fields = [
    { key: "Category",    id: "filterCategory"   },
    { key: "Property",    id: "filterProperty"   },
    { key: "App status",  id: "filterAppStatus"  },
    { key: "Lead Status", id: "filterLeadStatus" },
    { key: "NM",          id: "filterNM"         },
    { key: "MM",          id: "filterMM"         }
  ];
  fields.forEach(f => {
    const select = document.getElementById(f.id);
    if (!select) return;
    const current = select.value;
    const values = [...new Set(allData.map(r => r[f.key]).filter(Boolean))].sort();
    select.innerHTML = `<option value="">${f.key}</option>` +
      values.map(v => `<option value="${v}">${v}</option>`).join("");
    select.value = current;
  });
}

// ============================================================
// FEATURE: ADD POINT BY CLICKING MAP
// ============================================================

// Fields shown in the Add-Point modal — Lat/Long auto-filled & locked; all fields optional
const ADD_POINT_FIELDS = [
  { key: "Name",           label: "Name",            type: "text"  },
  { key: "Category",       label: "Category",        type: "text"  },
  { key: "Sub Category",   label: "Sub Category",    type: "text"  },
  { key: "Road",           label: "Road",            type: "text"  },
  { key: "Property",       label: "Property",        type: "text"  },
  { key: "App status",     label: "App Status",      type: "text"  },
  { key: "Lead Status",    label: "Lead Status",     type: "text"  },
  { key: "Final Status",   label: "Final Status",    type: "text"  },
  { key: "location",       label: "Maps Link",       type: "url"   },
  { key: "Contact Name",   label: "Contact Name",    type: "text"  },
  { key: "Contact number", label: "Contact Number",  type: "tel"   },
  { key: "Comment",        label: "Comment",         type: "text"  },
  { key: "Restroom ID",    label: "Restroom ID",     type: "text"  },
];

function toggleAddPointMode() {
  addPointMode = !addPointMode;
  const btn = document.getElementById("btnAddPoint");
  if (addPointMode) {
    btn.textContent = "❌ Cancel Add Point";
    btn.style.background = "#c0392b";
    map.getContainer().style.cursor = "crosshair";
    console.log("📍 Add-Point mode ON — click map to place");
  } else {
    btn.textContent = "➕ Add Point";
    btn.style.background = "";
    map.getContainer().style.cursor = "";
    // remove temporary marker if any
    if (addPointMarker) { map.removeLayer(addPointMarker); addPointMarker = null; }
    console.log("📍 Add-Point mode OFF");
  }
}

function openAddPointModal(lat, lng) {
  // place a temporary crosshair marker
  if (addPointMarker) map.removeLayer(addPointMarker);
  addPointMarker = L.marker([lat, lng], {
    icon: L.divIcon({ className: "custom-icon", html: `<div style="font-size:24px">📌</div>` })
  }).addTo(map);

  const container = document.getElementById("addPointFields");
  if (!container) {
    console.error("❌ #addPointFields not found in DOM");
    return;
  }

  container.innerHTML = `
    <!-- Lat/Long shown but locked -->
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
      <input type="${f.type}" id="ap_${f.key.replace(/[\s.]/g,'_')}"
             placeholder="${f.label}${f.required ? '' : ' (optional)'}" />
    </div>
  `).join("");

  document.getElementById("addPointModal").style.display = "flex";
}

function closeAddPointModal() {
  document.getElementById("addPointModal").style.display = "none";
  if (addPointMarker) { map.removeLayer(addPointMarker); addPointMarker = null; }
  // turn off add-point mode after placing
  if (addPointMode) toggleAddPointMode();
}

async function submitAddPoint() {
  const lat = document.getElementById("ap_Lat").value;
  const lng = document.getElementById("ap_Long").value;

  // New rows have no _rowIndex — Apps Script will appendRow
  const newRow = {
    "Lat":  parseFloat(lat),
    "Long": parseFloat(lng)
  };

  // collect optional fields
  ADD_POINT_FIELDS.forEach(f => {
    const el = document.getElementById("ap_" + f.key.replace(/[\s.]/g,'_'));
    if (el && el.value.trim()) {
      newRow[f.key] = f.type === "number" ? parseFloat(el.value) : el.value.trim();
    }
  });

  // auto-assign NM/MM
  const hood = assignHood({ lat: parseFloat(lat), lng: parseFloat(lng) });
  if (hood) {
    newRow.NM = hood.nano_market;
    newRow.MM = hood.micro_market;
    newRow["NM Id"] = hood.hood_id;
    console.log(`✅ Auto-assigned NM: ${newRow.NM}, MM: ${newRow.MM}`);
  }

  console.log("📤 Submitting new point:", newRow);

  try {
    const res = await fetch(CONFIG.API_URL, { method: "POST", body: JSON.stringify(newRow) });
    const txt = await res.text();
    const json = JSON.parse(txt);
    if (json.success) {
      alert(`✅ Point added\nNM: ${newRow.NM || "-"}, MM: ${newRow.MM || "-"}`);
      closeAddPointModal();
      loadData();
    } else {
      alert("❌ Failed: " + txt);
    }
  } catch (err) {
    console.error("❌ submitAddPoint failed:", err);
    alert("❌ Error: " + err.message);
  }
}

// ============================================================
// FEATURE: DOWNLOAD KML / CSV (WKT) — filtered output only
// ============================================================

/** Returns the currently-filtered dataset (or all data if no filters active) */
function getFilteredData() {
  const hasFilter = Object.values(activeFilters).some(v => v);
  if (!hasFilter) return allData;
  return allData.filter(row =>
    Object.keys(activeFilters).every(key => !activeFilters[key] || row[key] === activeFilters[key])
  );
}

/** Returns unique NM values present in filtered data */
function getFilteredNMs() {
  const data = getFilteredData();
  return [...new Set(data.map(r => r.NM).filter(Boolean))];
}

/** Returns unique MM values present in filtered data */
function getFilteredMMs() {
  const data = getFilteredData();
  return [...new Set(data.map(r => r.MM).filter(Boolean))];
}

/** Returns hood objects matching a list of NM names */
function hoodsByNM(nmList) {
  return hoods.filter(h => nmList.includes(h.nano_market));
}

/** Returns hood objects matching a list of MM names */
function hoodsByMM(mmList) {
  return hoods.filter(h => mmList.includes(h.micro_market));
}

// ── KML helpers ─────────────────────────────────────────────

function geojsonCoordToKmlRing(coords) {
  // coords: array of [lng, lat] pairs
  return coords.map(c => `${c[0]},${c[1]},0`).join(" ");
}

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
      const outer = poly[0];
      const inner = poly.slice(1);
      return `<Polygon>
        <outerBoundaryIs><LinearRing><coordinates>${geojsonCoordToKmlRing(outer)}</coordinates></LinearRing></outerBoundaryIs>
        ${inner.map(r => `<innerBoundaryIs><LinearRing><coordinates>${geojsonCoordToKmlRing(r)}</coordinates></LinearRing></innerBoundaryIs>`).join("")}
      </Polygon>`;
    }).join("")}</MultiGeometry>`;
  }
  return "";
}

function hoodsToKml(hoodList, layerName, color = "7f0000ff") {
  const placemarks = hoodList.map(h => `
  <Placemark>
    <name>${escXml(h.nano_market || h.micro_market || h.hood_id)}</name>
    <description><![CDATA[NM: ${h.nano_market || ""}<br>MM: ${h.micro_market || ""}<br>ID: ${h.hood_id || ""}]]></description>
    <Style><PolyStyle><color>${color}</color><outline>1</outline></PolyStyle></Style>
    ${geometryToKmlGeometry(h.geometry)}
  </Placemark>`).join("\n");

  return `<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2">
<Document>
  <name>${escXml(layerName)}</name>
${placemarks}
</Document>
</kml>`;
}

function pointsToKml(data, layerName) {
  const placemarks = data
    .filter(row => !isNaN(parseFloat(row.Lat)) && !isNaN(parseFloat(row.Long)))
    .map(row => `
  <Placemark>
    <name>${escXml(row.Name || "Point")}</name>
    <description><![CDATA[
      Category: ${row.Category || ""}<br>
      NM: ${row.NM || ""}<br>
      MM: ${row.MM || ""}<br>
      Road: ${row.Road || ""}<br>
      Status: ${row["Final Status"] || ""}
    ]]></description>
    <Point><coordinates>${parseFloat(row.Long)},${parseFloat(row.Lat)},0</coordinates></Point>
  </Placemark>`).join("\n");

  return `<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2">
<Document>
  <name>${escXml(layerName)}</name>
${placemarks}
</Document>
</kml>`;
}

function escXml(str) {
  return String(str || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

// ── WKT CSV helpers ──────────────────────────────────────────

function coordsToWktPolygon(coordinates) {
  if (!coordinates || !coordinates.length) return "";
  const ring = coordinates[0].map(c => `${c[0]} ${c[1]}`).join(", ");
  return `POLYGON((${ring}))`;
}

function geometryToWkt(geometry) {
  if (!geometry) return "";
  if (geometry.type === "Polygon") return coordsToWktPolygon(geometry.coordinates);
  if (geometry.type === "MultiPolygon") {
    const parts = geometry.coordinates.map(poly => {
      const ring = poly[0].map(c => `${c[0]} ${c[1]}`).join(", ");
      return `((${ring}))`;
    });
    return `MULTIPOLYGON(${parts.join(", ")})`;
  }
  return "";
}

function hoodsToCsvWkt(hoodList, nameField) {
  const headers = ["WKT", "name", "nm", "mm", "hood_id"];
  const rows = hoodList.map(h => [
    `"${geometryToWkt(h.geometry)}"`,
    `"${h[nameField] || h.nano_market || h.micro_market || ""}"`,
    `"${h.nano_market || ""}"`,
    `"${h.micro_market || ""}"`,
    `"${h.hood_id || ""}"`
  ].join(","));
  return [headers.join(","), ...rows].join("\n");
}

function pointsToCsvWkt(data) {
  const headers = ["WKT", "name", "category", "nm", "mm", "road", "final_status", "lat", "long"];
  const rows = data
    .filter(row => !isNaN(parseFloat(row.Lat)) && !isNaN(parseFloat(row.Long)))
    .map(row => {
      const lat = parseFloat(row.Lat);
      const lng = parseFloat(row.Long);
      return [
        `"POINT(${lng} ${lat})"`,
        `"${(row.Name || "").replace(/"/g,'""')}"`,
        `"${row.Category || ""}"`,
        `"${row.NM || ""}"`,
        `"${row.MM || ""}"`,
        `"${(row.Road || "").replace(/"/g,'""')}"`,
        `"${row["Final Status"] || ""}"`,
        lat,
        lng
      ].join(",");
    });
  return [headers.join(","), ...rows].join("\n");
}

// ── Download trigger ─────────────────────────────────────────

function downloadBlob(content, filename, mimeType) {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

function downloadLayerKML(type) {
  const filteredData = getFilteredData();
  const label = activeFilters.NM || activeFilters.MM || "filtered";

  if (type === "nm") {
    const nmList = getFilteredNMs();
    const hoodList = hoodsByNM(nmList);
    if (!hoodList.length) { alert("No NM hoods found for current filter."); return; }
    downloadBlob(hoodsToKml(hoodList, `NM Layer — ${label}`, "7f0000ff"), `nm_layer_${label}.kml`, "application/vnd.google-earth.kml+xml");
    console.log(`📥 NM KML downloaded — ${hoodList.length} hoods`);

  } else if (type === "mm") {
    const mmList = getFilteredMMs();
    // de-duplicate: one polygon per MM (merge by first match)
    const seen = new Set();
    const hoodList = hoods.filter(h => {
      if (!mmList.includes(h.micro_market)) return false;
      if (seen.has(h.micro_market)) return false;
      seen.add(h.micro_market);
      return true;
    });
    if (!hoodList.length) { alert("No MM hoods found for current filter."); return; }
    downloadBlob(hoodsToKml(hoodList, `MM Layer — ${label}`, "7fff0000"), `mm_layer_${label}.kml`, "application/vnd.google-earth.kml+xml");
    console.log(`📥 MM KML downloaded — ${hoodList.length} hoods`);

  } else if (type === "points") {
    if (!filteredData.length) { alert("No data points to export."); return; }
    downloadBlob(pointsToKml(filteredData, `Data Points — ${label}`), `points_${label}.kml`, "application/vnd.google-earth.kml+xml");
    console.log(`📥 Points KML downloaded — ${filteredData.length} rows`);
  }
}

function downloadLayerCSV(type) {
  const filteredData = getFilteredData();
  const label = activeFilters.NM || activeFilters.MM || "filtered";

  if (type === "nm") {
    const hoodList = hoodsByNM(getFilteredNMs());
    if (!hoodList.length) { alert("No NM hoods found for current filter."); return; }
    downloadBlob(hoodsToCsvWkt(hoodList, "nano_market"), `nm_layer_${label}.csv`, "text/csv");
    console.log(`📥 NM CSV (WKT) downloaded — ${hoodList.length} hoods`);

  } else if (type === "mm") {
    const mmList = getFilteredMMs();
    const seen = new Set();
    const hoodList = hoods.filter(h => {
      if (!mmList.includes(h.micro_market) || seen.has(h.micro_market)) return false;
      seen.add(h.micro_market);
      return true;
    });
    if (!hoodList.length) { alert("No MM hoods found for current filter."); return; }
    downloadBlob(hoodsToCsvWkt(hoodList, "micro_market"), `mm_layer_${label}.csv`, "text/csv");
    console.log(`📥 MM CSV (WKT) downloaded — ${hoodList.length} hoods`);

  } else if (type === "points") {
    if (!filteredData.length) { alert("No data points to export."); return; }
    downloadBlob(pointsToCsvWkt(filteredData), `points_${label}.csv`, "text/csv");
    console.log(`📥 Points CSV (WKT) downloaded — ${filteredData.length} rows`);
  }
}

// ============================================================
// URL RESOLUTION (unchanged)
// ============================================================
async function resolveGoogleMapsCoords(url) {
  const directMatch = url.match(/@(-?\d+\.\d+),(-?\d+\.\d+)/);
  if (directMatch) return { lat: parseFloat(directMatch[1]), lng: parseFloat(directMatch[2]) };

  const backendUrl = CONFIG.API_URL + "?action=resolveUrl&url=" + encodeURIComponent(url);
  const res = await fetch(backendUrl);
  if (!res.ok) throw new Error(`Backend resolve failed: ${res.status}`);
  const text = await res.text();

  let expandedUrl = "";
  let json = null;
  try {
    json = JSON.parse(text);
    expandedUrl = json.url || json.expandedUrl || json.resolved || "";
  } catch (e) {
    expandedUrl = text.trim();
  }

  if (json && json.lat && json.lng) return { lat: parseFloat(json.lat), lng: parseFloat(json.lng) };

  const d3Match = expandedUrl.match(/!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)/);
  if (d3Match) return { lat: parseFloat(d3Match[1]), lng: parseFloat(d3Match[2]) };

  const expandedMatch = expandedUrl.match(/@(-?\d+\.\d+),(-?\d+\.\d+)/);
  if (expandedMatch) return { lat: parseFloat(expandedMatch[1]), lng: parseFloat(expandedMatch[2]) };

  throw new Error(`Could not extract coords from expanded URL: ${expandedUrl}`);
}

async function fillLatLong() {
  let updated = [], skippedCount = 0, failedCount = 0;
  for (let row of allData) {
    if (!row.location) { skippedCount++; continue; }
    const latOk = row.Lat && !isNaN(parseFloat(row.Lat)) && parseFloat(row.Lat) !== 0;
    const lngOk = row.Long && !isNaN(parseFloat(row.Long)) && parseFloat(row.Long) !== 0;
    if (latOk && lngOk) { skippedCount++; continue; }
    try {
      const coords = await resolveGoogleMapsCoords(row.location);
      row.Lat = coords.lat;
      row.Long = coords.lng;
      updated.push(row);
    } catch (e) {
      failedCount++;
      console.error(`❌ Row _rowIndex:${row._rowIndex} — failed:`, e.message);
    }
  }
  updated.forEach(r => fetch(CONFIG.API_URL, { method: "POST", body: JSON.stringify(r) }));
  alert(`✅ Updated ${updated.length} rows\n⏭️ Skipped: ${skippedCount}\n❌ Failed: ${failedCount}`);
  if (updated.length > 0) loadData();
}

// ============================================================
// FIX MISSING NM
// ============================================================
function fixMissingNM() {
  let updated = [];
  allData.forEach((row, i) => {
    const lat = parseFloat(row.Lat);
    const lng = parseFloat(row.Long);
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
  updated.forEach(r => fetch(CONFIG.API_URL, { method: "POST", body: JSON.stringify(r) }));
  alert(`✅ Updated ${updated.length} rows`);
}

// ============================================================
// HOOD / GEOMETRY UTILITIES
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

// Show only hoods whose NM/MM match the active filter; show all when no filter
function updateHoodVisibility() {
  const filterNM = activeFilters.NM || "";
  const filterMM = activeFilters.MM || "";
  const noFilter = !filterNM && !filterMM;

  hoodLayers.forEach(({ layer, nm, mm }) => {
    const nmMatch = !filterNM || nm === filterNM;
    const mmMatch = !filterMM || mm === filterMM;
    const visible = noFilter || (nmMatch && mmMatch);

    if (visible) {
      if (!map.hasLayer(layer)) map.addLayer(layer);
      // reset highlight to default style
      layer.setStyle({ color: "blue", weight: 1, fillColor: "#4da6ff", fillOpacity: 0.15 });
    } else {
      if (map.hasLayer(layer)) map.removeLayer(layer);
    }
  });
}

function assignHood(coords) {
  const pt = turf.point([coords.lng, coords.lat]);
  let nearest = null, minDist = Infinity;
  for (let h of hoods) {
    const polygon = { type: "Feature", geometry: h.geometry };
    try {
      if (turf.booleanPointInPolygon(pt, polygon)) return h;
      const center = turf.centroid(polygon);
      const dist = turf.distance(pt, center);
      if (dist < minDist) { minDist = dist; nearest = h; }
    } catch (e) {}
  }
  return nearest;
}

function isEmpty(val) {
  return !val || val.toString().trim() === "" || val === "NA";
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
  const table = document.getElementById("detailsTable");
  table.innerHTML = `
    <tr><td><b>NM</b></td><td>${h.nano_market}</td></tr>
    <tr><td><b>MM</b></td><td>${h.micro_market}</td></tr>
    <tr><td><b>Region</b></td><td>${h.region}</td></tr>
    <tr><td><b>Hood ID</b></td><td>${h.hood_id}</td></tr>
  `;
}

function showDetails(row) {
  currentRow = row;
  const table = document.getElementById("detailsTable");
  table.innerHTML = "";
  Object.keys(row).forEach(key => {
    if (key === "_rowIndex") return; // internal field — hide from UI
    table.innerHTML += `<tr><td>${key}</td><td contenteditable="true" data-key="${key}">${row[key] || ""}</td></tr>`;
  });
}

function saveCurrent() {
  if (!currentRow) return;
  document.querySelectorAll("[contenteditable]").forEach(cell => {
    currentRow[cell.dataset.key] = cell.innerText;
  });
  if (!currentRow._rowIndex) { alert("❌ Cannot save — row index missing, try refreshing data first"); return; }
  fetch(CONFIG.API_URL, { method: "POST", body: JSON.stringify(currentRow) })
    .then(() => alert("✅ Saved"))
    .catch(err => alert("❌ Save failed: " + err.message));
}

// ============================================================
// SHEET PREVIEW
// ============================================================
function renderSheetPreview(data) {
  const table = document.getElementById("sheetPreviewTable");
  const countEl = document.getElementById("sheetRowCount");
  if (!data.length) { table.innerHTML = "<tr><td>No data</td></tr>"; return; }
  countEl.textContent = `${data.length} rows`;
  const previewCols = ["Name", "Category", "NM", "MM", "Lead Status", "Final Status", "App status", "Road", "Contact Name", "Contact number"];
  const availableCols = previewCols.filter(c => data[0].hasOwnProperty(c));
  table.innerHTML =
    `<thead><tr>${availableCols.map(c => `<th>${c}</th>`).join("")}</tr></thead>` +
    `<tbody>${data.map((row, idx) => `
      <tr data-idx="${idx}" style="cursor:pointer">
        ${availableCols.map(c => `<td title="${String(row[c]||'').replace(/"/g,'&quot;')}">${row[c]||""}</td>`).join("")}
      </tr>`).join("")}</tbody>`;
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

function refreshData() { loadData(); }

// ============================================================
// MAP SEARCH (unchanged)
// ============================================================
let searchMarker = null;
let searchDebounceTimer = null;

function initMapSearch() {
  const input = document.getElementById("mapSearchInput");
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
    const url = `https://nominatim.openstreetmap.org/search?q=${encodeURIComponent(query)}&format=json&limit=6&addressdetails=1`;
    const res = await fetch(url, { headers: { "Accept-Language": "en" } });
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
    icon: L.divIcon({ className: "custom-icon", html: `<div class="search-pin" style="font-size:28px">📌</div>` })
  }).addTo(map).bindPopup(`<b>${name}</b><br><small>${lat}, ${lng}</small>`).openPopup();
  map.setView(latlng, 16);
  document.getElementById("mapSearchInput").value = name;
  document.getElementById("mapSearchResults").classList.remove("open");
}