let map;
let currentRow = null;
let allData = [];
let hoods = [];
let markers = [];
let activeFilters = {};
console.log("🚀 App initializing...");
init();

async function init() {
  console.log("🗺️ init() called");

  map = L.map('map').setView([12.9, 77.65], 12);
  console.log("🗺️ Map created with default view [12.9, 77.65], zoom 12");

  L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png')
    .addTo(map);
  console.log("🗺️ Tile layer added");

  // ✅ FIX: force map resize
  setTimeout(() => {
    console.log("🔄 Forcing map invalidateSize()");
    map.invalidateSize();
  }, 300);

  console.log("📦 Fetching hoods.json...");
  hoods = await fetch("hoods.json").then(r => r.json());
  console.log(`📦 hoods.json loaded — ${hoods.length} hoods`);

  drawHoods();

  await loadData();

  console.log("⏱️ Auto-refresh set for every 30 seconds");
  setInterval(loadData, 30000);
}

async function loadData() {
  const url = CONFIG.API_URL + "?t=" + Date.now();
  console.log("📡 loadData() called — fetching:", url);

  try {
    const res = await fetch(url);
    console.log("📡 Response status:", res.status, res.statusText);

    const text = await res.text();
    console.log("RAW RESPONSE:", text);

    const data = JSON.parse(text);
    console.log(`✅ Parsed data — ${data.length} rows`);

    allData = data;
    console.log("💾 allData updated with", allData.length, "rows");

    if (Object.values(activeFilters).some(v => v)) {
      console.log("🔽 Active filters detected — running filterAndRender()");
      filterAndRender();
    } else {
      console.log("🔽 No active filters — running renderMarkers()");
      renderMarkers();
    };

  } catch (err) {
    console.error("❌ Fetch failed:", err);
  }
  populateFilters();
}

function renderMarkers() {
  console.log("📍 renderMarkers() called");

  // ✅ CLEAR OLD MARKERS
  console.log(`🧹 Clearing ${markers.length} existing markers`);
  markers.forEach(m => map.removeLayer(m));
  markers = [];

  let skipped = 0;

  allData.forEach((row, i) => {
    const lat = parseFloat(row.Lat);
    const lng = parseFloat(row.Long);

    if (!isNaN(lat) && !isNaN(lng)) {

      // ✅ AUTO ASSIGN NM/MM
      if (!row.NM || !row.MM) {
        console.log(`🔍 Row ${i} missing NM/MM — assigning hood for [${lat}, ${lng}]`);
        const hood = assignHood({ lat, lng });

        if (hood) {
          row.NM = hood.nano_market;
          row.MM = hood.micro_market;
          row["NM Id"] = hood.hood_id;
          console.log(`✅ Row ${i} assigned NM: ${row.NM}, MM: ${row.MM}, Hood ID: ${row["NM Id"]}`);

          // 🔥 KEY FIX: immediately save to backend
          // renderMarkers() mutates the row in memory BEFORE fixMissingNM() runs,
          // so fixMissingNM() sees NM/MM already filled and skips these rows (Updated: 0 bug).
          // Solution: save right here, as soon as we assign.
          if (!row["Sr. No"]) {
            console.warn(`⚠️ Row ${i} — cannot auto-save, missing Sr. No`);
          } else {
            console.log(`📤 Auto-saving row ${i} (Sr. No: ${row["Sr. No"]}) to backend...`);
            fetch(CONFIG.API_URL, {
              method: "POST",
              body: JSON.stringify(row)
            })
            .then(() => console.log(`✅ Auto-saved Sr. No: ${row["Sr. No"]} — NM: ${row.NM}, MM: ${row.MM}`))
            .catch(err => console.error(`❌ Auto-save failed for Sr. No: ${row["Sr. No"]}`, err));
          }
        } else {
          console.warn(`⚠️ Row ${i} — no hood found for [${lat}, ${lng}]`);
        }
      }

      const marker = L.marker([lat, lng], {
        icon: getCategoryIcon(row.Category)
      })
        .addTo(map)
        .bindPopup(`
          <b>${row.Name || "No Name"}</b><br>
          ${row.Category || ""}<br>
          NM: ${row.NM || "-"}<br>
          MM: ${row.MM || "-"}
        `)
        .on('click', () => showDetails(row));

      markers.push(marker);
    } else {
      skipped++;
      console.warn(`⚠️ Row ${i} skipped — invalid Lat/Long:`, row.Lat, row.Long);
    }
  });

  console.log(`📍 renderMarkers() done — ${markers.length} markers added, ${skipped} rows skipped`);
}

function fixMissingNM() {
  console.log("🔧 fixMissingNM() called");

  let updated = [];

  allData.forEach((row, i) => {
    const lat = parseFloat(row.Lat);
    const lng = parseFloat(row.Long);

    // 🔬 DEEP DIAGNOSTIC: log raw values and types for every row
    console.log(
      `🔬 Row ${i} | Sr.No: ${row["Sr. No"]} | Name: ${row.Name} | ` +
      `Lat raw: ${JSON.stringify(row.Lat)} (${typeof row.Lat}) → parsed: ${lat} | ` +
      `Long raw: ${JSON.stringify(row.Long)} (${typeof row.Long}) → parsed: ${lng} | ` +
      `NM: ${JSON.stringify(row.NM)} | MM: ${JSON.stringify(row.MM)}`
    );

    if (isNaN(lat) || isNaN(lng)) {
      console.warn(`⚠️ Row ${i} (Sr.No ${row["Sr. No"]}) SKIPPED — isNaN lat:${isNaN(lat)} lng:${isNaN(lng)}`);
      return;
    }

    const nmEmpty = isEmpty(row.NM);
    const mmEmpty = isEmpty(row.MM);
    console.log(`🔬 Row ${i} isEmpty check — NM empty: ${nmEmpty}, MM empty: ${mmEmpty}`);

    if (nmEmpty || mmEmpty) {
      console.log(`🔍 Row ${i} has empty NM/MM — assigning hood for [${lat}, ${lng}]`);
      const hood = assignHood({ lat, lng });

      if (hood) {
        row.NM = hood.nano_market;
        row.MM = hood.micro_market;
        row["NM Id"] = hood.hood_id;
        console.log(`✅ Row ${i} updated — NM: ${row.NM}, MM: ${row.MM}, Hood ID: ${row["NM Id"]}`);
        updated.push(row);
      } else {
        console.log("❌ No hood found for:", lat, lng);
      }
    }
  });

  console.log("Updated:", updated.length);

  updated.forEach((r, i) => {
    console.log(`📤 POSTing updated row ${i}:`, r["Sr. No"] || "(no Sr. No)");
    fetch(CONFIG.API_URL, {
      method: "POST",
      body: JSON.stringify(r)
    });
  });

  alert(`✅ Updated ${updated.length} rows`);
}

// Extracts lat/lng from a Google Maps URL (short or full).
// Strategy:
//   1. If already a full URL with @lat,lng — parse directly (no fetch needed)
//   2. Otherwise — ask YOUR OWN Apps Script backend to resolve the short URL
//      (avoids CORS issues entirely since Apps Script can fetch any URL server-side)
async function resolveGoogleMapsCoords(url) {
  console.log(`🔗 resolveGoogleMapsCoords() — resolving: ${url}`);

  // Strategy 1: full URL already has coords — parse directly, no network call needed
  const directMatch = url.match(/@(-?\d+\.\d+),(-?\d+\.\d+)/);
  if (directMatch) {
    const result = { lat: parseFloat(directMatch[1]), lng: parseFloat(directMatch[2]) };
    console.log(`✅ Strategy 1: parsed coords directly from URL:`, result);
    return result;
  }

  // Strategy 2: short URL — ask Apps Script backend to expand it
  // Apps Script can fetch any URL server-side without CORS restrictions.
  // Your backend needs to handle action=resolveUrl (see Apps Script snippet below).
  const backendUrl = CONFIG.API_URL + "?action=resolveUrl&url=" + encodeURIComponent(url);
  console.log(`🌐 Strategy 2: asking Apps Script to resolve short URL: ${backendUrl}`);

  const res = await fetch(backendUrl);
  if (!res.ok) throw new Error(`Backend resolve failed: ${res.status}`);

  const text = await res.text();
  console.log(`🔗 Apps Script returned: ${text}`);

  let expandedUrl = "";
  let json = null;
  try {
    json = JSON.parse(text);
    expandedUrl = json.url || json.expandedUrl || json.resolved || "";
    console.log(`🔗 Expanded URL from backend: ${expandedUrl}`);
  } catch (e) {
    expandedUrl = text.trim();
    console.log(`🔗 Expanded URL (plain text): ${expandedUrl}`);
  }

  // Best case: backend already parsed coords from HTML body
  if (json && json.lat && json.lng) {
    const result = { lat: parseFloat(json.lat), lng: parseFloat(json.lng) };
    console.log(`✅ Strategy 2: backend returned coords directly:`, result);
    return result;
  }

  // Priority 1: !3d{lat}!4d{lng} — this is the ACTUAL PIN location, most accurate
  // The @lat,lng in the URL is just the map viewport center, which can be offset
  const d3Match = expandedUrl.match(/!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)/);
  if (d3Match) {
    const result = { lat: parseFloat(d3Match[1]), lng: parseFloat(d3Match[2]) };
    console.log(`✅ Strategy 2: parsed PIN coords from !3d!4d (accurate):`, result);
    return result;
  }

  // Fallback: @lat,lng — map viewport center, less accurate (up to ~300m off)
  const expandedMatch = expandedUrl.match(/@(-?\d+\.\d+),(-?\d+\.\d+)/);
  if (expandedMatch) {
    const result = { lat: parseFloat(expandedMatch[1]), lng: parseFloat(expandedMatch[2]) };
    console.warn(`⚠️ Strategy 2 fallback: using viewport coords @lat,lng (may be ~300m off):`, result);
    return result;
  }

  throw new Error(`Could not extract coords from expanded URL: ${expandedUrl}`);
}

/*
  ============================================================
  ADD THIS TO YOUR GOOGLE APPS SCRIPT (Code.gs / doGet):
  ============================================================

  In your doGet(e) function, add this case:

    if (e.parameter.action === "resolveUrl") {
      const shortUrl = e.parameter.url;
      try {
        // UrlFetchApp follows redirects by default — this gives us the final URL
        const response = UrlFetchApp.fetch(shortUrl, {
          followRedirects: true,
          muteHttpExceptions: true
        });
        const finalUrl = response.getHeaders()["Location"] || response.getContentText().match(/href="([^"]+)"/)?.[1] || shortUrl;
        return ContentService
          .createTextOutput(JSON.stringify({ url: finalUrl }))
          .setMimeType(ContentService.MimeType.JSON);
      } catch(err) {
        return ContentService
          .createTextOutput(JSON.stringify({ error: err.toString() }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

  ============================================================
*/

async function fillLatLong() {
  console.log("📍 fillLatLong() called");

  let updated = [];
  let skippedCount = 0;
  let failedCount = 0;

  for (let row of allData) {

    // Skip if no location link at all
    if (!row.location) {
      console.log(`⏭️ Row Sr.No ${row["Sr. No"]} — no location link, skipping`);
      skippedCount++;
      continue;
    }

    // Skip only if BOTH Lat and Long are already valid non-zero numbers
    const latOk = row.Lat && !isNaN(parseFloat(row.Lat)) && parseFloat(row.Lat) !== 0;
    const lngOk = row.Long && !isNaN(parseFloat(row.Long)) && parseFloat(row.Long) !== 0;

    if (latOk && lngOk) {
      console.log(`⏭️ Row Sr.No ${row["Sr. No"]} — already has valid Lat/Long (${row.Lat}, ${row.Long}), skipping`);
      skippedCount++;
      continue;
    }

    console.log(`🌐 Row Sr.No ${row["Sr. No"]} — missing coords, resolving: "${row.location}"`);

    try {
      const coords = await resolveGoogleMapsCoords(row.location);
      console.log(`✅ Row Sr.No ${row["Sr. No"]} — got coords: Lat ${coords.lat}, Long ${coords.lng}`);

      row.Lat = coords.lat;
      row.Long = coords.lng;

      updated.push(row);

    } catch (e) {
      failedCount++;
      console.error(`❌ Row Sr.No ${row["Sr. No"]} — failed to resolve "${row.location}":`, e.message);
    }
  }

  console.log(`📊 fillLatLong() summary — updated: ${updated.length}, skipped: ${skippedCount}, failed: ${failedCount}`);
  console.log(`📤 POSTing ${updated.length} updated rows`);

  updated.forEach((r, i) => {
    console.log(`📤 Posting row ${i} (Sr.No: ${r["Sr. No"]}) — Lat: ${r.Lat}, Long: ${r.Long}`);
    console.log(`📤 Full payload:`, JSON.stringify(r));
    fetch(CONFIG.API_URL, {
      method: "POST",
      body: JSON.stringify(r)
    })
    .then(res => res.text())
    .then(txt => {
      console.log(`📥 POST response for Sr.No ${r["Sr. No"]}:`, txt);
      try {
        const json = JSON.parse(txt);
        if (json.success) console.log(`✅ Saved Sr.No: ${r["Sr. No"]}`);
        else console.error(`❌ Backend rejected Sr.No ${r["Sr. No"]}:`, txt);
      } catch(e) {
        console.warn(`⚠️ Non-JSON response for Sr.No ${r["Sr. No"]}:`, txt);
      }
    })
    .catch(err => console.error(`❌ Save failed for Sr.No: ${r["Sr. No"]}`, err));
  });

  alert(`✅ Updated ${updated.length} rows\n⏭️ Skipped: ${skippedCount}\n❌ Failed: ${failedCount}`);

  // Reload fresh data so allData reflects what's now in the sheet
  // (prevents stale in-memory rows showing as "missing coords" on next run)
  if (updated.length > 0) {
    console.log("🔄 Reloading data to sync allData with sheet...");
    loadData();
  }
}

function applyFilters() {
  console.log("🔽 applyFilters() called");

  activeFilters = {
    Category: document.getElementById("filterCategory").value,
    Property: document.getElementById("filterProperty").value,
    "App status": document.getElementById("filterAppStatus").value,
    "Lead Status": document.getElementById("filterLeadStatus").value,
    NM: document.getElementById("filterNM").value,
    MM: document.getElementById("filterMM").value
  };

  console.log("🔽 Active filters:", activeFilters);

  filterAndRender();
}

function filterAndRender() {
  console.log("🔽 filterAndRender() called");

  const filtered = allData.filter(row => {
    return Object.keys(activeFilters).every(key => {
      if (!activeFilters[key]) return true;
      return row[key] === activeFilters[key];
    });
  });

  console.log(`🔽 Filter result: ${filtered.length} / ${allData.length} rows match`);

  renderFilteredMarkers(filtered);
}

function renderFilteredMarkers(data) {
  console.log(`📍 renderFilteredMarkers() called with ${data.length} rows`);

  markers.forEach(m => map.removeLayer(m));
  markers = [];

  const bounds = [];
  let skipped = 0;

  data.forEach((row, i) => {
    const lat = parseFloat(row.Lat);
    const lng = parseFloat(row.Long);

    if (!isNaN(lat) && !isNaN(lng)) {
      const marker = L.marker([lat, lng], {
        icon: getCategoryIcon(row.Category)
      })
        .addTo(map)
        .on('click', () => showDetails(row));

      markers.push(marker);
      bounds.push([lat, lng]);
    } else {
      skipped++;
      console.warn(`⚠️ Row ${i} skipped — invalid Lat/Long:`, row.Lat, row.Long);
    }
  });

  console.log(`📍 renderFilteredMarkers() done — ${markers.length} markers, ${skipped} skipped`);

  if (bounds.length) {
    console.log("🗺️ Fitting map to bounds:", bounds.length, "points");
    map.fitBounds(bounds);
  } else {
    console.warn("⚠️ No valid bounds — map not adjusted");
  }
}

function clearFilters() {
  console.log("🧹 clearFilters() called — resetting all filters");

  activeFilters = {};

  document.querySelectorAll("select").forEach(s => s.value = "");

  console.log("🧹 All dropdowns reset");

  renderMarkers();
}

function getCategoryIcon(category) {
  console.log(`🎨 getCategoryIcon() called for category: "${category}"`);

  const icons = {
    shop: "📍",
    apartment: "📍",
    bus_stop: "📍",
    park: "📍",
    pg: "📍",
    restaurant: "📍"
  };

  const emoji = icons[category?.toLowerCase()] || "📍";
  console.log(`🎨 Icon resolved: "${emoji}" for category: "${category}"`);

  return L.divIcon({
    className: "custom-icon",
    html: `<div style="font-size:18px">${emoji}</div>`
  });
}

function isEmpty(val) {
  const result = !val || val.toString().trim() === "" || val === "NA";
  console.log(`🔍 isEmpty("${val}") → ${result}`);
  return result;
}

function showHoodDetails(h) {
  console.log("🏘️ showHoodDetails() called for hood:", h.nano_market, "|", h.micro_market, "| ID:", h.hood_id);

  const table = document.getElementById("detailsTable");
  table.innerHTML = `
    <tr><td><b>NM</b></td><td>${h.nano_market}</td></tr>
    <tr><td><b>MM</b></td><td>${h.micro_market}</td></tr>
    <tr><td><b>Region</b></td><td>${h.region}</td></tr>
    <tr><td><b>Hood ID</b></td><td>${h.hood_id}</td></tr>
  `;
}

function populateFilters() {
  console.log("🔽 populateFilters() called");

  const fields = [
    { key: "Category", label: "Category" },
    { key: "Property", label: "Property" },
    { key: "App status", label: "App Status" },
    { key: "Lead Status", label: "Lead Status" },
    { key: "NM", label: "NM" },
    { key: "MM", label: "MM" }
  ];

  fields.forEach(f => {
    const id = "filter" + f.key.replace(/ /g, "");
    const select = document.getElementById(id);

    if (!select) {
      console.warn(`⚠️ populateFilters: element #${id} not found in DOM`);
      return;
    }

    const values = [...new Set(allData.map(r => r[f.key]).filter(Boolean))];
    console.log(`🔽 Filter "${f.key}" — ${values.length} unique values:`, values);

    // preserve current selection
    const currentValue = select.value;

    select.innerHTML =
      `<option value="">${f.label}</option>` +
      values.map(v => `<option value="${v}">${v}</option>`).join("");

    // restore selection
    select.value = currentValue;
    console.log(`🔽 Filter "${f.key}" restored to: "${currentValue}"`);
  });
}

function drawHoods() {
  console.log(`🗺️ drawHoods() called — drawing ${hoods.length} hoods`);

  hoods.forEach((h, i) => {
    if (!h.geometry) {
      console.warn(`⚠️ Hood ${i} skipped — no geometry`);
      return;
    }

    const layer = L.geoJSON(h.geometry, {
      style: {
        color: "blue",
        weight: 1,
        fillColor: "#4da6ff",
        fillOpacity: 0.15
      }
    }).addTo(map);

    console.log(`🗺️ Hood drawn: ${h.nano_market} (${h.hood_id})`);

    // ✅ CLICK ON NM
    layer.on("click", () => {
      console.log(`🖱️ Hood clicked: ${h.nano_market} | ${h.micro_market} | ID: ${h.hood_id}`);

      // highlight
      layer.setStyle({
        fillColor: "orange",
        fillOpacity: 0.4
      });

      showHoodDetails(h);
    });
  });

  console.log("🗺️ drawHoods() complete");
}

function assignHood(coords) {
  console.log(`📌 assignHood() called for coords:`, coords);

  const pt = turf.point([coords.lng, coords.lat]);

  let nearest = null;
  let minDist = Infinity;

  for (let h of hoods) {
    const polygon = {
      type: "Feature",
      geometry: h.geometry
    };

    try {
      if (turf.booleanPointInPolygon(pt, polygon)) {
        console.log(`✅ assignHood() — point inside hood: ${h.nano_market} (${h.hood_id})`);
        return h;
      }

      // fallback: nearest
      const center = turf.centroid(polygon);
      const dist = turf.distance(pt, center);

      if (dist < minDist) {
        minDist = dist;
        nearest = h;
      }

    } catch (e) {
      console.warn(`⚠️ assignHood() error for hood ${h.hood_id}:`, e.message);
    }
  }

  console.log(`📌 assignHood() fallback — nearest hood: ${nearest?.nano_market} (${nearest?.hood_id}), dist: ${minDist.toFixed(3)} km`);
  return nearest; // 🔥 fallback
}

function showDetails(row) {
  console.log("📋 showDetails() called for row:", row["Sr. No"] || "(no Sr. No)", "|", row.Name || "(no name)");
  console.log("📋 Full row data:", row);

  currentRow = row;

  const table = document.getElementById("detailsTable");
  table.innerHTML = "";

  Object.keys(row).forEach(key => {
    table.innerHTML += `
      <tr>
        <td>${key}</td>
        <td contenteditable="true" data-key="${key}">
          ${row[key] || ""}
        </td>
      </tr>
    `;
  });

  console.log(`📋 Details table populated with ${Object.keys(row).length} fields`);
}

function saveCurrent() {
  console.log("💾 saveCurrent() called");

  if (!currentRow) {
    console.warn("⚠️ saveCurrent() — no currentRow selected");
    return;
  }

  console.log("💾 Reading editable cells...");
  document.querySelectorAll("[contenteditable]").forEach(cell => {
    const key = cell.dataset.key;
    const oldVal = currentRow[key];
    currentRow[key] = cell.innerText;
    if (oldVal !== cell.innerText) {
      console.log(`✏️ Field "${key}" changed: "${oldVal}" → "${cell.innerText}"`);
    }
  });

  // ✅ ensure identifier is sent
  if (!currentRow["Sr. No"]) {
    console.error("❌ saveCurrent() — missing Sr. No, aborting save");
    alert("❌ Missing Sr. No — cannot update row");
    return;
  }

  console.log(`📤 POSTing row Sr. No: ${currentRow["Sr. No"]} to:`, CONFIG.API_URL);
  console.log("📤 Payload:", currentRow);

  fetch(CONFIG.API_URL, {
    method: "POST",
    body: JSON.stringify(currentRow)
  })
  .then(() => {
    console.log("✅ saveCurrent() — save successful");
    alert("✅ Saved");
  })
  .catch((err) => {
    console.error("❌ saveCurrent() — save failed:", err);
    alert("❌ Save failed");
  });
}

function refreshData() {
  console.log("🔄 refreshData() called — triggering loadData()");
  loadData(); // ✅ better than reload
}