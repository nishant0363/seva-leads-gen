// ============================================================
// CONFIG — Region-aware
// ============================================================

const CONFIG = {

  // ── Main property sheet (Form Responses 1) ──────────────────
  API_URL: "https://script.google.com/macros/s/AKfycbwWrWpI52IecmST4dDRXrYDeePWFH_GSCgk-q_4myRvQA8UQMTYT9f68GklnOIGXRI9/exec",

  // ── Hotspots sheet ──────────────────────────────────────────
  HOTSPOT_URL: "https://script.google.com/macros/s/AKfycbwVBVyUxogZMVbIR_ybi9zMOdBftWUwyU5_E6oXDZw0PpUlVrJvKzjgLu9WeqFRuVQe/exec",

  // ── Demand sheet ────────────────────────────────────────────
  DEMAND_URL: "https://script.google.com/macros/s/AKfycbwSj8Br28NHbp7ZDHnnT51Q_dteXur_O5QqoBWDBcsqP8aPluqf4lWB733V7lMDTD2oEQ/exec",

  // ── Idle sheet ──────────────────────────────────────────────
  IDLE_URL: "https://script.google.com/macros/s/AKfycbxNdYhsSBL8unx054iJ4XmlCxt-P1Q87F1xBKEm5cNZlZvXLk8PU3FLuajl0neuPy-I/exec",

  // ── Demand Centroid sheet ────────────────────────────────────
  CENTROID_URL: "https://script.google.com/macros/s/AKfycbxJgK3i2VE9i6Q48uMzG0TYjqUII-I_SDgyNakLq4GAoGDvdYiUOm6EJGTQISYkSFjyjg/exec",

};

// ── Active Region (set by region-select page, persisted in sessionStorage) ──
// Value is the lowercase region string, e.g. "bangalore", "hyderabad", "noida"
// or "all" for no region filter.
const SELECTED_REGION = (function () {
  try {
    return sessionStorage.getItem("selectedRegion") || null;
  } catch (e) {
    return null;
  }
})();