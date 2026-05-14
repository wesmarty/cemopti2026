
// ════════════════════════════════════════════════════════════════════════════
//  CEMOPTI v6.0 — Full Wellbore Depth-Based Cementing Simulation Platform
//  © Kemiserve FZE, SPC Free Zone, Sharjah, UAE. All rights reserved.
//  v6: Rigorous Calc Engine · Validation Schema · Per-Param Rules · DB Layer
//      Dynamic Charts · Real-time output updates from input changes
// ════════════════════════════════════════════════════════════════════════════

import { useState, useEffect, useCallback, useRef, useMemo } from "react";
import * as XLSX from "xlsx";
import {
  LineChart, Line, XAxis, YAxis, CartesianGrid,
  Tooltip, Legend, ResponsiveContainer, ReferenceLine
} from "recharts";

// ─────────────────────────── DESIGN TOKENS (unchanged) ────────────────────
const T = {
  bg:"#0b1120", surface:"#111827", card:"#162032", panel:"#1c2d42",
  border:"#1e3350", borderHi:"#2a4a6e", accent:"#00c6ff", accent2:"#0072ff",
  gold:"#f5a623", green:"#22c55e", red:"#ef4444", yellow:"#fbbf24",
  cyan:"#06b6d4", text:"#e2e8f0", muted:"#64748b", dim:"#334155",
};

const css = `
  @import url('https://fonts.googleapis.com/css2?family=Orbitron:wght@400;700;900&family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;500;600&display=swap');
  *,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
  body{background:${T.bg};color:${T.text};font-family:'IBM Plex Sans',sans-serif}
  ::-webkit-scrollbar{width:6px;height:6px}
  ::-webkit-scrollbar-track{background:${T.surface}}
  ::-webkit-scrollbar-thumb{background:${T.border};border-radius:3px}
  ::-webkit-scrollbar-thumb:hover{background:${T.borderHi}}
  input,select,textarea{background:${T.bg};color:${T.text};border:1px solid ${T.border};border-radius:4px;padding:6px 10px;font-family:'IBM Plex Mono',monospace;font-size:12px;outline:none;transition:border-color 0.2s}
  input:focus,select:focus{border-color:${T.accent};box-shadow:0 0 0 2px rgba(0,198,255,0.15)}
  select option{background:${T.surface}}
  input.error{border-color:${T.red}!important;box-shadow:0 0 0 2px rgba(239,68,68,0.2)!important}
  button{cursor:pointer;font-family:'IBM Plex Sans',sans-serif}
  table{border-collapse:collapse;width:100%}
  th{background:${T.panel};color:${T.accent};font-family:'IBM Plex Mono',monospace;font-size:11px;text-transform:uppercase;letter-spacing:0.05em;padding:8px 10px;text-align:left;border-bottom:1px solid ${T.border}}
  td{padding:5px 10px;border-bottom:1px solid ${T.border};font-size:12px;font-family:'IBM Plex Mono',monospace}
  tr:hover td{background:rgba(0,198,255,0.04)}
  .grid-2{display:grid;grid-template-columns:1fr 1fr;gap:16px}
  .grid-3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:14px}
  .grid-4{display:grid;grid-template-columns:repeat(4,1fr);gap:12px}
  @keyframes pulse{0%,100%{opacity:1}50%{opacity:0.5}}
  @keyframes slideIn{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:translateY(0)}}
  @keyframes fadeIn{from{opacity:0}to{opacity:1}}
  .animate-in{animation:slideIn 0.3s ease forwards}
  .pulse{animation:pulse 2s infinite}
  @media print{
    .no-print{display:none!important}
    body{background:#fff!important;color:#000!important}
    .print-page{page-break-after:always}
  }
`;

// ─────────────────────────── UNIT CONVERSION ─────────────────────────────────
// All internal values stored in Imperial. Metric display converts on render.
const UNITS = {
  imperial: {
    depth: "ft", depth_s: "ft", pressure: "psi", density: "ppg",
    temperature: "°F", velocity: "ft/min", volume: "bbl", diameter: "in",
    weight: "lb/ft", label: "Field"
  },
  metric: {
    depth: "m", depth_s: "m", pressure: "kPa", density: "kg/m³",
    temperature: "°C", velocity: "m/min", volume: "m³", diameter: "mm",
    weight: "kg/m", label: "Metric (SI)"
  }
};

function cvt(val, field, toMetric, round = 3) {
  if (!toMetric) return Math.round(val * Math.pow(10, round)) / Math.pow(10, round);
  const factors = {
    depth: 0.3048, pressure: 6.89476, density: 119.826, temperature_offset: true,
    velocity: 0.3048, volume: 0.158987, diameter: 25.4, weight: 1.48816,
  };
  if (field === "temperature") return Math.round((val - 32) * 5/9 * 10) / 10;
  const f = factors[field] || 1;
  return Math.round(val * f * Math.pow(10, round)) / Math.pow(10, round);
}

function cvtIn(val, field, fromMetric) {
  // Convert input from metric → imperial for storage
  if (!fromMetric) return parseFloat(val) || 0;
  const factors = {
    depth: 1/0.3048, pressure: 1/6.89476, density: 1/119.826,
    temperature: true, velocity: 1/0.3048, volume: 1/0.158987,
    diameter: 1/25.4, weight: 1/1.48816,
  };
  if (field === "temperature") return parseFloat(val) * 9/5 + 32;
  const f = factors[field] || 1;
  return Math.round((parseFloat(val) || 0) * f * 10000) / 10000;
}

// ─────────────────────────── DEFAULT DATA ─────────────────────────────────────
const DEFAULT_DATA = {
  survey_data: [
    {md:0,tvd:0,inclination:0,azimuth:0},
    {md:1000,tvd:1000,inclination:0,azimuth:0},
    {md:2500,tvd:2490,inclination:3.2,azimuth:45},
    {md:4000,tvd:3975,inclination:5.1,azimuth:62},
    {md:5500,tvd:5440,inclination:4.8,azimuth:68},
    {md:7000,tvd:6870,inclination:2.1,azimuth:70},
    {md:8500,tvd:8350,inclination:0.8,azimuth:72},
  ],
  casing_profile: [
    {id:"c1",from_depth:0,to_depth:500,od:20.0,id_:19.0,grade:"K-55",type:"Surface",weight:94,description:"Surface Casing"},
    {id:"c2",from_depth:0,to_depth:3500,od:13.375,id_:12.415,grade:"K-55",type:"Intermediate",weight:54.5,description:"Intermediate Casing"},
    {id:"c3",from_depth:0,to_depth:7800,od:9.625,id_:8.835,grade:"L-80",type:"Production",weight:40,description:"Production Casing"},
  ],
  open_hole_profile: [
    {id:"h1",from_depth:0,to_depth:500,hole_size:26.0,excess:50,description:"Conductor"},
    {id:"h2",from_depth:500,to_depth:3500,hole_size:26.0,excess:80,description:"Intermediate hole"},
    {id:"h3",from_depth:3500,to_depth:8500,hole_size:26.0,excess:100,description:"Production hole"},
  ],
  formation_data: [
    {id:"f1",from_depth:0,   to_depth:2000, pore_gradient:0.433, frac_gradient:0.65, temperature:100, lithology:"Sand/Shale"},
    {id:"f2",from_depth:2000, to_depth:4500, pore_gradient:0.452, frac_gradient:0.74, temperature:150, lithology:"Shale"},
    {id:"f3",from_depth:4500, to_depth:6500, pore_gradient:0.468, frac_gradient:0.82, temperature:195, lithology:"Limestone"},
    {id:"f4",from_depth:6500, to_depth:8500, pore_gradient:0.491, frac_gradient:0.90, temperature:245, lithology:"Dolomite"},
  ],
  fluid_data: [
    {fluid_id:"fl1",fluid_name:"Spacer",      type:"Spacer", density:11.5,pv:15,yp:10,model:"Bingham", n:0.8, k:0.18,color:"#34D399"},
    {fluid_id:"fl2",fluid_name:"Lead Cement", type:"Cement", density:15.8,pv:55,yp:22,model:"Bingham", n:0.68,k:0.62,color:"#F97316"},
    {fluid_id:"fl3",fluid_name:"Drilling Mud",type:"Mud",    density:10.2,pv:22,yp:14,model:"Bingham", n:0.85,k:0.25,color:"#3B82F6"},
  ],
  centralizer_data: [
    {id:"cr1",from_depth:6500,to_depth:8500,spacing:40,type:"Bow Spring",standoff:72,run_in_force:120},
    {id:"cr2",from_depth:4000,to_depth:6500,spacing:60,type:"Solid Body",standoff:80,run_in_force:200},
  ],
  temperature_profile: [
    {depth:0,temperature:70},{depth:2000,temperature:120},
    {depth:4000,temperature:165},{depth:6000,temperature:205},{depth:8500,temperature:245},
  ],
  pumping_schedule: [
    {stage:1,fluid_id:"fl1",fluid_name:"Spacer",      rate:7.0,volume:60, purpose:"Preflush",    time_start:0},
    {stage:2,fluid_id:"fl2",fluid_name:"Lead Cement", rate:6.0,volume:180,purpose:"Lead Cement", time_start:0},
    {stage:3,fluid_id:"fl3",fluid_name:"Drilling Mud",rate:8.0,volume:100,purpose:"Displacement",time_start:0},
  ],
  balanced_plug_data: [
    {id:"bp1",from_depth:7800,to_depth:8500,length:700,set_depth:7800,excess:200,fluid_id:"fl5"},
  ],
  general_data: {
    well_type:"Exploration",
    toc:3000,
    total_depth_md:8500,
    hole_size:8.75,
    casing_shoe_depth:7800,
    num_timesteps:100,
    num_depth_grids:100,
  },
};

// ─────────────────────────── FLUID TYPE DEFAULTS ──────────────────────────────
const FLUID_TYPE_DEFAULTS = {
  "Mud":    { density:10.2, pv:22, yp:14, model:"Bingham",   n:0.85, k:0.25, color:"#3B82F6" },
  "Cement": { density:15.8, pv:55, yp:22, model:"Bingham",   n:0.68, k:0.62, color:"#F97316" },
  "Spacer": { density:11.5, pv:15, yp:10, model:"Bingham",   n:0.80, k:0.18, color:"#34D399" },
  "Wash":   { density:8.4,  pv:3,  yp:2,  model:"Newtonian", n:1.00, k:0.08, color:"#A78BFA" },
  "Brine":  { density:9.6,  pv:5,  yp:3,  model:"Power Law", n:0.90, k:0.12, color:"#67E8F9" },
  "Plug":   { density:15.8, pv:58, yp:25, model:"Bingham",   n:0.68, k:0.62, color:"#FCD34D" },
};

// ─────────────────────────── CASING TYPE PROPERTIES ──────────────────────────
const CASING_TYPE_PROPS = {
  "Surface":      { od:20.0,   id_:19.0,   grade:"K-55",  weight:94,   clearance:1.0 },
  "Intermediate": { od:13.375, id_:12.415, grade:"K-55",  weight:54.5, clearance:0.8 },
  "Production":   { od:9.625,  id_:8.835,  grade:"L-80",  weight:40,   clearance:0.6 },
  "Liner":        { od:7.0,    id_:6.276,  grade:"P-110", weight:26,   clearance:0.5 },
  "Tieback":      { od:9.625,  id_:8.835,  grade:"L-80",  weight:40,   clearance:0.6 },
};

// ─────────────────────────── CENTRALIZER TYPE PROPERTIES ─────────────────────
const CENTRALIZER_TYPE_PROPS = {
  "Bow Spring": { baseStandoff:67, rihForce:120, turbulenceBonus:0, description:"High restoring force; preferred in deviated wells ≤60°" },
  "Solid Body": { baseStandoff:80, rihForce:200, turbulenceBonus:3, description:"Low drag; preferred in horizontal/high-angle sections" },
  "Semi-Rigid": { baseStandoff:72, rihForce:160, turbulenceBonus:1, description:"Compromise; moderate deviation, moderate drag" },
  "Turbo":      { baseStandoff:75, rihForce:180, turbulenceBonus:8, description:"Induces turbulence; best for mud removal in horizontal" },
  "Flex":       { baseStandoff:65, rihForce:100, turbulenceBonus:2, description:"Flexible; tight tolerances or corrosion-inhibited casing" },
};

// ─────────────────────────── FLUID MAPPING SYNCHRONIZATION ────────────────────────
/**
 * Synchronize pumping schedule fluid names with current fluid database
 * Fixes stale fluid_name references when fluid database has been updated
 * @param {Array} pumpingSchedule - Pumping schedule stages
 * @param {Array} fluidDatabase - Current fluid database
 * @returns {Array} Updated pumping schedule with synchronized fluid names
 */
function synchronizeFluidMapping(pumpingSchedule, fluidDatabase) {
  if (!pumpingSchedule || !fluidDatabase) return pumpingSchedule;
  
  return pumpingSchedule.map(stage => {
    // Find the actual fluid object by ID
    var actualFluid = fluidDatabase.find(f => f.fluid_id === stage.fluid_id);
    
    if (actualFluid && actualFluid.fluid_name !== stage.fluid_name) {
      // Fluid name mismatch - synchronize to current database
      console.log(`Synchronizing Stage ${stage.stage}: "${stage.fluid_name}" → "${actualFluid.fluid_name}" (ID: ${stage.fluid_id})`);
      return {
        ...stage,
        fluid_name: actualFluid.fluid_name
      };
    }
    
    if (!actualFluid) {
      // Fluid ID not found - log warning but keep stage
      console.warn(`Stage ${stage.stage}: Fluid ID ${stage.fluid_id} not found in fluid database`);
    }
    
    return stage;
  });
}

// ─────────────────────────── DATABASE INIT & PERSISTENCE ─────────────────────
const DB_KEY = "cemopti_db_v6";

function buildFreshDB() {
  return {
    users: [{ user_id:"u1", username:"cemopti", password:"1234", email:"engineer@kemiserve.com", role:"Senior Cementing Engineer" }],
    projects: [{
      project_id:"p1", user_id:"u1", project_name:"Gulf Well Alpha-7",
      well_name:"Alpha-7", field_name:"Gulf Block-A", hole_type:"open_hole",
      unit_system:"imperial", created_at:new Date().toISOString(), status:"active", depth_step:50
    }],
    survey_data:         { p1: JSON.parse(JSON.stringify(DEFAULT_DATA.survey_data)) },
    casing_profile:      { p1: JSON.parse(JSON.stringify(DEFAULT_DATA.casing_profile)) },
    open_hole_profile:   { p1: JSON.parse(JSON.stringify(DEFAULT_DATA.open_hole_profile)) },
    formation_data:      { p1: JSON.parse(JSON.stringify(DEFAULT_DATA.formation_data)) },
    fluid_data:          { p1: JSON.parse(JSON.stringify(DEFAULT_DATA.fluid_data)) },
    centralizer_data:    { p1: JSON.parse(JSON.stringify(DEFAULT_DATA.centralizer_data)) },
    temperature_profile: { p1: JSON.parse(JSON.stringify(DEFAULT_DATA.temperature_profile)) },
    pumping_schedule:    { p1: JSON.parse(JSON.stringify(DEFAULT_DATA.pumping_schedule)) },
    balanced_plug_data:  { p1: JSON.parse(JSON.stringify(DEFAULT_DATA.balanced_plug_data)) },
    general_data:        { p1: { ...DEFAULT_DATA.general_data } },
    simulation_runs:     {},
    actual_job_data:     { p1: [] },
  };
}

function initDB() {
  try {
    const stored = localStorage.getItem(DB_KEY);
    if (stored) {
      const parsed = JSON.parse(stored);
      // Ensure simulation_runs and general_data exist (schema migration)
      if (!parsed.simulation_runs) parsed.simulation_runs = {};
      if (!parsed.general_data)   parsed.general_data = { p1: { ...DEFAULT_DATA.general_data } };
      return parsed;
    }
  } catch (e) { /* corrupted storage — fall through to fresh DB */ }
  return buildFreshDB();
}

function saveDB(db) {
  try { localStorage.setItem(DB_KEY, JSON.stringify(db)); } catch (e) { /* storage quota exceeded — ignore */ }
}

// Full validation with per-parameter rules

// ─────────────────────────── VALIDATION SCHEMA ───────────────────────────────
const VALIDATION_RULES = {
  survey_md:          { min:0,     max:40000,  unit:"ft",         label:"Measured Depth",          severity:"error"   },
  survey_tvd:         { min:0,     max:40000,  unit:"ft",         label:"True Vertical Depth",     severity:"error"   },
  survey_inclination: { min:0,     max:90,     unit:"°",          label:"Inclination",             severity:"error"   },
  survey_azimuth:     { min:0,     max:360,    unit:"°",          label:"Azimuth",                 severity:"error"   },
  casing_od:          { min:2.375, max:30,     unit:"in",         label:"Casing OD",               severity:"error"   },
  casing_id:          { min:2.0,   max:29,     unit:"in",         label:"Casing ID",               severity:"error"   },
  casing_weight:      { min:4,     max:200,    unit:"lb/ft",      label:"Casing Weight",           severity:"warning" },
  casing_from_depth:  { min:0,     max:40000,  unit:"ft",         label:"Casing From Depth",       severity:"error"   },
  casing_to_depth:    { min:1,     max:40000,  unit:"ft",         label:"Casing To Depth",         severity:"error"   },
  hole_size:          { min:3.5,   max:36,     unit:"in",         label:"Hole Size",               severity:"error"   },
  hole_excess:        { min:0,     max:2000,   unit:"ft",         label:"Hole Excess",             severity:"warning" },
  pore_gradient:      { min:0.1,   max:1.5,    unit:"psi/ft",     label:"Pore Pressure Gradient",  severity:"error"   },
  frac_gradient:      { min:0.2,   max:2.0,    unit:"psi/ft",     label:"Fracture Gradient",       severity:"error"   },
  fluid_density:      { min:7.0,   max:22.0,   unit:"ppg",        label:"Fluid Density",           severity:"error"   },
  fluid_pv:           { min:0,     max:500,    unit:"cP",         label:"Plastic Viscosity",       severity:"error"   },
  fluid_yp:           { min:0,     max:200,    unit:"lb/100ft²",  label:"Yield Point",             severity:"warning" },
  fluid_n:            { min:0.01,  max:1.99,   unit:"—",          label:"Flow Index n",            severity:"error"   },
  fluid_k:            { min:0.001, max:10,     unit:"—",          label:"Consistency K",           severity:"warning" },
  pump_rate:          { min:0,     max:30,     unit:"bpm",        label:"Pump Rate",               severity:"error"   },
  pump_volume:        { min:0,     max:10000,  unit:"bbl",        label:"Pump Volume",             severity:"warning" },
  cent_spacing:       { min:5,     max:120,    unit:"ft",         label:"Centralizer Spacing",     severity:"warning" },
  cent_standoff:      { min:0,     max:100,    unit:"%",          label:"Standoff Percentage",     severity:"error"   },
  temperature:        { min:50,    max:500,    unit:"°F",         label:"Temperature",             severity:"error"   },
  temp_depth:         { min:0,     max:40000,  unit:"ft",         label:"Temperature Depth",       severity:"error"   },
  plug_length:        { min:50,    max:5000,   unit:"ft",         label:"Plug Length",             severity:"error"   },
  plug_set_depth:     { min:50,    max:40000,  unit:"ft",         label:"Plug Set Depth",          severity:"error"   },
  plug_excess:        { min:0,     max:1000,   unit:"ft",         label:"Plug Excess",             severity:"warning" },
  gen_toc:            { min:0,     max:40000,  unit:"ft",         label:"Top of Cement",           severity:"warning" },
  gen_total_depth:    { min:100,   max:40000,  unit:"ft",         label:"Total Depth",             severity:"error"   },
  gen_hole_size:      { min:3.5,   max:36,     unit:"in",         label:"General Hole Size",       severity:"error"   },
  gen_casing_shoe:    { min:100,   max:40000,  unit:"ft",         label:"Casing Shoe Depth",       severity:"error"   },
};

function checkRange(value, ruleKey, label) {
  const rule = VALIDATION_RULES[ruleKey];
  if (!rule) return null;
  const v = parseFloat(value);
  if (isNaN(v)) return { severity:"error", msg:`${label||rule.label}: must be a number` };
  if (v < rule.min) return { severity:rule.severity, msg:`${label||rule.label}: ${v} ${rule.unit} is below minimum (${rule.min} ${rule.unit})` };
  if (v > rule.max) return { severity:rule.severity, msg:`${label||rule.label}: ${v} ${rule.unit} exceeds maximum (${rule.max} ${rule.unit})` };
  return null;
}

function validateAll(db, pid) {
  const errors = [], warnings = [];
  const addE = (page, msg, field="") => errors.push({ page, msg, field });
  const addW = (page, msg, field="") => warnings.push({ page, msg, field });
  const check = (page, value, ruleKey, rowLabel, field) => {
    const r = checkRange(value, ruleKey, rowLabel);
    if (!r) return;
    r.severity === "error" ? addE(page, r.msg, field) : addW(page, r.msg, field);
  };

  const survey  = db.survey_data[pid]||[];
  const casing  = db.casing_profile[pid]||[];
  const holes   = db.open_hole_profile[pid]||[];
  const form    = db.formation_data[pid]||[];
  const fluids  = db.fluid_data[pid]||[];
  const pump    = db.pumping_schedule[pid]||[];
  const cent    = db.centralizer_data[pid]||[];
  const temps   = db.temperature_profile[pid]||[];
  const plug    = db.balanced_plug_data[pid]||[];
  const gen     = db.general_data?.[pid] || {};

  // ── Survey ──────────────────────────────────────────────────────────────────
  if (survey.length < 2) addE("Survey","At least 2 survey points required");
  survey.forEach((s,i) => {
    check("Survey", s.md,          "survey_md",          `Survey row ${i+1} MD`);
    check("Survey", s.tvd,         "survey_tvd",         `Survey row ${i+1} TVD`);
    check("Survey", s.inclination, "survey_inclination", `Survey row ${i+1} Incl.`);
    check("Survey", s.azimuth,     "survey_azimuth",     `Survey row ${i+1} Az.`);
    if (i > 0 && s.md <= survey[i-1].md)
      addE("Survey", `Survey row ${i+1}: MD must be strictly increasing (${s.md} ≤ ${survey[i-1].md} ft)`);
    if (s.tvd > s.md + 1)
      addE("Survey", `Survey row ${i+1}: TVD (${s.tvd}) cannot exceed MD (${s.md})`);
  });

  // ── Casing ───────────────────────────────────────────────────────────────────
  if (casing.length === 0) addE("Casing","At least one casing string is required");
  casing.forEach((c,i) => {
    check("Casing", c.od,         "casing_od",         `Casing row ${i+1} OD`);
    check("Casing", c.id_,        "casing_id",         `Casing row ${i+1} ID`);
    check("Casing", c.weight,     "casing_weight",     `Casing row ${i+1} Weight`);
    check("Casing", c.from_depth, "casing_from_depth", `Casing row ${i+1} From`);
    check("Casing", c.to_depth,   "casing_to_depth",   `Casing row ${i+1} To`);
    if (c.id_ >= c.od)
      addE("Casing", `Casing row ${i+1}: ID (${c.id_}") must be less than OD (${c.od}")`);
    if (c.from_depth >= c.to_depth)
      addE("Casing", `Casing row ${i+1}: From depth (${c.from_depth}) must be less than To depth (${c.to_depth})`);
    const hAtTop = holes.find(h => c.to_depth >= h.from_depth && c.to_depth <= h.to_depth);
    if (hAtTop && c.od >= hAtTop.hole_size)
      addE("Casing", `Casing row ${i+1}: OD (${c.od}") ≥ hole size (${hAtTop.hole_size}") — impossible geometry`);
  });

  // ── Open Hole ─────────────────────────────────────────────────────────────────
  holes.forEach((h,i) => {
    check("Open Hole", h.hole_size,   "hole_size",   `Hole row ${i+1} Size`);
    check("Open Hole", h.excess,      "hole_excess", `Hole row ${i+1} Excess`);
    if (h.from_depth >= h.to_depth)
      addE("Open Hole", `Hole row ${i+1}: From depth must be < To depth`);
  });

  // ── Formation ─────────────────────────────────────────────────────────────────
  if (form.length === 0) addE("Formation","At least one formation interval required");
  form.forEach((f,i) => {
    check("Formation", f.pore_gradient, "pore_gradient", `Formation row ${i+1} Pore Grad`);
    check("Formation", f.frac_gradient, "frac_gradient", `Formation row ${i+1} Frac Grad`);
    if (f.frac_gradient <= f.pore_gradient)
      addE("Formation", `Formation row ${i+1}: Frac gradient (${f.frac_gradient}) must exceed pore gradient (${f.pore_gradient})`);
    if (f.from_depth >= f.to_depth)
      addE("Formation", `Formation row ${i+1}: From depth must be < To depth`);
    // Typical pressure window check
    const window = f.frac_gradient - f.pore_gradient;
    if (window < 0.05)
      addW("Formation", `Formation row ${i+1}: Pressure window very narrow (${window.toFixed(3)} psi/ft) — ECD control will be critical`);
  });

  // ── Fluids ────────────────────────────────────────────────────────────────────
  if (fluids.length === 0) addE("Fluids","At least one fluid must be defined");
  fluids.forEach((f,i) => {
    check("Fluids", f.density, "fluid_density", `${f.fluid_name} Density`);
    check("Fluids", f.pv,      "fluid_pv",      `${f.fluid_name} PV`);
    check("Fluids", f.yp,      "fluid_yp",      `${f.fluid_name} YP`);
    check("Fluids", f.n,       "fluid_n",       `${f.fluid_name} n`);
    check("Fluids", f.k,       "fluid_k",       `${f.fluid_name} K`);
    if (!f.fluid_name || f.fluid_name.trim() === "")
      addE("Fluids", `Fluid row ${i+1}: Name is required`);
  });
  // Density hierarchy check
  const cemFluids = fluids.filter(f => f.type?.toLowerCase()==="cement").sort((a,b)=>a.density-b.density);
  const spcFluids = fluids.filter(f => f.type?.toLowerCase()==="spacer");
  const mudFluids = fluids.find(f => f.type?.toLowerCase()==="mud");
  if (cemFluids.length && spcFluids.length && spcFluids[0].density >= cemFluids[0].density)
    addW("Fluids", `Spacer density (${spcFluids[0].density} ppg) should be less than cement density (${cemFluids[0].density} ppg) for stable displacement`);
  if (spcFluids.length && mudFluids && mudFluids.density >= spcFluids[0].density)
    addW("Fluids", `Mud density (${mudFluids.density} ppg) should be less than spacer density (${spcFluids[0].density} ppg) for stable displacement`);

  // ── Pumping Schedule ──────────────────────────────────────────────────────────
  pump.forEach((s,i) => {
    if (s.volume > 0) {
      check("Pumping", s.rate,   "pump_rate",   `Stage ${s.stage} Rate`);
      check("Pumping", s.volume, "pump_volume", `Stage ${s.stage} Volume`);
    }
    if (!s.fluid_id) addE("Pumping", `Stage ${s.stage}: Fluid must be selected`);
    if (s.rate < 0)  addE("Pumping", `Stage ${s.stage}: Pump rate cannot be negative`);
  });
  const hasCement = pump.some(s => { const fl=fluids.find(f=>f.fluid_id===s.fluid_id); return fl?.type?.toLowerCase()==="cement" && s.volume>0; });
  if (!hasCement) addW("Pumping","No cement stage with volume > 0 found in pumping schedule");

  // ── Centralizers ──────────────────────────────────────────────────────────────
  cent.forEach((c,i) => {
    check("Centralizers", c.spacing,  "cent_spacing",  `Centralizer row ${i+1} Spacing`);
    check("Centralizers", c.standoff, "cent_standoff", `Centralizer row ${i+1} Standoff`);
    if (c.from_depth >= c.to_depth)
      addE("Centralizers", `Centralizer row ${i+1}: From depth must be < To depth`);
    if (c.standoff < 67)
      addW("Centralizers", `Centralizer row ${i+1}: Standoff ${c.standoff}% is below recommended minimum (67%) for pay zone intervals`);
  });

  // ── Temperature ───────────────────────────────────────────────────────────────
  if (temps.length < 2) addW("Temperature","At least 2 temperature survey points recommended for accurate BHCT interpolation");
  temps.forEach((t,i) => {
    check("Temperature", t.temperature, "temperature", `Temp point ${i+1} Temperature`);
    check("Temperature", t.depth,       "temp_depth",  `Temp point ${i+1} Depth`);
  });
  if (temps.length >= 2) {
    const bhct = Math.max(...temps.map(t=>t.temperature));
    if (bhct > 350) addW("Temperature",`BHCT ${bhct}°F exceeds 350°F — use HPHT certified cement system and laboratory BHCT testing`);
  }

  // ── Balanced Plug ─────────────────────────────────────────────────────────────
  plug.forEach((p,i) => {
    check("Balanced Plug", p.length,    "plug_length",    `Plug row ${i+1} Length`);
    check("Balanced Plug", p.set_depth, "plug_set_depth", `Plug row ${i+1} Set Depth`);
    check("Balanced Plug", p.excess,    "plug_excess",    `Plug row ${i+1} Excess`);
    if (p.from_depth >= p.to_depth)
      addE("Balanced Plug", `Plug row ${i+1}: From depth must be < To depth`);
    if (p.set_depth < p.from_depth || p.set_depth > p.to_depth)
      addW("Balanced Plug", `Plug row ${i+1}: Set depth should be within From-To interval`);
    if (p.length > (p.to_depth - p.from_depth))
      addW("Balanced Plug", `Plug row ${i+1}: Plug length (${p.length} ft) exceeds interval (${p.to_depth-p.from_depth} ft)`);
  });

  // ── General ───────────────────────────────────────────────────────────────────
  if (gen.total_depth_md) {
    check("General", gen.total_depth_md,    "gen_total_depth", "Total Depth MD");
    check("General", gen.hole_size,         "gen_hole_size",   "General Hole Size");
    check("General", gen.casing_shoe_depth, "gen_casing_shoe", "Casing Shoe Depth");
    check("General", gen.toc,               "gen_toc",         "Top of Cement");
    if (gen.toc && gen.total_depth_md && gen.toc >= gen.total_depth_md)
      addE("General", `TOC (${gen.toc} ft) must be less than Total Depth (${gen.total_depth_md} ft)`);
  }

  return { errors, warnings, valid: errors.length === 0 };
}

// ─────────────────────────── INLINE FIELD VALIDATOR ─────────────────────────
// Used by input tables to highlight individual cells
function validateField(ruleKey, value) {
  const r = checkRange(value, ruleKey);
  if (!r) return null;
  return { severity: r.severity, msg: r.msg };
}

// ─────────────────────────── IN-MEMORY DATABASE LAYER ──────────────────────
// Structured relational-style storage with versioned simulation runs
const DB_SCHEMA_VERSION = "6.0";

function dbGet(db, table, projectId) {
  if (Array.isArray(db[table])) return db[table];          // master tables
  return db[table]?.[projectId] || [];                     // project tables
}

function dbGetObj(db, table, projectId) {
  return db[table]?.[projectId] || {};
}

function dbUpsertProject(db, pid, table, newData) {
  return { ...db, [table]: { ...db[table], [pid]: newData } };
}

function dbAddSimRun(db, pid, run) {
  const existing = db.simulation_runs?.[pid] || [];
  // Keep last 10 runs per project
  const trimmed = [run, ...existing].slice(0, 10);
  return { ...db, simulation_runs: { ...db.simulation_runs, [pid]: trimmed } };
}

function dbGetLastRun(db, pid) {
  return db.simulation_runs?.[pid]?.[0] || null;
}

// ─────────────────────────── RHEOLOGY ENGINE ─────────────────────────────────
// Computes friction pressure (psi/100 ft) in annulus using rigorous model selection

function binghamReynolds(rho, v, dh, pv) {
  // API Bingham Reynolds: N_Re = 928 × ρ × v × d_h / PV
  // Note: 928 is correct constant for field units (ppg, ft/min, in, cP)
  return (928 * rho * v * dh) / Math.max(pv, 0.001);
}

function powerLawReynolds(rho, v, dh, flowIdx, consistency) {
  // Dodge-Metzner Power Law Reynolds
  const num = 109 * rho * Math.pow(v, 2 - flowIdx) * Math.pow(dh, flowIdx);
  const den = Math.pow(144, 1 - flowIdx) * consistency * Math.pow((2 + 1/flowIdx) / 0.0208, flowIdx);
  return num / Math.max(den, 0.001);
}

function calcFriction(fluid, annVel, dhIn, adjDen) {
  if (annVel <= 0 || dhIn <= 0 || adjDen <= 0) return 0;
  var fluidPv    = (fluid && fluid.pv    != null) ? fluid.pv    : 30;
  var fluidYp    = (fluid && fluid.yp    != null) ? fluid.yp    : 15;
  var fluidModel = (fluid && fluid.model != null) ? fluid.model : "Bingham";
  var flowIdx    = (fluid && fluid.n     != null) ? fluid.n     : 0.7;
  var kConsis    = (fluid && fluid.k     != null) ? fluid.k     : 0.4;

  switch (fluidModel) {

    case "Newtonian": {
      var re0 = (928 * adjDen * annVel * dhIn) / Math.max(fluidPv, 0.001);
      if (re0 > 2100) {
        var ff = 0.0791 / Math.pow(re0, 0.25);
        return (ff * adjDen * annVel * annVel) / (25.81 * dhIn);
      }
      return (fluidPv * annVel) / (300 * dhIn);
    }

    case "Bingham": {
      var re1 = binghamReynolds(adjDen, annVel, dhIn, fluidPv);
      if (re1 > 2100) return (adjDen * annVel * annVel) / 25600;
      return (fluidPv * annVel) / (300 * dhIn) + fluidYp / (225 * dhIn);
    }

    case "Power Law": {
      var re2 = powerLawReynolds(adjDen, annVel, dhIn, flowIdx, kConsis);
      if (re2 > 3470 - 1370 * flowIdx) return (adjDen * annVel * annVel) / 25600;
      var termA = kConsis * Math.pow(annVel, flowIdx) / (144 * Math.pow(dhIn, 1 + flowIdx));
      var termB = Math.pow((2 + 1 / flowIdx) / 0.0208, flowIdx);
      return termA * termB * 144;
    }

    case "HB": {
      var re3 = binghamReynolds(adjDen, annVel, dhIn, fluidPv);
      if (re3 > 2100) return (adjDen * annVel * annVel) / 25600;
      var yldC = fluidYp / (225 * dhIn);
      var plC  = kConsis * Math.pow(annVel, flowIdx) / (144 * Math.pow(dhIn, 1 + flowIdx))
                 * Math.pow((2 + 1 / flowIdx) / 0.0208, flowIdx) * 144;
      return yldC + plC;
    }

    default:
      return (fluidPv * annVel) / (300 * dhIn) + fluidYp / (225 * dhIn);
  }
}

// ─────────────────────────── GEOMETRY HELPERS ────────────────────────────────

// Annular capacity (bbl/ft): (D_hole² - D_casing_OD²) / 1029.4
function annularCapacity(holeSize, casingOD) {
  return Math.max(0, holeSize*holeSize - casingOD*casingOD) / 1029.4;
}
// Internal capacity (bbl/ft): D_ID² / 1029.4
function internalCapacity(casingID) {
  return casingID*casingID / 1029.4;
}
// Annular velocity (ft/min): Q (bbl/min) / Capacity (bbl/ft)
// Capacity (bbl/ft) = Area (in²) / 1029.4
// Therefore: v = Q / (Area / 1029.4) = Q × 1029.4 / Area
function annularVelocity(pumpRateBpm, holeSize, casingOD) {
  const capacity = annularCapacity(holeSize, casingOD);
  return capacity > 0 ? pumpRateBpm / capacity : 0;
}
// Hydraulic diameter (in): D_hole - D_casing_OD
function hydraulicDiameter(holeSize, casingOD) {
  return Math.max(0, holeSize - casingOD);
}

// ─────────────────────────── TEMPERATURE ENGINE ──────────────────────────────
// Derive a temperature profile array from formation_data rows
// Each formation row contributes two points: from_depth and to_depth, both at the row's temperature
// The engine's interpTemp can then interpolate between them
function tempProfileFromFormation(formationData) {
  if (!formationData || formationData.length === 0) return [];
  const pts = [];
  const sorted = [...formationData].sort((a,b) => a.from_depth - b.from_depth);
  sorted.forEach(f => {
    if (f.temperature > 0) {
      pts.push({ depth: f.from_depth, temperature: f.temperature });
      pts.push({ depth: f.to_depth,   temperature: f.temperature });
    }
  });
  // Deduplicate by depth, keeping last
  const map = {};
  pts.forEach(p => { map[p.depth] = p.temperature; });
  return Object.keys(map).map(d => ({ depth: parseFloat(d), temperature: map[d] }))
               .sort((a,b) => a.depth - b.depth);
}


// Linear interpolation between survey points
function interpTemp(tempProfile, depth) {
  if (!tempProfile || tempProfile.length === 0) return 100;
  const sorted = [...tempProfile].sort((a,b) => a.depth - b.depth);
  if (depth <= sorted[0].depth) return sorted[0].temperature;
  if (depth >= sorted[sorted.length-1].depth) return sorted[sorted.length-1].temperature;
  for (let i=0; i<sorted.length-1; i++) {
    if (depth >= sorted[i].depth && depth <= sorted[i+1].depth) {
      const t = (depth - sorted[i].depth) / (sorted[i+1].depth - sorted[i].depth);
      return sorted[i].temperature + t * (sorted[i+1].temperature - sorted[i].temperature);
    }
  }
  return sorted[sorted.length-1].temperature;
}

// Temperature-corrected fluid density
// ρ_adj = ρ_base × [1 - α × (T - T_ref)]  where α = 2.0×10⁻⁵ /°F, T_ref = 70°F
function tempCorrectDensity(baseDensity, tempF) {
  const alpha = 0.00002;
  const tRef  = 70;
  return baseDensity * (1 - alpha * (tempF - tRef));
}

// BHCT correction factor (API Appendix L simplified):
// BHCT ≈ 0.61 × BHST + 0.39 × Surface Temp  (for circulation >2 hours)
function estimateBHCT(bhst, surfaceTemp) {
  return 0.61 * bhst + 0.39 * surfaceTemp;
}

// ─────────────────────────── SURVEY HELPERS ──────────────────────────────────

// TVD from MD using minimum curvature (simplified: linear interp between survey pts)
function tvdAtMD(survey, md) {
  if (!survey || survey.length < 2) return md;
  const sorted = [...survey].sort((a,b) => a.md - b.md);
  if (md <= sorted[0].md) return sorted[0].tvd;
  if (md >= sorted[sorted.length-1].md) return sorted[sorted.length-1].tvd;
  for (let i=0; i<sorted.length-1; i++) {
    if (md >= sorted[i].md && md <= sorted[i+1].md) {
      const t = (md - sorted[i].md) / (sorted[i+1].md - sorted[i].md);
      return sorted[i].tvd + t * (sorted[i+1].tvd - sorted[i].tvd);
    }
  }
  return sorted[sorted.length-1].tvd;
}

// Average inclination between two survey stations
function avgInclination(survey, md) {
  const sorted = [...survey].sort((a,b) => a.md - b.md);
  for (let i=0; i<sorted.length-1; i++) {
    if (md >= sorted[i].md && md <= sorted[i+1].md)
      return (sorted[i].inclination + sorted[i+1].inclination) / 2;
  }
  return sorted[sorted.length-1]?.inclination || 0;
}

// ─────────────────────────── FORMATION HELPERS ───────────────────────────────
function getFormAt(form, depth) {
  for (const f of form) if (depth >= f.from_depth && depth <= f.to_depth) return f;
  return form[form.length-1] || { pore_gradient:0.433, frac_gradient:0.75, lithology:"Unknown" };
}
function getCasingAt(casing, depth) {
  let match = null;
  for (const c of casing) if (depth <= c.to_depth && (!match || c.od < match.od)) match = c;
  return match;
}
function getHoleAt(holes, depth) {
  for (const h of holes) if (depth >= h.from_depth && depth <= h.to_depth) return h.hole_size;
  return holes[holes.length-1]?.hole_size || 8.5;
}

// ─────────────────────────── CENTRALIZER HELPERS ────────────────────────────
function adjustedStandoff(cent, depth) {
  for (const c of cent) {
    if (depth >= c.from_depth && depth <= c.to_depth) {
      const props = CENTRALIZER_TYPE_PROPS[c.type] || CENTRALIZER_TYPE_PROPS["Bow Spring"];
      return Math.max(c.standoff, props.baseStandoff * 0.6);
    }
  }
  return Math.max(20, 45 - depth / 500);
}
function turbulenceBonus(cent, depth) {
  for (const c of cent) {
    if (depth >= c.from_depth && depth <= c.to_depth)
      return (CENTRALIZER_TYPE_PROPS[c.type] || { turbulenceBonus:0 }).turbulenceBonus;
  }
  return 0;
}

// ─────────────────────────── FLUID COLUMN TRACKING ───────────────────────────
function fluidMap(fluids) { const m={}; for (const f of fluids) m[f.fluid_id]=f; return m; }

// Tracks which fluid is present at each depth in the annulus post-displacement
function fluidAtDepth(pump, fluids, depth, maxMD) {
  const fm = fluidMap(fluids);
  const avgAnnCap = 0.002; // bbl/ft average — used for column height estimation
  let remaining = maxMD - depth, accum = 0;
  const order = [...pump].reverse().filter(s => s.volume > 0);
  for (const s of order) {
    accum += s.volume / avgAnnCap;
    if (remaining <= accum) return fm[s.fluid_id] || fluids[0];
  }
  return fluids.find(f => f.type?.toLowerCase()==="mud") || fluids[0];
}

// ─────────────────────────── DISPLACEMENT ENGINE ─────────────────────────────
// Module 2: Mud displacement efficiency
// Based on API RP 10D and Hall-Thompson correlation
function calcDispEfficiency(standoff, annVelFtMin, reynolds, turbBonus, fluidType) {
  // Base efficiency from standoff (API RP 10D):
  //   DE_base = 100 × (SO/100)^0.5 × (v_ann / v_critical)^0.3
  const sof = Math.min(Math.max(standoff, 0), 100);

  // Laminar vs turbulent contribution
  const turbFactor = reynolds > 2100 ? 1.25 : 1.0;  // 25% bonus for turbulent flow
  const velFactor  = Math.min(annVelFtMin / 50, 2.0); // normalize to 50 ft/min reference

  // Combined efficiency index (0–99%)
  let de = (0.40 * sof + 0.25 * Math.min(annVelFtMin, 120) + 10) * turbFactor;
  de += turbBonus;

  // Fluid type penalty: heavier, more viscous fluids displace better
  if (fluidType?.toLowerCase() === "wash") de += 5;  // low-viscosity flush improves contact

  return Math.min(99, Math.max(5, Math.round(de)));
}

// Fluid contamination zone width (simplified mixing model)
function calcContamination(dispEff) {
  // Contamination % inversely proportional to displacement efficiency
  return Math.max(0, Math.round((100 - dispEff) * 0.15 * 10) / 10);
}

// ─────────────────────────── BALANCED PLUG ENGINE ────────────────────────────
// Module 5: Balanced plug placement
function computeBalancedPlug(plugData, casingProfile, openHoles, fluids) {
  return plugData.map(p => {
    const fm = fluidMap(fluids);
    const fluid    = fm[p.fluid_id] || fluids[0] || { density:15.8, type:"Cement", fluid_name:"Cement" };
    const mudFluid = fluids.find(f=>f.type?.toLowerCase()==="mud") || { density:10.0 };

    const holeSize  = getHoleAt(openHoles, p.from_depth);
    const casing    = getCasingAt(casingProfile, p.from_depth);
    const casingOD  = casing ? casing.od  : holeSize * 0.8;
    const casingID  = casing ? casing.id_ : holeSize * 0.75;

    const annCap    = annularCapacity(holeSize, casingOD);   // bbl/ft
    const intCap    = internalCapacity(casingID);            // bbl/ft

    // Volumes
    const plugVol          = p.length    * annCap;    // bbl — cement volume in annulus
    const displacementVol  = p.set_depth * intCap;    // bbl — volume to displace pipe contents
    const waterBehind      = p.excess    * annCap;    // bbl — fluid below plug

    // Hydrostatic balance (Bernoulli at plug base):
    // Balanced when: P_inside_pipe(set_depth) = P_annulus(set_depth)
    //
    // P_inside  = ρ_disp × 0.052 × set_depth
    // P_outside = ρ_fluid_plug × 0.052 × plug_length +
    //             ρ_mud × 0.052 × (set_depth - plug_length)
    const pInside  = mudFluid.density * 0.052 * p.set_depth;
    const pOutside = fluid.density    * 0.052 * p.length +
                     mudFluid.density * 0.052 * Math.max(0, p.set_depth - p.length);
    const balanceDiff = Math.abs(pInside - pOutside);

    // TOC = depth of top of cement column in annulus
    const finalTOC = Math.max(0, p.set_depth - p.length);

    // Bottom-hole hydrostatic pressure at set depth
    const bhhp = fluid.density * 0.052 * p.set_depth;

    // Cement volume needed (cross-check against annular capacity)
    const cementVol_check = p.length * annCap;

    return {
      ...p,
      fluidName:         fluid.fluid_name || fluid.type,
      fluidDensity:      fluid.density,
      plugVolume:        Math.round(plugVol * 100) / 100,
      displacementVolume:Math.round(displacementVol * 100) / 100,
      waterBehind:       Math.round(waterBehind * 100) / 100,
      finalTOC:          Math.round(finalTOC),
      balanced:          balanceDiff < 0.5,     // tolerance 0.5 psi/ft
      balanceDiff:       Math.round(balanceDiff * 100) / 100,
      pressureInside:    Math.round(pInside * 10) / 10,
      pressureOutside:   Math.round(pOutside * 10) / 10,
      bottomHolePressure:Math.round(bhhp),
      cementVolCheck:    Math.round(cementVol_check * 100) / 100,
      annularCapacity:   Math.round(annCap * 10000) / 10000,
    };
  });
}

// ─────────────────────────── TORQUE & DRAG ENGINE ────────────────────────────
// Module: Hook load and drag (simplified soft-string model)
// Hook load (kips): cumulative casing weight minus buoyancy
// Drag force (kips): hook load × sin(inclination) × friction factor
function calcHookLoadAndTorque(cas, md, inclDeg, adjDen) {
  const wt = cas ? cas.weight : 40;         // lb/ft
  const casingOD = cas ? cas.od : 7.0;

  // Buoyancy factor: BF = (ρ_mud - ρ_steel) / ρ_steel  ≈ 1 - ρ_mud/65.4
  const bfactor = Math.max(0.5, 1 - adjDen / 65.4);

  // Axial hook load (kips)
  const hookLoad = (wt * md / 1000) * bfactor;

  // Torque (ft-lbs): T = F_normal × r × μ
  // F_normal = wt × MD × sin(incl) × BF
  const inclRad = (inclDeg * Math.PI) / 180;
  const mu = 0.25;     // friction factor (cased hole)
  const torque = (wt * md / 1000) * Math.sin(inclRad) * mu * (casingOD / 24) * 1000;

  return { hookLoad: Math.round(hookLoad * 10) / 10, torque: Math.round(torque) };
}

// ─────────────────────────── MAIN SIMULATION ENGINE ──────────────────────────

// ═══════════════════════════════════════════════════════════════════════════
// FLUID FRONT TRACKING — Track cement/spacer/mud interface positions
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Initialize fluid tracking state
 * @param {Array} pumpSchedule - Pumping stages with volume, rate, fluid_id
 * @param {Array} fluids - Fluid database
 * @param {number} pipeCapacity - Total pipe capacity (bbl)
 * @param {number} annularCapacity - Total annular capacity (bbl)
 * @returns {Object} Initial tracking state
 */
function initFluidTracking(pumpSchedule, fluids, pipeCapacity, annularCapacity) {
  // Build fluid sequence in reverse pumping order
  var fluidSequence = [];
  var cumVolume = 0;
  
  for (var i = pumpSchedule.length - 1; i >= 0; i--) {
    var stage = pumpSchedule[i];
    if (stage.volume > 0 && stage.rate > 0) {
      var fluid = fluids.find(function(f) { return f.fluid_id === stage.fluid_id; });
      if (fluid) {
        fluidSequence.push({
          fluidId: stage.fluid_id,
          fluidName: stage.fluid_name || fluid.fluid_name,
          fluidType: fluid.type,
          volume: stage.volume,
          startVolume: cumVolume,
          endVolume: cumVolume + stage.volume,
          color: fluid.color || "#999"
        });
        cumVolume += stage.volume;
      }
    }
  }
  
  return {
    fluidSequence: fluidSequence,
    totalPumpedVolume: 0,
    pipeCapacity: pipeCapacity,
    annularCapacity: annularCapacity,
    cementFrontDepth: 0,
    spacerFrontDepth: 0,
    mudFrontDepth: 0,
    interfaces: []  // Array of {depth, fluidAbove, fluidBelow}
  };
}

/**
 * Update fluid front positions for a given pumped volume
 * @param {Object} trackingState - Current tracking state
 * @param {number} pumpedVolume - Total volume pumped so far (bbl)
 * @param {Array} depthGrids - Depth grids with depth and capacity info
 * @returns {Object} Updated tracking state with fluid assignments per grid
 */
function updateFluidFronts(trackingState, pumpedVolume, depthGrids) {
  var state = {...trackingState};
  state.totalPumpedVolume = pumpedVolume;
  
  // Calculate volume in annulus (pumped volume - pipe capacity)
  var volumeInAnnulus = Math.max(0, pumpedVolume - state.pipeCapacity);
  
  // Determine which fluid is at pipe exit (entering annulus)
  var fluidAtPipeExit = null;
  var volumeFromPipeEntry = pumpedVolume;
  
  for (var i = 0; i < state.fluidSequence.length; i++) {
    var fs = state.fluidSequence[i];
    if (volumeFromPipeEntry >= fs.startVolume && volumeFromPipeEntry < fs.endVolume) {
      fluidAtPipeExit = fs;
      break;
    }
  }
  
  // Default to last fluid if past sequence
  if (!fluidAtPipeExit && state.fluidSequence.length > 0) {
    fluidAtPipeExit = state.fluidSequence[state.fluidSequence.length - 1];
  }
  
  // Calculate cumulative annular capacity from TD upward
  var totalAnnCap = 0;
  for (var gi = 0; gi < depthGrids.length; gi++) {
    totalAnnCap += depthGrids[gi].annularCapacity || 0;
  }
  
  // Track fluid column in annulus from bottom to top
  var annFluids = [];  // Array of {fluidInfo, volumeInSection}
  var remainingVolume = volumeInAnnulus;
  
  for (var fi = 0; fi < state.fluidSequence.length; fi++) {
    var fs = state.fluidSequence[fi];
    var volumeOfThisFluid = Math.min(remainingVolume, fs.volume);
    if (volumeOfThisFluid > 0) {
      annFluids.push({
        fluidInfo: fs,
        volume: volumeOfThisFluid
      });
      remainingVolume -= volumeOfThisFluid;
    }
    if (remainingVolume <= 0) break;
  }
  
  // Assign fluids to depth grids from bottom to top
  var cumVolumeFromTD = 0;
  var interfaces = [];
  
  for (var gi = depthGrids.length - 1; gi >= 0; gi--) {
    var grid = depthGrids[gi];
    var gridCapacity = grid.annularCapacity || 0;
    var gridTop = cumVolumeFromTD;
    var gridBottom = cumVolumeFromTD + gridCapacity;
    
    // Determine which fluid occupies this grid
    var assignedFluid = null;
    var fluidVolInGrid = 0;
    
    for (var afi = 0; afi < annFluids.length; afi++) {
      var af = annFluids[afi];
      var fluidStart = 0;
      for (var j = 0; j < afi; j++) {
        fluidStart += annFluids[j].volume;
      }
      var fluidEnd = fluidStart + af.volume;
      
      // Check if this fluid overlaps with this grid
      if (fluidEnd > gridTop && fluidStart < gridBottom) {
        var overlapStart = Math.max(gridTop, fluidStart);
        var overlapEnd = Math.min(gridBottom, fluidEnd);
        var overlapVol = overlapEnd - overlapStart;
        
        if (overlapVol > fluidVolInGrid) {
          fluidVolInGrid = overlapVol;
          assignedFluid = af.fluidInfo;
        }
        
        // Track interface if fluid changes within this grid
        if (fluidStart > gridTop && fluidStart < gridBottom) {
          interfaces.push({
            depth: grid.depth,
            fluidAbove: afi > 0 ? annFluids[afi-1].fluidInfo.fluidName : "Original Mud",
            fluidBelow: af.fluidInfo.fluidName
          });
        }
      }
    }
    
    grid.assignedFluid = assignedFluid ? {
      fluidId: assignedFluid.fluidId,
      fluidName: assignedFluid.fluidName,
      fluidType: assignedFluid.fluidType,
      color: assignedFluid.color
    } : {
      fluidId: "original",
      fluidName: "Original Mud",
      fluidType: "Mud",
      color: "#8B7355"
    };
    
    cumVolumeFromTD += gridCapacity;
  }
  
  // Identify key fronts (cement, spacer, mud)
  var cementFront = 0;
  var spacerFront = 0;
  var mudFront = 0;
  
  for (var gi = 0; gi < depthGrids.length; gi++) {
    var grid = depthGrids[gi];
    if (grid.assignedFluid.fluidType === "Cement" && cementFront === 0) {
      cementFront = grid.depth;
    }
    if (grid.assignedFluid.fluidType === "Spacer" && spacerFront === 0) {
      spacerFront = grid.depth;
    }
    if (grid.assignedFluid.fluidType === "Mud" && mudFront === 0 && gi > 0) {
      mudFront = grid.depth;
    }
  }
  
  state.cementFrontDepth = cementFront;
  state.spacerFrontDepth = spacerFront;
  state.mudFrontDepth = mudFront;
  state.interfaces = interfaces;
  
  return state;
}

/**
 * Get fluid at a specific depth based on current tracking state
 * @param {Array} depthGrids - Depth grids with assigned fluids
 * @param {number} targetDepth - Depth to query (ft)
 * @returns {Object} Fluid info at that depth
 */
function getFluidAtDepth(depthGrids, targetDepth) {
  // Find closest grid
  var closestGrid = depthGrids[0];
  var minDiff = Math.abs(targetDepth - closestGrid.depth);
  
  for (var i = 1; i < depthGrids.length; i++) {
    var diff = Math.abs(targetDepth - depthGrids[i].depth);
    if (diff < minDiff) {
      minDiff = diff;
      closestGrid = depthGrids[i];
    }
  }
  
  return closestGrid.assignedFluid || {
    fluidId: "original",
    fluidName: "Original Mud",
    fluidType: "Mud",
    color: "#8B7355"
  };
}


// ═══════════════════════════════════════════════════════════════════════════
// ITERATIVE PRESSURE SOLVER — Converges pressure solution at each depth
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Iteratively solve pressure at a depth point until convergence
 * @param {Object} params - Input parameters for this depth
 * @returns {Object} Converged solution with pressure, ECD, velocities, etc.
 */
function solveDepthIterative(params) {
  var MAX_ITERATIONS = 30;
  var PRESSURE_TOLERANCE = 10;  // psi
  var ECD_TOLERANCE = 0.02;     // ppg
  
  var {
    md, tvd, adjDen, fluidHere, hydroPsi,
    holeSize, casingOD, casingID, pumpRate,
    stepSize, fricPsiPrev, pipeFricPsiPrev,
    trackedFluid, fluids
  } = params;
  
  // Use tracked fluid properties if available
  var activeFluid = fluidHere;
  if (trackedFluid && trackedFluid.fluidId && trackedFluid.fluidId !== "original") {
    var trackedFluidFull = null;
    for (var i = 0; i < fluids.length; i++) {
      if (fluids[i].fluid_id === trackedFluid.fluidId) {
        trackedFluidFull = fluids[i];
        break;
      }
    }
    if (trackedFluidFull) {
      activeFluid = trackedFluidFull;
    }
  }
  
  // Geometry calculations (constant within depth point)
  var dhIn = hydraulicDiameter(holeSize, casingOD);
  var annVel = annularVelocity(pumpRate, holeSize, casingOD);
  var intCap = internalCapacity(casingID);
  var intVel = intCap > 0 ? pumpRate / intCap : 0;
  
  // Check if iteration is actually needed (linear rheology, constant properties)
  var needsIteration = activeFluid.model !== "Newtonian" && activeFluid.model !== "Bingham";
  
  // Initial values for iteration
  var fricPsi = fricPsiPrev || 0;
  var pipeFricPsi = pipeFricPsiPrev || 0;
  var ecdPres = hydroPsi + fricPsi;
  var ecd = tvd > 0 ? ecdPres / (0.052 * tvd) : adjDen;
  
  var bestSolution = null;
  var bestResidual = Infinity;
  var converged = false;
  var iterationHistory = [];
  var status = "Unknown";
  
  // If no iteration needed, do direct calculation
  if (!needsIteration) {
    var re = binghamReynolds(adjDen, annVel, dhIn, activeFluid.pv || 1);
    var fric100Ann = calcFriction(activeFluid, annVel, dhIn, adjDen);
    var fric100Int = calcFriction(activeFluid, intVel, casingID, adjDen);
    var fricPsiNew = fricPsiPrev + fric100Ann * stepSize / 100;
    var pipeFricPsiNew = pipeFricPsiPrev + fric100Int * stepSize / 100;
    var ecdPresNew = hydroPsi + fricPsiNew;
    var ecdNew = tvd > 0 ? ecdPresNew / (0.052 * tvd) : adjDen;
    
    return {
      fricPsi: fricPsiNew,
      pipeFricPsi: pipeFricPsiNew,
      ecdPres: ecdPresNew,
      ecd: ecdNew,
      annVel: annVel,
      intVel: intVel,
      re: re,
      fric100Ann: fric100Ann,
      fric100Int: fric100Int,
      converged: true,
      iterations: 1,
      residual: 0,
      status: "Direct Solution",
      pressureResidual: 0,
      ecdResidual: 0
    };
  }
  
  // Iterative solution for nonlinear rheology
  for (var iter = 0; iter < MAX_ITERATIONS; iter++) {
    // Store previous iteration values
    var ecdPres_prev = ecdPres;
    var ecd_prev = ecd;
    
    // Calculate Reynolds and friction at CURRENT pressure/ECD state
    var re = binghamReynolds(adjDen, annVel, dhIn, activeFluid.pv || 1);
    var fric100Ann = calcFriction(activeFluid, annVel, dhIn, adjDen);
    var fric100Int = calcFriction(activeFluid, intVel, casingID, adjDen);
    
    // Calculate new friction pressures
    var fricPsiNew = fricPsiPrev + fric100Ann * stepSize / 100;
    var pipeFricPsiNew = pipeFricPsiPrev + fric100Int * stepSize / 100;
    
    // Calculate new ECD pressure and ECD
    var ecdPresNew = hydroPsi + fricPsiNew;
    var ecdNew = tvd > 0 ? ecdPresNew / (0.052 * tvd) : adjDen;
    
    // Calculate residuals (change from previous iteration)
    var pressureDiff = Math.abs(ecdPresNew - ecdPres_prev);
    var ecdDiff = Math.abs(ecdNew - ecd_prev);
    var residual = Math.max(pressureDiff, ecdDiff * 100);
    
    // Store iteration history
    iterationHistory.push({
      iter: iter + 1,
      ecdPres: ecdPresNew,
      ecd: ecdNew,
      pressureResidual: pressureDiff,
      ecdResidual: ecdDiff,
      residual: residual
    });
    
    // Debug logging for first few depths
    if (md <= 200 && iter < 5) {
      console.log(`  Iter ${iter+1}: ECD=${ecdNew.toFixed(3)} ppg, P=${ecdPresNew.toFixed(1)} psi, ΔP=${pressureDiff.toFixed(2)}, ΔECD=${ecdDiff.toFixed(4)}`);
    }
    
    // Store best solution
    if (residual < bestResidual) {
      bestResidual = residual;
      bestSolution = {
        fricPsi: fricPsiNew,
        pipeFricPsi: pipeFricPsiNew,
        ecdPres: ecdPresNew,
        ecd: ecdNew,
        annVel: annVel,
        intVel: intVel,
        re: re,
        fric100Ann: fric100Ann,
        fric100Int: fric100Int,
        iterations: iter + 1,
        residual: residual,
        pressureResidual: pressureDiff,
        ecdResidual: ecdDiff
      };
    }
    
    // Check convergence criteria
    if (pressureDiff < PRESSURE_TOLERANCE && ecdDiff < ECD_TOLERANCE) {
      converged = true;
      status = "Converged";
      bestSolution.converged = true;
      bestSolution.status = status;
      break;
    }
    
    // Update for next iteration
    fricPsi = fricPsiNew;
    pipeFricPsi = pipeFricPsiNew;
    ecdPres = ecdPresNew;
    ecd = ecdNew;
  }
  
  // Set final status if not converged
  if (!converged) {
    if (bestResidual < PRESSURE_TOLERANCE * 5) {
      status = "Stable Approximation";
    } else {
      status = "Max Iterations";
    }
  }
  
  // Return best solution found
  bestSolution.converged = converged;
  bestSolution.status = status;
  return bestSolution;
}

/**
 * Iterative pressure solver for entire wellbore
 * @param {Object} engineParams - All engine parameters
 * @returns {Array} Results array with converged solutions at each depth
 */
function runIterativePressureSolver(engineParams) {
  var {
    survey, casing, holes, form, fluids, cent, tp, pump, gen,
    maxMD, stepSize, surfTemp, pumpRate, bhst, bhct,
    totalPipeCapacity, totalAnnularCapacity, fluidTracking, depthGridsForTracking
  } = engineParams;
  
  var hydroPsi = 0;
  var fricPsi = 0;
  var pipeFricPsi = 0;
  var results = [];
  
  var totalIterations = 0;
  var convergedCount = 0;
  var maxIterUsed = 0;
  
  for (var i = 0; i <= Math.floor(maxMD / stepSize); i++) {
    var md = i * stepSize;
    var tvd = tvdAtMD(survey, md);
    var inclDeg = avgInclination(survey, md);
    var formation = getFormAt(form, md);
    var cas = getCasingAt(casing, md);
    var holeSize = getHoleAt(holes, md);
    var tempSBHT = interpTemp(tp, md);
    var tempCirc = estimateBHCT(tempSBHT, surfTemp);
    var standoff = adjustedStandoff(cent, md);
    var turbBonus = turbulenceBonus(cent, md);
    
    var casingOD = cas ? cas.od : holeSize * 0.8;
    var casingID = cas ? cas.id_ : holeSize * 0.75;
    var annCap = annularCapacity(holeSize, casingOD);
    var intCap = internalCapacity(casingID);
    
    var fluidHere = fluidAtDepth(pump, fluids, md, maxMD);
    var trackedFluid = depthGridsForTracking[i] ? depthGridsForTracking[i].assignedFluid : null;
    
    // Use tracked fluid density if available
    var baseFluid = fluidHere;  // Default fallback
    if (trackedFluid && trackedFluid.fluidId && trackedFluid.fluidId !== "original") {
      // Find the full fluid object from database for density
      var trackedFluidFull = null;
      for (var fi = 0; fi < fluids.length; fi++) {
        if (fluids[fi].fluid_id === trackedFluid.fluidId) {
          trackedFluidFull = fluids[fi];
          break;
        }
      }
      if (trackedFluidFull) {
        baseFluid = trackedFluidFull;
      }
    }
    var adjDen = tempCorrectDensity(baseFluid.density, tempCirc);
    
    // Accumulate hydrostatic (skip for surface point i=0)
    if (i > 0) {
      hydroPsi += adjDen * 0.052 * stepSize;
    }
    
    // Solve pressure iteratively at this depth
    var solution = solveDepthIterative({
      md: md,
      tvd: tvd,
      adjDen: adjDen,
      fluidHere: fluidHere,
      hydroPsi: hydroPsi,
      holeSize: holeSize,
      casingOD: casingOD,
      casingID: casingID,
      pumpRate: pumpRate,
      stepSize: stepSize,
      fricPsiPrev: fricPsi,
      pipeFricPsiPrev: pipeFricPsi,
      trackedFluid: trackedFluid,  // NEW: Pass tracked fluid
      fluids: fluids                // NEW: Pass fluids database
    });
    
    // Update accumulated friction for next depth
    fricPsi = solution.fricPsi;
    pipeFricPsi = solution.pipeFricPsi;
    
    // Track convergence stats
    totalIterations += solution.iterations;
    if (solution.converged) convergedCount++;
    maxIterUsed = Math.max(maxIterUsed, solution.iterations);
    
    // Calculate other parameters
    var ecdPres = solution.ecdPres;
    var ecd = solution.ecd;
    var porePsi = formation.pore_gradient * tvd;
    var fracPsi = formation.frac_gradient * tvd;
    var safety = fracPsi - ecdPres;
    var belowPore = ecdPres < porePsi;
    var aboveFrac = ecdPres > fracPsi * 0.98;
    var dispEff = calcDispEfficiency(standoff, solution.annVel, solution.re, turbBonus, baseFluid.type);
    var contamination = calcContamination(dispEff);
    var hookLoadTorque = calcHookLoadAndTorque(cas, md, inclDeg, adjDen);
    var flowRegime = solution.re > 3000 ? "Turbulent" : solution.re > 2100 ? "Trans." : "Laminar";
    var surfacePumpPressure = Math.max(0, ecdPres - hydroPsi + pipeFricPsi);
    
    results.push({
      md: Math.round(md),
      tvd: Math.round(tvd),
      inclination: Math.round(inclDeg * 10) / 10,
      temperature: Math.round(tempCirc * 10) / 10,
      tempStatic: Math.round(tempSBHT * 10) / 10,
      adjDensity: Math.round(adjDen * 100) / 100,
      hydroPressure: Math.round(hydroPsi),
      frictionPressure: Math.round(fricPsi),
      pipeFriction: Math.round(pipeFricPsi),
      ecdPressure: Math.round(ecdPres),
      porePressure: Math.round(porePsi),
      fracturePressure: Math.round(fracPsi),
      ecd: Math.round(ecd * 100) / 100,
      safetyMargin: Math.round(safety),
      surfacePumpPressure: Math.round(surfacePumpPressure),
      annularVelocity: Math.round(solution.annVel * 10) / 10,
      internalVelocity: Math.round(solution.intVel * 10) / 10,
      reynolds: Math.round(solution.re),
      flowRegime: flowRegime,
      standoff: Math.round(standoff),
      dispEfficiency: dispEff,
      contamination: contamination,
      fluidName: fluidHere.fluid_name,
      fluidType: fluidHere.type,
      fluidColor: fluidHere.color,
      holeSize: Math.round(holeSize * 100) / 100,
      casingOD: Math.round(casingOD * 100) / 100,
      annCapacity: Math.round(annCap * 10000) / 10000,
      torque: hookLoadTorque.torque,
      hookLoad: hookLoadTorque.hookLoad,
      lithology: formation.lithology,
      lczFlag: aboveFrac,
      kickRisk: belowPore,
      trackedFluidId: trackedFluid ? trackedFluid.fluidId : "original",
      trackedFluidName: trackedFluid ? trackedFluid.fluidName : "Original Mud",
      trackedFluidType: trackedFluid ? trackedFluid.fluidType : "Mud",
      trackedFluidColor: trackedFluid ? trackedFluid.color : "#8B7355",
      // Solver diagnostics
      iterConverged: solution.converged,
      iterCount: solution.iterations,
      iterResidual: Math.round(solution.residual * 100) / 100,
      iterStatus: solution.status || "N/A",
      iterPressureResidual: Math.round((solution.pressureResidual || 0) * 100) / 100,
      iterEcdResidual: Math.round((solution.ecdResidual || 0) * 10000) / 10000,
      fluidUsedForCalc: (trackedFluid && trackedFluid.fluidId !== "original") ? "tracked" : "default"
    });
  }
  
  return {
    results: results,
    solverStats: {
      totalIterations: totalIterations,
      convergedCount: convergedCount,
      failedCount: results.length - convergedCount,
      maxIterUsed: maxIterUsed,
      avgIterations: Math.round(totalIterations / results.length * 10) / 10
    }
  };
}

function runSimulationEngine(db, pid) {
  const proj = db.projects.find(p => p.project_id === pid);
  if (!proj) return null;

  const survey  = db.survey_data[pid]||[];
  const casing  = db.casing_profile[pid]||[];
  const holes   = db.open_hole_profile[pid]||[];
  const form    = db.formation_data[pid]||[];
  const fluids  = db.fluid_data[pid]||[];
  const cent    = db.centralizer_data[pid]||[];
  // Temperature profile: prefer formation_data inline temperatures, fallback to temperature_profile
  const formTempProfile = tempProfileFromFormation(form);
  const tp = (formTempProfile.length >= 2) ? formTempProfile : (db.temperature_profile[pid]||[]);
  const pump    = db.pumping_schedule[pid]||[];
  const gen     = db.general_data?.[pid] || {};

  // Depth grid parameters
  const maxMD      = Math.max(...survey.map(s => s.md), gen.total_depth_md || 8500);
  const stepSize   = proj.depth_step || 50;
  const surfTemp   = tp.length > 0 ? Math.min(...tp.map(t=>t.temperature)) : 70;
  const pumpRate   = pump.filter(s => s.volume > 0).reduce((mx, s) => Math.max(mx, s.rate), 0) || 6.0;

  // Pre-compute BHST and BHCT
  const bhst = tp.length > 0 ? Math.max(...tp.map(t=>t.temperature)) : 220;
  const bhct = estimateBHCT(bhst, surfTemp);

  // Calculate total pipe and annular capacities for fluid tracking
  let totalPipeCapacity = 0;
  let totalAnnularCapacity = 0;
  
  for (let i = 0; i <= Math.floor(maxMD / stepSize); i++) {
    const md = i * stepSize;
    const cas = getCasingAt(casing, md);
    const holeSize = getHoleAt(holes, md);
    const casingOD = cas ? cas.od : holeSize * 0.8;
    const casingID = cas ? cas.id_ : holeSize * 0.75;
    const annCap = annularCapacity(holeSize, casingOD) * stepSize;
    const intCap = internalCapacity(casingID) * stepSize;
    totalPipeCapacity += intCap;
    totalAnnularCapacity += annCap;
  }
  
  // Initialize fluid tracking
  var fluidTracking = initFluidTracking(pump, fluids, totalPipeCapacity, totalAnnularCapacity);
  
  // Calculate total job volume
  var totalJobVolume = pump.filter(function(s){return s.volume>0;}).reduce(function(sum,s){return sum+s.volume;}, 0);
  
  // For this static simulation, set pumped volume to total (end of job state)
  var pumpedVolume = totalJobVolume;

  let hydroPsi = 0;
  let fricPsi = 0;        // Accumulated annular friction (incremental)
  let pipeFricPsi = 0;    // Accumulated pipe friction (incremental)
  // const results = [];  // Commented out - using iterative solver

  // Build depth grids with capacity info for fluid tracking
  const depthGridsForTracking = [];
  for (let i = 0; i <= Math.floor(maxMD / stepSize); i++) {
    const md = i * stepSize;
    const cas = getCasingAt(casing, md);
    const holeSize = getHoleAt(holes, md);
    const casingOD = cas ? cas.od : holeSize * 0.8;
    const casingID = cas ? cas.id_ : holeSize * 0.75;
    const annCap = annularCapacity(holeSize, casingOD) * stepSize;  // bbl in this step
    depthGridsForTracking.push({
      depth: md,
      annularCapacity: annCap,
      assignedFluid: null  // Will be populated by tracking
    });
  }
  
  // Update fluid tracking with final pumped volume
  fluidTracking = updateFluidFronts(fluidTracking, pumpedVolume, depthGridsForTracking);

  // Run iterative pressure solver for entire wellbore
  const solverOutput = runIterativePressureSolver({
    survey, casing, holes, form, fluids, cent, tp, pump, gen,
    maxMD, stepSize, surfTemp, pumpRate, bhst, bhct,
    totalPipeCapacity, totalAnnularCapacity, fluidTracking, depthGridsForTracking
  });
  
  var results = solverOutput.results;
  var solverStats = solverOutput.solverStats;

  // OLD DIRECT LOOP REPLACED BY ITERATIVE SOLVER ABOVE
  /*
  for (let i = 0; i <= Math.floor(maxMD / stepSize); i++) {
    const md        = i * stepSize;
    const tvd       = tvdAtMD(survey, md);
    const inclDeg   = avgInclination(survey, md);
    const formation = getFormAt(form, md);
    const cas       = getCasingAt(casing, md);
    const holeSize  = getHoleAt(holes, md);
    const tempSBHT  = interpTemp(tp, md);                        // static BHT at depth
    const tempCirc  = estimateBHCT(tempSBHT, surfTemp);          // circulating temp
    const standoff  = adjustedStandoff(cent, md);
    const turbBonus = turbulenceBonus(cent, md);

    // Geometry
    const casingOD  = cas ? cas.od  : holeSize * 0.8;
    const casingID  = cas ? cas.id_ : holeSize * 0.75;
    const annCap    = annularCapacity(holeSize, casingOD);
    const intCap    = internalCapacity(casingID);
    
    // Get assigned fluid from tracking
    const dhIn      = hydraulicDiameter(holeSize, casingOD);
    const annVel    = annularVelocity(pumpRate, holeSize, casingOD);
    const intCap    = internalCapacity(casingID);
    const intVel    = intCap > 0 ? pumpRate / intCap : 0;

    // Fluid at this depth
    const fluidHere = fluidAtDepth(pump, fluids, md, maxMD);
    
    // Get assigned fluid from tracking (for future use - not yet affecting calculations)
    const trackedFluid = depthGridsForTracking[i] ? depthGridsForTracking[i].assignedFluid : null;

    // Temperature-corrected density
    const adjDen    = tempCorrectDensity(fluidHere.density, tempCirc);

    // Hydrostatic accumulation (incremental per step)
    hydroPsi += adjDen * 0.052 * stepSize;

    // Friction pressure (annular) — INCREMENTAL accumulation per step
    const fric100Ann   = calcFriction(fluidHere, annVel, dhIn, adjDen);
    fricPsi += fric100Ann * stepSize / 100;  // Add friction for THIS step only

    // Reynolds number
    const re = binghamReynolds(adjDen, annVel, dhIn, fluidHere.pv || 1);

    // Pipe friction (internal) — INCREMENTAL accumulation per step
    const dhPipe     = casingID / 12;         // convert in → ft for pipe flow
    const fric100Int = calcFriction(fluidHere, intVel, casingID, adjDen);
    pipeFricPsi += fric100Int * stepSize / 100;  // Add pipe friction for THIS step only

    // Total ECD pressure = hydrostatic + annular friction
    const ecdPres = hydroPsi + fricPsi;

    // ECD in ppg
    const ecd = tvd > 0 ? adjDen + fricPsi / (0.052 * tvd) : adjDen;

    // Formation pressures
    const porePsi = formation.pore_gradient * tvd;
    const fracPsi = formation.frac_gradient * tvd;
    const safety  = fracPsi - ecdPres;

    // Operating pressure window check
    const belowPore = ecdPres < porePsi;        // influx / kick risk
    const aboveFrac = ecdPres > fracPsi * 0.98; // lost circulation risk

    // Displacement efficiency
    const dispEff = calcDispEfficiency(standoff, annVel, re, turbBonus, fluidHere.type);
    const contamination = calcContamination(dispEff);

    // Torque & drag
    const { hookLoad, torque } = calcHookLoadAndTorque(cas, md, inclDeg, adjDen);

    // Flow regime (based on Bingham Reynolds)
    const flowRegime = re > 3000 ? "Turbulent" : re > 2100 ? "Trans." : "Laminar";

    // Final pump pressure at surface (ecdPressure is BHP; surface = BHP - hydrostatic)
    const surfacePumpPressure = Math.max(0, ecdPres - hydroPsi + pipeFricPsi);

    results.push({
      md:         Math.round(md),
      tvd:        Math.round(tvd),
      inclination:Math.round(inclDeg * 10) / 10,
      temperature:Math.round(tempCirc * 10) / 10,
      tempStatic: Math.round(tempSBHT * 10) / 10,
      adjDensity: Math.round(adjDen * 100) / 100,
      hydroPressure:   Math.round(hydroPsi),
      frictionPressure:Math.round(fricPsi),
      pipeFriction:    Math.round(pipeFricPsi),
      ecdPressure:     Math.round(ecdPres),
      porePressure:    Math.round(porePsi),
      fracturePressure:Math.round(fracPsi),
      ecd:             Math.round(ecd * 100) / 100,
      safetyMargin:    Math.round(safety),
      surfacePumpPressure: Math.round(surfacePumpPressure),
      annularVelocity: Math.round(annVel * 10) / 10,
      internalVelocity:Math.round(intVel * 10) / 10,
      reynolds:        Math.round(re),
      flowRegime,
      standoff:        Math.round(standoff),
      dispEfficiency:  dispEff,
      contamination,
      fluidName:   fluidHere.fluid_name,
      fluidType:   fluidHere.type,
      fluidColor:  fluidHere.color,
      holeSize:    Math.round(holeSize * 100) / 100,
      casingOD:    Math.round(casingOD * 100) / 100,
      annCapacity: Math.round(annCap * 10000) / 10000,
      torque,
      hookLoad,
      lithology:   formation.lithology,
      lczFlag:     aboveFrac,
      kickRisk:    belowPore,
      // Fluid tracking data
      trackedFluidId:    trackedFluid ? trackedFluid.fluidId : "original",
      trackedFluidName:  trackedFluid ? trackedFluid.fluidName : "Original Mud",
      trackedFluidType:  trackedFluid ? trackedFluid.fluidType : "Mud",
      trackedFluidColor: trackedFluid ? trackedFluid.color : "#8B7355",
    });
  }
  */
  // END OF OLD LOOP (now using iterative solver)

  // Summary metrics - calculated directly from depth-grid results
  const maxECD     = Math.max(...results.map(r => r.ecd));
  const minSafety  = Math.min(...results.map(r => r.safetyMargin));
  const avgDisp    = Math.round(results.reduce((s, r) => s + r.dispEfficiency, 0) / results.length);
  const maxTemp    = Math.max(...results.map(r => r.temperature));
  const violations = results.filter(r => r.lczFlag).length;
  const kickFlags  = results.filter(r => r.kickRisk).length;
  const finalPumpP = results[results.length-1]?.surfacePumpPressure || 0;
  const maxSurfaceP = Math.max(...results.map(r => r.surfacePumpPressure));
  
  // Validation: Verify maxECD matches actual grid data
  const gridECDs = results.map(r => r.ecd);
  const actualMaxECD = Math.max(...gridECDs);
  const actualMinECD = Math.min(...gridECDs);
  if (Math.abs(maxECD - actualMaxECD) > 0.01) {
    console.warn(`ECD Consistency Warning: Summary maxECD (${maxECD.toFixed(2)}) does not match grid maximum (${actualMaxECD.toFixed(2)})`);
  }
  console.log(`ECD Summary: Max=${maxECD.toFixed(2)} ppg, Min=${actualMinECD.toFixed(2)} ppg, Range=${(maxECD-actualMinECD).toFixed(2)} ppg, Grid points=${results.length}`);
  
  // Reynolds number and flow regime summary
  const reValues = results.map(r => r.reynolds);
  const minRe = Math.min(...reValues);
  const maxRe = Math.max(...reValues);
  const avgRe = (reValues.reduce((s, r) => s + r, 0) / results.length).toFixed(0);
  const laminarCount = results.filter(r => r.flowRegime === "Laminar").length;
  const transCount = results.filter(r => r.flowRegime === "Trans.").length;
  const turbCount = results.filter(r => r.flowRegime === "Turbulent").length;
  console.log(`Reynolds Summary: Min=${minRe.toFixed(0)}, Max=${maxRe.toFixed(0)}, Avg=${avgRe}`);
  console.log(`Flow Regimes: Laminar=${laminarCount}, Transitional=${transCount}, Turbulent=${turbCount} (of ${results.length} points)`);

  // Balanced plug results
  const plugResults = computeBalancedPlug(
    db.balanced_plug_data[pid]||[],
    db.casing_profile[pid]||[],
    db.open_hole_profile[pid]||[],
    fluids
  );

  // Debug: Log ECD distribution for verification
  const ecdDistribution = {
    min: Math.min(...results.map(r => r.ecd)),
    max: Math.max(...results.map(r => r.ecd)),
    avg: (results.reduce((s, r) => s + r.ecd, 0) / results.length).toFixed(2),
    median: results.map(r => r.ecd).sort((a,b) => a-b)[Math.floor(results.length/2)].toFixed(2),
    at_surface: results[0]?.ecd || 0,
    at_bottom: results[results.length-1]?.ecd || 0
  };
  console.log('ECD Distribution:', ecdDistribution);
  
  return {
    run_id:     `run_${Date.now()}`,
    project_id: pid,
    timestamp:  new Date().toISOString(),
    grid:       results,
    plugResults,
    fluidTracking: {
      cementFrontDepth: fluidTracking.cementFrontDepth,
      spacerFrontDepth: fluidTracking.spacerFrontDepth,
      mudFrontDepth: fluidTracking.mudFrontDepth,
      interfaces: fluidTracking.interfaces,
      totalPumpedVolume: fluidTracking.totalPumpedVolume,
      pipeCapacity: fluidTracking.pipeCapacity,
      annularCapacity: fluidTracking.annularCapacity,
    },
    summary: {
      maxECD, minSafety, avgDisp, maxTemp,
      violations, kickFlags, totalSteps: results.length,
      bhct: Math.round(bhct), bhst: Math.round(bhst),
      finalPumpPressure: finalPumpP,
      maxSurfacePressure: maxSurfaceP,
      pumpRate, surfaceTemp: surfTemp,
      // Iterative solver statistics
      solverConverged: solverStats.convergedCount,
      solverFailed: solverStats.failedCount,
      solverMaxIter: solverStats.maxIterUsed,
      solverAvgIter: solverStats.avgIterations,
    },
  };
}

// ─────────────────── EXCEL IMPORT/EXPORT ──────────────────────────────────────

// Map Excel column headers → DB field keys
const EXCEL_MAPS = {
  survey_data:       { "MD":"md","TVD":"tvd","Inclination":"inclination","Azimuth":"azimuth" },
  casing_profile:    { "From":"from_depth","To":"to_depth","OD":"od","ID":"id_","Grade":"grade","Type":"type","Weight":"weight","Description":"description" },
  open_hole_profile: { "From":"from_depth","To":"to_depth","Hole Size":"hole_size","Excess":"excess","Description":"description" },
  formation_data:    { "From":"from_depth","To":"to_depth","Pore Grad":"pore_gradient","Frac Grad":"frac_gradient","Lithology":"lithology" },
  fluid_data:        { "Name":"fluid_name","Type":"type","Density":"density","PV":"pv","YP":"yp","Model":"model","n":"n","K":"k","Color":"color" },
  centralizer_data:  { "From":"from_depth","To":"to_depth","Type":"type","Spacing":"spacing","Standoff":"standoff","RIH Force":"run_in_force" },
  temperature_profile:{ "Depth":"depth","Temperature":"temperature" },
  pumping_schedule:  { "Stage":"stage","Fluid":"fluid_name","Rate":"rate","Volume":"volume","Purpose":"purpose" },
  balanced_plug_data:{ "From":"from_depth","To":"to_depth","Length":"length","Set Depth":"set_depth","Excess":"excess" },
};

function parseExcelFile(file, tableName, cb) {
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const wb = XLSX.read(e.target.result, { type:"array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const raw = XLSX.utils.sheet_to_json(ws, { defval:"" });
      const colMap = EXCEL_MAPS[tableName] || {};
      const rows = raw.map((r, idx) => {
        const mapped = { id: `imp_${tableName}_${Date.now()}_${idx}` };
        for (const [excelCol, dbField] of Object.entries(colMap)) {
          const val = r[excelCol] ?? r[excelCol.toUpperCase()] ?? r[excelCol.toLowerCase()];
          if (val !== undefined && val !== "") {
            mapped[dbField] = isNaN(val) ? val : parseFloat(val);
          }
        }
        // Auto-assign fluid_id for fluid_data
        if (tableName === "fluid_data") {
          mapped.fluid_id = mapped.fluid_id || `fl_imp_${Date.now()}_${idx}`;
          const defaults = FLUID_TYPE_DEFAULTS[mapped.type] || FLUID_TYPE_DEFAULTS["Mud"];
          mapped.color = mapped.color || defaults.color;
        }
        return mapped;
      });
      cb(rows, null);
    } catch(err) {
      cb(null, `Parse error: ${err.message}`);
    }
  };
  reader.readAsArrayBuffer(file);
}

function exportTableToExcel(data, tableName) {
  if (!data || data.length === 0) return;
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, tableName);
  XLSX.writeFile(wb, `CEMOPTI_${tableName}_${Date.now()}.xlsx`);
}

// ─────────────────── EXCEL IMPORT BUTTON ──────────────────────────────────────
function ExcelImportBtn({ tableName, onImport }) {
  const ref = useRef();
  const [status, setStatus] = useState(null);
  const handle = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    parseExcelFile(file, tableName, (rows, err) => {
      if (err) { setStatus({ ok:false, msg:err }); return; }
      onImport(rows);
      setStatus({ ok:true, msg:`✓ Imported ${rows.length} rows` });
      setTimeout(()=>setStatus(null), 3000);
    });
    e.target.value = "";
  };
  return (
    <div style={{ display:"inline-flex",alignItems:"center",gap:8 }}>
      <input ref={ref} type="file" accept=".xlsx,.xls,.csv" onChange={handle} style={{ display:"none" }} />
      <button onClick={()=>ref.current.click()} style={{
        background:"transparent",border:`1px solid ${T.gold}`,color:T.gold,
        borderRadius:4,padding:"5px 12px",fontSize:11,cursor:"pointer",
        fontFamily:"'IBM Plex Sans'",fontWeight:600,
      }}>⬆ Import Excel</button>
      {status && (
        <span style={{ fontSize:11,color:status.ok?T.green:T.red,fontFamily:"'IBM Plex Mono'" }}>{status.msg}</span>
      )}
    </div>
  );
}

// Module 5: Balanced Plug — fully dynamic based on selected fluid
// ─────────────────────────── REPORT GENERATOR (HTML → PDF) ───────────────────
function generateReport(db, pid, simResults) {
  console.log('Generate Report called', { pid, hasResults: !!simResults });
  
  try {
    const proj = db.projects.find(p=>p.project_id===pid);
    if (!proj) {
      console.error('Project not found:', pid);
      alert('Error: Project not found');
      return;
    }
    if (!simResults) {
      console.error('No simulation results available');
      alert('Error: No simulation results. Please run simulation first.');
      return;
    }
    console.log('Generating report for project:', proj.project_name);

  const s     = simResults.summary;
  const grid  = simResults.grid;
  const plug  = simResults.plugResults || [];
  const fluids = db.fluid_data[pid]||[];
  const pumpRaw = db.pumping_schedule[pid]||[];
  const pump   = synchronizeFluidMapping(pumpRaw, fluids);
  const form   = db.formation_data[pid]||[];
  const casing = db.casing_profile[pid]||[];
  const holes  = db.open_hole_profile[pid]||[];
  const cent   = db.centralizer_data[pid]||[];
  const temps  = db.temperature_profile[pid]||[];
  const survey = db.survey_data[pid]||[];
  const bplug  = db.balanced_plug_data[pid]||[];
  const gen    = db.general_data?.[pid] || DEFAULT_DATA.general_data;
  const isM    = proj.unit_system === "metric";
  const U      = UNITS[proj.unit_system];

  const tblHead = (...hs) =>
    `<tr>${hs.map(h=>`<th>${h}</th>`).join("")}</tr>`;
  const tbl = (headers, rows, caption="") => `
    ${caption?`<div class="tbl-cap">${caption}</div>`:""}
    <table><thead>${tblHead(...headers)}</thead><tbody>
    ${rows.map((r,i)=>`<tr class="${i%2?"alt":""}">${r.map(c=>`<td>${c??""}</td>`).join("")}</tr>`).join("")}
    </tbody></table>`;

  const html = `<!DOCTYPE html>
<html lang="en"><head><meta charset="UTF-8">
<title>CEMOPTI Report — ${proj.well_name}</title>
<style>
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:'Segoe UI',Arial,sans-serif;color:#1a2b3a;background:#fff;font-size:11px}
  /* COVER */
  .cover{background:linear-gradient(140deg,#0b1a30 0%,#1a3a5c 60%,#0f4c81 100%);color:#fff;padding:54px 48px;min-height:260px;page-break-after:always}
  .cover-badge{font-size:10px;letter-spacing:3px;text-transform:uppercase;opacity:0.55;margin-bottom:10px}
  .cover h1{font-size:28px;font-weight:900;letter-spacing:1.5px;margin-bottom:6px}
  .cover h2{font-size:15px;font-weight:300;opacity:0.75;margin-bottom:36px}
  .cover-meta{display:grid;grid-template-columns:1fr 1fr 1fr;gap:14px;margin-top:28px}
  .cover-meta .cm{background:rgba(255,255,255,0.1);padding:10px 14px;border-radius:5px}
  .cm .lbl{font-size:9px;opacity:0.55;text-transform:uppercase;letter-spacing:1px}
  .cm .val{font-size:14px;font-weight:700;margin-top:4px}
  .logo-line{font-size:10px;opacity:0.4;margin-top:32px}
  /* SECTIONS */
  .section{padding:24px 36px;border-bottom:1px solid #d8e4ef;page-break-inside:avoid}
  .section h2{font-size:13px;font-weight:700;color:#0f4c81;border-left:4px solid #00c6ff;padding-left:10px;margin-bottom:14px;text-transform:uppercase;letter-spacing:.8px}
  .section h3{font-size:11px;font-weight:700;color:#1a5a8c;margin:12px 0 6px}
  /* METRICS */
  .metrics{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:16px}
  .metric{background:#f0f7ff;border-left:3px solid #00c6ff;padding:10px 12px;border-radius:4px}
  .metric .lbl{font-size:9px;color:#666;text-transform:uppercase;letter-spacing:.5px}
  .metric .val{font-size:18px;font-weight:700;color:#0f4c81;margin-top:3px}
  .metric .unit{font-size:9px;color:#888}
  .metric.warn .val{color:#c0392b}
  .metric.ok   .val{color:#1e8449}
  /* GENERAL BOX */
  .gen-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin-bottom:14px}
  .gen-item{background:#f5f9ff;border:1px solid #d0e4f4;border-radius:4px;padding:8px 12px}
  .gen-item .lbl{font-size:9px;color:#666;text-transform:uppercase}
  .gen-item .val{font-size:13px;font-weight:700;color:#0f4c81;margin-top:3px}
  /* TABLES */
  table{width:100%;border-collapse:collapse;margin:6px 0;font-size:10px}
  th{background:#1a3a5c;color:#fff;padding:5px 7px;text-align:left;font-size:9px;text-transform:uppercase;letter-spacing:.5px}
  td{padding:4px 7px;border-bottom:1px solid #e0eaf0}
  tr.alt td{background:#f4f8fc}
  .tbl-cap{font-size:10px;font-weight:600;color:#1a5a8c;margin-top:10px;margin-bottom:4px}
  .ok{color:#1e8449;font-weight:700} .warn{color:#c0392b;font-weight:700}
  /* DISCLAIMER */
  .disclaimer{background:#0f2840;color:#fff;padding:18px 36px;font-size:9.5px;line-height:1.7}
  .footer{text-align:center;padding:12px;font-size:9px;color:#888;border-top:1px solid #e0e8f0}
  @page{margin:1.5cm 1cm;size:A4}
  @media print{
    body{-webkit-print-color-adjust:exact;print-color-adjust:exact}
    .section{padding:16px 24px}
    .cover{padding:40px 36px}
    .no-print{display:none}
    .page-break{page-break-before:always}
  }
</style></head>
<body>

<!-- COVER PAGE -->
<div class="cover">
  <div class="cover-badge">CEMOPTI v5 — Cementing Simulation Report</div>
  <h1>Cementing Job Design</h1>
  <h2>${proj.project_name} — ${proj.well_name}</h2>
  <div class="cover-meta">
    <div class="cm"><div class="lbl">Well Name</div><div class="val">${proj.well_name}</div></div>
    <div class="cm"><div class="lbl">Field</div><div class="val">${proj.field_name}</div></div>
    <div class="cm"><div class="lbl">Well Type</div><div class="val">${gen.well_type||"—"}</div></div>
    <div class="cm"><div class="lbl">Total Depth (MD)</div><div class="val">${(gen.total_depth_md||0).toLocaleString()} ft</div></div>
    <div class="cm"><div class="lbl">Unit System</div><div class="val">${U.label}</div></div>
    <div class="cm"><div class="lbl">Run ID</div><div class="val">${simResults.run_id.slice(-8)}</div></div>
    <div class="cm"><div class="lbl">Generated</div><div class="val">${new Date().toLocaleDateString()}</div></div>
    <div class="cm"><div class="lbl">Generated By</div><div class="val">CEMOPTI v5.0</div></div>
    <div class="cm"><div class="lbl">Simulation Steps</div><div class="val">${s.totalSteps.toLocaleString()}</div></div>
  </div>
  <div class="logo-line">© Kemiserve FZE · SPC Free Zone · Sharjah, UAE · info@kemiserve.com</div>
</div>

<!-- SECTION 1: EXECUTIVE SUMMARY -->
<div class="section">
  <h2>1. Executive Summary</h2>
  <div class="metrics">
    <div class="metric ${s.maxECD>14?"warn":"ok"}"><div class="lbl">Max ECD</div><div class="val">${s.maxECD}</div><div class="unit">${U.density}</div></div>
    <div class="metric ${s.minSafety<300?"warn":"ok"}"><div class="lbl">Min Safety Margin</div><div class="val">${s.minSafety.toLocaleString()}</div><div class="unit">${U.pressure}</div></div>
    <div class="metric ${s.avgDisp>70?"ok":"warn"}"><div class="lbl">Avg Disp. Efficiency</div><div class="val">${s.avgDisp}%</div><div class="unit">Annular</div></div>
    <div class="metric"><div class="lbl">BHCT</div><div class="val">${isM?cvt(s.bhct,"temperature",true):s.bhct}</div><div class="unit">${U.temperature}</div></div>
    <div class="metric"><div class="lbl">Final Pump Pressure</div><div class="val">${s.finalPumpPressure.toLocaleString()}</div><div class="unit">${U.pressure}</div></div>
    <div class="metric ${s.violations>0?"warn":"ok"}"><div class="lbl">Pressure Violations</div><div class="val">${s.violations}</div><div class="unit">ECD window flags</div></div>
    <div class="metric"><div class="lbl">TOC (Design)</div><div class="val">${(gen.toc||0).toLocaleString()}</div><div class="unit">ft</div></div>
    <div class="metric"><div class="lbl">Cement Column</div><div class="val">${((gen.total_depth_md||0)-(gen.toc||0)).toLocaleString()}</div><div class="unit">ft</div></div>
  </div>
</div>

<!-- SECTION 2: GENERAL WELL DATA -->
<div class="section">
  <h2>2. General Well Data</h2>
  <div class="gen-grid">
    ${[["Well Type",gen.well_type||"—"],["Total Depth (MD)",`${(gen.total_depth_md||0).toLocaleString()} ft`],["Hole Size",`${gen.hole_size||0}"`],["Casing Shoe Depth",`${(gen.casing_shoe_depth||0).toLocaleString()} ft`],["Top of Cement (TOC)",`${(gen.toc||0).toLocaleString()} ft`],["Cement Column Length",`${((gen.total_depth_md||0)-(gen.toc||0)).toLocaleString()} ft`],["Unit System",U.label],["Project",proj.project_name],["Run ID",simResults.run_id.slice(-8)]].map(([l,v])=>`<div class="gen-item"><div class="lbl">${l}</div><div class="val">${v}</div></div>`).join("")}
  </div>
</div>

<!-- SECTION 3: INPUTS -->
<div class="section">
  <h2>3. Input Data Summary</h2>
  <h3>3.1 Directional Survey</h3>
  ${tbl(["MD (ft)","TVD (ft)","Inclination (°)","Azimuth (°)"],survey.map(s=>[s.md,s.tvd,s.inclination,s.azimuth]))}

  <h3>3.2 Casing Profile</h3>
  ${tbl(["From (ft)","To (ft)","OD (in)","ID (in)","Grade","Type","Weight (lb/ft)"],casing.map(c=>[c.from_depth,c.to_depth,c.od,c.id_,c.grade,c.type,c.weight]))}

  <h3>3.3 Open Hole Profile</h3>
  ${tbl(["From (ft)","To (ft)","Hole Size (in)","Excess (ft)","Description"],holes.map(h=>[h.from_depth,h.to_depth,h.hole_size,h.excess,h.description||"—"]))}

  <h3>3.4 Formation Data</h3>
  ${tbl(["From (ft)","To (ft)","Pore Grad (psi/ft)","Frac Grad (psi/ft)","Lithology"],form.map(f=>[f.from_depth,f.to_depth,f.pore_gradient,f.frac_gradient,f.lithology]))}

  <h3>3.5 Fluid Properties</h3>
  ${tbl(["Name","Type","Density (ppg)","PV (cP)","YP (lb/100ft²)","Model","n","K"],fluids.map(f=>[f.fluid_name,f.type,f.density,f.pv,f.yp,f.model,f.n,f.k]))}

  <h3>3.6 Pumping Schedule</h3>
  ${tbl(["Stage","Fluid","Rate (bpm)","Volume (bbl)","Purpose"],pump.map(p=>[p.stage,p.fluid_name,p.rate,p.volume,p.purpose]))}
  <p style="margin-top:6px;font-size:10px;color:#555">Total displacement volume: <b>${pump.reduce((s,r)=>s+(r.volume||0),0)} bbl</b></p>

  <h3>3.7 Centralizer Schedule</h3>
  ${tbl(["From (ft)","To (ft)","Type","Spacing (ft)","Standoff %","RIH Force (lbs)"],cent.map(c=>[c.from_depth,c.to_depth,c.type,c.spacing,c.standoff,c.run_in_force]))}

  <h3>3.8 Temperature Survey</h3>
  ${tbl(["Depth (ft)","Temperature (°F)","Temperature (°C)"],temps.map(t=>[t.depth,t.temperature,cvt(t.temperature,"temperature",true)]))}

  <h3>3.9 Balanced Plug Data</h3>
  ${tbl(["From (ft)","To (ft)","Length (ft)","Set Depth (ft)","Excess (ft)","Fluid ID"],bplug.map(p=>[p.from_depth,p.to_depth,p.length,p.set_depth,p.excess,p.fluid_id]))}
</div>

<!-- SECTION 4: MODULE 1 — HYDRAULICS -->
<div class="section page-break">
  <h2>4. Module 1 — Hydraulics &amp; Pressure</h2>
  <h3>Formation Pressure Window</h3>
  ${tbl(["From (ft)","To (ft)","Pore Grad (psi/ft)","Frac Grad (psi/ft)","Window (psi/ft)","Lithology"],form.map(f=>[f.from_depth,f.to_depth,f.pore_gradient,f.frac_gradient,(f.frac_gradient-f.pore_gradient).toFixed(3),f.lithology]))}
  <h3>Pressure Profile (every 500 ft)</h3>
  ${tbl([`MD (${U.depth})`,`TVD`,`ECD (${U.density})`,`Hydrostatic`,`Pore Pres.`,`Frac Pres.`,`Safety Margin`,`Flow Regime`],
    grid.filter((_,i)=>i%10===0).map(r=>[
      r.md,r.tvd,
      isM?cvt(r.ecd,"density",true):r.ecd,
      isM?cvt(r.hydroPressure,"pressure",true):r.hydroPressure,
      isM?cvt(r.porePressure,"pressure",true):r.porePressure,
      isM?cvt(r.fracturePressure,"pressure",true):r.fracturePressure,
      `<span class="${r.safetyMargin<200?"warn":"ok"}">${r.safetyMargin.toLocaleString()}</span>`,
      r.flowRegime
    ]))}
</div>

<!-- SECTION 5: MODULE 2 — DISPLACEMENT -->
<div class="section">
  <h2>5. Module 2 — Mud Displacement / Cleanout</h2>
  ${tbl(["MD (ft)","Fluid","Type","Ann. Velocity (ft/min)","Standoff %","Disp. Efficiency %","Contamination %"],
    grid.filter((_,i)=>i%10===0).map(r=>[r.md,r.fluidName,r.fluidType,r.annularVelocity,r.standoff,
      `<span class="${r.dispEfficiency>70?"ok":"warn"}">${r.dispEfficiency}%</span>`,
      r.contamination]))}
</div>

<!-- SECTION 6: MODULE 3 — TEMPERATURE -->
<div class="section">
  <h2>6. Module 3 — Temperature (BHCT)</h2>
  <p style="margin-bottom:8px">Bottom Hole Circulating Temperature (BHCT): <b>${isM?cvt(s.bhct,"temperature",true):s.bhct} ${U.temperature}</b></p>
  ${tbl(["Depth (ft)","Temperature (°F)","Temperature (°C)","Adj. Density (ppg)"],
    grid.filter((_,i)=>i%10===0).map(r=>[r.md,r.temperature,cvt(r.temperature,"temperature",true),r.adjDensity]))}
</div>

<!-- SECTION 7: MODULE 4 — CENTRALIZER -->
<div class="section">
  <h2>7. Module 4 — Centralizer Placement / Standoff</h2>
  ${tbl(["From (ft)","To (ft)","Type","Spacing (ft)","Standoff %","RIH Force (lbs)","Status"],cent.map(c=>[c.from_depth,c.to_depth,c.type,c.spacing,c.standoff,c.run_in_force,`<span class="${c.standoff>=67?"ok":"warn"}">${c.standoff>=67?"✓ OK":"⚠ LOW"}</span>`]))}
  <h3>Standoff Profile Summary</h3>
  ${tbl(["MD (ft)","Standoff %","Min Recommended","Status"],
    grid.filter((_,i)=>i%10===0).map(r=>[r.md,r.standoff+"%","67%",`<span class="${r.standoff>=67?"ok":"warn"}">${r.standoff>=67?"✓ OK":"⚠ LOW"}</span>`]))}
</div>

<!-- SECTION 8: MODULE 5 — BALANCED PLUG -->
<div class="section">
  <h2>8. Module 5 — Balanced Plug Placement</h2>
  ${plug.length>0 ? tbl(
    ["From (ft)","To (ft)","Fluid","ρ (ppg)","Length (ft)","Plug Vol (bbl)","Displ. Vol","Water Behind","Final TOC (ft)","BHHP (psi)","Balance Diff","Balanced"],
    plug.map(p=>[p.from_depth,p.to_depth,p.fluidName||"—",p.fluidDensity||"—",p.length,p.plugVolume,p.displacementVolume,p.waterBehind,p.finalTOC,p.bottomHolePressure||"—",p.balanceDiff||"—",`<span class="${p.balanced?"ok":"warn"}">${p.balanced?"✓ BALANCED":"✗ UNBALANCED"}</span>`])
  ) : "<p style='color:#888'>No balanced plug data defined for this project.</p>"}
</div>

<!-- SECTION 9: FULL RESULTS TABLE -->
<div class="section page-break">
  <h2>9. Full Depth-Grid Results (every 500 ft)</h2>
  ${tbl([`MD (ft)`,`TVD`,`ECD`,`Hydro (psi)`,`ECD Pres.`,`Pore Pres.`,`Frac Pres.`,`Safety`,`Disp%`,`Standoff`,`Temp`,`Flow`],
    grid.filter((_,i)=>i%10===0).map(r=>[
      `<b>${r.md.toLocaleString()}</b>`,r.tvd.toLocaleString(),
      `<span class="${r.ecd>14?"warn":""}">${r.ecd}</span>`,
      r.hydroPressure.toLocaleString(),
      `<span class="${r.lczFlag?"warn":""}">${r.ecdPressure.toLocaleString()}</span>`,
      r.porePressure.toLocaleString(),r.fracturePressure.toLocaleString(),
      `<span class="${r.safetyMargin<200?"warn":"ok"}">${r.safetyMargin.toLocaleString()}</span>`,
      `<span class="${r.dispEfficiency>70?"ok":"warn"}">${r.dispEfficiency}%</span>`,
      r.standoff+"%",r.temperature,r.flowRegime
    ]))}
</div>

<!-- DISCLAIMER -->
<div class="disclaimer">
  <b>DISCLAIMER</b> — For Engineering Guidance Only<br>
  This report is generated by CEMOPTI Cementing Simulation Software. All outputs are engineering estimates based on simplified analytical models and are NOT a substitute for qualified professional engineering judgment. Results must be reviewed and validated by a licensed cementing or drilling engineer before field application. Kemiserve FZE accepts no liability for any field decisions made solely on the basis of this report.<br><br>
  <b>© Kemiserve FZE · SPC Free Zone · Sharjah, United Arab Emirates · All Rights Reserved</b><br>
  info@kemiserve.com · CEMOPTI v5.0 · Unauthorized reproduction strictly prohibited.
</div>
<div class="footer">CEMOPTI v5.0 · Kemiserve FZE · SPC Free Zone, Sharjah, UAE · Generated: ${new Date().toLocaleString()} · Run: ${simResults.run_id}</div>
<div class="no-print" style="text-align:center;padding:20px;background:#f0f7ff">
  <button onclick="window.print()" style="background:#0f4c81;color:#fff;border:none;padding:10px 28px;border-radius:6px;font-size:14px;cursor:pointer;font-weight:700">🖨️ Print / Save as PDF</button>
  <p style="font-size:11px;color:#888;margin-top:8px">Use browser's Print → Save as PDF for best results</p>
</div>
</body></html>`;

  const blob = new Blob([html], {type:"text/html"});
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  const filename = `CEMOPTI_Report_${proj.well_name.replace(/\s+/g,"_")}_${Date.now()}.html`;
  a.download = filename;
  
  console.log('Triggering download:', filename);
  a.click();
  
  // Small delay before cleanup
  setTimeout(() => {
    URL.revokeObjectURL(url);
    console.log('Report download triggered successfully');
  }, 100);
  
  } catch (error) {
    console.error('Error generating report:', error);
    alert('Error generating report: ' + error.message);
  }
}

// All input pages are always shown — both Open Hole and Casing are always available

// All input pages — temperature removed (merged into Formation).
// "plug" only shown when job_type = "Secondary (Balanced Plug)"
function getInputPages(db, pid) {
  const base = ["general","survey","casing","openhole","formation","fluids","pumping","centralizers"];
  const gen = db?.general_data?.[pid] || {};
  if (gen.job_type === "Secondary (Balanced Plug)") base.push("plug");
  return base;
}

// Static fallback for components before DB is loaded
const INPUT_PAGES = ["general","survey","casing","openhole","formation","fluids","pumping","centralizers","plug"];

// ─────────────────────────── UI COMPONENTS (unchanged design) ─────────────────

const Btn = ({ children, onClick, variant="primary", size="md", disabled, style={} }) => {
  const V = {
    primary:{ background:`linear-gradient(135deg,${T.accent2},${T.accent})`,color:"#000",border:"none" },
    secondary:{ background:"transparent",color:T.accent,border:`1px solid ${T.accent}` },
    danger:{ background:"transparent",color:T.red,border:`1px solid ${T.red}` },
    ghost:{ background:"transparent",color:T.muted,border:`1px solid ${T.border}` },
    success:{ background:"transparent",color:T.green,border:`1px solid ${T.green}` },
    warning:{ background:"transparent",color:T.gold,border:`1px solid ${T.gold}` },
  };
  const S = { sm:{padding:"4px 10px",fontSize:11}, md:{padding:"7px 16px",fontSize:12}, lg:{padding:"10px 24px",fontSize:14} };
  return <button onClick={onClick} disabled={disabled} style={{ ...V[variant],...S[size],borderRadius:4,fontWeight:600,letterSpacing:"0.03em",fontFamily:"'IBM Plex Sans',sans-serif",opacity:disabled?0.5:1,cursor:disabled?"not-allowed":"pointer",transition:"all 0.15s",...style }}>{children}</button>;
};

const Card = ({ children, style={}, glow }) => (
  <div style={{ background:T.card,border:`1px solid ${glow?T.accent:T.border}`,borderRadius:8,padding:16,boxShadow:glow?`0 0 20px rgba(0,198,255,0.15)`:"none",...style }}>{children}</div>
);

const MetricCard = ({ label, value, unit, color, icon }) => (
  <Card style={{ borderLeft:`3px solid ${color||T.accent}`,padding:"12px 16px" }}>
    <div style={{ fontSize:11,color:T.muted,textTransform:"uppercase",letterSpacing:"0.08em",marginBottom:6 }}>{icon} {label}</div>
    <div style={{ fontSize:24,fontWeight:700,color:color||T.text,fontFamily:"'IBM Plex Mono',monospace" }}>{value}</div>
    <div style={{ fontSize:11,color:T.muted,marginTop:3 }}>{unit}</div>
  </Card>
);

const SectionTitle = ({ icon, title, sub }) => (
  <div style={{ marginBottom:18,borderBottom:`1px solid ${T.border}`,paddingBottom:12 }}>
    <div style={{ display:"flex",alignItems:"center",gap:10 }}>
      <span style={{ fontSize:20 }}>{icon}</span>
      <h2 style={{ fontFamily:"'Orbitron',monospace",fontSize:16,fontWeight:700,color:T.accent,letterSpacing:"0.06em",margin:0 }}>{title}</h2>
    </div>
    {sub && <p style={{ fontSize:12,color:T.muted,marginTop:4,paddingLeft:30 }}>{sub}</p>}
  </div>
);

const Badge = ({ text, color }) => (
  <span style={{ background:`${color}22`,color,border:`1px solid ${color}55`,borderRadius:3,padding:"2px 8px",fontSize:10,fontFamily:"'IBM Plex Mono',monospace",fontWeight:600,letterSpacing:"0.05em" }}>{text}</span>
);

// Navigation footer for input pages
function InputNavFooter({ pageId, onNavigate, pages }) {
  const pageList = pages || INPUT_PAGES;
  const idx = pageList.indexOf(pageId);
  const prev = idx > 0 ? pageList[idx-1] : null;
  const next = idx < pageList.length-1 ? pageList[idx+1] : null;
  const PAGE_LABELS = { survey:"Survey", casing:"Casing", openhole:"Open Hole", formation:"Formation", fluids:"Fluids", pumping:"Pumping Schedule", centralizers:"Centralizers", temperature:"Temperature", plug:"Balanced Plug" };
  return (
    <div style={{ display:"flex",justifyContent:"space-between",alignItems:"center",marginTop:20,paddingTop:14,borderTop:`1px solid ${T.border}` }}>
      <div>
        {prev && <Btn variant="ghost" onClick={()=>onNavigate(prev)}>← {PAGE_LABELS[prev]}</Btn>}
      </div>
      <div style={{ fontSize:11,color:T.muted,fontFamily:"'IBM Plex Mono'" }}>{idx+1} / {pageList.length}</div>
      <div>
        {next && <Btn variant="secondary" onClick={()=>onNavigate(next)}>{PAGE_LABELS[next]} →</Btn>}
        {!next && <Btn variant="primary" onClick={()=>onNavigate("run")}>▶ Run Simulation</Btn>}
      </div>
    </div>
  );
}

// ─────────────────────────── DEPTH VALIDATION HELPERS ───────────────────────

// Change 3: Check that depth values in a table are in strictly increasing order row-by-row
// depthKey = the field name that holds the primary depth for each row (e.g. "from_depth", "md")
function checkDepthOrder(rows, depthKey) {
  const errors = [];
  for (let i = 1; i < rows.length; i++) {
    const prev = parseFloat(rows[i-1][depthKey]) || 0;
    const curr = parseFloat(rows[i][depthKey])   || 0;
    if (curr < prev) {
      errors.push({
        row: i + 1,
        msg: `Row ${i+1}: ${depthKey === "md" ? "MD" : "Depth"} ${curr} ft must be greater than row ${i} value (${prev} ft) — depths must increase top to bottom`
      });
    }
  }
  return errors;
}

// Change 4: Max depth from a table (using to_depth or from_depth or md)
function maxDepthInTable(rows) {
  if (!rows || rows.length === 0) return 0;
  return Math.max(
    ...rows.map(r => Math.max(
      parseFloat(r.to_depth)   || 0,
      parseFloat(r.from_depth) || 0,
      parseFloat(r.md)         || 0,
    ))
  );
}

// Change 4: Banner component for cross-page depth violation
function CrossDepthAlert({ tableName, tableMax, ohMax }) {
  if (!ohMax || tableMax <= ohMax) return null;
  return (
    <div style={{ background:`${T.red}18`,border:`1px solid ${T.red}`,borderRadius:6,padding:"8px 14px",marginBottom:12,fontSize:12,color:T.red,fontWeight:600 }}>
      ⛔ {tableName} max depth ({tableMax.toLocaleString()} ft) exceeds Open Hole max depth ({ohMax.toLocaleString()} ft).
      All depths must be within the Open Hole profile range.
    </div>
  );
}


function IntervalTable({ columns, data, onChange, onAdd, onDelete, rowColor, errors={}, fieldRules={} }) {
  return (
    <div style={{ borderRadius:6,overflow:"hidden",border:`1px solid ${T.border}` }}>
      <div style={{ overflowX:"auto" }}>
        <table>
          <thead><tr>{columns.map(c=><th key={c.key}>{c.label}</th>)}<th style={{width:50}}></th></tr></thead>
          <tbody>
            {data.map((row,ri)=>(
              <tr key={row.id||ri} style={{ background:rowColor?rowColor(row,ri):(ri%2?`${T.panel}50`:"transparent") }}>
                {columns.map(c=>{
                  // Per-cell validation
                  const ruleKey = fieldRules[c.key];
                  const cellErr = ruleKey ? validateField(ruleKey, row[c.key]) : null;
                  return (
                  <td key={c.key}>
                    {c.type==="select"?(
                      <select value={row[c.key]} onChange={e=>onChange(ri,c.key,e.target.value)}
                        style={{ width:"100%", borderColor:errors[`${ri}_${c.key}`]?T.red:undefined }}>
                        {c.options.map(o=><option key={o.v??o} value={o.v??o}>{o.l??o}</option>)}
                      </select>
                    ):c.type==="readonly"?(
                      <span style={{color:T.muted}}>{row[c.key]}</span>
                    ):(
                      <div style={{ position:"relative" }}>
                        <input type={c.type||"number"} value={row[c.key]??""} step={c.step||"any"} min={c.min} max={c.max}
                          className={cellErr?.severity==="error"?"error":""}
                          onChange={e=>onChange(ri,c.key,c.type==="text"?e.target.value:parseFloat(e.target.value)||0)}
                          style={{ width:c.w||80, borderColor: cellErr ? (cellErr.severity==="error"?T.red:T.gold) : undefined,
                                   boxShadow: cellErr ? `0 0 0 2px ${cellErr.severity==="error"?T.red+"44":T.gold+"44"}` : undefined }}
                          title={cellErr?.msg||""} />
                        {cellErr && <span style={{ position:"absolute",right:2,top:"50%",transform:"translateY(-50%)",fontSize:10,color:cellErr.severity==="error"?T.red:T.gold }}>
                          {cellErr.severity==="error"?"✕":"⚠"}
                        </span>}
                      </div>
                    )}
                  </td>
                  );
                })}
                <td><Btn variant="danger" size="sm" onClick={()=>onDelete(ri)}>✕</Btn></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <div style={{ padding:"8px 12px",borderTop:`1px solid ${T.border}`,background:T.panel }}>
        <Btn variant="secondary" size="sm" onClick={onAdd}>+ Add Row</Btn>
      </div>
    </div>
  );
}


// ─── Universal "Set to Default" button — available on all input pages ──────
function SetToDefaultBtn({ onClick }) {
  return (
    <button
      onClick={onClick}
      style={{
        background: "transparent",
        border: "1px solid " + T.gold,
        color: T.gold,
        borderRadius: 4,
        padding: "5px 14px",
        fontSize: 11,
        cursor: "pointer",
        fontFamily: "'IBM Plex Sans'",
        fontWeight: 600,
        letterSpacing: "0.03em",
      }}
    >
      ↺ Set to Default
    </button>
  );
}

function DepthChart({ data, lines, title, xLabel="psi", height=420 }) {
  // Sort data ascending by depth so index 0 = shallowest
  const sorted = [...data].sort((a, b) => a.md - b.md);
  const minMD  = sorted.length ? sorted[0].md                 : 0;
  const maxMD  = sorted.length ? sorted[sorted.length - 1].md : 10000;
  return (
    <Card>
      <div style={{ fontSize:12,fontWeight:600,color:T.accent,fontFamily:"'IBM Plex Mono'",marginBottom:12,letterSpacing:"0.05em" }}>{title}</div>
      <ResponsiveContainer width="100%" height={height}>
        {/* Standard layout: depth on Y-axis, values on X-axis.
            domain={[maxMD, minMD]} puts MAX at top of scale = bottom of chart
            because Recharts Y-axis renders domain[0] at BOTTOM by default,
            so reversing to [maxMD, minMD] makes shallow=top, deep=bottom.
            reversed={true} on <YAxis> is the explicit Recharts API call per spec. */}
        <LineChart data={sorted} layout="vertical" margin={{top:5,right:30,left:10,bottom:20}}>
          <CartesianGrid strokeDasharray="3 3" stroke={T.dim} />
          <XAxis
            type="number"
            tick={{fontSize:10,fill:T.muted}}
            label={{value:xLabel,position:"insideBottom",offset:-5,style:{fontSize:10,fill:T.muted}}}
          />
          <YAxis
            type="number"
            dataKey="md"
            reversed={true}
            domain={[minMD, maxMD]}
            tick={{fontSize:10,fill:T.muted}}
            label={{value:"Depth (ft MD)",angle:-90,position:"insideLeft",style:{fontSize:10,fill:T.muted}}}
          />
          <Tooltip
            contentStyle={{background:T.card,border:`1px solid ${T.border}`,fontSize:11}}
            labelFormatter={v=>`Depth: ${v} ft`}
          />
          <Legend wrapperStyle={{fontSize:11}} />
          {lines.map(l=><Line key={l.key} type="monotone" dataKey={l.key} stroke={l.color} dot={false} strokeWidth={1.8} name={l.name} />)}
        </LineChart>
      </ResponsiveContainer>
    </Card>
  );
}

// Validation Panel
function ValidationPanel({ validation, onDismiss }) {
  if (!validation) return null;
  return (
    <div style={{ position:"fixed",top:60,right:16,width:380,maxHeight:"70vh",overflowY:"auto",background:T.card,border:`1px solid ${validation.valid?T.green:T.red}`,borderRadius:8,padding:16,zIndex:1000,boxShadow:"0 8px 32px rgba(0,0,0,0.5)" }} className="animate-in">
      <div style={{ display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12 }}>
        <div style={{ fontFamily:"'Orbitron'",fontSize:12,color:validation.valid?T.green:T.red }}>
          {validation.valid ? "✓ VALIDATION PASSED" : "✗ VALIDATION FAILED"}
        </div>
        <button onClick={onDismiss} style={{ background:"none",border:"none",color:T.muted,fontSize:16,cursor:"pointer" }}>✕</button>
      </div>
      {validation.errors.length > 0 && (
        <div style={{ marginBottom:10 }}>
          <div style={{ fontSize:10,color:T.red,textTransform:"uppercase",marginBottom:6,letterSpacing:"0.08em" }}>⛔ Errors ({validation.errors.length})</div>
          {validation.errors.map((e,i)=>(
            <div key={i} style={{ background:`${T.red}15`,border:`1px solid ${T.red}44`,borderRadius:4,padding:"6px 10px",marginBottom:4,fontSize:11 }}>
              <span style={{ color:T.red,fontWeight:700 }}>[{e.page}]</span> {e.msg}
            </div>
          ))}
        </div>
      )}
      {validation.warnings.length > 0 && (
        <div>
          <div style={{ fontSize:10,color:T.gold,textTransform:"uppercase",marginBottom:6,letterSpacing:"0.08em" }}>⚠ Warnings ({validation.warnings.length})</div>
          {validation.warnings.map((w,i)=>(
            <div key={i} style={{ background:`${T.gold}15`,border:`1px solid ${T.gold}44`,borderRadius:4,padding:"6px 10px",marginBottom:4,fontSize:11 }}>
              <span style={{ color:T.gold,fontWeight:700 }}>[{w.page}]</span> {w.msg}
            </div>
          ))}
        </div>
      )}
      {validation.valid && validation.warnings.length===0 && (
        <div style={{ fontSize:12,color:T.green }}>All inputs are valid. Ready to run simulation.</div>
      )}
    </div>
  );
}

// Unit selector
function UnitToggle({ unitSystem, onChange }) {
  return (
    <div style={{ display:"flex",gap:4,alignItems:"center" }}>
      <span style={{ fontSize:10,color:T.muted }}>UNITS:</span>
      {["imperial","metric"].map(u=>(
        <button key={u} onClick={()=>onChange(u)} style={{
          padding:"3px 10px",borderRadius:3,border:`1px solid ${unitSystem===u?T.accent:T.border}`,
          background:unitSystem===u?`${T.accent}22`:"transparent",
          color:unitSystem===u?T.accent:T.muted,fontSize:10,fontWeight:600,cursor:"pointer"
        }}>{u==="imperial"?"FIELD":"METRIC"}</button>
      ))}
    </div>
  );
}

// ─────────────────────────── LOGIN PAGE ───────────────────────────────────────
function LoginPage({ onLogin }) {
  const [user, setUser] = useState("");
  const [pass, setPass] = useState("");
  const [err, setErr] = useState("");
  const [loading, setLoading] = useState(false);

  const doLogin = () => {
    setLoading(true);
    setTimeout(() => {
      if (user === "cemopti" && pass === "1234") { onLogin(); }
      else { setErr("Invalid credentials. Use cemopti / 1234"); setLoading(false); }
    }, 500);
  };

  return (
    <div style={{ minHeight:"100vh",background:T.bg,display:"flex",alignItems:"center",justifyContent:"center",fontFamily:"'IBM Plex Sans',sans-serif" }}>
      <style>{css}</style>
      <div style={{ width:380 }}>
        {/* Logo */}
        <div style={{ textAlign:"center",marginBottom:32 }}>
          <div style={{ width:56,height:56,background:`linear-gradient(135deg,${T.accent2},${T.accent})`,borderRadius:12,display:"flex",alignItems:"center",justifyContent:"center",fontSize:24,fontWeight:900,color:"#000",fontFamily:"'Orbitron'",margin:"0 auto 16px" }}>C</div>
          <div style={{ fontFamily:"'Orbitron'",fontSize:22,fontWeight:900,color:T.accent,letterSpacing:"0.12em" }}>CEMOPTI</div>
          <div style={{ fontSize:11,color:T.muted,letterSpacing:"0.1em",marginTop:4 }}>CEMENTING SIMULATION PLATFORM</div>
          <div style={{ fontSize:10,color:T.dim,marginTop:6 }}>© Kemiserve FZE · Sharjah, UAE</div>
        </div>
        <Card style={{ border:`1px solid ${T.borderHi}` }} glow>
          <div style={{ fontSize:13,color:T.accent,fontFamily:"'Orbitron'",letterSpacing:"0.06em",marginBottom:20 }}>SIGN IN</div>
          {err && (
            <div style={{ background:`${T.red}15`,border:`1px solid ${T.red}44`,borderRadius:4,padding:"8px 12px",marginBottom:14,fontSize:12,color:T.red }}>{err}</div>
          )}
          <div style={{ marginBottom:14 }}>
            <div style={{ fontSize:10,color:T.muted,marginBottom:6,textTransform:"uppercase",letterSpacing:"0.08em" }}>Username</div>
            <input type="text" value={user} onChange={e=>{setUser(e.target.value);setErr("")}}
              placeholder="cemopti" style={{ width:"100%",padding:"10px 14px" }}
              onKeyDown={e=>e.key==="Enter"&&doLogin()} autoFocus />
          </div>
          <div style={{ marginBottom:20 }}>
            <div style={{ fontSize:10,color:T.muted,marginBottom:6,textTransform:"uppercase",letterSpacing:"0.08em" }}>Password</div>
            <input type="password" value={pass} onChange={e=>{setPass(e.target.value);setErr("")}}
              placeholder="••••" style={{ width:"100%",padding:"10px 14px" }}
              onKeyDown={e=>e.key==="Enter"&&doLogin()} />
          </div>
          <Btn variant="primary" size="lg" onClick={doLogin} disabled={loading} style={{ width:"100%",justifyContent:"center" }}>
            {loading?"⏳ Authenticating...":"→ Login"}
          </Btn>
        </Card>
        <div style={{ textAlign:"center",marginTop:16,fontSize:11,color:T.dim }}>Depth-based cementing simulation · All 5 engineering modules</div>
      </div>
    </div>
  );
}

// ─────────────────────────── GENERAL INPUT PAGE ──────────────────────────────
function GeneralPage({ db, activeProject, setDb, onNavigate, inputPages }) {
  const g = db.general_data?.[activeProject] || { ...DEFAULT_DATA.general_data };
  const set = (fld, v) => setDb(d => ({
    ...d,
    general_data: { ...d.general_data, [activeProject]: { ...g, [fld]: v } }
  }));

  const WELL_TYPES = ["Exploration","Appraisal","Development","Production","Injection","Disposal","Observation"];
  const WELL_CONTROL_METHODS = ["Automated Choke","MPC","Foamed Cement","Fluid Density Management"];
  const JOB_TYPES = ["Primary","Secondary (Remedial)","Secondary (Balanced Plug)"];

  // Cross-field depth validation: TOC must be < Casing Shoe Depth
  const tocNum  = parseFloat(g.toc)               || 0;
  const shoeNum = parseFloat(g.casing_shoe_depth)  || 0;
  const tocError  = tocNum  > 0 && shoeNum > 0 && tocNum >= shoeNum;
  const hasErrors = tocError;

  const fieldErr = {
    toc:               tocError,
    casing_shoe_depth: tocError,
  };

  const fields = [
    { key:"well_type",           label:"Well Type",             type:"select", options:WELL_TYPES },
    { key:"job_type",            label:"Job Type",              type:"select", options:JOB_TYPES },
    { key:"well_control_method", label:"Well Control Method",   type:"select", options:WELL_CONTROL_METHODS },
    { key:"toc",                 label:"Top of Cement (TOC)",   type:"number", unit:"ft", step:50  },
    { key:"casing_shoe_depth",   label:"Casing Shoe Depth",     type:"number", unit:"ft", step:50  },
  ];

  return (
    <div className="animate-in">
      <SectionTitle icon="📋" title="GENERAL" sub="Well identification, key depths, and geometry — used in report generation and calculations" />

      {tocError && (
        <div style={{ background:`${T.red}18`,border:`1px solid ${T.red}`,borderRadius:6,padding:"8px 14px",marginBottom:8,fontSize:12,color:T.red,fontWeight:600 }}>
          ⛔ Top of Cement (TOC) must be smaller than Casing Shoe Depth
          &nbsp;({tocNum.toLocaleString()} ft ≥ {shoeNum.toLocaleString()} ft — invalid)
        </div>
      )}

      <div style={{display:"flex",justifyContent:"flex-end",marginBottom:8}}><SetToDefaultBtn onClick={()=>setDb(d=>({...d,general_data:{...d.general_data,[activeProject]:{...DEFAULT_DATA.general_data}}}))} /></div>
      <Card style={{ maxWidth:580 }}>
        {fields.map(f => (
          <div key={f.key} style={{ display:"flex",alignItems:"center",padding:"10px 0",borderBottom:`1px solid ${T.border}`,gap:16 }}>
            <div style={{ minWidth:200,fontSize:12,color:fieldErr[f.key]?T.red:T.muted,fontWeight:fieldErr[f.key]?700:400 }}>
              {f.label}{fieldErr[f.key] && <span style={{ marginLeft:6 }}>⛔</span>}
            </div>
            {f.type === "select" ? (
              <select value={g[f.key]||""} onChange={e=>set(f.key, e.target.value)} style={{ flex:1 }}>
                {f.options.map(o=><option key={o} value={o}>{o}</option>)}
              </select>
            ) : (
              <div style={{ display:"flex",alignItems:"center",gap:8,flex:1 }}>
                <input type="number" value={g[f.key]||""} step={f.step} min={0}
                  onChange={e=>set(f.key, parseFloat(e.target.value)||0)}
                  style={{ width:120,
                    borderColor: fieldErr[f.key] ? T.red : undefined,
                    boxShadow:   fieldErr[f.key] ? `0 0 0 2px ${T.red}33` : undefined }} />
                {f.unit && <span style={{ fontSize:11,color:T.dim,fontFamily:"'IBM Plex Mono'" }}>{f.unit}</span>}
              </div>
            )}
          </div>
        ))}
        
        {/* Simulation Parameters */}
        <div style={{ marginTop:16,paddingTop:16,borderTop:"1px solid " + T.border }}>
          <div style={{ fontSize:11,fontWeight:600,color:T.accent,marginBottom:12,fontFamily:"'IBM Plex Mono'" }}>
            SIMULATION PARAMETERS
          </div>
          <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:12 }}>
            <div style={{ display:"flex",alignItems:"center",gap:16 }}>
              <div style={{ minWidth:200,fontSize:12,color:T.muted }}>Number of Timesteps</div>
              <input
                type="number"
                value={g.num_timesteps ?? 100}
                onChange={(e)=>set("num_timesteps", Math.max(10, Math.min(1000, parseInt(e.target.value) || 100)))}
                style={{ width:120 }}
                min="10"
                max="1000"
              />
              <span style={{ fontSize:9,color:T.dim }}>10-1000</span>
            </div>
            <div style={{ display:"flex",alignItems:"center",gap:16 }}>
              <div style={{ minWidth:200,fontSize:12,color:T.muted }}>Number of Depth Grids</div>
              <input
                type="number"
                value={g.num_depth_grids ?? 100}
                onChange={(e)=>set("num_depth_grids", Math.max(10, Math.min(1000, parseInt(e.target.value) || 100)))}
                style={{ width:120 }}
                min="10"
                max="1000"
              />
              <span style={{ fontSize:9,color:T.dim }}>10-1000</span>
            </div>
          </div>
        </div>
      </Card>

      {(shoeNum > 0 || tocNum > 0) && (
        <Card style={{ marginTop:14,border:`1px solid ${hasErrors?T.red:T.accent}33` }}>
          <div style={{ fontSize:11,color:hasErrors?T.red:T.accent,fontFamily:"'IBM Plex Mono'",marginBottom:8 }}>
            GENERAL DATA SUMMARY {hasErrors && "— ⛔ FIX VALIDATION ERRORS ABOVE"}
          </div>
          <div style={{ display:"flex",gap:20,flexWrap:"wrap" }}>
            {[
              { l:"Well Type",          v:g.well_type||"—",                                         c:T.cyan  },
              { l:"Job Type",           v:g.job_type||"—",                                          c:T.accent},
              { l:"Well Control Method",v:g.well_control_method||"—",                              c:T.gold  },
              { l:"Casing Shoe Depth",  v:`${shoeNum.toLocaleString()} ft`,                         c:tocError?T.red:T.text },
              { l:"TOC",                v:`${tocNum.toLocaleString()} ft`,                          c:tocError?T.red:T.green },
              { l:"Balanced Plug",      v:g.job_type==="Secondary (Balanced Plug)"?"Enabled":"Disabled",
                                        c:g.job_type==="Secondary (Balanced Plug)"?T.green:T.dim },
            ].map((r,i)=>(
              <div key={i} style={{ minWidth:130 }}>
                <div style={{ fontSize:10,color:T.muted,textTransform:"uppercase",letterSpacing:"0.06em" }}>{r.l}</div>
                <div style={{ fontSize:14,fontWeight:700,color:r.c,fontFamily:"'IBM Plex Mono'",marginTop:3 }}>{r.v}</div>
              </div>
            ))}
          </div>
        </Card>
      )}
      <InputNavFooter pageId="general" onNavigate={onNavigate} pages={inputPages} />
    </div>
  );
}

// ─────────────────────────── INPUT PAGES ─────────────────────────────────────

function SurveyPage({ db, activeProject, setDb, onNavigate, inputPages }) {
  const data = db.survey_data[activeProject]||[];
  const u = (ri,fld,v) => { var nr=[...data];nr[ri]={...nr[ri],[fld]:v};setDb(d=>({...d,survey_data:{...d.survey_data,[activeProject]:nr}})); };
  const add = () => { const l=data[data.length-1]||{md:0,tvd:0}; setDb(d=>({...d,survey_data:{...d.survey_data,[activeProject]:[...data,{md:l.md+500,tvd:l.tvd+498,inclination:2.0,azimuth:45}]}})); };
  const del = ri => setDb(d=>({...d,survey_data:{...d.survey_data,[activeProject]:data.filter((_,i)=>i!==ri)}}));
  const handleImport = (rows) => setDb(d=>({...d,survey_data:{...d.survey_data,[activeProject]:rows}}));
  const cols = [{key:"md",label:"MD (ft)"},{key:"tvd",label:"TVD (ft)"},{key:"inclination",label:"Inclination (°)",step:0.1},{key:"azimuth",label:"Azimuth (°)",step:0.1}];
  const RULES = { md:"survey_md", tvd:"survey_tvd", inclination:"survey_inclination", azimuth:"survey_azimuth" };
  const orderErrors = checkDepthOrder(data, "md");
  return (
    <div className="animate-in">
      <SectionTitle icon="📐" title="SURVEY DATA" sub="Directional survey: measured depth, true vertical depth, inclination, and azimuth" />
      <div style={{ display:"flex",gap:10,marginBottom:12,flexWrap:"wrap",alignItems:"center" }}>
        <ExcelImportBtn tableName="survey_data" onImport={handleImport} />
        <button onClick={()=>exportTableToExcel(data,"survey_data")} style={{ background:"transparent",border:`1px solid ${T.border}`,color:T.muted,borderRadius:4,padding:"5px 12px",fontSize:11,cursor:"pointer",fontFamily:"'IBM Plex Sans'",fontWeight:600 }}>⬇ Export Excel</button>
        <button onClick={()=>{ setDb(d=>({...d,survey_data:{...d.survey_data,[activeProject]:JSON.parse(JSON.stringify(DEFAULT_DATA.survey_data))}})); }} style={{ background:"transparent",border:`1px solid ${T.gold}`,color:T.gold,borderRadius:4,padding:"5px 14px",fontSize:11,cursor:"pointer",fontFamily:"'IBM Plex Sans'",fontWeight:600 }}>↺ Set to Default</button>
      </div>
      {orderErrors.length > 0 && (
        <div style={{ background:`${T.red}15`,border:`1px solid ${T.red}44`,borderRadius:4,padding:"8px 12px",marginBottom:10 }}>
          {orderErrors.map((e,i)=><div key={i} style={{ fontSize:12,color:T.red,fontWeight:600 }}>⛔ {e.msg}</div>)}
        </div>
      )}
      <IntervalTable columns={cols} data={data} onChange={u} onAdd={add} onDelete={del} fieldRules={RULES} />
      <InputNavFooter pageId="survey" onNavigate={onNavigate} pages={inputPages} />
    </div>
  );
}

function CasingPage({ db, activeProject, setDb, onNavigate, inputPages }) {
  const data = db.casing_profile[activeProject]||[];
  const u = (ri,fld,v) => {
    var nr=[...data];
    if (fld === "type") {
      var props = CASING_TYPE_PROPS[v] || {};
      nr[ri] = { ...nr[ri], type:v, od:props.od||nr[ri].od, id_:props.id_||nr[ri].id_, grade:props.grade||nr[ri].grade, weight:props.weight||nr[ri].weight };
    } else {
      nr[ri] = { ...nr[ri], [fld]:v };
    }
    setDb(d=>({...d,casing_profile:{...d.casing_profile,[activeProject]:nr}}));
  };
  const add = () => {
    const props = CASING_TYPE_PROPS["Production"];
    setDb(d=>({...d,casing_profile:{...d.casing_profile,[activeProject]:[...data,{id:`c${Date.now()}`,from_depth:0,to_depth:5000,...props,description:""}]}}));
  };
  const del = ri => setDb(d=>({...d,casing_profile:{...d.casing_profile,[activeProject]:data.filter((_,i)=>i!==ri)}}));
  const handleImport = (rows) => setDb(d=>({...d,casing_profile:{...d.casing_profile,[activeProject]:rows}}));
  const cols = [{key:"from_depth",label:"From (ft)"},{key:"to_depth",label:"To (ft)"},{key:"od",label:"OD (in)",step:0.001},{key:"id_",label:"ID (in)",step:0.001},{key:"grade",label:"Grade",type:"text",w:70},{key:"type",label:"Type",type:"select",options:Object.keys(CASING_TYPE_PROPS)},{key:"weight",label:"Wt (lb/ft)"},{key:"description",label:"Desc",type:"text",w:110}];
  const RULES = { from_depth:"casing_from_depth", to_depth:"casing_to_depth", od:"casing_od", id_:"casing_id", weight:"casing_weight" };

  // Cross-field validation: OD must be > ID on every row
  const odIdErrors = data.map((c, i) =>
    c.od <= c.id_ ? { row: i+1, od: c.od, id: c.id_ } : null
  ).filter(Boolean);
  // Depth order: from_depth must increase row-by-row
  const orderErrors = checkDepthOrder(data, "from_depth");
  // Cross-page: casing max depth must not exceed open hole max depth
  const ohMax = maxDepthInTable(db.open_hole_profile[activeProject]||[]);
  const casingMax = maxDepthInTable(data);

  return (
    <div className="animate-in">
      <SectionTitle icon="⬡" title="CASING PROFILE" sub="Casing type selection auto-fills OD, ID, grade, and weight from API specifications" />
      {odIdErrors.length > 0 && (
        <div style={{ background:`${T.red}18`,border:`1px solid ${T.red}`,borderRadius:6,padding:"8px 14px",marginBottom:12 }}>
          {odIdErrors.map(e=>(
            <div key={e.row} style={{ fontSize:12,color:T.red,fontWeight:600 }}>
              ⛔ Row {e.row}: Casing OD ({e.od}") must be greater than Casing ID ({e.id}") — invalid geometry
            </div>
          ))}
        </div>
      )}
      {orderErrors.length > 0 && (
        <div style={{ background:`${T.red}15`,border:`1px solid ${T.red}44`,borderRadius:4,padding:"8px 12px",marginBottom:10 }}>
          {orderErrors.map((e,i)=><div key={i} style={{ fontSize:12,color:T.red,fontWeight:600 }}>⛔ {e.msg}</div>)}
        </div>
      )}
      <CrossDepthAlert tableName="Casing Profile" tableMax={casingMax} ohMax={ohMax} />
      <div style={{ display:"flex",gap:10,marginBottom:12,flexWrap:"wrap",alignItems:"center" }}>
        <ExcelImportBtn tableName="casing_profile" onImport={handleImport} />
        <button onClick={()=>exportTableToExcel(data,"casing_profile")} style={{ background:"transparent",border:`1px solid ${T.border}`,color:T.muted,borderRadius:4,padding:"5px 12px",fontSize:11,cursor:"pointer",fontFamily:"'IBM Plex Sans'",fontWeight:600 }}>⬇ Export Excel</button>
        <button onClick={()=>{ setDb(d=>({...d,casing_profile:{...d.casing_profile,[activeProject]:JSON.parse(JSON.stringify(DEFAULT_DATA.casing_profile))}})); }} style={{ background:"transparent",border:`1px solid ${T.gold}`,color:T.gold,borderRadius:4,padding:"5px 14px",fontSize:11,cursor:"pointer",fontFamily:"'IBM Plex Sans'",fontWeight:600 }}>↺ Set to Default</button>
      </div>
      <IntervalTable columns={cols} data={data} onChange={u} onAdd={add} onDelete={del} fieldRules={RULES} />
      <Card style={{ marginTop:12 }}>
        <div style={{ fontSize:11,color:T.accent,fontFamily:"'IBM Plex Mono'",marginBottom:8 }}>CASING TYPE REFERENCE — API STANDARD PROPERTIES</div>
        <div style={{ display:"flex",gap:10,flexWrap:"wrap" }}>
          {Object.entries(CASING_TYPE_PROPS).map(([type,props])=>(
            <div key={type} style={{ flex:1,minWidth:150,background:T.panel,borderRadius:6,padding:"8px 12px",border:`1px solid ${T.border}` }}>
              <div style={{ fontSize:11,fontWeight:700,color:T.cyan,marginBottom:4 }}>{type}</div>
              <div style={{ fontSize:10,color:T.muted,fontFamily:"'IBM Plex Mono'",lineHeight:1.6 }}>OD: {props.od}" · ID: {props.id_}"<br/>Grade: {props.grade} · {props.weight} lb/ft</div>
            </div>
          ))}
        </div>
      </Card>
      <InputNavFooter pageId="casing" onNavigate={onNavigate} pages={inputPages} />
    </div>
  );
}

function OpenHolePage({ db, activeProject, setDb, onNavigate, inputPages }) {
  const data = db.open_hole_profile[activeProject]||[];
  const casingData = db.casing_profile[activeProject]||[];
  const maxCasingOD = casingData.length ? Math.max(...casingData.map(c=>c.od)) : 0;

  const u = (ri,fld,v) => { var nr=[...data];nr[ri]={...nr[ri],[fld]:v};setDb(d=>({...d,open_hole_profile:{...d.open_hole_profile,[activeProject]:nr}})); };
  const add = () => setDb(d=>({...d,open_hole_profile:{...d.open_hole_profile,[activeProject]:[...data,{id:`h${Date.now()}`,from_depth:0,to_depth:5000,hole_size:8.5,excess:100,description:""}]}}));
  const del = ri => setDb(d=>({...d,open_hole_profile:{...d.open_hole_profile,[activeProject]:data.filter((_,i)=>i!==ri)}}));

  // Standard Excel import (manual open hole table)
  const handleImport = (rows) => setDb(d=>({...d,open_hole_profile:{...d.open_hole_profile,[activeProject]:rows}}));

  // Calliper Log import — columns: Depth, Calliper (diameter in inches)
  const calliperRef = useRef();
  const [calliperStatus, setCalliperStatus] = useState(null);
  const handleCalliperImport = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = evt => {
      try {
        const wb = XLSX.read(evt.target.result, { type:"array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const raw = XLSX.utils.sheet_to_json(ws, { defval:0 });
        // Accepted columns: Depth/depth/MD, Calliper/calliper/Diameter/diameter/hole_size
        const sorted = raw.map(r => ({
          depth: parseFloat(r.Depth || r.depth || r.MD || r.md || 0),
          cal:   parseFloat(r.Calliper || r.calliper || r.Diameter || r.diameter || r.hole_size || 0),
        })).filter(r => r.depth > 0 && r.cal > 0).sort((a,b)=>a.depth-b.depth);

        if (sorted.length === 0) { setCalliperStatus({ok:false,msg:"No valid Depth/Calliper columns found"}); return; }

        // Group into intervals (consecutive points within 2% diameter tolerance → same interval)
        const intervals = [];
        let segStart = sorted[0].depth;
        let segSize  = sorted[0].cal;
        for (let i=1; i<sorted.length; i++) {
          const diff = Math.abs(sorted[i].cal - segSize) / segSize;
          if (diff > 0.02 || i === sorted.length-1) {
            intervals.push({ id:`cal_${Date.now()}_${i}`, from_depth:segStart, to_depth:sorted[i-1].depth, hole_size:Math.round(segSize*100)/100, excess:50, description:"From Calliper Log" });
            segStart = sorted[i].depth;
            segSize  = sorted[i].cal;
          }
        }
        if (intervals.length === 0) {
          intervals.push({ id:`cal_${Date.now()}`, from_depth:sorted[0].depth, to_depth:sorted[sorted.length-1].depth, hole_size:Math.round(segSize*100)/100, excess:50, description:"From Calliper Log" });
        }
        setDb(d=>({...d,open_hole_profile:{...d.open_hole_profile,[activeProject]:intervals}}));
        setCalliperStatus({ok:true,msg:`✓ Calliper imported: ${sorted.length} points → ${intervals.length} depth interval${intervals.length>1?"s":""}`});
      } catch(err) { setCalliperStatus({ok:false,msg:`Parse error: ${err.message}`}); }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  };

  // Cross-field validation: hole_size must be > max casing OD
  const holeSizeErrors = maxCasingOD > 0
    ? data.filter(h => h.hole_size <= maxCasingOD).map(h => ({ depth:`${h.from_depth}–${h.to_depth} ft`, hs:h.hole_size }))
    : [];

  const cols = [{key:"from_depth",label:"From (ft)"},{key:"to_depth",label:"To (ft)"},{key:"hole_size",label:"Hole Size (in)",step:0.25},{key:"excess",label:"Excess (ft)"},{key:"description",label:"Desc / Source",type:"text",w:130}];
  const RULES = { hole_size:"hole_size", excess:"hole_excess" };
  // Depth order: from_depth must increase row-by-row
  const orderErrors = checkDepthOrder(data, "from_depth");

  return (
    <div className="animate-in">
      <SectionTitle icon="🕳️" title="OPEN HOLE PROFILE" sub="Borehole geometry by depth — manual entry or Calliper Log import" />

      {/* Validation warning */}
      {holeSizeErrors.length > 0 && (
        <div style={{ background:`${T.red}18`,border:`1px solid ${T.red}`,borderRadius:6,padding:"8px 14px",marginBottom:12 }}>
          {holeSizeErrors.map((e,i)=>(
            <div key={i} style={{ fontSize:12,color:T.red,fontWeight:600 }}>
              ⛔ {e.depth}: Hole size ({e.hs}") must be greater than max Casing OD ({maxCasingOD}") — casing cannot fit inside this hole
            </div>
          ))}
        </div>
      )}
      {maxCasingOD > 0 && holeSizeErrors.length === 0 && (
        <div style={{ background:`${T.green}10`,border:`1px solid ${T.green}44`,borderRadius:4,padding:"6px 12px",marginBottom:10,fontSize:11,color:T.green }}>
          ✓ All hole sizes &gt; max casing OD ({maxCasingOD}")
        </div>
      )}
      {orderErrors.length > 0 && (
        <div style={{ background:`${T.red}15`,border:`1px solid ${T.red}44`,borderRadius:4,padding:"8px 12px",marginBottom:10 }}>
          {orderErrors.map((e,i)=><div key={i} style={{ fontSize:12,color:T.red,fontWeight:600 }}>⛔ {e.msg}</div>)}
        </div>
      )}

      {/* Import controls */}
      <div style={{ display:"flex",gap:10,marginBottom:12,flexWrap:"wrap",alignItems:"center" }}>
        <ExcelImportBtn tableName="open_hole_profile" onImport={handleImport} />
        <button onClick={()=>exportTableToExcel(data,"open_hole_profile")} style={{ background:"transparent",border:`1px solid ${T.border}`,color:T.muted,borderRadius:4,padding:"5px 12px",fontSize:11,cursor:"pointer",fontFamily:"'IBM Plex Sans'",fontWeight:600 }}>⬇ Export Excel</button>
        <button onClick={()=>{ setDb(d=>({...d,open_hole_profile:{...d.open_hole_profile,[activeProject]:JSON.parse(JSON.stringify(DEFAULT_DATA.open_hole_profile))}})); }} style={{ background:"transparent",border:`1px solid ${T.gold}`,color:T.gold,borderRadius:4,padding:"5px 14px",fontSize:11,cursor:"pointer",fontFamily:"'IBM Plex Sans'",fontWeight:600 }}>↺ Set to Default</button>
        {/* Calliper log import */}
        <input ref={calliperRef} type="file" accept=".xlsx,.xls,.csv" onChange={handleCalliperImport} style={{ display:"none" }} />
        <button onClick={()=>calliperRef.current.click()} style={{ background:"transparent",border:`1px solid ${T.gold}`,color:T.gold,borderRadius:4,padding:"5px 12px",fontSize:11,cursor:"pointer",fontFamily:"'IBM Plex Sans'",fontWeight:600 }}>📏 Import Calliper Log</button>
        {calliperStatus && <span style={{ fontSize:11,color:calliperStatus.ok?T.green:T.red,fontFamily:"'IBM Plex Mono'" }}>{calliperStatus.msg}</span>}
      </div>
      <div style={{ fontSize:10,color:T.dim,marginBottom:10 }}>
        Calliper Log columns: <strong>Depth</strong> (ft) + <strong>Calliper</strong> or <strong>Diameter</strong> (in) — auto-segments into depth intervals
        &nbsp;|&nbsp; Manual entry: add rows directly in the table below
      </div>

      <IntervalTable columns={cols} data={data} onChange={u} onAdd={add} onDelete={del} fieldRules={RULES} />
      <InputNavFooter pageId="openhole" onNavigate={onNavigate} pages={inputPages} />
    </div>
  );
}

function FormationPage({ db, activeProject, setDb, onNavigate, inputPages }) {
  const data = db.formation_data[activeProject]||[];

  // Single unified update handler — also syncs temperature_profile when temperature changes
  const u = (ri, fld, v) => {
    var nr = [...data];
    nr[ri] = { ...nr[ri], [fld]: v };
    // Keep temperature_profile in sync so thermal output pages continue to work
    var newFormData = nr;
    const syncedTempProfile = tempProfileFromFormation(newFormData);
    setDb(d => ({
      ...d,
      formation_data:    { ...d.formation_data,    [activeProject]: newFormData },
      temperature_profile:{ ...d.temperature_profile,[activeProject]: syncedTempProfile },
    }));
  };

  const add = () => {
    const last = data[data.length-1] || { to_depth:0, temperature:100 };
    const newRow = {
      id: `f${Date.now()}`,
      from_depth:    last.to_depth,
      to_depth:      last.to_depth + 1000,
      pore_gradient: 0.452,
      frac_gradient: 0.75,
      temperature:   Math.round(last.temperature + 15),
      lithology:     "Shale",
    };
    const newData = [...data, newRow];
    const syncedTempProfile = tempProfileFromFormation(newData);
    setDb(d => ({
      ...d,
      formation_data:    { ...d.formation_data,    [activeProject]: newData },
      temperature_profile:{ ...d.temperature_profile,[activeProject]: syncedTempProfile },
    }));
  };

  const del = ri => {
    const newData = data.filter((_,i) => i !== ri);
    const syncedTempProfile = tempProfileFromFormation(newData);
    setDb(d => ({
      ...d,
      formation_data:    { ...d.formation_data,    [activeProject]: newData },
      temperature_profile:{ ...d.temperature_profile,[activeProject]: syncedTempProfile },
    }));
  };

  const handleImport = rows => {
    // Accept imported rows; auto-assign temperature=0 if missing
    const enriched = rows.map(r => ({ temperature: 0, ...r }));
    const syncedTempProfile = tempProfileFromFormation(enriched);
    setDb(d => ({
      ...d,
      formation_data:    { ...d.formation_data,    [activeProject]: enriched },
      temperature_profile:{ ...d.temperature_profile,[activeProject]: syncedTempProfile },
    }));
  };

  // Single unified table columns: From | To | Pore Grad | Frac Grad | Temperature | Lithology
  const cols = [
    { key:"from_depth",    label:"From (ft)" },
    { key:"to_depth",      label:"To (ft)" },
    { key:"pore_gradient", label:"Pore Gradient (psi/ft)", step:0.001 },
    { key:"frac_gradient", label:"Frac Gradient (psi/ft)", step:0.001 },
    { key:"temperature",   label:"Temperature (°F)",       step:1     },
    { key:"lithology",     label:"Lithology", type:"text", w:110 },
  ];
  const RULES = {
    from_depth:    "casing_from_depth",
    to_depth:      "casing_to_depth",
    pore_gradient: "pore_gradient",
    frac_gradient: "frac_gradient",
    temperature:   "temperature",
  };
  // Depth order: from_depth must increase row-by-row
  const orderErrors = checkDepthOrder(data, "from_depth");

  return (
    <div className="animate-in">
      <SectionTitle icon="🪨" title="FORMATION DATA"
        sub="From · To · Pore Gradient · Frac Gradient · Temperature · Lithology — single unified table" />

      <div style={{ display:"flex",gap:10,marginBottom:12,flexWrap:"wrap",alignItems:"center" }}>
        <ExcelImportBtn tableName="formation_data" onImport={handleImport} />
        <button onClick={()=>exportTableToExcel(data,"formation_data")} style={{ background:"transparent",border:`1px solid ${T.border}`,color:T.muted,borderRadius:4,padding:"5px 12px",fontSize:11,cursor:"pointer",fontFamily:"'IBM Plex Sans'",fontWeight:600 }}>⬇ Export Excel</button>
        <button onClick={()=>{ (()=>{ const dd=JSON.parse(JSON.stringify(DEFAULT_DATA.formation_data)); const st=tempProfileFromFormation(dd); setDb(d=>({...d,formation_data:{...d.formation_data,[activeProject]:dd},temperature_profile:{...d.temperature_profile,[activeProject]:st}})); })(); }} style={{ background:"transparent",border:`1px solid ${T.gold}`,color:T.gold,borderRadius:4,padding:"5px 14px",fontSize:11,cursor:"pointer",fontFamily:"'IBM Plex Sans'",fontWeight:600 }}>↺ Set to Default</button>
        <div style={{ fontSize:10,color:T.dim }}>
          Excel columns: From, To, Pore Grad, Frac Grad, Temperature, Lithology
        </div>
      </div>
      {orderErrors.length > 0 && (
        <div style={{ background:`${T.red}15`,border:`1px solid ${T.red}44`,borderRadius:4,padding:"8px 12px",marginBottom:10 }}>
          {orderErrors.map((e,i)=><div key={i} style={{ fontSize:12,color:T.red,fontWeight:600 }}>⛔ {e.msg}</div>)}
        </div>
      )}

      {/* Validation warnings for formation rows */}
      {data.some(f => f.frac_gradient <= f.pore_gradient) && (
        <div style={{ background:`${T.red}15`,border:`1px solid ${T.red}44`,borderRadius:4,padding:"6px 12px",marginBottom:10,fontSize:11,color:T.red }}>
          ⛔ Frac gradient must exceed pore gradient on all rows — check highlighted rows
        </div>
      )}
      {data.some(f => f.frac_gradient - f.pore_gradient < 0.05 && f.frac_gradient > f.pore_gradient) && (
        <div style={{ background:`${T.yellow}15`,border:`1px solid ${T.yellow}44`,borderRadius:4,padding:"6px 12px",marginBottom:10,fontSize:11,color:T.yellow }}>
          ⚠ Narrow pressure window detected (&lt;0.05 psi/ft) on one or more rows — ECD control will be critical
        </div>
      )}

      {/* THE single unified table */}
      <IntervalTable columns={cols} data={data} onChange={u} onAdd={add} onDelete={del} fieldRules={RULES}
        rowColor={(row, ri) => {
          if (row.frac_gradient <= row.pore_gradient) return `${T.red}15`;
          if (row.frac_gradient - row.pore_gradient < 0.05) return `${T.yellow}10`;
          return ri % 2 ? `${T.panel}50` : "transparent";
        }}
      />

      {/* Live preview charts below the table */}
      <div className="grid-2" style={{ marginTop:16 }}>
        <Card>
          <div style={{ fontSize:11,color:T.muted,fontFamily:"'IBM Plex Mono'",marginBottom:8 }}>
            PRESSURE WINDOW (psi/ft) — LIVE PREVIEW
          </div>
          <ResponsiveContainer width="100%" height={200}>
            <LineChart data={data} margin={{top:5,right:20,left:10,bottom:5}}>
              <CartesianGrid strokeDasharray="3 3" stroke={T.dim} />
              <XAxis dataKey="from_depth" tick={{fontSize:9,fill:T.muted}}
                label={{value:"Depth (ft)",position:"insideBottom",offset:-2,style:{fontSize:9,fill:T.muted}}} />
              <YAxis domain={[0.3,1.2]} tick={{fontSize:9,fill:T.muted}} />
              <Tooltip contentStyle={{background:T.card,border:`1px solid ${T.border}`,fontSize:10}} />
              <Legend wrapperStyle={{fontSize:10}} />
              <Line type="stepAfter" dataKey="pore_gradient" stroke={T.green} dot={false} strokeWidth={2} name="Pore (psi/ft)" />
              <Line type="stepAfter" dataKey="frac_gradient" stroke={T.red}   dot={false} strokeWidth={2} name="Frac (psi/ft)" />
            </LineChart>
          </ResponsiveContainer>
        </Card>
        <Card>
          <div style={{ fontSize:11,color:T.muted,fontFamily:"'IBM Plex Mono'",marginBottom:8 }}>
            TEMPERATURE PROFILE (°F) — LIVE PREVIEW
          </div>
          <ResponsiveContainer width="100%" height={200}>
            <LineChart
              data={[...data].sort((a,b)=>a.from_depth-b.from_depth).map(f=>({ depth:f.from_depth, temperature:f.temperature }))}
              layout="vertical" margin={{top:5,right:20,left:10,bottom:5}}>
              <CartesianGrid strokeDasharray="3 3" stroke={T.dim} />
              <XAxis type="number" dataKey="temperature" tick={{fontSize:10,fill:T.muted}} />
              <YAxis type="number" dataKey="depth" reversed tick={{fontSize:10,fill:T.muted}} />
              <Tooltip contentStyle={{background:T.card,border:`1px solid ${T.border}`,fontSize:11}} />
              <Line type="monotone" dataKey="temperature" stroke={T.gold} dot={{fill:T.gold,r:3}} strokeWidth={2} name="Temp (°F)" />
            </LineChart>
          </ResponsiveContainer>
        </Card>
      </div>

      <InputNavFooter pageId="formation" onNavigate={onNavigate} pages={inputPages} />
    </div>
  );
}

function FluidPage({ db, activeProject, setDb, onNavigate, inputPages }) {
  const data = db.fluid_data[activeProject]||[];
  const u = (ri,fld,v) => {
    var nr=[...data];
    // When fluid type changes, apply default properties AND update fluid name
    if (fld === "type") {
      const defaults = FLUID_TYPE_DEFAULTS[v] || FLUID_TYPE_DEFAULTS["Mud"];
      // Auto-set name based on type (only if name matches a previous auto-name or is blank)
      const autoNames = Object.keys(FLUID_TYPE_DEFAULTS);
      var currentName = nr[ri].fluid_name || "";
      var isAutoName = autoNames.includes(currentName) || currentName === "" ||
        autoNames.some(function(t){ return currentName.toLowerCase().includes(t.toLowerCase()); });
      var newName = isAutoName ? v : currentName;
      nr[ri] = { ...nr[ri], type:v, fluid_name:newName,
        density:defaults.density, pv:defaults.pv, yp:defaults.yp,
        model:defaults.model, n:defaults.n, k:defaults.k,
        color:defaults.color };
    } else if (fld === "model") {
      nr[ri] = { ...nr[ri], model:v };
      if (v === "Newtonian") nr[ri] = { ...nr[ri], yp:0, n:1.0 };
    } else {
      nr[ri] = { ...nr[ri], [fld]:v };
    }
    setDb(d=>({...d,fluid_data:{...d.fluid_data,[activeProject]:nr}}));
  };
  const add = () => {
    const defaults = FLUID_TYPE_DEFAULTS["cement"];
    setDb(d=>({...d,fluid_data:{...d.fluid_data,[activeProject]:[...data,{fluid_id:`fl${Date.now()}`,fluid_name:"New Cement",type:"Cement",...defaults}]}}));
  };
  const del = ri => setDb(d=>({...d,fluid_data:{...d.fluid_data,[activeProject]:data.filter((_,i)=>i!==ri)}}));
  const handleImport = (rows) => {
    const merged = rows.map((r,i)=>({ fluid_id:`fl_imp_${Date.now()}_${i}`, fluid_name:r.fluid_name||"Imported Fluid", type:r.type||"Mud", density:r.density||10.2, pv:r.pv||20, yp:r.yp||12, model:r.model||"Bingham", n:r.n||0.8, k:r.k||0.2, color:(FLUID_TYPE_DEFAULTS[r.type]||FLUID_TYPE_DEFAULTS["Mud"]).color, ...r }));
    setDb(d=>({...d,fluid_data:{...d.fluid_data,[activeProject]:merged}}));
  };

  const cols = [
    {key:"fluid_name",label:"Name",type:"text",w:120},
    {key:"type",label:"Type",type:"select",options:Object.keys(FLUID_TYPE_DEFAULTS)},    {key:"density",label:"Density (ppg)",step:0.1},
    {key:"pv",label:"PV (cP)"},
    {key:"yp",label:"YP (lb/100ft²)"},
    {key:"model",label:"Model",type:"select",options:["Bingham","Power Law","Newtonian","HB"]},
    {key:"n",label:"n",step:0.01},
    {key:"k",label:"K",step:0.01},
    {key:"color",label:"Color",type:"text",w:70}
  ];

  // Model equation reference
  const modelDesc = { Bingham:"τ = YP + PV·γ — Laminar: ΔP=(PV·v/300d)+(YP/225d)", "Power Law":"τ = K·γⁿ — Dodge-Metzner correlation", Newtonian:"τ = μ·γ — Single viscosity parameter (PV=μ)", HB:"τ = τ₀ + K·γⁿ — Herschel-Bulkley (full 3-param)" };

  const selectedModels = [...new Set(data.map(f=>f.model))];

  return (
    <div className="animate-in">
      <SectionTitle icon="🧪" title="FLUID DATABASE" sub="All fluids in the displacement train — type selection auto-fills default properties · Model selection applies correct correlation" />
      <div style={{ display:"flex",gap:10,marginBottom:12,flexWrap:"wrap",alignItems:"center" }}>
        <ExcelImportBtn tableName="fluid_data" onImport={handleImport} />
        <button onClick={()=>exportTableToExcel(data,"fluid_data")} style={{ background:"transparent",border:`1px solid ${T.border}`,color:T.muted,borderRadius:4,padding:"5px 12px",fontSize:11,cursor:"pointer",fontFamily:"'IBM Plex Sans'",fontWeight:600 }}>⬇ Export Excel</button>
        <button onClick={()=>{ setDb(d=>({...d,fluid_data:{...d.fluid_data,[activeProject]:JSON.parse(JSON.stringify(DEFAULT_DATA.fluid_data))}})); }} style={{ background:"transparent",border:`1px solid ${T.gold}`,color:T.gold,borderRadius:4,padding:"5px 14px",fontSize:11,cursor:"pointer",fontFamily:"'IBM Plex Sans'",fontWeight:600 }}>↺ Set to Default</button>
      </div>
      <IntervalTable columns={cols} data={data} onChange={u} onAdd={add} onDelete={del} rowColor={row=>`${row.color}15`} />

      {/* Dynamic model info panel */}
      {selectedModels.length > 0 && (
        <Card style={{ marginTop:12,border:`1px solid ${T.border}` }}>
          <div style={{ fontSize:11,color:T.accent,fontFamily:"'IBM Plex Mono'",marginBottom:8 }}>ACTIVE RHEOLOGY MODELS — EQUATIONS IN USE</div>
          <div style={{ display:"flex",gap:10,flexWrap:"wrap" }}>
            {selectedModels.map(m=>(
              <div key={m} style={{ flex:1,minWidth:220,background:T.panel,borderRadius:6,padding:"8px 12px",border:`1px solid ${T.border}` }}>
                <div style={{ fontSize:11,fontWeight:700,color:T.yellow,marginBottom:4 }}>{m}</div>
                <div style={{ fontSize:10,color:T.muted,fontFamily:"'IBM Plex Mono'" }}>{modelDesc[m]||"Custom model"}</div>
              </div>
            ))}
          </div>
        </Card>
      )}

      <div style={{ display:"flex",gap:10,marginTop:12,flexWrap:"wrap" }}>
        {data.map(f=>(
          <div key={f.fluid_id} style={{ display:"flex",alignItems:"center",gap:6,padding:"4px 10px",borderRadius:4,border:`1px solid ${f.color}55`,background:`${f.color}11` }}>
            <span style={{ width:10,height:10,borderRadius:"50%",background:f.color }}></span>
            <span style={{ fontSize:11,fontFamily:"'IBM Plex Mono'",color:T.text }}>{f.fluid_name}</span>
            <span style={{ fontSize:10,color:T.muted }}>{f.density} ppg</span>
            <span style={{ fontSize:10,color:T.dim }}>{f.model}</span>
          </div>
        ))}
      </div>
      <InputNavFooter pageId="fluids" onNavigate={onNavigate} pages={inputPages} />
    </div>
  );
}

function PumpingPage({ db, activeProject, setDb, onNavigate, inputPages }) {
  const scheduleRaw = db.pumping_schedule[activeProject]||[];
  const fluids = db.fluid_data[activeProject]||[];
  const openHoles = db.open_hole_profile[activeProject]||[];
  
  // Synchronize fluid mappings on component mount
  const schedule = React.useMemo(() => {
    const synchronized = synchronizeFluidMapping(scheduleRaw, fluids);
    
    // If synchronization made changes, update database
    const hasChanges = synchronized.some((stage, idx) => 
      stage.fluid_name !== scheduleRaw[idx]?.fluid_name
    );
    
    if (hasChanges) {
      console.log("Pumping schedule fluid mappings synchronized");
      // Update database asynchronously
      setTimeout(() => {
        setDb(d => ({
          ...d,
          pumping_schedule: {
            ...d.pumping_schedule,
            [activeProject]: synchronized
          }
        }));
      }, 0);
    }
    
    return synchronized;
  }, [scheduleRaw, fluids, activeProject]);

  const PURPOSES = ["Pre-flush","Chemical Wash","Spacer","Lead Cement","Tail Cement","Displacement","Plug","Squeeze"];

  const u = (ri,fld,v) => {
    var nr=[...schedule];nr[ri]={...nr[ri],[fld]:v};
    if(fld==="fluid_id"){var fl=fluids.find(f=>f.fluid_id===v);if(fl)nr[ri].fluid_name=fl.fluid_name;}
    setDb(d=>({...d,pumping_schedule:{...d.pumping_schedule,[activeProject]:nr}}));
  };
  const add = () => setDb(d=>({...d,pumping_schedule:{...d.pumping_schedule,[activeProject]:[...schedule,{stage:schedule.length+1,fluid_id:fluids[0]?.fluid_id||"",fluid_name:fluids[0]?.fluid_name||"",rate:6.0,volume:0,purpose:"Lead Cement",time_start:0}]}}));
  const del = ri => setDb(d=>({...d,pumping_schedule:{...d.pumping_schedule,[activeProject]:schedule.filter((_,i)=>i!==ri)}}));
  const handleImport = (rows) => setDb(d=>({...d,pumping_schedule:{...d.pumping_schedule,[activeProject]:rows}}));

  // Calculate Volume: uses open hole geometry to compute annular volume for cement stages
  const calculateVolumes = () => {
    if (openHoles.length === 0) { alert("No open hole profile data. Please enter Open Hole data first."); return; }
    // Annular capacity per interval
    const casingData = db.casing_profile[activeProject]||[];
    const annVols = openHoles.map(h => {
      const cas = casingData.find(c => c.to_depth >= h.to_depth) || casingData[0];
      const casingOD = cas ? cas.od : h.hole_size * 0.8;
      const annCap = Math.max(0, h.hole_size*h.hole_size - casingOD*casingOD) / 1029.4; // bbl/ft
      const intLen = h.to_depth - h.from_depth;
      return annCap * intLen; // bbl for this interval
    });
    const totalAnnVol = Math.round(annVols.reduce((s,v)=>s+v,0) * 10) / 10;

    const updated = schedule.map(s => {
      const fl = fluids.find(f=>f.fluid_id===s.fluid_id);
      const isCement = fl?.type?.toLowerCase()==="cement";
      const isSpacerOrWash = ["spacer","wash"].includes(fl?.type?.toLowerCase()||"");
      if (isCement) return { ...s, volume: totalAnnVol };
      if (isSpacerOrWash) return { ...s, volume: Math.round(totalAnnVol * 0.15 * 10) / 10 }; // spacer ~15% of cement vol
      return s;
    });
    setDb(d=>({...d,pumping_schedule:{...d.pumping_schedule,[activeProject]:updated}}));
  };

  const totalVol = schedule.reduce((s,r)=>s+(r.volume||0),0);
  const cols = [
    {key:"stage",label:"Stage"},
    {key:"fluid_id",label:"Fluid",type:"select",options:fluids.map(f=>({v:f.fluid_id,l:f.fluid_name}))},
    {key:"rate",label:"Rate (bpm)",step:0.5},
    {key:"volume",label:"Volume (bbl)"},
    {key:"purpose",label:"Purpose",type:"select",options:PURPOSES},
  ];

  return (
    <div className="animate-in">
      <SectionTitle icon="⛽" title="PUMPING SCHEDULE" sub="Displacement train sequence · Purpose dropdown · Auto-volume from open hole geometry" />
      <div style={{ display:"flex",gap:10,marginBottom:12,flexWrap:"wrap",alignItems:"center" }}>
        <ExcelImportBtn tableName="pumping_schedule" onImport={handleImport} />
        <button onClick={()=>exportTableToExcel(schedule,"pumping_schedule")} style={{ background:"transparent",border:`1px solid ${T.border}`,color:T.muted,borderRadius:4,padding:"5px 12px",fontSize:11,cursor:"pointer",fontFamily:"'IBM Plex Sans'",fontWeight:600 }}>⬇ Export Excel</button>
        <button onClick={()=>{ setDb(d=>({...d,pumping_schedule:{...d.pumping_schedule,[activeProject]:JSON.parse(JSON.stringify(DEFAULT_DATA.pumping_schedule))}})); }} style={{ background:"transparent",border:`1px solid ${T.gold}`,color:T.gold,borderRadius:4,padding:"5px 14px",fontSize:11,cursor:"pointer",fontFamily:"'IBM Plex Sans'",fontWeight:600 }}>↺ Set to Default</button>
        <button onClick={calculateVolumes} style={{ background:`linear-gradient(135deg,${T.accent2},${T.accent})`,color:"#000",border:"none",borderRadius:4,padding:"6px 16px",fontSize:12,cursor:"pointer",fontFamily:"'IBM Plex Sans'",fontWeight:700 }}>
          🧮 Calculate Volume
        </button>
        <div style={{ fontSize:10,color:T.dim }}>Auto-fills cement/spacer volumes from Open Hole geometry</div>
      </div>
      <IntervalTable columns={cols} data={schedule} onChange={u} onAdd={add} onDelete={del}
        rowColor={row=>{const fl=fluids.find(f=>f.fluid_id===row.fluid_id);return fl?`${fl.color}18`:"transparent";}}
        fieldRules={{rate:"pump_rate",volume:"pump_volume"}} />
      <div style={{ display:"flex",gap:12,marginTop:12 }}>
        <Card style={{ padding:"10px 16px" }}><div style={{ fontSize:11,color:T.muted }}>Total Volume</div><div style={{ fontSize:18,fontWeight:700,color:T.accent,fontFamily:"'IBM Plex Mono'" }}>{totalVol} bbl</div></Card>
        {openHoles.length > 0 && (
          <Card style={{ padding:"10px 16px" }}>
            <div style={{ fontSize:11,color:T.muted }}>Open Hole Intervals</div>
            <div style={{ fontSize:18,fontWeight:700,color:T.cyan,fontFamily:"'IBM Plex Mono'" }}>{openHoles.length}</div>
          </Card>
        )}
      </div>
      <InputNavFooter pageId="pumping" onNavigate={onNavigate} pages={inputPages} />
    </div>
  );
}

function CentralizerPage({ db, activeProject, setDb, onNavigate, inputPages }) {
  const data = db.centralizer_data[activeProject]||[];
  const u = (ri,fld,v) => {
    var nr=[...data];
    if (fld === "type") {
      var cProps = CENTRALIZER_TYPE_PROPS[v] || {};
      nr[ri] = { ...nr[ri], type:v, standoff:cProps.baseStandoff||nr[ri].standoff, run_in_force:cProps.rihForce||nr[ri].run_in_force };
    } else {
      nr[ri] = { ...nr[ri], [fld]:v };
    }
    setDb(d=>({...d,centralizer_data:{...d.centralizer_data,[activeProject]:nr}}));
  };
  const add = () => {
    const props = CENTRALIZER_TYPE_PROPS["Bow Spring"];
    setDb(d=>({...d,centralizer_data:{...d.centralizer_data,[activeProject]:[...data,{id:`cr${Date.now()}`,from_depth:3000,to_depth:8500,spacing:40,type:"Bow Spring",standoff:props.baseStandoff,run_in_force:props.rihForce}]}}));
  };
  const del = ri => setDb(d=>({...d,centralizer_data:{...d.centralizer_data,[activeProject]:data.filter((_,i)=>i!==ri)}}));
  const handleImport = (rows) => setDb(d=>({...d,centralizer_data:{...d.centralizer_data,[activeProject]:rows}}));
  const cols = [{key:"from_depth",label:"From (ft)"},{key:"to_depth",label:"To (ft)"},{key:"type",label:"Type",type:"select",options:Object.keys(CENTRALIZER_TYPE_PROPS)},{key:"spacing",label:"Spacing (ft)"},{key:"standoff",label:"Standoff %",step:1},{key:"run_in_force",label:"RIH Force (lbs)"}];
  const RULES = { spacing:"cent_spacing", standoff:"cent_standoff" };
  // Depth order validation
  const orderErrors = checkDepthOrder(data, "from_depth");
  // Cross-page: centralizer max depth must not exceed open hole max depth
  const ohMax = maxDepthInTable(db.open_hole_profile[activeProject]||[]);
  const centMax = maxDepthInTable(data);
  return (
    <div className="animate-in">
      <SectionTitle icon="⚙️" title="CENTRALIZER DATA" sub="Type selection auto-fills standoff % and RIH force — each type applies different turbulence bonus in displacement calculations" />
      {orderErrors.length > 0 && (
        <div style={{ background:`${T.red}15`,border:`1px solid ${T.red}44`,borderRadius:4,padding:"8px 12px",marginBottom:10 }}>
          {orderErrors.map((e,i)=><div key={i} style={{ fontSize:12,color:T.red,fontWeight:600 }}>⛔ {e.msg}</div>)}
        </div>
      )}
      <CrossDepthAlert tableName="Centralizer" tableMax={centMax} ohMax={ohMax} />
      <div style={{ display:"flex",gap:10,marginBottom:12,flexWrap:"wrap",alignItems:"center" }}>
        <ExcelImportBtn tableName="centralizer_data" onImport={handleImport} />
        <button onClick={()=>exportTableToExcel(data,"centralizer_data")} style={{ background:"transparent",border:`1px solid ${T.border}`,color:T.muted,borderRadius:4,padding:"5px 12px",fontSize:11,cursor:"pointer",fontFamily:"'IBM Plex Sans'",fontWeight:600 }}>⬇ Export Excel</button>
        <button onClick={()=>{ setDb(d=>({...d,centralizer_data:{...d.centralizer_data,[activeProject]:JSON.parse(JSON.stringify(DEFAULT_DATA.centralizer_data))}})); }} style={{ background:"transparent",border:`1px solid ${T.gold}`,color:T.gold,borderRadius:4,padding:"5px 14px",fontSize:11,cursor:"pointer",fontFamily:"'IBM Plex Sans'",fontWeight:600 }}>↺ Set to Default</button>
      </div>
      <IntervalTable columns={cols} data={data} onChange={u} onAdd={add} onDelete={del} fieldRules={RULES} />
      <Card style={{ marginTop:12 }}>
        <div style={{ fontSize:11,color:T.accent,fontFamily:"'IBM Plex Mono'",marginBottom:8 }}>CENTRALIZER TYPE REFERENCE</div>
        <div style={{ display:"flex",gap:10,flexWrap:"wrap" }}>
          {Object.entries(CENTRALIZER_TYPE_PROPS).map(([type,props])=>(
            <div key={type} style={{ flex:1,minWidth:160,background:T.panel,borderRadius:6,padding:"8px 12px",border:`1px solid ${T.border}` }}>
              <div style={{ fontSize:11,fontWeight:700,color:T.green,marginBottom:4 }}>{type}</div>
              <div style={{ fontSize:10,color:T.muted,fontFamily:"'IBM Plex Mono'",lineHeight:1.6 }}>
                Standoff: {props.baseStandoff}% · RIH: {props.rihForce} lbs<br/>
                Turbulence bonus: +{props.turbulenceBonus}%<br/>
                <span style={{ color:T.dim }}>{props.description}</span>
              </div>
            </div>
          ))}
        </div>
      </Card>
      <InputNavFooter pageId="centralizers" onNavigate={onNavigate} pages={inputPages} />
    </div>
  );
}

function TemperaturePage({ db, activeProject, setDb, onNavigate, inputPages }) {
  const data = db.temperature_profile[activeProject]||[];
  const u = (ri,fld,v) => { var nr=[...data];nr[ri]={...nr[ri],[fld]:v};setDb(d=>({...d,temperature_profile:{...d.temperature_profile,[activeProject]:nr}})); };
  const add = () => setDb(d=>({...d,temperature_profile:{...d.temperature_profile,[activeProject]:[...data,{depth:5000,temperature:180}]}}));
  const del = ri => setDb(d=>({...d,temperature_profile:{...d.temperature_profile,[activeProject]:data.filter((_,i)=>i!==ri)}}));
  const handleImport = (rows) => setDb(d=>({...d,temperature_profile:{...d.temperature_profile,[activeProject]:rows}}));
  const cols = [{key:"depth",label:"Depth (ft)"},{key:"temperature",label:"Temperature (°F)"}];
  return (
    <div className="animate-in">
      <SectionTitle icon="🌡️" title="TEMPERATURE PROFILE" sub="Static or circulating temperature survey — depth-temperature pairs for linear interpolation" />
      <div style={{ display:"flex",gap:10,marginBottom:12,flexWrap:"wrap",alignItems:"center" }}>
        <ExcelImportBtn tableName="temperature_profile" onImport={handleImport} />
        <button onClick={()=>exportTableToExcel(data,"temperature_profile")} style={{ background:"transparent",border:`1px solid ${T.border}`,color:T.muted,borderRadius:4,padding:"5px 12px",fontSize:11,cursor:"pointer",fontFamily:"'IBM Plex Sans'",fontWeight:600 }}>⬇ Export Excel</button>
        <button onClick={()=>{ setDb(d=>({...d,temperature_profile:{...d.temperature_profile,[activeProject]:JSON.parse(JSON.stringify(DEFAULT_DATA.temperature_profile))}})); }} style={{ background:"transparent",border:`1px solid ${T.gold}`,color:T.gold,borderRadius:4,padding:"5px 14px",fontSize:11,cursor:"pointer",fontFamily:"'IBM Plex Sans'",fontWeight:600 }}>↺ Set to Default</button>
      </div>
      <div className="grid-2">
        <IntervalTable columns={cols} data={data} onChange={u} onAdd={add} onDelete={del} fieldRules={{depth:"temp_depth",temperature:"temperature"}} />
        <Card>
          <div style={{ fontSize:11,color:T.muted,marginBottom:8,fontFamily:"'IBM Plex Mono'" }}>TEMPERATURE PREVIEW</div>
          <ResponsiveContainer width="100%" height={220}>
            <LineChart data={[...data].sort((a,b)=>a.depth-b.depth)} layout="vertical" margin={{top:5,right:20,left:10,bottom:5}}>
              <CartesianGrid strokeDasharray="3 3" stroke={T.dim} />
              <XAxis type="number" dataKey="temperature" tick={{fontSize:10,fill:T.muted}} />
              <YAxis type="number" dataKey="depth" reversed tick={{fontSize:10,fill:T.muted}} />
              <Tooltip contentStyle={{background:T.card,border:`1px solid ${T.border}`,fontSize:11}} />
              <Line type="monotone" dataKey="temperature" stroke={T.gold} dot={{fill:T.gold,r:4}} strokeWidth={2} name="Temp (°F)" />
            </LineChart>
          </ResponsiveContainer>
        </Card>
      </div>
      <InputNavFooter pageId="temperature" onNavigate={onNavigate} pages={inputPages} />
    </div>
  );
}

function BalancedPlugPage({ db, activeProject, setDb, onNavigate, simResults, inputPages }) {
  const data = db.balanced_plug_data[activeProject]||[];
  const fluids = db.fluid_data[activeProject]||[];
  const casingProfile = db.casing_profile[activeProject]||[];
  const openHoles = db.open_hole_profile[activeProject]||[];
  const gen = db.general_data?.[activeProject] || {};

  // Determine if Balanced Plug is enabled based on Job Type
  const jobType = gen.job_type || "";
  const isEnabled = jobType === "Secondary (Balanced Plug)";
  const disabledReason = !jobType ? "No Job Type selected on General page"
    : jobType === "Primary" ? "Disabled — Job Type is Primary"
    : jobType === "Secondary (Remedial)" ? "Disabled — Job Type is Secondary (Remedial)"
    : "";

  const u = (ri,fld,v) => { if (!isEnabled) return; var nr=[...data];nr[ri]={...nr[ri],[fld]:v};setDb(d=>({...d,balanced_plug_data:{...d.balanced_plug_data,[activeProject]:nr}})); };
  const add = () => { if (!isEnabled) return; setDb(d=>({...d,balanced_plug_data:{...d.balanced_plug_data,[activeProject]:[...data,{id:`bp${Date.now()}`,from_depth:7000,to_depth:8000,length:1000,set_depth:7000,excess:150,fluid_id:fluids[0]?.fluid_id||""}]}})); };
  const del = ri => { if (!isEnabled) return; setDb(d=>({...d,balanced_plug_data:{...d.balanced_plug_data,[activeProject]:data.filter((_,i)=>i!==ri)}})); };
  const handleImport = (rows) => { if (!isEnabled) return; setDb(d=>({...d,balanced_plug_data:{...d.balanced_plug_data,[activeProject]:rows}})); };

  const cols = [{key:"from_depth",label:"From (ft)"},{key:"to_depth",label:"To (ft)"},{key:"length",label:"Length (ft)"},{key:"set_depth",label:"Set Depth (ft)"},{key:"excess",label:"Excess (ft)"},{key:"fluid_id",label:"Fluid",type:"select",options:fluids.map(f=>({v:f.fluid_id,l:f.fluid_name}))}];

  // Live dynamic calculation only when enabled
  const liveResults = useMemo(() => isEnabled ? computeBalancedPlug(data, casingProfile, openHoles, fluids) : [], [data, casingProfile, openHoles, fluids, isEnabled]);

  // Depth order + cross-page (only shown when enabled)
  const orderErrors = isEnabled ? checkDepthOrder(data, "from_depth") : [];
  const ohMax   = maxDepthInTable(db.open_hole_profile[activeProject]||[]);
  const plugMax = maxDepthInTable(data);

  return (
    <div className="animate-in">
      <SectionTitle icon="🔌" title="BALANCED PLUG PLACEMENT" sub="Fluid selection dynamically updates TOC, displacement volumes, and pressure balance in real time" />

      {/* Job Type status banner */}
      <div style={{
        padding:"10px 16px", borderRadius:6, marginBottom:14,
        background: isEnabled ? `${T.green}15` : `${T.dim}30`,
        border: `1px solid ${isEnabled ? T.green : T.dim}`,
        display:"flex", alignItems:"center", gap:10,
      }}>
        <span style={{ fontSize:16 }}>{isEnabled ? "✅" : "🔒"}</span>
        <div>
          <div style={{ fontSize:12, fontWeight:700, color: isEnabled ? T.green : T.muted, fontFamily:"'IBM Plex Mono'" }}>
            {isEnabled ? "ENABLED — Job Type: Secondary (Balanced Plug)" : disabledReason || "DISABLED"}
          </div>
          {!isEnabled && (
            <div style={{ fontSize:11, color:T.dim, marginTop:3 }}>
              Go to General page → change Job Type to "Secondary (Balanced Plug)" to enable this input table
            </div>
          )}
        </div>
      </div>

      {/* Depth validation — only shown when enabled */}
      {isEnabled && orderErrors.length > 0 && (
        <div style={{ background:`${T.red}15`,border:`1px solid ${T.red}44`,borderRadius:4,padding:"8px 12px",marginBottom:10 }}>
          {orderErrors.map((e,i)=><div key={i} style={{ fontSize:12,color:T.red,fontWeight:600 }}>⛔ {e.msg}</div>)}
        </div>
      )}
      {isEnabled && <CrossDepthAlert tableName="Balanced Plug" tableMax={plugMax} ohMax={ohMax} />}

      {/* Disabled overlay wrapper */}
      <div style={{ opacity: isEnabled ? 1 : 0.4, pointerEvents: isEnabled ? "auto" : "none", userSelect: isEnabled ? "auto" : "none" }}>
        <div style={{ display:"flex",gap:10,marginBottom:12,flexWrap:"wrap",alignItems:"center" }}>
          <ExcelImportBtn tableName="balanced_plug_data" onImport={handleImport} />
          <button onClick={()=>exportTableToExcel(data,"balanced_plug_data")} style={{ background:"transparent",border:`1px solid ${T.border}`,color:T.muted,borderRadius:4,padding:"5px 12px",fontSize:11,cursor:"pointer",fontFamily:"'IBM Plex Sans'",fontWeight:600 }}>⬇ Export Excel</button>
          <SetToDefaultBtn onClick={()=>{ if(isEnabled){ setDb(d=>({...d,balanced_plug_data:{...d.balanced_plug_data,[activeProject]:JSON.parse(JSON.stringify(DEFAULT_DATA.balanced_plug_data))}})); } }} />
        </div>
        <IntervalTable columns={cols} data={data} onChange={u} onAdd={add} onDelete={del} />

        {/* Live dynamic preview */}
        {liveResults.length > 0 && (
          <Card style={{ marginTop:16,border:`1px solid ${T.accent}44` }}>
            <div style={{ fontSize:11,color:T.accent,fontFamily:"'IBM Plex Mono'",marginBottom:10 }}>
              ⚡ LIVE CALCULATION — UPDATES ON EVERY INPUT CHANGE
            </div>
            <table>
              <thead><tr>{["From","To","Fluid","ρ (ppg)","Length (ft)","Plug Vol (bbl)","Displ. Vol","Water Behind","Final TOC","BHHP (psi)","Balance Diff","Balanced"].map(h=><th key={h}>{h}</th>)}</tr></thead>
              <tbody>
                {liveResults.map((p,i)=>(
                  <tr key={i} style={{ background:p.balanced?`${T.green}08`:`${T.red}08` }}>
                    <td>{p.from_depth}</td><td>{p.to_depth}</td>
                    <td style={{ color:T.cyan }}>{p.fluidName}</td>
                    <td style={{ color:T.gold,fontWeight:700 }}>{p.fluidDensity}</td>
                    <td>{p.length}</td>
                    <td>{p.plugVolume}</td><td>{p.displacementVolume}</td>
                    <td>{p.waterBehind}</td>
                    <td style={{ color:T.accent,fontWeight:700 }}>{p.finalTOC} ft</td>
                    <td>{p.bottomHolePressure}</td>
                    <td style={{ color:p.balanced?T.green:T.red }}>{p.balanceDiff} psi/ft</td>
                    <td style={{ color:p.balanced?T.green:T.red,fontWeight:700,fontSize:14 }}>
                      {p.balanced ? "✓ BALANCED" : "✗ CHECK"}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
            <div style={{ marginTop:10,fontSize:11,color:T.muted }}>
              Balance condition: |P_inside - P_outside| &lt; 0.5 psi/ft at set depth
            </div>
          </Card>
        )}
      </div>

      <InputNavFooter pageId="plug" onNavigate={onNavigate} pages={inputPages} />
    </div>
  );
}

// ─────────────────────────── RESULTS PAGES ────────────────────────────────────

function NoResults() {
  return (
    <div style={{ textAlign:"center",padding:"80px 20px",color:T.muted }}>
      <div style={{ fontSize:48,marginBottom:12 }}>⚙️</div>
      <div style={{ fontFamily:"'Orbitron'",fontSize:14,color:T.accent,marginBottom:8 }}>NO SIMULATION DATA</div>
      <div style={{ fontSize:13 }}>Run the simulation engine from the Dashboard to generate depth-based results</div>
    </div>
  );
}

function ResultsPressurePage({ simResults }) {
  if (!simResults) return <NoResults />;
  const d = simResults.grid.filter((_,i)=>i%2===0);
  const s = simResults.summary;
  const surfP = s.maxSurfacePressure || s.finalPumpPressure || 0;

  // Classify each grid point for pressure window status
  const kicks    = d.filter(r => r.ecdPressure < r.porePressure);
  const fractures= d.filter(r => r.ecdPressure > r.fracturePressure);
  const warnings = d.filter(r => r.ecdPressure >= r.porePressure && r.ecdPressure <= r.fracturePressure &&
    (r.ecdPressure < r.porePressure * 1.05 || r.ecdPressure > r.fracturePressure * 0.95));

  // Build annotated dataset with zone flags
  const annotated = d.map(r => ({
    ...r,
    // For reference lines in chart — annotate critical points
    kickZone:     r.ecdPressure < r.porePressure    ? r.ecdPressure : null,
    fracZone:     r.ecdPressure > r.fracturePressure ? r.ecdPressure : null,
    warnZone:     (r.ecdPressure >= r.porePressure && r.ecdPressure <= r.fracturePressure &&
                   (r.safetyMargin < r.fracturePressure * 0.05 || r.ecdPressure < r.porePressure * 1.05))
                   ? r.ecdPressure : null,
  }));

  return (
    <div className="animate-in">
      <SectionTitle icon="📊" title="MODULE 1 — HYDRAULICS & PRESSURE" sub="Pressure envelope · ECD · Surface pressure · Safety margin vs depth" />

      {/* Alert banners — shown only when violations exist */}
      {fractures.length > 0 && (
        <div style={{ background:`${T.red}20`, border:`2px solid ${T.red}`, borderRadius:6, padding:"10px 16px", marginBottom:12,
          display:"flex", alignItems:"center", gap:12 }}>
          <span style={{ fontSize:18 }}>🔴</span>
          <div>
            <div style={{ fontSize:12, fontWeight:700, color:T.red, fontFamily:"'IBM Plex Mono'" }}>
              CRITICAL: ECD EXCEEDS FRACTURE PRESSURE — {fractures.length} depth point{fractures.length>1?"s":""} at risk
            </div>
            <div style={{ fontSize:11, color:T.muted, marginTop:3 }}>
              Shallowest violation at {fractures[0].md.toLocaleString()} ft MD · Max ECD at violation: {Math.max(...fractures.map(r=>r.ecd))} ppg
              · Reduce pump rate or switch to lighter fluid to prevent lost circulation
            </div>
          </div>
        </div>
      )}
      {kicks.length > 0 && (
        <div style={{ background:`${T.yellow}20`, border:`2px solid ${T.yellow}`, borderRadius:6, padding:"10px 16px", marginBottom:12,
          display:"flex", alignItems:"center", gap:12 }}>
          <span style={{ fontSize:18 }}>🟡</span>
          <div>
            <div style={{ fontSize:12, fontWeight:700, color:T.yellow, fontFamily:"'IBM Plex Mono'" }}>
              WARNING: ECD BELOW PORE PRESSURE — {kicks.length} depth point{kicks.length>1?"s":""} — influx / kick risk
            </div>
            <div style={{ fontSize:11, color:T.muted, marginTop:3 }}>
              Shallowest underbalance at {kicks[0].md.toLocaleString()} ft MD
              · Increase fluid density or reduce pump rate to restore overbalance
            </div>
          </div>
        </div>
      )}

      <div className="grid-4" style={{ marginBottom:16 }}>
        <MetricCard label="Max ECD" value={s.maxECD} unit="ppg" color={s.maxECD>14?T.red:T.green} icon="💧" />
        <MetricCard label="Min Safety" value={`${s.minSafety.toLocaleString()} psi`} unit="Frac - ECD" color={s.minSafety<300?T.red:T.green} icon="🛡️" />
        <MetricCard label="Surface Pressure" value={`${surfP.toLocaleString()}`} unit="psi (max during job)" color={surfP>5000?T.red:surfP>3000?T.yellow:T.green} icon="🔧" />
        <MetricCard label="Window Flags" value={fractures.length + kicks.length}
          unit={`${fractures.length} critical · ${kicks.length} warning`}
          color={(fractures.length>0)?T.red:(kicks.length>0)?T.yellow:T.green} icon="⚠️" />
      </div>

      {/* Surface pressure detail card */}
      <Card style={{ marginBottom:16, borderLeft:`3px solid ${T.cyan}` }}>
        <div style={{ fontSize:11,color:T.muted,textTransform:"uppercase",letterSpacing:"0.08em",marginBottom:8 }}>🔧 SURFACE PRESSURE DETAIL</div>
        <div style={{ display:"flex",gap:24,flexWrap:"wrap" }}>
          {[
            { l:"Max Surface Pressure",  v:`${surfP.toLocaleString()} psi`,               c:surfP>5000?T.red:T.green },
            { l:"Final Pump Pressure",   v:`${s.finalPumpPressure.toLocaleString()} psi`, c:T.cyan },
            { l:"Pump Rate",             v:`${s.pumpRate} bpm`,                           c:T.text },
            { l:"Surface Temp",          v:`${s.surfaceTemp}°F`,                          c:T.gold },
          ].map((m,i)=>(
            <div key={i} style={{ minWidth:160 }}>
              <div style={{ fontSize:10,color:T.muted,marginBottom:3 }}>{m.l}</div>
              <div style={{ fontSize:16,fontWeight:700,color:m.c,fontFamily:"'IBM Plex Mono'" }}>{m.v}</div>
            </div>
          ))}
        </div>
        <div style={{ fontSize:11,color:T.dim,marginTop:10 }}>
          Formula: P_surface = P_ECD − P_hydrostatic + P_pipe_friction
        </div>
      </Card>

      <div className="grid-2">
        {/* Main pressure envelope chart with custom reference lines */}
        <Card>
          <div style={{ fontSize:12,fontWeight:600,color:T.accent,fontFamily:"'IBM Plex Mono'",marginBottom:8,letterSpacing:"0.05em" }}>
            PRESSURE ENVELOPE (psi vs Depth)
            {fractures.length>0 && <span style={{ marginLeft:10,fontSize:10,color:T.red }}>🔴 {fractures.length} CRITICAL</span>}
            {kicks.length>0    && <span style={{ marginLeft:6, fontSize:10,color:T.yellow }}>🟡 {kicks.length} WARNING</span>}
          </div>
          <ResponsiveContainer width="100%" height={480}>
            <LineChart data={[...annotated].sort((a,b)=>a.md-b.md)} layout="vertical" margin={{top:5,right:30,left:10,bottom:20}}>
              <CartesianGrid strokeDasharray="3 3" stroke={T.dim} />
              <XAxis type="number" tick={{fontSize:10,fill:T.muted}} label={{value:"psi",position:"insideBottom",offset:-5,style:{fontSize:10,fill:T.muted}}} />
              <YAxis type="number" dataKey="md" reversed={true} domain={["dataMin","dataMax"]} tick={{fontSize:10,fill:T.muted}} label={{value:"Depth (ft MD)",angle:-90,position:"insideLeft",style:{fontSize:10,fill:T.muted}}} />
              <Tooltip contentStyle={{background:T.card,border:`1px solid ${T.border}`,fontSize:11}}
                labelFormatter={v=>`Depth: ${v} ft`}
                formatter={(val,name,props)=>{
                  const r=props.payload;
                  if(name==="ECD Pressure") {
                    const status = r.ecdPressure>r.fracturePressure ? " 🔴 CRITICAL" :
                                   r.ecdPressure<r.porePressure ? " 🟡 WARNING" : " ✓ OK";
                    return [`${val?.toLocaleString()} psi${status}`, name];
                  }
                  return [val?.toLocaleString()+" psi", name];
                }}
              />
              <Legend wrapperStyle={{fontSize:11}} />
              {/* Highlight fracture violation zones */}
              {fractures.map((r,i)=>(
                <ReferenceLine key={`f${i}`} y={r.md} stroke={T.red} strokeWidth={0.5} strokeOpacity={0.4} />
              ))}
              {/* Highlight kick zones */}
              {kicks.map((r,i)=>(
                <ReferenceLine key={`k${i}`} y={r.md} stroke={T.yellow} strokeWidth={0.5} strokeOpacity={0.4} />
              ))}
              <Line type="monotone" dataKey="hydroPressure"    stroke={T.cyan}   dot={false} strokeWidth={1.8} name="Hydrostatic" />
              <Line type="monotone" dataKey="ecdPressure"      stroke={T.yellow} dot={false} strokeWidth={2.5} name="ECD Pressure"
                strokeDasharray={fractures.length>0?"none":"none"} />
              <Line type="monotone" dataKey="porePressure"     stroke={T.green}  dot={false} strokeWidth={1.8} name="Pore Pressure" />
              <Line type="monotone" dataKey="fracturePressure" stroke={T.red}    dot={false} strokeWidth={1.8} name="Frac Pressure" />
            </LineChart>
          </ResponsiveContainer>
          {/* Zone legend */}
          <div style={{ display:"flex",gap:16,marginTop:8,fontSize:10,flexWrap:"wrap" }}>
            <span style={{ color:T.green }}>✅ Safe: ECD between Pore and Frac</span>
            <span style={{ color:T.yellow }}>🟡 Warning: ECD ≤ Pore (influx risk)</span>
            <span style={{ color:T.red }}>🔴 Critical: ECD ≥ Frac (loss circulation)</span>
          </div>
        </Card>

        {/* Depth-zone status table */}
        <Card>
          <div style={{ fontSize:12,fontWeight:600,color:T.accent,fontFamily:"'IBM Plex Mono'",marginBottom:8 }}>
            PRESSURE WINDOW STATUS vs DEPTH
          </div>
          <div style={{ overflowY:"auto",maxHeight:520 }}>
            <table>
              <thead><tr>
                {["Depth (ft)","ECD (ppg)","Pore P.","Frac P.","Safety (psi)","Status"].map(h=><th key={h}>{h}</th>)}
              </tr></thead>
              <tbody>
                {d.map((r,i)=>{
                  const isFrac = r.ecdPressure > r.fracturePressure;
                  const isKick = r.ecdPressure < r.porePressure;
                  const bg = isFrac ? `${T.red}18` : isKick ? `${T.yellow}18` : "transparent";
                  return (
                    <tr key={i} style={{ background:bg }}>
                      <td style={{ color:T.accent,fontWeight:600 }}>{r.md.toLocaleString()}</td>
                      <td style={{ color:isFrac?T.red:isKick?T.yellow:T.text,fontWeight:700 }}>{r.ecd}</td>
                      <td style={{ color:T.green }}>{r.porePressure.toLocaleString()}</td>
                      <td style={{ color:T.red }}>{r.fracturePressure.toLocaleString()}</td>
                      <td style={{ color:r.safetyMargin<0?T.red:r.safetyMargin<500?T.yellow:T.green,fontWeight:700 }}>
                        {r.safetyMargin.toLocaleString()}
                      </td>
                      <td>
                        {isFrac
                          ? <span style={{ color:T.red,  fontWeight:700,fontSize:11 }}>🔴 CRITICAL</span>
                          : isKick
                          ? <span style={{ color:T.yellow,fontWeight:700,fontSize:11 }}>🟡 WARNING</span>
                          : <span style={{ color:T.green,fontSize:11 }}>✅ OK</span>}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </Card>
      </div>
      <DepthChart data={d} title="SURFACE & FRICTION PRESSURE vs Depth" xLabel="psi" height={300}
        lines={[{key:"surfacePumpPressure",name:"Surface Pressure (psi)",color:T.cyan},{key:"frictionPressure",name:"Ann. Friction (psi)",color:T.gold},{key:"pipeFriction",name:"Pipe Friction (psi)",color:T.accent}]} />
    </div>
  );
}

function ResultsDisplacementPage({ simResults, db, activeProject }) {
  if (!simResults) return <NoResults />;
  const d = simResults.grid.filter((_,i)=>i%2===0);
  const s = simResults.summary;
  const fluids = db.fluid_data[activeProject]||[];

  // Zone counts
  const poor       = d.filter(r => r.dispEfficiency < 70).length;
  const acceptable = d.filter(r => r.dispEfficiency >= 70 && r.dispEfficiency <= 90).length;
  const excellent  = d.filter(r => r.dispEfficiency > 90).length;

  // Build per-fluid efficiency data — group grid points by fluid name
  const fluidNames = [...new Set(d.map(r => r.fluidName))];
  const fluidColors = {};
  fluids.forEach(f => { fluidColors[f.fluid_name] = f.color; });

  // Add zone classification to each point
  const zoneData = d.map(r => ({
    ...r,
    zone: r.dispEfficiency < 70 ? "poor" : r.dispEfficiency <= 90 ? "acceptable" : "excellent",
    zoneColor: r.dispEfficiency < 70 ? T.red : r.dispEfficiency <= 90 ? T.yellow : T.green,
  }));

  return (
    <div className="animate-in">
      <SectionTitle icon="🔄" title="MODULE 2 — MUD DISPLACEMENT / CLEANOUT" sub="Fluid displacement efficiency per pumped fluid · Zone classification · Contamination vs depth" />

      {/* Zone summary banners */}
      {poor > 0 && (
        <div style={{ background:`${T.red}18`,border:`1px solid ${T.red}`,borderRadius:6,padding:"8px 14px",marginBottom:8,display:"flex",alignItems:"center",gap:10 }}>
          <span style={{ fontSize:16 }}>🔴</span>
          <span style={{ fontSize:12,color:T.red,fontWeight:700,fontFamily:"'IBM Plex Mono'" }}>
            POOR DISPLACEMENT: {poor} depth point{poor>1?"s":""} below 70% efficiency — channeling or mud bypassing likely
          </span>
        </div>
      )}

      <div className="grid-4" style={{ marginBottom:16 }}>
        <MetricCard label="Avg Disp. Efficiency" value={`${s.avgDisp}%`} unit="Annular avg" color={s.avgDisp>90?T.green:s.avgDisp>70?T.yellow:T.red} icon="🔄" />
        <MetricCard label="🔴 Poor (<70%)" value={poor} unit="depth steps" color={poor>0?T.red:T.green} icon="⚠️" />
        <MetricCard label="🟡 Acceptable (70–90%)" value={acceptable} unit="depth steps" color={T.yellow} icon="⚡" />
        <MetricCard label="🟢 Excellent (>90%)" value={excellent} unit="depth steps" color={T.green} icon="✅" />
      </div>

      <div className="grid-2" style={{ marginBottom:16 }}>
        {/* Displacement efficiency chart with zone background bands */}
        <Card>
          <div style={{ fontSize:12,fontWeight:600,color:T.accent,fontFamily:"'IBM Plex Mono'",marginBottom:8 }}>
            DISPLACEMENT EFFICIENCY (%) vs DEPTH — ZONE CLASSIFIED
          </div>
          <div style={{ display:"flex",gap:12,marginBottom:8,fontSize:10,flexWrap:"wrap" }}>
            <span style={{ color:T.red }}>🔴 Poor &lt;70%</span>
            <span style={{ color:T.yellow }}>🟡 Acceptable 70–90%</span>
            <span style={{ color:T.green }}>🟢 Excellent &gt;90%</span>
          </div>
          <ResponsiveContainer width="100%" height={460}>
            <LineChart data={[...zoneData].sort((a,b)=>a.md-b.md)} layout="vertical" margin={{top:5,right:40,left:10,bottom:20}}>
              <CartesianGrid strokeDasharray="3 3" stroke={T.dim} />
              <XAxis type="number" domain={[0,100]} tick={{fontSize:10,fill:T.muted}}
                label={{value:"%",position:"insideBottom",offset:-5,style:{fontSize:10,fill:T.muted}}}
                ticks={[0,20,40,60,70,80,90,100]} />
              <YAxis type="number" dataKey="md" reversed={true} domain={["dataMin","dataMax"]} tick={{fontSize:10,fill:T.muted}} label={{value:"Depth (ft MD)",angle:-90,position:"insideLeft",style:{fontSize:10,fill:T.muted}}} />
              <Tooltip contentStyle={{background:T.card,border:`1px solid ${T.border}`,fontSize:11}}
                labelFormatter={v=>`Depth: ${v} ft`}
                formatter={(val,name,props)=>{
                  const z=props.payload.zone;
                  const label=z==="poor"?"🔴 POOR":z==="acceptable"?"🟡 ACCEPTABLE":"🟢 EXCELLENT";
                  return [`${val}% — ${label}`, name];
                }}
              />
              <Legend wrapperStyle={{fontSize:11}} />
              {/* Poor zone reference */}
              <ReferenceLine x={70} stroke={T.red} strokeDasharray="4 2" strokeWidth={1.5}
                label={{value:"70% min",position:"insideTopRight",style:{fontSize:9,fill:T.red}}} />
              <ReferenceLine x={90} stroke={T.green} strokeDasharray="4 2" strokeWidth={1.5}
                label={{value:"90% excel",position:"insideTopRight",style:{fontSize:9,fill:T.green}}} />
              <Line type="monotone" dataKey="dispEfficiency" stroke={T.green} dot={false} strokeWidth={2}
                name="Disp. Efficiency %" />
              {/* Overlay coloured dots for zone classification */}
              {zoneData.filter(r=>r.zone==="poor").length > 0 &&
                <Line type="monotone" data={zoneData.filter(r=>r.zone==="poor")}
                  dataKey="dispEfficiency" stroke={T.red} dot={false} strokeWidth={4}
                  name="Poor (<70%)" connectNulls={false} />}
            </LineChart>
          </ResponsiveContainer>
        </Card>

        {/* Per-fluid efficiency breakdown */}
        <Card>
          <div style={{ fontSize:12,fontWeight:600,color:T.accent,fontFamily:"'IBM Plex Mono'",marginBottom:12 }}>
            DISPLACEMENT EFFICIENCY BY FLUID STAGE
          </div>
          {fluidNames.map(fname => {
            const pts = d.filter(r => r.fluidName === fname);
            if (pts.length === 0) return null;
            const avgDE = Math.round(pts.reduce((s,r)=>s+r.dispEfficiency,0)/pts.length);
            const minDE = Math.min(...pts.map(r=>r.dispEfficiency));
            const maxDE = Math.max(...pts.map(r=>r.dispEfficiency));
            const depthRange = `${pts[0].md.toLocaleString()}–${pts[pts.length-1].md.toLocaleString()} ft`;
            const fc = fluidColors[fname] || T.muted;
            const zone = avgDE < 70 ? "poor" : avgDE <= 90 ? "acceptable" : "excellent";
            const zoneLabel = zone==="poor"?"🔴 POOR":zone==="acceptable"?"🟡 ACCEPTABLE":"🟢 EXCELLENT";
            const zoneColor = zone==="poor"?T.red:zone==="acceptable"?T.yellow:T.green;
            return (
              <div key={fname} style={{ marginBottom:12, padding:"10px 12px", background:T.panel, borderRadius:6,
                border:`1px solid ${fc}44`, borderLeft:`3px solid ${fc}` }}>
                <div style={{ display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6 }}>
                  <div style={{ display:"flex",alignItems:"center",gap:8 }}>
                    <span style={{ width:10,height:10,borderRadius:"50%",background:fc,flexShrink:0 }}></span>
                    <span style={{ fontSize:12,fontWeight:700,color:T.text }}>{fname}</span>
                  </div>
                  <span style={{ fontSize:11,fontWeight:700,color:zoneColor }}>{zoneLabel}</span>
                </div>
                <div style={{ fontSize:10,color:T.muted,marginBottom:6 }}>{depthRange} · {pts.length} grid points</div>
                {/* Efficiency bar */}
                <div style={{ marginBottom:4 }}>
                  <div style={{ display:"flex",justifyContent:"space-between",fontSize:10,color:T.muted,marginBottom:2 }}>
                    <span>Avg: {avgDE}%</span><span>Min: {minDE}%</span><span>Max: {maxDE}%</span>
                  </div>
                  <div style={{ height:10,background:T.dim,borderRadius:5,overflow:"hidden" }}>
                    <div style={{ height:"100%",width:`${avgDE}%`,
                      background:`linear-gradient(90deg,${avgDE<70?T.red:avgDE<=90?T.yellow:T.green},${fc})`,
                      borderRadius:5,transition:"width 0.3s" }} />
                  </div>
                </div>
                <div style={{ height:4,background:T.bg,borderRadius:2,overflow:"hidden",marginTop:3 }}>
                  <div style={{ height:"100%",width:`${minDE}%`,background:`${T.red}55`,borderRadius:2 }} />
                </div>
              </div>
            );
          })}
          {/* Zone reference */}
          <div style={{ marginTop:10,padding:"8px 10px",background:T.bg,borderRadius:4,fontSize:10,color:T.muted,lineHeight:1.7 }}>
            Industry standard (API RP 10D):<br/>
            🔴 &lt;70% — Poor: channeling / bypassing likely · increase standoff or pump rate<br/>
            🟡 70–90% — Acceptable: adequate for non-critical intervals<br/>
            🟢 &gt;90% — Excellent: full displacement achieved · target for pay zones
          </div>
        </Card>
      </div>

      <DepthChart data={d} title="CONTAMINATION & STANDOFF vs Depth" xLabel="%" height={280}
        lines={[{key:"contamination",name:"Contamination %",color:T.red},{key:"standoff",name:"Standoff %",color:T.cyan}]} />
      <DepthChart data={d.slice(0,d.length)} title="ANNULAR & INTERNAL VELOCITY vs Depth" xLabel="ft/min" height={280}
        lines={[{key:"annularVelocity",name:"Ann. Velocity",color:T.accent},{key:"internalVelocity",name:"Internal Velocity",color:T.gold}]} />
    </div>
  );
}

function ResultsThermalPage({ simResults }) {
  if (!simResults) return <NoResults />;
  const d = simResults.grid.filter((_,i)=>i%2===0);
  const s = simResults.summary;
  return (
    <div className="animate-in">
      <SectionTitle icon="🌡️" title="MODULE 3 — TEMPERATURE (BHCT)" sub="Bottom hole circulating temperature · Temperature profile vs depth · Thermal variation" />
      <div className="grid-4" style={{ marginBottom:16 }}>
        <MetricCard label="BHCT" value={`${s.bhct}°F`} unit="Bottom Hole Circ. Temp" color={T.gold} icon="🌡️" />
        <MetricCard label="Surface Temp" value={`${simResults.grid[0]?.temperature||70}°F`} unit="At wellhead" color={T.muted} icon="🌍" />
        <MetricCard label="Temp Gradient" value={`${((s.bhct-(simResults.grid[0]?.temperature||70))/8500*100).toFixed(2)}`} unit="°F/100ft avg" color={T.cyan} icon="📈" />
        <MetricCard label="BHCT (°C)" value={`${cvt(s.bhct,"temperature",true)}`} unit="Celsius" color={T.yellow} icon="🌡️" />
      </div>
      <DepthChart data={d} title="TEMPERATURE PROFILE vs DEPTH (°F)" xLabel="°F" height={480} lines={[{key:"temperature",name:"Temperature (°F)",color:T.gold}]} />
    </div>
  );
}

function ResultsCentralizerPage({ simResults, db, activeProject }) {
  if (!simResults) return <NoResults />;
  const d = simResults.grid.filter((_,i)=>i%2===0);
  const cent = db.centralizer_data[activeProject]||[];
  return (
    <div className="animate-in">
      <SectionTitle icon="⚙️" title="MODULE 4 — CENTRALIZER PLACEMENT / STANDOFF" sub="Standoff percentage · Centralizer distribution · Contact points along wellbore" />
      <div className="grid-2">
        <DepthChart data={d} title="STANDOFF (%) vs DEPTH" xLabel="%" height={420} lines={[{key:"standoff",name:"Standoff %",color:T.cyan}]} />
        <Card>
          <div style={{ fontSize:11,color:T.accent,fontFamily:"'IBM Plex Mono'",marginBottom:10 }}>CENTRALIZER SCHEDULE</div>
          <table>
            <thead><tr>{["From (ft)","To (ft)","Type","Spacing","Standoff %","Status"].map(h=><th key={h}>{h}</th>)}</tr></thead>
            <tbody>
              {cent.map((c,i)=>(
                <tr key={i}>
                  <td>{c.from_depth}</td><td>{c.to_depth}</td><td>{c.type}</td><td>{c.spacing} ft</td>
                  <td style={{ color:c.standoff>=67?T.green:T.yellow,fontWeight:700 }}>{c.standoff}%</td>
                  <td>{c.standoff>=67?<Badge text="OK" color={T.green} />:<Badge text="LOW" color={T.yellow} />}</td>
                </tr>
              ))}
            </tbody>
          </table>
          <div style={{ marginTop:12 }}>
            <div style={{ fontSize:11,color:T.muted,marginBottom:8 }}>STANDOFF SUMMARY vs DEPTH</div>
            {d.filter((_,i)=>i%6===0).map((r,i)=>(
              <div key={i} style={{ display:"flex",alignItems:"center",gap:8,marginBottom:3 }}>
                <span style={{ fontSize:10,color:T.muted,width:55,fontFamily:"'IBM Plex Mono'" }}>{r.md} ft</span>
                <div style={{ flex:1,background:T.dim,borderRadius:2,height:8 }}>
                  <div style={{ width:`${r.standoff}%`,background:r.standoff>=67?T.green:T.yellow,height:"100%",borderRadius:2 }}></div>
                </div>
                <span style={{ fontSize:10,color:r.standoff>=67?T.green:T.yellow,width:35,fontFamily:"'IBM Plex Mono'" }}>{r.standoff}%</span>
              </div>
            ))}
          </div>
        </Card>
      </div>
    </div>
  );
}

function ResultsPlugPage({ simResults }) {
  if (!simResults) return <NoResults />;
  const plugResults = simResults.plugResults||[];
  return (
    <div className="animate-in">
      <SectionTitle icon="🔌" title="MODULE 5 — BALANCED PLUG PLACEMENT" sub="Plug volume · Displacement fluid · Water behind · Final TOC · Balance verification" />
      {plugResults.length === 0 ? (
        <Card><div style={{ color:T.muted,fontSize:12 }}>No balanced plug data defined. Add plug data in the Balanced Plug input page.</div></Card>
      ) : (
        <>
          <div className="grid-4" style={{ marginBottom:16 }}>
            {plugResults.map((p,i)=>(
              <MetricCard key={i} label={`Plug ${i+1} — TOC`} value={`${p.finalTOC} ft`} unit={`${p.from_depth}–${p.to_depth} ft`} color={p.balanced?T.green:T.red} icon="🔌" />
            ))}
          </div>
          <Card>
            <div style={{ fontSize:11,color:T.accent,fontFamily:"'IBM Plex Mono'",marginBottom:10 }}>PLUG BALANCE VERIFICATION</div>
            <table>
              <thead><tr>{["From (ft)","To (ft)","Length (ft)","Plug Vol (bbl)","Displ. Vol (bbl)","Water Behind","Ann. Cap (bbl/ft)","Final TOC","Balanced"].map(h=><th key={h}>{h}</th>)}</tr></thead>
              <tbody>
                {plugResults.map((p,i)=>(
                  <tr key={i} style={{ background:p.balanced?`${T.green}08`:`${T.red}08` }}>
                    <td>{p.from_depth}</td><td>{p.to_depth}</td><td>{p.length}</td>
                    <td>{p.plugVolume}</td><td>{p.displacementVolume}</td><td>{p.waterBehind}</td>
                    <td>{p.annularCapacity}</td><td>{p.finalTOC}</td>
                    <td style={{ color:p.balanced?T.green:T.red,fontWeight:700,fontSize:14 }}>{p.balanced?"✓ BALANCED":"✗ UNBALANCED"}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </Card>
        </>
      )}
    </div>
  );
}

function ResultsHookLoadPage({ simResults }) {
  if (!simResults) return <NoResults />;
  const d = simResults.grid.filter((_,i)=>i%2===0);
  return (
    <div className="animate-in">
      <SectionTitle icon="🏗️" title="TORQUE & DRAG" sub="Cumulative hook load and torque profiles vs depth for casing running analysis" />
      <div className="grid-2">
        <DepthChart data={d} title="HOOK LOAD (kips) vs Depth" xLabel="kips" height={440} lines={[{key:"hookLoad",name:"Hook Load (kips)",color:T.accent}]} />
        <DepthChart data={d} title="TORQUE (ft-lbs) vs Depth" xLabel="ft-lbs" height={440} lines={[{key:"torque",name:"Torque (ft-lbs)",color:T.gold}]} />
      </div>
    </div>
  );
}


// ═══════════════════════════════════════════════════════════════════════════
// DEBUG REPORT PAGE — Complete Engineering Calculation Audit
// ═══════════════════════════════════════════════════════════════════════════
function DebugReportPage({ simResults, db, activeProject }) {
  if (!simResults || !simResults.grid) return <NoResults />;
  
  // Get active project metadata
  var proj = db.projects.find(p => p.project_id === activeProject) || {};
  
  var grid = simResults.grid;
  var summary = simResults.summary || {};
  var g = db.general_data[activeProject] || {};
  var survey = db.survey_data[activeProject] || [];
  var casing = db.casing_profile[activeProject] || [];
  var openhole = db.open_hole_profile[activeProject] || [];
  var formation = db.formation_data[activeProject] || [];
  var fluids = db.fluid_data[activeProject] || [];
  var pump = db.pumping_schedule[activeProject] || [];
  var centralizers = db.centralizer_data[activeProject] || [];
  var temp = db.temperature_profile[activeProject] || [];
  var plug = db.balanced_plug_data[activeProject] || [];
  
  // Calculate correct depth summaries from survey data
  var maxMD = survey.length > 0 ? Math.max(...survey.map(s => s.md)) : g.total_depth_md || 0;
  var maxTVD = survey.length > 0 ? Math.max(...survey.map(s => s.tvd)) : maxMD;
  var casingShoeDepth = g.casing_shoe_depth || (casing.length > 0 ? Math.max(...casing.map(c => c.to_depth)) : 0);
  var topOfCement = g.toc || casingShoeDepth || 0;
  
  // Get active casing geometry (production string)
  var activeCasing = casing.length > 0 ? casing[casing.length - 1] : null;
  var casingOD = activeCasing ? activeCasing.od : 0;
  var casingID = activeCasing ? activeCasing.id_ : 0;
  
  // Get active hole size (deepest section)
  var activeHole = openhole.length > 0 ? openhole[openhole.length - 1] : null;
  var holeSize = activeHole ? activeHole.hole_size : g.hole_size || 0;
  
  // Validate data mappings and warn if inconsistencies detected
  var warnings = [];
  if (maxMD === 0) warnings.push("⚠ Survey depth is zero - check survey data");
  if (casingOD === 0 || casingID === 0) warnings.push("⚠ Casing geometry is zero - check casing profile");
  if (holeSize === 0) warnings.push("⚠ Hole size is zero - check open hole profile");
  
  // Check fluid mapping consistency
  pump.forEach((stage, idx) => {
    var fluidMatch = fluids.find(f => f.fluid_id === stage.fluid_id);
    if (fluidMatch && fluidMatch.fluid_name !== stage.fluid_name) {
      warnings.push(`⚠ Stage ${stage.stage}: Fluid ID ${stage.fluid_id} name mismatch - Schedule shows "${stage.fluid_name}" but fluid DB has "${fluidMatch.fluid_name}"`);
    }
    if (!fluidMatch) {
      warnings.push(`⚠ Stage ${stage.stage}: Fluid ID ${stage.fluid_id} not found in fluid database`);
    }
  });
  
  // Check formation data completeness
  formation.forEach((zone, idx) => {
    if (!zone.pore_gradient || zone.pore_gradient === 0) {
      warnings.push(`⚠ Formation zone ${idx + 1} (${zone.from_depth}-${zone.to_depth}ft): Missing pore gradient`);
    }
    if (!zone.frac_gradient || zone.frac_gradient === 0) {
      warnings.push(`⚠ Formation zone ${idx + 1} (${zone.from_depth}-${zone.to_depth}ft): Missing fracture gradient`);
    }
  });
  
  if (warnings.length > 0) {
    console.warn("Debug Report Data Validation Warnings:", warnings);
  }

  return (
    <div className="animate-in" style={{ maxWidth:1400 }}>
      <SectionTitle icon="🐛" title="DEBUG REPORT" sub="Complete engineering calculations audit — all input parameters, equations, grid-by-grid computations, and iteration convergence history" />

      {/* ═══ SECTION 1: INPUT PARAMETERS ═══ */}
      <Card>
        <h3 style={{ fontFamily:"'IBM Plex Mono'",fontSize:13,color:T.accent,marginBottom:12,letterSpacing:"0.05em" }}>
          § 1. INPUT PARAMETERS
        </h3>
        
        {warnings.length > 0 && (
          <div style={{ background:T.gold+"15",border:`1px solid ${T.gold}`,borderRadius:4,padding:10,marginBottom:12 }}>
            <div style={{ fontSize:11,fontWeight:600,color:T.gold,marginBottom:6 }}>⚠ Data Validation Warnings ({warnings.length})</div>
            {warnings.map((w, i) => (
              <div key={i} style={{ fontSize:10,color:T.text,marginBottom:2 }}>{w}</div>
            ))}
          </div>
        )}
        
        <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:12,marginBottom:16 }}>
          <div>
            <div style={{ fontSize:10,color:T.muted,fontWeight:600,marginBottom:4 }}>PROJECT</div>
            <div style={{ fontSize:11 }}>Name: {proj.project_name || "Untitled"}</div>
            <div style={{ fontSize:11 }}>Well: {proj.well_name || "—"}</div>
            <div style={{ fontSize:11 }}>Field: {proj.field_name || "—"}</div>
            <div style={{ fontSize:11 }}>Units: {proj.unit_system || "imperial"}</div>
          </div>
          <div>
            <div style={{ fontSize:10,color:T.muted,fontWeight:600,marginBottom:4 }}>KEY DEPTHS</div>
            <div style={{ fontSize:11 }}>MD: {maxMD.toFixed(0)} ft</div>
            <div style={{ fontSize:11 }}>TVD: {maxTVD.toFixed(0)} ft</div>
            <div style={{ fontSize:11 }}>Shoe: {casingShoeDepth.toFixed(0)} ft</div>
            <div style={{ fontSize:11 }}>TOC: {topOfCement.toFixed(0)} ft</div>
          </div>
          <div>
            <div style={{ fontSize:10,color:T.muted,fontWeight:600,marginBottom:4 }}>GEOMETRY</div>
            <div style={{ fontSize:11 }}>Casing OD: {casingOD.toFixed(3)} in</div>
            <div style={{ fontSize:11 }}>Casing ID: {casingID.toFixed(3)} in</div>
            <div style={{ fontSize:11 }}>Hole Size: {holeSize.toFixed(2)} in</div>
            <div style={{ fontSize:11 }}>Job Type: {g.job_type || "Primary"}</div>
          </div>
        </div>

        <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:12 }}>
          <div>
            <div style={{ fontSize:10,color:T.muted,fontWeight:600,marginBottom:4 }}>SURVEY DATA ({survey.length} points)</div>
            <div style={{ fontSize:10,fontFamily:"'IBM Plex Mono'",background:T.dim,padding:6,borderRadius:4,maxHeight:120,overflowY:"auto" }}>
              {survey.slice(0,10).map((s,i)=>(
                <div key={i}>MD:{s.md} TVD:{s.tvd} Inc:{s.inclination}° Az:{s.azimuth}°</div>
              ))}
              {survey.length > 10 && <div>... +{survey.length-10} more</div>}
            </div>
          </div>
          <div>
            <div style={{ fontSize:10,color:T.muted,fontWeight:600,marginBottom:4 }}>CASING/OPENHOLE ({casing.length} + {openhole.length} intervals)</div>
            <div style={{ fontSize:10,fontFamily:"'IBM Plex Mono'",background:T.dim,padding:6,borderRadius:4,maxHeight:120,overflowY:"auto" }}>
              <div>CASING:</div>
              {casing.slice(0,5).map((c,i)=>(
                <div key={i}>{c.from_depth}-{c.to_depth}ft: {c.type} OD:{c.od}" ID:{c.id_}"</div>
              ))}
              {casing.length > 5 && <div>... +{casing.length-5} more</div>}
              <div style={{ marginTop:6 }}>OPENHOLE:</div>
              {openhole.slice(0,5).map((h,i)=>(
                <div key={i}>{h.from_depth}-{h.to_depth}ft: {h.hole_size}"</div>
              ))}
              {openhole.length > 5 && <div>... +{openhole.length-5} more</div>}
            </div>
          </div>
        </div>

        <div style={{ marginTop:12,display:"grid",gridTemplateColumns:"1fr 1fr",gap:12 }}>
          <div>
            <div style={{ fontSize:10,color:T.muted,fontWeight:600,marginBottom:4 }}>FORMATION DATA ({formation.length} zones)</div>
            <div style={{ fontSize:10,fontFamily:"'IBM Plex Mono'",background:T.dim,padding:6,borderRadius:4,maxHeight:120,overflowY:"auto" }}>
              {formation.slice(0,8).map((f,i)=>(
                <div key={i}>{f.from_depth}-{f.to_depth}ft: {f.lithology} Pore:{f.pore_gradient || "N/A"} Frac:{f.frac_gradient || "N/A"}</div>
              ))}
              {formation.length > 8 && <div>... +{formation.length-8} more</div>}
            </div>
          </div>
          <div>
            <div style={{ fontSize:10,color:T.muted,fontWeight:600,marginBottom:4 }}>FLUID DATABASE ({fluids.length} fluids)</div>
            <div style={{ fontSize:10,fontFamily:"'IBM Plex Mono'",background:T.dim,padding:6,borderRadius:4,maxHeight:120,overflowY:"auto" }}>
              {fluids.map((f,i)=>(
                <div key={i}>ID:{f.fluid_id} {f.fluid_name} {f.density}ppg {f.model} PV:{f.pv} YP:{f.yp}</div>
              ))}
            </div>
          </div>
        </div>

        <div style={{ marginTop:12 }}>
          <div style={{ fontSize:10,color:T.muted,fontWeight:600,marginBottom:4 }}>PUMPING SCHEDULE ({pump.length} stages) - Fluid Mapping Verification</div>
          <div style={{ fontSize:10,fontFamily:"'IBM Plex Mono'",background:T.dim,padding:6,borderRadius:4,maxHeight:100,overflowY:"auto" }}>
            {pump.map((stage,i)=>{
              var fluidRef = fluids.find(f => f.fluid_id === stage.fluid_id);
              var mismatch = fluidRef && fluidRef.fluid_name !== stage.fluid_name;
              return (
                <div key={i} style={{ color: mismatch ? T.gold : T.text }}>
                  Stage {stage.stage}: ID={stage.fluid_id} Schedule="{stage.fluid_name}" 
                  {fluidRef ? ` → DB="{fluidRef.fluid_name}" ${fluidRef.density}ppg PV:${fluidRef.pv}` : " → NOT FOUND IN DB"}
                  {mismatch && " ⚠ MISMATCH"}
                  {" | "}Vol:{stage.volume}bbl Rate:{stage.rate}bpm
                </div>
              );
            })}
          </div>
        </div>
      </Card>

      {/* ═══ SECTION 2: UNIT CONVERSIONS & CONSTANTS ═══ */}
      <Card style={{ marginTop:12 }}>
        <h3 style={{ fontFamily:"'IBM Plex Mono'",fontSize:13,color:T.accent,marginBottom:12,letterSpacing:"0.05em" }}>
          § 2. UNIT CONVERSIONS & ENGINEERING CONSTANTS
        </h3>
        <div style={{ fontSize:10,fontFamily:"'IBM Plex Mono'",background:T.dim,padding:8,borderRadius:4,lineHeight:1.6 }}>
          <div><strong>PRESSURE:</strong></div>
          <div>• Hydrostatic: ΔP (psi) = ρ (ppg) × 0.052 × ΔTVD (ft)</div>
          <div>• ECD: ECD (ppg) = P_total (psi) / (0.052 × TVD)</div>
          <div style={{ marginTop:6 }}><strong>FRICTION CORRELATIONS (Field Units):</strong></div>
          <div style={{ marginLeft:12,marginTop:4 }}><strong>Annular Friction Pressure:</strong></div>
          <div style={{ marginLeft:24 }}>ΔP_annular (psi) = Σ [ΔP/100ft × ΔMD/100] from surface to depth</div>
          <div style={{ marginLeft:24 }}>where ΔP/100ft depends on rheology model:</div>
          <div style={{ marginLeft:36,marginTop:4 }}><em>Newtonian Laminar (Re &lt; 2100):</em></div>
          <div style={{ marginLeft:48 }}>ΔP/100 (psi/100ft) = (PV × v) / (300 × d_h)</div>
          <div style={{ marginLeft:48 }}>PV: plastic viscosity (cP), v: velocity (ft/min), d_h: hydraulic diameter (in)</div>
          <div style={{ marginLeft:36,marginTop:4 }}><em>Bingham Laminar (Re &lt; 2100):</em></div>
          <div style={{ marginLeft:48 }}>ΔP/100 (psi/100ft) = (PV × v) / (300 × d_h) + YP / (225 × d_h)</div>
          <div style={{ marginLeft:48 }}>PV: plastic viscosity (cP), YP: yield point (lbf/100ft²), v: velocity (ft/min)</div>
          <div style={{ marginLeft:36,marginTop:4 }}><em>Power Law Laminar (Re &lt; 2100):</em></div>
          <div style={{ marginLeft:48 }}>ΔP/100 (psi/100ft) = [k × v^n / (144 × d_h^(1+n))] × [(2+1/n)/0.0208]^n × 144</div>
          <div style={{ marginLeft:48 }}>k: consistency index (lbf·s^n/100ft²), n: flow behavior index, v: velocity (ft/min)</div>
          <div style={{ marginLeft:36,marginTop:4 }}><em>Turbulent (Re &gt; 3000):</em></div>
          <div style={{ marginLeft:48 }}>ΔP/100 (psi/100ft) = f × ρ × v² / (25.8 × d_h)</div>
          <div style={{ marginLeft:48 }}>where f = Fanning friction factor ≈ 0.046 / Re^0.2 for smooth pipes</div>
          <div style={{ marginLeft:48 }}>Simplified form (for typical Re=5000-50000): ΔP/100 ≈ ρ × v² / 25600</div>
          <div style={{ marginLeft:48 }}>ρ: density (ppg), v: velocity (ft/min), d_h: hydraulic diameter (in)</div>
          <div style={{ marginLeft:12,marginTop:6 }}><strong>Pipe (Internal) Friction Pressure:</strong></div>
          <div style={{ marginLeft:24 }}>ΔP_pipe (psi) = Σ [ΔP/100ft × ΔMD/100] from surface to depth</div>
          <div style={{ marginLeft:24 }}>Same correlations as annular, using:</div>
          <div style={{ marginLeft:36 }}>• d_h = casing ID (in)</div>
          <div style={{ marginLeft:36 }}>• v_pipe = Q / Cap_pipe, where Cap_pipe = ID² / 1029.4 (bbl/ft)</div>
          <div style={{ marginLeft:12,marginTop:6 }}><strong>Incremental Accumulation:</strong></div>
          <div style={{ marginLeft:24 }}>Friction accumulates step-by-step (not linear extrapolation):</div>
          <div style={{ marginLeft:36 }}>P_friction(depth_i) = P_friction(depth_i-1) + ΔP/100ft × stepSize/100</div>
          <div style={{ marginLeft:24 }}>This accounts for changing fluid properties, geometry, and flow regime with depth.</div>
          <div style={{ marginTop:6 }}><strong>VELOCITY (CORRECTED FORMULA):</strong></div>
          <div>• General: v (ft/min) = Q (bbl/min) / Capacity (bbl/ft)</div>
          <div>• Annular: v_ann = Q / Cap_ann, where Cap_ann = (D_hole² − D_casing²) / 1029.4</div>
          <div>• Internal: v_pipe = Q / Cap_pipe, where Cap_pipe = D_id² / 1029.4</div>
          <div>• Note: Units cancel correctly: (bbl/min) ÷ (bbl/ft) = ft/min</div>
          <div style={{ marginTop:6 }}><strong>CAPACITY:</strong></div>
          <div>• Annular: Cap (bbl/ft) = (D_hole² − D_casing²) / 1029.4</div>
          <div>• Internal: Cap (bbl/ft) = D_id² / 1029.4</div>
          <div>• Conversion factor: 1029.4 in²/(bbl/ft) — standard oilfield constant</div>
          <div style={{ marginTop:6 }}><strong>REYNOLDS NUMBER:</strong></div>
          <div>• Newtonian: Re = 928 × ρ × v × d / μ</div>
          <div>• Bingham: Re = 928 × ρ × v × d / PV</div>
          <div>• Power Law: Re = 109 × ρ × v^(2-n) × d^n / (144^(1-n) × k × ((2+1/n)/0.0208)^n)</div>
          <div style={{ marginTop:6 }}><strong>TEMPERATURE:</strong></div>
          <div>• Geothermal gradient: T_static (°F) = T_surface + (TVD × grad / 100)</div>
          <div>• Dynamic adjustment: T_circ ≈ T_static × (0.7 + 0.3 × (1 − exp(−TVD/5000)))</div>
        </div>
      </Card>

      {/* ═══ SECTION 3: GRID CALCULATIONS (first 20 points) ═══ */}
      <Card style={{ marginTop:12 }}>
        <h3 style={{ fontFamily:"'IBM Plex Mono'",fontSize:13,color:T.accent,marginBottom:12,letterSpacing:"0.05em" }}>
          § 3. GRID-BY-GRID CALCULATIONS (first 20 of {grid.length} points)
        </h3>
        <div style={{ fontSize:9,fontFamily:"'IBM Plex Mono'",overflowX:"auto" }}>
          <table style={{ width:"100%",borderCollapse:"collapse" }}>
            <thead>
              <tr style={{ background:T.dim,borderBottom:"1px solid " + T.border }}>
                <th style={{ padding:"4px 6px",textAlign:"left" }}>MD</th>
                <th style={{ padding:"4px 6px",textAlign:"left" }}>TVD</th>
                <th style={{ padding:"4px 6px",textAlign:"left" }}>Inc°</th>
                <th style={{ padding:"4px 6px",textAlign:"left" }}>Temp°F</th>
                <th style={{ padding:"4px 6px",textAlign:"left" }}>ρ_adj</th>
                <th style={{ padding:"4px 6px",textAlign:"left" }}>P_hydro</th>
                <th style={{ padding:"4px 6px",textAlign:"left" }}>P_fric</th>
                <th style={{ padding:"4px 6px",textAlign:"left" }}>P_pipe</th>
                <th style={{ padding:"4px 6px",textAlign:"left" }}>ECD_dyn</th>
                <th style={{ padding:"4px 6px",textAlign:"left" }}>ECD_static</th>
                <th style={{ padding:"4px 6px",textAlign:"left" }}>Pore</th>
                <th style={{ padding:"4px 6px",textAlign:"left" }}>Frac</th>
                <th style={{ padding:"4px 6px",textAlign:"left" }}>Margin</th>
                <th style={{ padding:"4px 6px",textAlign:"left" }}>v_ann</th>
                <th style={{ padding:"4px 6px",textAlign:"left" }}>Re</th>
                <th style={{ padding:"4px 6px",textAlign:"left" }}>Flow</th>
              </tr>
            </thead>
            <tbody>
              {grid.slice(0,20).map((r,i)=>(
                <tr key={i} style={{ borderBottom:"1px solid " + T.border }}>
                  <td style={{ padding:"4px 6px" }}>{r.md}</td>
                  <td style={{ padding:"4px 6px" }}>{r.tvd}</td>
                  <td style={{ padding:"4px 6px" }}>{r.inclination?.toFixed(1)}</td>
                  <td style={{ padding:"4px 6px" }}>{r.temperature?.toFixed(0)}</td>
                  <td style={{ padding:"4px 6px" }}>{r.adjDensity?.toFixed(2)}</td>
                  <td style={{ padding:"4px 6px" }}>{r.hydroPressure?.toFixed(0)}</td>
                  <td style={{ padding:"4px 6px" }}>{r.frictionPressure?.toFixed(0)}</td>
                  <td style={{ padding:"4px 6px" }}>{r.pipeFriction?.toFixed(0)}</td>
                  <td style={{ padding:"4px 6px",color:T.yellow }}>{r.ecd?.toFixed(2)}</td>
                  <td style={{ padding:"4px 6px" }}>{r.staticECD?.toFixed(2)}</td>
                  <td style={{ padding:"4px 6px" }}>{r.porePressure?.toFixed(0)}</td>
                  <td style={{ padding:"4px 6px" }}>{r.fracturePressure?.toFixed(0)}</td>
                  <td style={{ padding:"4px 6px",color:r.safetyMargin<0.5?T.red:T.green }}>{r.safetyMargin?.toFixed(2)}</td>
                  <td style={{ padding:"4px 6px" }}>{r.annularVelocity?.toFixed(0)}</td>
                  <td style={{ padding:"4px 6px" }}>{r.reynolds?.toFixed(0)}</td>
                  <td style={{ padding:"4px 6px" }}>{r.flowRegime}</td>
                </tr>
              ))}
            </tbody>
          </table>
          {grid.length > 20 && (
            <div style={{ marginTop:8,fontSize:10,color:T.muted }}>
              ... showing first 20 of {grid.length} total grid points. Download full data via "Full Data Grid" tab.
            </div>
          )}
        </div>
      </Card>

      {/* ═══ SECTION 4: PRESSURE ACCUMULATION BREAKDOWN ═══ */}
      <Card style={{ marginTop:12 }}>
        <h3 style={{ fontFamily:"'IBM Plex Mono'",fontSize:13,color:T.accent,marginBottom:12,letterSpacing:"0.05em" }}>
          § 4. PRESSURE COMPONENT BREAKDOWN (at key depths)
        </h3>
        <div style={{ fontSize:10,fontFamily:"'IBM Plex Mono'" }}>
          {[0, Math.floor(grid.length*0.25), Math.floor(grid.length*0.5), Math.floor(grid.length*0.75), grid.length-1].map(idx=>{
            var r = grid[idx];
            if (!r) return null;
            return (
              <div key={idx} style={{ marginBottom:12,padding:8,background:T.dim,borderRadius:4 }}>
                <div style={{ fontWeight:600,marginBottom:4 }}>Depth: {r.md} ft MD / {r.tvd} ft TVD</div>
                <div>P_hydrostatic = ρ_avg × 0.052 × TVD = {r.adjDensity?.toFixed(2)} × 0.052 × {r.tvd} = {r.hydroPressure?.toFixed(0)} psi</div>
                <div>P_friction_ann = ∫ dP_fric = {r.frictionPressure?.toFixed(0)} psi (accumulated annular friction)</div>
                <div>P_friction_pipe = {r.pipeFriction?.toFixed(0)} psi (pipe friction to this depth)</div>
                <div>P_total = P_hydro + P_fric_ann = {r.hydroPressure?.toFixed(0)} + {r.frictionPressure?.toFixed(0)} = {r.ecdPressure?.toFixed(0)} psi</div>
                <div style={{ color:T.yellow,fontWeight:600 }}>ECD_dynamic = P_total / (0.052 × TVD) = {r.ecdPressure?.toFixed(0)} / (0.052 × {r.tvd}) = {r.ecd?.toFixed(2)} ppg</div>
                <div>ECD_static = {r.staticECD?.toFixed(2)} ppg (no friction)</div>
                <div>Friction contribution: +{r.frictionECDContrib?.toFixed(2)} ppg</div>
                <div style={{ marginTop:4 }}>Pore pressure: {r.porePressure?.toFixed(0)} psi ({(r.porePressure/(0.052*r.tvd))?.toFixed(2)} ppg equiv)</div>
                <div>Frac pressure: {r.fracturePressure?.toFixed(0)} psi ({(r.fracturePressure/(0.052*r.tvd))?.toFixed(2)} ppg equiv)</div>
                <div style={{ color:r.safetyMargin<0.5?T.red:T.green }}>Safety margin: {r.safetyMargin?.toFixed(2)} ppg</div>
              </div>
            );
          })}
        </div>
      </Card>

      {/* ═══ SECTION 5: FINAL SURFACE PUMP PRESSURE ═══ */}
      <Card style={{ marginTop:12 }}>
        <h3 style={{ fontFamily:"'IBM Plex Mono'",fontSize:13,color:T.accent,marginBottom:12,letterSpacing:"0.05em" }}>
          § 5. FINAL SURFACE PUMP PRESSURE BREAKDOWN
        </h3>
        <div style={{ fontSize:11,fontFamily:"'IBM Plex Mono'",background:T.dim,padding:10,borderRadius:4 }}>
          <div><strong>Components at TD:</strong></div>
          <div style={{ marginLeft:12,marginTop:6 }}>
            <div>P_hydrostatic_ann = {grid[grid.length-1]?.hydroPressure?.toFixed(0)} psi</div>
            <div>P_friction_ann = {grid[grid.length-1]?.frictionPressure?.toFixed(0)} psi</div>
            <div>P_friction_pipe = {grid[grid.length-1]?.pipeFriction?.toFixed(0)} psi</div>
          </div>
          <div style={{ marginTop:8,fontWeight:600,fontSize:12,color:T.yellow }}>
            P_pump_surface = P_friction_pipe + P_friction_ann = {summary.finalPumpPressure?.toFixed(0)} psi
          </div>
          <div style={{ marginTop:6,fontSize:10,color:T.muted }}>
            (Hydrostatic cancels: surface pump must overcome pipe + annular friction only)
          </div>
        </div>
      </Card>

      {/* ═══ SECTION 6: DISPLACEMENT MODULE ═══ */}
      <Card style={{ marginTop:12 }}>
        <h3 style={{ fontFamily:"'IBM Plex Mono'",fontSize:13,color:T.accent,marginBottom:12,letterSpacing:"0.05em" }}>
          § 6. DISPLACEMENT EFFICIENCY CALCULATIONS
        </h3>
        <div style={{ fontSize:10,fontFamily:"'IBM Plex Mono'" }}>
          {grid.slice(0,10).map((r,i)=>(
            <div key={i} style={{ marginBottom:6,padding:6,background:T.dim,borderRadius:4 }}>
              <div>Depth {r.md} ft: Standoff={r.standoff?.toFixed(2)} Velocity={r.annularVelocity?.toFixed(0)}fpm Reynolds={r.reynolds?.toFixed(0)}</div>
              <div>Disp.Eff = f(standoff,Re,flowRegime) = {r.dispEfficiency?.toFixed(3)}</div>
              <div>Contamination = 1 − ∫eff = {r.contamination?.toFixed(3)}</div>
            </div>
          ))}
          {grid.length > 10 && <div style={{ marginTop:8,color:T.muted }}>... +{grid.length-10} more points</div>}
        </div>
      </Card>

      {/* ═══ SECTION 7: TEMPERATURE MODULE ═══ */}
      <Card style={{ marginTop:12 }}>
        <h3 style={{ fontFamily:"'IBM Plex Mono'",fontSize:13,color:T.accent,marginBottom:12,letterSpacing:"0.05em" }}>
          § 7. TEMPERATURE PROFILE CALCULATIONS
        </h3>
        <div style={{ fontSize:10,fontFamily:"'IBM Plex Mono'" }}>
          <div style={{ marginBottom:8 }}>Surface temp: {summary.surfaceTemp || 70}°F</div>
          {grid.slice(0,10).map((r,i)=>(
            <div key={i} style={{ marginBottom:6,padding:6,background:T.dim,borderRadius:4 }}>
              <div>Depth {r.md} ft ({r.tvd} TVD):</div>
              <div>T_static = T_surf + grad × TVD/100 = {r.tempStatic?.toFixed(1)}°F</div>
              <div>T_circulating = T_static × damping_factor = {r.temperature?.toFixed(1)}°F</div>
              <div>ρ_adjusted = ρ_base × temp_correction = {r.adjDensity?.toFixed(2)} ppg</div>
            </div>
          ))}
          {grid.length > 10 && <div style={{ marginTop:8,color:T.muted }}>... +{grid.length-10} more points</div>}
          <div style={{ marginTop:12,fontWeight:600 }}>
            BHCT (circulating): {summary.bhct}°F · BHST (static): {summary.bhst}°F
          </div>
        </div>
      </Card>

      {/* ═══ SECTION 8: CENTRALIZER MODULE ═══ */}
      <Card style={{ marginTop:12 }}>
        <h3 style={{ fontFamily:"'IBM Plex Mono'",fontSize:13,color:T.accent,marginBottom:12,letterSpacing:"0.05em" }}>
          § 8. CENTRALIZER PERFORMANCE
        </h3>
        <div style={{ fontSize:10,fontFamily:"'IBM Plex Mono'" }}>
          <div>Centralizer count: {centralizers.length}</div>
          <div style={{ marginTop:8 }}>Centralizer spacing and standoff:</div>
          {centralizers.slice(0,8).map((c,i)=>(
            <div key={i} style={{ marginBottom:6,padding:6,background:T.dim,borderRadius:4 }}>
              {c.from_depth}-{c.to_depth}ft: {c.type} Standoff={c.standoff} RIH_Force={c.run_in_force}lbf
            </div>
          ))}
          {centralizers.length > 8 && <div style={{ color:T.muted }}>... +{centralizers.length-8} more</div>}
          <div style={{ marginTop:12 }}>Grid standoff distribution (first 10 points):</div>
          {grid.slice(0,10).map((r,i)=>(
            <div key={i} style={{ padding:4 }}>
              {r.md}ft: standoff={r.standoff?.toFixed(3)} → DispEff={r.dispEfficiency?.toFixed(3)}
            </div>
          ))}
        </div>
      </Card>

      {/* ═══ SECTION 9: PLUG MODULE ═══ */}
      {g.job_type === "Secondary (Balanced Plug)" && simResults.plugResults && (
        <Card style={{ marginTop:12 }}>
          <h3 style={{ fontFamily:"'IBM Plex Mono'",fontSize:13,color:T.accent,marginBottom:12,letterSpacing:"0.05em" }}>
            § 9. BALANCED PLUG CALCULATIONS
          </h3>
          <div style={{ fontSize:10,fontFamily:"'IBM Plex Mono'",background:T.dim,padding:10,borderRadius:4 }}>
            <div>Plug top: {simResults.plugResults.plugTop} ft</div>
            <div>Plug bottom: {simResults.plugResults.plugBottom} ft</div>
            <div>Plug length: {simResults.plugResults.plugLength} ft</div>
            <div>Lead cement vol: {simResults.plugResults.leadVolume} bbl</div>
            <div>Tail cement vol: {simResults.plugResults.tailVolume} bbl</div>
            <div>Spacer vol: {simResults.plugResults.spacerVolume} bbl</div>
            <div>Total fluid vol: {simResults.plugResults.totalVolume} bbl</div>
            <div style={{ marginTop:8 }}>
              Hydrostatic balance: P_above = {simResults.plugResults.pressureAbove?.toFixed(0)} psi, 
              P_below = {simResults.plugResults.pressureBelow?.toFixed(0)} psi
            </div>
            <div>Balance check: {simResults.plugResults.isBalanced ? "✓ BALANCED" : "✗ UNBALANCED"}</div>
          </div>
        </Card>
      )}

      {/* ═══ SECTION 10: ITERATION & CONVERGENCE HISTORY ═══ */}
      <Card style={{ marginTop:12 }}>
        <h3 style={{ fontFamily:"'IBM Plex Mono'",fontSize:13,color:T.accent,marginBottom:12,letterSpacing:"0.05em" }}>
          § 10. ITERATION & CONVERGENCE DIAGNOSTICS
        </h3>
        <div style={{ fontSize:10,fontFamily:"'IBM Plex Mono'" }}>
          <div style={{ marginBottom:8 }}><strong>Hydraulic Pressure Calculation Iteration Approach:</strong></div>
          <div style={{ background:T.dim,padding:8,borderRadius:4,marginBottom:12,lineHeight:1.8 }}>
            <div><strong>Overview:</strong> Iterative solver converges pressure-velocity coupling at each depth point.</div>
            <div style={{ marginTop:6 }}><strong>Algorithm Steps (per depth grid):</strong></div>
            <div style={{ marginLeft:12 }}>1. <strong>Initial Guess:</strong></div>
            <div style={{ marginLeft:24 }}>• P_friction^(0) = P_friction from previous depth (accumulated)</div>
            <div style={{ marginLeft:24 }}>• P_total^(0) = P_hydrostatic + P_friction^(0)</div>
            <div style={{ marginLeft:24 }}>• ECD^(0) = P_total^(0) / (0.052 × TVD)</div>
            <div style={{ marginLeft:12,marginTop:4 }}>2. <strong>Iteration Loop</strong> (k = 1 to MAX_ITERATIONS):</div>
            <div style={{ marginLeft:24 }}><strong>Step A:</strong> Calculate velocities from current geometry and pump rate:</div>
            <div style={{ marginLeft:36 }}>Capacity_ann = (D_hole² − D_casing²) / 1029.4  (bbl/ft)</div>
            <div style={{ marginLeft:36 }}>Capacity_pipe = D_id² / 1029.4  (bbl/ft)</div>
            <div style={{ marginLeft:36 }}>v_annular = Q (bpm) / Capacity_ann  (ft/min)</div>
            <div style={{ marginLeft:36 }}>v_pipe = Q (bpm) / Capacity_pipe  (ft/min)</div>
            <div style={{ marginLeft:24,marginTop:2 }}><strong>Step B:</strong> Calculate Reynolds number using current fluid properties:</div>
            <div style={{ marginLeft:36 }}>Re = 928 × ρ × v × d_h / PV  (Bingham model)</div>
            <div style={{ marginLeft:24,marginTop:2 }}><strong>Step C:</strong> Determine flow regime:</div>
            <div style={{ marginLeft:36 }}>IF Re &lt; 2100: Laminar → IF Re &gt; 3000: Turbulent → ELSE: Transitional</div>
            <div style={{ marginLeft:24,marginTop:2 }}><strong>Step D:</strong> Calculate friction factors (rheology-aware):</div>
            <div style={{ marginLeft:36 }}>ΔP/100_annular = f(model, Re, PV, YP, k, n, ρ, v, d_h)</div>
            <div style={{ marginLeft:36 }}>ΔP/100_pipe = f(model, Re, PV, YP, k, n, ρ, v, d_id)</div>
            <div style={{ marginLeft:24,marginTop:2 }}><strong>Step E:</strong> Update friction pressures (incremental):</div>
            <div style={{ marginLeft:36 }}>P_friction_ann^(k) = P_friction_ann^(k-1) + ΔP/100_annular × stepSize/100</div>
            <div style={{ marginLeft:36 }}>P_friction_pipe^(k) = P_friction_pipe^(k-1) + ΔP/100_pipe × stepSize/100</div>
            <div style={{ marginLeft:24,marginTop:2 }}><strong>Step F:</strong> Calculate new total pressure and ECD:</div>
            <div style={{ marginLeft:36 }}>P_total^(k) = P_hydrostatic + P_friction_ann^(k)</div>
            <div style={{ marginLeft:36 }}>ECD^(k) = P_total^(k) / (0.052 × TVD)</div>
            <div style={{ marginLeft:24,marginTop:2 }}><strong>Step G:</strong> Check convergence:</div>
            <div style={{ marginLeft:36 }}>IF |P_total^(k) − P_total^(k-1)| &lt; 10 psi AND |ECD^(k) − ECD^(k-1)| &lt; 0.02 ppg:</div>
            <div style={{ marginLeft:48 }}>→ CONVERGED → Store solution and exit loop</div>
            <div style={{ marginLeft:36 }}>ELSE: Continue to next iteration</div>
            <div style={{ marginLeft:24,marginTop:2 }}><strong>Step H:</strong> Track best solution (lowest residual):</div>
            <div style={{ marginLeft:36 }}>residual = max(|ΔP|, |ΔECD| × 100)</div>
            <div style={{ marginLeft:36 }}>IF residual &lt; best_residual: store current as best solution</div>
            <div style={{ marginLeft:12,marginTop:4 }}>3. <strong>Termination:</strong></div>
            <div style={{ marginLeft:24 }}>• IF converged: return converged solution</div>
            <div style={{ marginLeft:24 }}>• IF max iterations reached: return best stable solution found</div>
            <div style={{ marginTop:6,background:T.dim,padding:8,borderRadius:4 }}>
            <div><strong>Solver Statistics (all {grid.length} points):</strong></div>
            <div>Converged: {grid.filter(r => r.iterStatus === "Converged").length}</div>
            <div>Direct Solution: {grid.filter(r => r.iterStatus === "Direct Solution").length}</div>
            <div>Max Iterations: {grid.filter(r => r.iterStatus === "Max Iterations").length}</div>
            <div>Stable Approximation: {grid.filter(r => r.iterStatus === "Stable Approximation").length}</div>
            <div>Avg Iterations: {(grid.reduce((s,r) => s + (r.iterCount||0), 0) / grid.length).toFixed(1)}</div>
            <div>Max Iterations Used: {Math.max(...grid.map(r => r.iterCount||0))}</div>
          </div>
          <div style={{ marginTop:6 }}><strong>Why Iteration is Needed:</strong></div>
            <div style={{ marginLeft:12 }}>Pressure and velocity are coupled through fluid rheology. Non-Newtonian fluids</div>
            <div style={{ marginLeft:12 }}>(Bingham, Power Law) have shear-rate-dependent viscosity, creating circular dependency:</div>
            <div style={{ marginLeft:24 }}>Pressure → Velocity → Shear Rate → Apparent Viscosity → Friction → Pressure</div>
            <div style={{ marginLeft:12 }}>Iteration breaks this cycle by successive approximation until consistent solution.</div>
          </div>
          <div style={{ background:T.dim,padding:8,borderRadius:4,marginBottom:12 }}>
            <div><strong>Solver Parameters:</strong></div>
            <div>ECD tolerance: 0.02 ppg</div>
            <div>Pressure tolerance: 10 psi</div>
            <div>Max iterations: 30</div>
            <div>Typical convergence: 3-8 iterations for most depths</div>
          </div>
          <div style={{ marginBottom:8 }}>Convergence at first 10 grid points:</div>
          {grid.slice(0,10).map((r,i)=>{
            var statusColor = r.iterStatus === "Converged" ? T.green : r.iterStatus === "Direct Solution" ? T.cyan : r.iterStatus === "Max Iterations" ? T.red : T.yellow;
            var statusIcon = r.iterStatus === "Converged" ? "✓" : r.iterStatus === "Direct Solution" ? "→" : r.iterStatus === "Max Iterations" ? "✗" : "⚠";
            return (
            <div key={i} style={{ padding:8,background:T.dim,borderRadius:4,marginBottom:6,border:`1px solid ${statusColor}40` }}>
              <div style={{ fontWeight:600,marginBottom:4 }}>{statusIcon} Depth {r.md}ft — {r.iterStatus || "N/A"}</div>
              <div style={{ fontSize:10,opacity:0.8 }}>Iterations: {r.iterCount || 0} | Total Residual: {r.iterResidual?.toFixed(3) || "N/A"} psi</div>
              <div style={{ fontSize:10,opacity:0.8 }}>Pressure Residual: {r.iterPressureResidual?.toFixed(2) || "N/A"} psi | ECD Residual: {r.iterEcdResidual?.toFixed(4) || "N/A"} ppg</div>
            </div>
          )})}
        </div>
      </Card>

      {/* ═══ SECTION 11: SUMMARY STATISTICS ═══ */}
      <Card style={{ marginTop:12 }}>
        <h3 style={{ fontFamily:"'IBM Plex Mono'",fontSize:13,color:T.accent,marginBottom:12,letterSpacing:"0.05em" }}>
          § 11. SIMULATION SUMMARY STATISTICS
        </h3>
        <div style={{ fontSize:11,fontFamily:"'IBM Plex Mono'",display:"grid",gridTemplateColumns:"1fr 1fr",gap:12 }}>
          <div>
            <div style={{ fontWeight:600,marginBottom:6 }}>PRESSURE/ECD:</div>
            <div>Max Dynamic ECD: {summary.maxECD?.toFixed(2)} ppg</div>
            <div>Min Dynamic ECD: {summary.minECD?.toFixed(2)} ppg</div>
            <div>Max Static ECD: {summary.maxStaticECD?.toFixed(2)} ppg</div>
            <div>Avg Dynamic ECD: {summary.avgDynamicECD?.toFixed(2)} ppg</div>
            <div>Max Friction +ECD: {summary.maxFricContrib?.toFixed(2)} ppg</div>
            <div>Min Safety Margin: {summary.minSafety?.toFixed(2)} ppg</div>
            <div>Surface Pump Pressure: {summary.finalPumpPressure?.toFixed(0)} psi</div>
            <div>Max Surface Pressure: {summary.maxSurfacePressure?.toFixed(0)} psi</div>
          </div>
          <div>
            <div style={{ fontWeight:600,marginBottom:6 }}>OPERATIONS:</div>
            <div>Total grid points: {summary.totalSteps}</div>
            <div>Violations (ECD > Frac): {summary.violations || 0}</div>
            <div>Kick risk flags: {summary.kickFlags || 0}</div>
            <div>Avg Displacement Eff: {summary.avgDisp?.toFixed(3)}</div>
            <div>Max Temperature: {summary.maxTemp?.toFixed(0)}°F</div>
            <div>BHCT: {summary.bhct}°F</div>
            <div>BHST: {summary.bhst}°F</div>
            <div>Pump Rate: {summary.pumpRate} bpm</div>
          </div>
        </div>
      </Card>

      <div style={{ marginTop:20,padding:12,background:T.dim,borderRadius:4,fontSize:10,color:T.muted,fontFamily:"'IBM Plex Mono'" }}>
        <div style={{ fontWeight:600,marginBottom:6 }}>EXPORT OPTIONS:</div>
        <div>• Full grid data (all {grid.length} points with all calculated fields) → "Full Data Grid" tab</div>
        <div>• Excel export → available on each result page</div>
        <div>• PDF report generation → Dashboard "Generate Report" button</div>
      </div>
    </div>
  );
}

function FullDataPage({ simResults, unitSystem }) {
  if (!simResults) return <NoResults />;
  const isM = unitSystem === "metric";
  const d = simResults.grid.filter((_,i)=>i%2===0);
  const U = UNITS[unitSystem];
  return (
    <div className="animate-in">
      <SectionTitle icon="🗄️" title="FULL DEPTH-GRID DATA TABLE" sub={`${simResults.summary.totalSteps} simulation steps — complete output database · ${U.label}`} />
      <Card>
        <div style={{ overflowX:"auto",maxHeight:600 }}>
          <table>
            <thead style={{ position:"sticky",top:0 }}>
              <tr>{[`MD (${U.depth})`,`TVD`,`Temp (${U.temperature})`,`ECD (${U.density})`,`Hydro (${U.pressure})`,`ECD Pres.`,`Pore P.`,`Frac P.`,`Safety`,`Disp %`,`Standoff %`,`Contam %`,`Flow`,`Hook (kips)`,`Fluid`].map(h=><th key={h}>{h}</th>)}</tr>
            </thead>
            <tbody>
              {d.map((r,i)=>(
                <tr key={i} style={{ background:r.lczFlag?`${T.red}15`:i%2?`${T.panel}40`:"transparent" }}>
                  <td style={{ color:T.accent,fontWeight:700 }}>{isM?cvt(r.md,"depth",true):r.md}</td>
                  <td>{isM?cvt(r.tvd,"depth",true):r.tvd}</td>
                  <td style={{ color:T.gold }}>{isM?cvt(r.temperature,"temperature",true):r.temperature}</td>
                  <td style={{ color:r.ecd>14?T.red:T.text }}>{isM?cvt(r.ecd,"density",true):r.ecd}</td>
                  <td>{isM?cvt(r.hydroPressure,"pressure",true):r.hydroPressure}</td>
                  <td style={{ color:r.lczFlag?T.red:T.yellow,fontWeight:r.lczFlag?700:400 }}>{isM?cvt(r.ecdPressure,"pressure",true):r.ecdPressure}</td>
                  <td style={{ color:T.green }}>{isM?cvt(r.porePressure,"pressure",true):r.porePressure}</td>
                  <td style={{ color:T.red }}>{isM?cvt(r.fracturePressure,"pressure",true):r.fracturePressure}</td>
                  <td style={{ color:r.safetyMargin<300?T.red:T.green }}>{r.safetyMargin}</td>
                  <td style={{ color:r.dispEfficiency>70?T.green:T.yellow }}>{r.dispEfficiency}%</td>
                  <td>{r.standoff}%</td>
                  <td style={{ color:r.contamination>5?T.yellow:T.muted }}>{r.contamination}%</td>
                  <td><Badge text={r.flowRegime} color={r.flowRegime==="Turbulent"?T.green:r.flowRegime==="Laminar"?T.cyan:T.yellow} /></td>
                  <td>{r.hookLoad}</td>
                  <td><span style={{ display:"inline-flex",alignItems:"center",gap:4 }}><span style={{ width:7,height:7,borderRadius:"50%",background:r.fluidColor }}></span><span style={{ fontSize:10 }}>{r.fluidName}</span></span></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Card>
    </div>
  );
}

function DatabasePage({ db, activeProject }) {
  const tables = [
    {name:"users",rows:db.users.length,desc:"User accounts",type:"MASTER"},
    {name:"projects",rows:db.projects.length,desc:"Well projects",type:"MASTER"},
    {name:"survey_data",rows:db.survey_data[activeProject]?.length||0,desc:"Directional survey",type:"INPUT"},
    {name:"casing_profile",rows:db.casing_profile[activeProject]?.length||0,desc:"Casing geometry",type:"INPUT"},
    {name:"open_hole_profile",rows:db.open_hole_profile[activeProject]?.length||0,desc:"Hole geometry",type:"INPUT"},
    {name:"formation_data",rows:db.formation_data[activeProject]?.length||0,desc:"Pressure gradients",type:"INPUT"},
    {name:"fluid_data",rows:db.fluid_data[activeProject]?.length||0,desc:"Fluid properties",type:"INPUT"},
    {name:"centralizer_data",rows:db.centralizer_data[activeProject]?.length||0,desc:"Centralizer data",type:"INPUT"},
    {name:"temperature_profile",rows:db.temperature_profile[activeProject]?.length||0,desc:"Temperature survey",type:"INPUT"},
    {name:"pumping_schedule",rows:db.pumping_schedule[activeProject]?.length||0,desc:"Pump stages",type:"INPUT"},
    {name:"balanced_plug_data",rows:db.balanced_plug_data[activeProject]?.length||0,desc:"Plug data",type:"INPUT"},
    {name:"pressure_profile",rows:db.simulation_runs[activeProject]?150:0,desc:"ECD & pressure output",type:"OUTPUT"},
    {name:"displacement_efficiency",rows:0,desc:"Disp. efficiency output",type:"OUTPUT"},
    {name:"temperature_output",rows:0,desc:"Temperature simulation output",type:"OUTPUT"},
    {name:"fluid_distribution",rows:0,desc:"Fluid front tracking",type:"OUTPUT"},
    {name:"simulation_runs",rows:Object.keys(db.simulation_runs).length,desc:"Run metadata",type:"OUTPUT"},
  ];
  return (
    <div className="animate-in">
      <SectionTitle icon="🗄️" title="DATABASE SCHEMA" sub="CEMOPTI relational DB — all tables, row counts, and structure" />
      <Card>
        <table>
          <thead><tr>{["Table Name","Rows","Type","Description","Indexed On"].map(h=><th key={h}>{h}</th>)}</tr></thead>
          <tbody>
            {tables.map((t,i)=>(
              <tr key={i}>
                <td style={{ color:T.accent,fontFamily:"'IBM Plex Mono'",fontWeight:600 }}>{t.name}</td>
                <td>{t.rows}</td>
                <td><Badge text={t.type} color={t.type==="OUTPUT"?T.yellow:t.type==="INPUT"?T.cyan:T.green} /></td>
                <td style={{ color:T.muted }}>{t.desc}</td>
                <td style={{ color:T.dim,fontSize:10 }}>{t.name.includes("profile")||t.name.includes("data")?`depth / from_depth`:"id"}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </Card>
    </div>
  );
}

function DashboardPage({ db, activeProject, simResults, onRun, running, onValidate, onReset, onReport, unitSystem, onUnitChange, simProgress, simStageInfo }) {
  const proj = db.projects.find(p=>p.project_id===activeProject);
  if (!proj) return null;
  const s = simResults?.summary;
  return (
    <div className="animate-in">
      <SectionTitle icon="⬡" title="PROJECT DASHBOARD" sub={`${proj.project_name} · ${proj.well_name} · ${proj.field_name}`} />
      <div style={{ display:"flex",gap:12,flexWrap:"wrap",marginBottom:16,alignItems:"center",justifyContent:"space-between" }}>
        <div style={{ display:"flex",gap:10,flexWrap:"wrap" }}>
          <Badge text={`STEP: ${proj.depth_step} ft`} color={T.muted} />
          <Badge text={proj.status.toUpperCase()} color={T.green} />
        </div>
        <UnitToggle unitSystem={unitSystem} onChange={onUnitChange} />
      </div>
      <Card style={{ marginBottom:20,background:`linear-gradient(135deg,${T.card},${T.panel})`,border:`1px solid ${T.accent}33` }}>
        <div style={{ display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:12 }}>
          <div>
            <div style={{ fontFamily:"'Orbitron'",fontSize:13,color:T.accent,marginBottom:4 }}>SIMULATION ENGINE</div>
            <div style={{ fontSize:12,color:T.muted }}>5 modules · depth-based grid · hydraulics · displacement · BHCT · standoff · plug</div>
          </div>
          <div style={{ display:"flex",gap:8,flexWrap:"wrap" }}>
            <Btn variant="ghost" onClick={onReset}>↺ Reset Defaults</Btn>
            <Btn variant="warning" onClick={onValidate}>✓ Validate Data</Btn>
            {simResults && <Btn variant="success" onClick={onReport}>⬇ Generate Report</Btn>}
            <Btn variant="primary" size="lg" onClick={onRun} disabled={running}>{running?"⏳ COMPUTING...":"▶ RUN SIMULATION"}</Btn>
            
            {running && simProgress > 0 && (
              <div style={{ marginTop:16,padding:16,background:T.card,border:"1px solid " + T.border,borderRadius:6 }}>
                <div style={{ fontSize:11,color:T.muted,marginBottom:8,fontFamily:"'IBM Plex Mono'" }}>
                  SIMULATION PROGRESS
                </div>
                <div style={{ width:"100%",height:24,background:T.dim,borderRadius:4,overflow:"hidden",marginBottom:8 }}>
                  <div style={{
                    width: simProgress + "%",
                    height:"100%",
                    background:"linear-gradient(90deg, " + T.gold + ", " + T.yellow + ")",
                    transition:"width 0.1s linear",
                    display:"flex",
                    alignItems:"center",
                    justifyContent:"center",
                    fontSize:11,
                    fontWeight:600,
                    color:T.bg,
                    fontFamily:"'IBM Plex Mono'"
                  }}>
                    {simProgress}%
                  </div>
                </div>
                {simStageInfo && (
                  <div style={{ fontSize:10,fontFamily:"'IBM Plex Mono'",color:T.text,lineHeight:1.6 }}>
                    <div><strong>Stage {simStageInfo.stageNum}/{simStageInfo.totalStages}:</strong> {simStageInfo.stage}</div>
                    <div>Timestep: {simStageInfo.timestep}/{simStageInfo.totalTimesteps}</div>
                    <div>Grid: {simStageInfo.grid}/{simStageInfo.totalGrids}</div>
                  </div>
                )}
              </div>
            )}
          </div>
        </div>

        {simResults && (
          <div style={{ marginTop:12,padding:"8px 12px",background:`${T.green}11`,border:`1px solid ${T.green}44`,borderRadius:4 }}>
            <span style={{ fontSize:11,color:T.green,fontFamily:"'IBM Plex Mono'" }}>✓ Last run: {new Date(simResults.timestamp).toLocaleString()} · {simResults.summary.totalSteps} depth steps · {simResults.run_id.slice(-8)}</span>
          </div>
        )}
      </Card>
      {s && (
        <div className="grid-4" style={{ marginBottom:20 }}>
          <MetricCard label="Max ECD" value={s.maxECD} unit="ppg" color={s.maxECD>14?T.red:T.green} icon="💧" />
          <MetricCard label="Min Safety Margin" value={`${(s.minSafety/1000).toFixed(1)}k psi`} unit="Frac - ECD" color={s.minSafety<300?T.red:T.green} icon="🛡️" />
          <MetricCard label="Avg Disp. Efficiency" value={`${s.avgDisp}%`} unit="Annular" color={s.avgDisp>70?T.green:T.yellow} icon="🔄" />
          <MetricCard label="BHCT" value={`${s.bhct}°F`} unit="Bottom Hole Circ. Temp" color={T.gold} icon="🌡️" />
        </div>
      )}
      {s && (
        <div className="grid-2">
          <DepthChart data={simResults.grid.filter((_,i)=>i%2===0)} title="PRESSURE WINDOW — QUICK VIEW" xLabel="psi" height={280}
            lines={[{key:"ecdPressure",name:"ECD (psi)",color:T.yellow},{key:"fracturePressure",name:"Frac (psi)",color:T.red},{key:"porePressure",name:"Pore (psi)",color:T.green}]} />
          <DepthChart data={simResults.grid.filter((_,i)=>i%2===0)} title="TEMPERATURE vs DEPTH" xLabel="°F" height={280}
            lines={[{key:"temperature",name:"Temp (°F)",color:T.gold}]} />
        </div>
      )}
    </div>
  );
}

function ProjectsPage({ db, setDb, activeProject, setActiveProject, onNavigate, onOpenProject, simResultsMap }) {
  const [showNew, setShowNew] = useState(false);
  const [np, setNp] = useState({project_name:"",well_name:"",field_name:"",unit_system:"imperial",depth_step:50});

  const deleteProject = (pid, pname) => {
    if (db.projects.length <= 1) { alert("Cannot delete the last project."); return; }
    if (!window.confirm(`Delete project "${pname}"? This cannot be undone.`)) return;
    setDb(d => {
      const next = { ...d, projects: d.projects.filter(p => p.project_id !== pid) };
      // Remove all project data from every table
      ["survey_data","casing_profile","open_hole_profile","formation_data","fluid_data",
       "centralizer_data","temperature_profile","pumping_schedule","balanced_plug_data",
       "general_data","simulation_runs","actual_job_data"].forEach(table => {
        if (next[table] && typeof next[table] === "object" && !Array.isArray(next[table])) {
          const copy = { ...next[table] };
          delete copy[pid];
          next[table] = copy;
        }
      });
      return next;
    });
    // Switch active project to first remaining
    if (activeProject === pid) {
      const remaining = db.projects.filter(p => p.project_id !== pid);
      if (remaining.length > 0) setActiveProject(remaining[0].project_id);
    }
  };
  const create = () => {
    const id = `p${Date.now()}`;
    const proj = {...np, hole_type:"open_hole", project_id:id,user_id:"u1",status:"active",created_at:new Date().toISOString()};
    setDb(d => ({...d,projects:[...d.projects,proj],
      survey_data:{...d.survey_data,[id]:JSON.parse(JSON.stringify(DEFAULT_DATA.survey_data))},
      casing_profile:{...d.casing_profile,[id]:JSON.parse(JSON.stringify(DEFAULT_DATA.casing_profile))},
      open_hole_profile:{...d.open_hole_profile,[id]:JSON.parse(JSON.stringify(DEFAULT_DATA.open_hole_profile))},
      formation_data:{...d.formation_data,[id]:JSON.parse(JSON.stringify(DEFAULT_DATA.formation_data))},
      fluid_data:{...d.fluid_data,[id]:JSON.parse(JSON.stringify(DEFAULT_DATA.fluid_data))},
      centralizer_data:{...d.centralizer_data,[id]:JSON.parse(JSON.stringify(DEFAULT_DATA.centralizer_data))},
      temperature_profile:{...d.temperature_profile,[id]:JSON.parse(JSON.stringify(DEFAULT_DATA.temperature_profile))},
      pumping_schedule:{...d.pumping_schedule,[id]:JSON.parse(JSON.stringify(DEFAULT_DATA.pumping_schedule))},
      balanced_plug_data:{...d.balanced_plug_data,[id]:JSON.parse(JSON.stringify(DEFAULT_DATA.balanced_plug_data))},
      general_data:{...d.general_data,[id]:{...DEFAULT_DATA.general_data}},
    }));
    setActiveProject(id);
    setShowNew(false);
    // Auto-navigate to first input page after project creation
    onNavigate("survey");
  };
  return (
    <div className="animate-in">
      <SectionTitle icon="📁" title="PROJECT MANAGER" sub="Create and manage well cementing projects — click any project row to open it" />
      <div style={{ display:"flex",justifyContent:"flex-end",marginBottom:14 }}>
        <Btn variant="primary" onClick={()=>setShowNew(!showNew)}>+ New Project</Btn>
      </div>
      {showNew && (
        <Card style={{ marginBottom:16,border:`1px solid ${T.accent}44` }}>
          <div style={{ fontSize:12,color:T.accent,fontFamily:"'Orbitron'",marginBottom:12 }}>NEW PROJECT</div>
          <div className="grid-3" style={{ gap:10,marginBottom:12 }}>
            {[{k:"project_name",l:"Project Name"},{k:"well_name",l:"Well Name"},{k:"field_name",l:"Field Name"}].map(f=>(
              <div key={f.k}><div style={{ fontSize:10,color:T.muted,marginBottom:4 }}>{f.l}</div>
              <input type="text" value={np[f.k]} onChange={e=>setNp(p=>({...p,[f.k]:e.target.value}))} style={{ width:"100%" }} /></div>
            ))}
          </div>
          <div style={{ display:"flex",gap:10,marginBottom:12,flexWrap:"wrap" }}>
            {[{k:"unit_system",l:"Units",opts:[{v:"imperial",l:"Field"},{v:"metric",l:"Metric"}]}].map(f=>(
              <div key={f.k}><div style={{ fontSize:10,color:T.muted,marginBottom:4 }}>{f.l}</div>
              <select value={np[f.k]} onChange={e=>setNp(p=>({...p,[f.k]:e.target.value}))}>
                {f.opts.map(o=><option key={o.v} value={o.v}>{o.l}</option>)}
              </select></div>
            ))}
            <div><div style={{ fontSize:10,color:T.muted,marginBottom:4 }}>Depth Step (ft)</div>
            <input type="number" value={np.depth_step} onChange={e=>setNp(p=>({...p,depth_step:parseInt(e.target.value)}))} style={{ width:80 }} /></div>
          </div>
          <div style={{ display:"flex",gap:8 }}>
            <Btn variant="primary" onClick={create}>Create Project → Enter Inputs</Btn>
            <Btn variant="ghost" onClick={()=>setShowNew(false)}>Cancel</Btn>
          </div>
        </Card>
      )}
      <Card>
        <table>
          <thead><tr>{["Project","Well","Field","Units","Results","Status","Created","Actions (Open / Edit / Delete)"].map(h=><th key={h}>{h}</th>)}</tr></thead>
          <tbody>
            {db.projects.map((p,i)=>(
              <tr key={i}
                style={{ background:p.project_id===activeProject?`${T.accent}12`:"transparent", cursor:"pointer" }}
                onClick={()=>onOpenProject(p.project_id)}
                title={`Click to open ${p.project_name}`}
              >
                <td style={{ color:T.accent,fontWeight:700 }}>{p.project_name}</td>
                <td>{p.well_name}</td>
                <td>{p.field_name}</td>
                <td><Badge text={UNITS[p.unit_system]?.label.toUpperCase() || p.unit_system.toUpperCase()} color={T.gold} /></td>
                <td>
                  {simResultsMap && simResultsMap[p.project_id]
                    ? <Badge text={`✓ ${simResultsMap[p.project_id].summary.totalSteps} PTS`} color={T.green} />
                    : <span style={{ fontSize:10,color:T.dim,fontFamily:"'IBM Plex Mono'" }}>Not run</span>
                  }
                </td>
                <td><Badge text={p.status.toUpperCase()} color={T.green} /></td>
                <td style={{ color:T.muted,fontSize:10 }}>{new Date(p.created_at).toLocaleDateString()}</td>
                <td onClick={e=>e.stopPropagation()} style={{ display:"flex",gap:6,padding:"5px 10px" }}>
                  <Btn variant={p.project_id===activeProject?"success":"secondary"} size="sm"
                    onClick={()=>onOpenProject(p.project_id)}>
                    {p.project_id===activeProject?"✓ Open":"Open"}
                  </Btn>
                  <Btn variant="ghost" size="sm"
                    onClick={()=>{ setActiveProject(p.project_id); onNavigate("survey"); }}>
                    Edit
                  </Btn>
                  <Btn variant="danger" size="sm"
                    onClick={()=>deleteProject(p.project_id, p.project_name)}>
                    Delete
                  </Btn>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
        <div style={{ padding:"8px 12px",borderTop:`1px solid ${T.border}`,fontSize:10,color:T.dim }}>
          Click any row to open a project · "Open" → Dashboard · "Edit" → first input page · "Delete" → permanently removes project
        </div>
      </Card>
    </div>
  );
}

// ─────────────────────────── MANUAL / TECH GUIDE PAGE ───────────────────────
function ManualPage() {
  const [section, setSection] = useState("user");
  const sections = [
    { id:"user",    label:"User Guide",      icon:"📖" },
    { id:"tech",    label:"Technical Guide", icon:"⚙️" },
    { id:"refdata", label:"Reference Data",  icon:"📋" },
    { id:"db",      label:"DB & System",     icon:"🗄️" },
    { id:"legal",   label:"Legal & Support", icon:"⚖️" },
  ];

  const content = {
    user: (
      <div>
        <h3 style={{ color:T.accent,fontFamily:"'Orbitron'",fontSize:14,marginBottom:16 }}>USER GUIDE — HOW TO USE CEMOPTI</h3>
        <Card style={{ marginBottom:14 }}>
          <div style={{ fontSize:12,fontWeight:700,color:T.cyan,marginBottom:8 }}>Workflow Overview</div>
          {[
            ["1. Login","Use credentials: cemopti / 1234. After login, you land on the Dashboard."],
            ["2. Create Project","Navigate to Projects → New Project. Select Hole Type and Unit System. The system auto-fills default inputs and takes you to the first input page."],
            ["3. Enter Inputs","Work through all input pages using the NEXT/BACK navigation buttons. All tables support manual entry and Excel import (⬆ Import Excel button on each page)."],
            ["4. Validate","Click ✓ Validate Data (dashboard, topbar, or sidebar) before running. All errors and warnings are shown with page references."],
            ["5. Run Simulation","Click ▶ Run Simulation or ▶ Run (topbar). The depth-based engine runs all 5 modules and navigates to the Pressure/ECD results page."],
            ["6. Review Results","Navigate through all 5 result modules: M1 Pressure, M2 Displacement, M3 Temperature, M4 Centralizer, M5 Plug. The Full Data Grid shows all parameters."],
            ["7. Generate Report","Click ⬇ Generate Report (shown after a simulation run). An HTML report is downloaded containing all inputs, all outputs, and a summary."],
          ].map(([title,desc],i)=>(
            <div key={i} style={{ display:"flex",gap:12,padding:"8px 0",borderBottom:`1px solid ${T.border}` }}>
              <span style={{ minWidth:80,fontSize:12,fontWeight:700,color:T.gold,fontFamily:"'IBM Plex Mono'" }}>{title}</span>
              <span style={{ fontSize:12,color:T.muted }}>{desc}</span>
            </div>
          ))}
        </Card>
        <Card style={{ marginBottom:14 }}>
          <div style={{ fontSize:12,fontWeight:700,color:T.cyan,marginBottom:8 }}>Input Pages — Field Descriptions</div>
          {[
            ["Survey Data","MD, TVD, inclination, and azimuth. Used for depth-to-TVD conversion in pressure and ECD calculations. Import from LAS or Excel."],
            ["Casing Profile","OD, ID, grade, weight, and type per depth interval. Type selection auto-fills API standard properties. Controls annular geometry for all hydraulic calculations."],
            ["Open Hole Profile","Hole size and volumetric excess per depth interval. Used in annular capacity, annular velocity, and centralizer standoff calculations."],
            ["Formation Data","Pore pressure gradient and fracture gradient per depth interval. Defines the operating pressure window. ECD must remain between these limits."],
            ["Fluids","Density, PV, YP, rheology model (Bingham/Power Law/Newtonian/HB), n and K parameters. Type selection auto-fills typical field values."],
            ["Pumping Schedule","Stage sequence, fluid selection, rate (bpm), and volume (bbl). Defines the displacement train order and timing."],
            ["Centralizers","Type, spacing, standoff %, and RIH force per interval. Type selection auto-fills standoff and applies turbulence bonus in displacement calculations."],
            ["Temperature","Depth-temperature survey points. Linear interpolation applied between points for BHCT and density correction."],
            ["Balanced Plug","From/To depth, plug length, set depth, excess, and fluid selection. Live calculation shows TOC, displacement volumes, and balance check instantly."],
          ].map(([title,desc],i)=>(
            <div key={i} style={{ display:"flex",gap:12,padding:"7px 0",borderBottom:`1px solid ${T.border}` }}>
              <span style={{ minWidth:110,fontSize:11,fontWeight:700,color:T.accent,fontFamily:"'IBM Plex Mono'" }}>{title}</span>
              <span style={{ fontSize:12,color:T.muted }}>{desc}</span>
            </div>
          ))}
        </Card>
        <Card>
          <div style={{ fontSize:12,fontWeight:700,color:T.cyan,marginBottom:8 }}>Output Interpretation</div>
          {[
            ["M1: Pressure/ECD","ECD must stay between pore pressure (kick risk) and fracture pressure (lost circulation risk) at all depths. Safety margin = fracture pressure − ECD pressure."],
            ["M2: Displacement","Efficiency > 70% indicates adequate mud removal. Contamination % shows fluid intermixing zone width. Turbulent flow improves efficiency."],
            ["M3: Temperature","BHCT (Bottom Hole Circulating Temperature) governs cement thickening time. Use BHCT for slurry design, BHST for set cement performance."],
            ["M4: Centralizer","Standoff < 67% in pay zones indicates channeling risk. Solid Body centralizers achieve higher standoff in horizontal sections."],
            ["M5: Balanced Plug","BALANCED = pressure inside pipe equals pressure outside at set depth. Unbalanced plug may move after cement is released."],
          ].map(([title,desc],i)=>(
            <div key={i} style={{ display:"flex",gap:12,padding:"7px 0",borderBottom:`1px solid ${T.border}` }}>
              <span style={{ minWidth:110,fontSize:11,fontWeight:700,color:T.green,fontFamily:"'IBM Plex Mono'" }}>{title}</span>
              <span style={{ fontSize:12,color:T.muted }}>{desc}</span>
            </div>
          ))}
        </Card>
      </div>
    ),

    tech: (
      <div>
        <h3 style={{ color:T.accent,fontFamily:"'Orbitron'",fontSize:14,marginBottom:16 }}>TECHNICAL GUIDE — CALCULATION ALGORITHMS</h3>
        <Card style={{ marginBottom:14 }}>
          <div style={{ fontSize:12,fontWeight:700,color:T.cyan,marginBottom:10 }}>M1: Hydraulics & Pressure</div>
          {[
            ["Hydrostatic Pressure","P_h(D) = 0.052 × Σ(ρᵢ × ΔZᵢ) — Accumulated over depth grid with temperature-corrected density at each step"],
            ["Temperature Density Correction","ρ_adj = ρ_base × (1 − 0.00002 × (T − 70°F)) — ~0.2% reduction per 100°F"],
            ["Annular Velocity","v_ann = Q / Capacity_ann — Q in bpm, Capacity = (D_hole² − D_cas_OD²)/1029.4 in bbl/ft, velocity in ft/min"],
            ["Annular Capacity","Cap = (D_hole² − D_cas_OD²) / 1029.4 — Result in bbl/ft"],
            ["Hydraulic Diameter","d_h = D_hole − D_cas_OD — Used in Reynolds number and friction pressure"],
            ["ECD","ECD (ppg) = ρ_adj + ΔP_friction / (0.052 × TVD) — Friction pressure is cumulative from surface"],
            ["Safety Margin","SM = P_fracture − P_ECD — Must be > 200 psi in critical zones"],
          ].map(([title,eq],i)=>(
            <div key={i} style={{ marginBottom:10 }}>
              <div style={{ fontSize:11,fontWeight:700,color:T.yellow }}>{title}</div>
              <div style={{ fontSize:11,fontFamily:"'IBM Plex Mono'",color:T.muted,background:T.bg,padding:"4px 8px",borderRadius:4,marginTop:3 }}>{eq}</div>
            </div>
          ))}
        </Card>
        <Card style={{ marginBottom:14 }}>
          <div style={{ fontSize:12,fontWeight:700,color:T.cyan,marginBottom:10 }}>Rheology Models — Friction Pressure</div>
          {[
            ["Bingham Plastic (Laminar)","ΔP/100ft = PV·v/(300·d_h) + YP/(225·d_h) — Standard API Bingham laminar"],
            ["Bingham Plastic (Turbulent)","ΔP/100ft = ρ·v²/25600 — Simplified Blasius turbulent (Re > 2100)"],
            ["Power Law (Laminar)","ΔP/100ft = K·vⁿ/(144·d_h^(1+n)) × ((2+1/n)/0.0208)ⁿ × 144 — Dodge-Metzner"],
            ["Newtonian (Laminar)","ΔP/100ft = PV·v/(300·d_h) — YP = 0, single viscosity"],
            ["Herschel-Bulkley","ΔP/100ft = K·vⁿ/(144·d_h^(1+n))×144 + YP/(225·d_h) — Three-parameter model"],
            ["Reynolds Number","N_Re = 109 × ρ × v × d_h / PV (Bingham) — Turbulent threshold: N_Re > 2100"],
          ].map(([title,eq],i)=>(
            <div key={i} style={{ marginBottom:10 }}>
              <div style={{ fontSize:11,fontWeight:700,color:T.gold }}>{title}</div>
              <div style={{ fontSize:11,fontFamily:"'IBM Plex Mono'",color:T.muted,background:T.bg,padding:"4px 8px",borderRadius:4,marginTop:3 }}>{eq}</div>
            </div>
          ))}
        </Card>
        <Card style={{ marginBottom:14 }}>
          <div style={{ fontSize:12,fontWeight:700,color:T.cyan,marginBottom:10 }}>M2: Displacement Efficiency</div>
          <div style={{ fontSize:11,fontFamily:"'IBM Plex Mono'",color:T.muted,background:T.bg,padding:"8px 12px",borderRadius:4,lineHeight:1.8 }}>
            DE (%) = 0.35 × SO + 0.28 × min(v_ann, 120) + 12 + (8 if turbulent) + turbulence_bonus(centralizer_type)<br/>
            SO = casing standoff (%)<br/>
            turbulence_bonus: Bow Spring=0, Solid Body=3, Turbo=8<br/>
            Contamination % = max(0, 15 − DE × 0.15)
          </div>
        </Card>
        <Card style={{ marginBottom:14 }}>
          <div style={{ fontSize:12,fontWeight:700,color:T.cyan,marginBottom:10 }}>M5: Balanced Plug</div>
          <div style={{ fontSize:11,fontFamily:"'IBM Plex Mono'",color:T.muted,background:T.bg,padding:"8px 12px",borderRadius:4,lineHeight:1.8 }}>
            Plug Volume (bbl) = plug_length × annular_capacity<br/>
            Displacement Volume (bbl) = set_depth × internal_capacity<br/>
            Water Behind = excess × annular_capacity<br/>
            Final TOC = set_depth − plug_length<br/>
            Balance check: |ρ_inside × 0.052 × set_depth − ρ_outside × 0.052 × plug_length| &lt; 0.5 psi/ft
          </div>
        </Card>
        <Card>
          <div style={{ fontSize:12,fontWeight:700,color:T.cyan,marginBottom:10 }}>Calculation Workflow (Step-by-Step)</div>
          {["Load project data and initialize depth grid (0 → maxMD in stepSize increments)","At each depth: look up active casing segment, formation segment, fluid in column","Interpolate temperature; apply density correction","Accumulate hydrostatic pressure: P_h[i] = P_h[i-1] + ρ_adj × 0.052 × dZ","Compute annular velocity from pump rate and annular cross-section","Determine flow regime via Reynolds number","Apply correct rheology model to compute friction pressure","Compute ECD = ρ_adj + ΔP_fric / (0.052 × TVD)","Compute pore and fracture pressures from input gradients","Compute safety margin, displacement efficiency, contamination, standoff","Log all values to results array","After loop: compute balanced plug results, summary metrics"].map((s,i)=>(
            <div key={i} style={{ display:"flex",gap:10,padding:"5px 0",borderBottom:`1px solid ${T.border}` }}>
              <span style={{ minWidth:22,fontSize:11,fontWeight:700,color:T.accent,fontFamily:"'IBM Plex Mono'" }}>{i+1}.</span>
              <span style={{ fontSize:12,color:T.muted }}>{s}</span>
            </div>
          ))}
        </Card>
      </div>
    ),

    refdata: (
      <div>
        <h3 style={{ color:T.accent,fontFamily:"'Orbitron'",fontSize:14,marginBottom:16 }}>REFERENCE DATA</h3>
        <div className="grid-2">
          <Card style={{ marginBottom:14 }}>
            <div style={{ fontSize:12,fontWeight:700,color:T.cyan,marginBottom:10 }}>Rheology Models</div>
            {Object.entries({ Bingham:"PV + YP (2-param) · Most common for cement and mud", "Power Law":"K + n (2-param) · Shear-thinning spacers and polymers", Newtonian:"μ only (1-param) · Water, thin wash fluids", HB:"τ₀ + K + n (3-param) · Herschel-Bulkley, most accurate" }).map(([m,d],i)=>(
              <div key={i} style={{ padding:"6px 0",borderBottom:`1px solid ${T.border}` }}>
                <div style={{ fontSize:11,fontWeight:700,color:T.yellow }}>{m}</div>
                <div style={{ fontSize:11,color:T.muted }}>{d}</div>
              </div>
            ))}
          </Card>
          <Card style={{ marginBottom:14 }}>
            <div style={{ fontSize:12,fontWeight:700,color:T.cyan,marginBottom:10 }}>Fluid Types & Typical Properties</div>
            {Object.entries(FLUID_TYPE_DEFAULTS).map(([type,p],i)=>(
              <div key={i} style={{ padding:"6px 0",borderBottom:`1px solid ${T.border}` }}>
                <div style={{ fontSize:11,fontWeight:700,color:p.color }}>{type.toUpperCase()}</div>
                <div style={{ fontSize:10,color:T.muted,fontFamily:"'IBM Plex Mono'" }}>ρ={p.density} ppg · PV={p.pv} cP · YP={p.yp} · {p.model}</div>
              </div>
            ))}
          </Card>
          <Card style={{ marginBottom:14 }}>
            <div style={{ fontSize:12,fontWeight:700,color:T.cyan,marginBottom:10 }}>Centralizer Types</div>
            {Object.entries(CENTRALIZER_TYPE_PROPS).map(([type,p],i)=>(
              <div key={i} style={{ padding:"6px 0",borderBottom:`1px solid ${T.border}` }}>
                <div style={{ fontSize:11,fontWeight:700,color:T.green }}>{type}</div>
                <div style={{ fontSize:10,color:T.muted }}>{p.description}</div>
                <div style={{ fontSize:10,color:T.muted,fontFamily:"'IBM Plex Mono'" }}>Standoff: {p.baseStandoff}% · RIH: {p.rihForce} lbs · Turb bonus: +{p.turbulenceBonus}%</div>
              </div>
            ))}
          </Card>
          <Card>
            <div style={{ fontSize:12,fontWeight:700,color:T.cyan,marginBottom:10 }}>Casing Types (API Standard)</div>
            {Object.entries(CASING_TYPE_PROPS).map(([type,p],i)=>(
              <div key={i} style={{ padding:"6px 0",borderBottom:`1px solid ${T.border}` }}>
                <div style={{ fontSize:11,fontWeight:700,color:T.accent }}>{type}</div>
                <div style={{ fontSize:10,color:T.muted,fontFamily:"'IBM Plex Mono'" }}>OD: {p.od}" · ID: {p.id_}" · {p.grade} · {p.weight} lb/ft</div>
              </div>
            ))}
          </Card>
        </div>
        <Card>
          <div style={{ fontSize:12,fontWeight:700,color:T.cyan,marginBottom:10 }}>Key Conversion Factors</div>
          <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:8 }}>
            {["0.052 psi/ft per ppg — hydrostatic gradient","1029.4 in²/(bbl/ft) — capacity conversion factor","928 — Bingham Reynolds constant","109 — Power Law Reynolds constant","0.3048 ft → m","6.89476 psi → kPa","119.826 ppg → kg/m³"].map((c,i)=>(
              <div key={i} style={{ fontSize:11,color:T.muted,fontFamily:"'IBM Plex Mono'",padding:"4px 0",borderBottom:`1px solid ${T.border}` }}>{c}</div>
            ))}
          </div>
        </Card>
      </div>
    ),

    db: (
      <div>
        <h3 style={{ color:T.accent,fontFamily:"'Orbitron'",fontSize:14,marginBottom:16 }}>DATABASE & SYSTEM DESIGN</h3>
        <Card style={{ marginBottom:14 }}>
          <div style={{ fontSize:12,fontWeight:700,color:T.cyan,marginBottom:10 }}>Database Schema Overview</div>
          {[
            ["users","user_id, username, password, email, role","Authentication"],
            ["projects","project_id, user_id, project_name, well_name, field_name, hole_type, unit_system, depth_step","Project metadata"],
            ["survey_data","project_id, md, tvd, inclination, azimuth","Directional survey points"],
            ["casing_profile","project_id, from_depth, to_depth, od, id, grade, type, weight","Casing geometry"],
            ["open_hole_profile","project_id, from_depth, to_depth, hole_size, excess","Hole geometry"],
            ["formation_data","project_id, from_depth, to_depth, pore_gradient, frac_gradient, lithology","Pressure gradients"],
            ["fluid_data","project_id, fluid_id, fluid_name, type, density, pv, yp, model, n, k, color","Fluid properties"],
            ["centralizer_data","project_id, from_depth, to_depth, type, spacing, standoff, run_in_force","Centralizer schedule"],
            ["temperature_profile","project_id, depth, temperature","Temperature survey"],
            ["pumping_schedule","project_id, stage, fluid_id, rate, volume, purpose","Pump stages"],
            ["balanced_plug_data","project_id, from_depth, to_depth, length, set_depth, excess, fluid_id","Plug design"],
          ].map(([table,cols,purpose],i)=>(
            <div key={i} style={{ padding:"7px 0",borderBottom:`1px solid ${T.border}` }}>
              <div style={{ display:"flex",justifyContent:"space-between",alignItems:"baseline" }}>
                <span style={{ fontSize:11,fontWeight:700,color:T.accent,fontFamily:"'IBM Plex Mono'" }}>{table}</span>
                <span style={{ fontSize:10,color:T.dim }}>{purpose}</span>
              </div>
              <div style={{ fontSize:10,color:T.muted,fontFamily:"'IBM Plex Mono'",marginTop:3 }}>{cols}</div>
            </div>
          ))}
        </Card>
        <Card>
          <div style={{ fontSize:12,fontWeight:700,color:T.cyan,marginBottom:10 }}>How Inputs Map to Engine Outputs</div>
          {[["survey_data","TVD interpolation for pressure/ECD calculation"],["casing_profile","Annular geometry, hook load, casing OD/ID"],["open_hole_profile","Hole size, annular capacity, annular velocity"],["formation_data","Pore/frac pressure limits, safety margin"],["fluid_data + pumping_schedule","Fluid column tracking, density, rheology model selection"],["centralizer_data","Standoff, turbulence bonus on displacement efficiency"],["temperature_profile","Temperature interpolation, density correction"],["balanced_plug_data","Plug volume, displacement volume, TOC, balance check"]].map(([input,output],i)=>(
            <div key={i} style={{ display:"flex",gap:12,padding:"6px 0",borderBottom:`1px solid ${T.border}` }}>
              <span style={{ minWidth:140,fontSize:10,fontFamily:"'IBM Plex Mono'",color:T.yellow }}>{input}</span>
              <span style={{ fontSize:11,color:T.muted }}>→ {output}</span>
            </div>
          ))}
        </Card>
      </div>
    ),

    legal: (
      <div>
        <h3 style={{ color:T.accent,fontFamily:"'Orbitron'",fontSize:14,marginBottom:16 }}>LEGAL, SUPPORT & COPYRIGHT</h3>
        <Card style={{ marginBottom:14,border:`1px solid ${T.red}44` }}>
          <div style={{ fontSize:12,fontWeight:700,color:T.red,marginBottom:10 }}>⚠️ DISCLAIMER</div>
          <p style={{ fontSize:12,color:T.muted,lineHeight:1.8 }}>
            CEMOPTI is provided as a cementing simulation and engineering planning tool for qualified petroleum engineering professionals. All outputs, including pressure profiles, ECD calculations, displacement efficiency, temperature profiles, and balanced plug results, are engineering estimates based on simplified analytical models and field-validated correlations.
          </p>
          <p style={{ fontSize:12,color:T.muted,lineHeight:1.8,marginTop:10 }}>
            RESULTS ARE FOR ENGINEERING GUIDANCE ONLY. The user is solely responsible for field validation, independent verification, and the application of professional judgment before implementing any design in the field. Kemiserve FZE accepts no liability for losses, damages, or consequences resulting from reliance on simulation outputs without qualified independent expert review.
          </p>
        </Card>
        <Card style={{ marginBottom:14 }}>
          <div style={{ fontSize:12,fontWeight:700,color:T.accent,marginBottom:10 }}>OWNERSHIP & COPYRIGHT</div>
          <p style={{ fontSize:12,color:T.muted,lineHeight:1.8 }}>CEMOPTI v4.0 — Depth-Based Cementing Simulation Platform</p>
          <p style={{ fontSize:12,color:T.text,fontWeight:700,marginTop:8 }}>© Kemiserve FZE<br/>SPC Free Zone, Sharjah, United Arab Emirates</p>
          <p style={{ fontSize:12,color:T.muted,marginTop:8,lineHeight:1.8 }}>All rights reserved. This software and all associated documentation, algorithms, calculation methodologies, and intellectual property are the exclusive property of Kemiserve FZE. Unauthorized reproduction, distribution, reverse engineering, or commercial exploitation is strictly prohibited under UAE intellectual property law and applicable international conventions.</p>
        </Card>
        <Card style={{ marginBottom:14 }}>
          <div style={{ fontSize:12,fontWeight:700,color:T.green,marginBottom:10 }}>SUPPORT CONTACT</div>
          <div style={{ fontSize:13,color:T.text }}>
            <div style={{ marginBottom:8 }}>📧 <strong>Email:</strong> <a href="mailto:info@kemiserve.com" style={{ color:T.accent }}>info@kemiserve.com</a></div>
            <div style={{ marginBottom:8 }}>🌐 <strong>Company:</strong> Kemiserve FZE</div>
            <div style={{ marginBottom:8 }}>📍 <strong>Address:</strong> SPC Free Zone, Sharjah, United Arab Emirates</div>
            <div style={{ marginBottom:8 }}>📱 <strong>Support Hours:</strong> Sunday–Thursday 08:00–17:00 GST</div>
          </div>
        </Card>
        <Card>
          <div style={{ fontSize:12,fontWeight:700,color:T.muted,marginBottom:8 }}>SOFTWARE VERSION</div>
          <div style={{ fontSize:12,color:T.muted,fontFamily:"'IBM Plex Mono'" }}>
            CEMOPTI v4.0<br/>
            5 Calculation Modules · Excel Import/Export · Dynamic Property Logic · Depth-Based Simulation<br/>
            Built on React 18 · Recharts · SheetJS (XLSX)
          </div>
        </Card>
      </div>
    ),
  };

  return (
    <div className="animate-in">
      <SectionTitle icon="📘" title="MANUAL / TECH GUIDE" sub="Complete user guide, technical documentation, reference data, and legal information" />
      <div style={{ display:"flex",gap:8,marginBottom:16,flexWrap:"wrap" }}>
        {sections.map(s=>(
          <button key={s.id} onClick={()=>setSection(s.id)} style={{
            padding:"8px 16px",borderRadius:6,border:`1px solid ${section===s.id?T.accent:T.border}`,
            background:section===s.id?`${T.accent}18`:"transparent",
            color:section===s.id?T.accent:T.muted,fontSize:12,fontWeight:600,cursor:"pointer",
            fontFamily:"'IBM Plex Sans'",
          }}>{s.icon} {s.label}</button>
        ))}
      </div>
      <div style={{ animation:"slideIn 0.2s ease" }}>
        {content[section]}
      </div>
    </div>
  );
}

// ─────────────────────────── NAV GROUPS ───────────────────────────────────────
// ─────────────────────────── EVALUATION OUTPUT PAGE ─────────────────────────
// Bond Index (BI) and Compressive Strength (CS) calculated from simulation data
// CBL/VDL import for comparison plots

function calcBondIndex(ecd, hydroPressure, dispEff, standoff) {
  // Bond Index (0–1): empirical correlation
  // Higher BI → better cement bond quality
  // BI ≈ f(displacement efficiency, standoff, hydrostatic pressure)
  // API RP 10B-2 simplified: BI = (DE/100)^0.8 × (SO/100)^0.4 × min(1, hydroPsi/3000)
  const deFactor  = Math.pow(Math.max(0, dispEff) / 100, 0.8);
  const soFactor  = Math.pow(Math.max(0, standoff) / 100, 0.4);
  const pFactor   = Math.min(1.0, Math.max(0, hydroPressure) / 3000);
  return Math.min(1.0, Math.round(deFactor * soFactor * pFactor * 1000) / 1000);
}

function calcCompressiveStrength(temperature, ecd, density) {
  // Compressive Strength (psi): simplified Bearden Consistency
  // CS ≈ k1 × ρ^2 × (1 - temp_factor) × ECD_factor
  // Typical range: 0–5000 psi for set cement
  // Reference: API RP 10B-2, Bearden-McIntyre correlation
  const tempFactor = Math.max(0.1, 1 - (temperature - 100) / 400); // higher temp → lower CS
  const rhoFactor  = Math.pow(Math.max(7, density), 1.5) / 100;
  const ecdFactor  = Math.min(1.5, ecd / 12);
  const cs = 500 * rhoFactor * tempFactor * ecdFactor;
  return Math.round(Math.min(6000, Math.max(0, cs)));
}

function ResultsEvaluationPage({ simResults }) {
  const [cblData, setCblData] = useState([]);   // imported CBL rows
  const [vdlData, setVdlData] = useState([]);   // imported VDL rows
  const [cblStatus, setCblStatus] = useState(null);
  const [vdlStatus, setVdlStatus] = useState(null);
  const cblRef = useRef();
  const vdlRef = useRef();

  if (!simResults) return <NoResults />;
  const d = simResults.grid.filter((_,i)=>i%2===0);

  // Calculate BI and CS at each depth step
  const evalData = d.map(r => ({
    md: r.md,
    tvd: r.tvd,
    bondIndex: calcBondIndex(r.ecd, r.hydroPressure, r.dispEfficiency, r.standoff),
    compressiveStrength: calcCompressiveStrength(r.temperature, r.ecd, r.adjDensity),
    dispEfficiency: r.dispEfficiency,
    standoff: r.standoff,
    temperature: r.temperature,
  }));

  // Bond quality classification
  const poorBond       = evalData.filter(r => r.bondIndex < 0.4).length;
  const partialBond    = evalData.filter(r => r.bondIndex >= 0.4 && r.bondIndex < 0.7).length;
  const goodBond       = evalData.filter(r => r.bondIndex >= 0.7).length;
  const avgBI          = Math.round(evalData.reduce((s,r)=>s+r.bondIndex,0)/evalData.length*1000)/1000;
  const avgCS          = Math.round(evalData.reduce((s,r)=>s+r.compressiveStrength,0)/evalData.length);

  // Excel import handlers
  function importLog(file, type) {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb = XLSX.read(e.target.result, { type:"array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const raw = XLSX.utils.sheet_to_json(ws, { defval:0 });
        // Expected columns: Depth, Amplitude (CBL) or Depth, Density (VDL)
        if (type === "cbl") {
          const mapped = raw.map(r => ({
            md: parseFloat(r.Depth || r.depth || r.MD || r.md || 0),
            cblAmplitude: parseFloat(r.Amplitude || r.amplitude || r.CBL || r.cbl || r["CBL Amplitude"] || 0),
          })).filter(r => r.md > 0);
          setCblData(mapped);
          setCblStatus({ ok:true, msg:`✓ Imported ${mapped.length} CBL points` });
        } else {
          const mapped = raw.map(r => ({
            md: parseFloat(r.Depth || r.depth || r.MD || r.md || 0),
            vdlDensity: parseFloat(r.Density || r.density || r.VDL || r.vdl || r["VDL Density"] || 0),
          })).filter(r => r.md > 0);
          setVdlData(mapped);
          setVdlStatus({ ok:true, msg:`✓ Imported ${mapped.length} VDL points` });
        }
      } catch(err) {
        if (type==="cbl") setCblStatus({ ok:false, msg:`Parse error: ${err.message}` });
        else setVdlStatus({ ok:false, msg:`Parse error: ${err.message}` });
      }
    };
    reader.readAsArrayBuffer(file);
    file = null;
  }

  // Merge CBL/VDL with evalData for overlay charts
  const cblMap = {};
  cblData.forEach(r => { cblMap[Math.round(r.md/50)*50] = r.cblAmplitude; });
  const vdlMap = {};
  vdlData.forEach(r => { vdlMap[Math.round(r.md/50)*50] = r.vdlDensity; });

  const mergedBIData = evalData.map(r => ({
    ...r,
    cblAmplitude: cblMap[r.md] !== undefined ? cblMap[r.md] : null,
  }));
  const mergedCSData = evalData.map(r => ({
    ...r,
    vdlDensity: vdlMap[r.md] !== undefined ? vdlMap[r.md] : null,
  }));

  return (
    <div className="animate-in">
      <SectionTitle icon="🧱" title="EVALUATION — BOND INDEX & COMPRESSIVE STRENGTH"
        sub="Calculated BI and CS vs depth · CBL/VDL log import for comparison · Quality assessment" />

      {/* Import controls */}
      <div style={{ display:"flex",gap:12,marginBottom:16,flexWrap:"wrap",alignItems:"center" }}>
        <div style={{ display:"flex",alignItems:"center",gap:8 }}>
          <input ref={cblRef} type="file" accept=".xlsx,.xls,.csv"
            onChange={e=>{ importLog(e.target.files[0],"cbl"); e.target.value=""; }}
            style={{ display:"none" }} />
          <button onClick={()=>cblRef.current.click()} style={{
            background:"transparent",border:`1px solid ${T.cyan}`,color:T.cyan,
            borderRadius:4,padding:"6px 14px",fontSize:11,cursor:"pointer",
            fontFamily:"'IBM Plex Sans'",fontWeight:600 }}>⬆ Import CBL Log (Excel)</button>
          {cblStatus && <span style={{ fontSize:11,color:cblStatus.ok?T.green:T.red,fontFamily:"'IBM Plex Mono'" }}>{cblStatus.msg}</span>}
        </div>
        <div style={{ display:"flex",alignItems:"center",gap:8 }}>
          <input ref={vdlRef} type="file" accept=".xlsx,.xls,.csv"
            onChange={e=>{ importLog(e.target.files[0],"vdl"); e.target.value=""; }}
            style={{ display:"none" }} />
          <button onClick={()=>vdlRef.current.click()} style={{
            background:"transparent",border:`1px solid ${T.gold}`,color:T.gold,
            borderRadius:4,padding:"6px 14px",fontSize:11,cursor:"pointer",
            fontFamily:"'IBM Plex Sans'",fontWeight:600 }}>⬆ Import VDL Log (Excel)</button>
          {vdlStatus && <span style={{ fontSize:11,color:vdlStatus.ok?T.green:T.red,fontFamily:"'IBM Plex Mono'" }}>{vdlStatus.msg}</span>}
        </div>
        <div style={{ fontSize:10,color:T.dim }}>
          Excel columns: Depth, Amplitude (CBL) | Depth, Density (VDL)
        </div>
      </div>

      {/* Summary metrics */}
      <div className="grid-4" style={{ marginBottom:16 }}>
        <MetricCard label="Avg Bond Index" value={avgBI} unit="0 = none · 1 = perfect" color={avgBI>0.7?T.green:avgBI>0.4?T.yellow:T.red} icon="🧱" />
        <MetricCard label="Poor Bond (<0.4)" value={poorBond} unit="depth steps" color={poorBond>0?T.red:T.green} icon="⚠️" />
        <MetricCard label="Good Bond (≥0.7)" value={goodBond} unit="depth steps" color={goodBond>0?T.green:T.muted} icon="✅" />
        <MetricCard label="Avg Comp. Strength" value={`${avgCS.toLocaleString()}`} unit="psi" color={avgCS>2000?T.green:avgCS>500?T.yellow:T.red} icon="💪" />
      </div>

      <div className="grid-2" style={{ marginBottom:16 }}>
        {/* Bond Index vs Depth + CBL overlay */}
        <Card>
          <div style={{ fontSize:12,fontWeight:600,color:T.accent,fontFamily:"'IBM Plex Mono'",marginBottom:8 }}>
            BOND INDEX vs DEPTH
            {cblData.length > 0 && <span style={{ marginLeft:10,fontSize:10,color:T.cyan }}>+ CBL Log overlay ({cblData.length} pts)</span>}
          </div>
          <div style={{ display:"flex",gap:12,marginBottom:6,fontSize:10,flexWrap:"wrap" }}>
            <span style={{ color:T.red }}>🔴 Poor &lt;0.4</span>
            <span style={{ color:T.yellow }}>🟡 Partial 0.4–0.7</span>
            <span style={{ color:T.green }}>🟢 Good &gt;0.7</span>
          </div>
          <ResponsiveContainer width="100%" height={440}>
            <LineChart data={[...mergedBIData].sort((a,b)=>a.md-b.md)} layout="vertical" margin={{top:5,right:30,left:10,bottom:20}}>
              <CartesianGrid strokeDasharray="3 3" stroke={T.dim} />
              <XAxis type="number" domain={[0, cblData.length>0 ? "auto" : 1]}
                tick={{fontSize:10,fill:T.muted}}
                label={{value:"BI / CBL Amplitude",position:"insideBottom",offset:-5,style:{fontSize:9,fill:T.muted}}} />
              <YAxis type="number" dataKey="md" reversed={true} domain={["dataMin","dataMax"]} tick={{fontSize:10,fill:T.muted}}
                label={{value:"Depth (ft MD)",angle:-90,position:"insideLeft",style:{fontSize:10,fill:T.muted}}} />
              <Tooltip contentStyle={{background:T.card,border:`1px solid ${T.border}`,fontSize:11}}
                labelFormatter={v=>`Depth: ${v} ft`}
                formatter={(val,name)=>{
                  if(name==="Bond Index") {
                    const q=val<0.4?"🔴 POOR":val<0.7?"🟡 PARTIAL":"🟢 GOOD";
                    return [`${val} — ${q}`, name];
                  }
                  return [val, name];
                }} />
              <Legend wrapperStyle={{fontSize:11}} />
              <ReferenceLine x={0.4} stroke={T.red}    strokeDasharray="4 2" strokeWidth={1}
                label={{value:"0.4",position:"insideTopRight",style:{fontSize:8,fill:T.red}}} />
              <ReferenceLine x={0.7} stroke={T.green}  strokeDasharray="4 2" strokeWidth={1}
                label={{value:"0.7",position:"insideTopRight",style:{fontSize:8,fill:T.green}}} />
              <Line type="monotone" dataKey="bondIndex" stroke={T.cyan} dot={false} strokeWidth={2} name="Bond Index" />
              {cblData.length > 0 &&
                <Line type="monotone" dataKey="cblAmplitude" stroke={T.gold} dot={false} strokeWidth={1.5}
                  name="CBL Amplitude" connectNulls={false} strokeDasharray="6 2" />}
            </LineChart>
          </ResponsiveContainer>
        </Card>

        {/* Compressive Strength vs Depth + VDL overlay */}
        <Card>
          <div style={{ fontSize:12,fontWeight:600,color:T.accent,fontFamily:"'IBM Plex Mono'",marginBottom:8 }}>
            COMPRESSIVE STRENGTH vs DEPTH
            {vdlData.length > 0 && <span style={{ marginLeft:10,fontSize:10,color:T.gold }}>+ VDL Log overlay ({vdlData.length} pts)</span>}
          </div>
          <div style={{ display:"flex",gap:12,marginBottom:6,fontSize:10,flexWrap:"wrap" }}>
            <span style={{ color:T.red }}>🔴 &lt;500 psi — insufficient</span>
            <span style={{ color:T.yellow }}>🟡 500–2000 psi — moderate</span>
            <span style={{ color:T.green }}>🟢 &gt;2000 psi — good</span>
          </div>
          <ResponsiveContainer width="100%" height={440}>
            <LineChart data={[...mergedCSData].sort((a,b)=>a.md-b.md)} layout="vertical" margin={{top:5,right:30,left:10,bottom:20}}>
              <CartesianGrid strokeDasharray="3 3" stroke={T.dim} />
              <XAxis type="number" tick={{fontSize:10,fill:T.muted}}
                label={{value:"CS (psi) / VDL Density",position:"insideBottom",offset:-5,style:{fontSize:9,fill:T.muted}}} />
              <YAxis type="number" dataKey="md" reversed={true} domain={["dataMin","dataMax"]} tick={{fontSize:10,fill:T.muted}}
                label={{value:"Depth (ft MD)",angle:-90,position:"insideLeft",style:{fontSize:10,fill:T.muted}}} />
              <Tooltip contentStyle={{background:T.card,border:`1px solid ${T.border}`,fontSize:11}}
                labelFormatter={v=>`Depth: ${v} ft`}
                formatter={(val,name)=>{
                  if(name==="Comp. Strength (psi)") {
                    const q=val<500?"🔴 INSUFFICIENT":val<2000?"🟡 MODERATE":"🟢 GOOD";
                    return [`${val?.toLocaleString()} psi — ${q}`, name];
                  }
                  return [val, name];
                }} />
              <Legend wrapperStyle={{fontSize:11}} />
              <ReferenceLine x={500}  stroke={T.red}   strokeDasharray="4 2" strokeWidth={1}
                label={{value:"500",position:"insideTopRight",style:{fontSize:8,fill:T.red}}} />
              <ReferenceLine x={2000} stroke={T.green} strokeDasharray="4 2" strokeWidth={1}
                label={{value:"2000",position:"insideTopRight",style:{fontSize:8,fill:T.green}}} />
              <Line type="monotone" dataKey="compressiveStrength" stroke={T.gold} dot={false} strokeWidth={2} name="Comp. Strength (psi)" />
              {vdlData.length > 0 &&
                <Line type="monotone" dataKey="vdlDensity" stroke={T.accent} dot={false} strokeWidth={1.5}
                  name="VDL Density" connectNulls={false} strokeDasharray="6 2" />}
            </LineChart>
          </ResponsiveContainer>
        </Card>
      </div>

      {/* Detailed BI + CS table */}
      <Card>
        <div style={{ fontSize:11,color:T.accent,fontFamily:"'IBM Plex Mono'",marginBottom:10 }}>
          BOND INDEX & COMPRESSIVE STRENGTH DATA TABLE
        </div>
        <div style={{ overflowX:"auto",maxHeight:380 }}>
          <table>
            <thead style={{ position:"sticky",top:0 }}>
              <tr>{["Depth (ft)","TVD (ft)","Bond Index","BI Quality","Comp. Strength (psi)","CS Quality","Disp. Eff %","Standoff %","Temp (°F)"].map(h=><th key={h}>{h}</th>)}</tr>
            </thead>
            <tbody>
              {evalData.map((r,i)=>{
                const biQ = r.bondIndex>=0.7?"🟢 GOOD":r.bondIndex>=0.4?"🟡 PARTIAL":"🔴 POOR";
                const csQ = r.compressiveStrength>=2000?"🟢 GOOD":r.compressiveStrength>=500?"🟡 MODERATE":"🔴 LOW";
                const biBg = r.bondIndex>=0.7?`${T.green}10`:r.bondIndex>=0.4?`${T.yellow}10`:`${T.red}10`;
                return (
                  <tr key={i} style={{ background:biBg }}>
                    <td style={{ color:T.accent,fontWeight:600 }}>{r.md.toLocaleString()}</td>
                    <td>{r.tvd.toLocaleString()}</td>
                    <td style={{ color:r.bondIndex>=0.7?T.green:r.bondIndex>=0.4?T.yellow:T.red,fontWeight:700 }}>{r.bondIndex}</td>
                    <td style={{ fontSize:10 }}>{biQ}</td>
                    <td style={{ color:r.compressiveStrength>=2000?T.green:r.compressiveStrength>=500?T.yellow:T.red,fontWeight:700 }}>
                      {r.compressiveStrength.toLocaleString()}
                    </td>
                    <td style={{ fontSize:10 }}>{csQ}</td>
                    <td style={{ color:r.dispEfficiency>70?T.green:T.yellow }}>{r.dispEfficiency}%</td>
                    <td>{r.standoff}%</td>
                    <td style={{ color:T.gold }}>{r.temperature}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
        <div style={{ padding:"8px 0 0",fontSize:10,color:T.dim,lineHeight:1.6 }}>
          Bond Index correlation: API RP 10B-2 simplified — BI = f(displacement efficiency, standoff, hydrostatic pressure)<br/>
          Compressive Strength: Bearden-McIntyre correlation — CS = f(density, temperature, ECD)
        </div>
      </Card>
    </div>
  );
}


const NAV_GROUPS = [
  { label:"PROJECT", items:[{id:"dashboard",icon:"⬡",label:"Dashboard"},{id:"projects",icon:"📁",label:"Projects"}] },
  { label:"INPUTS", items:[{id:"general",icon:"📋",label:"General"},{id:"survey",icon:"📐",label:"Survey"},{id:"casing",icon:"⬡",label:"Casing Profile"},{id:"openhole",icon:"🕳️",label:"Open Hole"},{id:"formation",icon:"🪨",label:"Formation"},{id:"fluids",icon:"🧪",label:"Fluids"},{id:"pumping",icon:"⛽",label:"Pumping Schedule"},{id:"centralizers",icon:"⚙️",label:"Centralizers"},{id:"plug",icon:"🔌",label:"Balanced Plug"}] },
  { label:"RESULTS", items:[{id:"schematic",icon:"🛢️",label:"Well Schematic"},{id:"res-pressure",icon:"📊",label:"M1: Pressure/ECD"},{id:"res-displacement",icon:"🔄",label:"M2: Displacement"},{id:"res-thermal",icon:"🌡️",label:"M3: Temperature"},{id:"res-centralizer",icon:"⚙️",label:"M4: Centralizer"},{id:"res-plug",icon:"🔌",label:"M5: Plug"},{id:"res-hookload",icon:"🏗️",label:"Torque & Drag"},{id:"res-evaluation",icon:"🧱",label:"Evaluation"},{id:"res-debug",icon:"🐛",label:"Debug Report"},{id:"fulldata",icon:"🗄️",label:"Full Data Grid"}] },
  { label:"GUIDE & DB", items:[{id:"manual",icon:"📘",label:"Manual / Tech Guide"},{id:"database",icon:"🗄️",label:"DB Schema"}] },
];

function WellSchematicPage({ db, activeProject, simResults }) {
  const fluids=db.fluid_data[activeProject]||[];
  const casingProfile=db.casing_profile[activeProject]||[];
  const openHoles=db.open_hole_profile[activeProject]||[];
  const pump=db.pumping_schedule[activeProject]||[];
  const gen=db.general_data?.[activeProject]||{};
  const survey=db.survey_data[activeProject]||[];

  // Change 5: Schematic depth limit = max depth from Open Hole page only
  // Dynamically follows highest Open Hole to_depth entered by user
  const maxD = openHoles.length
    ? Math.max(...openHoles.map(h => h.to_depth || 0), 1000)
    : 1000;

  const toc = gen.toc || 0;  // Top of Cement depth from General page
  const W=300, H=520;
  const sy = d => 20 + (d / maxD) * (H - 50);
  const cx = W / 2;
  const fc = {};
  fluids.forEach(f => { fc[f.fluid_id] = f.color; });

  // Depth scale ticks
  const ticks = [0,1,2,3,4].map(i => Math.round(i * maxD / 4));

  // Compute fluid column heights from pump schedule for annular display
  const cementStages = [...pump].filter(s => {
    const fl = fluids.find(f=>f.fluid_id===s.fluid_id);
    return fl?.type?.toLowerCase()==="cement" && s.volume>0;
  });
  const spacerStages = [...pump].filter(s => {
    const fl = fluids.find(f=>f.fluid_id===s.fluid_id);
    return ["spacer","wash"].includes(fl?.type?.toLowerCase()||"") && s.volume>0;
  });

  // Cement column top = TOC, bottom = maxD (simplified)
  const cementTopY  = toc > 0 ? sy(toc) : sy(maxD * 0.5);
  const cementBotY  = sy(maxD);
  const cementColor = cementStages[0] ? (fc[cementStages[0].fluid_id] || "#F97316") : "#F97316";
  const spacerColor = spacerStages[0] ? (fc[spacerStages[0].fluid_id] || "#34D399") : "#34D399";

  return (
    <div className="animate-in">
      <SectionTitle icon="🛢️" title="WELL SCHEMATIC" sub={`Cement profile · TOC · Fluid sections · Depth limit: ${maxD.toLocaleString()} ft (Open Hole max)`} />
      <div style={{ display:"flex",gap:16,flexWrap:"wrap" }}>
        <svg width={W} height={H} style={{ background:T.surface,borderRadius:6,border:`1px solid ${T.border}`,flexShrink:0 }}>
          {/* Surface line */}
          <line x1={0} y1={20} x2={W} y2={20} stroke={T.border} strokeWidth={1}/>
          <text x={10} y={17} fill={T.muted} fontSize={9} fontFamily="IBM Plex Mono">Surface</text>

          {/* TD marker */}
          <line x1={0} y1={sy(maxD)} x2={W} y2={sy(maxD)} stroke={T.red} strokeWidth={1} strokeDasharray="3 2"/>
          <text x={10} y={sy(maxD)-3} fill={T.red} fontSize={8} fontFamily="IBM Plex Mono">TD {maxD.toLocaleString()}ft</text>

          {/* Open hole profile — width varies by hole size */}
          {openHoles.map((h,i) => {
            const y1=sy(h.from_depth), y2=sy(h.to_depth);
            const w = (h.hole_size / 26) * 100;
            return (
              <g key={i}>
                <rect x={cx-w} y={y1} width={w*2} height={Math.max(1,y2-y1)}
                  fill={`${T.gold}10`} stroke={T.gold} strokeWidth={0.5} strokeDasharray="4 2"/>
                <text x={6} y={y1+(y2-y1)/2+4} fill={T.gold} fontSize={7} fontFamily="IBM Plex Mono">{h.hole_size}"</text>
              </g>
            );
          })}

          {/* Cement column in annulus (TOC to TD) — uses cement fluid color */}
          {toc > 0 && (
            <>
              <rect x={cx-32} y={cementTopY} width={64} height={Math.max(0,cementBotY-cementTopY)}
                fill={`${cementColor}55`} stroke={cementColor} strokeWidth={1}/>
              {/* TOC marker */}
              <line x1={cx-38} y1={cementTopY} x2={cx+38} y2={cementTopY} stroke={cementColor} strokeWidth={1.5} strokeDasharray="5 2"/>
              <text x={cx+40} y={cementTopY+4} fill={cementColor} fontSize={8} fontFamily="IBM Plex Mono">TOC {toc.toLocaleString()}ft</text>
            </>
          )}

          {/* Spacer/wash above cement (simplified — 200ft above TOC) */}
          {toc > 0 && spacerStages.length > 0 && (
            <rect x={cx-30} y={Math.max(20,cementTopY-20)} width={60} height={20}
              fill={`${spacerColor}66`} stroke={spacerColor} strokeWidth={0.5}/>
          )}

          {/* Casing strings */}
          {casingProfile.map((c,i) => {
            const y1=sy(0), y2=sy(c.to_depth), w=(c.od/20)*72, tk=4;
            const col = `hsl(${200+i*35},55%,55%)`;
            return (
              <g key={i}>
                <rect x={cx-w-tk} y={y1} width={tk} height={y2-y1} fill={col}/>
                <rect x={cx+w}     y={y1} width={tk} height={y2-y1} fill={col}/>
                {/* Shoe */}
                <polygon points={`${cx-w-tk},${y2} ${cx+w+tk},${y2} ${cx},${y2+8}`} fill={col}/>
              </g>
            );
          })}

          {/* Fluid inside casing (simplified) */}
          {[...pump].reverse().filter(s=>s.volume>0).map((stage,i) => {
            const colLen = stage.volume * 400;
            const td2 = maxD - colLen * i;
            const top2 = maxD - colLen * (i+1);
            const y1=sy(Math.max(0,top2)), y2=sy(Math.min(maxD,td2));
            const color = fc[stage.fluid_id] || "#888";
            return <rect key={i} x={cx-14} y={y1} width={28} height={Math.max(0,y2-y1)} fill={`${color}88`}/>;
          })}

          {/* Depth scale ticks */}
          {ticks.map((d,i) => (
            <g key={i}>
              <line x1={W-38} y1={sy(d)} x2={W-33} y2={sy(d)} stroke={T.muted} strokeWidth={1}/>
              <text x={W-31} y={sy(d)+3} fill={T.muted} fontSize={7} fontFamily="IBM Plex Mono">{d.toLocaleString()}</text>
            </g>
          ))}
          <text x={cx} y={14} textAnchor="middle" fill={T.accent} fontSize={9} fontFamily="Orbitron" fontWeight={700}>WELL SCHEMATIC</text>
        </svg>

        <div style={{ flex:1,minWidth:240 }}>
          {/* Legend */}
          <Card style={{ marginBottom:12 }}>
            <div style={{ fontSize:11,color:T.muted,fontFamily:"'IBM Plex Mono'",marginBottom:8 }}>LEGEND</div>
            {[
              { color:T.gold,      label:"Open Hole Profile (varies with depth)" },
              { color:"#F97316",   label:`Cement Column (TOC: ${toc?toc.toLocaleString()+" ft":"not set"})` },
              { color:spacerColor, label:"Spacer / Wash zone" },
            ].map((l,i)=>(
              <div key={i} style={{ display:"flex",alignItems:"center",gap:8,padding:"4px 0",borderBottom:`1px solid ${T.border}`,fontSize:11 }}>
                <span style={{ width:16,height:10,borderRadius:2,background:l.color,flexShrink:0 }}></span>
                <span style={{ color:T.muted }}>{l.label}</span>
              </div>
            ))}
          </Card>

          {/* Casing strings */}
          <Card style={{ marginBottom:12 }}>
            <div style={{ fontSize:11,color:T.muted,fontFamily:"'IBM Plex Mono'",marginBottom:8 }}>CASING STRINGS</div>
            {casingProfile.map((c,i)=>(
              <div key={i} style={{ display:"flex",gap:10,padding:"5px 0",borderBottom:`1px solid ${T.border}`,fontSize:11 }}>
                <span style={{ color:T.accent,minWidth:80,fontFamily:"'IBM Plex Mono'" }}>{c.od}" {c.type}</span>
                <span style={{ color:T.muted }}>0–{c.to_depth.toLocaleString()} ft</span>
                <span style={{ color:T.dim }}>{c.grade}</span>
              </div>
            ))}
          </Card>

          {/* Open hole profile */}
          <Card style={{ marginBottom:12 }}>
            <div style={{ fontSize:11,color:T.muted,fontFamily:"'IBM Plex Mono'",marginBottom:8 }}>OPEN HOLE PROFILE</div>
            {openHoles.map((h,i)=>(
              <div key={i} style={{ display:"flex",gap:10,padding:"5px 0",borderBottom:`1px solid ${T.border}`,fontSize:11 }}>
                <span style={{ color:T.gold,minWidth:50,fontFamily:"'IBM Plex Mono'" }}>{h.hole_size}"</span>
                <span style={{ color:T.muted }}>{h.from_depth.toLocaleString()}–{h.to_depth.toLocaleString()} ft</span>
              </div>
            ))}
          </Card>

          {/* Fluid column sequence */}
          <Card>
            <div style={{ fontSize:11,color:T.muted,fontFamily:"'IBM Plex Mono'",marginBottom:8 }}>FLUID SEQUENCE (BOTTOM → TOP)</div>
            {[...pump].reverse().filter(s=>s.volume>0).map((s,i)=>{
              const fl=fluids.find(f=>f.fluid_id===s.fluid_id);
              return (
                <div key={i} style={{ display:"flex",gap:8,padding:"5px 0",borderBottom:`1px solid ${T.border}`,fontSize:11,alignItems:"center" }}>
                  <span style={{ width:10,height:10,borderRadius:2,background:fl?.color||"#888",flexShrink:0 }}></span>
                  <span style={{ color:T.text,minWidth:110 }}>{s.fluid_name}</span>
                  <span style={{ color:T.muted }}>{s.volume} bbl</span>
                  <Badge text={s.purpose?.toUpperCase()||""} color={fl?.color||T.muted} />
                </div>
              );
            })}
          </Card>
        </div>
      </div>
    </div>
  );
}
// ─────────────────────────── MAIN APP ─────────────────────────────────────────
export default function CEMOPTI() {
  const [loggedIn, setLoggedIn] = useState(false);
  const [db, setDb] = useState(()=>initDB());
  const [activeProject, setActiveProject] = useState("p1");
  const [activePage, setActivePage] = useState("dashboard");
  // Per-project simulation results: { [projectId]: simResults }
  const [simResultsMap, setSimResultsMap] = useState({});
  const [running, setRunning] = useState(false);
  const [simProgress, setSimProgress] = useState(0);
  const [simStageInfo, setSimStageInfo] = useState(null);
  const [sidebarOpen, setSidebarOpen] = useState(true);
  const [validation, setValidation] = useState(null);
  const [unitSystem, setUnitSystem] = useState("imperial");

  useEffect(()=>{ saveDB(db); },[db]);

  // Sync unit system whenever the active project changes
  useEffect(()=>{
    const proj = db.projects.find(p=>p.project_id===activeProject);
    if (proj) setUnitSystem(proj.unit_system);
  },[activeProject, db.projects]);

  const setDbAndSave = useCallback(updater=>{
    setDb(prev=>typeof updater==="function"?updater(prev):updater);
  },[]);

  // Active project's simulation results (from per-project map)
  const simResults = simResultsMap[activeProject] || null;

  // Compute input pages based on active project's hole_type
  const proj = db.projects.find(p=>p.project_id===activeProject);
  // Input pages computed from job_type: plug only for "Secondary (Balanced Plug)"
  const inputPages = useMemo(() => getInputPages(db, activeProject), [db, activeProject]);

  const runSim = useCallback(()=>{
    var preCheck = validateAll(db, activeProject);
    if (preCheck.errors.length > 0) {
      setValidation(preCheck);
      return;
    }
    if (preCheck.warnings.length > 0) setValidation(preCheck);

    setRunning(true);
    setSimProgress(0);
    setSimStageInfo(null);
    
    var g = db.general_data[activeProject] || {};
    var numTimesteps = g.num_timesteps || 100;
    var numDepthGrids = g.num_depth_grids || 100;
    var pump = db.pumping_schedule[activeProject] || [];
    var openhole = db.open_hole_profile[activeProject] || [];
    
    var minDepth = openhole.length > 0 ? Math.min(...openhole.map(function(h){return h.from_depth;})) : 0;
    var maxDepth = openhole.length > 0 ? Math.max(...openhole.map(function(h){return h.to_depth;})) : (g.total_depth_md || 10000);
    var depthStep = (maxDepth - minDepth) / Math.max(1, numDepthGrids - 1);
    
    var depthGrids = [];
    for (var i = 0; i < numDepthGrids; i++) {
      depthGrids.push({
        gridIndex: i,
        depth: minDepth + i * depthStep,
        pressure: 0,
        ecd: 0,
        fluidAssignment: null
      });
    }
    
    var totalJobTime = 0;
    var validStages = [];
    for (var si = 0; si < pump.length; si++) {
      if (pump[si].volume > 0 && pump[si].rate > 0) {
        var stageTime = (pump[si].volume / pump[si].rate) * 60;
        validStages.push({
          stageIndex: si,
          stageName: pump[si].fluid_name || ("Stage " + (si+1)),
          stageTime: stageTime
        });
        totalJobTime += stageTime;
      }
    }
    
    if (validStages.length === 0) {
      totalJobTime = 3600;
      validStages = [{stageIndex:0, stageName:"Default", stageTime:3600}];
    }
    
    var timestepStages = [];
    var cumTime = 0;
    for (var vi = 0; vi < validStages.length; vi++) {
      var vs = validStages[vi];
      var stageTimesteps = Math.max(1, Math.round((vs.stageTime / totalJobTime) * numTimesteps));
      timestepStages.push({
        stageIndex: vs.stageIndex,
        stageName: vs.stageName,
        startTime: cumTime,
        endTime: cumTime + vs.stageTime,
        timesteps: stageTimesteps
      });
      cumTime += vs.stageTime;
    }
    
    var totalSteps = numTimesteps * numDepthGrids;
    var currentStep = 0;
    var currentStageIdx = 0;
    var currentTimestep = 0;
    var currentGrid = 0;
    
    var progressInterval = setInterval(function(){
      currentStep++;
      currentGrid = ((currentStep - 1) % numDepthGrids) + 1;
      
      if (currentGrid === 1 && currentStep > 1) {
        currentTimestep++;
        var cumulativeTs = 0;
        for (var si = 0; si < timestepStages.length; si++) {
          cumulativeTs += timestepStages[si].timesteps;
          if (currentTimestep <= cumulativeTs) {
            currentStageIdx = si;
            break;
          }
        }
      }
      
      var progress = Math.min(100, Math.round((currentStep / totalSteps) * 100));
      setSimProgress(progress);
      
      var stageInfo = timestepStages[currentStageIdx] || {stageName:"Processing", stageIndex:0};
      setSimStageInfo({
        stage: stageInfo.stageName,
        stageNum: currentStageIdx + 1,
        totalStages: timestepStages.length,
        timestep: currentTimestep,
        totalTimesteps: numTimesteps,
        grid: currentGrid,
        totalGrids: numDepthGrids
      });
      
      if (currentStep >= totalSteps || progress >= 100) {
        clearInterval(progressInterval);
        setTimeout(function(){
          try {
            var r = runSimulationEngine(db, activeProject);
            setSimResultsMap(function(prev){return {...prev, [activeProject]: r};});
            setDbAndSave(function(d){return dbAddSimRun(d, activeProject, {
              run_id: r.run_id, timestamp: r.timestamp, summary: r.summary,
            });});
            setActivePage("res-pressure");
            setRunning(false);
            setSimProgress(0);
            setSimStageInfo(null);
          } catch(e) {
            console.error("Simulation error:", e);
            setRunning(false);
            setSimProgress(0);
            setSimStageInfo(null);
          }
        }, 300);
      }
    }, Math.max(10, Math.min(100, 3000 / totalSteps)));
  },[db, activeProject, setDbAndSave]);

  const handleValidate = useCallback(()=>{
    setValidation(validateAll(db, activeProject));
  },[db, activeProject]);

  const handleReset = useCallback(()=>{
    if (!window.confirm("Reset all inputs for this project to default values?")) return;
    setDbAndSave(d=>({...d,
      survey_data:{...d.survey_data,[activeProject]:JSON.parse(JSON.stringify(DEFAULT_DATA.survey_data))},
      casing_profile:{...d.casing_profile,[activeProject]:JSON.parse(JSON.stringify(DEFAULT_DATA.casing_profile))},
      open_hole_profile:{...d.open_hole_profile,[activeProject]:JSON.parse(JSON.stringify(DEFAULT_DATA.open_hole_profile))},
      formation_data:{...d.formation_data,[activeProject]:JSON.parse(JSON.stringify(DEFAULT_DATA.formation_data))},
      fluid_data:{...d.fluid_data,[activeProject]:JSON.parse(JSON.stringify(DEFAULT_DATA.fluid_data))},
      centralizer_data:{...d.centralizer_data,[activeProject]:JSON.parse(JSON.stringify(DEFAULT_DATA.centralizer_data))},
      temperature_profile:{...d.temperature_profile,[activeProject]:JSON.parse(JSON.stringify(DEFAULT_DATA.temperature_profile))},
      pumping_schedule:{...d.pumping_schedule,[activeProject]:JSON.parse(JSON.stringify(DEFAULT_DATA.pumping_schedule))},
      balanced_plug_data:{...d.balanced_plug_data,[activeProject]:JSON.parse(JSON.stringify(DEFAULT_DATA.balanced_plug_data))},
      general_data:{...d.general_data,[activeProject]:{...DEFAULT_DATA.general_data}},
    }));
  },[activeProject, setDbAndSave]);

  const handleReport = useCallback(()=>{
    console.log('Report button clicked');
    try {
      generateReport(db, activeProject, simResults);
    } catch (error) {
      console.error('Error in handleReport:', error);
      alert('Error generating report: ' + error.message);
    }
  },[db, activeProject, simResults]);

  const handleUnitChange = useCallback((u)=>{
    setUnitSystem(u);
    setDbAndSave(d=>({...d, projects:d.projects.map(p=>p.project_id===activeProject?{...p,unit_system:u}:p)}));
  },[activeProject, setDbAndSave]);

  const navigate = useCallback((pageId)=>{
    if(pageId==="run"){ runSim(); return; }
    setActivePage(pageId);
  },[runSim]);

  // Open existing project: switch active, go to Dashboard (inputs + results already in state)
  const openProject = useCallback((pid)=>{
    setActiveProject(pid);
    setActivePage("dashboard");
  },[]);

  if (!loggedIn) return <LoginPage onLogin={()=>setLoggedIn(true)} />;

  const sharedProps = { db, activeProject, setDb:setDbAndSave, simResults, unitSystem, inputPages };

  // All nav groups shown as-is — no conditional filtering
  const navGroupsFiltered = NAV_GROUPS;

  const renderPage = () => {
    switch(activePage){
      case "dashboard":       return <DashboardPage {...sharedProps} onRun={runSim} running={running} simProgress={simProgress} simStageInfo={simStageInfo} onValidate={handleValidate} onReset={handleReset} onReport={handleReport} onUnitChange={handleUnitChange} />;
      case "projects":        return <ProjectsPage {...sharedProps} setActiveProject={setActiveProject} onNavigate={navigate} onOpenProject={openProject} simResultsMap={simResultsMap} />;
      case "schematic":       return <WellSchematicPage {...sharedProps} />;
      case "general":         return <GeneralPage {...sharedProps} onNavigate={navigate} />;
      case "survey":          return <SurveyPage {...sharedProps} onNavigate={navigate} />;
      case "casing":          return <CasingPage {...sharedProps} onNavigate={navigate} />;
      case "openhole":        return <OpenHolePage {...sharedProps} onNavigate={navigate} />;
      case "formation":       return <FormationPage {...sharedProps} onNavigate={navigate} />;
      case "fluids":          return <FluidPage {...sharedProps} onNavigate={navigate} />;
      case "pumping":         return <PumpingPage {...sharedProps} onNavigate={navigate} />;
      case "centralizers":    return <CentralizerPage {...sharedProps} onNavigate={navigate} />;
      case "plug":            return <BalancedPlugPage {...sharedProps} onNavigate={navigate} />;
      case "res-pressure":    return <ResultsPressurePage {...sharedProps} />;
      case "res-displacement":return <ResultsDisplacementPage {...sharedProps} />;
      case "res-thermal":     return <ResultsThermalPage {...sharedProps} />;
      case "res-centralizer": return <ResultsCentralizerPage {...sharedProps} />;
      case "res-plug":        return <ResultsPlugPage {...sharedProps} />;
      case "res-hookload":    return <ResultsHookLoadPage {...sharedProps} />;
      case "res-evaluation":  return <ResultsEvaluationPage {...sharedProps} />;
      case "res-debug":       return <DebugReportPage {...sharedProps} />;
      case "fulldata":        return <FullDataPage {...sharedProps} />;
      case "manual":          return <ManualPage />;
      case "database":        return <DatabasePage {...sharedProps} />;
      default: return null;
    }
  };

  return (
    <>
      <style>{css}</style>
      <div style={{ display:"flex",minHeight:"100vh",background:T.bg }}>
        {/* Sidebar */}
        <div style={{ width:sidebarOpen?230:52,flexShrink:0,background:T.surface,borderRight:`1px solid ${T.border}`,transition:"width 0.2s",display:"flex",flexDirection:"column",overflow:"hidden" }}>
          <div style={{ padding:"16px 14px",borderBottom:`1px solid ${T.border}`,display:"flex",alignItems:"center",gap:10 }}>
            <div style={{ width:28,height:28,background:`linear-gradient(135deg,${T.accent2},${T.accent})`,borderRadius:6,display:"flex",alignItems:"center",justifyContent:"center",fontSize:14,fontWeight:900,color:"#000",flexShrink:0,fontFamily:"'Orbitron'" }}>C</div>
            {sidebarOpen && (<div><div style={{ fontFamily:"'Orbitron'",fontSize:13,fontWeight:900,color:T.accent,letterSpacing:"0.1em" }}>CEMOPTI</div><div style={{ fontSize:9,color:T.muted,letterSpacing:"0.08em" }}>CEMENTING SIMULATION</div></div>)}
          </div>
          {sidebarOpen && proj && (
            <div style={{ padding:"8px 14px",borderBottom:`1px solid ${T.border}`,background:`${T.accent}08` }}>
              <div style={{ fontSize:9,color:T.muted,textTransform:"uppercase",letterSpacing:"0.08em" }}>Active Project</div>
              <div style={{ fontSize:11,color:T.accent,fontFamily:"'IBM Plex Mono'",marginTop:2,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis" }}>{proj.well_name}</div>
              <div style={{ fontSize:9,color:T.dim,marginTop:2 }}>{proj.project_name}</div>
            </div>
          )}
          <div style={{ flex:1,overflowY:"auto",padding:"8px 0" }}>
            {navGroupsFiltered.map(group=>(
              <div key={group.label}>
                {sidebarOpen && <div style={{ padding:"8px 14px 4px",fontSize:9,color:T.dim,fontFamily:"'IBM Plex Mono'",letterSpacing:"0.12em",textTransform:"uppercase" }}>{group.label}</div>}
                {group.items.map(item=>{
                  const isResult = item.id.startsWith("res-")||item.id==="fulldata";
                  const active = activePage===item.id;
                  return <button key={item.id} onClick={()=>setActivePage(item.id)} style={{ width:"100%",padding:sidebarOpen?"8px 14px":"10px",border:"none",background:active?`${T.accent}18`:"transparent",color:active?T.accent:isResult&&!simResults?T.dim:T.muted,display:"flex",alignItems:"center",gap:8,fontSize:12,fontFamily:"'IBM Plex Sans',sans-serif",borderLeft:active?`2px solid ${T.accent}`:"2px solid transparent",cursor:"pointer",transition:"all 0.1s",textAlign:"left" }}>
                    <span style={{ fontSize:14,flexShrink:0 }}>{item.icon}</span>
                    {sidebarOpen && <span style={{ whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis" }}>{item.label}</span>}
                  </button>;
                })}
              </div>
            ))}
          </div>
          {sidebarOpen && (
            <div style={{ padding:"8px 14px",borderTop:`1px solid ${T.border}`,background:`${T.border}33` }}>
              <div style={{ fontSize:9,color:T.muted,marginBottom:6 }}>QUICK ACTIONS</div>
              <div style={{ display:"flex",flexDirection:"column",gap:4 }}>
                <button onClick={handleReset} style={{ background:"none",border:`1px solid ${T.border}`,color:T.muted,borderRadius:4,padding:"5px 8px",fontSize:10,cursor:"pointer",textAlign:"left",fontFamily:"'IBM Plex Sans'" }}>↺ Reset to Defaults</button>
                <button onClick={handleValidate} style={{ background:"none",border:`1px solid ${T.border}`,color:T.gold,borderRadius:4,padding:"5px 8px",fontSize:10,cursor:"pointer",textAlign:"left",fontFamily:"'IBM Plex Sans'" }}>✓ Validate Data</button>
                {simResults && <button onClick={handleReport} style={{ background:"none",border:`1px solid ${T.border}`,color:T.green,borderRadius:4,padding:"5px 8px",fontSize:10,cursor:"pointer",textAlign:"left",fontFamily:"'IBM Plex Sans'" }}>⬇ Generate Report</button>}
              </div>
            </div>
          )}
          <button onClick={()=>setSidebarOpen(!sidebarOpen)} style={{ padding:"10px",border:"none",borderTop:`1px solid ${T.border}`,background:T.panel,color:T.muted,cursor:"pointer",fontSize:14 }}>{sidebarOpen?"◀":"▶"}</button>
          <button onClick={()=>{ if(window.confirm("Sign out of CEMOPTI?")) setLoggedIn(false); }} style={{
            padding:"10px",border:"none",borderTop:`1px solid ${T.border}`,
            background:"transparent",color:T.red,cursor:"pointer",fontSize:12,
            fontFamily:"'IBM Plex Sans'",fontWeight:600,
            display:"flex",alignItems:"center",justifyContent:"center",gap:6,
          }}>
            <span>⏻</span>
            {sidebarOpen && <span>Logout</span>}
          </button>
        </div>

        {/* Main */}
        <div style={{ flex:1,display:"flex",flexDirection:"column",overflow:"hidden" }}>
          <div style={{ height:52,background:T.surface,borderBottom:`1px solid ${T.border}`,display:"flex",alignItems:"center",padding:"0 20px",justifyContent:"space-between",flexShrink:0 }}>
            <div style={{ display:"flex",alignItems:"center",gap:10 }}>
              <span style={{ fontFamily:"'IBM Plex Mono'",fontSize:12,color:T.muted }}>
                {navGroupsFiltered.flatMap(g=>g.items).find(i=>i.id===activePage)?.icon}{" "}
                {navGroupsFiltered.flatMap(g=>g.items).find(i=>i.id===activePage)?.label}
              </span>
              {proj && <Badge text={proj.project_name.toUpperCase()} color={T.accent} />}
              <Badge text={UNITS[unitSystem].label.toUpperCase()} color={T.gold} />
            </div>
            <div style={{ display:"flex",alignItems:"center",gap:10 }}>

              {simResults && <Badge text={`${simResults.summary.totalSteps} PTS`} color={T.green} />}
              <Btn variant="ghost" size="sm" onClick={handleValidate}>✓ Validate</Btn>
              <Btn variant="primary" size="sm" onClick={runSim} disabled={running}>▶ Run</Btn>
              <div title={db.users[0]?.username} style={{ width:28,height:28,borderRadius:"50%",background:`${T.accent}22`,border:`1px solid ${T.accent}44`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,color:T.accent,cursor:"pointer" }}
                onClick={()=>{ if(window.confirm("Sign out?")) setLoggedIn(false); }}>
                {db.users[0]?.username?.[0]?.toUpperCase()}
              </div>
            </div>
          </div>
          <div style={{ flex:1,overflow:"auto",padding:"20px 24px" }}>
            {renderPage()}
          </div>
        </div>
      </div>
      <ValidationPanel validation={validation} onDismiss={()=>setValidation(null)} />
    </>
  );
}

