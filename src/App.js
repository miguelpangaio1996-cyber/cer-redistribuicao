import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import { initializeApp } from "firebase/app";
import { getFirestore, collection, doc, setDoc, deleteDoc, onSnapshot, getDoc } from "firebase/firestore";

// ─── Firebase Config ──────────────────────────────────────────────────────────
const firebaseConfig = {
  apiKey: "AIzaSyC7LTXETEQ0dpV1GqaTmY5b6k19ez-VY7o",
  authDomain: "cer-redistribuicao.firebaseapp.com",
  projectId: "cer-redistribuicao",
  storageBucket: "cer-redistribuicao.firebasestorage.app",
  messagingSenderId: "1090241154765",
  appId: "1:1090241154765:web:d54f16d29ecb667901e384"
};

const app = initializeApp(firebaseConfig);
const db = getFirestore(app);

// ─── Helpers ──────────────────────────────────────────────────────────────────
function uid() { return Date.now().toString(36) + Math.random().toString(36).slice(2); }

function haversine(lat1, lon1, lat2, lon2) {
  const R = 6371;
  const dLat = (lat2 - lat1) * Math.PI / 180;
  const dLon = (lon2 - lon1) * Math.PI / 180;
  const a = Math.sin(dLat/2)**2 + Math.cos(lat1*Math.PI/180)*Math.cos(lat2*Math.PI/180)*Math.sin(dLon/2)**2;
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
}

// ─── Point in Polygon (Ray Casting) ──────────────────────────────────────────
function pointInPolygon(lat, lon, polygonCoords) {
  // polygonCoords é array de [lon, lat] (formato GeoJSON)
  let inside = false;
  const x = lon, y = lat;
  for (let i = 0, j = polygonCoords.length - 1; i < polygonCoords.length; j = i++) {
    const xi = polygonCoords[i][0], yi = polygonCoords[i][1];
    const xj = polygonCoords[j][0], yj = polygonCoords[j][1];
    const intersect = ((yi > y) !== (yj > y)) && (x < (xj - xi) * (y - yi) / (yj - yi) + xi);
    if (intersect) inside = !inside;
  }
  return inside;
}

function pointInFeature(lat, lon, feature) {
  if (!feature) return true; // sem limite definido, aceita tudo
  const geom = feature.geometry;
  if (!geom) return true;
  try {
    if (geom.type === "Polygon") {
      return pointInPolygon(lat, lon, geom.coordinates[0]);
    } else if (geom.type === "MultiPolygon") {
      return geom.coordinates.some(poly => pointInPolygon(lat, lon, poly[0]));
    }
  } catch { return true; }
  return true;
}

// ─── Estilos ──────────────────────────────────────────────────────────────────
const S = {
  app: { minHeight: "100vh", background: "#f0f7f4", color: "#2d4a3e", fontFamily: "'Georgia', 'Times New Roman', serif", display: "flex", flexDirection: "column" },
  header: { background: "linear-gradient(135deg, #2d6a4f 0%, #40916c 100%)", padding: "20px 32px", display: "flex", alignItems: "center", gap: 16, boxShadow: "0 2px 12px rgba(45,106,79,0.15)" },
  logo: { width: 44, height: 44, background: "rgba(255,255,255,0.2)", borderRadius: "50%", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 22 },
  title: { fontSize: 20, fontWeight: 700, color: "#ffffff", margin: 0, letterSpacing: 0.5 },
  subtitle: { fontSize: 12, color: "rgba(255,255,255,0.75)", marginTop: 2, fontStyle: "italic" },
  nav: { display: "flex", background: "#ffffff", padding: "0 32px", gap: 0, boxShadow: "0 1px 4px rgba(0,0,0,0.06)", borderBottom: "1px solid #d8ede6" },
  navBtn: (active) => ({ padding: "16px 22px", fontSize: 13, fontWeight: active ? 700 : 500, fontFamily: "inherit", background: "none", border: "none", borderBottom: active ? "3px solid #2d6a4f" : "3px solid transparent", color: active ? "#2d6a4f" : "#7a9e8e", cursor: "pointer", transition: "all 0.2s" }),
  main: { flex: 1, padding: 32, maxWidth: 1100, margin: "0 auto", width: "100%" },
  card: { background: "#ffffff", border: "1px solid #d8ede6", borderRadius: 14, padding: 28, marginBottom: 20, boxShadow: "0 2px 8px rgba(45,106,79,0.06)" },
  cardTitle: { fontSize: 14, color: "#2d6a4f", marginBottom: 20, fontWeight: 700 },
  grid2: { display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 },
  grid3: { display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 16 },
  field: { display: "flex", flexDirection: "column", gap: 6 },
  label: { fontSize: 11, color: "#7a9e8e", fontWeight: 600, textTransform: "uppercase", letterSpacing: 1 },
  input: { background: "#f8fdfb", border: "1.5px solid #b7ddd0", borderRadius: 8, padding: "10px 14px", color: "#2d4a3e", fontFamily: "inherit", fontSize: 13, outline: "none" },
  select: { background: "#f8fdfb", border: "1.5px solid #b7ddd0", borderRadius: 8, padding: "10px 14px", color: "#2d4a3e", fontFamily: "inherit", fontSize: 13, outline: "none" },
  btn: (variant = "primary") => ({ padding: "10px 22px", borderRadius: 8, border: "none", fontFamily: "inherit", fontSize: 13, fontWeight: 600, cursor: "pointer", background: variant === "primary" ? "#2d6a4f" : variant === "danger" ? "#e05c5c" : variant === "green" ? "#40916c" : "#e8f4ef", color: variant === "primary" ? "#ffffff" : variant === "danger" ? "#ffffff" : variant === "green" ? "#ffffff" : "#2d6a4f", transition: "opacity 0.2s", boxShadow: variant === "primary" || variant === "green" ? "0 2px 6px rgba(45,106,79,0.2)" : "none" }),
  table: { width: "100%", borderCollapse: "collapse", fontSize: 13 },
  th: { padding: "11px 14px", textAlign: "left", fontSize: 11, color: "#7a9e8e", textTransform: "uppercase", letterSpacing: 1, borderBottom: "2px solid #d8ede6", fontWeight: 600 },
  td: { padding: "13px 14px", borderBottom: "1px solid #eaf5f0", color: "#3d5a4e" },
  badge: (color) => ({ display: "inline-block", padding: "4px 10px", borderRadius: 20, fontSize: 11, fontWeight: 600, background: color === "green" ? "#d8f3e8" : color === "orange" ? "#fef3dc" : color === "red" ? "#fde8e8" : "#dceefb", color: color === "green" ? "#1e7a4a" : color === "orange" ? "#b07c10" : color === "red" ? "#c0392b" : "#2563a8" }),
  empty: { textAlign: "center", padding: "52px 0", color: "#b7ddd0", fontSize: 15 },
  alert: (t) => ({ padding: "13px 18px", borderRadius: 10, fontSize: 13, marginBottom: 16, background: t === "ok" ? "#d8f3e8" : t === "err" ? "#fde8e8" : t === "warn" ? "#fef3dc" : "#dceefb", color: t === "ok" ? "#1e7a4a" : t === "err" ? "#c0392b" : t === "warn" ? "#b07c10" : "#2563a8", border: `1px solid ${t === "ok" ? "#a8dfc0" : t === "err" ? "#f0b0b0" : t === "warn" ? "#f0d080" : "#a8c8f0"}` }),
  divider: { borderColor: "#d8ede6", margin: "20px 0" },
  statBox: { background: "linear-gradient(135deg, #f0faf5 0%, #e8f5ef 100%)", border: "1px solid #c8e6d8", borderRadius: 12, padding: "20px 16px", textAlign: "center" },
  statNum: { fontSize: 30, fontWeight: 700, color: "#2d6a4f" },
  statLabel: { fontSize: 11, color: "#7a9e8e", textTransform: "uppercase", letterSpacing: 1, marginTop: 4 },
};

// ─── Mapa Leaflet ─────────────────────────────────────────────────────────────
function LeafletMap({ pontos, mostrarRaios, config, polygonFeature, linhasRedistribuicao }) {
  const mapRef = useRef(null);
  const mapInstanceRef = useRef(null);
  const layersRef = useRef([]);
  const initializedRef = useRef(false);

  const renderMap = () => {
    const L = window.L;
    const map = mapInstanceRef.current;
    if (!L || !map) return;

    layersRef.current.forEach(l => map.removeLayer(l));
    layersRef.current = [];

    const bounds = [];

    // Desenhar polígono da freguesia
    if (polygonFeature) {
      try {
        const geoLayer = L.geoJSON(polygonFeature, {
          style: { color: "#2d6a4f", weight: 2.5, fillColor: "#2d6a4f", fillOpacity: 0.06, dashArray: "6,4" }
        }).addTo(map);
        layersRef.current.push(geoLayer);
        const polyBounds = geoLayer.getBounds();
        if (polyBounds.isValid()) bounds.push(...Object.values(polyBounds));
      } catch {}
    }

    pontos.forEach(p => {
      const lat = parseFloat(String(p.lat).replace(",", ".").trim());
      const lon = parseFloat(String(p.lon).replace(",", ".").trim());
      if (isNaN(lat) || isNaN(lon) || lat === 0 || lon === 0) return;
      console.log("Pin:", p.nome, lat, lon);
      bounds.push([lat, lon]);

      const frg = p.freguesia;
      const limFregMapa = polygonFeature?.properties?.freguesia;
      const dentroLimite = !limFregMapa || !frg || frg === limFregMapa;
      const cor = p.tipo === "produtor" ? "#e8820c" : "#2d6a4f";
      const corFora = "#aaaaaa";
      const emoji = p.tipo === "produtor" ? "☀" : "🏠";
      const corFinal = dentroLimite ? cor : corFora;
      const iconHtml = `<div style="background:${corFinal};width:34px;height:34px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:16px;border:3px solid white;box-shadow:0 2px 8px rgba(0,0,0,0.25);opacity:${dentroLimite ? 1 : 0.5}">${emoji}</div>`;
      const icon = L.divIcon({ html: iconHtml, className: "", iconSize: [34, 34], iconAnchor: [17, 17] });

      const foraTexto = dentroLimite ? "" : "<br><span style='color:#e05c5c;font-size:11px'>⚠ Fora do limite da freguesia</span>";
      const marker = L.marker([lat, lon], { icon }).addTo(map)
        .bindPopup(`<b>${p.nome}</b><br><small style="color:#666">${p.cpe}</small>${p.localidade ? `<br>${p.localidade}` : ""}${foraTexto}`);
      layersRef.current.push(marker);

      if (mostrarRaios && p.tipo === "produtor" && dentroLimite) {
        const raiosProd = config?.raiosProdutores || {};
        const raioKm = parseFloat(raiosProd[p.id] || config?.raioPadrao || 3);
        const circle = L.circle([lat, lon], {
          radius: raioKm * 1000, color: "#e8820c", fillColor: "#e8820c",
          fillOpacity: 0.07, weight: 2, dashArray: "6,4",
        }).addTo(map);
        layersRef.current.push(circle);
      }
    });

    // Desenhar linhas de redistribuição
    if (linhasRedistribuicao && linhasRedistribuicao.length > 0) {
      const pontosPorId = {};
      pontos.forEach(p => {
        const lat = parseFloat(String(p.lat).replace(",", ".").trim());
        const lon = parseFloat(String(p.lon).replace(",", ".").trim());
        if (!isNaN(lat) && !isNaN(lon)) pontosPorId[p.id] = [lat, lon];
      });

      // Agrupar por par produtor-beneficiário para evitar linhas duplicadas
      const pares = {};
      linhasRedistribuicao.forEach(l => {
        const key = `${l.prodId}-${l.benId}`;
        if (!pares[key]) pares[key] = { ...l, kwTotal: 0 };
        pares[key].kwTotal += l.kw;
      });

      Object.values(pares).forEach(l => {
        const latP = pontosPorId[l.prodId];
        const latB = pontosPorId[l.benId];
        if (!latP || !latB) return;

        const linha = L.polyline([latP, latB], {
          color: "#2d6a4f",
          weight: Math.min(Math.max(l.kwTotal / 20, 1.5), 5),
          opacity: 0.7,
          dashArray: "8,4",
        }).addTo(map);
        linha.bindPopup(`<b>${l.prodNome || l.prodCPE}</b> → <b>${l.benNome || l.benCPE}</b><br><small>${l.kwTotal.toFixed(1)} kW transferidos</small>`);
        layersRef.current.push(linha);

        // Seta no meio da linha
        const midLat = (latP[0] + latB[0]) / 2;
        const midLon = (latP[1] + latB[1]) / 2;
        const arrowIcon = L.divIcon({
          html: `<div style="background:#2d6a4f;color:white;border-radius:50%;width:22px;height:22px;display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:700;box-shadow:0 1px 4px rgba(0,0,0,0.3)">${l.kwTotal.toFixed(0)}</div>`,
          className: "", iconSize: [22, 22], iconAnchor: [11, 11]
        });
        const midMarker = L.marker([midLat, midLon], { icon: arrowIcon }).addTo(map);
        layersRef.current.push(midMarker);
      });
    }

    if (bounds.length > 0) {
      try {
        const lBounds = window.L.latLngBounds(bounds.map(b => window.L.latLng(b[0], b[1])));
        if (lBounds.isValid()) {
          map.fitBounds(lBounds, { padding: [50, 50], maxZoom: 14, animate: false });
        } else {
          map.setView([40.9, -8.5], 12);
        }
      } catch(e) {
        console.error("fitBounds error:", e);
        map.setView([40.9, -8.5], 12);
      }
    } else {
      map.setView([40.9, -8.5], 12);
    }
  };

  useEffect(() => {
    if (!document.getElementById("leaflet-css")) {
      const link = document.createElement("link");
      link.id = "leaflet-css";
      link.rel = "stylesheet";
      link.href = "https://unpkg.com/leaflet@1.9.4/dist/leaflet.css";
      document.head.appendChild(link);
    }

    const initMap = () => {
      if (!mapRef.current || mapInstanceRef.current) return;
      const L = window.L;
      if (!L) return;
      const map = L.map(mapRef.current, { zoomControl: true }).setView([39.7, -8.5], 9);
      L.tileLayer("https://{s}.basemaps.cartocdn.com/rastertiles/voyager/{z}/{x}/{y}{r}.png", {
        attribution: "© OpenStreetMap © CARTO",
        subdomains: "abcd",
        maxZoom: 19,
      }).addTo(map);
      mapInstanceRef.current = map;
      setTimeout(() => renderMap(), 600);
    };

    if (!window.L) {
      const script = document.createElement("script");
      script.src = "https://unpkg.com/leaflet@1.9.4/dist/leaflet.js";
      script.onload = () => setTimeout(initMap, 200);
      document.head.appendChild(script);
    } else {
      setTimeout(initMap, 200);
    }

    return () => {
      if (mapInstanceRef.current) { mapInstanceRef.current.remove(); mapInstanceRef.current = null; }
    };
  }, []);

  useEffect(() => {
    if (!mapInstanceRef.current) return;
    const timer = setTimeout(() => renderMap(), 300);
    return () => clearTimeout(timer);
  }, [pontos, mostrarRaios, config, polygonFeature, linhasRedistribuicao]);

  return (
    <div style={{ borderRadius: 10, overflow: "hidden", border: "1px solid #d8ede6" }}>
      <div ref={mapRef} style={{ height: 400, width: "100%" }} />
    </div>
  );
}

// ─── Modal ────────────────────────────────────────────────────────────────────
function Modal({ title, onClose, children }) {
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.35)", zIndex: 100, display: "flex", alignItems: "center", justifyContent: "center" }}>
      <div style={{ background: "#ffffff", border: "1px solid #d8ede6", borderRadius: 16, padding: 28, width: "min(680px, 95vw)", maxHeight: "90vh", overflowY: "auto", boxShadow: "0 8px 32px rgba(45,106,79,0.15)" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 24 }}>
          <div style={S.cardTitle}>{title}</div>
          <button onClick={onClose} style={{ background: "none", border: "none", color: "#7a9e8e", fontSize: 22, cursor: "pointer" }}>×</button>
        </div>
        {children}
      </div>
    </div>
  );
}

// ─── Formulário Produtor ──────────────────────────────────────────────────────
const EMPTY_PROD = { cpe: "", nome: "", nif: "", morada: "", localidade: "", cp: "", lat: "", lon: "", potencia: "", distrito: "", municipio: "", freguesia: "" };

function FormProdutor({ initial, onSave, onCancel, polygonFeature, listaFreguesias }) {
  const [f, setF] = useState(initial || EMPTY_PROD);
  const set = k => e => setF(p => ({ ...p, [k]: e.target.value }));

  const distritos = listaFreguesias ? [...new Set(listaFreguesias.map(x => x.distrito))].sort() : [];
  const municipios = f.distrito ? [...new Set(listaFreguesias.filter(x => x.distrito === f.distrito).map(x => x.municipio))].sort() : [];
  const freguesias = f.municipio ? listaFreguesias.filter(x => x.municipio === f.municipio).map(x => x.freguesia).sort() : [];

  const setDistrito = e => setF(p => ({ ...p, distrito: e.target.value, municipio: "", freguesia: "" }));
  const setMunicipio = e => setF(p => ({ ...p, municipio: e.target.value, freguesia: "" }));
  const setFreguesia = e => setF(p => ({ ...p, freguesia: e.target.value }));

  const limiteFreguesia = polygonFeature?.properties?.freguesia;
  const matchLimite = !limiteFreguesia || f.freguesia === limiteFreguesia;

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
      <div style={S.grid2}>
        <div style={S.field}><label style={S.label}>CPE *</label><input style={S.input} value={f.cpe} onChange={set("cpe")} placeholder="PT0002000..." /></div>
        <div style={S.field}><label style={S.label}>Nome / Razão Social *</label><input style={S.input} value={f.nome} onChange={set("nome")} /></div>
      </div>
      <div style={S.grid2}>
        <div style={S.field}><label style={S.label}>NIF</label><input style={S.input} value={f.nif} onChange={set("nif")} /></div>
        <div style={S.field}><label style={S.label}>Potência instalada (kW)</label><input style={S.input} type="number" value={f.potencia} onChange={set("potencia")} /></div>
      </div>
      <div style={S.field}><label style={S.label}>Morada</label><input style={S.input} value={f.morada} onChange={set("morada")} /></div>
      <div style={S.grid2}>
        <div style={S.field}><label style={S.label}>Código Postal</label><input style={S.input} value={f.cp} onChange={set("cp")} /></div>
        <div style={{ display: "flex", gap: 8 }}>
          <div style={S.field}><label style={S.label}>Lat.</label><input style={S.input} type="number" value={f.lat} onChange={set("lat")} placeholder="38.7..." /></div>
          <div style={S.field}><label style={S.label}>Lon.</label><input style={S.input} type="number" value={f.lon} onChange={set("lon")} placeholder="-9.1..." /></div>
        </div>
      </div>
      {listaFreguesias && listaFreguesias.length > 0 ? (
        <div style={S.grid3}>
          <div style={S.field}><label style={S.label}>Distrito</label>
            <select style={S.select} value={f.distrito} onChange={setDistrito}>
              <option value="">— Selecionar —</option>
              {distritos.map(d => <option key={d}>{d}</option>)}
            </select>
          </div>
          <div style={S.field}><label style={S.label}>Município</label>
            <select style={S.select} value={f.municipio} onChange={setMunicipio} disabled={!f.distrito}>
              <option value="">— Selecionar —</option>
              {municipios.map(m => <option key={m}>{m}</option>)}
            </select>
          </div>
          <div style={S.field}><label style={S.label}>Freguesia</label>
            <select style={S.select} value={f.freguesia} onChange={setFreguesia} disabled={!f.municipio}>
              <option value="">— Selecionar —</option>
              {freguesias.map(fr => <option key={fr}>{fr}</option>)}
            </select>
          </div>
        </div>
      ) : (
        <div style={S.field}><label style={S.label}>Localidade</label><input style={S.input} value={f.localidade} onChange={set("localidade")} /></div>
      )}
      <div style={S.field}><label style={S.label}>Localidade / Lugar</label><input style={S.input} value={f.localidade || ""} onChange={set("localidade")} placeholder="Ex: Espargo (opcional)" /></div>
      {limiteFreguesia && f.freguesia && !matchLimite && (
        <div style={S.alert("warn")}>⚠ Este produtor está na freguesia <b>{f.freguesia}</b>, diferente do limite ativo (<b>{limiteFreguesia}</b>). Não participará na redistribuição.</div>
      )}
      {limiteFreguesia && f.freguesia && matchLimite && (
        <div style={S.alert("ok")}>✓ Freguesia dentro do limite ativo.</div>
      )}
      <div style={{ display: "flex", gap: 12, justifyContent: "flex-end", marginTop: 8 }}>
        <button style={S.btn("ghost")} onClick={onCancel}>Cancelar</button>
        <button style={S.btn("primary")} onClick={() => { if (!f.cpe || !f.nome) return alert("CPE e Nome são obrigatórios."); onSave(f); }}>Guardar</button>
      </div>
    </div>
  );
}

// ─── Formulário Beneficiário ──────────────────────────────────────────────────
const PROF_OPTIONS = ["Empregado", "Desempregado", "Reformado", "Estudante", "Incapacitado", "Outro"];
const EMPTY_BEN = { cpe: "", nome: "", nif: "", morada: "", localidade: "", cp: "", lat: "", lon: "", membros: [{ nome: "", parentesco: "", idade: "", situacao: "" }], distrito: "", municipio: "", freguesia: "" };

function FormBeneficiario({ initial, onSave, onCancel, polygonFeature, listaFreguesias }) {
  const [f, setF] = useState(initial ? JSON.parse(JSON.stringify(initial)) : JSON.parse(JSON.stringify(EMPTY_BEN)));
  const set = k => e => setF(p => ({ ...p, [k]: e.target.value }));
  const setMembro = (i, k) => e => { const m = [...f.membros]; m[i] = { ...m[i], [k]: e.target.value }; setF(p => ({ ...p, membros: m })); };
  const addMembro = () => setF(p => ({ ...p, membros: [...p.membros, { nome: "", parentesco: "", idade: "", situacao: "" }] }));
  const remMembro = i => setF(p => ({ ...p, membros: p.membros.filter((_, j) => j !== i) }));

  const distritos = listaFreguesias ? [...new Set(listaFreguesias.map(x => x.distrito))].sort() : [];
  const municipios = f.distrito ? [...new Set(listaFreguesias.filter(x => x.distrito === f.distrito).map(x => x.municipio))].sort() : [];
  const freguesias = f.municipio ? listaFreguesias.filter(x => x.municipio === f.municipio).map(x => x.freguesia).sort() : [];

  const setDistrito = e => setF(p => ({ ...p, distrito: e.target.value, municipio: "", freguesia: "" }));
  const setMunicipio = e => setF(p => ({ ...p, municipio: e.target.value, freguesia: "" }));
  const setFreguesia = e => setF(p => ({ ...p, freguesia: e.target.value }));

  const limiteFreguesia = polygonFeature?.properties?.freguesia;
  const matchLimite = !limiteFreguesia || f.freguesia === limiteFreguesia;

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
      <div style={S.grid2}>
        <div style={S.field}><label style={S.label}>CPE *</label><input style={S.input} value={f.cpe} onChange={set("cpe")} placeholder="PT0002000..." /></div>
        <div style={S.field}><label style={S.label}>Nome *</label><input style={S.input} value={f.nome} onChange={set("nome")} /></div>
      </div>
      <div style={S.grid3}>
        <div style={S.field}><label style={S.label}>NIF</label><input style={S.input} value={f.nif} onChange={set("nif")} /></div>
        <div style={S.field}><label style={S.label}>Código Postal</label><input style={S.input} value={f.cp} onChange={set("cp")} /></div>
        <div style={{ display: "flex", gap: 8 }}>
          <div style={S.field}><label style={S.label}>Lat.</label><input style={S.input} type="number" value={f.lat} onChange={set("lat")} /></div>
          <div style={S.field}><label style={S.label}>Lon.</label><input style={S.input} type="number" value={f.lon} onChange={set("lon")} /></div>
        </div>
      </div>
      <div style={S.field}><label style={S.label}>Morada</label><input style={S.input} value={f.morada} onChange={set("morada")} /></div>
      {listaFreguesias && listaFreguesias.length > 0 ? (
        <div style={S.grid3}>
          <div style={S.field}><label style={S.label}>Distrito</label>
            <select style={S.select} value={f.distrito} onChange={setDistrito}>
              <option value="">— Selecionar —</option>
              {distritos.map(d => <option key={d}>{d}</option>)}
            </select>
          </div>
          <div style={S.field}><label style={S.label}>Município</label>
            <select style={S.select} value={f.municipio} onChange={setMunicipio} disabled={!f.distrito}>
              <option value="">— Selecionar —</option>
              {municipios.map(m => <option key={m}>{m}</option>)}
            </select>
          </div>
          <div style={S.field}><label style={S.label}>Freguesia</label>
            <select style={S.select} value={f.freguesia} onChange={setFreguesia} disabled={!f.municipio}>
              <option value="">— Selecionar —</option>
              {freguesias.map(fr => <option key={fr}>{fr}</option>)}
            </select>
          </div>
        </div>
      ) : (
        <div style={S.field}><label style={S.label}>Localidade</label><input style={S.input} value={f.localidade} onChange={set("localidade")} /></div>
      )}
      {limiteFreguesia && f.freguesia && !matchLimite && (
        <div style={S.alert("warn")}>⚠ Este beneficiário está na freguesia <b>{f.freguesia}</b>, diferente do limite ativo (<b>{limiteFreguesia}</b>). Não receberá kW na redistribuição.</div>
      )}
      {limiteFreguesia && f.freguesia && matchLimite && (
        <div style={S.alert("ok")}>✓ Freguesia dentro do limite ativo.</div>
      )}
      <div style={S.field}><label style={S.label}>Localidade / Lugar</label><input style={S.input} value={f.localidade || ""} onChange={set("localidade")} placeholder="Ex: Espargo (opcional)" /></div>
      <div style={{ marginTop: 8 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
          <div style={{ fontSize: 13, color: "#2d6a4f", fontWeight: 700 }}>Agregado Familiar *</div>
          <button style={S.btn("ghost")} onClick={addMembro}>+ Membro</button>
        </div>
        {f.membros.map((m, i) => (
          <div key={i} style={{ background: "#f8fdfb", borderRadius: 10, padding: 14, marginBottom: 10, border: "1px solid #d8ede6" }}>
            <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 10 }}>
              <div style={{ fontSize: 11, color: "#7a9e8e", fontWeight: 600 }}>MEMBRO {i + 1}</div>
              {f.membros.length > 1 && <button style={{ background: "none", border: "none", color: "#e05c5c", cursor: "pointer", fontSize: 12 }} onClick={() => remMembro(i)}>Remover</button>}
            </div>
            <div style={S.grid2}>
              <div style={S.field}><label style={S.label}>Nome</label><input style={S.input} value={m.nome} onChange={setMembro(i, "nome")} /></div>
              <div style={S.field}><label style={S.label}>Parentesco</label><input style={S.input} value={m.parentesco} onChange={setMembro(i, "parentesco")} placeholder="Titular, Cônjuge, Filho..." /></div>
            </div>
            <div style={{ ...S.grid2, marginTop: 10 }}>
              <div style={S.field}><label style={S.label}>Idade</label><input style={S.input} type="number" value={m.idade} onChange={setMembro(i, "idade")} /></div>
              <div style={S.field}><label style={S.label}>Situação Profissional</label>
                <select style={S.select} value={m.situacao} onChange={setMembro(i, "situacao")}>
                  <option value="">—</option>
                  {PROF_OPTIONS.map(o => <option key={o}>{o}</option>)}
                </select>
              </div>
            </div>
          </div>
        ))}
      </div>
      <div style={{ display: "flex", gap: 12, justifyContent: "flex-end" }}>
        <button style={S.btn("ghost")} onClick={onCancel}>Cancelar</button>
        <button style={S.btn("primary")} onClick={() => { if (!f.cpe || !f.nome) return alert("CPE e Nome são obrigatórios."); if (!f.membros.length) return alert("É necessário pelo menos um membro do agregado."); onSave(f); }}>Guardar</button>
      </div>

      {/* Gráfico de Cobertura */}
      <div style={S.card}>
        <div style={S.cardTitle}>🎯 Cobertura da Comunidade</div>
        {(() => {
          const totalBeneficiariosFreguesia = limFrg ? beneficiarios.filter(b => b.freguesia === limFrg) : beneficiarios;
          const totalProdutoresFreguesia = limFrg ? produtores.filter(p => p.freguesia === limFrg) : produtores;

          // Beneficiários servidos no último relatório
          const ultimoRel = relatorios.sort((a, b) => b.data > a.data ? 1 : -1)[0];
          const benServidos = ultimoRel ? new Set(ultimoRel.linhas.map(l => l.benId)).size : 0;
          const benNaoServidos = totalBeneficiariosFreguesia.length - benServidos;

          // kW totais disponíveis vs distribuídos no último mês
          const kwInstalado = totalProdutoresFreguesia.reduce((s, p) => s + (parseFloat(p.potencia) || 0), 0);
          const kwDistribuidoUltimo = ultimoRel ? ultimoRel.linhas.reduce((s, l) => s + l.kw, 0) : 0;
          const kwNaoAlocadoUltimo = ultimoRel ? Object.values(ultimoRel.naoAlocado || {}).reduce((s, v) => s + v, 0) : 0;

          // Direito total dos beneficiários
          const direitoTotal = totalBeneficiariosFreguesia.reduce((s, b) => s + (b.membros?.length || 1) * (parseFloat(config.kwPorMembro) || 50), 0);
          const pctBenServidos = totalBeneficiariosFreguesia.length > 0 ? (benServidos / totalBeneficiariosFreguesia.length * 100).toFixed(0) : 0;
          const pctDireitoSatisfeito = direitoTotal > 0 ? Math.min((kwDistribuidoUltimo / direitoTotal * 100), 100).toFixed(0) : 0;

          return (
            <div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 24 }}>

                {/* Cobertura de famílias */}
                <div>
                  <div style={{ fontSize: 12, fontWeight: 700, color: "#3d5a4e", marginBottom: 16 }}>Famílias Beneficiadas</div>
                  <div style={{ position: "relative", width: 160, height: 160, margin: "0 auto 16px" }}>
                    <svg viewBox="0 0 36 36" style={{ width: "100%", height: "100%", transform: "rotate(-90deg)" }}>
                      <circle cx="18" cy="18" r="15.9" fill="none" stroke="#e8f5ef" strokeWidth="3.5" />
                      <circle cx="18" cy="18" r="15.9" fill="none" stroke="#2d6a4f" strokeWidth="3.5"
                        strokeDasharray={`${pctBenServidos} ${100 - pctBenServidos}`}
                        strokeLinecap="round" />
                    </svg>
                    <div style={{ position: "absolute", inset: 0, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center" }}>
                      <div style={{ fontSize: 28, fontWeight: 700, color: "#2d6a4f" }}>{pctBenServidos}%</div>
                      <div style={{ fontSize: 10, color: "#7a9e8e", textTransform: "uppercase", letterSpacing: 1 }}>cobertura</div>
                    </div>
                  </div>
                  <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                      <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                        <div style={{ width: 10, height: 10, borderRadius: "50%", background: "#2d6a4f" }} />
                        <span style={{ fontSize: 12, color: "#3d5a4e" }}>Servidas</span>
                      </div>
                      <span style={{ fontSize: 13, fontWeight: 700, color: "#2d6a4f" }}>{benServidos}</span>
                    </div>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                      <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                        <div style={{ width: 10, height: 10, borderRadius: "50%", background: "#e8f5ef", border: "1px solid #c8e6d8" }} />
                        <span style={{ fontSize: 12, color: "#3d5a4e" }}>Por servir</span>
                      </div>
                      <span style={{ fontSize: 13, fontWeight: 700, color: "#7a9e8e" }}>{benNaoServidos}</span>
                    </div>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                      <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                        <div style={{ width: 10, height: 10, borderRadius: 2, background: "#d8ede6" }} />
                        <span style={{ fontSize: 12, color: "#3d5a4e" }}>Total registados</span>
                      </div>
                      <span style={{ fontSize: 13, fontWeight: 700, color: "#3d5a4e" }}>{totalBeneficiariosFreguesia.length}</span>
                    </div>
                  </div>
                </div>

                {/* Cobertura de kW */}
                <div>
                  <div style={{ fontSize: 12, fontWeight: 700, color: "#3d5a4e", marginBottom: 16 }}>Direito Energético Satisfeito</div>
                  <div style={{ position: "relative", width: 160, height: 160, margin: "0 auto 16px" }}>
                    <svg viewBox="0 0 36 36" style={{ width: "100%", height: "100%", transform: "rotate(-90deg)" }}>
                      <circle cx="18" cy="18" r="15.9" fill="none" stroke="#e8f5ef" strokeWidth="3.5" />
                      <circle cx="18" cy="18" r="15.9" fill="none" stroke="#f5a623" strokeWidth="3.5"
                        strokeDasharray={`${pctDireitoSatisfeito} ${100 - pctDireitoSatisfeito}`}
                        strokeLinecap="round" />
                    </svg>
                    <div style={{ position: "absolute", inset: 0, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center" }}>
                      <div style={{ fontSize: 28, fontWeight: 700, color: "#f5a623" }}>{pctDireitoSatisfeito}%</div>
                      <div style={{ fontSize: 10, color: "#7a9e8e", textTransform: "uppercase", letterSpacing: 1 }}>satisfeito</div>
                    </div>
                  </div>
                  <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                      <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                        <div style={{ width: 10, height: 10, borderRadius: "50%", background: "#f5a623" }} />
                        <span style={{ fontSize: 12, color: "#3d5a4e" }}>kW distribuídos</span>
                      </div>
                      <span style={{ fontSize: 13, fontWeight: 700, color: "#f5a623" }}>{kwDistribuidoUltimo.toFixed(0)}</span>
                    </div>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                      <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                        <div style={{ width: 10, height: 10, borderRadius: "50%", background: "#fde8e8", border: "1px solid #f0b0b0" }} />
                        <span style={{ fontSize: 12, color: "#3d5a4e" }}>kW não alocados</span>
                      </div>
                      <span style={{ fontSize: 13, fontWeight: 700, color: "#e05c5c" }}>{kwNaoAlocadoUltimo.toFixed(0)}</span>
                    </div>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                      <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                        <div style={{ width: 10, height: 10, borderRadius: 2, background: "#d8ede6" }} />
                        <span style={{ fontSize: 12, color: "#3d5a4e" }}>Direito total</span>
                      </div>
                      <span style={{ fontSize: 13, fontWeight: 700, color: "#3d5a4e" }}>{direitoTotal.toFixed(0)} kW</span>
                    </div>
                  </div>
                </div>
              </div>

              {!ultimoRel && (
                <div style={{ ...S.alert("info"), marginTop: 16 }}>
                  Executa a primeira redistribuição para ver os dados de cobertura.
                </div>
              )}
            </div>
          );
        })()}
      </div>
    </div>
  );
}

// ─── Aba Produtores ───────────────────────────────────────────────────────────
function TabProdutores({ produtores, polygonFeature, listaFreguesias }) {
  const [modal, setModal] = useState(null);
  const [search, setSearch] = useState("");
  const [verMapa, setVerMapa] = useState(false);

  const handleSave = async (data) => {
    const id = modal === "new" ? uid() : modal.data.id;
    await setDoc(doc(db, "produtores", id), { ...data, id });
    setModal(null);
  };

  const filtered = produtores.filter(p => p.cpe.toLowerCase().includes(search.toLowerCase()) || p.nome.toLowerCase().includes(search.toLowerCase()));
  const pontosMapa = produtores.filter(p => p.lat && p.lon).map(p => ({ ...p, tipo: "produtor" }));
  const limFrg = polygonFeature?.properties?.freguesia;
  const foraLimite = limFrg ? produtores.filter(p => p.freguesia && p.freguesia !== limFrg).length : 0;

  return (
    <div>
      {modal && <Modal title={modal === "new" ? "Novo Produtor" : "Editar Produtor"} onClose={() => setModal(null)}>
        <FormProdutor initial={modal !== "new" ? modal.data : null} onSave={handleSave} onCancel={() => setModal(null)} polygonFeature={polygonFeature} listaFreguesias={listaFreguesias} />
      </Modal>}
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
        <input style={{ ...S.input, width: 280 }} placeholder="Pesquisar por CPE ou nome..." value={search} onChange={e => setSearch(e.target.value)} />
        <div style={{ display: "flex", gap: 10 }}>
          <button style={S.btn(verMapa ? "primary" : "ghost")} onClick={() => setVerMapa(v => !v)}>🗺 {verMapa ? "Ocultar Mapa" : "Ver Mapa"}</button>
          <button style={S.btn("primary")} onClick={() => setModal("new")}>+ Novo Produtor</button>
        </div>
      </div>
      {foraLimite > 0 && <div style={S.alert("warn")}>⚠ {foraLimite} produtor(es) fora do limite da freguesia — não participarão na redistribuição.</div>}
      {verMapa && (
        <div style={S.card}>
          <div style={S.cardTitle}>☀ Localização dos Produtores</div>
          {pontosMapa.length === 0
            ? <div style={S.empty}>Nenhum produtor com coordenadas definidas</div>
            : <LeafletMap pontos={pontosMapa} mostrarRaios={false} polygonFeature={polygonFeature} />
          }
          <div style={{ marginTop: 10, fontSize: 12, color: "#7a9e8e", fontStyle: "italic" }}>
            {pontosMapa.length} de {produtores.length} produtores com coordenadas — pins cinzentos estão fora do limite
          </div>
        </div>
      )}
      <div style={S.card}>
        <div style={S.cardTitle}>Produtores Registados ({filtered.length})</div>
        {filtered.length === 0 ? <div style={S.empty}>☀ Nenhum produtor registado</div> : (
          <table style={S.table}>
            <thead><tr><th style={S.th}>CPE</th><th style={S.th}>Nome</th><th style={S.th}>Freguesia</th><th style={S.th}>Potência (kW)</th><th style={S.th}>Limite</th><th style={S.th}></th></tr></thead>
            <tbody>
              {filtered.map(p => {
                const limFrg = polygonFeature?.properties?.freguesia;
                const dentro = !limFrg || p.freguesia === limFrg;
                return (
                  <tr key={p.id}>
                    <td style={S.td}><span style={S.badge("blue")}>{p.cpe}</span></td>
                    <td style={S.td}>{p.nome}</td>
                    <td style={S.td}>{p.freguesia || "—"}</td>
                    <td style={S.td}>{p.potencia || "—"}</td>
                    <td style={S.td}>{polygonFeature ? <span style={S.badge(dentro ? "green" : "red")}>{dentro ? "✓ Dentro" : "✗ Fora"}</span> : <span style={{ color: "#b7ddd0", fontSize: 11 }}>—</span>}</td>
                    <td style={S.td}>
                      <div style={{ display: "flex", gap: 8 }}>
                        <button style={S.btn("ghost")} onClick={() => setModal({ data: p })}>Editar</button>
                        <button style={S.btn("danger")} onClick={async () => { if (confirm("Remover produtor?")) await deleteDoc(doc(db, "produtores", p.id)); }}>✕</button>
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        )}
      </div>
    </div>
  );
}

// ─── Aba Beneficiários ────────────────────────────────────────────────────────
function TabBeneficiarios({ beneficiarios, polygonFeature, listaFreguesias }) {
  const [modal, setModal] = useState(null);
  const [search, setSearch] = useState("");
  const [verMapa, setVerMapa] = useState(false);

  const handleSave = async (data) => {
    const id = modal === "new" ? uid() : modal.data.id;
    await setDoc(doc(db, "beneficiarios", id), { ...data, id });
    setModal(null);
  };

  const filtered = beneficiarios.filter(b => b.cpe.toLowerCase().includes(search.toLowerCase()) || b.nome.toLowerCase().includes(search.toLowerCase()));
  const pontosMapa = beneficiarios.filter(b => b.lat && b.lon).map(b => ({ ...b, tipo: "beneficiario" }));
  const limFrgB = polygonFeature?.properties?.freguesia;
  const foraLimite = limFrgB ? beneficiarios.filter(b => b.freguesia && b.freguesia !== limFrgB).length : 0;

  return (
    <div>
      {modal && <Modal title={modal === "new" ? "Novo Beneficiário" : "Editar Beneficiário"} onClose={() => setModal(null)}>
        <FormBeneficiario initial={modal !== "new" ? modal.data : null} onSave={handleSave} onCancel={() => setModal(null)} polygonFeature={polygonFeature} listaFreguesias={listaFreguesias} />
      </Modal>}
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
        <input style={{ ...S.input, width: 280 }} placeholder="Pesquisar por CPE ou nome..." value={search} onChange={e => setSearch(e.target.value)} />
        <div style={{ display: "flex", gap: 10 }}>
          <button style={S.btn(verMapa ? "primary" : "ghost")} onClick={() => setVerMapa(v => !v)}>🗺 {verMapa ? "Ocultar Mapa" : "Ver Mapa"}</button>
          <button style={S.btn("primary")} onClick={() => setModal("new")}>+ Novo Beneficiário</button>
        </div>
      </div>
      {foraLimite > 0 && <div style={S.alert("warn")}>⚠ {foraLimite} beneficiário(s) fora do limite da freguesia — não receberão kW.</div>}
      {verMapa && (
        <div style={S.card}>
          <div style={S.cardTitle}>🏠 Localização dos Beneficiários</div>
          {pontosMapa.length === 0
            ? <div style={S.empty}>Nenhum beneficiário com coordenadas definidas</div>
            : <LeafletMap pontos={pontosMapa} mostrarRaios={false} polygonFeature={polygonFeature} />
          }
          <div style={{ marginTop: 10, fontSize: 12, color: "#7a9e8e", fontStyle: "italic" }}>
            {pontosMapa.length} de {beneficiarios.length} beneficiários com coordenadas — pins cinzentos estão fora do limite
          </div>
        </div>
      )}
      <div style={S.card}>
        <div style={S.cardTitle}>Beneficiários Registados ({filtered.length})</div>
        {filtered.length === 0 ? <div style={S.empty}>⚡ Nenhum beneficiário registado</div> : (
          <table style={S.table}>
            <thead><tr><th style={S.th}>CPE</th><th style={S.th}>Nome</th><th style={S.th}>Freguesia</th><th style={S.th}>Agregado</th><th style={S.th}>Limite</th><th style={S.th}></th></tr></thead>
            <tbody>
              {filtered.map(b => {
                const limFrg = polygonFeature?.properties?.freguesia;
                const dentro = !limFrg || b.freguesia === limFrg;
                return (
                  <tr key={b.id}>
                    <td style={S.td}><span style={S.badge("green")}>{b.cpe}</span></td>
                    <td style={S.td}>{b.nome}</td>
                    <td style={S.td}>{b.freguesia || "—"}</td>
                    <td style={S.td}>{b.membros?.length || 0} membro(s)</td>
                    <td style={S.td}>{polygonFeature ? <span style={S.badge(dentro ? "green" : "red")}>{dentro ? "✓ Dentro" : "✗ Fora"}</span> : <span style={{ color: "#b7ddd0", fontSize: 11 }}>—</span>}</td>
                    <td style={S.td}>
                      <div style={{ display: "flex", gap: 8 }}>
                        <button style={S.btn("ghost")} onClick={() => setModal({ data: b })}>Editar</button>
                        <button style={S.btn("danger")} onClick={async () => { if (confirm("Remover beneficiário?")) await deleteDoc(doc(db, "beneficiarios", b.id)); }}>✕</button>
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        )}
      </div>
    </div>
  );
}

// ─── Aba Configurações ────────────────────────────────────────────────────────
function TabConfiguracoes({ config, produtores, beneficiarios, polygonFeature, setPolygonFeature, setListaFreguesias }) {
  const [local, setLocal] = useState(config);
  const [geoData, setGeoData] = useState(null);
  const [distritos, setDistritos] = useState([]);
  const [municipios, setMunicipos] = useState([]);
  const [freguesias, setFreguesias] = useState([]);
  const [selDistrito, setSelDistrito] = useState(config.limiteDistrito || "");
  const [selMunicipio, setSelMunicipio] = useState(config.limiteMunicipio || "");
  const [selFreguesia, setSelFreguesia] = useState(config.limiteFreguesia || "");

  // Inicializar local apenas uma vez com o config inicial

  const set = k => e => setLocal(p => ({ ...p, [k]: e.target.value }));
  const setRaio = (prodId, val) => setLocal(p => ({ ...p, raiosProdutores: { ...(p.raiosProdutores || {}), [prodId]: val } }));

  const handleGeoUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const data = JSON.parse(ev.target.result);
        setGeoData(data);
        const dists = [...new Set(data.features.map(f => f.properties.distrito_ilha))].sort();
        setDistritos(dists);
        setSelDistrito(""); setSelMunicipio(""); setSelFreguesia("");
        setMunicipos([]); setFreguesias([]);
      } catch { alert("Ficheiro GeoJSON inválido."); }
    };
    reader.readAsText(file);
  };

  useEffect(() => {
    if (!geoData || !selDistrito) return;
    const muns = [...new Set(geoData.features.filter(f => f.properties.distrito_ilha === selDistrito).map(f => f.properties.municipio))].sort();
    setMunicipos(muns);
    setSelMunicipio(""); setSelFreguesia(""); setFreguesias([]);
  }, [selDistrito, geoData]);

  useEffect(() => {
    if (!geoData || !selMunicipio) return;
    const fregs = geoData.features.filter(f => f.properties.municipio === selMunicipio).map(f => f.properties.freguesia).sort();
    setFreguesias(fregs);
    setSelFreguesia("");
  }, [selMunicipio, geoData]);

  useEffect(() => {
    if (!geoData || !selFreguesia) return;
    const feature = geoData.features.find(f => f.properties.freguesia === selFreguesia && f.properties.municipio === selMunicipio);
    if (feature) setPolygonFeature(feature);
  }, [selFreguesia, geoData, selMunicipio]);

  const guardar = async () => {
    const cfg = { ...local, limiteDistrito: selDistrito, limiteMunicipio: selMunicipio, limiteFreguesia: selFreguesia };
    await setDoc(doc(db, "config", "global"), cfg);
    // Guardar o polígono da freguesia no Firebase
    if (polygonFeature) {
      await setDoc(doc(db, "config", "poligono"), { feature: JSON.stringify(polygonFeature) });
    }
    // Guardar lista de freguesias no Firebase (só nomes, sem geometrias)
    if (geoData) {
      const lista = geoData.features.map(f => ({
        distrito: f.properties.distrito_ilha,
        municipio: f.properties.municipio,
        freguesia: f.properties.freguesia,
      }));
      await setDoc(doc(db, "config", "listaFreguesias"), { lista: JSON.stringify(lista) });
      setListaFreguesias(lista);
    }
    alert("Configurações guardadas!");
  };

  const migrarRegistos = async () => {
    if (!polygonFeature) { alert("Primeiro seleciona e guarda um limite de freguesia."); return; }
    const freguesia = polygonFeature.properties.freguesia;
    const municipio = polygonFeature.properties.municipio;
    const distrito = polygonFeature.properties.distrito_ilha;
    if (!freguesia) { alert("Limite ativo não tem freguesia definida. Volta a selecionar e guardar."); return; }
    if (!confirm(`Aplicar a freguesia "${freguesia}" a todos os produtores e beneficiários sem freguesia definida?`)) return;
    let count = 0;

    for (const p of produtores) {
      if (!p.freguesia) {
        await setDoc(doc(db, "produtores", p.id), { ...p, freguesia, municipio, distrito });
        count++;
      }
    }
    for (const b of beneficiarios) {
      if (!b.freguesia) {
        await setDoc(doc(db, "beneficiarios", b.id), { ...b, freguesia, municipio, distrito });
        count++;
      }
    }
    alert(`✓ ${count} registo(s) atualizados com a freguesia "${freguesia}".`);
  };

  return (
    <div>
      {/* Limite Geográfico */}
      <div style={S.card}>
        <div style={S.cardTitle}>🗺 Limite Geográfico — Freguesia</div>
        <div style={{ marginBottom: 16 }}>
          <label style={{ ...S.label, display: "block", marginBottom: 8 }}>1. Carregar ficheiro de freguesias (cont_freguesias.json)</label>
          <input type="file" accept=".json,.geojson" onChange={handleGeoUpload} style={{ color: "#3d5a4e", fontSize: 13 }} />
          <div style={{ fontSize: 11, color: "#7a9e8e", marginTop: 6, fontStyle: "italic" }}>Ficheiro cont_freguesias.json do CAOP</div>
        </div>
        {geoData && (
          <div style={{ marginTop: 16 }}>
            <div style={{ marginBottom: 8, fontSize: 11, color: "#7a9e8e", fontWeight: 600, textTransform: "uppercase", letterSpacing: 1 }}>2. Selecionar freguesia</div>
            <div style={S.grid3}>
              <div style={S.field}>
                <label style={S.label}>Distrito</label>
                <select style={S.select} value={selDistrito} onChange={e => setSelDistrito(e.target.value)}>
                  <option value="">— Selecionar —</option>
                  {distritos.map(d => <option key={d}>{d}</option>)}
                </select>
              </div>
              <div style={S.field}>
                <label style={S.label}>Município</label>
                <select style={S.select} value={selMunicipio} onChange={e => setSelMunicipio(e.target.value)} disabled={!selDistrito}>
                  <option value="">— Selecionar —</option>
                  {municipios.map(m => <option key={m}>{m}</option>)}
                </select>
              </div>
              <div style={S.field}>
                <label style={S.label}>Freguesia</label>
                <select style={S.select} value={selFreguesia} onChange={e => setSelFreguesia(e.target.value)} disabled={!selMunicipio}>
                  <option value="">— Selecionar —</option>
                  {freguesias.map(f => <option key={f}>{f}</option>)}
                </select>
              </div>
            </div>
            {selFreguesia && (
              <div style={{ ...S.alert("ok"), marginTop: 16 }}>
                ✓ Limite ativo: <b>{selFreguesia}</b>, {selMunicipio}, {selDistrito}
              </div>
            )}
          </div>
        )}
        {!geoData && polygonFeature && (
          <div style={S.alert("ok")}>✓ Limite ativo: <b>{config.limiteFreguesia}</b>, {config.limiteMunicipio} — carrega o ficheiro novamente para alterar.</div>
        )}
        {!geoData && !polygonFeature && (
          <div style={S.alert("info")}>Sem limite geográfico definido — todos os produtores e beneficiários são elegíveis.</div>
        )}
        {polygonFeature && (
          <div style={{ marginTop: 16, padding: "14px 18px", background: "#f0faf5", borderRadius: 10, border: "1px solid #c8e6d8" }}>
            <div style={{ fontSize: 13, color: "#2d6a4f", fontWeight: 700, marginBottom: 8 }}>Migrar registos existentes</div>
            <div style={{ fontSize: 12, color: "#7a9e8e", marginBottom: 12, fontStyle: "italic" }}>
              Preenche automaticamente a freguesia em todos os produtores e beneficiários que estejam dentro do limite geográfico ativo e ainda não tenham freguesia definida.
            </div>
            <button style={S.btn("primary")} onClick={migrarRegistos}>
              ⚡ Aplicar Freguesia Automaticamente
            </button>
          </div>
        )}
      </div>

      {/* Parâmetros Globais */}
      <div style={S.card}>
        <div style={S.cardTitle}>Parâmetros Globais</div>
        <div style={S.grid2}>
          <div style={S.field}>
            <label style={S.label}>kW por membro de agregado</label>
            <input style={S.input} type="number" value={local.kwPorMembro} onChange={set("kwPorMembro")} />
            <div style={{ fontSize: 11, color: "#7a9e8e", marginTop: 4, fontStyle: "italic" }}>Ex: 50 kW → agregado de 2 = direito de 100 kW</div>
          </div>
          <div style={S.field}>
            <label style={S.label}>Raio padrão de redistribuição (km)</label>
            <input style={S.input} type="number" value={local.raioPadrao} onChange={set("raioPadrao")} />
            <div style={{ fontSize: 11, color: "#7a9e8e", marginTop: 4, fontStyle: "italic" }}>Aplicado a produtores sem raio específico</div>
          </div>
        </div>
      </div>

      {/* Raio por Produtor */}
      <div style={S.card}>
        <div style={S.cardTitle}>Raio por Produtor</div>
        {produtores.length === 0 ? <div style={S.empty}>Adiciona produtores primeiro</div> : (
          <table style={S.table}>
            <thead><tr><th style={S.th}>CPE</th><th style={S.th}>Nome</th><th style={S.th}>Raio (km)</th></tr></thead>
            <tbody>
              {produtores.map(p => (
                <tr key={p.id}>
                  <td style={S.td}><span style={S.badge("orange")}>{p.cpe}</span></td>
                  <td style={S.td}>{p.nome}</td>
                  <td style={S.td}><input style={{ ...S.input, width: 100 }} type="number" placeholder={local.raioPadrao} value={(local.raiosProdutores || {})[p.id] || ""} onChange={e => setRaio(p.id, e.target.value)} /></td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>
      {/* Método de redistribuição */}
      <div style={S.card}>
        <div style={S.cardTitle}>⚖ Método de Redistribuição</div>
        <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
          {[
            {
              id: "proporcional",
              titulo: "Proporcional ao agregado",
              descricao: "Cada produtor distribui os seus kW proporcionalmente ao número de membros de cada agregado. Famílias maiores recebem mais, famílias menores recebem menos. Garante equidade relativa entre todos os beneficiários dentro do raio."
            },
            {
              id: "satisfacao",
              titulo: "Satisfação completa por distância",
              descricao: "O produtor tenta satisfazer completamente cada beneficiário antes de passar ao próximo, começando pelo mais próximo. Garante que alguns beneficiários ficam totalmente satisfeitos, mas outros podem não receber nada se os kW esgotarem."
            },
            {
              id: "igualitario",
              titulo: "Igualitário",
              descricao: "Todos os beneficiários dentro do raio recebem exatamente o mesmo valor de kW, independentemente do tamanho do agregado. Método mais simples e transparente, ideal para comunidades com agregados de dimensão semelhante."
            }
          ].map(m => (
            <div
              key={m.id}
              onClick={() => setLocal(p => ({ ...p, metodoRedistribuicao: m.id }))}
              style={{
                padding: "16px 18px",
                borderRadius: 10,
                border: `2px solid ${local.metodoRedistribuicao === m.id ? "#2d6a4f" : "#d8ede6"}`,
                background: local.metodoRedistribuicao === m.id ? "#f0faf5" : "#fafafa",
                cursor: "pointer",
                transition: "all 0.2s"
              }}
            >
              <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 6 }}>
                <div style={{
                  width: 18, height: 18, borderRadius: "50%",
                  border: `2px solid ${local.metodoRedistribuicao === m.id ? "#2d6a4f" : "#b7ddd0"}`,
                  background: local.metodoRedistribuicao === m.id ? "#2d6a4f" : "white",
                  display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0
                }}>
                  {local.metodoRedistribuicao === m.id && <div style={{ width: 8, height: 8, borderRadius: "50%", background: "white" }} />}
                </div>
                <div style={{ fontSize: 13, fontWeight: 700, color: local.metodoRedistribuicao === m.id ? "#2d6a4f" : "#3d5a4e" }}>{m.titulo}</div>
              </div>
              <div style={{ fontSize: 12, color: "#7a9e8e", lineHeight: 1.6, paddingLeft: 28, fontStyle: "italic" }}>{m.descricao}</div>
            </div>
          ))}
        </div>
      </div>

      <div style={{ display: "flex", justifyContent: "flex-end" }}>
        <button style={S.btn("primary")} onClick={guardar}>Guardar Configurações</button>
      </div>
    </div>
  );
}

// ─── Lógica de Redistribuição ─────────────────────────────────────────────────
function executarRedistribuicao(produtores, beneficiarios, config, kwDisponivel, polygonFeature) {
  const metodo = config.metodoRedistribuicao || "proporcional";
  const kwPorMembro = parseFloat(config.kwPorMembro) || 50;
  const raioPadrao = parseFloat(config.raioPadrao) || 3;
  const raiosProd = config.raiosProdutores || {};
  const direito = {};
  const recebido = {};

  // Filtrar por freguesia selecionada E por polígono geográfico
  const limiteFreguesia = polygonFeature?.properties?.freguesia;
  const produtoresValidos = produtores.filter(p => {
    if (!limiteFreguesia) return p.lat && p.lon;
    return p.freguesia === limiteFreguesia && p.lat && p.lon;
  });
  const beneficiariosValidos = beneficiarios.filter(b => {
    if (!limiteFreguesia) return b.lat && b.lon;
    return b.freguesia === limiteFreguesia && b.lat && b.lon;
  });

  beneficiariosValidos.forEach(b => { direito[b.id] = (b.membros?.length || 1) * kwPorMembro; recebido[b.id] = 0; });
  const linhas = [];
  const naoAlocado = {};

  produtoresValidos.forEach(prod => {
    const kwTotal = parseFloat(kwDisponivel[prod.id] || 0);
    if (!kwTotal) return;
    const raio = parseFloat(raiosProd[prod.id] || raioPadrao);
    const latP = parseFloat(prod.lat), lonP = parseFloat(prod.lon);

    // Beneficiários elegíveis dentro do raio com direito por satisfazer
    const elegiveis = beneficiariosValidos.filter(b => {
      const latB = parseFloat(b.lat), lonB = parseFloat(b.lon);
      if (isNaN(latP) || isNaN(lonP) || isNaN(latB) || isNaN(lonB)) return false;
      return haversine(latP, lonP, latB, lonB) <= raio && recebido[b.id] < direito[b.id];
    });

    if (!elegiveis.length) { naoAlocado[prod.id] = kwTotal; return; }

    let kwRestante = kwTotal;

    if (metodo === "proporcional") {
      // Distribuição proporcional ao nº de membros do agregado
      let continuar = true;
      while (continuar && kwRestante > 0.001) {
        continuar = false;
        const ativos = elegiveis.filter(b => recebido[b.id] < direito[b.id]);
        if (!ativos.length) break;
        const totalMembros = ativos.reduce((s, b) => s + (b.membros?.length || 1), 0);
        ativos.forEach(b => {
          if (kwRestante <= 0) return;
          const membros = b.membros?.length || 1;
          const quota = (membros / totalMembros) * kwRestante;
          const necessario = direito[b.id] - recebido[b.id];
          const atribuir = Math.min(quota, necessario, kwRestante);
          if (atribuir > 0.001) {
            recebido[b.id] += atribuir;
            kwRestante -= atribuir;
            if (atribuir < quota - 0.001) continuar = true;
            linhas.push({ prodId: prod.id, prodCPE: prod.cpe, prodNome: prod.nome, benId: b.id, benCPE: b.cpe, benNome: b.nome, kw: atribuir, distKm: haversine(latP, lonP, parseFloat(b.lat), parseFloat(b.lon)).toFixed(2) });
          }
        });
      }

    } else if (metodo === "satisfacao") {
      // Satisfação completa por distância (mais próximo primeiro)
      const elegiveisOrdenados = [...elegiveis].sort((a, b) => {
        const dA = haversine(latP, lonP, parseFloat(a.lat), parseFloat(a.lon));
        const dB = haversine(latP, lonP, parseFloat(b.lat), parseFloat(b.lon));
        return dA - dB;
      });
      for (const b of elegiveisOrdenados) {
        if (kwRestante <= 0) break;
        const necessario = direito[b.id] - recebido[b.id];
        const atribuir = Math.min(necessario, kwRestante);
        if (atribuir > 0.001) {
          recebido[b.id] += atribuir;
          kwRestante -= atribuir;
          linhas.push({ prodId: prod.id, prodCPE: prod.cpe, prodNome: prod.nome, benId: b.id, benCPE: b.cpe, benNome: b.nome, kw: atribuir, distKm: haversine(latP, lonP, parseFloat(b.lat), parseFloat(b.lon)).toFixed(2) });
        }
      }

    } else if (metodo === "igualitario") {
      // Distribuição igualitária — mesmo valor para todos
      let continuar = true;
      while (continuar && kwRestante > 0.001) {
        continuar = false;
        const ativos = elegiveis.filter(b => recebido[b.id] < direito[b.id]);
        if (!ativos.length) break;
        const quotaIgual = kwRestante / ativos.length;
        ativos.forEach(b => {
          if (kwRestante <= 0) return;
          const necessario = direito[b.id] - recebido[b.id];
          const atribuir = Math.min(quotaIgual, necessario, kwRestante);
          if (atribuir > 0.001) {
            recebido[b.id] += atribuir;
            kwRestante -= atribuir;
            if (atribuir < quotaIgual - 0.001) continuar = true;
            linhas.push({ prodId: prod.id, prodCPE: prod.cpe, prodNome: prod.nome, benId: b.id, benCPE: b.cpe, benNome: b.nome, kw: atribuir, distKm: haversine(latP, lonP, parseFloat(b.lat), parseFloat(b.lon)).toFixed(2) });
          }
        });
      }
    }

    if (kwRestante > 0.001) naoAlocado[prod.id] = kwRestante;
  });
  return { linhas, naoAlocado, recebido, direito, produtoresValidos, beneficiariosValidos };
}

// ─── Aba Redistribuição ───────────────────────────────────────────────────────
function TabRedistribuicao({ produtores, beneficiarios, config, polygonFeature }) {
  const [kwDisponivel, setKwDisponivel] = useState({});
  const [resultado, setResultado] = useState(null);
  const [msg, setMsg] = useState(null);
  const [periodo, setPeriodo] = useState(() => new Date().toISOString().slice(0, 7));
  const [verMapa, setVerMapa] = useState(false);

  const pontosMapa = [
    ...produtores.filter(p => p.lat && p.lon).map(p => ({ ...p, tipo: "produtor" })),
    ...beneficiarios.filter(b => b.lat && b.lon).map(b => ({ ...b, tipo: "beneficiario" })),
  ];

  const handleExcel = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      const wb = XLSX.read(ev.target.result, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
      const kw = {};
      rows.slice(1).forEach(r => {
        const cpe = String(r[0] || "").trim();
        const val = parseFloat(r[1]);
        if (cpe && !isNaN(val)) { const prod = produtores.find(p => p.cpe === cpe); if (prod) kw[prod.id] = val; }
      });
      setKwDisponivel(kw);
      setMsg({ type: "ok", text: `Excel lido: ${Object.keys(kw).length} produtor(es) encontrado(s).` });
    };
    reader.readAsArrayBuffer(file);
  };

  const executar = async () => {
    if (!produtores.length || !beneficiarios.length) { setMsg({ type: "err", text: "É necessário ter produtores e beneficiários registados." }); return; }
    const res = executarRedistribuicao(produtores, beneficiarios, config, kwDisponivel, polygonFeature);
    setResultado(res);
    const relatorio = {
      id: uid(), periodo, data: new Date().toISOString(),
      linhas: res.linhas, naoAlocado: res.naoAlocado, recebido: res.recebido, direito: res.direito,
      produtores: produtores.map(p => ({ id: p.id, cpe: p.cpe, nome: p.nome })),
      beneficiarios: beneficiarios.map(b => ({ id: b.id, cpe: b.cpe, nome: b.nome, membros: b.membros?.length || 1 })),
      kwDisponivel: { ...kwDisponivel },
      limiteFreguesia: config.limiteFreguesia || null,
    };
    await setDoc(doc(db, "relatorios", relatorio.id), relatorio);
    setMsg({ type: "ok", text: `Redistribuição executada — ${res.produtoresValidos.length} produtores e ${res.beneficiariosValidos.length} beneficiários dentro do limite.` });
  };

  return (
    <div>
      {msg && <div style={S.alert(msg.type)}>{msg.text}</div>}

      {/* Mapa conjunto */}
      <div style={S.card}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div style={{ ...S.cardTitle, marginBottom: 0 }}>🗺 Mapa da Rede</div>
          <button style={S.btn(verMapa ? "primary" : "ghost")} onClick={() => setVerMapa(v => !v)}>
            {verMapa ? "Ocultar" : "Ver Mapa"}
          </button>
        </div>
        {verMapa && (
          <div style={{ marginTop: 20 }}>
            {pontosMapa.length === 0
              ? <div style={S.empty}>Nenhum ponto com coordenadas definidas</div>
              : <LeafletMap pontos={pontosMapa} mostrarRaios={true} config={config} polygonFeature={polygonFeature} linhasRedistribuicao={resultado ? resultado.linhas : []} />
            }
            <div style={{ display: "flex", gap: 20, marginTop: 10, fontSize: 12, color: "#7a9e8e", fontStyle: "italic", flexWrap: "wrap" }}>
              <span>☀ Produtores ({produtores.filter(p => p.lat && p.lon).length})</span>
              <span>🏠 Beneficiários ({beneficiarios.filter(b => b.lat && b.lon).length})</span>
              <span>Círculos = raio de redistribuição</span>
              {resultado && <span>― Linhas = fluxo de energia (kW)</span>}
              {polygonFeature && <span>Contorno verde = limite da freguesia</span>}
            </div>
          </div>
        )}
      </div>

      <div style={S.card}>
        <div style={S.cardTitle}>Período</div>
        <div style={S.field}><label style={S.label}>Mês / Ano</label><input style={{ ...S.input, width: 180 }} type="month" value={periodo} onChange={e => setPeriodo(e.target.value)} /></div>
      </div>

      <div style={S.card}>
        <div style={S.cardTitle}>kW Disponíveis por Produtor</div>
        <div style={{ marginBottom: 16 }}>
          <label style={{ ...S.label, display: "block", marginBottom: 8 }}>Importar Excel (CPE | kW)</label>
          <input type="file" accept=".xlsx,.xls,.csv" onChange={handleExcel} style={{ color: "#3d5a4e", fontSize: 13 }} />
          <div style={{ fontSize: 11, color: "#7a9e8e", marginTop: 6, fontStyle: "italic" }}>Formato: coluna A = CPE, coluna B = kW disponíveis</div>
        </div>
        <hr style={S.divider} />
        <div style={{ marginBottom: 12, fontSize: 11, color: "#7a9e8e", fontWeight: 600, textTransform: "uppercase", letterSpacing: 1 }}>Ou introduzir manualmente</div>
        {produtores.length === 0 ? <div style={S.empty}>Sem produtores registados</div> : (
          <table style={S.table}>
            <thead><tr><th style={S.th}>CPE</th><th style={S.th}>Nome</th><th style={S.th}>Limite</th><th style={S.th}>kW disponíveis</th></tr></thead>
            <tbody>
              {produtores.map(p => {
                const limFrg = polygonFeature?.properties?.freguesia;
                const dentro = !limFrg || p.freguesia === limFrg;
                return (
                  <tr key={p.id} style={{ opacity: dentro ? 1 : 0.5 }}>
                    <td style={S.td}><span style={S.badge("orange")}>{p.cpe}</span></td>
                    <td style={S.td}>{p.nome}</td>
                    <td style={S.td}>{polygonFeature ? <span style={S.badge(dentro ? "green" : "red")}>{dentro ? "✓" : "✗ Fora"}</span> : "—"}</td>
                    <td style={S.td}><input style={{ ...S.input, width: 120 }} type="number" value={kwDisponivel[p.id] || ""} onChange={e => setKwDisponivel(prev => ({ ...prev, [p.id]: e.target.value }))} placeholder="0" disabled={!dentro} /></td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        )}
      </div>

      <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: 24 }}>
        <button style={S.btn("green")} onClick={executar}>▶ Executar Redistribuição</button>
      </div>

      {resultado && (
        <div style={S.card}>
          <div style={S.cardTitle}>Resultado da Redistribuição</div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 16, marginBottom: 24 }}>
            <div style={S.statBox}><div style={S.statNum}>{resultado.linhas.length}</div><div style={S.statLabel}>Atribuições</div></div>
            <div style={S.statBox}><div style={S.statNum}>{resultado.linhas.reduce((s, l) => s + l.kw, 0).toFixed(1)}</div><div style={S.statLabel}>kW Distribuídos</div></div>
            <div style={S.statBox}><div style={{ ...S.statNum, color: Object.keys(resultado.naoAlocado).length ? "#e05c5c" : "#2d6a4f" }}>{Object.values(resultado.naoAlocado).reduce((s, v) => s + v, 0).toFixed(1)}</div><div style={S.statLabel}>kW Não Alocados</div></div>
          </div>
          <table style={S.table}>
            <thead><tr><th style={S.th}>Produtor CPE</th><th style={S.th}>Beneficiário CPE</th><th style={S.th}>Nome</th><th style={S.th}>kW</th><th style={S.th}>Dist. (km)</th></tr></thead>
            <tbody>
              {resultado.linhas.map((l, i) => (
                <tr key={i}>
                  <td style={S.td}><span style={S.badge("orange")}>{l.prodCPE}</span></td>
                  <td style={S.td}><span style={S.badge("green")}>{l.benCPE}</span></td>
                  <td style={S.td}>{l.benNome}</td>
                  <td style={S.td}>{l.kw.toFixed(2)}</td>
                  <td style={S.td}>{l.distKm}</td>
                </tr>
              ))}
            </tbody>
          </table>
          {Object.keys(resultado.naoAlocado).length > 0 && (
            <div style={{ ...S.alert("err"), marginTop: 16 }}>
              ⚠ kW não alocados: {produtores.filter(p => resultado.naoAlocado[p.id]).map(p => `${p.cpe}: ${resultado.naoAlocado[p.id].toFixed(2)} kW`).join(" | ")}
            </div>
          )}
        </div>
      )}
    </div>
  );
}

// ─── Aba Relatórios ───────────────────────────────────────────────────────────
function gerarPDF(rel) {
  const kwDistribuidos = rel.linhas.reduce((s, l) => s + l.kw, 0);
  const kwNaoAlocados = Object.values(rel.naoAlocado || {}).reduce((s, v) => s + v, 0);
  const familias = new Set(rel.linhas.map(l => l.benId)).size;
  const data = new Date(rel.data).toLocaleDateString("pt-PT");

  const html = `<!DOCTYPE html>
<html lang="pt">
<head>
<meta charset="UTF-8">
<title>Relatório CER - ${rel.periodo}</title>
<style>
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body { font-family: Georgia, serif; color: #2d4a3e; padding: 40px; font-size: 13px; }
  .header { background: linear-gradient(135deg, #2d6a4f, #40916c); color: white; padding: 28px 32px; border-radius: 10px; margin-bottom: 28px; }
  .header h1 { font-size: 22px; font-weight: 700; margin-bottom: 4px; }
  .header p { font-size: 12px; opacity: 0.85; font-style: italic; }
  .header .meta { margin-top: 12px; font-size: 11px; opacity: 0.75; }
  .stats { display: grid; grid-template-columns: repeat(4, 1fr); gap: 14px; margin-bottom: 24px; }
  .stat { background: #f0faf5; border: 1px solid #c8e6d8; border-radius: 8px; padding: 14px; text-align: center; }
  .stat .num { font-size: 26px; font-weight: 700; color: #2d6a4f; }
  .stat .label { font-size: 10px; color: #7a9e8e; text-transform: uppercase; letter-spacing: 1px; margin-top: 3px; }
  .section-title { font-size: 13px; font-weight: 700; color: #2d6a4f; margin-bottom: 12px; padding-bottom: 6px; border-bottom: 2px solid #d8ede6; }
  table { width: 100%; border-collapse: collapse; margin-bottom: 24px; font-size: 12px; }
  th { padding: 9px 12px; text-align: left; font-size: 10px; color: #7a9e8e; text-transform: uppercase; letter-spacing: 1px; border-bottom: 2px solid #d8ede6; font-weight: 600; }
  td { padding: 10px 12px; border-bottom: 1px solid #eaf5f0; color: #3d5a4e; }
  tr:last-child td { border-bottom: none; }
  .badge-prod { background: #fef3dc; color: #b07c10; padding: 2px 8px; border-radius: 10px; font-size: 10px; font-weight: 600; }
  .badge-ben { background: #d8f3e8; color: #1e7a4a; padding: 2px 8px; border-radius: 10px; font-size: 10px; font-weight: 600; }
  .alert-red { background: #fde8e8; border: 1px solid #f0b0b0; border-radius: 8px; padding: 12px 16px; margin-bottom: 24px; font-size: 12px; color: #c0392b; }
  .footer { margin-top: 32px; padding-top: 16px; border-top: 1px solid #d8ede6; font-size: 11px; color: #7a9e8e; display: flex; justify-content: space-between; }
  .section { margin-bottom: 24px; }
  @media print { body { padding: 20px; } }
</style>
</head>
<body>
  <div class="header">
    <h1>Banco de Energia Solidário</h1>
    <p>Comunidade de Energia Renovável${rel.limiteFreguesia ? ` · ${rel.limiteFreguesia}` : ""}</p>
    <div class="meta">Relatório de Redistribuição Energética · Período: ${rel.periodo} · Gerado em: ${data}</div>
  </div>

  <div class="stats">
    <div class="stat"><div class="num">${rel.linhas.length}</div><div class="label">Atribuições</div></div>
    <div class="stat"><div class="num">${kwDistribuidos.toFixed(1)}</div><div class="label">kW Distribuídos</div></div>
    <div class="stat"><div class="num">${kwNaoAlocados.toFixed(1)}</div><div class="label">kW Não Alocados</div></div>
    <div class="stat"><div class="num">${familias}</div><div class="label">Famílias Servidas</div></div>
  </div>

  <div class="section">
    <div class="section-title">Detalhe das Atribuições</div>
    <table>
      <thead>
        <tr>
          <th>Produtor CPE</th>
          <th>Produtor Nome</th>
          <th>Beneficiário CPE</th>
          <th>Beneficiário Nome</th>
          <th>kW Atribuídos</th>
          <th>Distância (km)</th>
        </tr>
      </thead>
      <tbody>
        ${rel.linhas.map(l => `<tr>
          <td><span class="badge-prod">${l.prodCPE}</span></td>
          <td>${l.prodNome || "—"}</td>
          <td><span class="badge-ben">${l.benCPE}</span></td>
          <td>${l.benNome}</td>
          <td>${parseFloat(l.kw).toFixed(2)} kW</td>
          <td>${l.distKm} km</td>
        </tr>`).join("")}
      </tbody>
    </table>
  </div>

  ${kwNaoAlocados > 0 ? `
  <div class="section">
    <div class="section-title">kW Não Alocados</div>
    <div class="alert-red">
      Os seguintes produtores tiveram excedente de energia não distribuída por ausência de beneficiários elegíveis dentro do raio de redistribuição:<br><br>
      ${(rel.produtores || []).filter(p => rel.naoAlocado[p.id]).map(p => `<b>${p.cpe}</b> (${p.nome}): ${parseFloat(rel.naoAlocado[p.id]).toFixed(2)} kW`).join(" &nbsp;|&nbsp; ")}
    </div>
  </div>` : ""}

  <div class="footer">
    <span>Banco de Energia Solidário · Relatório gerado automaticamente</span>
    <span>Período: ${rel.periodo} · ${data}</span>
  </div>
</body>
</html>`;

  const janela = window.open("", "_blank");
  janela.document.write(html);
  janela.document.close();
  setTimeout(() => janela.print(), 500);
}

function TabRelatorios({ relatorios }) {
  const [selecionado, setSelecionado] = useState(null);

  const exportarExcel = (rel) => {
    const rows = [["Produtor CPE", "Produtor Nome", "Beneficiário CPE", "Beneficiário Nome", "kW Atribuídos", "Distância (km)"]];
    rel.linhas.forEach(l => rows.push([l.prodCPE, l.prodNome, l.benCPE, l.benNome, parseFloat(l.kw).toFixed(2), l.distKm]));
    if (Object.keys(rel.naoAlocado || {}).length) {
      rows.push([]); rows.push(["--- kW NÃO ALOCADOS ---"]);
      (rel.produtores || []).forEach(p => { if (rel.naoAlocado[p.id]) rows.push([p.cpe, p.nome, "", "", parseFloat(rel.naoAlocado[p.id]).toFixed(2), ""]); });
    }
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Redistribuição");
    XLSX.writeFile(wb, `relatorio_banco_energia_${rel.periodo}.xlsx`);
  };

  const agrupados = {};
  relatorios.forEach(r => {
    const ano = r.periodo?.slice(0, 4) || "?";
    if (!agrupados[ano]) agrupados[ano] = [];
    agrupados[ano].push(r);
  });

  return (
    <div>
      {selecionado && (
        <Modal title={`Relatório — ${selecionado.periodo}`} onClose={() => setSelecionado(null)}>
          <div style={{ marginBottom: 16, display: "flex", gap: 12, alignItems: "center", flexWrap: "wrap" }}>
            <button style={S.btn("green")} onClick={() => exportarExcel(selecionado)}>⬇ Exportar Excel</button>
            <button style={S.btn("primary")} onClick={() => gerarPDF(selecionado)}>📄 Gerar PDF</button>
            {selecionado.limiteFreguesia && <span style={S.badge("green")}>Freguesia: {selecionado.limiteFreguesia}</span>}
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 12, marginBottom: 20 }}>
            <div style={S.statBox}><div style={S.statNum}>{selecionado.linhas.length}</div><div style={S.statLabel}>Atribuições</div></div>
            <div style={S.statBox}><div style={S.statNum}>{selecionado.linhas.reduce((s, l) => s + l.kw, 0).toFixed(1)}</div><div style={S.statLabel}>kW Distribuídos</div></div>
            <div style={S.statBox}><div style={{ ...S.statNum, color: "#e05c5c" }}>{Object.values(selecionado.naoAlocado || {}).reduce((s, v) => s + v, 0).toFixed(1)}</div><div style={S.statLabel}>kW Não Alocados</div></div>
          </div>
          <table style={S.table}>
            <thead><tr><th style={S.th}>Produtor CPE</th><th style={S.th}>Beneficiário CPE</th><th style={S.th}>Nome</th><th style={S.th}>kW</th></tr></thead>
            <tbody>{selecionado.linhas.map((l, i) => (<tr key={i}><td style={S.td}><span style={S.badge("orange")}>{l.prodCPE}</span></td><td style={S.td}><span style={S.badge("green")}>{l.benCPE}</span></td><td style={S.td}>{l.benNome}</td><td style={S.td}>{parseFloat(l.kw).toFixed(2)}</td></tr>))}</tbody>
          </table>
        </Modal>
      )}
      <div style={S.card}>
        <div style={S.cardTitle}>Histórico de Relatórios ({relatorios.length})</div>
        {relatorios.length === 0 ? <div style={S.empty}>📄 Nenhum relatório gerado ainda</div> : (
          <table style={S.table}>
            <thead><tr><th style={S.th}>Período</th><th style={S.th}>Data execução</th><th style={S.th}>Freguesia</th><th style={S.th}>Atribuições</th><th style={S.th}>kW Distribuídos</th><th style={S.th}>kW Não Alocados</th><th style={S.th}></th></tr></thead>
            <tbody>
              {relatorios.sort((a, b) => b.data > a.data ? 1 : -1).map(r => (
                <tr key={r.id}>
                  <td style={S.td}><span style={S.badge("blue")}>{r.periodo}</span></td>
                  <td style={{ ...S.td, fontSize: 11, color: "#7a9e8e" }}>{new Date(r.data).toLocaleString("pt-PT")}</td>
                  <td style={S.td}>{r.limiteFreguesia ? <span style={S.badge("green")}>{r.limiteFreguesia}</span> : <span style={{ color: "#b7ddd0", fontSize: 11 }}>—</span>}</td>
                  <td style={S.td}>{r.linhas.length}</td>
                  <td style={S.td}>{r.linhas.reduce((s, l) => s + l.kw, 0).toFixed(1)} kW</td>
                  <td style={S.td}>{Object.values(r.naoAlocado || {}).reduce((s, v) => s + v, 0) > 0 ? <span style={S.badge("orange")}>{Object.values(r.naoAlocado).reduce((s, v) => s + v, 0).toFixed(1)} kW</span> : <span style={S.badge("green")}>0 kW</span>}</td>
                  <td style={S.td}>
                    <div style={{ display: "flex", gap: 8 }}>
                      <button style={S.btn("ghost")} onClick={() => setSelecionado(r)}>Ver</button>
                      <button style={S.btn("green")} onClick={() => exportarExcel(r)}>Excel</button>
                      <button style={S.btn("primary")} onClick={() => gerarPDF(r)}>PDF</button>
                      <button style={S.btn("danger")} onClick={async () => { if (confirm("Eliminar relatório?")) await deleteDoc(doc(db, "relatorios", r.id)); }}>✕</button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>
      {Object.keys(agrupados).length > 0 && (
        <div style={S.card}>
          <div style={S.cardTitle}>Resumo Estatístico Anual</div>
          {Object.entries(agrupados).sort(([a], [b]) => b - a).map(([ano, dados]) => {
            const totalKw = dados.reduce((s, r) => s + r.linhas.reduce((ss, l) => ss + l.kw, 0), 0);
            const totalNao = dados.reduce((s, r) => s + Object.values(r.naoAlocado || {}).reduce((ss, v) => ss + v, 0), 0);
            return (
              <div key={ano} style={{ marginBottom: 20 }}>
                <div style={{ fontSize: 13, color: "#2d6a4f", fontWeight: 700, marginBottom: 12 }}>Ano {ano}</div>
                <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 12 }}>
                  <div style={S.statBox}><div style={S.statNum}>{dados.length}</div><div style={S.statLabel}>Meses</div></div>
                  <div style={S.statBox}><div style={S.statNum}>{dados.reduce((s, r) => s + r.linhas.length, 0)}</div><div style={S.statLabel}>Atribuições</div></div>
                  <div style={S.statBox}><div style={S.statNum}>{totalKw.toFixed(0)}</div><div style={S.statLabel}>kW Distribuídos</div></div>
                  <div style={S.statBox}><div style={{ ...S.statNum, color: totalNao > 0 ? "#e05c5c" : "#2d6a4f" }}>{totalNao.toFixed(0)}</div><div style={S.statLabel}>kW Não Alocados</div></div>
                </div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}


// ─── Aba Dashboard ────────────────────────────────────────────────────────────
function TabDashboard({ produtores, beneficiarios, relatorios, config, polygonFeature }) {
  const limFrg = polygonFeature?.properties?.freguesia;

  // Estatísticas gerais
  const totalProdutores = limFrg ? produtores.filter(p => p.freguesia === limFrg).length : produtores.length;
  const totalBeneficiarios = limFrg ? beneficiarios.filter(b => b.freguesia === limFrg).length : beneficiarios.length;
  const totalMembros = beneficiarios.reduce((s, b) => s + (b.membros?.length || 0), 0);
  const totalPotencia = produtores.reduce((s, p) => s + (parseFloat(p.potencia) || 0), 0);

  // Estatísticas dos relatórios
  const totalKwDistribuidos = relatorios.reduce((s, r) => s + r.linhas.reduce((ss, l) => ss + l.kw, 0), 0);
  const totalKwNaoAlocados = relatorios.reduce((s, r) => s + Object.values(r.naoAlocado || {}).reduce((ss, v) => ss + v, 0), 0);
  const totalAtribuicoes = relatorios.reduce((s, r) => s + r.linhas.length, 0);
  const ultimoRelatorio = relatorios.sort((a, b) => b.data > a.data ? 1 : -1)[0];

  // Dados para gráfico mensal (últimos 6 meses)
  const dadosMensais = [];
  const hoje = new Date();
  for (let i = 5; i >= 0; i--) {
    const d = new Date(hoje.getFullYear(), hoje.getMonth() - i, 1);
    const periodo = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
    const rel = relatorios.filter(r => r.periodo === periodo);
    const kw = rel.reduce((s, r) => s + r.linhas.reduce((ss, l) => ss + l.kw, 0), 0);
    const nao = rel.reduce((s, r) => s + Object.values(r.naoAlocado || {}).reduce((ss, v) => ss + v, 0), 0);
    dadosMensais.push({ periodo: periodo.slice(5) + "/" + periodo.slice(2, 4), kw: Math.round(kw), nao: Math.round(nao) });
  }

  const maxKw = Math.max(...dadosMensais.map(d => d.kw + d.nao), 1);

  // Taxa de satisfação do último mês
  const taxaSatisfacao = ultimoRelatorio ? (
    (ultimoRelatorio.linhas.reduce((s, l) => s + l.kw, 0) /
    (ultimoRelatorio.linhas.reduce((s, l) => s + l.kw, 0) + Object.values(ultimoRelatorio.naoAlocado || {}).reduce((s, v) => s + v, 0))) * 100
  ).toFixed(0) : 0;

  return (
    <div>
      {/* Cabeçalho */}
      <div style={{ marginBottom: 24 }}>
        <div style={{ fontSize: 18, fontWeight: 700, color: "#2d6a4f", marginBottom: 4 }}>
          Banco de Energia Solidário
        </div>
        <div style={{ fontSize: 13, color: "#7a9e8e", fontStyle: "italic" }}>
          {limFrg ? `${limFrg} · ` : ""}Visão geral da comunidade de energia renovável
        </div>
      </div>

      {/* Estatísticas principais */}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 16, marginBottom: 20 }}>
        <div style={{ ...S.card, marginBottom: 0, borderLeft: "4px solid #f5a623", padding: 20 }}>
          <div style={{ fontSize: 11, color: "#7a9e8e", textTransform: "uppercase", letterSpacing: 1, marginBottom: 8 }}>Produtores Ativos</div>
          <div style={{ fontSize: 36, fontWeight: 700, color: "#f5a623" }}>{totalProdutores}</div>
          <div style={{ fontSize: 11, color: "#7a9e8e", marginTop: 4 }}>{totalPotencia.toFixed(1)} kW instalados</div>
        </div>
        <div style={{ ...S.card, marginBottom: 0, borderLeft: "4px solid #2d6a4f", padding: 20 }}>
          <div style={{ fontSize: 11, color: "#7a9e8e", textTransform: "uppercase", letterSpacing: 1, marginBottom: 8 }}>Beneficiários</div>
          <div style={{ fontSize: 36, fontWeight: 700, color: "#2d6a4f" }}>{totalBeneficiarios}</div>
          <div style={{ fontSize: 11, color: "#7a9e8e", marginTop: 4 }}>{totalMembros} membros de agregado</div>
        </div>
        <div style={{ ...S.card, marginBottom: 0, borderLeft: "4px solid #40916c", padding: 20 }}>
          <div style={{ fontSize: 11, color: "#7a9e8e", textTransform: "uppercase", letterSpacing: 1, marginBottom: 8 }}>kW Distribuídos</div>
          <div style={{ fontSize: 36, fontWeight: 700, color: "#40916c" }}>{totalKwDistribuidos.toFixed(0)}</div>
          <div style={{ fontSize: 11, color: "#7a9e8e", marginTop: 4 }}>{totalAtribuicoes} atribuições totais</div>
        </div>
        <div style={{ ...S.card, marginBottom: 0, borderLeft: "4px solid #4a90d9", padding: 20 }}>
          <div style={{ fontSize: 11, color: "#7a9e8e", textTransform: "uppercase", letterSpacing: 1, marginBottom: 8 }}>Taxa de Satisfação</div>
          <div style={{ fontSize: 36, fontWeight: 700, color: "#4a90d9" }}>{taxaSatisfacao}%</div>
          <div style={{ fontSize: 11, color: "#7a9e8e", marginTop: 4 }}>último mês processado</div>
        </div>
      </div>

      {/* Gráfico mensal */}
      <div style={S.card}>
        <div style={S.cardTitle}>📊 Evolução Mensal — kW Distribuídos</div>
        {relatorios.length === 0 ? (
          <div style={S.empty}>Sem dados de relatórios ainda</div>
        ) : (
          <div>
            <div style={{ display: "flex", alignItems: "flex-end", gap: 12, height: 180, padding: "0 8px", marginBottom: 8 }}>
              {dadosMensais.map((d, i) => (
                <div key={i} style={{ flex: 1, display: "flex", flexDirection: "column", alignItems: "center", gap: 4 }}>
                  <div style={{ fontSize: 10, color: "#2d6a4f", fontWeight: 700 }}>{d.kw > 0 ? d.kw : ""}</div>
                  <div style={{ width: "100%", display: "flex", flexDirection: "column", gap: 2 }}>
                    {d.nao > 0 && (
                      <div style={{ width: "100%", height: Math.max((d.nao / maxKw) * 140, 2), background: "#fde8e8", borderRadius: "4px 4px 0 0", border: "1px solid #f0b0b0" }} title={`Não alocado: ${d.nao} kW`} />
                    )}
                    <div style={{ width: "100%", height: Math.max((d.kw / maxKw) * 140, d.kw > 0 ? 4 : 0), background: d.kw > 0 ? "linear-gradient(to top, #2d6a4f, #40916c)" : "#e8f5ef", borderRadius: d.nao > 0 ? "0" : "4px 4px 0 0", border: d.kw > 0 ? "none" : "1px dashed #d8ede6" }} title={`Distribuído: ${d.kw} kW`} />
                  </div>
                  <div style={{ fontSize: 10, color: "#7a9e8e" }}>{d.periodo}</div>
                </div>
              ))}
            </div>
            <div style={{ display: "flex", gap: 20, fontSize: 11, color: "#7a9e8e", paddingLeft: 8 }}>
              <span style={{ display: "flex", alignItems: "center", gap: 6 }}><span style={{ width: 12, height: 12, background: "linear-gradient(to top, #2d6a4f, #40916c)", borderRadius: 2, display: "inline-block" }} /> kW Distribuídos</span>
              <span style={{ display: "flex", alignItems: "center", gap: 6 }}><span style={{ width: 12, height: 12, background: "#fde8e8", border: "1px solid #f0b0b0", borderRadius: 2, display: "inline-block" }} /> kW Não Alocados</span>
            </div>
          </div>
        )}
      </div>

      {/* Último relatório + Beneficiários com mais kW */}
      <div style={S.grid2}>
        <div style={S.card}>
          <div style={S.cardTitle}>📋 Último Relatório</div>
          {!ultimoRelatorio ? (
            <div style={S.empty}>Sem relatórios</div>
          ) : (
            <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <span style={{ fontSize: 12, color: "#7a9e8e" }}>Período</span>
                <span style={S.badge("blue")}>{ultimoRelatorio.periodo}</span>
              </div>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <span style={{ fontSize: 12, color: "#7a9e8e" }}>kW distribuídos</span>
                <span style={{ fontSize: 14, fontWeight: 700, color: "#2d6a4f" }}>{ultimoRelatorio.linhas.reduce((s, l) => s + l.kw, 0).toFixed(1)} kW</span>
              </div>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <span style={{ fontSize: 12, color: "#7a9e8e" }}>kW não alocados</span>
                <span style={{ fontSize: 14, fontWeight: 700, color: "#e05c5c" }}>{Object.values(ultimoRelatorio.naoAlocado || {}).reduce((s, v) => s + v, 0).toFixed(1)} kW</span>
              </div>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <span style={{ fontSize: 12, color: "#7a9e8e" }}>Famílias servidas</span>
                <span style={{ fontSize: 14, fontWeight: 700, color: "#2d6a4f" }}>{new Set(ultimoRelatorio.linhas.map(l => l.benId)).size}</span>
              </div>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <span style={{ fontSize: 12, color: "#7a9e8e" }}>Data</span>
                <span style={{ fontSize: 11, color: "#7a9e8e" }}>{new Date(ultimoRelatorio.data).toLocaleDateString("pt-PT")}</span>
              </div>
            </div>
          )}
        </div>

        <div style={S.card}>
          <div style={S.cardTitle}>⚡ Direito por Agregado</div>
          {beneficiarios.length === 0 ? (
            <div style={S.empty}>Sem beneficiários</div>
          ) : (
            <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
              {beneficiarios.slice(0, 5).map(b => {
                const kwDireito = (b.membros?.length || 1) * (parseFloat(config.kwPorMembro) || 50);
                const kwRecebido = ultimoRelatorio ? (ultimoRelatorio.linhas.filter(l => l.benId === b.id).reduce((s, l) => s + l.kw, 0)) : 0;
                const pct = Math.min((kwRecebido / kwDireito) * 100, 100);
                return (
                  <div key={b.id}>
                    <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 4 }}>
                      <span style={{ fontSize: 12, color: "#3d5a4e" }}>{b.nome}</span>
                      <span style={{ fontSize: 11, color: "#7a9e8e" }}>{kwRecebido.toFixed(0)}/{kwDireito} kW</span>
                    </div>
                    <div style={{ height: 8, background: "#e8f5ef", borderRadius: 4, overflow: "hidden" }}>
                      <div style={{ height: "100%", width: `${pct}%`, background: pct >= 100 ? "#2d6a4f" : pct > 50 ? "#40916c" : "#f5a623", borderRadius: 4, transition: "width 0.5s" }} />
                    </div>
                  </div>
                );
              })}
              {beneficiarios.length > 5 && <div style={{ fontSize: 11, color: "#7a9e8e", fontStyle: "italic", textAlign: "center" }}>+{beneficiarios.length - 5} beneficiários</div>}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}


// ─── Aba Simulador ────────────────────────────────────────────────────────────
function TabSimulador({ produtores, beneficiarios, config, polygonFeature }) {
  const [kwSim, setKwSim] = useState({});
  const [resultado, setResultado] = useState(null);
  const [metodoSim, setMetodoSim] = useState(config.metodoRedistribuicao || "proporcional");
  const [kwPorMembroSim, setKwPorMembroSim] = useState(config.kwPorMembro || 50);
  const [raioPadraoSim, setRaioPadraoSim] = useState(config.raioPadrao || 3);
  const [verMapa, setVerMapa] = useState(false);

  const limFrg = polygonFeature?.properties?.freguesia;
  const produtoresValidos = limFrg ? produtores.filter(p => p.freguesia === limFrg && p.lat && p.lon) : produtores.filter(p => p.lat && p.lon);
  const beneficiariosValidos = limFrg ? beneficiarios.filter(b => b.freguesia === limFrg && b.lat && b.lon) : beneficiarios.filter(b => b.lat && b.lon);

  const pontosMapa = [
    ...produtoresValidos.map(p => ({ ...p, tipo: "produtor" })),
    ...beneficiariosValidos.map(b => ({ ...b, tipo: "beneficiario" })),
  ];

  const simularTudo = () => {
    const kw = {};
    produtoresValidos.forEach(p => { kw[p.id] = parseFloat(p.potencia) || 0; });
    setKwSim(kw);
  };

  const executarSimulacao = () => {
    const configSim = { ...config, kwPorMembro: kwPorMembroSim, raioPadrao: raioPadraoSim, metodoRedistribuicao: metodoSim };
    const res = executarRedistribuicao(produtores, beneficiarios, configSim, kwSim, polygonFeature);
    setResultado(res);
  };

  const limpar = () => { setKwSim({}); setResultado(null); };

  const taxaSatisfacao = resultado ? (
    resultado.beneficiariosValidos.length > 0
      ? (resultado.beneficiariosValidos.filter(b => (resultado.recebido[b.id] || 0) >= (resultado.direito[b.id] || 1)).length / resultado.beneficiariosValidos.length * 100).toFixed(0)
      : 0
  ) : null;

  return (
    <div>
      {/* Aviso simulador */}
      <div style={{ ...S.alert("info"), marginBottom: 20 }}>
        🧪 <b>Modo Simulação</b> — Os resultados aqui não são guardados nem afetam os dados reais. Experimenta diferentes cenários antes de executar a redistribuição oficial.
      </div>

      {/* Parâmetros de simulação */}
      <div style={S.card}>
        <div style={S.cardTitle}>⚙ Parâmetros da Simulação</div>
        <div style={S.grid3}>
          <div style={S.field}>
            <label style={S.label}>kW por membro de agregado</label>
            <input style={S.input} type="number" value={kwPorMembroSim} onChange={e => setKwPorMembroSim(e.target.value)} />
            <div style={{ fontSize: 11, color: "#7a9e8e", marginTop: 4, fontStyle: "italic" }}>Configuração atual: {config.kwPorMembro} kW</div>
          </div>
          <div style={S.field}>
            <label style={S.label}>Raio padrão (km)</label>
            <input style={S.input} type="number" value={raioPadraoSim} onChange={e => setRaioPadraoSim(e.target.value)} />
            <div style={{ fontSize: 11, color: "#7a9e8e", marginTop: 4, fontStyle: "italic" }}>Configuração atual: {config.raioPadrao} km</div>
          </div>
          <div style={S.field}>
            <label style={S.label}>Método de redistribuição</label>
            <select style={S.select} value={metodoSim} onChange={e => setMetodoSim(e.target.value)}>
              <option value="proporcional">Proporcional ao agregado</option>
              <option value="satisfacao">Satisfação completa por distância</option>
              <option value="igualitario">Igualitário</option>
            </select>
          </div>
        </div>
      </div>

      {/* kW disponíveis */}
      <div style={S.card}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
          <div style={{ ...S.cardTitle, marginBottom: 0 }}>☀ kW Disponíveis por Produtor</div>
          <div style={{ display: "flex", gap: 10 }}>
            <button style={S.btn("ghost")} onClick={simularTudo}>Usar potência instalada</button>
            <button style={S.btn("ghost")} onClick={limpar}>Limpar</button>
          </div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 12, padding: "12px 16px", background: "#f0faf5", borderRadius: 8, marginBottom: 16, border: "1px solid #c8e6d8" }}>
          <div style={{ fontSize: 12, color: "#2d6a4f", fontWeight: 600, whiteSpace: "nowrap" }}>Aplicar a todos:</div>
          <input
            style={{ ...S.input, width: 120 }}
            type="number"
            placeholder="kW para simular"
            id="kwGlobal"
          />
          <button style={S.btn("primary")} onClick={() => {
            const val = document.getElementById("kwGlobal").value;
            if (!val) return;
            const kw = {};
            produtoresValidos.forEach(p => { kw[p.id] = val; });
            setKwSim(kw);
          }}>Aplicar</button>
          <div style={{ fontSize: 11, color: "#7a9e8e", fontStyle: "italic" }}>Podes ajustar cada produtor individualmente depois</div>
        </div>
        {produtoresValidos.length === 0 ? <div style={S.empty}>Sem produtores elegíveis</div> : (
          <table style={S.table}>
            <thead><tr><th style={S.th}>CPE</th><th style={S.th}>Nome</th><th style={S.th}>Pot. Instalada</th><th style={S.th}>kW para simular</th></tr></thead>
            <tbody>
              {produtoresValidos.map(p => (
                <tr key={p.id}>
                  <td style={S.td}><span style={S.badge("orange")}>{p.cpe}</span></td>
                  <td style={S.td}>{p.nome}</td>
                  <td style={{ ...S.td, color: "#7a9e8e" }}>{p.potencia || "—"} kW</td>
                  <td style={S.td}>
                    <input style={{ ...S.input, width: 120 }} type="number" value={kwSim[p.id] || ""} onChange={e => setKwSim(prev => ({ ...prev, [p.id]: e.target.value }))} placeholder="0" />
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>

      <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: 24 }}>
        <button style={S.btn("green")} onClick={executarSimulacao}>🧪 Simular Redistribuição</button>
      </div>

      {/* Resultado da simulação */}
      {resultado && (
        <div>
          <div style={S.card}>
            <div style={S.cardTitle}>📊 Resultado da Simulação</div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 16, marginBottom: 24 }}>
              <div style={S.statBox}><div style={S.statNum}>{resultado.linhas.length}</div><div style={S.statLabel}>Atribuições</div></div>
              <div style={S.statBox}><div style={S.statNum}>{resultado.linhas.reduce((s, l) => s + l.kw, 0).toFixed(1)}</div><div style={S.statLabel}>kW Distribuídos</div></div>
              <div style={S.statBox}><div style={{ ...S.statNum, color: Object.keys(resultado.naoAlocado).length ? "#e05c5c" : "#2d6a4f" }}>{Object.values(resultado.naoAlocado).reduce((s, v) => s + v, 0).toFixed(1)}</div><div style={S.statLabel}>kW Não Alocados</div></div>
              <div style={S.statBox}><div style={{ ...S.statNum, color: taxaSatisfacao >= 100 ? "#2d6a4f" : taxaSatisfacao >= 50 ? "#f5a623" : "#e05c5c" }}>{taxaSatisfacao}%</div><div style={S.statLabel}>Famílias Satisfeitas</div></div>
            </div>

            {/* Satisfação por beneficiário */}
            <div style={{ marginBottom: 20 }}>
              <div style={{ fontSize: 12, color: "#2d6a4f", fontWeight: 700, marginBottom: 12 }}>Satisfação por Beneficiário</div>
              {resultado.beneficiariosValidos.map(b => {
                const kwR = resultado.recebido[b.id] || 0;
                const kwD = resultado.direito[b.id] || 1;
                const pct = Math.min((kwR / kwD) * 100, 100);
                return (
                  <div key={b.id} style={{ marginBottom: 12 }}>
                    <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 4 }}>
                      <span style={{ fontSize: 12, color: "#3d5a4e" }}>{b.nome} <span style={{ color: "#7a9e8e", fontSize: 11 }}>({b.membros?.length || 1} membro(s))</span></span>
                      <span style={{ fontSize: 12, fontWeight: 700, color: pct >= 100 ? "#2d6a4f" : pct > 0 ? "#f5a623" : "#e05c5c" }}>{kwR.toFixed(1)} / {kwD} kW ({pct.toFixed(0)}%)</span>
                    </div>
                    <div style={{ height: 10, background: "#e8f5ef", borderRadius: 5, overflow: "hidden" }}>
                      <div style={{ height: "100%", width: `${pct}%`, background: pct >= 100 ? "#2d6a4f" : pct > 50 ? "#40916c" : "#f5a623", borderRadius: 5, transition: "width 0.5s" }} />
                    </div>
                  </div>
                );
              })}
            </div>

            <table style={S.table}>
              <thead><tr><th style={S.th}>Produtor CPE</th><th style={S.th}>Beneficiário CPE</th><th style={S.th}>Nome</th><th style={S.th}>kW</th><th style={S.th}>Dist. (km)</th></tr></thead>
              <tbody>
                {resultado.linhas.map((l, i) => (
                  <tr key={i}>
                    <td style={S.td}><span style={S.badge("orange")}>{l.prodCPE}</span></td>
                    <td style={S.td}><span style={S.badge("green")}>{l.benCPE}</span></td>
                    <td style={S.td}>{l.benNome}</td>
                    <td style={S.td}>{l.kw.toFixed(2)}</td>
                    <td style={S.td}>{l.distKm}</td>
                  </tr>
                ))}
              </tbody>
            </table>

            {Object.keys(resultado.naoAlocado).length > 0 && (
              <div style={{ ...S.alert("err"), marginTop: 16 }}>
                ⚠ kW não alocados: {produtores.filter(p => resultado.naoAlocado[p.id]).map(p => `${p.cpe}: ${resultado.naoAlocado[p.id].toFixed(2)} kW`).join(" | ")}
              </div>
            )}
          </div>

          {/* Mapa da simulação */}
          <div style={S.card}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <div style={{ ...S.cardTitle, marginBottom: 0 }}>🗺 Mapa da Simulação</div>
              <button style={S.btn(verMapa ? "primary" : "ghost")} onClick={() => setVerMapa(v => !v)}>
                {verMapa ? "Ocultar" : "Ver Mapa"}
              </button>
            </div>
            {verMapa && (
              <div style={{ marginTop: 20 }}>
                <LeafletMap pontos={pontosMapa} mostrarRaios={true} config={{ ...config, raioPadrao: raioPadraoSim }} polygonFeature={polygonFeature} linhasRedistribuicao={resultado.linhas} />
                <div style={{ display: "flex", gap: 20, marginTop: 10, fontSize: 12, color: "#7a9e8e", fontStyle: "italic" }}>
                  <span>☀ Produtores</span>
                  <span>🏠 Beneficiários</span>
                  <span>― Linhas = fluxo simulado</span>
                </div>
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  );
}

// ─── App Principal ────────────────────────────────────────────────────────────
const TABS = ["Dashboard", "Produtores", "Beneficiários", "Configurações", "Redistribuição", "Simulador", "Relatórios"];
const DEFAULT_CONFIG = { kwPorMembro: 50, raioPadrao: 3, raiosProdutores: {}, metodoRedistribuicao: "proporcional" };

export default function App() {
  const [tab, setTab] = useState(0);
  const [produtores, setProdutores] = useState([]);
  const [beneficiarios, setBeneficiarios] = useState([]);
  const [config, setConfig] = useState(DEFAULT_CONFIG);
  const [relatorios, setRelatorios] = useState([]);
  const [loaded, setLoaded] = useState(false);
  const [polygonFeature, setPolygonFeature] = useState(null);
  const [listaFreguesias, setListaFreguesias] = useState([]);

  useEffect(() => {
    const unsubs = [
      onSnapshot(collection(db, "produtores"), snap => setProdutores(snap.docs.map(d => d.data()))),
      onSnapshot(collection(db, "beneficiarios"), snap => setBeneficiarios(snap.docs.map(d => d.data()))),
      onSnapshot(collection(db, "relatorios"), snap => setRelatorios(snap.docs.map(d => d.data()))),
    ];
    onSnapshot(doc(db, "config", "global"), d => { if (d.exists()) setConfig(d.data()); });
    getDoc(doc(db, "config", "poligono")).then(d => {
      if (d.exists() && d.data().feature) {
        try { setPolygonFeature(JSON.parse(d.data().feature)); } catch {}
      }
    });
    getDoc(doc(db, "config", "listaFreguesias")).then(d => {
      if (d.exists() && d.data().lista) {
        try {
          const lista = JSON.parse(d.data().lista);
          setListaFreguesias(lista);
        } catch(e) { console.error("Erro ao carregar lista de freguesias:", e); }
      } else {
        console.log("Lista de freguesias não encontrada no Firebase");
      }
      setLoaded(true);
    }).catch(e => {
      console.error("Erro Firebase listaFreguesias:", e);
      setLoaded(true);
    });
    return () => unsubs.forEach(u => u());
  }, []);

  if (!loaded) return (
    <div style={{ ...S.app, alignItems: "center", justifyContent: "center", gap: 16 }}>
      <div style={{ fontSize: 40 }}>☀</div>
      <div style={{ color: "#7a9e8e", fontSize: 14, fontStyle: "italic" }}>A ligar ao servidor...</div>
    </div>
  );

  return (
    <div style={S.app}>
      <div style={S.header}>
        <div style={S.logo}>☀</div>
        <div>
          <div style={S.title}>Banco de Energia Solidário</div>
          <div style={S.subtitle}>
            Comunidade de Energia Renovável
            {config.limiteFreguesia && ` · ${config.limiteFreguesia}`}
          </div>
        </div>
        <div style={{ marginLeft: "auto", fontSize: 11, color: "rgba(255,255,255,0.8)" }}>● ligado</div>
      </div>
      <div style={S.nav}>
        {TABS.map((t, i) => <button key={t} style={S.navBtn(tab === i)} onClick={() => setTab(i)}>{t}</button>)}
      </div>
      <div style={S.main}>
        {tab === 0 && <TabDashboard produtores={produtores} beneficiarios={beneficiarios} relatorios={relatorios} config={config} polygonFeature={polygonFeature} />}
        {tab === 1 && <TabProdutores produtores={produtores} polygonFeature={polygonFeature} listaFreguesias={listaFreguesias} />}
        {tab === 2 && <TabBeneficiarios beneficiarios={beneficiarios} polygonFeature={polygonFeature} listaFreguesias={listaFreguesias} />}
        {tab === 3 && <TabConfiguracoes config={config} produtores={produtores} beneficiarios={beneficiarios} polygonFeature={polygonFeature} setPolygonFeature={setPolygonFeature} setListaFreguesias={setListaFreguesias} />}
        {tab === 4 && <TabRedistribuicao produtores={produtores} beneficiarios={beneficiarios} config={config} polygonFeature={polygonFeature} />}
        {tab === 5 && <TabSimulador produtores={produtores} beneficiarios={beneficiarios} config={config} polygonFeature={polygonFeature} />}
        {tab === 6 && <TabRelatorios relatorios={relatorios} />}
      </div>
    </div>
  );
}
