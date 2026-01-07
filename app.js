/* =========================================================
   TASATOP — Cronograma de Inversión (Replica VBA)
   - Sin frameworks
   - Lógica financiera separada (funciones puras)
   - UI/DOM separado
   - Export PDF con logo embebido (dataURL)
========================================================= */

/** ===== CONFIG ===== */
const LOGO_URL = "https://tasatop.com.pe/wp-content/uploads/elementor/thumbs/logos-17-r320c27cra7m7te2fafiia4mrbqd3aqj7ifttvy33g.png";
const TASA_IMPUESTO_2DA = 0.05;

/** ===== DOM ===== */
const elLogo = document.getElementById("brandLogo");
const elGeneratedAt = document.getElementById("generatedAt");

const form = document.getElementById("form");
const btnLimpiar = document.getElementById("btnLimpiar");
const btnPdf = document.getElementById("btnPdf");

const errorSummary = document.getElementById("errorSummary");
const errorSummaryList = document.getElementById("errorSummaryList");
const logoWarning = document.getElementById("logoWarning");

const resumen = document.getElementById("resumen");
const tableWrap = document.getElementById("tableWrap");
const tbody = document.getElementById("tbody");
const tfoot = document.getElementById("tfoot");
const emptyState = document.getElementById("emptyState");

const sumTA = document.getElementById("sumTA");
const sumMonto = document.getElementById("sumMonto");
const sumProducto = document.getElementById("sumProducto");
const sumPlazo = document.getElementById("sumPlazo");
const sumFreqInt = document.getElementById("sumFreqInt");
const sumFreqCap = document.getElementById("sumFreqCap");

const thMontoBase = document.getElementById("thMontoBase");
const thIntBruto = document.getElementById("thIntBruto");
const thImpuesto = document.getElementById("thImpuesto");
const thIntDep = document.getElementById("thIntDep");
const thDevCap = document.getElementById("thDevCap");
const thSaldo = document.getElementById("thSaldo");
const thTotal = document.getElementById("thTotal");

/** ===== STATE ===== */
let lastResult = null; // { inputs, rows, totals, generatedAtISO, logoDataUrl }

/** ===== INIT ===== */
setGeneratedNow();
loadLogoToUI();

/* =========================================================
   UTILIDADES — replicando VBA
========================================================= */

/** MonedaSimbolo(ByVal moneda As String) */
function monedaSimbolo(moneda) {
  const m = String(moneda ?? "").trim().toUpperCase();
  if (m === "") return "S/.";
  if (m.includes("$") || m.includes("USD") || m.includes("DOL")) return "$";
  if (m.includes("S/") || m.includes("SOL")) return "S/.";
  return String(moneda ?? "").trim(); // si viene raro, lo respeta
}

/** NormalizarClave: mayúsculas + sin tildes */
function normalizarClave(s) {
  let x = String(s ?? "").trim().toUpperCase();

  const map = {
    "Á":"A","É":"E","Í":"I","Ó":"O","Ú":"U",
    "Ü":"U","Ñ":"N"
  };
  x = x.replace(/[ÁÉÍÓÚÜÑ]/g, (ch) => map[ch] || ch);

  while (x.includes("  ")) x = x.replace(/  /g, " ");
  return x;
}

/** ObtenerDiaPago(producto) */
function obtenerDiaPago(producto) {
  switch (normalizarClave(producto)) {
    case "IKB": return 15;
    case "ALI": return 28;
    case "PET": return 10;
    case "M&L": return 20;
    default: return 15;
  }
}

/** FrecuenciaAMeses(frecuencia, plazoMeses) */
function frecuenciaAMeses(frecuencia, plazoMeses) {
  switch (normalizarClave(frecuencia)) {
    case "MENSUAL": return 1;
    case "BIMESTRAL": return 2;
    case "TRIMESTRAL": return 3;
    case "SEMESTRAL": return 6;
    case "ANUAL": return 12;
    case "AL FINALIZAR": return Number(plazoMeses);
    default: return 1;
  }
}

/** Excel/VBA Round(x,2): banker’s rounding (ties to even) */
function vbaRound(value, decimals = 0) {
  // Maneja negativos y casos tie 0.5 a par
  const factor = Math.pow(10, decimals);
  const x = value * factor;

  if (!isFinite(x)) return value;

  const sign = x < 0 ? -1 : 1;
  const ax = Math.abs(x);

  const floor = Math.floor(ax);
  const diff = ax - floor;

  // tolerancia por floating
  const EPS = 1e-12;

  let rounded;
  if (diff > 0.5 + EPS) {
    rounded = floor + 1;
  } else if (diff < 0.5 - EPS) {
    rounded = floor;
  } else {
    // tie: 0.5 exacto -> al par
    rounded = (floor % 2 === 0) ? floor : floor + 1;
  }

  return (sign * rounded) / factor;
}

/** Parse YYYY-MM-DD (input type=date) -> Date local (sin TZ issues) */
function parseDateInput(val) {
  if (!val) return null;
  const [y, m, d] = val.split("-").map(Number);
  if (!y || !m || !d) return null;
  return new Date(y, m - 1, d, 0, 0, 0, 0);
}

/** DateDiff("d", a, b) — igual a VBA para fechas (b - a) en días */
function dateDiffDays(a, b) {
  const msPerDay = 24 * 60 * 60 * 1000;
  const ua = Date.UTC(a.getFullYear(), a.getMonth(), a.getDate());
  const ub = Date.UTC(b.getFullYear(), b.getMonth(), b.getDate());
  return Math.floor((ub - ua) / msPerDay);
}

/** last day of month */
function lastDayOfMonth(year, monthIndex0) {
  // monthIndex0: 0..11
  return new Date(year, monthIndex0 + 1, 0).getDate();
}

/** FechaPagoMes(fechaBase, mesOffset, diaPago) */
function fechaPagoMes(fechaBase, mesOffset, diaPago) {
  const baseY = fechaBase.getFullYear();
  const baseM = fechaBase.getMonth();
  const target = new Date(baseY, baseM + Number(mesOffset), 1);
  const y = target.getFullYear();
  const m = target.getMonth();
  const ultimo = lastDayOfMonth(y, m);
  const d = Math.min(Number(diaPago), ultimo);
  return new Date(y, m, d, 0, 0, 0, 0);
}

/** CalcularPrimeraFechaPago(fechaInicio, diaPago, opcionPrimerPago) */
function calcularPrimeraFechaPago(fechaInicio, diaPago, opcionPrimerPago) {
  const op = normalizarClave(opcionPrimerPago);

  if (fechaInicio.getDate() > diaPago) {
    return fechaPagoMes(fechaInicio, 1, diaPago);
  }

  if (op.includes("MES") && op.includes("INVERSION")) {
    return fechaPagoMes(fechaInicio, 0, diaPago);
  }
  return fechaPagoMes(fechaInicio, 1, diaPago);
}

/** EsMesDePago(i, freqMeses) */
function esMesDePago(i, freqMeses) {
  if (freqMeses <= 0) return false;
  return (i % freqMeses) === 0;
}

/** Formatos UI */
function pad2(n) { return String(n).padStart(2, "0"); }
function formatDateDMY(d) {
  if (!d) return "";
  return `${pad2(d.getDate())}/${pad2(d.getMonth() + 1)}/${d.getFullYear()}`;
}
function formatMoney(n) {
  // Solo formato visual, lógica usa números puros
  const x = Number(n);
  if (!isFinite(x)) return "--";
  return x.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
function formatPercentEA(tasaEA) {
  // En VBA: Format(tasaEA, "0.000%") (tasaEA ya es decimal)
  const p = tasaEA * 100;
  return `${p.toFixed(3)}%`;
}

/* =========================================================
   LÓGICA FINANCIERA — replica exacta del VBA
========================================================= */
function generarCronograma(inputs) {
  const {
    fechaInicio,
    monto,
    monedaRaw,
    tasaEA_pct,
    plazo,
    producto,
    frecuenciaInteres,
    frecuenciaCapital,
    opcionPrimerPago
  } = inputs;

  const moneda = monedaSimbolo(monedaRaw);
  const tasaEA = Number(tasaEA_pct) / 100;

  const diaPago = obtenerDiaPago(producto);
  const freqIntMeses = frecuenciaAMeses(frecuenciaInteres, plazo);
  const freqCapMeses = frecuenciaAMeses(frecuenciaCapital, plazo);

  let fechaPagoAnterior = calcularPrimeraFechaPago(fechaInicio, diaPago, opcionPrimerPago);

  // Inicialización
  let saldo = Number(monto);
  let pagosInteresCont = 0;
  let pagosCapitalCont = 0;
  let mesesDesdeUltPagoInteres = 0;

  // numPagosCapital
  let numPagosCapital;
  if (normalizarClave(frecuenciaCapital) === "AL FINALIZAR") {
    numPagosCapital = 1;
  } else {
    numPagosCapital = Math.floor((plazo + freqCapMeses - 1) / freqCapMeses);
    if (numPagosCapital < 1) numPagosCapital = 1;
  }
  const amortBase = Number(monto) / numPagosCapital;

  const rows = [];

  for (let i = 0; i <= plazo; i++) {
    const mes = i;

    let pagaInteres;
    let pagaCapital;

    let fechaPago = null;
    let fechaCrono = null;

    // días informativos (VBA lo calcula, pero NO lo imprime en la tabla)
    let diasInfo = 0;

    if (i === 0) {
      fechaCrono = fechaInicio;
      fechaPago = null;
      diasInfo = 0;
    } else {
      if (i === 1) {
        fechaPago = fechaPagoAnterior;
      } else {
        fechaPago = fechaPagoMes(fechaPagoAnterior, 1, diaPago);
        fechaPagoAnterior = fechaPago;
      }

      // REGLA CLAVE: desde mes 1, Fecha Cronograma = Fecha Pago
      fechaCrono = fechaPago;

      if (i === 1) {
        diasInfo = dateDiffDays(fechaInicio, fechaPago);
        if (diasInfo < 0) diasInfo = 0;
      } else {
        const fechaPagoPrev = fechaPagoMes(fechaPago, -1, diaPago);
        diasInfo = dateDiffDays(fechaPagoPrev, fechaPago);
        if (diasInfo <= 0) diasInfo = 30;
      }
    }

    // =========================
    // INTERÉS (solo cuando toca pago)
    // =========================
    if (i === 0) {
      pagaInteres = false;
    } else {
      mesesDesdeUltPagoInteres = mesesDesdeUltPagoInteres + 1;

      if (freqIntMeses === 1) {
        pagaInteres = true;
      } else if (esMesDePago(mes, freqIntMeses) || mes === plazo) {
        pagaInteres = true;
      } else {
        pagaInteres = false;
      }
    }

    let diasInteres = 0;
    let interesMes = 0;
    let interesBrutoPago = 0;
    let impuesto = 0;
    let interesDepositar = 0;

    if (pagaInteres) {
      if (pagosInteresCont === 0) {
        diasInteres = dateDiffDays(fechaInicio, fechaPago);
        if (diasInteres < 0) diasInteres = 0;
      } else {
        diasInteres = 30 * mesesDesdeUltPagoInteres;
      }

      interesBrutoPago = ((Math.pow(1 + tasaEA, diasInteres / 360) - 1) * saldo);
      interesBrutoPago = vbaRound(interesBrutoPago, 2);

      impuesto = vbaRound(interesBrutoPago * TASA_IMPUESTO_2DA, 2);
      interesDepositar = vbaRound(interesBrutoPago - impuesto, 2);

      interesMes = interesBrutoPago;
      pagosInteresCont = pagosInteresCont + 1;
      mesesDesdeUltPagoInteres = 0;
    }

    // =========================
    // CAPITAL
    // =========================
    if (i === 0) {
      pagaCapital = false;
    } else {
      if (normalizarClave(frecuenciaCapital) === "AL FINALIZAR") {
        pagaCapital = (mes === plazo);
      } else {
        pagaCapital = esMesDePago(mes, freqCapMeses) || (mes === plazo);
      }
    }

    let devolucionCapital = 0;

    if (pagaCapital && saldo > 0) {
      pagosCapitalCont = pagosCapitalCont + 1;

      if (mes === plazo || pagosCapitalCont === numPagosCapital) {
        devolucionCapital = saldo;
      } else {
        devolucionCapital = vbaRound(amortBase, 2);
        if (devolucionCapital > saldo) devolucionCapital = saldo;
      }

      saldo = vbaRound(saldo - devolucionCapital, 2);
      if (saldo < 0) saldo = 0;
    }

    const totalDepositar = vbaRound(interesDepositar + devolucionCapital, 2);

    // Monto base (E): Round(saldo + devolucionCapital, 2)
    const montoBase = vbaRound(saldo + devolucionCapital, 2);

    // Columna D (Días) según VBA:
    // i=0 => "--"
    // else => IIf(pagaInteres, diasInteres, "--")
    const diasCol = (i === 0) ? "--" : (pagaInteres ? diasInteres : "--");

    rows.push({
      mes: i,
      fechaCrono,
      fechaPago: (i === 0 ? null : fechaPago),
      dias: diasCol,
      montoBase,
      interesBruto: interesMes,
      impuesto,
      interesDepositar,
      devolucionCapital,
      saldo,
      totalDepositar,

      // No se imprime en la tabla del VBA, pero lo conservamos por trazabilidad
      _diasInfo: diasInfo,
      _pagaInteres: pagaInteres,
      _pagaCapital: pagaCapital
    });
  }

  // Totales (como VBA): sum H, I, K (desde filaIni..ultimaFila incluyendo mes 0)
  const totalInteresDepositar = vbaRound(rows.reduce((acc, r) => acc + Number(r.interesDepositar || 0), 0), 2);
  const totalDevolucionCapital = vbaRound(rows.reduce((acc, r) => acc + Number(r.devolucionCapital || 0), 0), 2);
  const totalTotalDepositar = vbaRound(rows.reduce((acc, r) => acc + Number(r.totalDepositar || 0), 0), 2);

  return {
    moneda,
    tasaEA,
    diaPago,
    freqIntMeses,
    freqCapMeses,
    rows,
    totals: {
      interesDepositar: totalInteresDepositar,
      devolucionCapital: totalDevolucionCapital,
      totalDepositar: totalTotalDepositar
    }
  };
}

/* =========================================================
   VALIDACIONES — equivalentes al VBA (mensajes en español)
========================================================= */
function validateInputs(raw) {
  const errors = {};

  // Fecha válida
  if (!(raw.fechaInicio instanceof Date) || isNaN(raw.fechaInicio.getTime())) {
    errors.fechaInicio = "La fecha de inicio no es una fecha válida.";
  }

  // Monto > 0
  if (raw.monto === "" || raw.monto === null || raw.monto === undefined || !isFinite(Number(raw.monto)) || Number(raw.monto) <= 0) {
    errors.monto = "El monto debe ser numérico y mayor a 0.";
  }

  // Moneda válida (no vacía)
  const mon = monedaSimbolo(raw.monedaRaw);
  if (!String(mon).trim()) {
    errors.moneda = "La moneda está vacía. Usa S/ o $.";
  }

  // Tasa numérica
  if (raw.tasaEA_pct === "" || raw.tasaEA_pct === null || raw.tasaEA_pct === undefined || !isFinite(Number(raw.tasaEA_pct))) {
    errors.tasaEA = "La tasa debe ser numérica.";
  }

  // Plazo > 0
  if (raw.plazo === "" || raw.plazo === null || raw.plazo === undefined || !Number.isFinite(Number(raw.plazo)) || Number(raw.plazo) <= 0) {
    errors.plazo = "El plazo (meses) debe ser numérico y mayor a 0.";
  }

  // Producto no vacío
  if (!String(raw.producto || "").trim()) {
    errors.producto = "El producto está vacío.";
  }

  // Frecuencias no vacías
  if (!String(raw.frecuenciaInteres || "").trim()) {
    errors.freqInteres = "La frecuencia de intereses está vacía.";
  }
  if (!String(raw.frecuenciaCapital || "").trim()) {
    errors.freqCapital = "La devolución de capital está vacía.";
  }

  // Opción primer pago: si vacío => "Próximo mes" (como VBA)
  // (no es error)

  return errors;
}

/* =========================================================
   UI / DOM
========================================================= */
function setGeneratedNow() {
  const now = new Date();
  elGeneratedAt.textContent = formatDateTime(now);
}

function formatDateTime(d) {
  const dd = pad2(d.getDate());
  const mm = pad2(d.getMonth() + 1);
  const yyyy = d.getFullYear();
  const hh = pad2(d.getHours());
  const mi = pad2(d.getMinutes());
  const ss = pad2(d.getSeconds());
  return `${dd}/${mm}/${yyyy} ${hh}:${mi}:${ss}`;
}

function clearFieldErrors() {
  document.querySelectorAll(".field__error").forEach(el => el.textContent = "");
  errorSummary.hidden = true;
  errorSummaryList.innerHTML = "";
}

function showErrors(errors) {
  clearFieldErrors();

  const entries = Object.entries(errors);
  if (entries.length === 0) return;

  // per-field
  for (const [key, msg] of entries) {
    const el = document.querySelector(`[data-error-for="${key}"]`);
    if (el) el.textContent = msg;
  }

  // summary
  errorSummaryList.innerHTML = "";
  for (const [, msg] of entries) {
    const li = document.createElement("li");
    li.textContent = msg;
    errorSummaryList.appendChild(li);
  }
  errorSummary.hidden = false;
}

function readForm() {
  const fechaInicio = parseDateInput(document.getElementById("fechaInicio").value);

  return {
    fechaInicio,
    monto: document.getElementById("monto").value,
    monedaRaw: document.getElementById("moneda").value,
    tasaEA_pct: document.getElementById("tasaEA").value,
    plazo: Number(document.getElementById("plazo").value),
    producto: document.getElementById("producto").value,
    frecuenciaInteres: document.getElementById("freqInteres").value,
    frecuenciaCapital: document.getElementById("freqCapital").value,
    opcionPrimerPago: document.getElementById("opcionPrimerPago").value || "Próximo mes"
  };
}

function renderResult(inputs, result, generatedAt) {
  // Summary card
  resumen.hidden = false;
  sumTA.textContent = formatPercentEA(result.tasaEA);
  sumMonto.textContent = `${result.moneda} ${formatMoney(Number(inputs.monto))}`;
  sumProducto.textContent = inputs.producto;
  sumPlazo.textContent = `${Number(inputs.plazo) * 30} Días`;
  sumFreqInt.textContent = inputs.frecuenciaInteres;
  sumFreqCap.textContent = inputs.frecuenciaCapital;

  // Table headers with currency
  thMontoBase.textContent = `Monto base (${result.moneda})`;
  thIntBruto.textContent = `Interés bruto (${result.moneda})`;
  thImpuesto.textContent = `Impuesto 2da categ. (${result.moneda})`;
  thIntDep.textContent = `Interés a depositar (${result.moneda})`;
  thDevCap.textContent = `Devolución capital (${result.moneda})`;
  thSaldo.textContent = `Saldo capital (${result.moneda})`;
  thTotal.textContent = `Total a depositar (${result.moneda})`;

  // Body
  tbody.innerHTML = "";
  for (const r of result.rows) {
    const tr = document.createElement("tr");

    const cells = [
      r.mes,
      formatDateDMY(r.fechaCrono),
      r.mes === 0 ? "" : formatDateDMY(r.fechaPago),
      r.dias,
      formatMoney(r.montoBase),
      formatMoney(r.interesBruto),
      formatMoney(r.impuesto),
      formatMoney(r.interesDepositar),
      formatMoney(r.devolucionCapital),
      formatMoney(r.saldo),
      formatMoney(r.totalDepositar),
    ];

    for (const c of cells) {
      const td = document.createElement("td");
      td.textContent = String(c);
      tr.appendChild(td);
    }

    tbody.appendChild(tr);
  }

  // Footer totals (como VBA)
  tfoot.innerHTML = "";
  const trTot = document.createElement("tr");

  // 11 columnas: "Total:" en col 1, vacías 2..7, totales en H (8), I (9), K (11)
  const footerCells = new Array(11).fill("");
  footerCells[0] = "Total:";
  footerCells[7] = formatMoney(result.totals.interesDepositar);
  footerCells[8] = formatMoney(result.totals.devolucionCapital);
  footerCells[10] = formatMoney(result.totals.totalDepositar);

  footerCells.forEach((val) => {
    const td = document.createElement("td");
    td.textContent = val;
    trTot.appendChild(td);
  });

  tfoot.appendChild(trTot);

  tableWrap.hidden = false;
  emptyState.hidden = true;
  btnPdf.disabled = false;

  // update generatedAt in header (for the current run)
  elGeneratedAt.textContent = formatDateTime(generatedAt);
}

/* =========================================================
   LOGO: UI + PDF (dataURL)
========================================================= */
async function loadLogoToUI() {
  elLogo.src = LOGO_URL;
  elLogo.loading = "eager";
}

async function fetchLogoAsDataURL() {
  // Intento 1: fetch blob -> FileReader (requiere CORS ok)
  try {
    const res = await fetch(LOGO_URL, { mode: "cors", cache: "no-store" });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const blob = await res.blob();
    const dataUrl = await blobToDataURL(blob);
    return dataUrl;
  } catch (_) {
    // Intento 2: usar <img> + canvas (también depende CORS)
    try {
      const dataUrl = await imgUrlToDataURLViaCanvas(LOGO_URL);
      return dataUrl;
    } catch (e2) {
      return null;
    }
  }
}

function blobToDataURL(blob) {
  return new Promise((resolve, reject) => {
    const fr = new FileReader();
    fr.onload = () => resolve(fr.result);
    fr.onerror = reject;
    fr.readAsDataURL(blob);
  });
}

function imgUrlToDataURLViaCanvas(url) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.crossOrigin = "anonymous";
    img.onload = () => {
      try{
        const canvas = document.createElement("canvas");
        canvas.width = img.naturalWidth || img.width;
        canvas.height = img.naturalHeight || img.height;
        const ctx = canvas.getContext("2d");
        ctx.drawImage(img, 0, 0);
        const dataUrl = canvas.toDataURL("image/png");
        resolve(dataUrl);
      } catch (e) {
        reject(e);
      }
    };
    img.onerror = reject;
    img.src = url;
  });
}

/* =========================================================
   PDF EXPORT (A4 horizontal, membrete + resumen + tabla)
========================================================= */
function buildPdfFileName(d) {
  const yyyy = d.getFullYear();
  const mm = pad2(d.getMonth() + 1);
  const dd = pad2(d.getDate());
  const hh = pad2(d.getHours());
  const mi = pad2(d.getMinutes());
  const ss = pad2(d.getSeconds());
  return `TASATOP_Cronograma_${yyyy}${mm}${dd}_${hh}${mi}${ss}.pdf`;
}

async function exportPDF(state) {
  const { inputs, result, generatedAt, logoDataUrl } = state;

  const jspdf = window.jspdf;
  if (!jspdf || !jspdf.jsPDF) {
    alert("No se pudo cargar jsPDF. Revisa tu conexión a internet (solo para cargar el CDN).");
    return;
  }

  const doc = new jspdf.jsPDF({ orientation: "landscape", unit: "pt", format: "a4" });
  const pageW = doc.internal.pageSize.getWidth();
  const margin = 28;

  // Header (membrete)
  const headerY = 28;

  if (logoDataUrl) {
    // logo a la izquierda
    const logoW = 46;
    const logoH = 46;
    doc.addImage(logoDataUrl, "PNG", margin, headerY, logoW, logoH);
  }

  const xText = margin + (logoDataUrl ? 58 : 0);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(14);
  doc.text("TASATOP", xText, headerY + 18);

  doc.setFont("helvetica", "normal");
  doc.setFontSize(11);
  doc.text("Cronograma de Inversión", xText, headerY + 34);

  doc.setFontSize(10);
  doc.setTextColor(80);
  doc.text(`Generado: ${formatDateTime(generatedAt)}`, pageW - margin, headerY + 18, { align: "right" });
  doc.setTextColor(0);

  // Summary "tarjeta"
  const sumTop = headerY + 62;
  const boxH = 74;
  doc.setDrawColor(220);
  doc.setFillColor(15, 23, 42); // oscuro tipo PDF del VBA
  doc.roundedRect(margin, sumTop, pageW - margin * 2, boxH, 10, 10, "F");

  doc.setTextColor(255);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(10);

  const moneda = result.moneda;
  const tasaEA = result.tasaEA;
  const plazoDias = Number(inputs.plazo) * 30;

  // layout de 4 columnas (similar al VBA)
  const colW = (pageW - margin * 2) / 4;
  const row1Y = sumTop + 22;
  const row2Y = sumTop + 50;

  // labels
  doc.text("TA:", margin + colW * 0 + 14, row1Y);
  doc.text("Monto Invertido", margin + colW * 1 + 14, row1Y);
  doc.text("Producto", margin + colW * 2 + 14, row1Y);
  doc.text("Plazo", margin + colW * 3 + 14, row1Y);

  doc.setFont("helvetica", "normal");
  doc.text(formatPercentEA(tasaEA), margin + colW * 0 + 14, row1Y + 14);
  doc.text(`${moneda} ${formatMoney(Number(inputs.monto))}`, margin + colW * 1 + 14, row1Y + 14);
  doc.text(String(inputs.producto), margin + colW * 2 + 14, row1Y + 14);
  doc.text(`${plazoDias} Días`, margin + colW * 3 + 14, row1Y + 14);

  doc.setFont("helvetica", "bold");
  doc.text("Frecuencia", margin + colW * 0 + 14, row2Y);
  doc.text("Tipo tasa", margin + colW * 2 + 14, row2Y);
  doc.text("Devolución de capital", margin + colW * 3 + 14, row2Y);

  doc.setFont("helvetica", "normal");
  doc.text(String(inputs.frecuenciaInteres), margin + colW * 0 + 14, row2Y + 14);
  doc.text("Tasa Efectiva Anual", margin + colW * 2 + 14, row2Y + 14);
  doc.text(String(inputs.frecuenciaCapital), margin + colW * 3 + 14, row2Y + 14);

  doc.setTextColor(0);

  // Table
  const head = [[
    "Mes",
    "Fecha de cronograma (1)",
    "Fecha de pago (2)",
    "Días",
    `Monto base (${moneda})`,
    `Interés bruto (${moneda})`,
    `Impuesto 2da categ. (${moneda})`,
    `Interés a Depositar (${moneda})`,
    `Devolución Capital (${moneda})`,
    `Saldo Capital (${moneda})`,
    `Total a depositar (${moneda})`,
  ]];

  const body = result.rows.map(r => ([
    r.mes,
    formatDateDMY(r.fechaCrono),
    r.mes === 0 ? "" : formatDateDMY(r.fechaPago),
    String(r.dias),
    formatMoney(r.montoBase),
    formatMoney(r.interesBruto),
    formatMoney(r.impuesto),
    formatMoney(r.interesDepositar),
    formatMoney(r.devolucionCapital),
    formatMoney(r.saldo),
    formatMoney(r.totalDepositar),
  ]));

  // Footer totals row (solo H, I, K)
  const foot = [[
    "Total:",
    "", "", "",
    "", "", "",
    formatMoney(result.totals.interesDepositar),
    formatMoney(result.totals.devolucionCapital),
    "",
    formatMoney(result.totals.totalDepositar),
  ]];

  // AutoTable: forzar que quepa a lo ancho (A4 landscape)
  doc.autoTable({
    head,
    body,
    foot,
    startY: sumTop + boxH + 14,
    margin: { left: margin, right: margin },
    styles: {
      font: "helvetica",
      fontSize: 8.5,
      cellPadding: 4,
      halign: "center",
      valign: "middle",
      lineColor: [225, 231, 235],
      lineWidth: 0.6,
      overflow: "linebreak"
    },
    headStyles: {
      fillColor: [238, 242, 247],
      textColor: [11, 18, 32],
      fontStyle: "bold"
    },
    footStyles: {
      fillColor: [243, 244, 246],
      textColor: [11, 18, 32],
      fontStyle: "bold"
    },
    alternateRowStyles: { fillColor: [250, 251, 252] },
    tableWidth: "auto",
    // Reduce si fuera necesario:
    didDrawPage: () => {},
  });

  doc.save(buildPdfFileName(generatedAt));
}

/* =========================================================
   EVENTOS
========================================================= */
form.addEventListener("submit", async (e) => {
  e.preventDefault();

  clearFieldErrors();

  const raw = readForm();

  // VBA: si opcionPrimerPago vacío => "Próximo mes"
  if (!String(raw.opcionPrimerPago || "").trim()) raw.opcionPrimerPago = "Próximo mes";

  const errors = validateInputs(raw);
  if (Object.keys(errors).length > 0) {
    showErrors(errors);
    lastResult = null;
    btnPdf.disabled = true;
    return;
  }

  const inputs = {
    fechaInicio: raw.fechaInicio,
    monto: Number(raw.monto),
    monedaRaw: raw.monedaRaw,
    tasaEA_pct: Number(raw.tasaEA_pct),
    plazo: Number(raw.plazo),
    producto: String(raw.producto).trim(),
    frecuenciaInteres: String(raw.frecuenciaInteres).trim(),
    frecuenciaCapital: String(raw.frecuenciaCapital).trim(),
    opcionPrimerPago: String(raw.opcionPrimerPago).trim()
  };

  const generatedAt = new Date();

  // Generar (replica VBA)
  const result = generarCronograma(inputs);

  // Logo embebido (para PDF)
  const logoDataUrl = await fetchLogoAsDataURL();
  logoWarning.hidden = !!logoDataUrl;

  // Render
  renderResult(inputs, result, generatedAt);

  // State
  lastResult = { inputs, result, generatedAt, logoDataUrl };
});

btnLimpiar.addEventListener("click", () => {
  form.reset();
  clearFieldErrors();

  resumen.hidden = true;
  tableWrap.hidden = true;
  emptyState.hidden = false;

  btnPdf.disabled = true;
  lastResult = null;

  setGeneratedNow();
});

btnPdf.addEventListener("click", async () => {
  if (!lastResult) return;
  await exportPDF(lastResult);
});

/* =========================================================
   Defaults (solo visuales para probar rápido)
========================================================= */
// Puedes comentar esto si no quieres valores por defecto.
(function setDefaults() {
  const today = new Date();
  const yyyy = today.getFullYear();
  const mm = pad2(today.getMonth() + 1);
  const dd = pad2(today.getDate());

  const elFecha = document.getElementById("fechaInicio");
  if (!elFecha.value) elFecha.value = `${yyyy}-${mm}-${dd}`;

  const elMonto = document.getElementById("monto");
  if (!elMonto.value) elMonto.value = "100000";

  const elTasa = document.getElementById("tasaEA");
  if (!elTasa.value) elTasa.value = "18";

  const elPlazo = document.getElementById("plazo");
  if (!elPlazo.value) elPlazo.value = "15";

  const elProd = document.getElementById("producto");
  if (!elProd.value) elProd.value = "PET";

  const elFI = document.getElementById("freqInteres");
  if (!elFI.value) elFI.value = "Semestral";

  const elFC = document.getElementById("freqCapital");
  if (!elFC.value) elFC.value = "Al finalizar";

  const elOP = document.getElementById("opcionPrimerPago");
  if (!elOP.value) elOP.value = "Próximo mes";
})();
