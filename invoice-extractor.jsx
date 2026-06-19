import React, { useState, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import { Upload, FileText, Trash2, Download, Loader2, CheckCircle2, AlertCircle, Plus, FileSpreadsheet, X, Pencil } from "lucide-react";

// ---------- helpers ----------

function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result.split(",")[1]);
    reader.onerror = () => reject(new Error("No se pudo leer el archivo"));
    reader.readAsDataURL(file);
  });
}

function mediaTypeFor(file) {
  if (file.type === "application/pdf") return "application/pdf";
  if (file.type === "image/png") return "image/png";
  if (file.type === "image/webp") return "image/webp";
  if (file.type === "image/gif") return "image/gif";
  return "image/jpeg";
}

const EXTRACTION_PROMPT = `Eres un asistente experto en contabilidad. Vas a analizar la imagen o PDF de una factura de un proveedor.

Extrae TODOS los conceptos/líneas de la factura (cada producto o servicio facturado) y los datos generales del documento.

Responde ÚNICAMENTE con un objeto JSON válido, sin texto adicional, sin explicaciones, sin markdown, con esta forma exacta:

{
  "proveedor": "nombre del proveedor o emisor",
  "cif_nif": "CIF/NIF/RFC del proveedor si aparece, si no, cadena vacía",
  "numero_factura": "número o folio de factura",
  "fecha": "fecha de la factura en formato YYYY-MM-DD si es posible, si no, como aparece",
  "moneda": "moneda detectada, ej. EUR, USD, MXN",
  "conceptos": [
    {
      "descripcion": "texto del concepto o producto/servicio",
      "cantidad": numero,
      "precio_unitario": numero,
      "impuesto_pct": numero o null si no aplica/no se ve,
      "importe": numero
    }
  ],
  "subtotal": numero o null,
  "impuestos_total": numero o null,
  "total": numero o null
}

Reglas:
- Si un campo numérico no aparece o no se puede determinar con certeza, usa null (no inventes cifras).
- "cantidad" por defecto es 1 si la factura no especifica cantidad para esa línea.
- Usa números (no texto) para cantidad, precio_unitario, impuesto_pct, importe, subtotal, impuestos_total, total. Sin símbolos de moneda ni separadores de miles.
- Incluye cada línea de concepto que exista en la factura, no resumas ni agrupes conceptos distintos.
- No incluyas líneas de "total", "subtotal" o "IVA" dentro de la lista de conceptos; esas van en los campos generales.
- Devuelve solo el JSON, nada más.`;

function safeParseJSON(text) {
  let cleaned = text.trim();
  cleaned = cleaned.replace(/^```json\s*/i, "").replace(/^```\s*/i, "").replace(/```\s*$/i, "");
  const firstBrace = cleaned.indexOf("{");
  const lastBrace = cleaned.lastIndexOf("}");
  if (firstBrace !== -1 && lastBrace !== -1) {
    cleaned = cleaned.slice(firstBrace, lastBrace + 1);
  }
  return JSON.parse(cleaned);
}

function numOrEmpty(v) {
  return typeof v === "number" && !Number.isNaN(v) ? v : "";
}

function fmtMoney(v, currency) {
  if (typeof v !== "number" || Number.isNaN(v)) return "—";
  try {
    return new Intl.NumberFormat("es-ES", { style: "currency", currency: currency || "EUR", maximumFractionDigits: 2 }).format(v);
  } catch {
    return v.toFixed(2);
  }
}

let idCounter = 1;
function nextId() {
  return `inv_${idCounter++}_${Date.now()}`;
}

// ---------- main component ----------

export default function InvoiceExtractor() {
  const [invoices, setInvoices] = useState([]); // {id, fileName, status, error, data, previewUrl}
  const [selectedId, setSelectedId] = useState(null);
  const [isDragging, setIsDragging] = useState(false);
  const fileInputRef = useRef(null);

  const selected = invoices.find((i) => i.id === selectedId) || null;

  const updateInvoice = useCallback((id, patch) => {
    setInvoices((prev) => prev.map((inv) => (inv.id === id ? { ...inv, ...patch } : inv)));
  }, []);

  const processFile = useCallback(async (id, file) => {
    updateInvoice(id, { status: "processing" });
    try {
      const base64 = await fileToBase64(file);
      const mediaType = mediaTypeFor(file);
      const contentBlock =
        mediaType === "application/pdf"
          ? { type: "document", source: { type: "base64", media_type: mediaType, data: base64 } }
          : { type: "image", source: { type: "base64", media_type: mediaType, data: base64 } };

      const response = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: "claude-sonnet-4-6",
          max_tokens: 4000,
          messages: [
            {
              role: "user",
              content: [contentBlock, { type: "text", text: EXTRACTION_PROMPT }],
            },
          ],
        }),
      });

      if (!response.ok) {
        throw new Error(`Error del servicio (${response.status})`);
      }

      const data = await response.json();
      const textBlock = (data.content || []).find((b) => b.type === "text");
      if (!textBlock) throw new Error("No se recibió respuesta de texto");

      const parsed = safeParseJSON(textBlock.text);

      const conceptos = Array.isArray(parsed.conceptos)
        ? parsed.conceptos.map((c) => ({
            id: nextId(),
            descripcion: c.descripcion || "",
            cantidad: typeof c.cantidad === "number" ? c.cantidad : 1,
            precio_unitario: typeof c.precio_unitario === "number" ? c.precio_unitario : "",
            impuesto_pct: typeof c.impuesto_pct === "number" ? c.impuesto_pct : "",
            importe: typeof c.importe === "number" ? c.importe : "",
          }))
        : [];

      updateInvoice(id, {
        status: "done",
        data: {
          proveedor: parsed.proveedor || "",
          cif_nif: parsed.cif_nif || "",
          numero_factura: parsed.numero_factura || "",
          fecha: parsed.fecha || "",
          moneda: parsed.moneda || "EUR",
          conceptos,
          subtotal: typeof parsed.subtotal === "number" ? parsed.subtotal : "",
          impuestos_total: typeof parsed.impuestos_total === "number" ? parsed.impuestos_total : "",
          total: typeof parsed.total === "number" ? parsed.total : "",
        },
      });
    } catch (err) {
      updateInvoice(id, { status: "error", error: err.message || "No se pudo procesar la factura" });
    }
  }, [updateInvoice]);

  const addFiles = useCallback((fileList) => {
    const files = Array.from(fileList).filter((f) => {
      const okType =
        f.type === "application/pdf" ||
        f.type === "image/jpeg" ||
        f.type === "image/png" ||
        f.type === "image/webp" ||
        f.type === "image/gif";
      return okType;
    });
    if (files.length === 0) return;

    const newInvoices = files.map((file) => ({
      id: nextId(),
      fileName: file.name,
      status: "queued",
      error: null,
      data: null,
      previewUrl: file.type === "application/pdf" ? null : URL.createObjectURL(file),
      isPdf: file.type === "application/pdf",
    }));

    setInvoices((prev) => [...prev, ...newInvoices]);
    if (!selectedId && newInvoices[0]) setSelectedId(newInvoices[0].id);

    newInvoices.forEach((inv, idx) => {
      processFile(inv.id, files[idx]);
    });
  }, [processFile, selectedId]);

  const handleDrop = useCallback((e) => {
    e.preventDefault();
    setIsDragging(false);
    if (e.dataTransfer.files?.length) addFiles(e.dataTransfer.files);
  }, [addFiles]);

  const removeInvoice = useCallback((id) => {
    setInvoices((prev) => prev.filter((i) => i.id !== id));
    setSelectedId((prev) => (prev === id ? null : prev));
  }, []);

  const retryInvoice = useCallback((id, fileName) => {
    // We don't keep the raw File object around for re-processing after error,
    // so guide the user to re-upload instead in that edge case.
  }, []);

  const updateConcepto = useCallback((invId, conceptoId, field, value) => {
    setInvoices((prev) =>
      prev.map((inv) => {
        if (inv.id !== invId || !inv.data) return inv;
        const conceptos = inv.data.conceptos.map((c) =>
          c.id === conceptoId ? { ...c, [field]: value } : c
        );
        return { ...inv, data: { ...inv.data, conceptos } };
      })
    );
  }, []);

  const updateField = useCallback((invId, field, value) => {
    setInvoices((prev) =>
      prev.map((inv) => (inv.id === invId && inv.data ? { ...inv, data: { ...inv.data, [field]: value } } : inv))
    );
  }, []);

  const addConcepto = useCallback((invId) => {
    setInvoices((prev) =>
      prev.map((inv) => {
        if (inv.id !== invId || !inv.data) return inv;
        const nuevo = { id: nextId(), descripcion: "", cantidad: 1, precio_unitario: "", impuesto_pct: "", importe: "" };
        return { ...inv, data: { ...inv.data, conceptos: [...inv.data.conceptos, nuevo] } };
      })
    );
  }, []);

  const removeConcepto = useCallback((invId, conceptoId) => {
    setInvoices((prev) =>
      prev.map((inv) => {
        if (inv.id !== invId || !inv.data) return inv;
        return { ...inv, data: { ...inv.data, conceptos: inv.data.conceptos.filter((c) => c.id !== conceptoId) } };
      })
    );
  }, []);

  const doneInvoices = invoices.filter((i) => i.status === "done" && i.data);
  const totalConceptos = doneInvoices.reduce((sum, i) => sum + i.data.conceptos.length, 0);
  const processingCount = invoices.filter((i) => i.status === "processing" || i.status === "queued").length;

  const exportToExcel = useCallback(() => {
    if (doneInvoices.length === 0) return;

    // Hoja 1: detalle de conceptos
    const detalleRows = [];
    doneInvoices.forEach((inv) => {
      inv.data.conceptos.forEach((c) => {
        detalleRows.push({
          "Archivo": inv.fileName,
          "Proveedor": inv.data.proveedor,
          "CIF/NIF": inv.data.cif_nif,
          "Nº Factura": inv.data.numero_factura,
          "Fecha": inv.data.fecha,
          "Moneda": inv.data.moneda,
          "Concepto": c.descripcion,
          "Cantidad": numOrEmpty(c.cantidad),
          "Precio unitario": numOrEmpty(c.precio_unitario),
          "Impuesto %": numOrEmpty(c.impuesto_pct),
          "Importe": numOrEmpty(c.importe),
        });
      });
    });

    // Hoja 2: resumen por factura
    const resumenRows = doneInvoices.map((inv) => ({
      "Archivo": inv.fileName,
      "Proveedor": inv.data.proveedor,
      "CIF/NIF": inv.data.cif_nif,
      "Nº Factura": inv.data.numero_factura,
      "Fecha": inv.data.fecha,
      "Moneda": inv.data.moneda,
      "Nº Conceptos": inv.data.conceptos.length,
      "Subtotal": numOrEmpty(inv.data.subtotal),
      "Impuestos": numOrEmpty(inv.data.impuestos_total),
      "Total": numOrEmpty(inv.data.total),
    }));

    const wb = XLSX.utils.book_new();
    const wsResumen = XLSX.utils.json_to_sheet(resumenRows);
    const wsDetalle = XLSX.utils.json_to_sheet(detalleRows);

    wsResumen["!cols"] = [
      { wch: 26 }, { wch: 22 }, { wch: 16 }, { wch: 16 }, { wch: 12 },
      { wch: 8 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 },
    ];
    wsDetalle["!cols"] = [
      { wch: 26 }, { wch: 22 }, { wch: 16 }, { wch: 16 }, { wch: 12 }, { wch: 8 },
      { wch: 36 }, { wch: 10 }, { wch: 14 }, { wch: 10 }, { wch: 12 },
    ];

    XLSX.utils.book_append_sheet(wb, wsResumen, "Resumen facturas");
    XLSX.utils.book_append_sheet(wb, wsDetalle, "Conceptos");

    XLSX.writeFile(wb, `facturas_extraidas_${new Date().toISOString().slice(0, 10)}.xlsx`);
  }, [doneInvoices]);

  return (
    <div
      style={{
        minHeight: "100vh",
        background: "var(--ledger-bg)",
        color: "var(--ink)",
        fontFamily: "var(--font-body)",
      }}
    >
      <style>{`
        :root {
          --ledger-bg: #f3eee2;
          --paper: #fbf8f0;
          --ink: #2b2620;
          --ink-soft: #6b6256;
          --rule: #d8cfb8;
          --forest: #2f4a3c;
          --forest-deep: #1f342a;
          --rust: #a14e2a;
          --gold: #b8893a;
          --ok: #2f6b4f;
          --err: #a13c2a;
          --font-display: 'Source Serif Pro', Georgia, 'Times New Roman', serif;
          --font-body: 'Source Serif Pro', Georgia, serif;
          --font-mono: 'IBM Plex Mono', 'Courier New', monospace;
        }
        @import url('https://fonts.googleapis.com/css2?family=Source+Serif+Pro:ital,wght@0,400;0,600;0,700;1,400&family=IBM+Plex+Mono:wght@400;500;600&display=swap');

        * { box-sizing: border-box; }
        ::selection { background: var(--gold); color: var(--paper); }

        .ledger-rule {
          background-image: repeating-linear-gradient(
            to bottom,
            transparent 0px,
            transparent 27px,
            var(--rule) 27px,
            var(--rule) 28px
          );
        }

        .mono { font-family: var(--font-mono); }

        input.cell-input {
          font-family: var(--font-mono);
          background: transparent;
          border: none;
          border-bottom: 1px solid transparent;
          width: 100%;
          padding: 2px 4px;
          color: var(--ink);
          font-size: 13px;
        }
        input.cell-input:focus {
          outline: none;
          border-bottom: 1px solid var(--forest);
          background: rgba(47,74,60,0.05);
        }
        input.cell-input.text-mode {
          font-family: var(--font-body);
        }

        .scrollbar-thin::-webkit-scrollbar { width: 8px; height: 8px; }
        .scrollbar-thin::-webkit-scrollbar-thumb { background: var(--rule); border-radius: 4px; }
        .scrollbar-thin::-webkit-scrollbar-track { background: transparent; }

        .stamp {
          transform: rotate(-3deg);
        }

        @keyframes pulse-soft {
          0%, 100% { opacity: 1; }
          50% { opacity: 0.5; }
        }
        .pulse { animation: pulse-soft 1.4s ease-in-out infinite; }
      `}</style>

      {/* Header */}
      <header
        style={{
          borderBottom: "2px solid var(--forest-deep)",
          background: "var(--forest)",
        }}
        className="px-6 py-5"
      >
        <div className="max-w-7xl mx-auto flex items-center justify-between flex-wrap gap-4">
          <div className="flex items-center gap-3">
            <div
              style={{
                width: 42,
                height: 42,
                border: "2px solid #e8dfc8",
                borderRadius: 4,
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                background: "var(--forest-deep)",
              }}
            >
              <FileText size={22} color="#e8dfc8" />
            </div>
            <div>
              <h1
                style={{ fontFamily: "var(--font-display)", color: "#f3eee2", letterSpacing: "0.01em" }}
                className="text-2xl font-semibold leading-tight"
              >
                Libro de Facturas
              </h1>
              <p className="mono text-xs" style={{ color: "#cfd9c9" }}>
                extracción de conceptos · proveedores · exportación a Excel
              </p>
            </div>
          </div>

          <div className="flex items-center gap-4">
            <div className="text-right">
              <div className="mono text-xs" style={{ color: "#cfd9c9" }}>facturas cargadas</div>
              <div className="mono text-lg font-medium" style={{ color: "#f3eee2" }}>{invoices.length}</div>
            </div>
            <div className="text-right">
              <div className="mono text-xs" style={{ color: "#cfd9c9" }}>conceptos extraídos</div>
              <div className="mono text-lg font-medium" style={{ color: "#f3eee2" }}>{totalConceptos}</div>
            </div>
            <button
              onClick={exportToExcel}
              disabled={doneInvoices.length === 0}
              style={{
                background: doneInvoices.length === 0 ? "#5a6f63" : "var(--gold)",
                color: doneInvoices.length === 0 ? "#9aa89e" : "#2b2113",
                border: "none",
                cursor: doneInvoices.length === 0 ? "not-allowed" : "pointer",
              }}
              className="flex items-center gap-2 px-4 py-2.5 rounded font-medium text-sm transition-colors"
            >
              <FileSpreadsheet size={17} />
              Exportar a Excel
            </button>
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-6 py-7">
        <div className="grid grid-cols-1 lg:grid-cols-[320px_1fr] gap-6">
          {/* Left column: upload + list */}
          <div className="flex flex-col gap-4">
            <div
              onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
              onDragLeave={() => setIsDragging(false)}
              onDrop={handleDrop}
              onClick={() => fileInputRef.current?.click()}
              style={{
                background: "var(--paper)",
                border: `2px dashed ${isDragging ? "var(--forest)" : "var(--rule)"}`,
                borderRadius: 6,
                cursor: "pointer",
                transition: "border-color 0.15s",
              }}
              className="p-6 flex flex-col items-center justify-center text-center gap-2"
            >
              <Upload size={26} color="var(--forest)" strokeWidth={1.6} />
              <p className="font-medium text-sm" style={{ fontFamily: "var(--font-display)" }}>
                Arrastra facturas aquí
              </p>
              <p className="text-xs" style={{ color: "var(--ink-soft)" }}>
                o haz clic para elegir archivos
              </p>
              <p className="mono text-[11px] mt-1" style={{ color: "var(--ink-soft)" }}>
                JPG · PNG · WEBP · PDF
              </p>
              <input
                ref={fileInputRef}
                type="file"
                multiple
                accept="image/jpeg,image/png,image/webp,image/gif,application/pdf"
                className="hidden"
                onChange={(e) => { if (e.target.files?.length) addFiles(e.target.files); e.target.value = ""; }}
              />
            </div>

            <div
              style={{ background: "var(--paper)", border: "1px solid var(--rule)", borderRadius: 6 }}
              className="flex-1 overflow-hidden flex flex-col"
            >
              <div
                style={{ borderBottom: "1px solid var(--rule)" }}
                className="px-4 py-3 flex items-center justify-between"
              >
                <span className="mono text-xs uppercase tracking-wide" style={{ color: "var(--ink-soft)" }}>
                  Bandeja
                </span>
                {processingCount > 0 && (
                  <span className="mono text-xs flex items-center gap-1.5" style={{ color: "var(--gold)" }}>
                    <Loader2 size={13} className="animate-spin" />
                    procesando {processingCount}
                  </span>
                )}
              </div>

              <div className="scrollbar-thin overflow-y-auto" style={{ maxHeight: 560 }}>
                {invoices.length === 0 ? (
                  <p className="px-4 py-6 text-sm text-center" style={{ color: "var(--ink-soft)" }}>
                    Aún no hay facturas. Sube una para empezar.
                  </p>
                ) : (
                  invoices.map((inv) => (
                    <div
                      key={inv.id}
                      onClick={() => setSelectedId(inv.id)}
                      style={{
                        borderBottom: "1px solid var(--rule)",
                        background: selectedId === inv.id ? "rgba(47,74,60,0.08)" : "transparent",
                        borderLeft: selectedId === inv.id ? "3px solid var(--forest)" : "3px solid transparent",
                        cursor: "pointer",
                      }}
                      className="px-4 py-3 flex items-center gap-3 group"
                    >
                      <div
                        style={{
                          width: 34, height: 34, borderRadius: 4, flexShrink: 0,
                          background: "#eee7d4", border: "1px solid var(--rule)",
                          display: "flex", alignItems: "center", justifyContent: "center", overflow: "hidden",
                        }}
                      >
                        {inv.previewUrl ? (
                          <img src={inv.previewUrl} alt="" style={{ width: "100%", height: "100%", objectFit: "cover" }} />
                        ) : (
                          <FileText size={16} color="var(--ink-soft)" />
                        )}
                      </div>

                      <div className="flex-1 min-w-0">
                        <p className="text-sm font-medium truncate" title={inv.fileName}>{inv.fileName}</p>
                        <div className="flex items-center gap-1.5 mt-0.5">
                          {inv.status === "queued" && (
                            <span className="mono text-[11px]" style={{ color: "var(--ink-soft)" }}>en cola…</span>
                          )}
                          {inv.status === "processing" && (
                            <span className="mono text-[11px] flex items-center gap-1 pulse" style={{ color: "var(--gold)" }}>
                              <Loader2 size={11} className="animate-spin" /> leyendo factura
                            </span>
                          )}
                          {inv.status === "done" && (
                            <span className="mono text-[11px] flex items-center gap-1" style={{ color: "var(--ok)" }}>
                              <CheckCircle2 size={11} /> {inv.data.conceptos.length} conceptos
                            </span>
                          )}
                          {inv.status === "error" && (
                            <span className="mono text-[11px] flex items-center gap-1" style={{ color: "var(--err)" }} title={inv.error}>
                              <AlertCircle size={11} /> error
                            </span>
                          )}
                        </div>
                      </div>

                      <button
                        onClick={(e) => { e.stopPropagation(); removeInvoice(inv.id); }}
                        style={{ color: "var(--ink-soft)" }}
                        className="opacity-0 group-hover:opacity-100 transition-opacity p-1 hover:text-[var(--err)]"
                        title="Quitar"
                      >
                        <X size={15} />
                      </button>
                    </div>
                  ))
                )}
              </div>
            </div>
          </div>

          {/* Right column: detail */}
          <div style={{ background: "var(--paper)", border: "1px solid var(--rule)", borderRadius: 6 }} className="min-h-[600px] flex flex-col">
            {!selected && (
              <div className="flex-1 flex flex-col items-center justify-center gap-2 p-12 text-center">
                <FileSpreadsheet size={40} color="var(--rule)" strokeWidth={1.3} />
                <p style={{ fontFamily: "var(--font-display)", color: "var(--ink-soft)" }} className="text-lg">
                  Selecciona una factura de la bandeja
                </p>
                <p className="text-sm" style={{ color: "var(--ink-soft)" }}>
                  Cada concepto extraído se podrá revisar y editar antes de exportar.
                </p>
              </div>
            )}

            {selected && selected.status === "processing" && (
              <div className="flex-1 flex flex-col items-center justify-center gap-3 p-12">
                <Loader2 size={32} className="animate-spin" color="var(--forest)" />
                <p style={{ fontFamily: "var(--font-display)" }} className="text-lg">Leyendo {selected.fileName}…</p>
                <p className="text-sm" style={{ color: "var(--ink-soft)" }}>Extrayendo proveedor, fecha y conceptos</p>
              </div>
            )}

            {selected && selected.status === "queued" && (
              <div className="flex-1 flex flex-col items-center justify-center gap-3 p-12">
                <p style={{ fontFamily: "var(--font-display)" }} className="text-lg">En cola…</p>
              </div>
            )}

            {selected && selected.status === "error" && (
              <div className="flex-1 flex flex-col items-center justify-center gap-3 p-12 text-center">
                <AlertCircle size={32} color="var(--err)" />
                <p style={{ fontFamily: "var(--font-display)" }} className="text-lg">No se pudo leer esta factura</p>
                <p className="text-sm max-w-md" style={{ color: "var(--ink-soft)" }}>{selected.error}</p>
                <p className="text-xs" style={{ color: "var(--ink-soft)" }}>
                  Quítala de la bandeja y vuelve a subir el archivo para intentarlo de nuevo.
                </p>
              </div>
            )}

            {selected && selected.status === "done" && selected.data && (
              <InvoiceDetail
                invoice={selected}
                onUpdateField={updateField}
                onUpdateConcepto={updateConcepto}
                onAddConcepto={addConcepto}
                onRemoveConcepto={removeConcepto}
              />
            )}
          </div>
        </div>
      </main>

      <footer className="max-w-7xl mx-auto px-6 pb-10 pt-2">
        <p className="mono text-[11px]" style={{ color: "var(--ink-soft)" }}>
          Los datos se procesan en esta sesión y no se almacenan. Revisa los importes antes de exportar.
        </p>
      </footer>
    </div>
  );
}

function InvoiceDetail({ invoice, onUpdateField, onUpdateConcepto, onAddConcepto, onRemoveConcepto }) {
  const { data } = invoice;
  const currency = data.moneda || "EUR";

  const calcTotal = data.conceptos.reduce((sum, c) => sum + (typeof c.importe === "number" ? c.importe : 0), 0);

  return (
    <div className="flex flex-col h-full">
      {/* Header fields */}
      <div style={{ borderBottom: "1px solid var(--rule)" }} className="p-6 pb-5">
        <div className="flex items-start justify-between mb-4">
          <div>
            <p className="mono text-[11px] uppercase tracking-wide" style={{ color: "var(--ink-soft)" }}>
              {invoice.fileName}
            </p>
          </div>
          <span
            className="stamp mono text-[10px] px-2 py-1 rounded uppercase tracking-wider"
            style={{ border: "1.5px solid var(--ok)", color: "var(--ok)" }}
          >
            extraída
          </span>
        </div>

        <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
          <Field label="Proveedor" value={data.proveedor} onChange={(v) => onUpdateField(invoice.id, "proveedor", v)} />
          <Field label="CIF / NIF" value={data.cif_nif} onChange={(v) => onUpdateField(invoice.id, "cif_nif", v)} mono />
          <Field label="Nº factura" value={data.numero_factura} onChange={(v) => onUpdateField(invoice.id, "numero_factura", v)} mono />
          <Field label="Fecha" value={data.fecha} onChange={(v) => onUpdateField(invoice.id, "fecha", v)} mono />
        </div>
      </div>

      {/* Conceptos table */}
      <div className="flex-1 overflow-auto scrollbar-thin px-6 py-4">
        <table className="w-full text-sm">
          <thead>
            <tr style={{ borderBottom: "2px solid var(--ink)" }}>
              <th className="text-left pb-2 font-medium mono text-xs uppercase" style={{ color: "var(--ink-soft)" }}>Concepto</th>
              <th className="text-right pb-2 font-medium mono text-xs uppercase w-20" style={{ color: "var(--ink-soft)" }}>Cant.</th>
              <th className="text-right pb-2 font-medium mono text-xs uppercase w-28" style={{ color: "var(--ink-soft)" }}>Precio u.</th>
              <th className="text-right pb-2 font-medium mono text-xs uppercase w-20" style={{ color: "var(--ink-soft)" }}>Imp. %</th>
              <th className="text-right pb-2 font-medium mono text-xs uppercase w-28" style={{ color: "var(--ink-soft)" }}>Importe</th>
              <th className="w-8"></th>
            </tr>
          </thead>
          <tbody>
            {data.conceptos.length === 0 && (
              <tr>
                <td colSpan={6} className="py-6 text-center text-sm" style={{ color: "var(--ink-soft)" }}>
                  No se detectaron conceptos. Añade uno manualmente si es necesario.
                </td>
              </tr>
            )}
            {data.conceptos.map((c) => (
              <tr key={c.id} style={{ borderBottom: "1px solid var(--rule)" }} className="group">
                <td className="py-1.5">
                  <input
                    className="cell-input text-mode"
                    value={c.descripcion}
                    onChange={(e) => onUpdateConcepto(invoice.id, c.id, "descripcion", e.target.value)}
                    placeholder="Descripción del concepto"
                  />
                </td>
                <td className="py-1.5">
                  <input
                    className="cell-input text-right"
                    type="number"
                    value={c.cantidad}
                    onChange={(e) => onUpdateConcepto(invoice.id, c.id, "cantidad", e.target.value === "" ? "" : Number(e.target.value))}
                  />
                </td>
                <td className="py-1.5">
                  <input
                    className="cell-input text-right"
                    type="number"
                    value={c.precio_unitario}
                    onChange={(e) => onUpdateConcepto(invoice.id, c.id, "precio_unitario", e.target.value === "" ? "" : Number(e.target.value))}
                    placeholder="—"
                  />
                </td>
                <td className="py-1.5">
                  <input
                    className="cell-input text-right"
                    type="number"
                    value={c.impuesto_pct}
                    onChange={(e) => onUpdateConcepto(invoice.id, c.id, "impuesto_pct", e.target.value === "" ? "" : Number(e.target.value))}
                    placeholder="—"
                  />
                </td>
                <td className="py-1.5">
                  <input
                    className="cell-input text-right font-medium"
                    type="number"
                    value={c.importe}
                    onChange={(e) => onUpdateConcepto(invoice.id, c.id, "importe", e.target.value === "" ? "" : Number(e.target.value))}
                    placeholder="—"
                  />
                </td>
                <td className="py-1.5 text-right">
                  <button
                    onClick={() => onRemoveConcepto(invoice.id, c.id)}
                    className="opacity-0 group-hover:opacity-100 transition-opacity"
                    style={{ color: "var(--ink-soft)" }}
                    title="Eliminar concepto"
                  >
                    <Trash2 size={14} />
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>

        <button
          onClick={() => onAddConcepto(invoice.id)}
          className="mt-3 flex items-center gap-1.5 text-sm font-medium"
          style={{ color: "var(--forest)" }}
        >
          <Plus size={15} /> Añadir concepto
        </button>
      </div>

      {/* Totals */}
      <div style={{ borderTop: "1px solid var(--rule)", background: "rgba(47,74,60,0.04)" }} className="p-6 pt-4">
        <div className="flex justify-end gap-8">
          <TotalLine label="Suma de conceptos" value={fmtMoney(calcTotal, currency)} />
          <TotalLine
            label="Subtotal (factura)"
            value={typeof data.subtotal === "number" ? fmtMoney(data.subtotal, currency) : "—"}
          />
          <TotalLine
            label="Impuestos"
            value={typeof data.impuestos_total === "number" ? fmtMoney(data.impuestos_total, currency) : "—"}
          />
          <TotalLine
            label="Total factura"
            value={typeof data.total === "number" ? fmtMoney(data.total, currency) : "—"}
            emphasized
          />
        </div>
      </div>
    </div>
  );
}

function Field({ label, value, onChange, mono }) {
  return (
    <div>
      <label className="mono text-[10px] uppercase tracking-wide block mb-1" style={{ color: "var(--ink-soft)" }}>
        {label}
      </label>
      <input
        value={value}
        onChange={(e) => onChange(e.target.value)}
        className={mono ? "cell-input" : "cell-input text-mode"}
        style={{
          fontSize: 14,
          fontWeight: 500,
          borderBottom: "1px solid var(--rule)",
          padding: "3px 2px",
        }}
      />
    </div>
  );
}

function TotalLine({ label, value, emphasized }) {
  return (
    <div className="text-right">
      <div className="mono text-[10px] uppercase tracking-wide mb-0.5" style={{ color: "var(--ink-soft)" }}>
        {label}
      </div>
      <div
        className="mono"
        style={{
          fontSize: emphasized ? 18 : 15,
          fontWeight: emphasized ? 700 : 500,
          color: emphasized ? "var(--forest-deep)" : "var(--ink)",
        }}
      >
        {value}
      </div>
    </div>
  );
}
