(() => {
  const API = "/api/proveedor";
  const LINK_STORAGE_KEY = "qaMantoProveedorOneDriveUrl";
  let selectedBaseFile = null;

  const esc = value => String(value ?? "").replace(/[&<>"']/g, ch => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
  }[ch]));

  async function request(url, options = {}) {
    const response = await fetch(url, options);
    const data = await response.json().catch(() => ({}));
    if (!response.ok) {
      throw new Error(data.error || `Error ${response.status}`);
    }
    return data;
  }

  function downloadBase64Xlsx(base64, filename) {
    const binary = atob(base64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
    const blob = new Blob([bytes], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename || "Proveedor_Actualizado.xlsx";
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 5000);
  }

  function render() {
    const savedLink = localStorage.getItem(LINK_STORAGE_KEY) || "";
    return `
      <section class="screen">
        ${window.QAMantoHeaderHTML ? window.QAMantoHeaderHTML() : ""}
        <div class="page proveedor-page">
          <div class="topbar">
            <div>
              <h2>Proveedor</h2>
              <p class="ad-module-subtitle">Sincronización directa usando el enlace compartido de OneDrive (sin iniciar sesión).</p>
            </div>
            <div class="topbar-actions">
              <button id="proveedor-back" class="btn btn-secondary" type="button">← Dashboard</button>
            </div>
          </div>

          <section class="search-card proveedor-card">
            <h3>1. Enlace compartido del Excel en OneDrive</h3>
            <div class="form-group">
              <label for="proveedor-onedrive-url">URL compartida del archivo maestro (.xlsx)</label>
              <div style="display:flex; gap:10px; flex-wrap:wrap;">
                <input id="proveedor-onedrive-url" type="url" style="flex:1; min-width:260px;"
                  placeholder="https://1drv.ms/x/s!... o https://onedrive.live.com/..."
                  value="${esc(savedLink)}">
                <button id="proveedor-test-link" class="btn btn-outline" type="button">Verificar enlace</button>
              </div>
              <div id="proveedor-link-status" class="field-help">
                En OneDrive haz clic en <strong>Compartir → Copiar vínculo</strong> (configurado como <em>"Cualquier persona que tenga el vínculo"</em>). No requiere iniciar sesión.
              </div>
            </div>
          </section>

          <section class="search-card proveedor-card">
            <h3>2. Excel base (desde tu computadora)</h3>
            <p class="field-help">Selecciona el archivo local con los registros y acciones (Agregar, Actualizar o Eliminar) que se aplicarán al Excel de OneDrive.</p>
            <div class="proveedor-upload-row">
              <label class="btn btn-outline" for="proveedor-base-file">Seleccionar Excel base</label>
              <input id="proveedor-base-file" class="hidden" type="file" accept=".xlsx">
              <span id="proveedor-base-name" class="proveedor-file-name">Ningún archivo seleccionado</span>
            </div>
            <div class="field-help">Formato requerido: .xlsx con 41 columnas y encabezados Supplier ID, Allocation Number y Action Code.</div>
          </section>

          <section class="search-card proveedor-card">
            <h3>3. Sincronización y generación de Excel actualizado</h3>
            <div class="proveedor-rules">
              <div><strong>Agregar</strong><span>Registros nuevos que no existen en OneDrive.</span></div>
              <div><strong>Actualizar</strong><span>Registros existentes con diferencias en sus columnas.</span></div>
              <div><strong>Eliminar</strong><span>Registros con Action Code deleven o sysdel.</span></div>
            </div>
            <p class="proveedor-warning">El sistema lee el archivo directamente desde el enlace de OneDrive, crea un respaldo local automático y descarga el Excel actualizado listo para subir.</p>
            <button id="proveedor-sync" class="btn btn-primary" type="button" disabled>Sincronizar con enlace de OneDrive</button>
            <span id="proveedor-progress" class="field-help" aria-live="polite"></span>
          </section>

          <section id="proveedor-report" class="results-card">
            <div class="results-head"><h3>Reporte de sincronización</h3><span class="result-count">Sin ejecutar</span></div>
            <div class="empty-results"><strong>Sin cambios registrados</strong><span>El resumen aparecerá aquí después de sincronizar.</span></div>
          </section>
        </div>
      </section>`;
  }

  function updateSyncButton() {
    const button = document.getElementById("proveedor-sync");
    const url = document.getElementById("proveedor-onedrive-url")?.value.trim();
    if (button) button.disabled = !(selectedBaseFile && url);
  }

  async function testLink() {
    const input = document.getElementById("proveedor-onedrive-url");
    const statusEl = document.getElementById("proveedor-link-status");
    const btn = document.getElementById("proveedor-test-link");
    const url = input?.value.trim();
    if (!url) {
      if (statusEl) statusEl.textContent = "Ingresa primero el enlace compartido de OneDrive.";
      return;
    }
    localStorage.setItem(LINK_STORAGE_KEY, url);
    btn.disabled = true;
    btn.textContent = "Verificando...";
    if (statusEl) statusEl.textContent = "Conectando con el enlace de OneDrive y validando estructura de 41 columnas...";
    try {
      const info = await request(`${API}/test-link`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ oneDriveUrl: url })
      });
      if (statusEl) {
        statusEl.innerHTML = `<span style="color:var(--success); font-weight:600;">✓ Enlace válido:</span> ${esc(info.fileName)} (${Number(info.rowsCount).toLocaleString("es-PA")} filas, ${info.columnsCount} columnas).`;
      }
      if (window.QAMantoToast) window.QAMantoToast("Enlace de OneDrive verificado correctamente.", "success");
    } catch (error) {
      if (statusEl) {
        statusEl.innerHTML = `<span style="color:var(--danger); font-weight:600;">✗ Error:</span> ${esc(error.message)}`;
      }
      if (window.QAMantoToast) window.QAMantoToast(error.message, "error");
    } finally {
      btn.disabled = false;
      btn.textContent = "Verificar enlace";
      updateSyncButton();
    }
  }

  function renderReport(report) {
    const area = document.getElementById("proveedor-report");
    if (!area || !report) return;
    const rows = (report.changes || []).slice(0, 500).map(change => `
      <tr><td>${esc(change.action)}</td><td>${esc(change.key)}</td><td>${esc(change.fields?.join(", ") || "—")}</td></tr>`).join("");
    area.innerHTML = `
      <div class="results-head"><h3>Reporte de sincronización</h3><span class="result-count">${new Date(report.completedAt).toLocaleString("es-PA")}</span></div>
      <div class="proveedor-report-grid">
        <div><strong>${Number(report.added).toLocaleString("es-PA")}</strong><span>Agregados</span></div>
        <div><strong>${Number(report.updated).toLocaleString("es-PA")}</strong><span>Actualizados</span></div>
        <div><strong>${Number(report.deleted).toLocaleString("es-PA")}</strong><span>Eliminados</span></div>
        <div><strong>${Number(report.unchanged).toLocaleString("es-PA")}</strong><span>Sin cambios</span></div>
      </div>
      <div class="proveedor-report-meta">
        <strong>${esc(report.message || "Proceso terminado")}</strong><br>
        Archivo origen: ${esc(report.fileName || "—")}<br>
        Respaldo guardado en servidor: ${esc(report.backupName || "—")}<br>
        Filas base: ${Number(report.totalBase || 0).toLocaleString("es-PA")} · Filas en OneDrive antes de sincronizar: ${Number(report.totalDestino || 0).toLocaleString("es-PA")}
      </div>
      ${rows ? `<div class="table-wrap"><table class="proveedor-report-table"><thead><tr><th>Acción</th><th>Clave</th><th>Columnas modificadas</th></tr></thead><tbody>${rows}</tbody></table></div>` : '<div class="empty-results">No se detectaron cambios.</div>'}
      ${(report.changes || []).length > 500 ? `<p class="field-help">Se muestran 500 de ${(report.changes || []).length.toLocaleString()} movimientos.</p>` : ""}`;
  }

  async function sync() {
    const url = document.getElementById("proveedor-onedrive-url")?.value.trim();
    if (!selectedBaseFile || !url) return;
    localStorage.setItem(LINK_STORAGE_KEY, url);

    const button = document.getElementById("proveedor-sync");
    const progress = document.getElementById("proveedor-progress");
    button.disabled = true;
    button.textContent = "Sincronizando...";
    progress.textContent = "Descargando Excel desde el enlace de OneDrive y comparando registros...";

    try {
      const form = new FormData();
      form.append("baseExcel", selectedBaseFile);
      form.append("oneDriveUrl", url);
      const report = await request(`${API}/sync-link`, { method: "POST", body: form });
      renderReport(report);
      progress.textContent = report.message;

      if (report.updatedFileBase64) {
        downloadBase64Xlsx(report.updatedFileBase64, report.fileName);
      }

      if (window.QAMantoToast) window.QAMantoToast(report.message, "success");
    } catch (error) {
      progress.textContent = error.message;
      if (window.QAMantoToast) window.QAMantoToast(error.message, "error");
      else alert(error.message);
    } finally {
      button.textContent = "Sincronizar con enlace de OneDrive";
      updateSyncButton();
    }
  }

  async function init() {
    document.getElementById("proveedor-back")?.addEventListener("click", () => window.QAMantoNavigateDashboard?.());
    document.getElementById("proveedor-test-link")?.addEventListener("click", testLink);
    document.getElementById("proveedor-onedrive-url")?.addEventListener("input", event => {
      localStorage.setItem(LINK_STORAGE_KEY, event.target.value.trim());
      updateSyncButton();
    });
    document.getElementById("proveedor-base-file")?.addEventListener("change", event => {
      selectedBaseFile = event.target.files?.[0] || null;
      const name = document.getElementById("proveedor-base-name");
      if (name) name.textContent = selectedBaseFile ? selectedBaseFile.name : "Ningún archivo seleccionado";
      updateSyncButton();
    });
    document.getElementById("proveedor-sync")?.addEventListener("click", sync);

    // Cargar enlace por defecto desde .env si existe y no hay uno guardado
    const input = document.getElementById("proveedor-onedrive-url");
    if (input && !input.value) {
      try {
        const cfg = await request(`${API}/config`);
        if (cfg.defaultOneDriveUrl) {
          input.value = cfg.defaultOneDriveUrl;
          localStorage.setItem(LINK_STORAGE_KEY, cfg.defaultOneDriveUrl);
        }
      } catch {}
    }
    updateSyncButton();
  }

  window.ProveedorModule = { render, init };
})();
