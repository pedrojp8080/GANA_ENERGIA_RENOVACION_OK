import React, { useMemo, useState } from 'react';
import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';

const BROWSER_APP_URL = 'https://externos.ganaenergia.com';
const CONTRACT_URL = `${BROWSER_APP_URL}/contract`;

async function saveFile({ fileName, mimeType, description, extensions, content }) {
  if (typeof window !== 'undefined' && 'showSaveFilePicker' in window) {
    const handle = await window.showSaveFilePicker({
      suggestedName: fileName,
      types: [
        {
          description,
          accept: {
            [mimeType]: extensions
          }
        }
      ]
    });

    const writable = await handle.createWritable();
    await writable.write(content);
    await writable.close();
    return;
  }

  const blob = content instanceof Blob ? content : new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

async function saveWorkbook(workbook, fileName) {
  const fileBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
  await saveFile({
    fileName,
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    description: 'Excel Workbook',
    extensions: ['.xlsx'],
    content: fileBuffer
  });
}

async function saveTextFile(text, fileName) {
  await saveFile({
    fileName,
    mimeType: 'text/plain',
    description: 'Archivo de texto',
    extensions: ['.txt'],
    content: text
  });
}

function isSaveDialogCancelled(error) {
  return error?.name === 'AbortError';
}

const TEMPLATE_HEADERS = [
  'fecha_creacion',
  'bonoSocialAplicable',
  'titular_antiguo',
  'fecha_tramitacion',
  'promocion',
  'gas',
  'nacionalidad',
  'origen_contrato',
  'id_externo',
  'suministro_activo',
  'titular.razon_social',
  'titular.contacto_nombre',
  'titular.contacto_documento',
  'titular.documento_antiguo',
  'titular.fijo',
  'titular.tipo',
  'titular.nombre',
  'titular.apellido1',
  'titular.apellido2',
  'titular.documento',
  'titular.fijo',
  'titular.movil',
  'titular.prefijo',
  'titular.email',
  'titular.cnae',
  'titular.comerciales',
  'suministro.duplicador',
  'suministro.escalera',
  'suministro.planta',
  'suministro.puerta',
  'suministro.tipo_aclarador',
  'suministro.aclarador',
  'suministro.observaciones',
  'suministro.tipo_via',
  'suministro.nombre_via',
  'suministro.numero',
  'suministro.localidad_ine',
  'suministro.provincia_ine',
  'suministro.localidad',
  'suministro.provincia',
  'suministro.cp',
  'suministro.cups',
  'suministro.tarifa_id',
  'suministro.tarifa_tipo',
  'suministro.atr',
  'suministro.tarifa_papel',
  'suministro.potencia_contratada_array',
  'suministro.precio_potencia',
  'suministro.cliente_paga',
  'suministro.codigo_cambio',
  'datos_bancarios.domiciliado',
  'datos_bancarios.titular',
  'datos_bancarios.documento',
  'datos_bancarios.iban'
];

const TEMPLATE_EXAMPLES = [
  {
    fecha_creacion: '10/03/2026',
    bonoSocialAplicable: '',
    titular_antiguo: '',
    fecha_tramitacion: '',
    promocion: '',
    gas: '',
    nacionalidad: '',
    origen_contrato: 0,
    id_externo: 'data',
    suministro_activo: true,
    'titular.razon_social': '',
    'titular.contacto_nombre': '',
    'titular.contacto_documento': '',
    'titular.documento_antiguo': '',
    'titular.fijo': '',
    'titular.tipo': 0,
    'titular.nombre': 'ANA MARIA',
    'titular.apellido1': 'PAY',
    'titular.apellido2': 'LUNA',
    'titular.documento': '77504220Q',
    'titular.movil': '654647851',
    'titular.prefijo': '34',
    'titular.email': 'jose.piber@gmail.com',
    'titular.cnae': '9820',
    'titular.comerciales': false,
    'suministro.duplicador': '',
    'suministro.escalera': '',
    'suministro.planta': '',
    'suministro.puerta': '',
    'suministro.tipo_aclarador': '',
    'suministro.aclarador': '',
    'suministro.observaciones': '',
    'suministro.tipo_via': 'CL',
    'suministro.nombre_via': 'FRANCISCO RABAL',
    'suministro.numero': '8',
    'suministro.localidad_ine': 600,
    'suministro.provincia_ine': 30,
    'suministro.localidad': 'ARCHENA',
    'suministro.provincia': 'MURCIA',
    'suministro.cp': '30600',
    'suministro.cups': 'ES0021000005847734DG',
    'suministro.tarifa_id': '695389679d3a42b3e194c473',
    'suministro.tarifa_tipo': '2.0TD',
    'suministro.atr': 18,
    'suministro.tarifa_papel': false,
    'suministro.potencia_contratada_array': '["3.3","3.3"]',
    'suministro.precio_potencia': '10.94',
    'suministro.cliente_paga': 2,
    'suministro.codigo_cambio': 'CZ',
    'datos_bancarios.domiciliado': true,
    'datos_bancarios.titular': 'ANA MARIA PAY LUNA',
    'datos_bancarios.documento': '77504220Q',
    'datos_bancarios.iban': 'ES3821008184091300490814'
  }
];

function setNestedValue(obj, path, value) {
  const keys = path.split('.');
  let current = obj;
  for (let index = 0; index < keys.length - 1; index += 1) {
    const key = keys[index];
    if (!current[key] || typeof current[key] !== 'object' || Array.isArray(current[key])) {
      current[key] = {};
    }
    current = current[key];
  }
  current[keys[keys.length - 1]] = value;
}

function normalizeBoolean(value) {
  if (typeof value === 'boolean') return value;
  if (typeof value === 'number') return value === 1;
  if (typeof value !== 'string') return value;
  const normalized = value.trim().toLowerCase();
  if (['true', '1', 'si', 'sí', 'yes', 'y'].includes(normalized)) return true;
  if (['false', '0', 'no', 'n'].includes(normalized)) return false;
  return value;
}

function normalizeNumber(value) {
  if (typeof value === 'number') return value;
  if (typeof value !== 'string') return value;
  const normalized = value.trim();
  if (normalized === '') return '';
  if (/^-?\d+$/.test(normalized)) return Number(normalized);
  return value;
}

function normalizeArray(value) {
  if (Array.isArray(value)) {
    return value.map((item) => String(item ?? '').trim()).filter(Boolean);
  }

  if (typeof value === 'string') {
    const normalized = value.trim();
    if (!normalized) return [];

    try {
      const parsed = JSON.parse(normalized);
      if (Array.isArray(parsed)) {
        return parsed.map((item) => String(item ?? '').trim()).filter(Boolean);
      }
    } catch {
      // Permite texto simple como "3.3,3.3".
    }

    return normalized
      .split(/[;,|]/)
      .map((item) => item.trim())
      .filter(Boolean);
  }

  if (value === undefined || value === null || value === '') return [];
  return [String(value)];
}

function cleanValue(key, value) {
  if (value === undefined || value === null) return '';
  if (typeof value === 'string') value = value.trim();

  const lowerKey = key.toLowerCase();
  const booleanKeys = ['suministro_activo', 'comerciales', 'domiciliado', 'tarifa_papel'];
  const numberKeys = ['origen_contrato', 'tipo', 'localidad_ine', 'provincia_ine', 'atr', 'cliente_paga'];

  if (lowerKey.endsWith('potencia_contratada_array')) return normalizeArray(value);
  if (booleanKeys.some((token) => lowerKey.includes(token))) return normalizeBoolean(value);
  if (numberKeys.some((token) => lowerKey.endsWith(token))) return normalizeNumber(value);
  return value;
}

function shouldIncludePayloadValue(value) {
  if (value === undefined || value === null) return false;
  if (typeof value === 'string') return value.trim() !== '';
  if (Array.isArray(value)) return value.length > 0;
  return true;
}

function rowToPayload(row) {
  const payload = {};
  Object.entries(row).forEach(([rawKey, rawValue]) => {
    const key = String(rawKey ?? '').trim();
    if (!key || key === 'fecha_creacion') return;
    const value = cleanValue(key, rawValue);
    if (!shouldIncludePayloadValue(value)) return;
    setNestedValue(payload, key, value);
  });
  return payload;
}

function getWorksheetHeaders(worksheet) {
  const [headers = []] = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: '',
    blankrows: false,
    raw: false
  });

  return headers.map((header) => String(header ?? '').trim()).filter(Boolean);
}

function compareHeaders(headers) {
  const missing = TEMPLATE_HEADERS.filter((header) => !headers.includes(header));
  const unexpected = headers.filter((header) => !TEMPLATE_HEADERS.includes(header));
  const sameOrder =
    headers.length === TEMPLATE_HEADERS.length &&
    headers.every((header, index) => header === TEMPLATE_HEADERS[index]);

  return {
    missing,
    unexpected,
    sameOrder,
    valid: missing.length === 0 && unexpected.length === 0
  };
}

function formatDateAsIso(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function normalizeExcelDate(value) {
  if (value == null || value === '') return '';
  if (value instanceof Date && !Number.isNaN(value.getTime())) return formatDateAsIso(value);

  if (typeof value === 'number') {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (!parsed) return '';
    return `${parsed.y}-${String(parsed.m).padStart(2, '0')}-${String(parsed.d).padStart(2, '0')}`;
  }

  const text = String(value).trim();
  if (!text) return '';

  const ddmmyyyy = text.split(' ')[0].match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (ddmmyyyy) return `${ddmmyyyy[3]}-${ddmmyyyy[2]}-${ddmmyyyy[1]}`;

  const iso = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return `${iso[1]}-${iso[2]}-${iso[3]}`;

  const parsed = new Date(text);
  if (Number.isNaN(parsed.getTime())) return '';
  return formatDateAsIso(parsed);
}

function buildConsoleScript(rowsToProcess, modeLabel) {
  const payloads = rowsToProcess.map((row, index) => ({
    rowNumber: index + 2,
    idExterno: String(row.id_externo ?? '').trim(),
    fechaCreacion: normalizeExcelDate(row.fecha_creacion),
    payload: rowToPayload(row)
  }));

  const serializedPayloads = JSON.stringify(payloads, null, 2);

  return `(() => {
  const BASE_URL = '${BROWSER_APP_URL}';
  const CONTRACT_URL = BASE_URL + '/contract';
  const MODE_LABEL = ${JSON.stringify(modeLabel)};
  const PAYLOADS = ${serializedPayloads};

  function firstNonEmpty(...values) {
    for (const value of values) {
      if (value === undefined || value === null) continue;
      if (typeof value === 'string' && value.trim() === '') continue;
      return value;
    }
    return '';
  }

  function extractContractData(data) {
    if (typeof data === 'string' || typeof data === 'number') {
      return {
        id: String(data),
        error: false,
        message: ''
      };
    }

    const res = data?.res || data || {};

    return {
      id: String(firstNonEmpty(
        res?.id_contrato,
        data?.id_contrato,
        res?.contrato_id,
        data?.contrato_id,
        res?._id,
        res?.id,
        res?.contractId,
        res?.contract_id,
        data?.contractId,
        data?.contract_id,
        data?.id
      ) || ''),
      error: Boolean(data?.error),
      message: data?.msg?.message || data?.message || ''
    };
  }

  function getStorageToken(storage) {
    if (!storage) return '';
    const exactKeys = ['token', 'access_token', 'jwt', 'authToken', 'authorization'];

    for (const key of exactKeys) {
      const rawValue = storage.getItem(key);
      if (rawValue) return rawValue.replace(/^Bearer\\s+/i, '').trim();
    }

    for (let index = 0; index < storage.length; index += 1) {
      const key = storage.key(index);
      if (!key) continue;
      const normalizedKey = key.toLowerCase();
      if (!normalizedKey.includes('token') && !normalizedKey.includes('auth') && !normalizedKey.includes('jwt')) {
        continue;
      }

      const rawValue = storage.getItem(key);
      if (!rawValue) continue;

      try {
        const parsed = JSON.parse(rawValue);
        const nestedValue = parsed?.token || parsed?.access_token || parsed?.jwt || parsed?.authorization;
        if (nestedValue) return String(nestedValue).replace(/^Bearer\\s+/i, '').trim();
      } catch {
        return rawValue.replace(/^Bearer\\s+/i, '').trim();
      }
    }

    return '';
  }

  async function parseResponse(response) {
    const rawText = await response.text().catch(() => '');
    if (!rawText) return { parsed: {}, rawText: '' };

    try {
      return { parsed: JSON.parse(rawText), rawText };
    } catch {
      return { parsed: {}, rawText };
    }
  }

  function escapeHtml(value) {
    return String(value ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function downloadResultsExcel(results) {
    const fileName = 'resultado_' + MODE_LABEL.replace(/[^a-z0-9_-]+/gi, '_').toLowerCase() + '.xls';
    const headers = [
      'index',
      'rowNumber',
      'id_externo',
      'fecha_creacion',
      'estado_api',
      'contrato_id',
      'error',
      'respuesta_body',
      'payload_enviado'
    ];

    const rows = results.map((row) => (
      '<tr>' +
      headers.map((header) => '<td>' + escapeHtml(row[header] ?? '') + '</td>').join('') +
      '</tr>'
    )).join('');

    const html = '<html><head><meta charset="utf-8"></head><body><table border="1"><thead><tr>' +
      headers.map((header) => '<th>' + escapeHtml(header) + '</th>').join('') +
      '</tr></thead><tbody>' + rows + '</tbody></table></body></html>';

    const blob = new Blob([html], { type: 'application/vnd.ms-excel' });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = fileName;
    document.body.appendChild(anchor);
    anchor.click();
    document.body.removeChild(anchor);
    URL.revokeObjectURL(url);
  }

  async function run() {
    if (!Array.isArray(PAYLOADS) || PAYLOADS.length === 0) {
      console.warn('[Gana Energia] No hay filas para procesar en este script.');
      return;
    }

    if (!window.location.origin.includes('ganaenergia.com')) {
      console.warn('[Gana Energia] Abre la web de Gana Energia con la sesion iniciada antes de pegar este codigo.');
      return;
    }

    const token = getStorageToken(window.localStorage) || getStorageToken(window.sessionStorage);
    const manualToken = token || window.prompt(
      'No se encontro token automaticamente. Si tu sesion usa Bearer token, pegalo aqui. Si no, deja el campo vacio para probar con cookies.',
      ''
    );

    const headers = { 'Content-Type': 'application/json' };
    if (manualToken && manualToken.trim()) {
      headers.Authorization = 'Bearer ' + manualToken.replace(/^Bearer\\s+/i, '').trim();
    }

    const results = [];
    console.info('[Gana Energia] Inicio modo:', MODE_LABEL, '- contratos:', PAYLOADS.length);

    for (let index = 0; index < PAYLOADS.length; index += 1) {
      const item = PAYLOADS[index];

      try {
        const response = await fetch(CONTRACT_URL, {
          method: 'POST',
          credentials: 'include',
          headers,
          body: JSON.stringify(item.payload)
        });

        const { parsed, rawText } = await parseResponse(response);
        const contract = extractContractData(parsed);
        const ok = response.ok && !contract.error;

        const rowResult = {
          index: index + 1,
          rowNumber: item.rowNumber,
          id_externo: item.idExterno || '',
          fecha_creacion: item.fechaCreacion || '',
          estado_api: ok ? 'OK' : 'ERROR',
          contrato_id: ok ? contract.id : '',
          error: ok ? '' : (rawText || contract.message || ('Error HTTP ' + response.status)),
          respuesta_body: rawText || JSON.stringify(parsed || {}),
          payload_enviado: JSON.stringify(item.payload)
        };

        results.push(rowResult);
        console.log('[Gana Energia][' + (index + 1) + '/' + PAYLOADS.length + ']', rowResult);
      } catch (error) {
        const rowResult = {
          index: index + 1,
          rowNumber: item.rowNumber,
          id_externo: item.idExterno || '',
          fecha_creacion: item.fechaCreacion || '',
          estado_api: 'ERROR',
          contrato_id: '',
          error: error?.message || 'Error de red',
          respuesta_body: '',
          payload_enviado: JSON.stringify(item.payload)
        };

        results.push(rowResult);
        console.error('[Gana Energia][' + (index + 1) + '/' + PAYLOADS.length + ']', rowResult);
      }
    }

    window.__ganaenergiaLastResults = results;
    console.table(results);
    downloadResultsExcel(results);
    console.info('[Gana Energia] Fin. Resultado guardado en window.__ganaenergiaLastResults');
  }

  run();
})();`;
}

export default function App() {
  const [fileName, setFileName] = useState('');
  const [rows, setRows] = useState([]);
  const [appError, setAppError] = useState('');
  const [templateCheck, setTemplateCheck] = useState(null);
  const [scriptPreview, setScriptPreview] = useState('');
  const systemDate = useMemo(() => formatDateAsIso(new Date()), []);

  const rowsForSystemDate = useMemo(
    () => rows.filter((row) => normalizeExcelDate(row.fecha_creacion) === systemDate),
    [rows, systemDate]
  );

  const allRowsScript = useMemo(
    () => (rows.length ? buildConsoleScript(rows, 'todas_las_filas') : ''),
    [rows]
  );

  const systemDateScript = useMemo(
    () => (rowsForSystemDate.length ? buildConsoleScript(rowsForSystemDate, `fecha_creacion_${systemDate}`) : ''),
    [rowsForSystemDate, systemDate]
  );

  async function handleExcelUpload(event) {
    const file = event.target.files?.[0];
    if (!file) return;

    setFileName(file.name);
    setRows([]);
    setAppError('');
    setTemplateCheck(null);
    setScriptPreview('');

    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: 'array' });
      const firstSheet = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheet];
      const headers = getWorksheetHeaders(worksheet);
      const jsonRows = XLSX.utils.sheet_to_json(worksheet, { defval: '', raw: false });

      setTemplateCheck(compareHeaders(headers));
      setRows(jsonRows);
    } catch {
      setAppError('No se pudo leer el Excel. Revisa que sea un archivo .xlsx o .xls valido.');
    }
  }

  async function downloadTemplate() {
    try {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('contracts', {
        views: [{ state: 'frozen', ySplit: 1 }]
      });

      worksheet.columns = TEMPLATE_HEADERS.map((header) => ({
        header,
        key: header,
        width: Math.max(header.length + 2, 16)
      }));

      TEMPLATE_EXAMPLES.forEach((row) => {
        worksheet.addRow(TEMPLATE_HEADERS.map((header) => row[header] ?? ''));
      });

      const lastColumn = worksheet.columnCount || TEMPLATE_HEADERS.length;
      const lastRow = Math.max(worksheet.rowCount, 1);
      worksheet.autoFilter = {
        from: { row: 1, column: 1 },
        to: { row: lastRow, column: lastColumn }
      };

      worksheet.getRow(1).font = { bold: true };
      const fileBuffer = await workbook.xlsx.writeBuffer();
      await saveFile({
        fileName: 'plantilla_ganaenergia.xlsx',
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        description: 'Excel Workbook',
        extensions: ['.xlsx'],
        content: fileBuffer
      });
    } catch (error) {
      if (!isSaveDialogCancelled(error)) {
        setAppError('No se pudo guardar la plantilla.');
      }
    }
  }

  async function handleDownloadScript(mode) {
    const script = mode === 'all' ? allRowsScript : systemDateScript;
    const targetRows = mode === 'all' ? rows.length : rowsForSystemDate.length;

    if (!targetRows || !script) {
      setAppError(
        mode === 'all'
          ? 'No hay filas cargadas para generar el TXT.'
          : `No hay filas con fecha_creacion igual a ${systemDate}.`
      );
      return;
    }

    const fileSuffix = mode === 'all' ? 'todas_las_filas' : `fecha_creacion_${systemDate}`;

    try {
      await saveTextFile(script, `pegar_en_consola_${fileSuffix}.txt`);
      setScriptPreview(script);
      setAppError('');
    } catch (error) {
      if (!isSaveDialogCancelled(error)) {
        setAppError('No se pudo guardar el TXT con el codigo de consola.');
      }
    }
  }

  return (
    <div className="app-shell">
      <div className="container">
        <header className="hero">
          <div>
            <h1>Importador de contratos Gana Energia</h1>
            <p>
              Sube el Excel y descarga un <strong>.txt</strong> con el codigo listo para pegar en la consola del
              navegador dentro de <strong>{BROWSER_APP_URL}</strong>. Hay dos modos: todas las filas o solo las que
              coinciden con <strong>fecha_creacion = {systemDate}</strong>.
            </p>
          </div>
          <button className="secondary" onClick={downloadTemplate}>Descargar plantilla ejemplo</button>
        </header>

        {appError && <div className="error-box">{appError}</div>}

        <section className="grid two-columns">
          <div className="card">
            <h2>1. Cargar Excel</h2>
            <p className="muted">La primera hoja debe usar los encabezados de la plantilla.</p>

            <div className="form">
              <div>
                <label>Plantilla Excel (.xlsx / .xls)</label>
                <input type="file" accept=".xlsx,.xls" onChange={handleExcelUpload} />
              </div>
              <div className="file-info">
                <div><strong>Archivo:</strong> {fileName || 'Ninguno'}</div>
                <div><strong>Filas:</strong> {rows.length}</div>
              </div>

              {templateCheck && (
                <div className={templateCheck.valid ? 'token-box' : 'warning-box'}>
                  <div className="token-title">Validacion de plantilla</div>
                  <div className="token-value">
                    {templateCheck.valid
                      ? templateCheck.sameOrder
                        ? 'Cabeceras correctas y en el mismo orden que la plantilla.'
                        : 'Cabeceras correctas, aunque el orden no coincide con la plantilla.'
                      : 'El Excel no coincide del todo con la plantilla esperada.'}
                  </div>
                  {!!templateCheck.missing.length && (
                    <div className="muted">Faltan: {templateCheck.missing.join(', ')}</div>
                  )}
                  {!!templateCheck.unexpected.length && (
                    <div className="muted">Sobran: {templateCheck.unexpected.join(', ')}</div>
                  )}
                </div>
              )}
            </div>
          </div>

          <div className="card">
            <h2>2. Descargar TXT</h2>
            <p className="muted">
              Cada TXT contiene el codigo autocontenido que crea todos los contratos al pegarlo en la consola del
              navegador con la sesion iniciada.
            </p>

            <div className="stats-grid">
              <div className="file-info">
                <div><strong>Total Excel:</strong> {rows.length}</div>
              </div>
              <div className="file-info">
                <div><strong>Fecha filtro:</strong> {systemDate}</div>
              </div>
              <div className="file-info">
                <div><strong>Coinciden por fecha:</strong> {rowsForSystemDate.length}</div>
              </div>
            </div>

            <div className="form">
              <button onClick={() => handleDownloadScript('all')} disabled={!rows.length}>
                Descargar TXT con todas las filas
              </button>
              <button
                className="secondary"
                onClick={() => handleDownloadScript('date')}
                disabled={!rowsForSystemDate.length}
              >
                Descargar TXT filtrado por fecha_creacion
              </button>
            </div>
          </div>
        </section>

        <section className="card">
          <h2>3. Como usar el TXT</h2>
          <ol>
            <li>Inicia sesion en <strong>{BROWSER_APP_URL}</strong>.</li>
            <li>Abre la pantalla de la web donde tengas la sesion activa.</li>
            <li>Abre la consola del navegador.</li>
            <li>Pega el contenido completo del TXT y pulsa Enter.</li>
            <li>El script ejecutara un POST por contrato y descargara un JSON con el resultado.</li>
          </ol>
        </section>

        <section className="card">
          <h2>4. Vista previa del codigo</h2>
          <p className="muted">
            Se muestra el ultimo script generado para que puedas revisarlo antes de pegarlo en la consola.
          </p>

          {!scriptPreview ? (
            <div className="empty-state">Todavia no se ha generado ningun TXT.</div>
          ) : (
            <details className="details-box" open>
              <summary>Ver codigo</summary>
              <pre>{scriptPreview}</pre>
            </details>
          )}
        </section>
      </div>
    </div>
  );
}
