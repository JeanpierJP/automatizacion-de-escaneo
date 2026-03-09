import { useEffect, useMemo, useState } from 'react';
import * as XLSX from 'xlsx';
import { Toaster, toast } from 'sonner';
import {
  ArrowDownTrayIcon,
  ChartBarSquareIcon,
  CheckIcon,
  CheckCircleIcon,
  CloudArrowUpIcon,
  DocumentMagnifyingGlassIcon,
  DocumentTextIcon,
  MoonIcon,
  PencilSquareIcon,
  QueueListIcon,
  SunIcon,
  XMarkIcon,
  XCircleIcon,
} from '@heroicons/react/24/outline';
import api from './api';
import './App.css';

const asBool = (value) => {
  if (typeof value === 'boolean') return value;
  if (typeof value === 'number') return value > 0;
  if (typeof value === 'string') {
    const normalized = value.trim().toLowerCase();
    return ['si', 'sí', 'yes', 'true', '1', 'ok', 'presente'].includes(normalized);
  }
  return false;
};

const normalizeShiftLabel = (value) => {
  const text = (value || '').toString().trim().toUpperCase();
  if (!text) return '';
  if (text === 'D' || text === 'DIA' || text === 'DIURNO') return 'D';
  if (text === 'N' || text === 'NOCHE' || text === 'NOCTURNO') return 'N';
  return text;
};

const isDayShift = (value) => normalizeShiftLabel(value) === 'D';
const isNightShift = (value) => normalizeShiftLabel(value) === 'N';

const inferShift = (entryTime, existingShift) => {
  const normalizedExisting = normalizeShiftLabel(existingShift);
  if (normalizedExisting) return normalizedExisting;
  if (!entryTime || typeof entryTime !== 'string') return 'No definido';
  const [hoursText] = entryTime.split(':');
  const hour = Number(hoursText);
  if (Number.isNaN(hour)) return 'No definido';
  return hour >= 6 && hour < 19 ? 'D' : 'N';
};

const normalizeRecord = (item, index) => {
  const workerName =
    item.worker_name || item.employee_name || item.nombre_trabajador || item.provider_name || '-';
  const dni = item.dni || item.document_number || item.provider_ruc || '-';
  const date = item.date || item.work_date || item.issue_date || '-';
  const entryTime = item.entry_time || item.check_in || item.hora_entrada || '-';
  const exitTime = item.exit_time || item.check_out || item.hora_salida || '-';
  const shift = inferShift(entryTime, item.shift || item.turno);
  const hasSignature = asBool(item.signature_present ?? item.has_signature ?? item.firma_presente);

  return {
    id: item.id ?? `${dni}-${date}-${index}`,
    workerName,
    dni,
    date,
    entryTime,
    exitTime,
    shift,
    hasSignature,
  };
};

function App() {
  const [uploadingSingle, setUploadingSingle] = useState(false);
  const [uploadingBulk, setUploadingBulk] = useState(false);
  const [topView, setTopView] = useState('registros');
  const [records, setRecords] = useState([]);
  const [editingId, setEditingId] = useState(null);
  const [editForm, setEditForm] = useState({});

  const fetchRecords = async () => {
    try {
      const response = await api.get('/invoices');
      const data = Array.isArray(response.data) ? response.data : [];
      setRecords(data.map(normalizeRecord));
    } catch (error) {
      toast.error('No se pudo cargar el historial', {
        description: 'Verifica que el backend este activo.',
      });
    }
  };

  useEffect(() => {
    fetchRecords();
  }, []);

  const uploadPdf = (file) => {
    const formData = new FormData();
    formData.append('file', file);
    return api.post('/upload', formData, {
      headers: { 'Content-Type': 'multipart/form-data' },
    });
  };

  const handleSingleFileChange = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;

    setUploadingSingle(true);

    await toast.promise(
      uploadPdf(file),
      {
        loading: 'Procesando PDF...',
        success: () => {
          fetchRecords();
          return 'Documento procesado correctamente.';
        },
        error: (errorResponse) => {
          return errorResponse.response?.data?.detail || 'Error al procesar el archivo.';
        },
      }
    );

    event.target.value = '';
    setUploadingSingle(false);
  };

  const handleMultiFileChange = async (event) => {
    const files = Array.from(event.target.files || []);
    if (files.length === 0) return;

    const pdfFiles = files.filter(
      (file) => file.type === 'application/pdf' || file.name.toLowerCase().endsWith('.pdf')
    );

    if (pdfFiles.length === 0) {
      toast.error('Selecciona archivos PDF validos.');
      event.target.value = '';
      return;
    }

    const selected = pdfFiles.slice(0, 15);
    if (pdfFiles.length > 15) {
      toast.warning('Solo se procesaran los primeros 15 PDFs.');
    }

    setUploadingBulk(true);

    const bulkPromise = Promise.allSettled(selected.map((file) => uploadPdf(file))).then((results) => {
      const successCount = results.filter((item) => item.status === 'fulfilled').length;
      const failedCount = results.length - successCount;
      if (successCount === 0) {
        throw new Error('No se pudo procesar ningun PDF.');
      }
      return { successCount, failedCount };
    });

    await toast.promise(bulkPromise, {
      loading: `Procesando ${selected.length} PDFs...`,
      success: ({ successCount, failedCount }) => {
        fetchRecords();
        return failedCount > 0
          ? `${successCount} PDF(s) procesados, ${failedCount} con error.`
          : `${successCount} PDF(s) procesados correctamente.`;
      },
      error: (errorResponse) => {
        return errorResponse?.message || 'Error en la carga masiva.';
      },
    });

    event.target.value = '';
    setUploadingBulk(false);
  };

  const exportToExcel = () => {
    const rows = records.map((record) => ({
      Nombre_Trabajador: record.workerName,
      DNI: record.dni,
      Fecha: record.date,
      Hora_Entrada: record.entryTime,
      Hora_Salida: record.exitTime,
      Turno: record.shift,
      Firma: record.hasSignature ? 'Si' : 'No',
    }));

    const worksheet = XLSX.utils.json_to_sheet(rows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Asistencia');
    XLSX.writeFile(workbook, 'Reporte_Asistencia.xlsx');
    toast.success('Excel exportado');
  };

  const stats = useMemo(() => {
    const signatureOk = records.filter((item) => item.hasSignature).length;
    const dayShift = records.filter((item) => isDayShift(item.shift)).length;
    const nightShift = records.filter((item) => isNightShift(item.shift)).length;
    return { signatureOk, dayShift, nightShift };
  }, [records]);

  const isBusy = uploadingSingle || uploadingBulk;
  const isEditMode = topView === 'editar';

  const startEdit = (record) => {
    setEditingId(record.id);
    setEditForm({ ...record });
  };

  const cancelEdit = () => {
    setEditingId(null);
    setEditForm({});
  };

  const saveEdit = () => {
    setRecords((prev) =>
      prev.map((item) =>
        item.id === editingId
          ? {
              ...item,
              ...editForm,
              shift: inferShift(editForm.entryTime, editForm.shift),
            }
          : item
      )
    );
    toast.success('Registro actualizado');
    setEditingId(null);
    setEditForm({});
  };

  const handleEditChange = (field, value) => {
    setEditForm((prev) => ({ ...prev, [field]: value }));
  };

  return (
    <div className="min-h-screen app-surface p-4 md:p-8">
      <Toaster position="top-center" richColors />
      <div className="mx-auto w-full max-w-7xl space-y-6">
        <section className="hero-panel rounded-3xl p-6 md:p-10">
          <div className="hero-heading-row">
            <h1 className="text-xl font-bold text-slate-900 md:text-3xl">Control de Asistencia por PDF</h1>
            <div className="hero-topbar">
              <div className="view-switcher">
                <button
                  type="button"
                  onClick={() => setTopView('registros')}
                  className={`view-pill ${topView === 'registros' ? 'view-pill-active' : ''}`}
                >
                  <DocumentTextIcon className="h-4 w-4" />
                  Registros PDF
                </button>
                <button
                  type="button"
                  onClick={() => setTopView('editar')}
                  className={`view-pill ${topView === 'editar' ? 'view-pill-active' : ''}`}
                >
                  <QueueListIcon className="h-4 w-4" />
                  Editar Registros
                </button>
              </div>
            </div>
          </div>

          <div>
            <p className="mt-3 max-w-4xl text-sm text-slate-700 md:text-[1.05rem]">
              Plataforma para extraccion automatica de datos laborales
            </p>
          </div>

          <div className="upload-grid mt-6">
            <article className="upload-card upload-card-single">
              <div className="upload-content">
                <div className="upload-icon-wrap">
                  <CloudArrowUpIcon className="h-6 w-6 text-indigo-600" />
                </div>
                <div>
                  <h3 className="upload-title">Cargar Registro Individual</h3>
                  <p className="upload-subtitle">Sube 1 PDF para procesarlo</p>
                </div>
              </div>
              <label className={`upload-action upload-action-single ${isBusy ? 'opacity-50 pointer-events-none' : ''}`}>
                {uploadingSingle ? 'Procesando...' : 'Seleccionar PDF'}
                <input
                  type="file"
                  accept="application/pdf"
                  onChange={handleSingleFileChange}
                  className="hidden"
                  disabled={isBusy}
                />
              </label>
            </article>

            <article className="upload-card upload-card-mass">
              <div className="upload-content">
                <div className="upload-icon-wrap upload-icon-wrap-mass">
                  <QueueListIcon className="h-6 w-6 text-emerald-600" />
                </div>
                <div>
                  <h3 className="upload-title">Carga Masiva</h3>
                  <p className="upload-subtitle upload-subtitle-mass">Sube hasta 15 PDFs simultaneamente</p>
                </div>
              </div>
              <label className={`upload-action upload-action-mass ${isBusy ? 'opacity-50 pointer-events-none' : ''}`}>
                {uploadingBulk ? 'Procesando...' : 'Seleccionar Multiples'}
                <input
                  type="file"
                  accept="application/pdf"
                  multiple
                  onChange={handleMultiFileChange}
                  className="hidden"
                  disabled={isBusy}
                />
              </label>
            </article>
          </div>
        </section>

        <section className="grid gap-4 md:grid-cols-2 xl:grid-cols-4">
          <article className="glass-panel">
            <div className="flex items-center justify-between">
              <p className="panel-label">Registros cargados</p>
              <ChartBarSquareIcon className="h-6 w-6 text-amber-700" />
            </div>
            <p className="panel-value">{records.length}</p>
          </article>
          <article className="glass-panel">
            <div className="flex items-center justify-between">
              <p className="panel-label">Turno dia</p>
              <SunIcon className="h-6 w-6 text-orange-600" />
            </div>
            <p className="panel-value">{stats.dayShift}</p>
          </article>
          <article className="glass-panel">
            <div className="flex items-center justify-between">
              <p className="panel-label">Turno nocturno</p>
              <MoonIcon className="h-6 w-6 text-slate-800" />
            </div>
            <p className="panel-value">{stats.nightShift}</p>
          </article>
          <article className="glass-panel">
            <div className="flex items-center justify-between">
              <p className="panel-label">Firmas validas</p>
              <CheckCircleIcon className="h-6 w-6 text-emerald-700" />
            </div>
            <p className="panel-value">{stats.signatureOk}</p>
          </article>
        </section>

        <section className="table-shell overflow-hidden rounded-3xl">
          <div className="flex flex-col gap-4 border-b border-slate-200/70 bg-white/70 p-4 md:flex-row md:items-center md:justify-between">
            <h2 className="flex items-center gap-2 text-lg font-semibold text-slate-800">
              <DocumentMagnifyingGlassIcon className="h-5 w-5" />
              {isEditMode ? `Edicion de Registros (${records.length})` : 'Datos extraidos por documento'}
            </h2>
            <button onClick={exportToExcel} className="export-button">
              <ArrowDownTrayIcon className="h-4 w-4" />
              Exportar Excel
            </button>
          </div>

          <div className="max-h-[560px] overflow-auto bg-white/80">
            <table className="w-full min-w-[920px] text-sm">
              <thead className="sticky top-0 bg-[#f8f4eb] text-xs uppercase tracking-wide text-slate-600">
                <tr>
                  <th className="px-4 py-3 text-left">Trabajador</th>
                  <th className="px-4 py-3 text-left">DNI</th>
                  <th className="px-4 py-3 text-left">Fecha</th>
                  <th className="px-4 py-3 text-left">Entrada</th>
                  <th className="px-4 py-3 text-left">Salida</th>
                  <th className="px-4 py-3 text-left">Turno</th>
                  <th className="px-4 py-3 text-left">Firma</th>
                  {isEditMode && <th className="px-4 py-3 text-left">Accion</th>}
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {records.length === 0 ? (
                  <tr>
                    <td className="px-4 py-10 text-center text-slate-500" colSpan={isEditMode ? 8 : 7}>
                      Aun no hay registros. Sube el primer PDF para comenzar.
                    </td>
                  </tr>
                ) : (
                  records.map((record) => (
                    <tr key={record.id} className="hover:bg-[#fff8ef] transition-colors">
                      <td className="px-4 py-3 font-semibold text-slate-800">
                        {editingId === record.id ? (
                          <input
                            type="text"
                            className="edit-input"
                            value={editForm.workerName || ''}
                            onChange={(e) => handleEditChange('workerName', e.target.value)}
                          />
                        ) : (
                          record.workerName
                        )}
                      </td>
                      <td className="px-4 py-3 text-slate-700">
                        {editingId === record.id ? (
                          <input
                            type="text"
                            className="edit-input"
                            value={editForm.dni || ''}
                            onChange={(e) => handleEditChange('dni', e.target.value)}
                          />
                        ) : (
                          record.dni
                        )}
                      </td>
                      <td className="px-4 py-3 text-slate-700">
                        {editingId === record.id ? (
                          <input
                            type="text"
                            className="edit-input"
                            value={editForm.date || ''}
                            onChange={(e) => handleEditChange('date', e.target.value)}
                          />
                        ) : (
                          record.date
                        )}
                      </td>
                      <td className="px-4 py-3 text-slate-700">
                        {editingId === record.id ? (
                          <input
                            type="text"
                            className="edit-input"
                            value={editForm.entryTime || ''}
                            onChange={(e) => handleEditChange('entryTime', e.target.value)}
                          />
                        ) : (
                          record.entryTime
                        )}
                      </td>
                      <td className="px-4 py-3 text-slate-700">
                        {editingId === record.id ? (
                          <input
                            type="text"
                            className="edit-input"
                            value={editForm.exitTime || ''}
                            onChange={(e) => handleEditChange('exitTime', e.target.value)}
                          />
                        ) : (
                          record.exitTime
                        )}
                      </td>
                      <td className="px-4 py-3">
                        {(() => {
                          const currentShift =
                            editingId === record.id ? inferShift(editForm.entryTime, editForm.shift) : record.shift;
                          return (
                        <span
                          className={`chip ${
                            isNightShift(currentShift)
                              ? 'chip-night'
                              : 'chip-day'
                          }`}
                        >
                          {currentShift}
                        </span>
                          );
                        })()}
                      </td>
                      <td className="px-4 py-3">
                        {(editingId === record.id ? editForm.hasSignature : record.hasSignature) ? (
                          <span className="chip chip-sign-ok">
                            <CheckCircleIcon className="h-4 w-4" />
                            Si
                          </span>
                        ) : (
                          <span className="chip chip-sign-no">
                            <XCircleIcon className="h-4 w-4" />
                            No
                          </span>
                        )}
                      </td>
                      {isEditMode && (
                        <td className="px-4 py-3">
                          {editingId === record.id ? (
                            <div className="action-group">
                              <button type="button" onClick={saveEdit} className="action-btn action-btn-save">
                                <CheckIcon className="h-4 w-4" />
                              </button>
                              <button type="button" onClick={cancelEdit} className="action-btn action-btn-cancel">
                                <XMarkIcon className="h-4 w-4" />
                              </button>
                            </div>
                          ) : (
                            <button type="button" onClick={() => startEdit(record)} className="action-btn action-btn-edit">
                              <PencilSquareIcon className="h-5 w-5" />
                            </button>
                          )}
                        </td>
                      )}
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        </section>
      </div>
    </div>
  );
}

export default App;
