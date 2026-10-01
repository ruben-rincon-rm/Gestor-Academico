import { initializeApp } from "firebase/app";
import { getAuth, signInAnonymously } from "firebase/auth";
import { 
    getFirestore, 
    collection, 
    doc, 
    setDoc, 
    getDoc, 
    getDocs, 
    updateDoc, 
    deleteDoc 
} from "firebase/firestore";
import * as XLSX from "xlsx";

// Interfaces
interface StudentRecord {
    id: string;
    matricula: string;
    tipo_documento: string;
    documento: string;
    nombres: string;
    apellidos: string;
    nombre_completo: string;
    nombre: string;
    grado: string;
    foto?: string;
    fotoTemp?: string;
    isStaff?: boolean;
}

interface StaffRecord {
    id: string;
    nombre: string;
    grado?: string;
    rol?: string;
    clave?: string;
    foto?: string;
}

interface InventoryItem {
    id: string;
    nombre: string;
}

interface ExcelValidationError {
    hoja: string;
    fila: number;
    campo: string;
}

declare global {
    interface Window {
        Html5Qrcode: any;
        QRCode: any;
        html2canvas: any;
    }
}

// Configuración Firebase
const firebaseConfig = { 
    apiKey: "AIzaSyBzLBoIhwBSBEolnbGJOPgaDorGRrVvt0k", 
    authDomain: "ie-ramon-munera-lopera.firebaseapp.com", 
    projectId: "ie-ramon-munera-lopera", 
    storageBucket: "ie-ramon-munera-lopera.firebasestorage.app", 
    messagingSenderId: "635796226651", 
    appId: "1:635796226651:web:97f862e1539ab98366757e", 
    measurementId: "G-X5L993QSR1" 
};

const app = initializeApp(firebaseConfig);
const auth = getAuth(app); 
const db = getFirestore(app);

// Rutas de Colecciones
const getSettingsPath = () => doc(db, 'configuracion_app', 'institucion');
const getStudentsPath = () => collection(db, 'estudiantes');
const getStaffPath = () => collection(db, 'docentes');
const getInventarioPath = () => collection(db, 'inventario');
const getPAEBenefPath = () => collection(db, 'pae_beneficiarios');
const getPermissionsPath = () => doc(db, 'configuracion_app', 'permisos_roles');

const getMealsPath = () => collection(db, 'registros_pae');
const getAttendancePath = () => collection(db, 'registros_asistencia');
const getLoansPath = () => collection(db, 'registros_prestamos');
const getNewsPath = () => collection(db, 'registros_novedades');
const getSalidasPath = () => collection(db, 'registros_salidas');
const getPruebasPath = () => collection(db, 'registros_pruebas');
const getEvaluacionesPath = () => collection(db, 'registros_evaluaciones');
const getPlanillasPath = () => collection(db, 'planillas_docentes');

// Variables Globales del Sistema
let currentStaff: any = null; 
let studentsDict: Record<string, StudentRecord> = {}; 
let staffDict: Record<string, StaffRecord> = {}; 
let inventarioDict: Record<string, InventoryItem> = {}; 
let paeBeneficiariosDict: Record<string, boolean> = {}; 
let rolePermissions: Record<string, string[]> = { 
    docente: ['asistencia', 'evaluacion', 'prestamo', 'pruebas', 'pae', 'novedad', 'clase_ef'], 
    vigilante: ['salida'], 
    coordinador: ['asistencia', 'evaluacion', 'salida', 'prestamo', 'pruebas', 'pae', 'novedad', 'clase_ef'], 
    administrador: ['asistencia', 'evaluacion', 'salida', 'prestamo', 'pruebas', 'pae', 'novedad', 'clase_ef'] 
};

const modulesList = [
    { id: 'asistencia', name: 'Asistencia' }, 
    { id: 'evaluacion', name: 'Evaluación (Notas)' }, 
    { id: 'salida', name: 'Salidas/Portería' }, 
    { id: 'prestamo', name: 'Préstamos' }, 
    { id: 'pruebas', name: 'Pruebas' }, 
    { id: 'pae', name: 'PAE' }, 
    { id: 'novedad', name: 'Novedades' }, 
    { id: 'clase_ef', name: 'Ed. Física' }
];

let html5QrcodeScanner: any = null; 
let isProcessing = false;
let institucionData = { nombre: "Sistema Institucional", logo: null as string | null, color1: "#2563eb", color2: "#0ea5e9" };
let appMode = 'asistencia'; 
let efPhase = 'asistencia'; 
let loanTempStudent: any = null; 
let tempNovedadStudent: any = null; 
let tempEvalStudent: any = null;
let pruebasCooldown: Record<string, number> = {}; 
let currentStudentForCarnet: any = null;

// Audio Beeps
let audioCtx: AudioContext | null = null;
function playBeep(type = 'success') {
    try {
        const AudioContextClass = window.AudioContext || (window as any).webkitAudioContext;
        if(!audioCtx) audioCtx = new AudioContextClass();
        if(audioCtx.state === 'suspended') audioCtx.resume();
        const oscillator = audioCtx.createOscillator();
        const gainNode = audioCtx.createGain();
        oscillator.connect(gainNode);
        gainNode.connect(audioCtx.destination);
        if(type === 'success') {
            oscillator.type = 'sine'; 
            oscillator.frequency.setValueAtTime(880, audioCtx.currentTime); 
            gainNode.gain.setValueAtTime(0.1, audioCtx.currentTime); 
            oscillator.start(); 
            oscillator.stop(audioCtx.currentTime + 0.15);
        } else {
            oscillator.type = 'sawtooth'; 
            oscillator.frequency.setValueAtTime(150, audioCtx.currentTime);
            oscillator.frequency.linearRampToValueAtTime(100, audioCtx.currentTime + 0.5);
            gainNode.gain.setValueAtTime(0.2, audioCtx.currentTime); 
            oscillator.start(); 
            oscillator.stop(audioCtx.currentTime + 0.6);
        }
    } catch(e) {}
}

// Master 2000
let masterWorkbook: any = null;
let masterStudentMap: Record<string, { sheetName: string; r: number }> = {}; 
let masterFileName = "Planilla_Calificada.xlsx";
let masterGradesCount = 0;

// Utilidades de Texto y Formato
export const normalizeNameMatch = (str: string) => {
    if(!str) return "";
    return str.normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[,.-]/g, ' ').replace(/\s+/g, ' ').toLowerCase().trim();
};

export const getFullName = (s: any) => {
    if (!s) return 'Sin Nombre';
    if (s.nombres || s.apellidos) return `${s.nombres || ''} ${s.apellidos || ''}`.trim();
    return s.nombre_completo || s.nombre || 'Sin Nombre';
};

export const getTodayString = () => { 
    const d = new Date(); 
    d.setMinutes(d.getMinutes() - d.getTimezoneOffset()); 
    return d.toISOString().split('T')[0]; 
};

export const formatTime = (d: Date) => d.toLocaleTimeString('es-CO', { hour: '2-digit', minute: '2-digit' });
export const formatTimeWithSeconds = (d: Date) => d.toLocaleTimeString('es-CO', { hour: '2-digit', minute: '2-digit', second: '2-digit' });

export const showToast = (msg: string, type = 'success', dur = 3000) => {
    const container = document.getElementById('toast-container'); 
    if (!container) return;
    const t = document.createElement('div');
    const bg = type === 'success' ? 'bg-green-600' : type === 'error' ? 'bg-red-600 border-l-8 border-red-800' : type === 'warning' ? 'bg-yellow-500' : 'bg-blue-600';
    const icon = type === 'success' ? 'fa-check-circle' : type === 'error' ? 'fa-times-circle' : type === 'warning' ? 'fa-exclamation-triangle' : 'fa-info-circle';
    
    t.className = dur > 5000 
        ? `${bg} text-white p-5 rounded-2xl shadow-2xl flex flex-col items-center justify-center animate-slide-3 text-center mb-2 z-50 border-4 border-white transform scale-105` 
        : `${bg} text-white p-4 rounded-xl shadow-2xl flex items-center animate-slide-3 font-medium text-left mb-2 z-50 text-sm`;
    
    t.innerHTML = dur > 5000 
        ? `<i class="fas ${icon} text-5xl mb-3"></i> <div class="w-full leading-tight">${msg}</div>` 
        : `<i class="fas ${icon} text-2xl mr-3"></i> <div class="flex-1 leading-tight">${msg}</div>`;
    
    container.appendChild(t); 
    setTimeout(() => { 
        t.style.opacity = '0'; 
        t.style.transform = 'translateY(-10px)'; 
        setTimeout(() => t.remove(), 300); 
    }, dur);
};

// ==========================================================
// RF-04: SEPARACIÓN AUTOMÁTICA DE NOMBRES Y APELLIDOS
// ==========================================================
/**
 * Separa automáticamente apellidos y nombres respetando partículas compuestas
 * Ej: "BERNAL GARCIA JUSTIN ANDRES" -> apellidos: "BERNAL GARCIA", nombres: "JUSTIN ANDRES"
 * Ej: "DE LA OSSA PEREZ JUAN CARLOS" -> apellidos: "DE LA OSSA PEREZ", nombres: "JUAN CARLOS"
 */
export function splitApellidosNombres(fullName: string): { apellidos: string; nombres: string; nombre_completo: string } {
    if (!fullName) return { apellidos: "", nombres: "", nombre_completo: "" };
    
    const cleaned = fullName.trim().replace(/\s+/g, ' ');
    const parts = cleaned.split(' ');
    
    if (parts.length <= 1) {
        return { apellidos: "", nombres: cleaned, nombre_completo: cleaned };
    }
    if (parts.length === 2) {
        return { apellidos: parts[0], nombres: parts[1], nombre_completo: cleaned };
    }
    if (parts.length === 3) {
        // e.g. "PEREZ GOMEZ JUAN" -> Apellidos: "PEREZ GOMEZ", Nombres: "JUAN"
        return { apellidos: `${parts[0]} ${parts[1]}`, nombres: parts[2], nombre_completo: cleaned };
    }

    // Partículas y prefijos hispánicos comunes en apellidos
    const prefixes = ['DE LA', 'DE LOS', 'DE LAS', 'DEL', 'DE', 'SAN', 'SANTA', 'LA', 'LOS', 'VON', 'VAN', 'DA', 'DO', 'DI'];
    let idx = 0;
    
    // Primer apellido
    let apellido1 = "";
    const twoWordPrefix1 = parts.slice(idx, idx + 2).join(' ').toUpperCase();
    const oneWordPrefix1 = parts[idx].toUpperCase();
    
    if (prefixes.includes(twoWordPrefix1) && parts.length > idx + 2) {
        apellido1 = parts.slice(idx, idx + 3).join(' ');
        idx += 3;
    } else if (prefixes.includes(oneWordPrefix1) && parts.length > idx + 1) {
        apellido1 = parts.slice(idx, idx + 2).join(' ');
        idx += 2;
    } else {
        apellido1 = parts[idx];
        idx += 1;
    }
    
    // Segundo apellido
    let apellido2 = "";
    if (idx < parts.length - 1) {
        const twoWordPrefix2 = parts.slice(idx, idx + 2).join(' ').toUpperCase();
        const oneWordPrefix2 = parts[idx].toUpperCase();
        
        if (prefixes.includes(twoWordPrefix2) && parts.length > idx + 3) {
            apellido2 = parts.slice(idx, idx + 3).join(' ');
            idx += 3;
        } else if (prefixes.includes(oneWordPrefix2) && parts.length > idx + 2) {
            apellido2 = parts.slice(idx, idx + 2).join(' ');
            idx += 2;
        } else {
            apellido2 = parts[idx];
            idx += 1;
        }
    }
    
    const apellidos = `${apellido1} ${apellido2}`.trim();
    const nombres = parts.slice(idx).join(' ').trim();
    
    return {
        apellidos: apellidos || parts.slice(0, 2).join(' '),
        nombres: nombres || parts.slice(2).join(' '),
        nombre_completo: cleaned
    };
}

// ==========================================================
// RF-02: DETECCIÓN DINÁMICA DE ENCABEZADOS DE COLUMNAS
// ==========================================================
interface HeaderMapping {
    headerRowIndex: number;
    colNumero: number;
    colMatricula: number;
    colTipoDoc: number;
    colDocumento: number;
    colEstudiante: number;
    colApellidos: number;
    colNombres: number;
    isOldFormat: boolean;
}

function normalizeHeaderString(str: any): string {
    if (!str) return "";
    return String(str)
        .toUpperCase()
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .replace(/[^A-Z0-9#]/g, " ")
        .replace(/\s+/g, " ")
        .trim();
}

export function detectExcelHeaders(rows: any[][]): HeaderMapping | null {
    // Buscar en las primeras 25 filas
    const maxSearchRows = Math.min(rows.length, 25);
    
    for (let r = 0; r < maxSearchRows; r++) {
        const row = rows[r];
        if (!row || !Array.isArray(row) || row.length === 0) continue;
        
        let colNumero = -1;
        let colMatricula = -1;
        let colTipoDoc = -1;
        let colDocumento = -1;
        let colEstudiante = -1;
        let colApellidos = -1;
        let colNombres = -1;
        
        for (let c = 0; c < row.length; c++) {
            const h = normalizeHeaderString(row[c]);
            if (!h) continue;

            // N° / Orden
            if (h === 'N' || h === 'NO' || h === 'NUM' || h === 'NUMERO' || h === '#' || h === 'ORDEN' || h === 'CONSECUTIVO') {
                colNumero = c;
            }
            // MATRÍCULA
            else if (h.includes('MATRICULA') || h === 'MAT' || h.includes('CODIGO') || h === 'ID MATRICULA') {
                colMatricula = c;
            }
            // TIPO DOC.
            else if (h.includes('TIPO DOC') || h.includes('TIPO DOCUMENTO') || h === 'TD' || h === 'T DOC' || h === 'TIPO DE DOCUMENTO') {
                colTipoDoc = c;
            }
            // DOCUMENTO (evitando confundir con Tipo Documento)
            else if ((h.includes('DOCUMENTO') || h.includes('IDENTIFICACION') || h.includes('DOC IDENT') || h === 'DOC' || h === 'IDENTIFICACION' || h === 'NRO DOC' || h === 'NUMERO DOC') && !h.includes('TIPO')) {
                colDocumento = c;
            }
            // ESTUDIANTE / NOMBRE COMPLETO / APELLIDOS Y NOMBRES
            else if (h.includes('ESTUDIANTE') || h.includes('ALUMNO') || h === 'NOMBRE COMPLETO' || h === 'APELLIDOS Y NOMBRES' || h === 'APELLIDOS NOMBRES' || h === 'NOMBRES Y APELLIDOS') {
                colEstudiante = c;
            }
            // Formato antiguo: APELLIDOS
            else if (h.includes('APELLIDO') && !h.includes('NOMBRE')) {
                colApellidos = c;
            }
            // Formato antiguo: NOMBRES
            else if (h.includes('NOMBRE') && !h.includes('APELLIDO') && !h.includes('COMPLETO')) {
                colNombres = c;
            }
        }
        
        const hasDoc = colDocumento !== -1;
        const hasMatricula = colMatricula !== -1;
        const hasName = colEstudiante !== -1 || (colApellidos !== -1 && colNombres !== -1);
        
        // Si detectamos al menos 2 columnas clave, esta es la fila de encabezados
        if ((hasDoc && (hasMatricula || hasName)) || (hasMatricula && hasName)) {
            return {
                headerRowIndex: r,
                colNumero,
                colMatricula,
                colTipoDoc,
                colDocumento,
                colEstudiante,
                colApellidos,
                colNombres,
                isOldFormat: colEstudiante === -1 && (colApellidos !== -1 || colNombres !== -1)
            };
        }
    }
    
    return null;
}

// ==========================================================
// NUEVO PROCESADOR EXCEL DINÁMICO (processExcelUpload)
// ==========================================================
export const processExcelUpload = async (file: File, type: 'students' | 'staff' | 'items' | 'pae') => {
    showToast('Analizando Excel...', 'warning', 2000);
    const countEl = document.getElementById(`count-${type}`);
    const progressEl = document.getElementById(`progress-${type}`);
    
    try {
        if(countEl) countEl.innerText = "Procesando...";
        if(progressEl) progressEl.style.width = "0%";

        const data = await file.arrayBuffer(); 
        const workbook = XLSX.read(data, { type: 'array' }); 
        
        if (type === 'students') {
            const parsedStudents: StudentRecord[] = [];
            const validationErrors: ExcelValidationError[] = [];
            
            // RF-01: Cada hoja representa un grupo académico. El nombre de la hoja tiene prioridad.
            workbook.SheetNames.forEach(rawSheetName => {
                const sheetName = rawSheetName.trim();
                const sheet = workbook.Sheets[rawSheetName];
                if (!sheet) return;
                
                const rows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
                if (!rows || rows.length === 0) return;
                
                // RF-02: Detección dinámica de encabezados
                const mapping = detectExcelHeaders(rows);
                
                let startRow = 1;
                let cMat = 0;
                let cTipo = -1;
                let cDoc = 1;
                let cEst = -1;
                let cApe = 3;
                let cNom = 2;
                let isOld = true;
                
                if (mapping) {
                    startRow = mapping.headerRowIndex + 1;
                    cMat = mapping.colMatricula !== -1 ? mapping.colMatricula : 0;
                    cTipo = mapping.colTipoDoc;
                    cDoc = mapping.colDocumento !== -1 ? mapping.colDocumento : 1;
                    cEst = mapping.colEstudiante;
                    cApe = mapping.colApellidos;
                    cNom = mapping.colNombres;
                    isOld = mapping.isOldFormat;
                } else {
                    // Fallback heurístico: buscar filas con datos
                    console.warn(`No se detectó fila de encabezados en hoja ${sheetName}, usando índices por defecto`);
                }
                
                // RF-03: Leer únicamente filas posteriores al encabezado
                for (let r = startRow; r < rows.length; r++) {
                    const row = rows[r];
                    if (!row || !Array.isArray(row) || row.length === 0) continue;
                    
                    // Verificar si la fila está completamente vacía
                    const rowJoined = row.map(cell => String(cell || '').trim()).join('');
                    if (!rowJoined) continue;
                    
                    // Extraer valores según mapeo
                    const rawMat = cMat !== -1 && row[cMat] !== undefined ? String(row[cMat]).trim() : "";
                    const rawTipo = cTipo !== -1 && row[cTipo] !== undefined ? String(row[cTipo]).trim().toUpperCase() : "";
                    const rawDoc = cDoc !== -1 && row[cDoc] !== undefined ? String(row[cDoc]).trim() : "";
                    
                    // Omitir filas de totales, firmas o pie de página
                    const rawFirstCell = String(row[0] || '').toUpperCase();
                    if (rawFirstCell.includes('TOTAL') || rawFirstCell.includes('FIRMA') || rawFirstCell.includes('RECTOR') || rawFirstCell.includes('OBSERVAC')) {
                        continue;
                    }
                    
                    let nombres = "";
                    let apellidos = "";
                    let nombreCompleto = "";
                    
                    if (!isOld && cEst !== -1) {
                        // RF-04: Nuevo Formato: Columna ESTUDIANTE (APELLIDOS NOMBRES)
                        const rawEst = String(row[cEst] || '').trim();
                        if (rawEst) {
                            const splitted = splitApellidosNombres(rawEst);
                            nombres = splitted.nombres;
                            apellidos = splitted.apellidos;
                            nombreCompleto = splitted.nombre_completo;
                        }
                    } else {
                        // Formato antiguo: Columnas APELLIDOS y NOMBRES separadas
                        const rawApe = cApe !== -1 && row[cApe] !== undefined ? String(row[cApe]).trim() : "";
                        const rawNom = cNom !== -1 && row[cNom] !== undefined ? String(row[cNom]).trim() : "";
                        
                        if (rawApe && rawNom) {
                            apellidos = rawApe;
                            nombres = rawNom;
                            nombreCompleto = `${apellidos} ${nombres}`.trim();
                        } else if (rawApe || rawNom) {
                            const single = rawApe || rawNom;
                            const splitted = splitApellidosNombres(single);
                            nombres = splitted.nombres;
                            apellidos = splitted.apellidos;
                            nombreCompleto = splitted.nombre_completo;
                        }
                    }
                    
                    // VALIDACIONES OBLIGATORIAS:
                    // - Documento obligatorio
                    // - Matrícula obligatoria
                    // - Nombre obligatorio
                    const rowNumberInExcel = r + 1;
                    const missingFields: string[] = [];
                    
                    if (!rawDoc) missingFields.push("DOCUMENTO");
                    if (!rawMat) missingFields.push("MATRÍCULA");
                    if (!nombres && !nombreCompleto) missingFields.push("NOMBRE / ESTUDIANTE");
                    
                    if (missingFields.length > 0) {
                        // Solo reportar error si la fila tenía algún dato relevante
                        if (rawDoc || rawMat || nombres || apellidos) {
                            missingFields.forEach(f => {
                                validationErrors.push({
                                    hoja: sheetName,
                                    fila: rowNumberInExcel,
                                    campo: f
                                });
                            });
                        }
                        continue; // No procesar fila incompleta
                    }
                    
                    // Tipo de documento (RF-05): Guardar tipo_documento y documento
                    const tipoDoc = rawTipo || (rawDoc.length === 10 ? "T.I." : (rawDoc.length === 8 ? "C.C." : "R.C."));
                    
                    // RF-06 y ESTRUCTURA DE DATOS FINAL:
                    const studentObj: StudentRecord = {
                        id: rawDoc,
                        matricula: rawMat,
                        tipo_documento: tipoDoc,
                        documento: rawDoc,
                        nombres: nombres,
                        apellidos: apellidos,
                        nombre_completo: nombreCompleto,
                        nombre: nombreCompleto, // compatibilidad hacia atrás
                        grado: sheetName // RF-01: Grado obtenido directamente del nombre de la hoja
                    };
                    
                    parsedStudents.push(studentObj);
                }
            });
            
            // Mostrar modal de errores si existen
            if (validationErrors.length > 0) {
                mostrarErroresExcel(validationErrors);
            }
            
            if (parsedStudents.length === 0) { 
                if(countEl) countEl.innerText = "0 válidos";
                return showToast('No se encontraron registros válidos de estudiantes.', 'error', 4000); 
            }
            
            showToast(`Sincronizando ${parsedStudents.length} alumnos a Firebase...`, 'warning', 4000);
            const colPath = getStudentsPath();
            
            // Obtener registros existentes para depuración y reemplazo exacto
            const existingSnap = await getDocs(colPath);
            const existingIds = new Set<string>();
            existingSnap.forEach(d => existingIds.add(d.id));

            const newIds = new Set(parsedStudents.map(item => String(item.id)));
            const idsToDelete = [...existingIds].filter(id => !newIds.has(id));

            if (idsToDelete.length > 0) {
                showToast(`Actualizando base: removiendo ${idsToDelete.length} registros no presentes...`, 'info', 3000);
                for (let i = 0; i < idsToDelete.length; i += 100) {
                    const chunk = idsToDelete.slice(i, i + 100);
                    await Promise.all(chunk.map(id => deleteDoc(doc(colPath, id))));
                }
            }

            const total = parsedStudents.length; 
            let current = 0; 
            const chunkSize = 100;
            
            for (let i = 0; i < total; i += chunkSize) {
                const chunk = parsedStudents.slice(i, i + chunkSize);
                await Promise.all(chunk.map(item => setDoc(doc(colPath, String(item.id)), item)));
                current += chunk.length;
                
                const pct = Math.round((current / total) * 100);
                if(progressEl) progressEl.style.width = `${pct}%`;
                if(countEl) countEl.innerText = `${current} / ${total}`;
            }
            
            if(countEl) countEl.innerText = `${total} estudiantes`;
            if(progressEl) setTimeout(() => progressEl.style.width = "0%", 2000); 
            
            // Respaldar también en cache local para alta disponibilidad
            try {
                localStorage.setItem('cached_students', JSON.stringify(parsedStudents));
            } catch(e) {}
            
            showToast(`¡Sincronización exitosa! ${total} estudiantes importados.`, 'success', 5000); 
            await loadDatabases();
            
        } else {
            // Manejo de Personal, Inventario y PAE
            let parsed: any[] = [];
            workbook.SheetNames.forEach(sheetName => {
                const rows: any[][] = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: "" }); 
                for (let i = 1; i < rows.length; i++) {
                    const row = rows[i]; 
                    if (!row || row.length === 0) continue;
                    
                    if (type === 'staff') { 
                        const idRaw = String(row[0]||'').trim(); 
                        if(!idRaw) continue;
                        const nomRaw = String(row[1]||'').trim(); 
                        const grupoRaw = String(row[2]||'').trim(); 
                        const rolRaw = String(row[3]||'').trim(); 
                        const claveRaw = row[4] ? String(row[4]).trim() : idRaw; 
                        if (nomRaw) parsed.push({ id: idRaw, nombre: nomRaw, grado: grupoRaw, rol: rolRaw, clave: claveRaw }); 
                    } 
                    else if (type === 'items') { 
                        const idRaw = String(row[0]||'').trim(); 
                        if(!idRaw) continue;
                        const nomRaw = String(row[1]||'').trim(); 
                        if (nomRaw) parsed.push({ id: idRaw, nombre: nomRaw }); 
                    } 
                    else if (type === 'pae') { 
                        const idRaw = String(row[0]||'').trim();
                        if(idRaw) parsed.push({ id: idRaw, beneficiario: true }); 
                    }
                }
            });
            
            if (parsed.length === 0) { 
                if(countEl) countEl.innerText = "Vacío";
                return showToast('Excel vacío o formato incorrecto.', 'error'); 
            }
            
            const colPath = type === 'staff' ? getStaffPath() : type === 'items' ? getInventarioPath() : getPAEBenefPath();
            const total = parsed.length; 
            let current = 0; 
            const chunkSize = 100;
            
            for (let i = 0; i < total; i += chunkSize) {
                const chunk = parsed.slice(i, i + chunkSize);
                await Promise.all(chunk.map(item => setDoc(doc(colPath, String(item.id)), item)));
                current += chunk.length;
                
                const pct = Math.round((current / total) * 100);
                if(progressEl) progressEl.style.width = `${pct}%`;
                if(countEl) countEl.innerText = `${current} / ${total}`;
            }
            
            if(countEl) countEl.innerText = `${total} registros`;
            if(progressEl) setTimeout(() => progressEl.style.width = "0%", 2000); 
            showToast(`¡Sincronización completada! ${total} registros.`, 'success', 4000); 
            await loadDatabases();
        }
        
    } catch (e) { 
        if(countEl) countEl.innerText = "Error";
        showToast('Error al procesar el archivo Excel.', 'error'); 
        console.error("Error en processExcelUpload:", e); 
    }
};

// ==========================================================
// MODAL DE INCONSISTENCIAS / ERRORES DE VALIDACIÓN EXCEL
// ==========================================================
function mostrarErroresExcel(errors: ExcelValidationError[]) {
    const modal = document.getElementById('modal-excel-errors');
    const list = document.getElementById('excel-errors-list');
    const summary = document.getElementById('excel-errors-summary');
    if (!modal || !list || !summary) return;
    
    summary.innerText = `Se encontraron ${errors.length} inconsistencias en el Excel. Las filas con datos faltantes fueron ignoradas para proteger la base de datos:`;
    list.innerHTML = '';
    
    // Agrupar por Hoja
    const grouped: Record<string, ExcelValidationError[]> = {};
    errors.forEach(err => {
        if (!grouped[err.hoja]) grouped[err.hoja] = [];
        grouped[err.hoja].push(err);
    });
    
    Object.keys(grouped).forEach(hoja => {
        const groupHeader = document.createElement('div');
        groupHeader.className = "font-bold text-red-800 bg-red-100 px-2 py-1 rounded mt-1";
        groupHeader.innerText = `📁 Hoja: ${hoja}`;
        list.appendChild(groupHeader);
        
        grouped[hoja].forEach(err => {
            const item = document.createElement('div');
            item.className = "py-1 px-2 text-gray-700 flex justify-between";
            item.innerHTML = `<span>Fila <b>${err.fila}</b></span> <span class="text-red-600 font-bold">Falta: ${err.campo}</span>`;
            list.appendChild(item);
        });
    });
    
    modal.classList.remove('hidden');
}

// Botones del Modal de Errores
document.getElementById('btn-close-excel-errors')?.addEventListener('click', () => {
    document.getElementById('modal-excel-errors')?.classList.add('hidden');
});
document.getElementById('btn-excel-errors-ok')?.addEventListener('click', () => {
    document.getElementById('modal-excel-errors')?.classList.add('hidden');
});

// ==========================================================
// DESCARGAR PLANTILLA EXCEL EJEMPLO NUEVO FORMATO
// ==========================================================
export function downloadSampleExcel() {
    const wb = XLSX.utils.book_new();
    
    // Grupo TS0501
    const dataTS0501 = [
        ["LISTADO DE ESTUDIANTES - AÑO 2026"],
        ["GRUPO: TS0501"],
        [],
        ["N°", "MATRÍCULA", "TIPO DOC.", "DOCUMENTO", "ESTUDIANTE"],
        [1, "261047", "R.C.", "1032033800", "BERNAL GARCIA JUSTIN ANDRES"],
        [2, "261048", "T.I.", "1032033801", "DE LA OSSA PEREZ JUAN CARLOS"],
        [3, "261049", "T.I.", "1032033802", "ZAPATA MARTINEZ VALENTINA"],
        [4, "261050", "T.I.", "1032033803", "SANCHEZ DEL VALLE LUIS MIGUEL"]
    ];
    const ws1 = XLSX.utils.aoa_to_sheet(dataTS0501);
    XLSX.utils.book_append_sheet(wb, ws1, "TS0501");
    
    // Grupo TS0502
    const dataTS0502 = [
        ["LISTADO DE ESTUDIANTES - AÑO 2026"],
        ["GRUPO: TS0502"],
        [],
        ["N°", "MATRÍCULA", "TIPO DOC.", "DOCUMENTO", "ESTUDIANTE"],
        [1, "262001", "T.I.", "1032033901", "GOMEZ ALVAREZ CAMILO"],
        [2, "262002", "T.I.", "1032033902", "RODRIGUEZ LOPEZ SOFIA MARIANA"],
        [3, "262003", "C.C.", "1032033903", "CASTRO HERRERA SEBASTIAN"]
    ];
    const ws2 = XLSX.utils.aoa_to_sheet(dataTS0502);
    XLSX.utils.book_append_sheet(wb, ws2, "TS0502");
    
    XLSX.writeFile(wb, "Plantilla_Estudiantes_Nuevo_Formato.xlsx");
    showToast("Plantilla Excel generada y descargada.", "success");
}

document.getElementById('btn-download-sample-excel')?.addEventListener('click', downloadSampleExcel);

// ==========================================================
// CARGA Y SINCRONIZACIÓN DE BASE DE DATOS
// ==========================================================
export const loadDatabases = async () => {
    try {
        const settingsSnap = await getDoc(getSettingsPath());
        if (settingsSnap.exists()) { 
            institucionData = settingsSnap.data() as any; 
            aplicarConfiguracionUI(); 
        }
        
        const studentsSnap = await getDocs(getStudentsPath()); 
        studentsDict = {};
        studentsSnap.forEach(docSnap => {
            studentsDict[docSnap.id] = docSnap.data() as StudentRecord;
        });
        const elStud = document.getElementById('count-students'); 
        if(elStud) elStud.innerText = Object.keys(studentsDict).length + ' registros';
        
        const staffSnap = await getDocs(getStaffPath()); 
        staffDict = {};
        staffSnap.forEach(docSnap => {
            staffDict[docSnap.id] = docSnap.data() as StaffRecord;
        });
        const elStaff = document.getElementById('count-staff'); 
        if(elStaff) elStaff.innerText = Object.keys(staffDict).length + ' registros';
        
        const inventarioSnap = await getDocs(getInventarioPath()); 
        inventarioDict = {};
        inventarioSnap.forEach(docSnap => {
            inventarioDict[docSnap.id] = docSnap.data() as InventoryItem;
        });
        const elItem = document.getElementById('count-items'); 
        if(elItem) elItem.innerText = Object.keys(inventarioDict).length + ' registros';

        const paeSnap = await getDocs(getPAEBenefPath()); 
        paeBeneficiariosDict = {};
        paeSnap.forEach(docSnap => {
            paeBeneficiariosDict[docSnap.id] = true;
        });
        const elPae = document.getElementById('count-pae'); 
        if(elPae) elPae.innerText = Object.keys(paeBeneficiariosDict).length + ' registros';

        const permSnap = await getDoc(getPermissionsPath());
        if (permSnap.exists()) { 
            rolePermissions = permSnap.data() as any; 
            if (!rolePermissions.administrador) {
                rolePermissions.administrador = ['asistencia', 'evaluacion', 'salida', 'prestamo', 'pruebas', 'pae', 'novedad', 'clase_ef'];
            }
        }
        
        renderRolesConfig(); 
        populateQRGroups();
        
        const badge = document.getElementById('firebase-status-badge');
        if (badge) {
            badge.className = "text-[10px] bg-green-800 text-green-200 px-2 py-0.5 rounded-full font-bold flex items-center gap-1";
            badge.innerHTML = '<span class="w-1.5 h-1.5 rounded-full bg-green-400"></span> BD Conectada';
        }
    } catch (error) { 
        console.error("Error cargando BD de Firebase:", error); 
        const badge = document.getElementById('firebase-status-badge');
        if (badge) {
            badge.className = "text-[10px] bg-yellow-800 text-yellow-200 px-2 py-0.5 rounded-full font-bold flex items-center gap-1";
            badge.innerHTML = '<span class="w-1.5 h-1.5 rounded-full bg-yellow-400"></span> Modo Local';
        }
        // Fallback a localStorage
        const cached = localStorage.getItem('cached_students');
        if (cached) {
            try {
                const list = JSON.parse(cached);
                studentsDict = {};
                list.forEach((s: StudentRecord) => { studentsDict[s.id] = s; });
                const elStud = document.getElementById('count-students'); 
                if(elStud) elStud.innerText = Object.keys(studentsDict).length + ' registros (Local)';
                populateQRGroups();
            } catch(e) {}
        }
    }
};

// Historial diario
export const loadTodayHistory = async (mode: string) => {
    if (!currentStaff) return;
    const container = document.getElementById('recent-scans'); 
    const counterEl = document.getElementById('counter-today');
    if (!container || !counterEl) return;
    
    container.innerHTML = '<div class="text-center py-4"><i class="fas fa-spinner fa-spin text-gray-300"></i></div>'; 
    counterEl.innerText = `0 registros`;
    
    try {
        let pathFn: any, color: string;
        if(mode==='pae') { pathFn=getMealsPath; color='blue'; } 
        else if(mode==='asistencia') { pathFn=getAttendancePath; color='green'; } 
        else if(mode==='evaluacion') { pathFn=getEvaluacionesPath; color='blue'; } 
        else if(mode==='pruebas') { pathFn=getPruebasPath; color='purple'; } 
        else if(mode==='prestamo') { pathFn=getLoansPath; color='yellow'; } 
        else if(mode==='novedad') { pathFn=getNewsPath; color='red'; } 
        else if(mode==='salida') { pathFn=getSalidasPath; color='gray'; } 
        else { pathFn=getAttendancePath; color='indigo'; }
        
        const snap = await getDocs(pathFn()); 
        const todayRecs: any[] = [];
        snap.forEach(d => { 
            const data: any = d.data(); 
            if (data.date === getTodayString() && data.recordedById === currentStaff.id) todayRecs.push(data); 
        });
        
        todayRecs.sort((a, b) => new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime()); 
        counterEl.innerText = `${todayRecs.length} registros`;
        
        if (todayRecs.length > 0) {
            container.innerHTML = '';
            todayRecs.slice(0, 15).forEach(r => {
                const l1 = r.nombre || r.studentName; 
                const l2 = mode==='pae'?`PAE`:mode==='asistencia'?'Asistencia':mode==='evaluacion'?`Nota: ${r.nota}`:mode==='pruebas'?'Prueba':mode==='prestamo'?`Llevó: ${r.itemName}`:mode==='salida'?(r.status==='fuera'?'Salió':'Volvió'):mode==='clase_ef'?'Ed. Física':'Novedad';
                const item = document.createElement('div'); 
                item.className = `flex justify-between items-center p-3 bg-${color}-50 rounded-lg border border-${color}-100 shadow-sm`;
                item.innerHTML = `<div class="flex flex-col"><span class="font-bold text-gray-800 text-sm mb-1">${l1}</span><span class="text-11px font-bold text-${color}-600 uppercase">${l2}</span></div><span class="text-xs font-bold text-gray-500 bg-white px-2 py-1 rounded shadow-sm">${(mode === 'pruebas') ? formatTimeWithSeconds(new Date(r.timestamp)) : formatTime(new Date(r.timestamp))}</span>`;
                container.appendChild(item);
            });
        } else {
            container.innerHTML = `<div class="text-gray-400 text-center py-4">Sin registros.</div>`;
        }
    } catch (e) { 
        container.innerHTML = `<div class="text-gray-400 text-center py-4">Sin registros (Base limpia).</div>`; 
    }
};

export const addRecentScanToUI = (line1: string, line2: string, dateObj: Date, colorBase = 'blue') => {
    const container = document.getElementById('recent-scans');
    if (!container) return;
    if (container.children.length === 1 && container.children[0].classList.contains('text-center')) container.innerHTML = '';
    
    const item = document.createElement('div'); 
    item.className = `flex justify-between items-center p-3 bg-${colorBase}-50 rounded-lg border border-${colorBase}-100 shadow-sm animate-slide-3`;
    const timeTxt = (appMode === 'pruebas') ? formatTimeWithSeconds(dateObj) : formatTime(dateObj);
    
    item.innerHTML = `<div class="flex flex-col"><span class="font-bold text-gray-800 text-sm mb-1">${line1}</span><span class="text-11px font-bold text-${colorBase}-600 uppercase">${line2}</span></div><span class="text-xs font-bold text-gray-500 bg-white px-2 py-1 rounded shadow-sm">${timeTxt}</span>`;
    container.insertBefore(item, container.firstChild); 
    if (container.children.length > 15 && container.lastChild) container.removeChild(container.lastChild);
    
    const counterEl = document.getElementById('counter-today'); 
    if (counterEl) {
        counterEl.innerText = `${(parseInt(counterEl.innerText) || 0) + 1} registros`;
    }
};

// ==========================================================
// MASTER 2000 - RELACIÓN MATRÍCULA Y NOTAS
// ==========================================================
const loadMasterMappings = () => {
    masterStudentMap = {}; 
    let matchedCount = 0;
    if (!masterWorkbook) return 0;
    
    masterWorkbook.SheetNames.forEach((sheetName: string) => {
        const sheet = masterWorkbook.Sheets[sheetName];
        const json: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
        for (let r = 2; r < json.length; r++) { 
            if (!json[r]) continue;
            const rawMatricula = String(json[r][1] || "").trim(); // Columna B: Matrícula
            if (rawMatricula) {
                for(const id in studentsDict) {
                    if (String(studentsDict[id].matricula).trim() === rawMatricula) {
                        masterStudentMap[id] = { sheetName: sheetName, r: r };
                        matchedCount++; 
                        break;
                    }
                }
            }
        }
    });
    return matchedCount;
};

// Carga de Excel Master 2000
document.getElementById('excel-master')?.addEventListener('change', async (e: any) => {
    const file = e.target.files[0];
    if (!file) return;
    masterFileName = file.name;
    masterGradesCount = 0;
    
    showToast("Verificando planilla...", "info", 2000);
    try {
        const data = await file.arrayBuffer();
        masterWorkbook = XLSX.read(data, { type: 'array' });
        const matchedCount = loadMasterMappings();

        if (matchedCount > 0) {
            showToast(`¡Aprobada! Subiendo planilla a Firebase...`, 'warning', 3000);
            
            const reader = new FileReader();
            reader.onload = async () => {
                try {
                    const base64Str = reader.result as string;
                    await setDoc(doc(getPlanillasPath(), currentStaff.id), {
                        fileName: masterFileName,
                        base64: base64Str,
                        timestamp: new Date().toISOString()
                    });
                    
                    const fnDisplay = document.getElementById('master-file-name-display');
                    if (fnDisplay) fnDisplay.innerText = masterFileName;
                    const gcDisplay = document.getElementById('master-grades-counter');
                    if (gcDisplay) gcDisplay.innerText = `Notas en esta sesión: 0`;
                    
                    document.getElementById('master-upload-wrapper')?.classList.add('hidden');
                    document.getElementById('master-actions-wrapper')?.classList.remove('hidden');
                    document.getElementById('btn-reset-master')?.classList.remove('hidden');
                    
                    showToast(`¡Planilla respaldada! (${matchedCount} alumnos listos)`, 'success', 5000);
                } catch(err) {
                    showToast('Error subiendo a Firebase.', 'error');
                }
            };
            reader.readAsDataURL(file);
        } else { 
            showToast('No se encontraron matrículas coincidentes en la columna B del archivo.', 'error', 4000);
            masterWorkbook = null;
        }
    } catch(e) {
        showToast("Error procesando Excel Master", "error"); 
        console.error(e);
    }
    e.target.value = '';
});

// Descargar Master
document.getElementById('btn-download-master')?.addEventListener('click', () => {
    if(!masterWorkbook) return;
    const btn = document.getElementById('btn-download-master'); 
    if(!btn) return;
    const orig = btn.innerHTML;
    btn.innerHTML = '<i class="fas fa-spinner fa-spin mr-2 text-lg"></i> Procesando...';
    try {
        XLSX.writeFile(masterWorkbook, masterFileName);
        showToast('¡Planilla descargada con éxito!', 'success');
    } catch(e) {
        showToast('Error al descargar', 'error');
    } finally { 
        btn.innerHTML = orig; 
    }
});

// Limpiar Master
document.getElementById('btn-reset-master')?.addEventListener('click', async () => {
    const btn = document.getElementById('btn-reset-master');
    if(!btn) return;
    btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i>';
    try {
        await deleteDoc(doc(getPlanillasPath(), currentStaff.id));
        masterWorkbook = null; 
        masterStudentMap = {}; 
        masterFileName = "Planilla_Calificada.xlsx"; 
        masterGradesCount = 0;
        document.getElementById('master-upload-wrapper')?.classList.remove('hidden');
        document.getElementById('master-actions-wrapper')?.classList.add('hidden');
        btn.classList.add('hidden');
        showToast('Planilla removida de la Nube.', 'info');
    } catch(err) {
        showToast('Error al borrar', 'error');
    } finally {
        btn.innerHTML = '<i class="fas fa-trash-alt"></i> Limpiar';
    }
});

// ==========================================================
// PROCESAMIENTO DE ESCANEO Y BÚSQUEDA
// Compatibilidad total: Búsqueda por Documento o Matrícula
// ==========================================================
export const processScan = async (scannedText: string) => {
    if(isProcessing || !currentStaff) return;
    isProcessing = true; 
    scannedText = scannedText.trim();
    if (!scannedText) { isProcessing = false; return; }

    // Compatibilidad: buscar por ID, Documento o Matrícula
    let resolvedId = scannedText;
    if (!studentsDict[scannedText]) {
        const found = Object.values(studentsDict).find(s => 
            String(s.matricula).trim() === scannedText || 
            String(s.documento).trim() === scannedText
        );
        if (found) resolvedId = found.id;
    }

    try {
        if (appMode === 'pae') await handlePAE(resolvedId);
        else if (appMode === 'asistencia') await handleAttendance(resolvedId);
        else if (appMode === 'evaluacion') await handleEvaluacion(resolvedId);
        else if (appMode === 'pruebas') await handlePruebas(resolvedId);
        else if (appMode === 'salida') await handleSalida(resolvedId);
        else if (appMode === 'prestamo') await handleLoan(resolvedId);
        else if (appMode === 'novedad') await handleNovedad(resolvedId);
        else if (appMode === 'clase_ef') await handleClaseEF(resolvedId);
    } catch(e) { 
        showToast("Error de conexión. Intente nuevamente.", "error"); 
        console.error(e);
    }
    
    if (appMode !== 'novedad' && appMode !== 'evaluacion' && (appMode !== 'clase_ef' || (efPhase !== 'novedad' && efPhase !== 'evaluacion'))) {
        setTimeout(() => { isProcessing = false; }, 300);
    } else { 
        isProcessing = false; 
    }
};

// Handlers por Módulo
const handlePAE = async (studentId: string) => {
    const student = studentsDict[studentId]; 
    if (!student) { playBeep('error'); return showToast(`Estudiante no registrado.`, 'error'); }
    
    if (Object.keys(paeBeneficiariosDict).length > 0 && !paeBeneficiariosDict[student.id]) {
        playBeep('error'); 
        return showToast(`<div class="mb-2"><span class="bg-white text-red-700 text-sm font-black px-3 py-1 rounded-full uppercase shadow">Atención</span></div><strong class="text-2xl block mt-2 text-white">NO ES BENEFICIARIO</strong><span class="text-lg font-bold mt-2 block">${getFullName(student)}</span>`, 'error', 5000);
    }
    
    const mealRef = doc(getMealsPath(), `${getTodayString()}_${student.id}`); 
    const mealSnap = await getDoc(mealRef);
    
    if (mealSnap.exists()) {
        const m = mealSnap.data(); 
        playBeep('error');
        showToast(`<div class="mb-2"><span class="bg-white text-red-700 text-sm font-black px-3 py-1 rounded-full uppercase shadow">Ración Duplicada</span></div><strong class="text-2xl block mt-2 text-white">¡YA COMIÓ HOY!</strong><span class="text-lg font-bold mt-2 block border-b-2 border-red-400 pb-2">${getFullName(student)}</span><div class="bg-red-900 bg-opacity-30 p-3 rounded-lg mt-3 text-sm border border-red-500">Hora: <b class="text-white text-lg">${formatTime(new Date(m.timestamp))}</b><br>Autorizó: <b class="text-white">${m.recordedByName}</b></div>`, 'error', 6000); 
    } else {
        await setDoc(mealRef, { 
            studentId: student.id, 
            nombre: getFullName(student), 
            grado: student.grado || 'Sin Grupo', 
            date: getTodayString(), 
            timestamp: new Date().toISOString(), 
            recordedById: currentStaff.id, 
            recordedByName: currentStaff.nombre 
        });
        playBeep('success'); 
        showToast(`<b>APROBADO PAE</b><br>${getFullName(student)}`, 'success', 1000); 
        addRecentScanToUI(getFullName(student), `PAE - ${student.grado}`, new Date(), 'blue');
    }
};

const handleAttendance = async (studentId: string) => {
    const student = studentsDict[studentId]; 
    if (!student) { playBeep('error'); return showToast(`Estudiante no encontrado.`, 'error'); }
    
    const attRef = doc(getAttendancePath(), `${getTodayString()}_${student.id}`); 
    const attSnap = await getDoc(attRef);
    if (attSnap.exists()) { 
        playBeep('error'); 
        showToast(`<b>${getFullName(student)}</b><br>Ya ingresó hoy.`, 'warning'); 
    } else { 
        await setDoc(attRef, { 
            studentId: student.id, 
            nombre: getFullName(student), 
            grado: student.grado || 'Sin Grupo', 
            date: getTodayString(), 
            timestamp: new Date().toISOString(), 
            recordedById: currentStaff.id, 
            recordedByName: currentStaff.nombre 
        });
        playBeep('success'); 
        showToast(`<b>ASISTENCIA</b><br>${getFullName(student)}`, 'success', 1000); 
        addRecentScanToUI(getFullName(student), 'Asistencia', new Date(), 'green'); 
    }
};

const handleEvaluacion = async (studentId: string) => {
    const student = studentsDict[studentId]; 
    if (!student) { playBeep('error'); return showToast(`Estudiante no encontrado.`, 'error'); }
    
    if (masterWorkbook) {
        const mapData = masterStudentMap[student.id];
        if (!mapData) {
            playBeep('error');
            return showToast(`<strong class="text-lg">ALUMNO NO VINCULADO</strong><br><b>${getFullName(student)}</b><br>Verifique la matrícula (${student.matricula}) en la planilla Master.`, 'error', 4000);
        }
        
        const select = document.getElementById('modal-eval-activity-select') as HTMLSelectElement;
        if (select) {
            select.innerHTML = '<option value="">Selecciona Actividad (1.1 a 1.17)</option>';
            select.classList.remove('hidden');
            
            const sheet = masterWorkbook.Sheets[mapData.sheetName];
            const json: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
            const actNames = ["1.1", "1.2", "1.3", "1.4", "1.5", "1.6", "1.7", "1.8", "1.9", "1.10", "1.11", "1.12", "1.13", "1.14", "1.15", "1.16", "1.17"];
            
            for(let col = 4; col <= 20; col++) {
                let hasGrades = false;
                for (let row = 2; row < json.length; row++) {
                    if (json[row] && String(json[row][col]).trim() !== "") {
                        hasGrades = true; break;
                    }
                }
                const symbol = hasGrades ? "🔴" : "🟢";
                const label = actNames[col - 4];
                select.innerHTML += `<option value="${col}">${symbol} Actividad ${label}</option>`;
            }
        }
    } else {
        document.getElementById('modal-eval-activity-select')?.classList.add('hidden');
    }

    tempEvalStudent = student; 
    playBeep('success');
    const nameEl = document.getElementById('eval-student-name');
    if (nameEl) nameEl.innerText = `${getFullName(student)} (${student.grado})`;
    
    const slider = document.getElementById('eval-slider') as HTMLInputElement;
    if (slider) slider.value = "50";
    updateEvalUI(50); 
    document.getElementById('modal-evaluacion')?.classList.remove('hidden');
};

const updateEvalUI = (val: number | string) => {
    const num = (Number(val) / 10).toFixed(1); 
    const valDisp = document.getElementById('eval-value-display');
    if (valDisp) valDisp.innerText = num;
    const circle = document.getElementById('eval-display-circle');
    if (!circle) return;
    
    if (Number(num) < 3.0) {
        circle.className = "w-28 h-28 rounded-full border-4 flex items-center justify-center text-5xl font-black transition-colors duration-200 border-red-500 text-red-600 bg-red-50";
    } else if (Number(num) < 4.0) {
        circle.className = "w-28 h-28 rounded-full border-4 flex items-center justify-center text-5xl font-black transition-colors duration-200 border-yellow-500 text-yellow-600 bg-yellow-50";
    } else {
        circle.className = "w-28 h-28 rounded-full border-4 flex items-center justify-center text-5xl font-black transition-colors duration-200 border-green-500 text-green-600 bg-green-50";
    }
};

document.getElementById('eval-slider')?.addEventListener('input', (e: any) => updateEvalUI(e.target.value));
document.getElementById('btn-eval-minus')?.addEventListener('click', () => { 
    const s = document.getElementById('eval-slider') as HTMLInputElement; 
    if(s && Number(s.value) > 0) { 
        s.value = String(parseInt(s.value) - 1); 
        updateEvalUI(s.value); 
    } 
});
document.getElementById('btn-eval-plus')?.addEventListener('click', () => { 
    const s = document.getElementById('eval-slider') as HTMLInputElement; 
    if(s && Number(s.value) < 50) { 
        s.value = String(parseInt(s.value) + 1); 
        updateEvalUI(s.value); 
    } 
});
document.getElementById('btn-cancel-eval')?.addEventListener('click', () => { 
    document.getElementById('modal-evaluacion')?.classList.add('hidden'); 
    tempEvalStudent = null; 
    isProcessing = false; 
});

document.getElementById('btn-save-eval')?.addEventListener('click', async () => {
    if(!tempEvalStudent) return;
    const slider = document.getElementById('eval-slider') as HTMLInputElement;
    const finalNota = (Number(slider.value) / 10).toFixed(1);
    
    let targetColIdx: any = -1;
    if (masterWorkbook) {
        const sel = document.getElementById('modal-eval-activity-select') as HTMLSelectElement;
        targetColIdx = sel.value;
        if (!targetColIdx) {
            return showToast("Debes seleccionar una actividad (1.1 a 1.17)", "warning");
        }
    }

    const btn = document.getElementById('btn-save-eval') as HTMLButtonElement; 
    btn.disabled = true; 
    btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i>';
    
    if (masterWorkbook && targetColIdx !== -1) {
        const mapData = masterStudentMap[tempEvalStudent.id]; 
        if (mapData) {
            const sheet = masterWorkbook.Sheets[mapData.sheetName];
            const cellRef = XLSX.utils.encode_cell({ r: mapData.r, c: parseInt(targetColIdx) });
            if (!sheet[cellRef]) { 
                sheet[cellRef] = { t: 'n', v: parseFloat(finalNota) }; 
            } else { 
                sheet[cellRef].t = 'n'; 
                sheet[cellRef].v = parseFloat(finalNota); 
            }
            
            masterGradesCount++;
            const counter = document.getElementById('master-grades-counter');
            if (counter) counter.innerText = `Notas en esta sesión: ${masterGradesCount}`;

            try {
                const wbout = XLSX.write(masterWorkbook, { bookType: 'xlsx', type: 'base64' });
                const base64Str = "data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64," + wbout;
                await updateDoc(doc(getPlanillasPath(), currentStaff.id), {
                    base64: base64Str,
                    lastUpdate: new Date().toISOString()
                });
            } catch (e) {
                console.error("Error actualizando planilla en Firebase", e);
            }
        }
    }

    try {
        await setDoc(doc(getEvaluacionesPath(), `${Date.now()}_${tempEvalStudent.id}`), { 
            studentId: tempEvalStudent.id, 
            studentName: getFullName(tempEvalStudent), 
            studentGrado: tempEvalStudent.grado, 
            nota: finalNota, 
            date: getTodayString(), 
            timestamp: new Date().toISOString(), 
            recordedById: currentStaff.id, 
            recordedByName: currentStaff.nombre 
        });
        showToast(`Nota ${finalNota} asignada.`, 'success'); 
        addRecentScanToUI(getFullName(tempEvalStudent), `Nota: ${finalNota}`, new Date(), 'blue'); 
        document.getElementById('modal-evaluacion')?.classList.add('hidden');
    } catch (e) { 
        showToast("Error al guardar evaluación", "error"); 
    } finally { 
        btn.disabled = false; 
        btn.innerHTML = '<i class="fas fa-save mr-2"></i>Guardar Nota'; 
        tempEvalStudent = null; 
        isProcessing = false; 
    }
});

const handlePruebas = async (studentId: string) => {
    const student = studentsDict[studentId]; 
    if (!student) { playBeep('error'); return showToast(`Estudiante no encontrado.`, 'error'); }
    
    const now = Date.now(); 
    if (pruebasCooldown[student.id] && (now - pruebasCooldown[student.id] < 3000)) { 
        playBeep('error'); 
        return showToast(`Espera 3 segundos.`, 'warning', 2000); 
    }
    pruebasCooldown[student.id] = now;
    
    await setDoc(doc(getPruebasPath(), `${now}_${student.id}`), { 
        studentId: student.id, 
        nombre: getFullName(student), 
        grado: student.grado || 'Sin Grupo', 
        date: getTodayString(), 
        timestamp: new Date().toISOString(), 
        recordedById: currentStaff.id, 
        recordedByName: currentStaff.nombre 
    });
    playBeep('success'); 
    showToast(`<b>⏱️ TIEMPO REGISTRADO</b><br>${getFullName(student)}`, 'success', 1000); 
    addRecentScanToUI(getFullName(student), 'Prueba Física', new Date(), 'purple');
};

const handleSalida = async (studentId: string) => {
    const student = studentsDict[studentId]; 
    if (!student) { playBeep('error'); return showToast(`Estudiante no encontrado.`, 'error'); }
    
    const ref = doc(getSalidasPath(), `${getTodayString()}_${student.id}`); 
    const snap = await getDoc(ref); 
    const data = snap.exists() ? snap.data() : null;
    const role = (currentStaff.rol || '').toLowerCase().trim();

    if (role.includes('vigilan') || role.includes('porter')) {
        if (data && data.status === 'autorizado') { 
            await updateDoc(ref, { status: 'salio', timeOut: new Date().toISOString() }); 
            playBeep('success'); 
            showToast(`<strong class="text-3xl block mt-2 text-white">✅ PUEDE SALIR</strong><span class="text-lg font-bold mt-2 block">${getFullName(student)}</span>`, 'success', 4000); 
            addRecentScanToUI(getFullName(student), 'Salió', new Date(), 'gray'); 
        } else { 
            playBeep('error'); 
            showToast(`<strong class="text-3xl block mt-2 text-white">❌ ACCESO DENEGADO</strong><span class="text-lg font-bold mt-2 block">${getFullName(student)}</span><span class="text-sm">No tiene autorización en el sistema.</span>`, 'error', 5000); 
        }
    } else if (role.includes('admin') || role.includes('coord') || role.includes('rector')) {
        await setDoc(ref, { 
            studentId: student.id, 
            nombre: getFullName(student), 
            grado: student.grado || 'Sin Grupo', 
            status: 'autorizado', 
            date: getTodayString(), 
            timestamp: new Date().toISOString(), 
            recordedById: currentStaff.id, 
            recordedByName: currentStaff.nombre 
        }); 
        playBeep('success'); 
        showToast(`<b>AUTORIZADO:</b><br>${getFullName(student)}`, 'success', 2500); 
        addRecentScanToUI(getFullName(student), 'Permiso Salida', new Date(), 'gray');
    } else {
        if (data && data.status === 'fuera') { 
            await updateDoc(ref, { status: 'regreso', timeIn: new Date().toISOString() }); 
            playBeep('success'); 
            showToast(`<b>REGRESÓ:</b> ${getFullName(student)}`, 'success', 1500); 
            addRecentScanToUI(getFullName(student), 'Volvió', new Date(), 'gray'); 
        } else { 
            await setDoc(ref, { 
                studentId: student.id, 
                nombre: getFullName(student), 
                grado: student.grado || 'Sin Grupo', 
                status: 'fuera', 
                date: getTodayString(), 
                timestamp: new Date().toISOString(), 
                timeOut: new Date().toISOString(), 
                recordedById: currentStaff.id, 
                recordedByName: currentStaff.nombre 
            }); 
            playBeep('success'); 
            showToast(`<b>SALIÓ:</b> ${getFullName(student)}`, 'warning', 1500); 
            addRecentScanToUI(getFullName(student), 'Salió', new Date(), 'gray'); 
        }
    }
};

const handleLoan = async (scannedText: string) => {
    const sb = document.getElementById('loan-status-bar'); 
    const st = document.getElementById('loan-status-text');
    if (!sb || !st) return;

    if (!loanTempStudent) {
        let student = studentsDict[scannedText];
        if (!student) {
            student = Object.values(studentsDict).find(s => 
                String(s.matricula).trim() === scannedText || 
                String(s.documento).trim() === scannedText
            )!;
        }
        if (student) { 
            loanTempStudent = student; 
            playBeep('success'); 
            sb.classList.remove('hidden'); 
            st.innerHTML = `<i class="fas fa-user text-yellow-600"></i> <b class="text-lg">${getFullName(student)}</b><br><span class="text-sm font-bold bg-white px-2 py-1 rounded mt-1 inline-block text-gray-700">Ahora escanea el elemento...</span>`; 
        } else { 
            playBeep('error'); 
            showToast(`Escanea carnet de estudiante primero.`, 'error'); 
        }
    } else {
        const itemName = inventarioDict[scannedText] ? inventarioDict[scannedText].nombre : scannedText;
        await setDoc(doc(getLoansPath(), `${Date.now()}_${loanTempStudent.id}`), { 
            studentId: loanTempStudent.id, 
            studentName: getFullName(loanTempStudent), 
            studentGrado: loanTempStudent.grado, 
            itemId: scannedText, 
            itemName: itemName, 
            status: 'prestado', 
            date: getTodayString(), 
            timestamp: new Date().toISOString(), 
            recordedById: currentStaff.id, 
            recordedByName: currentStaff.nombre 
        });
        playBeep('success'); 
        showToast(`<b>PRESTADO: ${itemName}</b>`, 'success'); 
        addRecentScanToUI(getFullName(loanTempStudent), `Llevó: ${itemName}`, new Date(), 'yellow');
        st.innerHTML = `<i class="fas fa-user text-yellow-600"></i> <b class="text-lg">${getFullName(loanTempStudent)}</b><br><span class="text-sm font-bold bg-green-100 px-2 py-1 rounded mt-1 inline-block text-green-800">Llevó: ${itemName}. Puedes escanear otro...</span>`;
    }
};

document.getElementById('btn-finish-loan')?.addEventListener('click', () => { 
    loanTempStudent = null; 
    document.getElementById('loan-status-bar')?.classList.add('hidden'); 
    showToast('Préstamos finalizados.', 'success'); 
});

// Novedades
const handleNovedad = async (studentId: string) => {
    const student = studentsDict[studentId]; 
    if (!student) { playBeep('error'); return showToast(`Estudiante no encontrado.`, 'error'); }
    tempNovedadStudent = student; 
    playBeep('success'); 
    const nName = document.getElementById('novedad-student-name');
    if (nName) nName.innerText = `${getFullName(student)} (${student.grado})`;
    const nText = document.getElementById('novedad-text') as HTMLTextAreaElement;
    if (nText) nText.value = ''; 
    document.getElementById('modal-novedad')?.classList.remove('hidden'); 
    setTimeout(() => nText?.focus(), 100);
};

document.getElementById('btn-cancel-novedad')?.addEventListener('click', () => { 
    document.getElementById('modal-novedad')?.classList.add('hidden'); 
    tempNovedadStudent = null; 
    isProcessing = false; 
});

document.getElementById('btn-save-novedad')?.addEventListener('click', async () => {
    const nText = document.getElementById('novedad-text') as HTMLTextAreaElement;
    const text = nText?.value.trim(); 
    if(!text) return showToast("Escribe la novedad", "warning");
    const btn = document.getElementById('btn-save-novedad') as HTMLButtonElement; 
    btn.disabled = true; 
    btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i>';
    
    try {
        await setDoc(doc(getNewsPath(), `${Date.now()}_${tempNovedadStudent.id}`), { 
            studentId: tempNovedadStudent.id, 
            studentName: getFullName(tempNovedadStudent), 
            studentGrado: tempNovedadStudent.grado, 
            observacion: text, 
            date: getTodayString(), 
            timestamp: new Date().toISOString(), 
            recordedById: currentStaff.id, 
            recordedByName: currentStaff.nombre 
        });
        showToast('Novedad guardada.', 'success'); 
        addRecentScanToUI(getFullName(tempNovedadStudent), `Novedad`, new Date(), 'red'); 
        document.getElementById('modal-novedad')?.classList.add('hidden');
    } catch (e) { 
        showToast("Error", "error"); 
    } finally { 
        btn.disabled = false; 
        btn.innerHTML = '<i class="fas fa-save mr-2"></i>Guardar'; 
        tempNovedadStudent = null; 
        isProcessing = false; 
    }
});

// Educación Física
const handleClaseEF = async (scannedText: string) => {
    if (efPhase === 'asistencia') await handleAttendance(scannedText);
    else if (efPhase === 'evaluacion') await handleEvaluacion(scannedText);
    else if (efPhase === 'prestamo') {
        let student = studentsDict[scannedText];
        if (!student) {
            student = Object.values(studentsDict).find(s => 
                String(s.matricula).trim() === scannedText || 
                String(s.documento).trim() === scannedText
            )!;
        }
        if (student && !loanTempStudent) {
            loanTempStudent = student; 
            playBeep('success'); 
            showToast(`<b>Estudiante: ${getFullName(student)}</b><br>Ahora escanea el elemento deportivo.`, 'success', 2500);
        } else {
            const studentObj = loanTempStudent || { id: 'EF_GENERAL', nombres: 'Clase Ed. Física', grado: 'General' };
            const itemName = inventarioDict[scannedText] ? inventarioDict[scannedText].nombre : scannedText;
            await setDoc(doc(getLoansPath(), `${Date.now()}_${studentObj.id}`), { 
                studentId: studentObj.id, 
                studentName: getFullName(studentObj), 
                studentGrado: studentObj.grado || 'EF', 
                itemId: scannedText, 
                itemName: itemName, 
                status: 'prestado', 
                date: getTodayString(), 
                timestamp: new Date().toISOString(), 
                recordedById: currentStaff.id, 
                recordedByName: currentStaff.nombre, 
                clase: 'Ed. Física' 
            });
            playBeep('success'); 
            showToast(`<b>Elemento Prestado: ${itemName}</b>`, 'success'); 
            addRecentScanToUI(getFullName(studentObj), `EF - Llevó: ${itemName}`, new Date(), 'indigo');
            loanTempStudent = null; 
        }
    } else if (efPhase === 'salida') {
        let student = studentsDict[scannedText];
        if (!student) {
            student = Object.values(studentsDict).find(s => 
                String(s.matricula).trim() === scannedText || 
                String(s.documento).trim() === scannedText
            )!;
        }
        if (!student) { playBeep('error'); return showToast('Estudiante no encontrado', 'error'); }
        await setDoc(doc(getSalidasPath(), `${getTodayString()}_${student.id}`), { 
            studentId: student.id, 
            nombre: getFullName(student), 
            grado: student.grado || 'Sin Grupo', 
            status: 'salio_ef', 
            date: getTodayString(), 
            timestamp: new Date().toISOString(), 
            recordedById: currentStaff.id, 
            recordedByName: currentStaff.nombre, 
            clase: 'Ed. Física' 
        });
        playBeep('success'); 
        showToast(`<b>SALIDA EF:</b> ${getFullName(student)}`, 'success', 1500); 
        addRecentScanToUI(getFullName(student), 'EF - Salió', new Date(), 'indigo');
    } else if (efPhase === 'novedad') {
        await handleNovedad(scannedText);
    } else if (efPhase === 'regreso') {
        let student = studentsDict[scannedText];
        if (!student) {
            student = Object.values(studentsDict).find(s => 
                String(s.matricula).trim() === scannedText || 
                String(s.documento).trim() === scannedText
            )!;
        }
        if (!student) { playBeep('error'); return showToast('Estudiante no encontrado', 'error'); }
        await setDoc(doc(getSalidasPath(), `${getTodayString()}_${student.id}`), { 
            studentId: student.id, 
            nombre: getFullName(student), 
            grado: student.grado || 'Sin Grupo', 
            status: 'regreso_ef', 
            date: getTodayString(), 
            timestamp: new Date().toISOString(), 
            recordedById: currentStaff.id, 
            recordedByName: currentStaff.nombre, 
            clase: 'Ed. Física' 
        });
        playBeep('success'); 
        showToast(`<b>REGRESO EF:</b> ${getFullName(student)}`, 'success', 1500); 
        addRecentScanToUI(getFullName(student), 'EF - Regresó', new Date(), 'indigo');
    } else if (efPhase === 'devolucion') {
        const itemName = inventarioDict[scannedText] ? inventarioDict[scannedText].nombre : scannedText;
        await setDoc(doc(getLoansPath(), `${Date.now()}_dev_${scannedText}`), { 
            studentId: 'DEV_EF', 
            studentName: 'Clase Ed. Física', 
            studentGrado: 'General', 
            itemId: scannedText, 
            itemName: itemName, 
            status: 'devuelto', 
            date: getTodayString(), 
            timestamp: new Date().toISOString(), 
            recordedById: currentStaff.id, 
            recordedByName: currentStaff.nombre, 
            clase: 'Ed. Física' 
        });
        playBeep('success'); 
        showToast(`<b>Elemento Devuelto: ${itemName}</b>`, 'success', 1500); 
        addRecentScanToUI('Sala de Deportes', `EF - Devolvió: ${itemName}`, new Date(), 'indigo');
    }
};

// ==========================================================
// CÁMARA Y ESCÁNER QR
// ==========================================================
const startScanner = () => {
    if (!document.getElementById('view-scanner')?.classList.contains('hidden') && currentStaff) {
        const overlay = document.getElementById('camera-overlay');
        const btnStop = document.getElementById('btn-stop-camera');
        
        if (overlay) overlay.classList.add('hidden');
        if (btnStop) btnStop.classList.remove('hidden');

        if (html5QrcodeScanner) return;
        
        if (window.Html5Qrcode) {
            html5QrcodeScanner = new window.Html5Qrcode("reader");
            html5QrcodeScanner.start(
                { facingMode: "environment" }, 
                { fps: 15 }, 
                (t: string) => processScan(t), 
                () => {}
            ).catch(() => {
                if (overlay) overlay.classList.remove('hidden');
                if (btnStop) btnStop.classList.add('hidden');
                const r = document.getElementById('reader');
                if (r) r.innerHTML = `<div class="p-6 text-center text-gray-400"><i class="fas fa-camera-slash text-4xl mb-3"></i><p>Sin acceso a cámara.<br>Utiliza la búsqueda manual por nombre o documento.</p></div>`;
            });
        }
    }
};

const stopScanner = () => { 
    const overlay = document.getElementById('camera-overlay');
    const btnStop = document.getElementById('btn-stop-camera');
    if (overlay) overlay.classList.remove('hidden');
    if (btnStop) btnStop.classList.add('hidden');

    if (html5QrcodeScanner) {
        try {
            html5QrcodeScanner.stop().then(() => {
                html5QrcodeScanner.clear();
                html5QrcodeScanner = null;
            }).catch(() => { 
                html5QrcodeScanner.clear();
                html5QrcodeScanner = null; 
            });
        } catch(e) { 
            html5QrcodeScanner = null; 
        }
    }
};

document.addEventListener('click', (e: any) => {
    if (e.target.closest('#btn-activate-camera')) {
        startScanner();
    }
    if (e.target.closest('#btn-stop-camera')) {
        stopScanner();
    }
});

// ==========================================================
// BÚSQUEDA MANUAL INTELIGENTE (DOCUMENTO, MATRÍCULA O NOMBRE)
// ==========================================================
const inputManual = document.getElementById('manual-id') as HTMLInputElement; 
const dropdownSearch = document.getElementById('search-dropdown');

if (inputManual && dropdownSearch) {
    inputManual.addEventListener('input', (e: any) => {
        const rawQuery = e.target.value.trim(); 
        const query = normalizeNameMatch(rawQuery); 
        dropdownSearch.innerHTML = '';
        if (query.length < 2) { 
            dropdownSearch.classList.add('hidden'); 
            return; 
        }
        
        const results: any[] = []; 
        const addedIds = new Set<string>();
        
        // Búsqueda en estudiantes por Nombre, Documento o Matrícula
        Object.values(studentsDict).forEach(st => {
            const fullName = getFullName(st); 
            const normName = normalizeNameMatch(fullName);
            const docStr = String(st.documento || st.id || '');
            const matStr = String(st.matricula || '');
            
            if (docStr.includes(query) || matStr.includes(query) || normName.includes(query)) {
                results.push({ 
                    id: st.id, 
                    title: fullName, 
                    subtitle: `Doc: ${docStr} • Mat: ${matStr || '-'} • Grado: ${st.grado}`, 
                    icon: 'fa-user' 
                }); 
                addedIds.add(st.id);
            }
        });
        
        // Búsqueda en personal
        Object.values(staffDict).forEach(st => {
            if(addedIds.has(st.id)) return;
            const fullName = getFullName(st); 
            const normName = normalizeNameMatch(fullName);
            if (st.id.includes(query) || normName.includes(query)) {
                results.push({ 
                    id: st.id, 
                    title: fullName, 
                    subtitle: 'Docente/Operador', 
                    icon: 'fa-chalkboard-teacher' 
                });
            }
        });
        
        // Búsqueda en inventario
        Object.values(inventarioDict).forEach(item => {
            const itemName = item.nombre || ''; 
            const normItem = normalizeNameMatch(itemName);
            if (item.id.includes(query) || normItem.includes(query)) {
                results.push({ 
                    id: item.id, 
                    title: itemName, 
                    subtitle: 'Elemento de Inventario', 
                    icon: 'fa-basketball-ball' 
                });
            }
        });

        const topResults = results.slice(0, 8);
        if (topResults.length > 0) {
            topResults.forEach(res => {
                const itemDiv = document.createElement('div'); 
                itemDiv.className = "p-3 hover:bg-blue-50 cursor-pointer flex items-center gap-3 transition-colors";
                itemDiv.innerHTML = `<i class="fas ${res.icon} text-gray-400 text-lg w-6 text-center"></i><div class="flex flex-col"><span class="font-bold text-gray-800 text-sm leading-tight">${res.title}</span><span class="text-xs text-gray-500">${res.subtitle}</span></div>`;
                itemDiv.addEventListener('click', () => { 
                    inputManual.value = ''; 
                    dropdownSearch.classList.add('hidden'); 
                    processScan(res.id); 
                });
                dropdownSearch.appendChild(itemDiv);
            });
            dropdownSearch.classList.remove('hidden');
        } else { 
            dropdownSearch.innerHTML = `<div class="p-3 text-center text-sm text-gray-500">No se encontraron resultados</div>`; 
            dropdownSearch.classList.remove('hidden'); 
        }
    });

    document.addEventListener('click', (e: any) => { 
        if (!inputManual.contains(e.target) && !dropdownSearch.contains(e.target)) {
            dropdownSearch.classList.add('hidden'); 
        }
    });
    
    document.getElementById('btn-manual-submit')?.addEventListener('click', () => { 
        const val = inputManual.value.trim(); 
        if(val) { 
            processScan(val); 
            inputManual.value = ''; 
            dropdownSearch.classList.add('hidden'); 
        } 
    });
    
    inputManual.addEventListener('keypress', (e) => { 
        if (e.key === 'Enter') document.getElementById('btn-manual-submit')?.click(); 
    });
}

// ==========================================================
// CARNET DIGITAL
// Nombres y Apellidos en líneas separadas
// ==========================================================
export const renderCarnet = (user: any) => {
    document.getElementById('sp-step-carnet')?.classList.remove('hidden'); 
    const fotoEl = document.getElementById('carnet-render-foto') as HTMLImageElement;
    if (fotoEl) fotoEl.src = user.foto || "";
    
    const logoEl = document.getElementById('carnet-render-logo') as HTMLImageElement;
    if (logoEl) {
        if(institucionData.logo) logoEl.src = institucionData.logo; 
        else logoEl.src = "data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='%239ca3af'%3E%3Cpath d='M12 3L1 9L12 15L21 10.09V17H23V9L12 3ZM12 12.8L3.93 9L12 5.2L20.07 9L12 12.8ZM5 13.09V15.9L12 19.5L19 15.9V13.09L12 16.5L5 13.09Z'/%3E%3C/svg%3E";
    }

    const nomEl = document.getElementById('carnet-render-nombres');
    const apeEl = document.getElementById('carnet-render-apellidos');
    
    // CARNET DIGITAL: Nombres y Apellidos estrictamente separados
    if (user.nombres && user.apellidos) { 
        if (nomEl) nomEl.innerText = user.nombres; 
        if (apeEl) {
            apeEl.innerText = user.apellidos; 
            apeEl.classList.remove('hidden'); 
        }
    } else { 
        const splitted = splitApellidosNombres(user.nombre || user.nombre_completo || 'USUARIO');
        if (nomEl) nomEl.innerText = splitted.nombres || splitted.nombre_completo; 
        if (apeEl) {
            if (splitted.apellidos) {
                apeEl.innerText = splitted.apellidos;
                apeEl.classList.remove('hidden');
            } else {
                apeEl.classList.add('hidden');
            }
        }
    }
    
    const docEl = document.getElementById('carnet-render-doc');
    const tipoDoc = user.tipo_documento ? `${user.tipo_documento} ` : '';
    if (docEl) docEl.innerText = `${tipoDoc}${user.documento || user.id}`;
    
    const matEl = document.getElementById('carnet-render-matricula');
    if (matEl) {
        if (user.matricula) {
            matEl.innerText = `MATRÍCULA: ${user.matricula}`;
            matEl.classList.remove('hidden');
        } else {
            matEl.classList.add('hidden');
        }
    }
    
    const gradoEl = document.getElementById('carnet-render-grado');
    if (gradoEl) {
        gradoEl.innerText = user.isStaff ? (user.rol || 'DOCENTE').toUpperCase() : `${user.grado || 'S/G'}`;
        if (institucionData.color1) gradoEl.style.color = institucionData.color1;
    }
    
    const qrBox = document.getElementById('carnet-render-qr-box');
    if (qrBox) {
        qrBox.innerHTML = '';
        if (window.QRCode) {
            new window.QRCode(qrBox, { text: String(user.id), width: 144, height: 144 });
        }
    }
};

// Portal del estudiante
document.getElementById('btn-open-student-portal')?.addEventListener('click', () => { 
    document.getElementById('staff-login-overlay')?.classList.add('hidden'); 
    document.getElementById('student-portal-overlay')?.classList.remove('hidden'); 
    const inputSp = document.getElementById('sp-student-id') as HTMLInputElement;
    if (inputSp) inputSp.value = ''; 
    document.getElementById('sp-step-search')?.classList.remove('hidden'); 
    document.getElementById('sp-step-photo')?.classList.add('hidden'); 
    document.getElementById('sp-step-carnet')?.classList.add('hidden'); 
});

document.getElementById('btn-back-to-login')?.addEventListener('click', () => { 
    document.getElementById('student-portal-overlay')?.classList.add('hidden'); 
    document.getElementById('staff-login-overlay')?.classList.remove('hidden'); 
});

// Búsqueda en portal de estudiante por Documento o Matrícula
document.getElementById('btn-sp-search')?.addEventListener('click', () => {
    const inputSp = document.getElementById('sp-student-id') as HTMLInputElement;
    const query = inputSp?.value.trim(); 
    if(!query) return showToast('Ingresa tu documento o matrícula', 'warning');
    
    let userObj: any = studentsDict[query]; 
    let isStaff = false; 
    
    if(!userObj) { 
        // Buscar por matrícula o documento alterno
        userObj = Object.values(studentsDict).find(s => 
            String(s.matricula).trim() === query || 
            String(s.documento).trim() === query
        );
    }
    
    if(!userObj) { 
        userObj = staffDict[query]; 
        if(userObj) isStaff = true; 
    }
    
    if(!userObj) return showToast('Estudiante o usuario no encontrado.', 'error');
    
    currentStudentForCarnet = { ...userObj, isStaff }; 
    document.getElementById('sp-step-search')?.classList.add('hidden');
    
    if(userObj.foto) {
        renderCarnet(userObj); 
    } else { 
        const welcomeName = document.getElementById('sp-welcome-name');
        if (welcomeName) welcomeName.innerText = `¡Hola ${userObj.nombres || userObj.nombre || ''}!`; 
        document.getElementById('sp-step-photo')?.classList.remove('hidden'); 
    }
});

// Foto del Carnet
document.getElementById('sp-photo-input')?.addEventListener('change', (e: any) => {
    const file = e.target.files[0]; 
    if (file) {
        const reader = new FileReader(); 
        reader.onload = (event: any) => { 
            const img = new Image(); 
            img.onload = () => {
                const canvas = document.createElement('canvas'); 
                const MAX_WIDTH = 400; 
                const scaleSize = MAX_WIDTH / img.width; 
                canvas.width = MAX_WIDTH; 
                canvas.height = img.height * scaleSize;
                const ctx = canvas.getContext('2d'); 
                ctx?.drawImage(img, 0, 0, canvas.width, canvas.height); 
                const base64 = canvas.toDataURL('image/jpeg', 0.8);
                
                const preview = document.getElementById('sp-photo-preview') as HTMLImageElement;
                if (preview) {
                    preview.src = base64; 
                    preview.classList.remove('hidden'); 
                }
                document.getElementById('btn-sp-save-photo')?.classList.remove('hidden'); 
                currentStudentForCarnet.fotoTemp = base64; 
            }; 
            img.src = event.target.result; 
        }; 
        reader.readAsDataURL(file);
    }
});

document.getElementById('btn-sp-save-photo')?.addEventListener('click', async () => {
    if(!currentStudentForCarnet || !currentStudentForCarnet.fotoTemp) return;
    const btn = document.getElementById('btn-sp-save-photo') as HTMLButtonElement; 
    btn.innerHTML = '<i class="fas fa-spinner fa-spin mr-2"></i>Guardando...'; 
    btn.disabled = true;
    
    try { 
        const collectionPath = currentStudentForCarnet.isStaff ? getStaffPath() : getStudentsPath(); 
        await updateDoc(doc(collectionPath, currentStudentForCarnet.id), { foto: currentStudentForCarnet.fotoTemp }); 
        currentStudentForCarnet.foto = currentStudentForCarnet.fotoTemp; 
        
        if(currentStudentForCarnet.isStaff) staffDict[currentStudentForCarnet.id].foto = currentStudentForCarnet.foto; 
        else studentsDict[currentStudentForCarnet.id].foto = currentStudentForCarnet.foto; 
        
        document.getElementById('sp-step-photo')?.classList.add('hidden'); 
        renderCarnet(currentStudentForCarnet); 
    } catch (e) { 
        showToast('Error al guardar la foto', 'error'); 
    } finally { 
        btn.innerHTML = '<i class="fas fa-check-circle mr-2"></i>Guardar y Ver Carnet'; 
        btn.disabled = false; 
    }
});

document.getElementById('btn-change-photo')?.addEventListener('click', () => { 
    document.getElementById('sp-step-carnet')?.classList.add('hidden'); 
    const preview = document.getElementById('sp-photo-preview') as HTMLImageElement;
    if (preview) { preview.classList.add('hidden'); preview.src = ''; }
    const photoInput = document.getElementById('sp-photo-input') as HTMLInputElement;
    if (photoInput) photoInput.value = ''; 
    document.getElementById('btn-sp-save-photo')?.classList.add('hidden'); 
    const welcome = document.getElementById('sp-welcome-name');
    if (welcome) welcome.innerText = `Actualizar Foto`; 
    document.getElementById('sp-step-photo')?.classList.remove('hidden'); 
});

// Descargar Carnet Digital como Imagen
document.getElementById('btn-download-carnet')?.addEventListener('click', () => {
    const btn = document.getElementById('btn-download-carnet') as HTMLButtonElement; 
    const originalHtml = btn.innerHTML; 
    btn.innerHTML = '<i class="fas fa-spinner fa-spin mr-2"></i> Procesando...'; 
    btn.disabled = true;
    
    const carnetNode = document.getElementById('carnet-card');
    if (window.html2canvas && carnetNode) {
        window.html2canvas(carnetNode, { scale: 3, useCORS: true, backgroundColor: null }).then((canvas: HTMLCanvasElement) => { 
            const imgData = canvas.toDataURL('image/jpeg', 0.95); 
            const a = document.createElement('a'); 
            a.href = imgData; 
            a.download = `Carnet_${currentStudentForCarnet.id}.jpg`; 
            document.body.appendChild(a); 
            a.click(); 
            document.body.removeChild(a); 
            btn.innerHTML = originalHtml; 
            btn.disabled = false; 
            showToast('Carnet descargado', 'success'); 
        }).catch(() => { 
            showToast('Error al descargar carnet', 'error'); 
            btn.innerHTML = originalHtml; 
            btn.disabled = false; 
        });
    } else {
        btn.innerHTML = originalHtml; 
        btn.disabled = false;
    }
});

// ==========================================================
// CONFIGURACIÓN DE IDENTIDAD Y UI
// ==========================================================
const aplicarConfiguracionUI = () => {
    const hName = document.getElementById('header-inst-name');
    const iInput = document.getElementById('inst-name-input') as HTMLInputElement;
    if (institucionData.nombre) { 
        if (hName) hName.innerText = institucionData.nombre; 
        if (iInput) iInput.value = institucionData.nombre; 
    }
    if (institucionData.logo) { 
        const hLogo = document.getElementById('header-logo') as HTMLImageElement;
        if (hLogo) hLogo.src = institucionData.logo; 
        document.getElementById('header-logo-container')?.classList.remove('hidden'); 
        document.getElementById('header-default-icon')?.classList.add('hidden'); 
        const lPrev = document.getElementById('logo-preview') as HTMLImageElement;
        if (lPrev) { lPrev.src = institucionData.logo; lPrev.classList.remove('hidden'); }
    }
    if (institucionData.color1) { 
        const c1 = document.getElementById('inst-color1') as HTMLInputElement;
        if (c1) c1.value = institucionData.color1; 
        const mHead = document.getElementById('main-header');
        if (mHead) mHead.style.backgroundColor = institucionData.color1; 
        document.getElementById('wave-3')?.setAttribute('fill', institucionData.color1); 
    }
    if (institucionData.color2) { 
        const c2 = document.getElementById('inst-color2') as HTMLInputElement;
        if (c2) c2.value = institucionData.color2; 
        document.getElementById('wave-1')?.setAttribute('fill', institucionData.color2); 
        document.getElementById('wave-2')?.setAttribute('fill', institucionData.color2); 
    }
};

document.getElementById('inst-logo-input')?.addEventListener('change', (e: any) => {
    const file = e.target.files[0]; 
    if (file) { 
        const reader = new FileReader(); 
        reader.onload = (event: any) => { 
            const img = new Image(); 
            img.onload = () => { 
                const canvas = document.createElement('canvas'); 
                const MAX_HEIGHT = 150; 
                const scaleSize = img.height > MAX_HEIGHT ? MAX_HEIGHT / img.height : 1; 
                canvas.width = img.width * scaleSize; 
                canvas.height = img.height * scaleSize; 
                const ctx = canvas.getContext('2d'); 
                ctx?.drawImage(img, 0, 0, canvas.width, canvas.height); 
                institucionData.logo = canvas.toDataURL('image/png'); 
                const lPrev = document.getElementById('logo-preview') as HTMLImageElement;
                if (lPrev) { lPrev.src = institucionData.logo; lPrev.classList.remove('hidden'); }
            }; 
            img.src = event.target.result; 
        }; 
        reader.readAsDataURL(file); 
    }
});

document.getElementById('btn-save-settings')?.addEventListener('click', async () => {
    const btn = document.getElementById('btn-save-settings') as HTMLButtonElement; 
    btn.innerHTML = '<i class="fas fa-spinner fa-spin mr-2"></i>Guardando...'; 
    btn.disabled = true;
    const iInput = document.getElementById('inst-name-input') as HTMLInputElement;
    const c1 = document.getElementById('inst-color1') as HTMLInputElement;
    const c2 = document.getElementById('inst-color2') as HTMLInputElement;
    institucionData.nombre = iInput?.value.trim() || 'Gestor Académico RM'; 
    institucionData.color1 = c1?.value || '#2563eb'; 
    institucionData.color2 = c2?.value || '#0ea5e9';
    
    try { 
        await setDoc(getSettingsPath(), institucionData); 
        aplicarConfiguracionUI(); 
        showToast('Identidad guardada.', 'success'); 
    } catch (error) { 
        localStorage.setItem('institucionData', JSON.stringify(institucionData)); 
        aplicarConfiguracionUI(); 
        showToast('Guardado localmente.', 'warning'); 
    } finally { 
        btn.innerHTML = '<i class="fas fa-save mr-2"></i>Guardar Identidad'; 
        btn.disabled = false; 
    }
});

// Event Listeners de Archivos Excel
document.getElementById('excel-students')?.addEventListener('change', (e: any) => { 
    if(e.target.files[0]) processExcelUpload(e.target.files[0], 'students'); 
    e.target.value=''; 
});
document.getElementById('excel-staff')?.addEventListener('change', (e: any) => { 
    if(e.target.files[0]) processExcelUpload(e.target.files[0], 'staff'); 
    e.target.value=''; 
});
document.getElementById('excel-items')?.addEventListener('change', (e: any) => { 
    if(e.target.files[0]) processExcelUpload(e.target.files[0], 'items'); 
    e.target.value=''; 
});
document.getElementById('excel-pae')?.addEventListener('change', (e: any) => { 
    if(e.target.files[0]) processExcelUpload(e.target.files[0], 'pae'); 
    e.target.value=''; 
});
document.getElementById('excel-pae-staff')?.addEventListener('change', (e: any) => { 
    if(e.target.files[0]) processExcelUpload(e.target.files[0], 'pae'); 
    e.target.value=''; 
});

// Roles y Permisos
const renderRolesConfig = () => {
    const container = document.getElementById('roles-config-container'); 
    if (!container) return;
    container.innerHTML = '';
    const roleLabels: Record<string, string> = { 'administrador': 'Admin', 'coordinador': 'Coordinadores', 'docente': 'Docentes', 'vigilante': 'Seguridad' };
    
    Object.keys(rolePermissions).forEach(role => {
        let html = `<div class="p-3 bg-gray-50 rounded-lg border border-gray-200"><h4 class="font-bold text-sm text-gray-700 mb-2">${roleLabels[role] || role}</h4><div class="flex flex-wrap gap-3">`;
        modulesList.forEach(mod => { 
            const checked = rolePermissions[role]?.includes(mod.id) ? 'checked' : ''; 
            html += `<label class="flex items-center text-xs font-medium text-gray-600 cursor-pointer"><input type="checkbox" class="mr-1 role-cb rounded" data-role="${role}" data-module="${mod.id}" ${checked}> ${mod.name}</label>`; 
        });
        container.innerHTML += html + `</div></div>`;
    });
};

document.getElementById('btn-save-roles')?.addEventListener('click', async () => {
    const newPerms: Record<string, string[]> = { docente: [], vigilante: [], coordinador: [], administrador: [] }; 
    document.querySelectorAll('.role-cb').forEach((cb: any) => { 
        if (cb.checked) newPerms[cb.dataset.role].push(cb.dataset.module); 
    });
    await setDoc(getPermissionsPath(), newPerms); 
    rolePermissions = newPerms; 
    showToast('Permisos guardados', 'success'); 
    if (currentStaff) applyRolesUI(currentStaff);
});

// Formatear Sistema
document.getElementById('btn-factory-reset')?.addEventListener('click', async () => {
    if (confirm("¿Estás seguro de que deseas formatear todos los datos del sistema? Esta acción no se puede deshacer.")) {
        const pin = prompt("Escribe 'BORRAR' para confirmar:");
        if (pin === "BORRAR") {
            try {
                showToast("Formateando base de datos...", "warning", 3000);
                const studentsSnap = await getDocs(getStudentsPath());
                await Promise.all(studentsSnap.docs.map(d => deleteDoc(d.ref)));
                studentsDict = {};
                showToast("Base de datos de estudiantes limpia.", "success");
                await loadDatabases();
            } catch(e) {
                showToast("Error al formatear", "error");
            }
        }
    }
});

// Navegación entre vistas
const switchTab = (tab: string) => {
    document.getElementById('main-overlays-container')?.classList.add('hidden');
    stopScanner();
    
    if (tab === 'scanner') {
        document.getElementById('view-scanner')?.classList.remove('hidden'); 
        document.getElementById('view-scanner')?.classList.add('flex');
        document.getElementById('view-admin')?.classList.add('hidden'); 
        document.getElementById('view-admin')?.classList.remove('flex');
        document.getElementById('btn-admin-gear')?.classList.remove('hidden'); 
    } else if (tab === 'admin') {
        document.getElementById('view-admin')?.classList.remove('hidden'); 
        document.getElementById('view-admin')?.classList.add('flex');
        document.getElementById('view-scanner')?.classList.add('hidden'); 
        document.getElementById('view-scanner')?.classList.remove('flex');
        document.getElementById('btn-admin-gear')?.classList.add('hidden'); 
    } else {
        document.getElementById('main-overlays-container')?.classList.remove('hidden');
        document.getElementById('view-scanner')?.classList.add('hidden'); 
        document.getElementById('view-admin')?.classList.add('hidden'); 
    }
};

document.getElementById('btn-admin-gear')?.addEventListener('click', () => switchTab('admin'));
document.getElementById('btn-back-from-login')?.addEventListener('click', () => switchTab('scanner'));
document.getElementById('btn-admin-logout')?.addEventListener('click', () => { 
    document.getElementById('admin-panel')?.classList.add('hidden'); 
    document.getElementById('admin-panel')?.classList.remove('flex'); 
    document.getElementById('admin-login')?.classList.remove('hidden'); 
    const pin = document.getElementById('admin-pin') as HTMLInputElement;
    if (pin) pin.value = ''; 
    switchTab('login'); 
});

document.getElementById('btn-admin-login')?.addEventListener('click', () => {
    const pin = document.getElementById('admin-pin') as HTMLInputElement;
    if (pin?.value === 'profe123') { 
        document.getElementById('admin-login')?.classList.add('hidden');
        document.getElementById('admin-panel')?.classList.remove('hidden'); 
        document.getElementById('admin-panel')?.classList.add('flex');
    } else {
        showToast('Clave de administrador incorrecta', 'error');
    }
});

const applyRolesUI = (staff: any) => {
    const role = (staff.rol || '').toLowerCase().trim();
    const btns = document.querySelectorAll('.mode-btn');
    
    let permRole = 'docente';
    if (role.includes('vigilan') || role.includes('porter')) permRole = 'vigilante';
    else if (role === 'administrador' || role.includes('admin')) permRole = 'administrador';
    else if (role.includes('coord') || role.includes('rector')) permRole = 'coordinador';

    const allowedModules = rolePermissions[permRole] || [];
    if(!allowedModules.includes('evaluacion') && (permRole==='docente' || permRole==='coordinador' || permRole==='administrador')) {
        allowedModules.push('evaluacion'); 
    }

    btns.forEach((b: any) => {
        if(allowedModules.includes(b.dataset.mode)) b.classList.remove('hidden');
        else b.classList.add('hidden');
    });

    const iconSal = document.getElementById('icon-salida');
    const textSal = document.getElementById('text-salida');
    if (permRole === 'vigilante') { 
        if (iconSal) iconSal.className = 'fas fa-shield-alt mb-1 text-lg'; 
        if (textSal) textSal.innerText = 'Portería'; 
    } else if (permRole === 'coordinador' || permRole === 'administrador') { 
        if (iconSal) iconSal.className = 'fas fa-sign-out-alt mb-1 text-lg'; 
        if (textSal) textSal.innerText = 'Autorizar Salida'; 
    } else { 
        if (iconSal) iconSal.className = 'fas fa-door-open mb-1 text-lg'; 
        if (textSal) textSal.innerText = 'Salidas'; 
    }
    
    if (permRole === 'coordinador' || permRole === 'administrador') { 
        document.getElementById('btn-upload-pae-staff')?.classList.remove('hidden'); 
    } else { 
        document.getElementById('btn-upload-pae-staff')?.classList.add('hidden'); 
    }

    if (allowedModules.length > 0 && !allowedModules.includes(appMode)) appMode = allowedModules[0];
    else if (allowedModules.length === 0) appMode = '';
    
    updateModeUI();
};

const updateModeUI = () => {
    document.querySelectorAll('.mode-btn').forEach((btn: any) => {
        if (btn.dataset.mode === appMode) { 
            btn.classList.add('module-active'); 
            btn.classList.remove('module-inactive'); 
        } else { 
            btn.classList.remove('module-active'); 
            btn.classList.add('module-inactive'); 
        }
    });
    
    const efBar = document.getElementById('ef-subtoolbar');
    if (appMode === 'clase_ef') { 
        efBar?.classList.remove('hidden'); 
        efBar?.classList.add('flex'); 
    } else { 
        efBar?.classList.add('hidden'); 
        efBar?.classList.remove('flex'); 
    }

    const masterBar = document.getElementById('eval-master-subtoolbar');
    if (appMode === 'evaluacion' || (appMode === 'clase_ef' && efPhase === 'evaluacion')) {
        masterBar?.classList.remove('hidden'); 
        masterBar?.classList.add('flex');
    } else {
        masterBar?.classList.add('hidden'); 
        masterBar?.classList.remove('flex');
    }

    loanTempStudent = null; 
    document.getElementById('loan-status-bar')?.classList.add('hidden');
    document.getElementById('btn-view-active-loans')?.classList.toggle('hidden', appMode !== 'prestamo');
    loadTodayHistory(appMode);
};

document.querySelectorAll('.mode-btn').forEach(btn => btn.addEventListener('click', (e: any) => { 
    appMode = e.currentTarget.dataset.mode; 
    updateModeUI(); 
}));

const efPhaseNames: Record<string, string> = { 
    'asistencia': '1. Asistencia en Salón', 
    'evaluacion': '2. Evaluación de Notas', 
    'prestamo': '3. Préstamo de Elementos', 
    'salida': '4. Salida por Portería', 
    'novedad': '5. Registro de Novedades', 
    'regreso': '6. Regreso a la Institución', 
    'devolucion': '7. Entrega de Elementos' 
};

document.querySelectorAll('.ef-phase-btn').forEach(btn => {
    btn.addEventListener('click', (e: any) => {
        const phase = e.currentTarget.dataset.efPhase; 
        efPhase = phase;
        const pLabel = document.getElementById('ef-phase-label');
        if (pLabel) pLabel.innerText = `Fase: ${efPhaseNames[phase]}`;
        
        document.querySelectorAll('.ef-phase-btn').forEach((b: any) => {
            if (b.dataset.efPhase === phase) { 
                b.className = "ef-phase-btn bg-indigo-600 text-white py-2 px-1 rounded text-xs font-bold shadow-sm flex flex-col items-center"; 
            } else { 
                b.className = "ef-phase-btn bg-white text-indigo-700 border border-indigo-300 py-2 px-1 rounded text-xs font-bold shadow-sm flex flex-col items-center"; 
            }
        });
        showToast(`Modo Ed. Física: ${efPhaseNames[phase]}`, 'info', 1500);
        updateModeUI();
    });
});

// Login de personal
document.getElementById('btn-staff-login')?.addEventListener('click', async () => {
    const sId = document.getElementById('staff-id') as HTMLInputElement;
    const sPin = document.getElementById('staff-pin') as HTMLInputElement;
    const id = sId?.value.trim(); 
    const pin = sPin?.value.trim();
    if (!id) return showToast('Ingrese documento', 'warning');
    
    const staff = staffDict[id];
    if (staff && pin === (staff.clave || staff.id)) {
        currentStaff = staff; 
        const sDisp = document.getElementById('current-staff-display');
        if (sDisp) sDisp.innerHTML = `<i class="fas fa-user-check mr-1"></i> ${currentStaff.nombre}`;
        applyRolesUI(currentStaff); 
        switchTab('scanner'); 
        showToast(`¡Hola, ${staff.nombre}!`, 'success');
        
        // Recuperar planilla guardada en la nube si existe
        try {
            const snap = await getDoc(doc(getPlanillasPath(), currentStaff.id));
            if (snap.exists()) {
                const data = snap.data();
                masterFileName = data.fileName;
                const res = await fetch(data.base64);
                const blob = await res.blob();
                const arrayBuffer = await blob.arrayBuffer();
                masterWorkbook = XLSX.read(arrayBuffer, { type: 'array' });
                const matchedCount = loadMasterMappings();
                
                const fnDisp = document.getElementById('master-file-name-display');
                if (fnDisp) fnDisp.innerText = masterFileName;
                const gcDisp = document.getElementById('master-grades-counter');
                if (gcDisp) gcDisp.innerText = `Planilla en vivo: ${matchedCount} alumnos listos`;
                
                document.getElementById('master-upload-wrapper')?.classList.add('hidden');
                document.getElementById('master-actions-wrapper')?.classList.remove('hidden');
                document.getElementById('btn-reset-master')?.classList.remove('hidden');
            }
        } catch(e) {}
    } else {
        showToast('Usuario o clave incorrecta.', 'error');
    }
});

document.getElementById('btn-staff-logout')?.addEventListener('click', () => {
    currentStaff = null; 
    stopScanner(); 
    switchTab('login'); 
    const sDisp = document.getElementById('current-staff-display');
    if (sDisp) sDisp.innerText = 'Sin iniciar sesión';
    const sId = document.getElementById('staff-id') as HTMLInputElement;
    const sPin = document.getElementById('staff-pin') as HTMLInputElement;
    if (sId) sId.value = ''; 
    if (sPin) sPin.value = '';
});

// Cambiar Contraseña de Personal
document.getElementById('btn-open-change-password')?.addEventListener('click', () => {
    document.getElementById('modal-cambiar-password')?.classList.remove('hidden');
    const oldP = document.getElementById('input-old-password') as HTMLInputElement;
    const newP = document.getElementById('input-new-password') as HTMLInputElement;
    if (oldP) oldP.value = '';
    if (newP) newP.value = '';
});
document.getElementById('btn-cancel-password')?.addEventListener('click', () => {
    document.getElementById('modal-cambiar-password')?.classList.add('hidden');
});
document.getElementById('btn-save-password')?.addEventListener('click', async () => {
    const oldP = (document.getElementById('input-old-password') as HTMLInputElement)?.value.trim();
    const newP = (document.getElementById('input-new-password') as HTMLInputElement)?.value.trim();
    if(!oldP || !newP) return showToast('Llena ambos campos', 'warning');
    if(oldP !== (currentStaff.clave || currentStaff.id)) return showToast('Contraseña actual incorrecta', 'error');
    
    const btn = document.getElementById('btn-save-password') as HTMLButtonElement; 
    btn.disabled = true; 
    btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i>';
    try {
        await updateDoc(doc(getStaffPath(), currentStaff.id), { clave: newP });
        staffDict[currentStaff.id].clave = newP; 
        currentStaff.clave = newP;
        showToast('Contraseña actualizada con éxito.', 'success');
        document.getElementById('modal-cambiar-password')?.classList.add('hidden');
    } catch(e) {
        showToast('Error de conexión', 'error');
    } finally {
        btn.disabled = false; 
        btn.innerHTML = '<i class="fas fa-save mr-2"></i>Guardar';
    }
});

// Reportes Docente
document.getElementById('btn-open-teacher-reports')?.addEventListener('click', () => { 
    document.getElementById('modal-docente-reportes')?.classList.remove('hidden'); 
});
document.getElementById('btn-close-docente-reportes')?.addEventListener('click', () => { 
    document.getElementById('modal-docente-reportes')?.classList.add('hidden'); 
});

document.getElementById('btn-download-teacher')?.addEventListener('click', async () => {
    const btn = document.getElementById('btn-download-teacher') as HTMLButtonElement; 
    const origTxt = btn.innerHTML; 
    btn.innerHTML = '<i class="fas fa-spinner fa-spin mr-2"></i>Generando...'; 
    btn.disabled = true;
    const sDateStart = (document.getElementById('teacher-report-date-start') as HTMLInputElement)?.value; 
    const sDateEnd = (document.getElementById('teacher-report-date-end') as HTMLInputElement)?.value; 
    const rType = (document.getElementById('teacher-report-type') as HTMLSelectElement)?.value; 
    const sGroup = (document.getElementById('teacher-report-group') as HTMLSelectElement)?.value || 'todos';
    
    if (!sDateStart || !sDateEnd) { 
        showToast('Seleccione un rango de fechas', 'warning'); 
        btn.innerHTML = origTxt; 
        btn.disabled = false; 
        return; 
    }

    try {
        let targetStudents = Object.values(studentsDict);
        if (sGroup !== 'todos') targetStudents = targetStudents.filter(s => s.grado === sGroup);
        targetStudents.sort((a, b) => getFullName(a).localeCompare(getFullName(b)));
        if (targetStudents.length === 0) { 
            showToast('No hay estudiantes en este grupo.', 'warning'); 
            btn.innerHTML = origTxt; 
            btn.disabled = false; 
            return; 
        }

        const dateList: string[] = []; 
        let currDate = new Date(sDateStart); 
        const endDate = new Date(sDateEnd);
        while (currDate <= endDate) { 
            dateList.push(currDate.toISOString().split('T')[0]); 
            currDate.setDate(currDate.getDate() + 1); 
        }

        const wb = XLSX.utils.book_new();
        const fetchRange = async (pathFn: any) => {
            const snap = await getDocs(pathFn()); 
            const recs: any[] = [];
            snap.forEach(d => { 
                const data: any = d.data(); 
                if (data.date >= sDateStart && data.date <= sDateEnd) recs.push(data); 
            }); 
            return recs;
        };

        if (rType === 'asistencia') { 
            const recs = await fetchRange(getAttendancePath); 
            const wsData: any[][] = [[institucionData.nombre.toUpperCase()], [`MATRIZ DE ASISTENCIA - GRUPO: ${sGroup.toUpperCase()}`], [], ["Documento", "Matrícula", "Nombres", "Grado", ...dateList]];
            targetStudents.forEach(st => {
                const row = [st.documento || st.id, st.matricula || '', getFullName(st), st.grado];
                dateList.forEach(d => {
                    const rec = recs.find((r: any) => r.date === d && r.studentId === st.id);
                    row.push(rec ? formatTime(new Date(rec.timestamp)) : "");
                });
                wsData.push(row);
            });
            const ws = XLSX.utils.aoa_to_sheet(wsData);
            XLSX.utils.book_append_sheet(wb, ws, "Asistencia"); 
            XLSX.writeFile(wb, `Matriz_Asistencia_${sGroup}.xlsx`); 
        } else if (rType === 'evaluacion') { 
            const recs = await fetchRange(getEvaluacionesPath); 
            const wsData: any[][] = [[institucionData.nombre.toUpperCase()], [`NOTAS DE EVALUACIÓN - GRUPO: ${sGroup.toUpperCase()} - DOCENTE: ${currentStaff.nombre}`], [], ["Documento", "Matrícula", "Nombres", "Grado", ...dateList]];
            targetStudents.forEach(st => {
                const row = [st.documento || st.id, st.matricula || '', getFullName(st), st.grado];
                dateList.forEach(d => {
                    const rec = recs.find((r: any) => r.date === d && r.studentId === st.id && r.recordedById === currentStaff.id);
                    row.push(rec ? rec.nota : "");
                });
                wsData.push(row);
            });
            const ws = XLSX.utils.aoa_to_sheet(wsData);
            XLSX.utils.book_append_sheet(wb, ws, "Notas"); 
            XLSX.writeFile(wb, `Planilla_Notas_${sGroup}.xlsx`); 
        } else {
            const pathMap: Record<string, any> = { pae: getMealsPath, pruebas: getPruebasPath, prestamo: getLoansPath, salida: getSalidasPath, novedad: getNewsPath };
            const fn = pathMap[rType] || getAttendancePath;
            const recs = await fetchRange(fn);
            const wsData: any[][] = [[institucionData.nombre.toUpperCase()], [`REPORTE ${rType.toUpperCase()} - GRUPO: ${sGroup.toUpperCase()}`], []];
            const formatted = recs.filter((r: any) => (sGroup === 'todos' || (r.studentGrado || r.grado) === sGroup)).map((rec: any) => ({
                Fecha: rec.date,
                Hora: formatTime(new Date(rec.timestamp)),
                Documento: rec.studentId,
                Nombre: rec.studentName || rec.nombre,
                Grado: rec.studentGrado || rec.grado,
                Detalle: rec.nota || rec.itemName || rec.observacion || rec.status || 'OK'
            }));
            const ws = (XLSX.utils as any).json_to_sheet(formatted, { origin: "A4" }); 
            XLSX.utils.sheet_add_aoa(ws, wsData, { origin: "A1" }); 
            XLSX.utils.book_append_sheet(wb, ws, "Reporte");
            XLSX.writeFile(wb, `Reporte_${rType}_${sGroup}.xlsx`);
        }
        showToast('¡Reporte generado!', 'success');
    } catch(e) { 
        showToast('Error generando reporte', 'error'); 
    } finally { 
        btn.innerHTML = origTxt; 
        btn.disabled = false; 
    }
});

// Grupos para Códigos QR
export const populateQRGroups = () => {
    const unique = new Set<string>(); 
    Object.values(studentsDict).forEach(i => { if (i.grado) unique.add(i.grado); });
    
    const adminGrp = document.getElementById('qr-group'); 
    if (adminGrp) adminGrp.innerHTML = '<option value="todos">Todos los grupos</option>';
    const teacherGrp = document.getElementById('teacher-report-group'); 
    if (teacherGrp) teacherGrp.innerHTML = '<option value="todos">Todos</option>';
    
    Array.from(unique).sort().forEach(g => { 
        if (adminGrp) adminGrp.innerHTML += `<option value="${g}">${g}</option>`; 
        if (teacherGrp) teacherGrp.innerHTML += `<option value="${g}">Grupo ${g}</option>`; 
    });
};

// Impresión Masiva de Códigos QR
const printMassiveQRs = (itemsArray: any[], title: string) => {
    if(itemsArray.length === 0) return showToast('No hay registros en este grupo para imprimir', 'warning');
    showToast('Generando ventana de impresión...', 'success'); 
    const pw = window.open('', '_blank');
    if (!pw) return;
    
    let html = `<!DOCTYPE html><html><head><title>${title}</title><script src="https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js"><\/script><style>body{font-family:sans-serif;margin:0;padding:15px;}.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:15px;}.card{border:2px dashed #ccc;padding:10px;text-align:center;display:flex;flex-direction:column;align-items:center;justify-content:space-between;page-break-inside:avoid;height:260px;}.name{font-weight:900;font-size:13px;text-transform:uppercase;margin-bottom:5px;min-height:32px;display:flex;align-items:center;justify-content:center;width:100%;line-height:1.1;color:#1f2937;}.qr{width:140px;height:140px;margin:5px 0;}.id{font-size:15px;font-weight:bold;background:#f3f4f6;padding:4px;width:100%;box-sizing:border-box;border-top:2px solid #000;border-bottom:2px solid #000;margin-bottom:3px;letter-spacing:1px;color:#111827;}.group{font-size:12px;text-transform:uppercase;font-weight:bold;color:#4b5563;}@media print{.grid{grid-template-columns:repeat(4,1fr);}@page{margin:1cm;}}</style></head><body><h2 style="text-align:center;font-family:sans-serif;text-transform:uppercase;margin-bottom:20px;color:#374151;">${title}</h2><div class="grid">`;
    
    itemsArray.forEach((item, idx) => { 
        const nameStr = item.nombres ? `${item.nombres} ${item.apellidos || ''}` : (item.nombre || 'SIN NOMBRE'); 
        const groupStr = item.grado || item.tipo || 'INVENTARIO'; 
        html += `<div class="card"><div class="name">${nameStr}</div><div class="qr" id="qr-${idx}"></div><div class="id">${item.id}</div><div class="group">${groupStr}</div></div>`; 
    });
    
    html += `</div><script>window.onload=function(){`;
    itemsArray.forEach((item, idx) => { 
        html += `new QRCode(document.getElementById('qr-${idx}'), { text: "${item.id}", width: 140, height: 140 });\n`; 
    });
    html += `setTimeout(function(){window.print();},2500);};<\/script></body></html>`;
    pw.document.write(html); 
    pw.document.close();
};

document.getElementById('btn-qr-print')?.addEventListener('click', () => { 
    const grpSelect = document.getElementById('qr-group') as HTMLSelectElement;
    const group = grpSelect?.value || 'todos'; 
    let studentsToPrint = Object.values(studentsDict);
    if (group !== 'todos') studentsToPrint = studentsToPrint.filter(s => s.grado === group);
    studentsToPrint.sort((a, b) => getFullName(a).localeCompare(getFullName(b)));
    const docTitle = group === 'todos' ? 'Impresión Masiva - Todos los Alumnos' : `Impresión Masiva - Grupo ${group}`;
    printMassiveQRs(studentsToPrint, docTitle);
});

document.getElementById('btn-qr-items')?.addEventListener('click', () => { 
    const itemsToPrint = Object.values(inventarioDict); 
    itemsToPrint.sort((a, b) => (a.nombre || '').localeCompare(b.nombre || ''));
    printMassiveQRs(itemsToPrint, 'Impresión Masiva - Inventario de Elementos');
});

// Búsqueda QR Individual
document.getElementById('btn-qr-search')?.addEventListener('click', () => {
    const qInput = document.getElementById('qr-search-input') as HTMLInputElement;
    const q = qInput?.value.trim(); 
    if(!q) return showToast('Ingresa un documento o matrícula', 'warning');
    
    let item: any = studentsDict[q] || inventarioDict[q] || staffDict[q]; 
    if (!item) {
        // Buscar por matrícula
        item = Object.values(studentsDict).find(s => String(s.matricula).trim() === q);
    }
    if(!item) return showToast('Registro no encontrado', 'error');
    
    const line1 = item.nombres || item.nombre || 'SIN NOMBRE'; 
    const line2 = item.apellidos || ''; 
    const group = item.grado || item.tipo || 'INVENTARIO';
    
    const qrName = document.getElementById('qr-ind-name');
    if (qrName) qrName.innerHTML = `${line1}<br>${line2}`; 
    const qrId = document.getElementById('qr-ind-id');
    if (qrId) qrId.innerText = item.id; 
    const qrGrp = document.getElementById('qr-ind-group');
    if (qrGrp) qrGrp.innerText = group;
    
    const qrImg = document.getElementById('qr-ind-img');
    if (qrImg) {
        qrImg.innerHTML = '';
        if (window.QRCode) {
            new window.QRCode(qrImg, { text: String(item.id), width: 130, height: 130 });
        }
    }
    document.getElementById('qr-individual-result')?.classList.remove('hidden');
});

document.getElementById('btn-qr-ind-print')?.addEventListener('click', () => {
    const name = document.getElementById('qr-ind-name')?.innerHTML || ''; 
    const id = document.getElementById('qr-ind-id')?.innerText || ''; 
    const group = document.getElementById('qr-ind-group')?.innerText || '';
    const pw = window.open('', '_blank', 'width=350,height=500');
    if (!pw) return;
    pw.document.write(`<!DOCTYPE html><html><head><script src="https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js"><\/script><style>body{font-family:sans-serif;margin:0;padding:5px 0;text-align:center;width:200px;display:flex;flex-direction:column;align-items:center;} .name{font-weight:900;font-size:13px;text-transform:uppercase;margin-bottom:6px;width:100%;color:#000;} .qr{width:150px;height:150px;margin:0;} .id{font-size:15px;font-weight:bold;background:#e5e7eb;padding:4px;margin-top:6px;width:100%;box-sizing:border-box;border-top:2px solid #000;border-bottom:2px solid #000;color:#000;} .group{font-size:12px;margin-top:4px;text-transform:uppercase;color:#000;font-weight:bold;}</style></head><body><div class="name">${name}</div><div class="qr" id="qr-ind-print"></div><div class="id">${id}</div><div class="group">${group}</div><script>window.onload=function(){new QRCode(document.getElementById('qr-ind-print'), { text: "${id}", width: 150, height: 150 }); setTimeout(function(){window.print();},1000);};<\/script></body></html>`);
    pw.document.close();
});

// Inicialización de la Aplicación
const initApp = async () => {
    const dDisplay = document.getElementById('date-display');
    if (dDisplay) {
        dDisplay.innerText = new Date().toLocaleDateString('es-CO', { weekday: 'short', month: 'short', day: 'numeric' });
    }
    
    try { 
        await signInAnonymously(auth); 
        const localData = localStorage.getItem('institucionData'); 
        if (localData) { 
            institucionData = JSON.parse(localData); 
            aplicarConfiguracionUI(); 
        }
        await loadDatabases();
    } catch (error) { 
        console.warn("Conexión inicial Firebase:", error);
        loadDatabases();
    }
};

initApp();
