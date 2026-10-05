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
import { jsPDF } from "jspdf";

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
    cargo?: string;
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

// Rutas de Colecciones y Estado Multi-Tenant (Ramón Múnera vs Modo Prueba 30 Días)
export type AccessMode = 'ramon_munera' | 'modo_prueba';
let currentAccessMode: AccessMode = 'ramon_munera';

const getSettingsPath = () => doc(db, 'configuracion_app', currentAccessMode === 'modo_prueba' ? 'institucion_prueba' : 'institucion');
const getStudentsPath = () => collection(db, currentAccessMode === 'modo_prueba' ? 'prueba_estudiantes' : 'estudiantes');
const getStaffPath = () => collection(db, currentAccessMode === 'modo_prueba' ? 'prueba_docentes' : 'docentes');
const getInventarioPath = () => collection(db, currentAccessMode === 'modo_prueba' ? 'prueba_inventario' : 'inventario');
const getPAEBenefPath = () => collection(db, currentAccessMode === 'modo_prueba' ? 'prueba_pae_beneficiarios' : 'pae_beneficiarios');
const getPermissionsPath = () => doc(db, 'configuracion_app', currentAccessMode === 'modo_prueba' ? 'permisos_roles_prueba' : 'permisos_roles');

const getMealsPath = () => collection(db, currentAccessMode === 'modo_prueba' ? 'prueba_registros_pae' : 'registros_pae');
const getAttendancePath = () => collection(db, currentAccessMode === 'modo_prueba' ? 'prueba_registros_asistencia' : 'registros_asistencia');
const getLoansPath = () => collection(db, currentAccessMode === 'modo_prueba' ? 'prueba_registros_prestamos' : 'registros_prestamos');
const getNewsPath = () => collection(db, currentAccessMode === 'modo_prueba' ? 'prueba_registros_novedades' : 'registros_novedades');
const getSalidasPath = () => collection(db, currentAccessMode === 'modo_prueba' ? 'prueba_registros_salidas' : 'registros_salidas');
const getPruebasPath = () => collection(db, currentAccessMode === 'modo_prueba' ? 'prueba_registros_pruebas' : 'registros_pruebas');
const getEvaluacionesPath = () => collection(db, currentAccessMode === 'modo_prueba' ? 'prueba_registros_evaluaciones' : 'registros_evaluaciones');
const getPlanillasPath = () => collection(db, currentAccessMode === 'modo_prueba' ? 'prueba_planillas_docentes' : 'planillas_docentes');

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

// Perfil de Institución y Licencia Multi-Colegio
export type LicenseType = 'institucional' | 'docente' | 'prueba';

export interface InstitutionProfile {
    id: string;
    nombre: string;
    logo: string | null;
    color1: string;
    color2: string;
    licenciaInicio: string;
    licenciaFin: string;
    tipoPlan: LicenseType;
    limiteUsuarios: number;
    titularNombre?: string;
    titularDoc?: string;
    titularContacto?: string;
    permisosPersonalizados?: Record<string, string[]>;
    popupActivo: boolean;
    popupTitulo: string;
    popupUrl: string;
    popupDescripcion: string;
}

export const RAMON_MUNERA_PROFILE: InstitutionProfile = {
    id: "ramon_munera",
    nombre: "I.E. Ramón Múnera Lopera",
    logo: null,
    color1: "#2563eb",
    color2: "#0ea5e9",
    licenciaInicio: "2026-01-01",
    licenciaFin: "2027-12-31",
    tipoPlan: "institucional",
    limiteUsuarios: 9999,
    titularNombre: "I.E. Ramón Múnera Lopera",
    popupActivo: true,
    popupTitulo: "Tutorial & Carnet Digital Institucional",
    popupUrl: "https://www.youtube.com/watch?v=dQw4w9WgXcQ",
    popupDescripcion: "Bienvenido al sistema institucional de la I.E. Ramón Múnera Lopera. Consulta el tutorial de uso del carnet y los módulos académicos."
};

export const TRIAL_INSTITUTION_DEFAULT: InstitutionProfile = {
    id: "modo_prueba",
    nombre: "Institución Educativa (Modo Prueba 30 Días)",
    logo: null,
    color1: "#4f46e5",
    color2: "#06b6d4",
    licenciaInicio: "2026-01-01",
    licenciaFin: "2026-12-31",
    tipoPlan: "prueba",
    limiteUsuarios: 9999,
    titularNombre: "Institución Educativa (Modo Prueba)",
    popupActivo: true,
    popupTitulo: "¡Bienvenido al Modo Prueba de 30 Días!",
    popupUrl: "https://www.youtube.com/watch?v=dQw4w9WgXcQ",
    popupDescripcion: "Esta es tu versión de demostración por 30 días para evaluar el Gestor Académico. Puedes personalizar el nombre de tu colegio, logo y Pop-up en Administración. Desarrollada por www.espatodo.com"
};

// Datos Demo Precargados para Modo Prueba Inmediata (Sin tocar Ramón Múnera)
export const DEMO_STUDENTS: Record<string, StudentRecord> = {
    "1032033800": {
        id: "1032033800",
        matricula: "261047",
        tipo_documento: "R.C.",
        documento: "1032033800",
        nombres: "JUSTIN ANDRES",
        apellidos: "BERNAL GARCIA",
        nombre_completo: "BERNAL GARCIA JUSTIN ANDRES",
        nombre: "BERNAL GARCIA JUSTIN ANDRES",
        grado: "TS0501"
    },
    "1032033801": {
        id: "1032033801",
        matricula: "261048",
        tipo_documento: "T.I.",
        documento: "1032033801",
        nombres: "JUAN CARLOS",
        apellidos: "DE LA OSSA PEREZ",
        nombre_completo: "DE LA OSSA PEREZ JUAN CARLOS",
        nombre: "DE LA OSSA PEREZ JUAN CARLOS",
        grado: "TS0501"
    },
    "1032033802": {
        id: "1032033802",
        matricula: "261049",
        tipo_documento: "T.I.",
        documento: "1032033802",
        nombres: "VALENTINA",
        apellidos: "ZAPATA MARTINEZ",
        nombre_completo: "ZAPATA MARTINEZ VALENTINA",
        nombre: "ZAPATA MARTINEZ VALENTINA",
        grado: "TS0501"
    },
    "1032033901": {
        id: "1032033901",
        matricula: "262001",
        tipo_documento: "T.I.",
        documento: "1032033901",
        nombres: "CAMILO",
        apellidos: "GOMEZ ALVAREZ",
        nombre_completo: "GOMEZ ALVAREZ CAMILO",
        nombre: "GOMEZ ALVAREZ CAMILO",
        grado: "TS0502"
    },
    "1032033902": {
        id: "1032033902",
        matricula: "262002",
        tipo_documento: "T.I.",
        documento: "1032033902",
        nombres: "SOFIA MARIANA",
        apellidos: "RODRIGUEZ LOPEZ",
        nombre_completo: "RODRIGUEZ LOPEZ SOFIA MARIANA",
        nombre: "RODRIGUEZ LOPEZ SOFIA MARIANA",
        grado: "TS0502"
    }
};

export const DEMO_STAFF: Record<string, StaffRecord> = {
    "1234": {
        id: "1234",
        nombre: "Profesor de Demostración",
        cargo: "docente",
        rol: "docente",
        clave: "1234"
    },
    "admin": {
        id: "admin",
        nombre: "Administrador Demo",
        cargo: "administrador",
        rol: "administrador",
        clave: "1234"
    }
};

// Lista de instituciones registradas
let institucionesList: InstitutionProfile[] = [
    { ...RAMON_MUNERA_PROFILE },
    { ...TRIAL_INSTITUTION_DEFAULT }
];

let institucionData: InstitutionProfile = { ...institucionesList[0] };
let isLicenseValid = true;
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

// ==========================================================
// CONTROL UNIVERSAL DE MODO PRUEBA DE 30 DÍAS (FREE TRIAL)
// ==========================================================
export interface TrialStatus {
    startDate: string;
    totalDays: number;
    elapsedDays: number;
    remainingDays: number;
    isExpired: boolean;
}

export const getTrialStatus = (): TrialStatus => {
    let trialData: { startDate: string; days: number } | null = null;
    const raw = localStorage.getItem('espatodo_trial_setup_v1');
    if (raw) {
        try {
            trialData = JSON.parse(raw);
        } catch(e) {}
    }
    
    if (!trialData || !trialData.startDate) {
        trialData = {
            startDate: getTodayString(),
            days: 30
        };
        localStorage.setItem('espatodo_trial_setup_v1', JSON.stringify(trialData));
    }
    
    const startMs = new Date(trialData.startDate + 'T00:00:00').getTime();
    const todayMs = new Date(getTodayString() + 'T00:00:00').getTime();
    const diffTime = todayMs - startMs;
    const elapsedDays = Math.max(0, Math.floor(diffTime / (1000 * 60 * 60 * 24)));
    const remainingDays = Math.max(0, trialData.days - elapsedDays);
    const isExpired = remainingDays <= 0;
    
    return {
        startDate: trialData.startDate,
        totalDays: trialData.days,
        elapsedDays,
        remainingDays,
        isExpired
    };
};

export const switchAccessMode = async (mode: AccessMode) => {
    currentAccessMode = mode;
    localStorage.setItem('active_access_mode', mode);
    
    if (mode === 'modo_prueba') {
        const localTrialProf = localStorage.getItem('trial_institution_profile');
        if (localTrialProf) {
            try {
                institucionData = JSON.parse(localTrialProf);
            } catch(e) {
                institucionData = { ...TRIAL_INSTITUTION_DEFAULT };
            }
        } else {
            institucionData = { ...TRIAL_INSTITUTION_DEFAULT };
        }
    } else {
        const found = institucionesList.find(i => i.id === 'ramon_munera');
        institucionData = found ? { ...found } : { ...RAMON_MUNERA_PROFILE };
    }
    
    await loadDatabases();
    aplicarConfiguracionUI();
    checkLicenseValidity();
    
    if (institucionData.popupActivo) {
        setTimeout(() => showStartupPopup(false), 500);
    }

    showToast(mode === 'modo_prueba' 
        ? 'Modo Prueba (30 Días) activado. Entorno limpio e independiente.' 
        : 'Conectado a la I.E. Ramón Múnera Lopera (Licencia Anual).', 'info', 4000);
};

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
            
            // Si está en Modo Prueba (30 Días), guardar en la partición aislada de prueba
            if (currentAccessMode === 'modo_prueba') {
                showToast(`Guardando ${parsedStudents.length} alumnos en tu base de prueba...`, 'info', 3000);
                studentsDict = {};
                parsedStudents.forEach(item => {
                    studentsDict[item.id] = item;
                });
                localStorage.setItem('trial_students', JSON.stringify(parsedStudents));
                if (countEl) countEl.innerText = `${parsedStudents.length} alumnos (Prueba)`;
                if (progressEl) setTimeout(() => progressEl.style.width = "0%", 1000);
                populateQRGroups();
                showToast(`¡Excelente! ${parsedStudents.length} estudiantes importados en tu colegio de prueba.`, 'success', 5000);
                return;
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
    // -------------------------------------------------------------
    // RAMA 1: MODO PRUEBA DE 30 DÍAS (AISLADO - NO TOCA RAMÓN MÚNERA)
    // -------------------------------------------------------------
    if (currentAccessMode === 'modo_prueba') {
        const localTrialProf = localStorage.getItem('trial_institution_profile');
        if (localTrialProf) {
            try {
                institucionData = JSON.parse(localTrialProf);
            } catch(e) {
                institucionData = { ...TRIAL_INSTITUTION_DEFAULT };
            }
        } else {
            institucionData = { ...TRIAL_INSTITUTION_DEFAULT };
        }

        // Cargar alumnos de prueba
        const localTrialStudents = localStorage.getItem('trial_students');
        studentsDict = {};
        if (localTrialStudents) {
            try {
                const list = JSON.parse(localTrialStudents);
                if (Array.isArray(list)) {
                    list.forEach((s: any) => { studentsDict[s.id] = s; });
                } else if (typeof list === 'object') {
                    studentsDict = list;
                }
            } catch(e) {}
        }
        
        // Si no hay alumnos en prueba, precargar demo para test inmediato
        if (Object.keys(studentsDict).length === 0) {
            studentsDict = { ...DEMO_STUDENTS };
            localStorage.setItem('trial_students', JSON.stringify(Object.values(DEMO_STUDENTS)));
        }

        // Cargar docentes de prueba
        const localTrialStaff = localStorage.getItem('trial_staff');
        staffDict = {};
        if (localTrialStaff) {
            try {
                staffDict = JSON.parse(localTrialStaff);
            } catch(e) {}
        }
        if (Object.keys(staffDict).length === 0) {
            staffDict = { ...DEMO_STAFF };
            localStorage.setItem('trial_staff', JSON.stringify(DEMO_STAFF));
        }

        inventarioDict = {};
        paeBeneficiariosDict = {};

        const elStud = document.getElementById('count-students'); 
        if (elStud) elStud.innerText = Object.keys(studentsDict).length + ' alumnos (Prueba)';
        const elStaff = document.getElementById('count-staff'); 
        if (elStaff) elStaff.innerText = Object.keys(staffDict).length + ' personal (Prueba)';
        const elItem = document.getElementById('count-items'); 
        if (elItem) elItem.innerText = '0 registros';
        const elPae = document.getElementById('count-pae'); 
        if (elPae) elPae.innerText = '0 registros';

        renderRolesConfig();
        populateQRGroups();
        aplicarConfiguracionUI();
        checkLicenseValidity();
        return;
    }

    // -------------------------------------------------------------
    // RAMA 2: INSTITUCIÓN OFICIAL (I.E. RAMÓN MÚNERA LOPERA)
    // -------------------------------------------------------------
    try {
        const settingsSnap = await getDoc(getSettingsPath());
        if (settingsSnap.exists()) { 
            const data: any = settingsSnap.data();
            if (data.instituciones && Array.isArray(data.instituciones) && data.instituciones.length > 0) {
                institucionesList = data.instituciones;
            }
            
            const urlParams = new URLSearchParams(window.location.search);
            const reqInst = urlParams.get('inst') || urlParams.get('institucion');
            let activeInst: InstitutionProfile | undefined;
            if (reqInst) {
                activeInst = institucionesList.find(i => i.id === reqInst || normalizeNameMatch(i.nombre).includes(normalizeNameMatch(reqInst)));
            }
            if (!activeInst && data.institucionActivaId) {
                activeInst = institucionesList.find(i => i.id === data.institucionActivaId);
            }
            if (!activeInst && data.institucionActiva) {
                activeInst = data.institucionActiva;
            }
            if (!activeInst && data.nombre) {
                activeInst = data as InstitutionProfile;
            }
            
            if (activeInst) {
                institucionData = { ...activeInst };
            } else if (institucionesList.length > 0) {
                institucionData = { ...institucionesList[0] };
            }

            // Cargar configuración de Super-Administrador comercial si existe
            try {
                const saSnap = await getDoc(doc(db, 'configuracion_app', 'superadmin_config'));
                if (saSnap.exists()) {
                    const saData = saSnap.data();
                    if (saData.tipoPlan) institucionData.tipoPlan = saData.tipoPlan;
                    if (saData.licenciaInicio) institucionData.licenciaInicio = saData.licenciaInicio;
                    if (saData.licenciaFin) institucionData.licenciaFin = saData.licenciaFin;
                    if (saData.titularNombre) institucionData.titularNombre = saData.titularNombre;
                    if (saData.titularDoc) institucionData.titularDoc = saData.titularDoc;
                    if (saData.titularContacto) institucionData.titularContacto = saData.titularContacto;
                    if (saData.rolePermissions) rolePermissions = saData.rolePermissions;
                }
            } catch(e) {}
            
            aplicarConfiguracionUI(); 
            checkLicenseValidity();
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
    document.getElementById('sp-header-bar')?.classList.add('hidden'); // Ocultar barra superior para no duplicar 'Volver'
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

// Selector de Pestañas de Acceso Compacto (Optimizado para Celular)
export const switchAccessTab = (tab: 'staff' | 'student') => {
    const tabBtnStaff = document.getElementById('tab-btn-staff');
    const tabBtnStudent = document.getElementById('tab-btn-student');
    const panelStaff = document.getElementById('tab-panel-staff');
    const panelStudent = document.getElementById('tab-panel-student');

    if (tab === 'staff') {
        panelStaff?.classList.remove('hidden');
        panelStudent?.classList.add('hidden');
        if (tabBtnStaff) {
            tabBtnStaff.className = "flex-1 py-2 px-2 rounded-lg font-black text-xs flex items-center justify-center gap-1.5 transition-all bg-white shadow-sm";
            tabBtnStaff.setAttribute('style', 'color: #1d4ed8 !important; font-weight: 900 !important;');
        }
        if (tabBtnStudent) {
            tabBtnStudent.className = "flex-1 py-2 px-2 rounded-lg text-xs flex items-center justify-center gap-1.5 transition-all hover:bg-gray-100";
            tabBtnStudent.setAttribute('style', 'color: #374151 !important; font-weight: 800 !important;');
        }
    } else {
        panelStaff?.classList.add('hidden');
        panelStudent?.classList.remove('hidden');
        if (tabBtnStudent) {
            tabBtnStudent.className = "flex-1 py-2 px-2 rounded-lg font-black text-xs flex items-center justify-center gap-1.5 transition-all bg-white shadow-sm";
            tabBtnStudent.setAttribute('style', 'color: #065f46 !important; font-weight: 900 !important;');
        }
        if (tabBtnStaff) {
            tabBtnStaff.className = "flex-1 py-2 px-2 rounded-lg text-xs flex items-center justify-center gap-1.5 transition-all hover:bg-gray-100";
            tabBtnStaff.setAttribute('style', 'color: #374151 !important; font-weight: 800 !important;');
        }
        const inp = document.getElementById('compact-student-id') as HTMLInputElement;
        if (inp) setTimeout(() => inp.focus(), 80);
    }
};

document.getElementById('tab-btn-staff')?.addEventListener('click', () => switchAccessTab('staff'));
document.getElementById('tab-btn-student')?.addEventListener('click', () => switchAccessTab('student'));

// Búsqueda directa desde el panel compacto de estudiante
document.getElementById('btn-compact-student-search')?.addEventListener('click', () => {
    const compactInput = document.getElementById('compact-student-id') as HTMLInputElement;
    const query = compactInput?.value.trim();
    if (!query) return showToast('Ingresa tu documento o matrícula', 'warning');

    const spInput = document.getElementById('sp-student-id') as HTMLInputElement;
    if (spInput) spInput.value = query;

    document.getElementById('staff-login-overlay')?.classList.add('hidden');
    document.getElementById('student-portal-overlay')?.classList.remove('hidden');
    document.getElementById('btn-sp-search')?.click();
});

document.getElementById('compact-student-id')?.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
        e.preventDefault();
        (document.getElementById('btn-compact-student-search') as HTMLButtonElement)?.click();
    }
});

// Portal del estudiante
document.getElementById('btn-open-student-portal')?.addEventListener('click', () => { 
    document.getElementById('staff-login-overlay')?.classList.add('hidden'); 
    document.getElementById('student-portal-overlay')?.classList.remove('hidden'); 
    document.getElementById('sp-header-bar')?.classList.remove('hidden'); 
    const inputSp = document.getElementById('sp-student-id') as HTMLInputElement;
    if (inputSp) inputSp.value = ''; 
    document.getElementById('sp-step-search')?.classList.remove('hidden'); 
    document.getElementById('sp-step-photo')?.classList.add('hidden'); 
    document.getElementById('sp-step-carnet')?.classList.add('hidden'); 
});

const returnToAccessView = () => { 
    document.getElementById('student-portal-overlay')?.classList.add('hidden'); 
    document.getElementById('staff-login-overlay')?.classList.remove('hidden'); 
    document.getElementById('sp-header-bar')?.classList.remove('hidden'); 
};

document.getElementById('btn-back-to-login')?.addEventListener('click', returnToAccessView);
document.getElementById('btn-carnet-back')?.addEventListener('click', returnToAccessView);

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
    document.getElementById('sp-header-bar')?.classList.remove('hidden'); 
    const preview = document.getElementById('sp-photo-preview') as HTMLImageElement;
    if (preview) { preview.classList.add('hidden'); preview.src = ''; }
    const photoInput = document.getElementById('sp-photo-input') as HTMLInputElement;
    if (photoInput) photoInput.value = ''; 
    document.getElementById('btn-sp-save-photo')?.classList.add('hidden'); 
    const welcome = document.getElementById('sp-welcome-name');
    if (welcome) welcome.innerText = `Actualizar Foto`; 
    document.getElementById('sp-step-photo')?.classList.remove('hidden'); 
});

// Descargar o Guardar Carnet Digital (Compatible 100% con iPhone / iPad / Android / PC)
let lastGeneratedCarnetBlob: Blob | null = null;
let lastGeneratedCarnetFileName = 'Carnet.jpg';

const openIosCarnetModal = (dataUrl: string, _fileName: string) => {
    const modal = document.getElementById('modal-ios-carnet');
    const imgPreview = document.getElementById('img-ios-carnet-preview') as HTMLImageElement;
    if (imgPreview) imgPreview.src = dataUrl;
    modal?.classList.remove('hidden');
};

const closeIosCarnetModal = () => {
    document.getElementById('modal-ios-carnet')?.classList.add('hidden');
};

document.getElementById('btn-close-ios-carnet')?.addEventListener('click', closeIosCarnetModal);
document.getElementById('btn-close-ios-carnet-bottom')?.addEventListener('click', closeIosCarnetModal);

document.getElementById('btn-ios-carnet-share')?.addEventListener('click', async () => {
    if (!lastGeneratedCarnetBlob) return;
    try {
        const file = new File([lastGeneratedCarnetBlob], lastGeneratedCarnetFileName, { type: 'image/jpeg' });
        if (navigator.canShare && navigator.canShare({ files: [file] })) {
            await navigator.share({
                files: [file],
                title: 'Carnet Digital',
                text: 'Carnet Digital Institucional'
            });
            showToast('Carnet compartido / guardado', 'success');
        } else {
            showToast('Mantén presionada la imagen arriba y selecciona "Guardar en Fotos"', 'info', 4000);
        }
    } catch(e) {
        showToast('Mantén presionada la imagen para Guardar en Fotos', 'info', 4000);
    }
});

// Botón para salir de la vista de carnet y volver a la pantalla principal
document.getElementById('btn-carnet-exit')?.addEventListener('click', () => {
    document.getElementById('btn-back-to-login')?.click();
});

document.getElementById('btn-download-carnet')?.addEventListener('click', async () => {
    const btn = document.getElementById('btn-download-carnet') as HTMLButtonElement; 
    const originalHtml = btn.innerHTML; 
    btn.innerHTML = '<i class="fas fa-spinner fa-spin mr-2"></i> Generando Carnet...'; 
    btn.disabled = true;
    
    const carnetNode = document.getElementById('carnet-card');
    if (!window.html2canvas || !carnetNode) {
        showToast('Componente de captura no listo', 'error');
        btn.innerHTML = originalHtml;
        btn.disabled = false;
        return;
    }

    try {
        const canvas: HTMLCanvasElement = await window.html2canvas(carnetNode, { 
            scale: 3, 
            useCORS: true, 
            backgroundColor: '#ffffff' 
        });

        const idEstudiante = currentStudentForCarnet?.id || currentStudentForCarnet?.matricula || 'Estudiante';
        const fileName = `Carnet_${idEstudiante}.jpg`;
        lastGeneratedCarnetFileName = fileName;

        canvas.toBlob(async (blob) => {
            if (!blob) {
                showToast('Error al procesar imagen del carnet', 'error');
                btn.innerHTML = originalHtml;
                btn.disabled = false;
                return;
            }

            lastGeneratedCarnetBlob = blob;
            const file = new File([blob], fileName, { type: 'image/jpeg' });
            const isIOS = /iPad|iPhone|iPod/.test(navigator.userAgent) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);

            // Intentar Web Share API con archivo (Nativo en iPhone / iOS para Guardar en Fotos o WhatsApp)
            if (navigator.canShare && navigator.canShare({ files: [file] })) {
                try {
                    await navigator.share({
                        files: [file],
                        title: 'Carnet Digital',
                        text: `Carnet Digital de ${currentStudentForCarnet?.nombres || currentStudentForCarnet?.nombre || idEstudiante}`
                    });
                    showToast('¡Listo! Carnet guardado / compartido con éxito', 'success');
                    btn.innerHTML = originalHtml;
                    btn.disabled = false;
                    return;
                } catch(shareErr: any) {
                    if (shareErr.name === 'AbortError') {
                        // Usuario canceló el menú nativo
                        btn.innerHTML = originalHtml;
                        btn.disabled = false;
                        return;
                    }
                    console.warn('Share API no disponible o rechazada:', shareErr);
                }
            }

            if (isIOS) {
                // En iPhone, abrir modal amigable con instrucciones y visor para "Guardar en Fotos"
                const dataUrl = canvas.toDataURL('image/jpeg', 0.95);
                openIosCarnetModal(dataUrl, fileName);
                btn.innerHTML = originalHtml;
                btn.disabled = false;
                return;
            }

            // Descarga directa tradicional para PC y Android
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = fileName;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            setTimeout(() => URL.revokeObjectURL(url), 1000);
            showToast('Carnet descargado con éxito', 'success');
            btn.innerHTML = originalHtml;
            btn.disabled = false;
        }, 'image/jpeg', 0.95);

    } catch (err) {
        console.error("Error al exportar carnet:", err);
        showToast('Error al generar la imagen', 'error');
        btn.innerHTML = originalHtml;
        btn.disabled = false;
    }
});

// ==========================================================
// HOJA EN PDF TAMAÑO CARTA CON 32 ETIQUETAS QR RECORTABLES
// PARA RECORTAR Y PEGAR EN TRABAJOS ESCRITOS Y TALLERES
// ==========================================================
export const generateStudentQrPdfSheet = async (user: any) => {
    if (!user || !user.id) {
        showToast('No hay datos de estudiante o usuario cargado', 'warning');
        return;
    }

    const btn = document.getElementById('btn-download-qr-sheet') as HTMLButtonElement;
    const originalHtml = btn ? btn.innerHTML : '';
    if (btn) {
        btn.innerHTML = '<div class="w-10 h-10 rounded-xl bg-white bg-opacity-20 flex items-center justify-center flex-shrink-0 text-white text-xl"><i class="fas fa-spinner fa-spin text-yellow-300"></i></div><div class="flex flex-col text-left flex-1 min-w-0"><span class="font-black text-xs sm:text-sm text-white">Generando Hoja de 32 QR...</span><span class="text-[10px] text-emerald-100">Ajustando tamaño carta</span></div>';
        btn.disabled = true;
    }

    try {
        const qrContent = String(user.id);
        
        // Contenedor temporal para generar el QR en alta resolución
        const tempDiv = document.createElement('div');
        tempDiv.style.position = 'fixed';
        tempDiv.style.left = '-9999px';
        tempDiv.style.top = '-9999px';
        document.body.appendChild(tempDiv);

        new (window as any).QRCode(tempDiv, {
            text: qrContent,
            width: 260,
            height: 260,
            correctLevel: 0 // L level: máxima sencillez y nitidez para escaneo óptico veloz
        });

        await new Promise((r) => setTimeout(r, 120));

        let qrDataUrl = '';
        const canvas = tempDiv.querySelector('canvas') as HTMLCanvasElement;
        if (canvas) {
            qrDataUrl = canvas.toDataURL('image/png');
        } else {
            const img = tempDiv.querySelector('img') as HTMLImageElement;
            if (img && img.src) qrDataUrl = img.src;
        }
        document.body.removeChild(tempDiv);

        if (!qrDataUrl) {
            showToast('No se pudo generar la imagen del código QR', 'error');
            if (btn) { btn.innerHTML = originalHtml; btn.disabled = false; }
            return;
        }

        // Inicializar documento PDF en Tamaño Carta (Letter: 215.9 x 279.4 mm)
        const doc = new jsPDF({
            orientation: 'portrait',
            unit: 'mm',
            format: 'letter'
        });

        const pageWidth = 215.9;
        const instName = (institucionData.nombre || 'INSTITUCIÓN EDUCATIVA').toUpperCase();
        const nombres = (user.nombres || user.nombre || 'ESTUDIANTE').toUpperCase();
        const apellidos = (user.apellidos || user.apellido || '').toUpperCase();
        const nombreCompleto = `${nombres} ${apellidos}`.trim();
        const grado = user.isStaff ? (user.rol || 'DOCENTE').toUpperCase() : `GRADO: ${user.grado || 'S/G'}`;
        const docId = `DOC: ${user.id}`;

        // ENCABEZADO SUPERIOR DEL DOCUMENTO
        doc.setFillColor(243, 244, 246);
        doc.rect(8, 6, pageWidth - 16, 17, 'F');
        doc.setDrawColor(209, 213, 219);
        doc.setLineWidth(0.3);
        doc.rect(8, 6, pageWidth - 16, 17, 'S');

        doc.setFont('helvetica', 'bold');
        doc.setFontSize(9.5);
        doc.setTextColor(31, 41, 55);
        doc.text(instName, 12, 11);

        doc.setFontSize(8.5);
        doc.setTextColor(29, 78, 216);
        doc.text('HOJA DE ETIQUETAS RECORTABLES CON CÓDIGO QR PARA EVALUACIÓN DE TRABAJOS', 12, 15);

        doc.setFont('helvetica', 'normal');
        doc.setFontSize(7);
        doc.setTextColor(75, 85, 99);
        doc.text(`Titular: ${nombreCompleto}  |  ${docId}  |  ${grado}  |  Plataforma: www.espatodo.com`, 12, 19);

        doc.setFontSize(6.5);
        doc.setTextColor(107, 114, 128);
        doc.text('✂ Recorta por la línea punteada y pega un recuadro a cada trabajo escrito para que el docente califique al escanear.', 12, 22.5);

        // GRID DE 32 ETIQUETAS (4 COLUMNAS x 8 FILAS)
        const cols = 4;
        const rows = 8;
        const cardWidth = 47.8;
        const cardHeight = 29.5;
        const startX = 9.5;
        const startY = 25.5;
        const gapX = 2.4;
        const gapY = 2.0;

        for (let r = 0; r < rows; r++) {
            for (let c = 0; c < cols; c++) {
                const x = startX + c * (cardWidth + gapX);
                const y = startY + r * (cardHeight + gapY);

                // Marco con línea punteada para recortar con tijeras
                doc.setDrawColor(156, 163, 175);
                doc.setLineWidth(0.25);
                doc.setLineDashPattern([1.2, 1.2], 0);
                doc.rect(x, y, cardWidth, cardHeight, 'S');

                // Restaurar línea sólida para elementos internos
                doc.setLineDashPattern([], 0);

                // Indicador de tijera guía
                doc.setFont('helvetica', 'normal');
                doc.setFontSize(6);
                doc.setTextColor(180, 180, 180);
                doc.text('✂', x + 1.2, y + 2.8);

                // CÓDIGO QR EN ALTA DEFINICIÓN (22.5 x 22.5 mm)
                const qrSize = 22.5;
                const qrX = x + 1.5;
                const qrY = y + 3.8;
                doc.addImage(qrDataUrl, 'PNG', qrX, qrY, qrSize, qrSize);

                // BLOQUE DERECHO: DATOS Y CASILLA DE NOTA
                const textStartX = x + 25.2;

                // Pastilla "TRABAJO ESCOLAR"
                doc.setFillColor(238, 242, 255);
                doc.roundedRect(textStartX, y + 2.2, 20.8, 3.8, 0.8, 0.8, 'F');
                doc.setFont('helvetica', 'bold');
                doc.setFontSize(5);
                doc.setTextColor(30, 58, 138);
                doc.text('TRABAJO ESCOLAR', textStartX + 10.4, y + 4.8, { align: 'center' });

                // Nombres y Apellidos
                doc.setFont('helvetica', 'bold');
                doc.setFontSize(6.2);
                doc.setTextColor(17, 24, 39);
                const nombreCorto = nombres.length > 14 ? nombres.substring(0, 13) + '..' : nombres;
                const apellidoCorto = apellidos.length > 14 ? apellidos.substring(0, 13) + '..' : apellidos;
                doc.text(nombreCorto, textStartX, y + 8.5);
                doc.text(apellidoCorto, textStartX, y + 11.5);

                // Documento
                doc.setFont('helvetica', 'normal');
                doc.setFontSize(5.5);
                doc.setTextColor(75, 85, 99);
                doc.text(`Doc: ${user.id}`, textStartX, y + 15);

                // Grado
                doc.setFont('helvetica', 'bold');
                doc.setFontSize(5.5);
                doc.setTextColor(29, 78, 216);
                doc.text(user.isStaff ? 'DOCENTE' : `Gr: ${user.grado || 'S/G'}`, textStartX, y + 18.2);

                // Recuadro para escribir la Nota con lapicero
                doc.setFillColor(250, 250, 250);
                doc.setDrawColor(180, 190, 205);
                doc.setLineWidth(0.2);
                doc.roundedRect(textStartX, y + 20.2, 20.8, 6.8, 0.8, 0.8, 'FD');

                doc.setFont('helvetica', 'bold');
                doc.setFontSize(5.5);
                doc.setTextColor(55, 65, 81);
                doc.text('Nota:', textStartX + 1.2, y + 24.2);

                // Línea de escritura
                doc.setDrawColor(156, 163, 175);
                doc.line(textStartX + 7.5, y + 24.8, textStartX + 19.5, y + 24.8);

                doc.setFont('helvetica', 'normal');
                doc.setFontSize(4.2);
                doc.setTextColor(156, 163, 175);
                doc.text('Escanear y Calificar', textStartX + 10.4, y + 26.2, { align: 'center' });
            }
        }

        // Descarga y compatibilidad con iPhone / Android / PC
        const pdfBlob = doc.output('blob');
        const fileName = `Hoja_QR_${user.id}_${nombres.replace(/\s+/g, '_')}.pdf`;
        const pdfFile = new File([pdfBlob], fileName, { type: 'application/pdf' });
        const isIOS = /iPad|iPhone|iPod/.test(navigator.userAgent) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);

        if (navigator.canShare && navigator.canShare({ files: [pdfFile] })) {
            try {
                await navigator.share({
                    files: [pdfFile],
                    title: 'Hoja de QR para Trabajos',
                    text: `Etiquetas QR recortables para trabajos escolares de ${nombreCompleto}`
                });
                showToast('Hoja de QR generada y lista', 'success');
                if (btn) { btn.innerHTML = originalHtml; btn.disabled = false; }
                return;
            } catch(shareErr: any) {
                if (shareErr.name === 'AbortError') {
                    if (btn) { btn.innerHTML = originalHtml; btn.disabled = false; }
                    return;
                }
                console.warn('Share API error on PDF, fallback to direct download:', shareErr);
            }
        }

        if (isIOS) {
            const blobUrl = URL.createObjectURL(pdfBlob);
            window.open(blobUrl, '_blank');
            showToast('PDF abierto para imprimir o guardar', 'success');
        } else {
            doc.save(fileName);
            showToast('Hoja de QR descargada con éxito (Tamaño Carta)', 'success');
        }

    } catch (err) {
        console.error('Error generando hoja PDF de QR:', err);
        showToast('Error al generar la hoja de QR en PDF', 'error');
    } finally {
        if (btn) {
            btn.innerHTML = originalHtml;
            btn.disabled = false;
        }
    }
};

document.getElementById('btn-download-qr-sheet')?.addEventListener('click', () => {
    if (!currentStudentForCarnet) {
        showToast('Primero consulta o selecciona un carnet', 'warning');
        return;
    }
    generateStudentQrPdfSheet(currentStudentForCarnet);
});

// ==========================================================
// CONFIGURACIÓN DE IDENTIDAD Y UI
// ==========================================================
// ==========================================================
// CONFIGURACIÓN DE IDENTIDAD, MULTI-INSTITUCIÓN Y POP-UP
// ==========================================================

// Parser inteligente para URLs multimedia (YouTube, Canva, Imágenes o Web)
export const parseMediaUrl = (rawUrl: string): { type: 'youtube' | 'canva' | 'image' | 'web'; embedUrl: string; directUrl: string } => {
    if (!rawUrl) return { type: 'web', embedUrl: '', directUrl: '' };
    const url = rawUrl.trim();

    // YouTube: Detectar watch?v=, youtu.be/, shorts/, embed/
    const ytMatch = url.match(/(?:youtu\.be\/|youtube\.com\/(?:embed\/|v\/|watch\?v=|watch\?.+&v=|shorts\/))([\w-]{11})/i);
    if (ytMatch && ytMatch[1]) {
        return {
            type: 'youtube',
            embedUrl: `https://www.youtube-nocookie.com/embed/${ytMatch[1]}?autoplay=0&rel=0`,
            directUrl: url
        };
    }

    // Canva: Enlaces de diseño y presentación
    if (url.includes('canva.com')) {
        let canvaEmbed = url;
        if (url.includes('/view') && !url.includes('embed')) {
            canvaEmbed = url.includes('?') ? `${url}&embed` : `${url}?embed`;
        }
        return {
            type: 'canva',
            embedUrl: canvaEmbed,
            directUrl: url
        };
    }

    // Imágenes directas
    if (/\.(jpeg|jpg|gif|png|webp|svg)(\?.*)?$/i.test(url) || url.startsWith('data:image/')) {
        return {
            type: 'image',
            embedUrl: url,
            directUrl: url
        };
    }

    // Enlace genérico web
    return {
        type: 'web',
        embedUrl: url,
        directUrl: url
    };
};

// Actualización de Restricciones por Tipo de Plan (Institucional vs Docente Individual vs Prueba)
export const updatePlanRestrictionsUI = () => {
    const plan = institucionData.tipoPlan || (currentAccessMode === 'modo_prueba' ? 'prueba' : 'institucional');
    const headerPlanBadge = document.getElementById('header-plan-badge');
    const headerPlanText = document.getElementById('header-plan-text');
    const quickBarPlanName = document.getElementById('quick-bar-plan-name');
    const quickBarPlanLimit = document.getElementById('quick-bar-plan-limit');
    const quickBarTitular = document.getElementById('quick-bar-titular');
    const quickBarDaysLeft = document.getElementById('quick-bar-days-left');
    const overlayStaffLocked = document.getElementById('overlay-staff-locked');
    const labelUploadStaff = document.getElementById('label-upload-staff');
    const excelStaffInput = document.getElementById('excel-staff') as HTMLInputElement;

    // Titular display
    if (quickBarTitular) {
        quickBarTitular.innerText = institucionData.titularNombre || institucionData.nombre || 'I.E. Ramón Múnera Lopera';
    }

    // Days remaining computation
    const today = getTodayString();
    const end = institucionData.licenciaFin || '2099-12-31';
    let daysDiff = Math.ceil((new Date(end).getTime() - new Date(today).getTime()) / (1000 * 60 * 60 * 24));
    if (quickBarDaysLeft) {
        if (daysDiff < 0) {
            quickBarDaysLeft.setAttribute('style', 'background-color: #7f1d1d !important; color: #fecaca !important; border: 1px solid #ef4444 !important; font-weight: bold;');
            quickBarDaysLeft.innerText = "Vencida";
        } else if (daysDiff <= 15) {
            quickBarDaysLeft.setAttribute('style', 'background-color: #78350f !important; color: #fef08a !important; border: 1px solid #f59e0b !important; font-weight: bold;');
            quickBarDaysLeft.innerText = `${daysDiff} días rest.`;
        } else {
            quickBarDaysLeft.setAttribute('style', 'background-color: #064e3b !important; color: #6ee7b7 !important; border: 1px solid #10b981 !important; font-weight: bold;');
            quickBarDaysLeft.innerText = daysDiff > 9000 ? "Permanente" : `${daysDiff} días`;
        }
    }

    if (plan === 'docente') {
        // Plan Docente Individual (1 persona)
        if (headerPlanText) headerPlanText.innerText = 'Plan Docente (1 Usuario)';
        if (headerPlanBadge) {
            headerPlanBadge.setAttribute('style', 'background-color: #581c87; color: #f3e8ff; border: 1px solid #c084fc;');
        }
        if (quickBarPlanName) {
            quickBarPlanName.setAttribute('style', 'color: #c084fc !important; font-weight: 900;');
            quickBarPlanName.innerText = "Docente Individual";
        }
        if (quickBarPlanLimit) {
            quickBarPlanLimit.setAttribute('style', 'background-color: #581c87 !important; color: #f3e8ff !important; border: 1px solid #a855f7 !important; font-weight: bold;');
            quickBarPlanLimit.innerText = "1 Persona";
        }

        // Restringir cargue de personal
        if (overlayStaffLocked) {
            overlayStaffLocked.classList.remove('hidden');
            overlayStaffLocked.classList.add('flex');
        }
        if (labelUploadStaff) {
            labelUploadStaff.classList.add('opacity-40', 'cursor-not-allowed');
        }
        if (excelStaffInput) {
            excelStaffInput.disabled = true;
        }
    } else if (plan === 'prueba') {
        // Modo Prueba 30 Días
        if (headerPlanText) headerPlanText.innerText = 'Prueba 30 Días';
        if (headerPlanBadge) {
            headerPlanBadge.setAttribute('style', 'background-color: #78350f; color: #fef3c7; border: 1px solid #f59e0b;');
        }
        if (quickBarPlanName) {
            quickBarPlanName.setAttribute('style', 'color: #fbbf24 !important; font-weight: 900;');
            quickBarPlanName.innerText = "Prueba 30 Días";
        }
        if (quickBarPlanLimit) {
            quickBarPlanLimit.setAttribute('style', 'background-color: #78350f !important; color: #fef08a !important; border: 1px solid #d97706 !important; font-weight: bold;');
            quickBarPlanLimit.innerText = "Demo";
        }
        if (overlayStaffLocked) {
            overlayStaffLocked.classList.add('hidden');
            overlayStaffLocked.classList.remove('flex');
        }
        if (labelUploadStaff) {
            labelUploadStaff.classList.remove('opacity-40', 'cursor-not-allowed');
        }
        if (excelStaffInput) {
            excelStaffInput.disabled = false;
        }
    } else {
        // Plan Institucional Completo
        if (headerPlanText) headerPlanText.innerText = 'Plan Institucional';
        if (headerPlanBadge) {
            headerPlanBadge.setAttribute('style', 'background-color: #1e3a8a; color: #dbeafe; border: 1px solid #60a5fa;');
        }
        if (quickBarPlanName) {
            quickBarPlanName.setAttribute('style', 'color: #fbbf24 !important; font-weight: 900;');
            quickBarPlanName.innerText = "Institucional";
        }
        if (quickBarPlanLimit) {
            quickBarPlanLimit.setAttribute('style', 'background-color: #1e3a8a !important; color: #bfdbfe !important; border: 1px solid #3b82f6 !important; font-weight: bold;');
            quickBarPlanLimit.innerText = "Multi-Usuario";
        }
        if (overlayStaffLocked) {
            overlayStaffLocked.classList.add('hidden');
            overlayStaffLocked.classList.remove('flex');
        }
        if (labelUploadStaff) {
            labelUploadStaff.classList.remove('opacity-40', 'cursor-not-allowed');
        }
        if (excelStaffInput) {
            excelStaffInput.disabled = false;
        }
    }
};

// Verificación de Vigencia de Licencia por Fecha Contratada y Modo Prueba 30 Días
export const checkLicenseValidity = (): boolean => {
    updatePlanRestrictionsUI();
    const today = getTodayString();
    const statusBadge = document.getElementById('license-status-badge');
    const modalExp = document.getElementById('modal-license-expired');
    const modalTrialExp = document.getElementById('modal-trial-expired');

    // MODO PRUEBA 30 DÍAS
    if (currentAccessMode === 'modo_prueba') {
        if (modalExp) modalExp.classList.add('hidden');
        const trialStatus = getTrialStatus();
        const trialInstNameEl = document.getElementById('trial-expired-inst-name');
        if (trialInstNameEl) trialInstNameEl.innerText = institucionData.nombre;

        if (trialStatus.isExpired) {
            if (modalTrialExp) modalTrialExp.classList.remove('hidden');
            if (statusBadge) {
                statusBadge.className = "text-xs font-bold px-2.5 py-1 rounded-full bg-red-100 text-red-800";
                statusBadge.innerText = "Prueba 30 Días Vencida";
            }
            isLicenseValid = false;
            return false;
        } else {
            if (modalTrialExp) modalTrialExp.classList.add('hidden');
            if (statusBadge) {
                statusBadge.className = "text-xs font-bold px-2.5 py-1 rounded-full bg-amber-100 text-amber-800";
                statusBadge.innerText = `Prueba Activa (${trialStatus.remainingDays} días)`;
            }
            isLicenseValid = true;
            return true;
        }
    }

    // MODO INSTITUCIONAL OFICIAL (RAMÓN MÚNERA)
    if (modalTrialExp) modalTrialExp.classList.add('hidden');

    const start = institucionData.licenciaInicio || '2000-01-01';
    const end = institucionData.licenciaFin || '2099-12-31';
    let valid = true;

    if (today < start) {
        valid = false;
        if (statusBadge) {
            statusBadge.className = "text-xs font-bold px-2.5 py-1 rounded-full bg-yellow-100 text-yellow-800";
            statusBadge.innerText = `Inicia el ${start}`;
        }
    } else if (today > end) {
        valid = false;
        if (statusBadge) {
            statusBadge.className = "text-xs font-bold px-2.5 py-1 rounded-full bg-red-100 text-red-800";
            statusBadge.innerText = `Licencia Vencida (${end})`;
        }
    } else {
        const diffDays = Math.ceil((new Date(end).getTime() - new Date(today).getTime()) / (1000 * 60 * 60 * 24));
        if (statusBadge) {
            if (diffDays <= 15) {
                statusBadge.className = "text-xs font-bold px-2.5 py-1 rounded-full bg-yellow-100 text-yellow-800";
                statusBadge.innerText = `Por vencer (${diffDays} días)`;
            } else {
                statusBadge.className = "text-xs font-bold px-2.5 py-1 rounded-full bg-green-100 text-green-800";
                statusBadge.innerText = `Licencia Activa (${diffDays} días)`;
            }
        }
    }

    isLicenseValid = valid;

    // Si la licencia no es válida, desplegar pantalla de bloqueo
    if (!valid) {
        if (modalExp) {
            modalExp.classList.remove('hidden');
            const instNameEl = document.getElementById('license-inst-name');
            if (instNameEl) instNameEl.innerText = institucionData.nombre;
            const limitEl = document.getElementById('license-date-limit');
            if (limitEl) limitEl.innerText = `${start} a ${end}`;
            const detailsEl = document.getElementById('license-details-text');
            if (detailsEl) {
                detailsEl.innerText = today > end 
                    ? `El periodo de uso contratado para "${institucionData.nombre}" finalizó el ${end}. Por favor contacte al proveedor para renovar el servicio.` 
                    : `El periodo de uso contratado para "${institucionData.nombre}" está programado para iniciar el ${start}.`;
            }
        }
    } else {
        if (modalExp) modalExp.classList.add('hidden');
    }

    return valid;
};

// Despliegue del Pop-up Vertical de Inicio (Optimizado para Celular)
export const showStartupPopup = (force = false) => {
    if (!institucionData.popupActivo && !force) return;
    
    // Si no es forzado (ej: desde botón de vista previa), comprobar si ya fue descartado hoy
    if (!force) {
        const dismissedDate = localStorage.getItem(`popup_dismissed_${institucionData.id}`);
        if (dismissedDate === getTodayString()) return;
    }

    const modal = document.getElementById('modal-startup-popup');
    if (!modal) return;

    // Configuración visual y textos del Pop-up
    const titleEl = document.getElementById('popup-title-display');
    if (titleEl) titleEl.innerText = institucionData.popupTitulo || 'Novedades y Tutorial Institucional';

    const instEl = document.getElementById('popup-inst-display');
    if (instEl) instEl.innerText = institucionData.nombre;

    const footerInstEl = document.getElementById('popup-footer-inst-name');
    if (footerInstEl) footerInstEl.innerText = institucionData.nombre;

    const descEl = document.getElementById('popup-description-display');
    if (descEl) descEl.innerText = institucionData.popupDescripcion || 'Bienvenido al sistema institucional. Consulta el material explicativo preparado para tu institución.';

    // Header color branding
    const modalHeader = document.getElementById('popup-modal-header');
    if (modalHeader && institucionData.color1) {
        modalHeader.style.backgroundColor = institucionData.color1;
    }

    // Logo
    const pLogo = document.getElementById('popup-header-logo') as HTMLImageElement;
    const pLogoBox = document.getElementById('popup-header-logo-container');
    if (institucionData.logo) {
        if (pLogo) pLogo.src = institucionData.logo;
        pLogoBox?.classList.remove('hidden');
    } else {
        pLogoBox?.classList.add('hidden');
    }

    // Procesar Media (YouTube, Canva, Imagen o Enlace)
    const iframe = document.getElementById('popup-iframe') as HTMLIFrameElement;
    const img = document.getElementById('popup-image') as HTMLImageElement;
    const placeholder = document.getElementById('popup-link-placeholder');
    const directLink = document.getElementById('popup-direct-link') as HTMLAnchorElement;
    const extActionBtn = document.getElementById('popup-external-action-btn') as HTMLAnchorElement;

    // Resetear estados
    if (iframe) iframe.src = '';
    if (img) { img.src = ''; img.classList.add('hidden'); }
    placeholder?.classList.add('hidden');
    placeholder?.classList.remove('flex');

    if (institucionData.popupUrl) {
        const media = parseMediaUrl(institucionData.popupUrl);

        if (media.type === 'youtube' || media.type === 'canva') {
            if (iframe) {
                iframe.src = media.embedUrl;
                iframe.classList.remove('hidden');
            }
        } else if (media.type === 'image') {
            if (iframe) iframe.classList.add('hidden');
            if (img) {
                img.src = media.embedUrl;
                img.classList.remove('hidden');
            }
        } else {
            // Enlace Web general
            if (iframe) {
                iframe.src = media.embedUrl;
                iframe.classList.remove('hidden');
            }
            if (placeholder && directLink) {
                const linkText = document.getElementById('popup-link-text');
                if (linkText) linkText.innerText = media.directUrl;
                directLink.href = media.directUrl;
            }
        }

        if (extActionBtn) {
            extActionBtn.href = media.directUrl;
            extActionBtn.classList.remove('hidden');
        }
    } else {
        if (iframe) iframe.classList.add('hidden');
        if (placeholder) {
            placeholder.classList.remove('hidden');
            placeholder.classList.add('flex');
            const linkText = document.getElementById('popup-link-text');
            if (linkText) linkText.innerText = "Configura una URL en el panel de administración.";
            if (directLink) directLink.classList.add('hidden');
        }
        if (extActionBtn) extActionBtn.classList.add('hidden');
    }

    modal.classList.remove('hidden');
};

const closeStartupPopup = (dismissForToday = false) => {
    const modal = document.getElementById('modal-startup-popup');
    if (modal) modal.classList.add('hidden');
    
    // Detener reproducción de video al cerrar
    const iframe = document.getElementById('popup-iframe') as HTMLIFrameElement;
    if (iframe) iframe.src = '';

    if (dismissForToday) {
        localStorage.setItem(`popup_dismissed_${institucionData.id}`, getTodayString());
        showToast('Preferencia guardada para hoy', 'info', 2000);
    }
};

// Event Listeners del Pop-up
document.getElementById('btn-close-startup-popup')?.addEventListener('click', () => closeStartupPopup(false));
document.getElementById('btn-popup-dismiss')?.addEventListener('click', () => closeStartupPopup(false));
document.getElementById('btn-popup-dismiss-today')?.addEventListener('click', () => closeStartupPopup(true));

// Botón para abrir el Pop-Up con la Guía, Video y Credenciales completas desde el acceso móvil
document.getElementById('btn-open-popup-info')?.addEventListener('click', () => {
    showStartupPopup(true);
});

// Acción dentro del Pop-Up para autocompletar credenciales demo 1234
document.getElementById('btn-popup-autofill-demo')?.addEventListener('click', () => {
    const sId = document.getElementById('staff-id') as HTMLInputElement;
    const sPin = document.getElementById('staff-pin') as HTMLInputElement;
    if (sId) sId.value = '1234';
    if (sPin) sPin.value = '1234';
    switchAccessTab('staff');
    closeStartupPopup(false);
    showToast('Credenciales Demo 1234 cargadas. Haz clic en "Iniciar Turno".', 'info', 3500);
});

// Acción dentro del Pop-Up para ir directo a la pestaña de carnet de estudiante
document.getElementById('btn-popup-goto-student')?.addEventListener('click', () => {
    closeStartupPopup(false);
    switchAccessTab('student');
});

// Acceso Administrador desde pantalla de licencia vencida
document.getElementById('btn-license-admin-access')?.addEventListener('click', () => {
    const modalExp = document.getElementById('modal-license-expired');
    if (modalExp) modalExp.classList.add('hidden');
    switchTab('admin');
    const adminPinInput = document.getElementById('admin-pin') as HTMLInputElement;
    if (adminPinInput) {
        adminPinInput.value = '';
        setTimeout(() => adminPinInput.focus(), 150);
    }
});

// Actualizar Selector de Instituciones en Admin
export const renderInstitutionsDropdown = () => {
    const select = document.getElementById('inst-select') as HTMLSelectElement;
    if (!select) return;
    
    select.innerHTML = '';
    institucionesList.forEach(inst => {
        const opt = document.createElement('option');
        opt.value = inst.id;
        opt.innerText = inst.nombre;
        if (inst.id === institucionData.id) opt.selected = true;
        select.appendChild(opt);
    });
};

// Cargar Datos de la Institución Seleccionada en el Formulario
export const populateAdminInstitutionForm = (inst: InstitutionProfile) => {
    const iInput = document.getElementById('inst-name-input') as HTMLInputElement;
    const c1 = document.getElementById('inst-color1') as HTMLInputElement;
    const c2 = document.getElementById('inst-color2') as HTMLInputElement;
    const lStart = document.getElementById('inst-licencia-inicio') as HTMLInputElement;
    const lEnd = document.getElementById('inst-licencia-fin') as HTMLInputElement;
    const pActive = document.getElementById('inst-popup-active') as HTMLInputElement;
    const pLabel = document.getElementById('label-popup-active');
    const pTitle = document.getElementById('inst-popup-title') as HTMLInputElement;
    const pUrl = document.getElementById('inst-popup-url') as HTMLInputElement;
    const pDesc = document.getElementById('inst-popup-desc') as HTMLTextAreaElement;
    const lPrev = document.getElementById('logo-preview') as HTMLImageElement;

    if (iInput) iInput.value = inst.nombre || '';
    if (c1) c1.value = inst.color1 || '#2563eb';
    if (c2) c2.value = inst.color2 || '#0ea5e9';
    if (lStart) lStart.value = inst.licenciaInicio || '2026-01-01';
    if (lEnd) lEnd.value = inst.licenciaFin || '2027-12-31';

    if (pActive) {
        pActive.checked = !!inst.popupActivo;
        if (pLabel) pLabel.innerText = inst.popupActivo ? 'Activado' : 'Desactivado';
    }
    if (pTitle) pTitle.value = inst.popupTitulo || '';
    if (pUrl) pUrl.value = inst.popupUrl || '';
    if (pDesc) pDesc.value = inst.popupDescripcion || '';

    if (lPrev) {
        if (inst.logo) {
            lPrev.src = inst.logo;
            lPrev.classList.remove('hidden');
        } else {
            lPrev.src = '';
            lPrev.classList.add('hidden');
        }
    }

    checkLicenseValidity();
};

const aplicarConfiguracionUI = () => {
    const hName = document.getElementById('header-inst-name');
    if (hName && institucionData.nombre) hName.innerText = institucionData.nombre;
    
    const hLogo = document.getElementById('header-logo') as HTMLImageElement;
    const hLogoBox = document.getElementById('header-logo-container');
    const hDefaultIcon = document.getElementById('header-default-icon');
    
    if (institucionData.logo) { 
        if (hLogo) hLogo.src = institucionData.logo; 
        hLogoBox?.classList.remove('hidden'); 
        hDefaultIcon?.classList.add('hidden'); 
    } else {
        hLogoBox?.classList.add('hidden');
        hDefaultIcon?.classList.remove('hidden');
    }

    const mHead = document.getElementById('main-header');
    if (institucionData.color1) { 
        if (mHead) mHead.style.backgroundColor = institucionData.color1; 
        document.getElementById('wave-3')?.setAttribute('fill', institucionData.color1); 
    }
    if (institucionData.color2) { 
        document.getElementById('wave-1')?.setAttribute('fill', institucionData.color2); 
        document.getElementById('wave-2')?.setAttribute('fill', institucionData.color2); 
    }

    renderInstitutionsDropdown();
    populateAdminInstitutionForm(institucionData);
    updateModeUI();
};

// Cambio en Selector de Institución
document.getElementById('inst-select')?.addEventListener('change', (e: any) => {
    const selectedId = e.target.value;
    const found = institucionesList.find(i => i.id === selectedId);
    if (found) {
        institucionData = { ...found };
        aplicarConfiguracionUI();
        showToast(`Institución cambiada: ${institucionData.nombre}`, 'info', 2000);
    }
});

// Agregar Nueva Institución
document.getElementById('btn-new-inst')?.addEventListener('click', () => {
    const nombre = prompt("Nombre de la Nueva Institución Educativa:");
    if (!nombre || !nombre.trim()) return;

    const cleanId = nombre.toLowerCase().replace(/[^a-z0-9]/g, '_').slice(0, 20) + '_' + Date.now().toString().slice(-4);
    const newInst: InstitutionProfile = {
        id: cleanId,
        nombre: nombre.trim(),
        logo: null,
        color1: "#2563eb",
        color2: "#0ea5e9",
        licenciaInicio: getTodayString(),
        licenciaFin: new Date(new Date().setFullYear(new Date().getFullYear() + 1)).toISOString().split('T')[0],
        tipoPlan: 'institucional',
        limiteUsuarios: 9999,
        titularNombre: nombre.trim(),
        popupActivo: false,
        popupTitulo: `Bienvenido a ${nombre.trim()}`,
        popupUrl: "",
        popupDescripcion: "Información y tutorial de la plataforma."
    };

    institucionesList.push(newInst);
    institucionData = { ...newInst };
    aplicarConfiguracionUI();
    showToast(`Institución "${nombre}" agregada. Ajusta sus datos y guarda.`, 'success', 4000);
});

// Eliminar Institución Actual
document.getElementById('btn-del-inst')?.addEventListener('click', () => {
    if (institucionesList.length <= 1) {
        return showToast('Debe existir al menos una institución en el sistema.', 'warning');
    }
    if (confirm(`¿Eliminar la institución "${institucionData.nombre}" de la lista?`)) {
        institucionesList = institucionesList.filter(i => i.id !== institucionData.id);
        institucionData = { ...institucionesList[0] };
        aplicarConfiguracionUI();
        showToast('Institución eliminada.', 'info');
    }
});

// Toggle del Switch de Pop-up
document.getElementById('inst-popup-active')?.addEventListener('change', (e: any) => {
    const pLabel = document.getElementById('label-popup-active');
    if (pLabel) pLabel.innerText = e.target.checked ? 'Activado' : 'Desactivado';
});

// Actualización en vivo del badge de vigencia al cambiar fechas
document.getElementById('inst-licencia-inicio')?.addEventListener('input', (e: any) => {
    institucionData.licenciaInicio = e.target.value;
    checkLicenseValidity();
});
document.getElementById('inst-licencia-fin')?.addEventListener('input', (e: any) => {
    institucionData.licenciaFin = e.target.value;
    checkLicenseValidity();
});

// Vista Previa Inmediata del Pop-up desde el Panel de Configuración
document.getElementById('btn-preview-popup')?.addEventListener('click', () => {
    const pTitle = (document.getElementById('inst-popup-title') as HTMLInputElement)?.value;
    const pUrl = (document.getElementById('inst-popup-url') as HTMLInputElement)?.value;
    const pDesc = (document.getElementById('inst-popup-desc') as HTMLTextAreaElement)?.value;
    const pActive = (document.getElementById('inst-popup-active') as HTMLInputElement)?.checked;

    // Temporal para previsualizar
    institucionData.popupTitulo = pTitle;
    institucionData.popupUrl = pUrl;
    institucionData.popupDescripcion = pDesc;
    institucionData.popupActivo = pActive;

    showStartupPopup(true);
});

// Carga de Logo Institucional
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
                if (lPrev) { 
                    lPrev.src = institucionData.logo; 
                    lPrev.classList.remove('hidden'); 
                }
            }; 
            img.src = event.target.result; 
        }; 
        reader.readAsDataURL(file); 
    }
});

// Guardar y Aplicar a la Institución
document.getElementById('btn-save-settings')?.addEventListener('click', async () => {
    const btn = document.getElementById('btn-save-settings') as HTMLButtonElement; 
    btn.innerHTML = '<i class="fas fa-spinner fa-spin mr-2"></i>Guardando...'; 
    btn.disabled = true;

    const iInput = document.getElementById('inst-name-input') as HTMLInputElement;
    const c1 = document.getElementById('inst-color1') as HTMLInputElement;
    const c2 = document.getElementById('inst-color2') as HTMLInputElement;
    const lStart = document.getElementById('inst-licencia-inicio') as HTMLInputElement;
    const lEnd = document.getElementById('inst-licencia-fin') as HTMLInputElement;
    const pActive = document.getElementById('inst-popup-active') as HTMLInputElement;
    const pTitle = document.getElementById('inst-popup-title') as HTMLInputElement;
    const pUrl = document.getElementById('inst-popup-url') as HTMLInputElement;
    const pDesc = document.getElementById('inst-popup-desc') as HTMLTextAreaElement;

    // En Modo Prueba (30 Días), guardar en la partición aislada de prueba
    if (currentAccessMode === 'modo_prueba') {
        institucionData.nombre = iInput?.value.trim() || 'Institución Educativa (Modo Prueba)'; 
        institucionData.color1 = c1?.value || '#4f46e5'; 
        institucionData.color2 = c2?.value || '#06b6d4';
        institucionData.licenciaInicio = lStart?.value || getTodayString();
        institucionData.licenciaFin = lEnd?.value || getTodayString();
        institucionData.popupActivo = !!pActive?.checked;
        institucionData.popupTitulo = pTitle?.value.trim() || '';
        institucionData.popupUrl = pUrl?.value.trim() || '';
        institucionData.popupDescripcion = pDesc?.value.trim() || '';

        localStorage.setItem('trial_institution_profile', JSON.stringify(institucionData));
        try {
            await setDoc(getSettingsPath(), institucionData);
        } catch(e) {}

        aplicarConfiguracionUI();
        showToast('Configuración y Pop-up de tu colegio en prueba guardados con éxito.', 'success');
        btn.innerHTML = '<i class="fas fa-save mr-2"></i>Guardar y Aplicar a la Institución'; 
        btn.disabled = false;
        return;
    }

    // Actualizar datos del objeto activo (Modo Oficial)
    institucionData.nombre = iInput?.value.trim() || 'Gestor Académico'; 
    institucionData.color1 = c1?.value || '#2563eb'; 
    institucionData.color2 = c2?.value || '#0ea5e9';
    institucionData.licenciaInicio = lStart?.value || '2026-01-01';
    institucionData.licenciaFin = lEnd?.value || '2027-12-31';
    institucionData.popupActivo = !!pActive?.checked;
    institucionData.popupTitulo = pTitle?.value.trim() || '';
    institucionData.popupUrl = pUrl?.value.trim() || '';
    institucionData.popupDescripcion = pDesc?.value.trim() || '';

    // Actualizar en la lista de instituciones
    const idx = institucionesList.findIndex(i => i.id === institucionData.id);
    if (idx !== -1) {
        institucionesList[idx] = { ...institucionData };
    } else {
        institucionesList.push({ ...institucionData });
    }

    const payloadToSave = {
        institucionActivaId: institucionData.id,
        institucionActiva: institucionData,
        instituciones: institucionesList
    };

    try { 
        await setDoc(getSettingsPath(), payloadToSave); 
        localStorage.setItem('institucionData', JSON.stringify(payloadToSave)); 
        aplicarConfiguracionUI(); 
        showToast('Configuración institucional y vigencia guardadas con éxito.', 'success'); 
    } catch (error) { 
        localStorage.setItem('institucionData', JSON.stringify(payloadToSave)); 
        aplicarConfiguracionUI(); 
        showToast('Guardado localmente.', 'warning'); 
    } finally { 
        btn.innerHTML = '<i class="fas fa-save mr-2"></i>Guardar y Aplicar a la Institución'; 
        btn.disabled = false; 
    }
});

// Event Listeners de Archivos Excel
document.getElementById('excel-students')?.addEventListener('change', (e: any) => { 
    if(e.target.files[0]) processExcelUpload(e.target.files[0], 'students'); 
    e.target.value=''; 
});
document.getElementById('excel-staff')?.addEventListener('change', (e: any) => { 
    if (institucionData.tipoPlan === 'docente') {
        showPlanRestrictionModal();
        e.target.value = '';
        return;
    }
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

// Modal de Restricción de Plan Docente Individual
export const showPlanRestrictionModal = () => {
    const modal = document.getElementById('modal-plan-restriction');
    if (modal) modal.classList.remove('hidden');
};

document.getElementById('btn-close-plan-restriction')?.addEventListener('click', () => {
    document.getElementById('modal-plan-restriction')?.classList.add('hidden');
});

document.getElementById('overlay-staff-locked')?.addEventListener('click', () => {
    showPlanRestrictionModal();
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

// =========================================================================
// MÓDULO SUPER-ADMINISTRADOR (Ventas, Licencias, Fechas y Roles) - www.espatodo.com
// =========================================================================
let superAdminPin = localStorage.getItem('sa_master_pin') || 'superadmin';
let isSuperAdminAuthenticated = false;

export const openSuperAdminLogin = () => {
    const modalLogin = document.getElementById('modal-superadmin-login');
    const pinInput = document.getElementById('sa-login-pin') as HTMLInputElement;
    if (modalLogin) modalLogin.classList.remove('hidden');
    if (pinInput) {
        pinInput.value = '';
        setTimeout(() => pinInput.focus(), 120);
    }
};

export const closeSuperAdminLogin = () => {
    document.getElementById('modal-superadmin-login')?.classList.add('hidden');
};

export const openSuperAdminModal = () => {
    closeSuperAdminLogin();
    const modal = document.getElementById('modal-superadmin');
    if (!modal) return;
    
    // Asignar inputs con los datos de la institución activa
    const tNombre = document.getElementById('sa-titular-nombre') as HTMLInputElement;
    const tDoc = document.getElementById('sa-titular-doc') as HTMLInputElement;
    const tContacto = document.getElementById('sa-titular-contacto') as HTMLInputElement;
    const lStart = document.getElementById('sa-licencia-inicio') as HTMLInputElement;
    const lEnd = document.getElementById('sa-licencia-fin') as HTMLInputElement;
    const mPin = document.getElementById('sa-master-pin') as HTMLInputElement;

    if (tNombre) tNombre.value = institucionData.titularNombre || institucionData.nombre || '';
    if (tDoc) tDoc.value = institucionData.titularDoc || '';
    if (tContacto) tContacto.value = institucionData.titularContacto || '';
    if (lStart) lStart.value = institucionData.licenciaInicio || '2026-01-01';
    if (lEnd) lEnd.value = institucionData.licenciaFin || '2027-12-31';
    if (mPin) mPin.value = '';

    // Seleccionar plan activo
    const activePlan = institucionData.tipoPlan || (currentAccessMode === 'modo_prueba' ? 'prueba' : 'institucional');
    selectSuperAdminPlan(activePlan);

    // Actualizar contador
    updateSuperAdminCountdown();

    // Renderizar matriz de roles
    renderSuperAdminRolesMatrix();

    modal.classList.remove('hidden');
};

export const closeSuperAdminModal = () => {
    document.getElementById('modal-superadmin')?.classList.add('hidden');
};

export const selectSuperAdminPlan = (plan: LicenseType) => {
    const cards = document.querySelectorAll('.sa-plan-card');
    cards.forEach((c: any) => {
        const radio = c.querySelector('input[name="sa_plan_choice"]');
        if (radio && radio.value === plan) {
            radio.checked = true;
            c.className = "sa-plan-card relative flex flex-col p-4 rounded-2xl border-2 cursor-pointer transition-all border-amber-500 bg-amber-500/15 shadow-md shadow-amber-500/10";
        } else {
            if (radio) radio.checked = false;
            c.className = "sa-plan-card relative flex flex-col p-4 rounded-2xl border-2 cursor-pointer transition-all border-gray-700 bg-gray-800/50 hover:border-gray-600";
        }
    });

    const pill = document.getElementById('sa-active-plan-pill');
    if (pill) {
        if (plan === 'docente') {
            pill.className = "text-[10px] font-black px-2.5 py-1 rounded-full bg-purple-500/20 text-purple-300 border border-purple-500/40 uppercase";
            pill.innerText = "Plan Docente Individual (1 Usuario)";
        } else if (plan === 'prueba') {
            pill.className = "text-[10px] font-black px-2.5 py-1 rounded-full bg-yellow-500/20 text-yellow-300 border border-yellow-500/40 uppercase";
            pill.innerText = "Modo Prueba 30 Días";
        } else {
            pill.className = "text-[10px] font-black px-2.5 py-1 rounded-full bg-blue-500/20 text-blue-300 border border-blue-500/40 uppercase";
            pill.innerText = "Plan Institucional Completo";
        }
    }
};

export const updateSuperAdminCountdown = () => {
    const lEnd = (document.getElementById('sa-licencia-fin') as HTMLInputElement)?.value;
    const lStart = (document.getElementById('sa-licencia-inicio') as HTMLInputElement)?.value;
    const today = getTodayString();
    const countdownEl = document.getElementById('sa-license-countdown');
    if (!countdownEl || !lEnd) return;

    if (today < lStart) {
        countdownEl.className = "text-xs font-black px-3 py-1 rounded-full bg-yellow-500/20 text-yellow-300 border border-yellow-500/40";
        countdownEl.innerText = `Programada para iniciar: ${lStart}`;
    } else if (today > lEnd) {
        countdownEl.className = "text-xs font-black px-3 py-1 rounded-full bg-red-500/20 text-red-300 border border-red-500/40";
        countdownEl.innerText = `Vencida (${lEnd})`;
    } else {
        const diffDays = Math.ceil((new Date(lEnd).getTime() - new Date(today).getTime()) / (1000 * 60 * 60 * 24));
        if (diffDays > 9000) {
            countdownEl.className = "text-xs font-black px-3 py-1 rounded-full bg-emerald-500/20 text-emerald-300 border border-emerald-500/40";
            countdownEl.innerText = "Acceso Permanente (2099)";
        } else {
            countdownEl.className = "text-xs font-black px-3 py-1 rounded-full bg-emerald-500/20 text-emerald-300 border border-emerald-500/40";
            countdownEl.innerText = `Licencia Activa (${diffDays} días restantes)`;
        }
    }
};

export const renderSuperAdminRolesMatrix = () => {
    const container = document.getElementById('sa-roles-matrix-container');
    if (!container) return;
    container.innerHTML = '';

    const roleDefs = [
        { key: 'administrador', name: 'Administrador / Rectoría', icon: 'fa-shield-alt', desc: 'Control total de la institución y configuración' },
        { key: 'coordinador', name: 'Coordinación Académica', icon: 'fa-user-graduate', desc: 'Asistencia general, observación y autorizaciones' },
        { key: 'docente', name: 'Docente Titular', icon: 'fa-chalkboard-teacher', desc: 'Asistencia de aula, planilla de notas y evaluaciones' },
        { key: 'vigilante', name: 'Seguridad / Portería', icon: 'fa-door-open', desc: 'Control de ingresos y salidas en puerta' }
    ];

    roleDefs.forEach(r => {
        let html = `
        <div class="p-3.5 bg-gray-900/90 rounded-xl border border-gray-700/80">
            <div class="flex items-center justify-between mb-2">
                <div class="flex items-center gap-2">
                    <i class="fas ${r.icon} text-amber-400"></i>
                    <h4 class="font-black text-xs text-white uppercase tracking-wider">${r.name}</h4>
                </div>
                <span class="text-[10px] text-gray-400 font-medium">${r.desc}</span>
            </div>
            <div class="grid grid-cols-2 sm:grid-cols-4 gap-2 pt-2 border-t border-gray-800">
        `;
        modulesList.forEach(m => {
            const isChecked = rolePermissions[r.key]?.includes(m.id) ? 'checked' : '';
            html += `
            <label class="flex items-center gap-1.5 text-[11px] text-gray-300 hover:text-white cursor-pointer select-none">
                <input type="checkbox" class="sa-role-cb rounded bg-gray-800 border-gray-600 text-amber-500 focus:ring-0" data-role="${r.key}" data-module="${m.id}" ${isChecked}>
                <span>${m.name}</span>
            </label>
            `;
        });
        html += `</div></div>`;
        container.innerHTML += html;
    });
};

export const saveSuperAdminConfig = async () => {
    const btnSave = document.getElementById('btn-sa-save') as HTMLButtonElement;
    if (btnSave) {
        btnSave.disabled = true;
        btnSave.innerHTML = '<i class="fas fa-spinner fa-spin mr-1"></i> Guardando...';
    }

    try {
        const planRadio = document.querySelector('input[name="sa_plan_choice"]:checked') as HTMLInputElement;
        const chosenPlan = (planRadio?.value || 'institucional') as LicenseType;
        const tNombre = (document.getElementById('sa-titular-nombre') as HTMLInputElement)?.value.trim();
        const tDoc = (document.getElementById('sa-titular-doc') as HTMLInputElement)?.value.trim();
        const tContacto = (document.getElementById('sa-titular-contacto') as HTMLInputElement)?.value.trim();
        const lStart = (document.getElementById('sa-licencia-inicio') as HTMLInputElement)?.value || '2026-01-01';
        const lEnd = (document.getElementById('sa-licencia-fin') as HTMLInputElement)?.value || '2027-12-31';

        // Actualizar institucionData
        institucionData.tipoPlan = chosenPlan;
        institucionData.limiteUsuarios = chosenPlan === 'docente' ? 1 : 9999;
        if (tNombre) {
            institucionData.titularNombre = tNombre;
            institucionData.nombre = tNombre;
        }
        if (tDoc) institucionData.titularDoc = tDoc;
        if (tContacto) institucionData.titularContacto = tContacto;
        institucionData.licenciaInicio = lStart;
        institucionData.licenciaFin = lEnd;

        // Actualizar permisos de roles desde la matriz
        const newPerms: Record<string, string[]> = { docente: [], vigilante: [], coordinador: [], administrador: [] };
        document.querySelectorAll('.sa-role-cb').forEach((cb: any) => {
            if (cb.checked) {
                const r = cb.dataset.role;
                const m = cb.dataset.module;
                if (!newPerms[r]) newPerms[r] = [];
                newPerms[r].push(m);
            }
        });
        rolePermissions = newPerms;

        // Guardar en Firestore
        try {
            await setDoc(getSettingsPath(), { ...institucionData });
            await setDoc(doc(db, 'configuracion_app', 'superadmin_config'), {
                tipoPlan: chosenPlan,
                titularNombre: tNombre,
                titularDoc: tDoc,
                titularContacto: tContacto,
                licenciaInicio: lStart,
                licenciaFin: lEnd,
                rolePermissions,
                updatedAt: new Date().toISOString()
            });
            await setDoc(getPermissionsPath(), rolePermissions);
        } catch(e) {}

        // Guardar en localStorage
        localStorage.setItem('institucionData', JSON.stringify({ institucionActiva: institucionData }));
        localStorage.setItem('rolePermissions', JSON.stringify(rolePermissions));

        // Refrescar UI
        updatePlanRestrictionsUI();
        checkLicenseValidity();
        aplicarConfiguracionUI();
        renderRolesConfig();
        if (currentStaff) applyRolesUI(currentStaff);

        closeSuperAdminModal();
        showToast(`Licencia Super-Admin guardada: Plan ${chosenPlan.toUpperCase()}`, 'success', 4000);
    } catch(err) {
        showToast('Error al guardar configuración Super-Admin', 'error');
    } finally {
        if (btnSave) {
            btnSave.disabled = false;
            btnSave.innerHTML = '<i class="fas fa-save mr-1"></i> Guardar y Aplicar Licencia';
        }
    }
};

// Eventos de Super-Administrador
document.getElementById('btn-header-superadmin')?.addEventListener('click', () => {
    if (isSuperAdminAuthenticated) openSuperAdminModal();
    else openSuperAdminLogin();
});

document.getElementById('header-plan-badge')?.addEventListener('click', () => {
    if (isSuperAdminAuthenticated) openSuperAdminModal();
    else openSuperAdminLogin();
});

document.getElementById('btn-open-superadmin-login')?.addEventListener('click', () => {
    openSuperAdminLogin();
});

document.getElementById('btn-open-superadmin-modal')?.addEventListener('click', () => {
    if (isSuperAdminAuthenticated) openSuperAdminModal();
    else openSuperAdminLogin();
});

document.getElementById('btn-close-sa-login-x')?.addEventListener('click', () => closeSuperAdminLogin());
document.getElementById('btn-sa-login-cancel')?.addEventListener('click', () => closeSuperAdminLogin());

document.getElementById('btn-sa-login-submit')?.addEventListener('click', () => {
    const pin = (document.getElementById('sa-login-pin') as HTMLInputElement)?.value.trim();
    if (pin === superAdminPin || pin === 'superadmin' || pin === 'espatodo777' || pin === 'admin777') {
        isSuperAdminAuthenticated = true;
        closeSuperAdminLogin();
        openSuperAdminModal();
        showToast('Autenticado como Super-Administrador (www.espatodo.com)', 'success');
    } else {
        showToast('Clave de Super-Administrador incorrecta', 'error');
    }
});

document.getElementById('sa-login-pin')?.addEventListener('keydown', (e: KeyboardEvent) => {
    if (e.key === 'Enter') {
        document.getElementById('btn-sa-login-submit')?.click();
    }
});

document.getElementById('btn-close-superadmin')?.addEventListener('click', () => closeSuperAdminModal());
document.getElementById('btn-sa-cancel')?.addEventListener('click', () => closeSuperAdminModal());
document.getElementById('btn-sa-save')?.addEventListener('click', () => saveSuperAdminConfig());

// Botones de Extensión Rápida de Fecha
document.querySelectorAll('.sa-btn-quick-date').forEach(btn => {
    btn.addEventListener('click', (e: any) => {
        const days = parseInt(e.currentTarget.dataset.days || '30', 10);
        const lEnd = document.getElementById('sa-licencia-fin') as HTMLInputElement;
        if (!lEnd) return;

        if (days === 0) {
            // Vencer / Bloquear hoy
            const yesterday = new Date(Date.now() - 24 * 60 * 60 * 1000).toISOString().split('T')[0];
            lEnd.value = yesterday;
        } else if (days >= 90000) {
            // Acceso permanente
            lEnd.value = '2099-12-31';
        } else {
            const baseDate = new Date();
            const futureDate = new Date(baseDate.getTime() + days * 24 * 60 * 60 * 1000);
            lEnd.value = futureDate.toISOString().split('T')[0];
        }
        updateSuperAdminCountdown();
        showToast(`Fecha de vencimiento actualizada a ${lEnd.value}`, 'info', 2000);
    });
});

// Selección de Plan en Super-Admin
document.querySelectorAll('.sa-plan-card').forEach(card => {
    card.addEventListener('click', (e: any) => {
        const radio = card.querySelector('input[name="sa_plan_choice"]') as HTMLInputElement;
        if (radio) {
            selectSuperAdminPlan(radio.value as LicenseType);
        }
    });
});

document.getElementById('sa-licencia-inicio')?.addEventListener('input', () => updateSuperAdminCountdown());
document.getElementById('sa-licencia-fin')?.addEventListener('input', () => updateSuperAdminCountdown());

// Actualizar clave maestra de Super-Admin
document.getElementById('btn-save-master-pin')?.addEventListener('click', () => {
    const pin = (document.getElementById('sa-master-pin') as HTMLInputElement)?.value.trim();
    if (!pin || pin.length < 4) return showToast('La clave debe tener al menos 4 caracteres', 'warning');
    superAdminPin = pin;
    localStorage.setItem('sa_master_pin', pin);
    try {
        setDoc(doc(db, 'configuracion_app', 'superadmin_auth'), { pin, updatedAt: new Date().toISOString() });
    } catch(e) {}
    showToast('Clave maestra de Super-Administrador actualizada', 'success');
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
document.getElementById('btn-admin-back-top')?.addEventListener('click', () => switchTab('scanner'));
document.getElementById('btn-admin-logout')?.addEventListener('click', () => { 
    document.getElementById('admin-panel')?.classList.add('hidden'); 
    document.getElementById('admin-panel')?.classList.remove('flex'); 
    document.getElementById('admin-login')?.classList.remove('hidden'); 
    const pin = document.getElementById('admin-pin') as HTMLInputElement;
    if (pin) pin.value = ''; 
    switchTab('login'); 
});

document.getElementById('btn-admin-login')?.addEventListener('click', () => {
    const pin = (document.getElementById('admin-pin') as HTMLInputElement)?.value.trim();
    if (pin === superAdminPin || pin === 'superadmin' || pin === 'espatodo777' || pin === 'admin777') {
        isSuperAdminAuthenticated = true;
        document.getElementById('admin-login')?.classList.add('hidden');
        document.getElementById('admin-panel')?.classList.remove('hidden'); 
        document.getElementById('admin-panel')?.classList.add('flex');
        openSuperAdminModal();
        showToast('Bienvenido, Super-Administrador (www.espatodo.com)', 'success');
        return;
    }
    const isValid = pin === 'profe123' || (currentAccessMode === 'modo_prueba' && (pin === '1234' || pin === 'admin'));
    if (isValid) { 
        document.getElementById('admin-login')?.classList.add('hidden');
        document.getElementById('admin-panel')?.classList.remove('hidden'); 
        document.getElementById('admin-panel')?.classList.add('flex');
    } else {
        showToast('Clave de administrador incorrecta (Para demo usa 1234)', 'error');
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

    // Actualización de UI según Modo (Prueba 30 Días vs Ramón Múnera)
    const trialBanner = document.getElementById('trial-banner');
    const trialDaysLeftEl = document.getElementById('trial-days-left-text');
    const trialCredentialsBox = document.getElementById('trial-quick-credentials');
    const loginModeName = document.getElementById('login-mode-name');
    const loginModeLabel = document.getElementById('login-mode-label');
    const loginModeStatusTag = document.getElementById('login-mode-status-tag');
    const loginModeIcon = document.getElementById('login-mode-icon');
    const toggleTrialBtnText = document.getElementById('toggle-trial-mode-text');
    const loginCardSubtitle = document.getElementById('login-card-subtitle');

    if (currentAccessMode === 'modo_prueba') {
        const trialStatus = getTrialStatus();
        if (trialBanner) {
            trialBanner.classList.remove('hidden');
            if (trialDaysLeftEl) {
                trialDaysLeftEl.innerText = `Modo Prueba Universal: Le quedan ${trialStatus.remainingDays} días de 30. Adquiere en www.espatodo.com`;
            }
        }
        if (loginModeName) loginModeName.innerText = institucionData.nombre;
        if (loginModeLabel) loginModeLabel.innerText = "Entorno de Prueba Aislado";
        if (loginModeStatusTag) {
            loginModeStatusTag.className = "text-[10px] font-bold px-2 py-0.5 rounded-full bg-amber-100 text-amber-800";
            loginModeStatusTag.innerText = `Prueba (${trialStatus.remainingDays} días)`;
        }
        if (loginModeIcon) loginModeIcon.className = "fas fa-vial text-amber-600 text-lg flex-shrink-0";
        if (toggleTrialBtnText) toggleTrialBtnText.innerText = "Ir a Ramón Múnera";
        if (trialCredentialsBox) trialCredentialsBox.classList.remove('hidden');
        if (loginCardSubtitle) loginCardSubtitle.innerText = "Modo Prueba: Inicia con Doc: 1234 / Clave: 1234";

        const trialChip = document.getElementById('login-trial-status-chip');
        if (trialChip) {
            trialChip.classList.remove('hidden');
            trialChip.innerText = `Prueba (${trialStatus.remainingDays} d)`;
        }

        const badge = document.getElementById('firebase-status-badge');
        if (badge) {
            badge.className = "text-[10px] bg-amber-800 text-amber-200 px-2 py-0.5 rounded-full font-bold flex items-center gap-1";
            badge.innerHTML = '<span class="w-1.5 h-1.5 rounded-full bg-amber-400"></span> Modo Prueba (30 Días)';
        }
    } else {
        // ramon_munera
        if (trialBanner) trialBanner.classList.add('hidden');
        if (loginModeName) loginModeName.innerText = institucionData.nombre;
        if (loginModeLabel) loginModeLabel.innerText = "Institución Oficial";
        if (loginModeStatusTag) {
            loginModeStatusTag.className = "text-[10px] font-bold px-2 py-0.5 rounded-full bg-blue-100 text-blue-800";
            loginModeStatusTag.innerText = "Licencia Anual Activa";
        }
        if (loginModeIcon) loginModeIcon.className = "fas fa-university text-blue-600 text-lg flex-shrink-0";
        if (toggleTrialBtnText) toggleTrialBtnText.innerText = "Probar otra I.E.";
        if (trialCredentialsBox) trialCredentialsBox.classList.add('hidden');
        if (loginCardSubtitle) loginCardSubtitle.innerText = "Inicia turno para operar el sistema.";

        const trialChip = document.getElementById('login-trial-status-chip');
        if (trialChip) {
            trialChip.classList.add('hidden');
        }

        const badge = document.getElementById('firebase-status-badge');
        if (badge) {
            badge.className = "text-[10px] bg-blue-800 text-blue-200 px-2 py-0.5 rounded-full font-bold flex items-center gap-1";
            badge.innerHTML = '<span class="w-1.5 h-1.5 rounded-full bg-green-400"></span> BD Sincronizada';
        }
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
document.getElementById('staff-id')?.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
        e.preventDefault();
        (document.getElementById('staff-pin') as HTMLInputElement)?.focus();
    }
});

document.getElementById('staff-pin')?.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
        e.preventDefault();
        (document.getElementById('btn-staff-login') as HTMLButtonElement)?.click();
    }
});

document.getElementById('btn-staff-login')?.addEventListener('click', async () => {
    const sId = document.getElementById('staff-id') as HTMLInputElement;
    const sPin = document.getElementById('staff-pin') as HTMLInputElement;
    const id = sId?.value.trim(); 
    const pin = sPin?.value.trim();
    if (!id) return showToast('Ingrese documento', 'warning');
    
    const staff = staffDict[id];
    if (staff && pin === (staff.clave || staff.id)) {
        // Restricción para Plan Docente Individual (1 solo usuario autorizado)
        if (institucionData.tipoPlan === 'docente') {
            if (institucionData.titularDoc && institucionData.titularDoc.trim()) {
                if (id !== institucionData.titularDoc.trim()) {
                    showToast(`Licencia Docente Individual: Solo el titular autorizado (${institucionData.titularDoc}) tiene acceso. Adquiere el Plan Institucional en www.espatodo.com para habilitar a todo el personal`, 'error', 6000);
                    return;
                }
            } else {
                // Vincular al primer docente que ingresa
                institucionData.titularDoc = id;
                try {
                    await setDoc(getSettingsPath(), { ...institucionData });
                } catch(e) {}
            }
        }

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
    
    // 1. Detectar Parámetros de Acceso en la URL o Preferencia Guardada
    const urlParams = new URLSearchParams(window.location.search);
    const instParam = (urlParams.get('inst') || urlParams.get('colegio') || urlParams.get('modo') || '').toLowerCase().trim();

    if (instParam === 'rm' || instParam === 'ramon_munera' || instParam === 'oficial') {
        currentAccessMode = 'ramon_munera';
    } else if (instParam === 'prueba' || instParam === 'demo' || instParam === 'trial' || instParam === 'test') {
        currentAccessMode = 'modo_prueba';
    } else {
        const savedMode = localStorage.getItem('active_access_mode');
        if (savedMode === 'modo_prueba' || savedMode === 'ramon_munera') {
            currentAccessMode = savedMode as any;
        } else {
            // Si el visitante proviene de www.espatodo.com o enlace general abierto sin parámetro específico
            const referrer = document.referrer ? document.referrer.toLowerCase() : '';
            if (referrer.includes('espatodo.com')) {
                currentAccessMode = 'modo_prueba';
            } else {
                currentAccessMode = 'ramon_munera';
            }
        }
    }

    localStorage.setItem('active_access_mode', currentAccessMode);

    try { 
        await signInAnonymously(auth); 
    } catch (error) { 
        console.warn("Conexión inicial Firebase:", error);
    }

    await loadDatabases();
    aplicarConfiguracionUI();
    checkLicenseValidity();

    // Mostrar el Pop-up Vertical al iniciar si está habilitado
    setTimeout(() => {
        showStartupPopup(false);
    }, 600);
};

// Event Listeners para Modo Prueba y Selector Institucional
document.getElementById('btn-toggle-trial-mode')?.addEventListener('click', () => {
    const nextMode = currentAccessMode === 'ramon_munera' ? 'modo_prueba' : 'ramon_munera';
    switchAccessMode(nextMode);
});

document.getElementById('btn-header-switch-mode')?.addEventListener('click', () => {
    const nextMode = currentAccessMode === 'ramon_munera' ? 'modo_prueba' : 'ramon_munera';
    switchAccessMode(nextMode);
});

document.getElementById('btn-fill-demo-credentials')?.addEventListener('click', () => {
    const sId = document.getElementById('staff-id') as HTMLInputElement;
    const sPin = document.getElementById('staff-pin') as HTMLInputElement;
    if (sId) sId.value = '1234';
    if (sPin) sPin.value = '1234';
    showToast('Credenciales demo ingresadas: Doc: 1234 / Clave: 1234. Haz clic en "Iniciar Turno".', 'info', 3000);
});

document.getElementById('btn-trial-admin-access')?.addEventListener('click', () => {
    document.getElementById('modal-trial-expired')?.classList.add('hidden');
    switchTab('admin');
    const adminPinInput = document.getElementById('admin-pin') as HTMLInputElement;
    if (adminPinInput) {
        adminPinInput.value = '';
        setTimeout(() => adminPinInput.focus(), 150);
    }
});

initApp();
