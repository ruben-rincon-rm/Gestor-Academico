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
    documento?: string;
    grado?: string;
    rol?: string;
    cargo?: string;
    clave?: string;
    foto?: string;
    isStaff?: boolean;
}

interface InventoryItem {
    id: string; // CÓDIGO DEL ELEMENTO o ID único
    nombre: string; // ELEMENTO
    codigo?: string; // CÓDIGO DEL ELEMENTO (Plaqueta)
    categoria?: string; // DEPORTES, SISTEMAS, BIBLIOTECA, OTROS
    caracteristicas?: string; // CARACTERÍSTICAS DEL ELEMENTO
    ubicacion?: string; // UBICACIÓN
    consecutivo?: string | number; // N°
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
export type AccessMode = 'ramon_munera' | 'modo_prueba' | 'institucion';
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
    titularClave?: string;
    titularRol?: string;
    titularContacto?: string;
    permisosPersonalizados?: Record<string, string[]>;
    modulosHabilitados?: string[];
    popupActivo: boolean;
    popupTitulo: string;
    popupUrl: string;
    popupDescripcion: string;
}

export const ALL_SYSTEM_MODULES = ['asistencia', 'evaluacion', 'salida', 'prestamo', 'pae', 'clase_ef', 'novedad'];

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
    modulosHabilitados: [...ALL_SYSTEM_MODULES],
    popupActivo: true,
    popupTitulo: "Tutorial & Carnet Digital Institucional",
    popupUrl: "https://www.youtube.com/watch?v=dQw4w9WgXcQ",
    popupDescripcion: "Bienvenido al sistema institucional de la I.E. Ramón Múnera Lopera. Consulta el tutorial de uso del carnet y los módulos académicos."
};

export const TRIAL_INSTITUTION_DEFAULT: InstitutionProfile = {
    id: "modo_prueba",
    nombre: "GESTOR ACADÉMICO (MODO PRUEBA 30 DÍAS)",
    logo: null,
    color1: "#1d4ed8",
    color2: "#059669",
    licenciaInicio: "2026-01-01",
    licenciaFin: "2026-12-31",
    tipoPlan: "prueba",
    limiteUsuarios: 9999,
    titularNombre: "Institución de Prueba Universal",
    modulosHabilitados: [...ALL_SYSTEM_MODULES],
    popupActivo: true,
    popupTitulo: "¡Bienvenido al Modo Prueba de 30 Días!",
    popupUrl: "https://www.youtube.com/watch?v=dQw4w9WgXcQ",
    popupDescripcion: "Plataforma escolar universal en Modo Prueba de 30 Días. Cualquier docente o institución puede evaluar todas las funciones usando el usuario Demo 1234 o personalizar su colegio en Administración. Desarrollada por www.espatodo.com"
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
    { ...TRIAL_INSTITUTION_DEFAULT },
    { ...RAMON_MUNERA_PROFILE }
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
        const registeredInst = localStorage.getItem('device_registered_institution') || 'ramon_munera';
        const found = institucionesList.find(i => i.id === registeredInst) || institucionesList.find(i => i.id === 'ramon_munera');
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
            
            if (type === 'items') {
                // CARGUE INTELIGENTE DE INVENTARIO MULTI-HOJA:
                // Hojas: DEPORTES, SISTEMAS, BIBLIOTECA, OTROS
                // Columnas: A1=N°, A2=ELEMENTO, A3=CÓDIGO DEL ELEMENTO, A4=CARACTERÍSTICAS DEL ELEMENTO, A5=UBICACIÓN
                workbook.SheetNames.forEach(sheetName => {
                    const rows: any[][] = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: "" });
                    if (rows.length < 2) return;

                    // Buscar fila de encabezados
                    let headerRowIdx = 0;
                    let cElem = -1;
                    let cCod = -1;
                    let cCaract = -1;
                    let cUbic = -1;
                    let cNum = -1;

                    for (let r = 0; r < Math.min(rows.length, 10); r++) {
                        const row = rows[r];
                        if (!row) continue;
                        row.forEach((cellVal: any, colIdx: number) => {
                            const valStr = String(cellVal || '').trim().toUpperCase();
                            if (valStr === 'N°' || valStr === 'NO' || valStr === 'NUM' || valStr === 'ITEM') cNum = colIdx;
                            else if (valStr.includes('ELEMENTO') && !valStr.includes('CÓDIGO') && !valStr.includes('CODIGO')) cElem = colIdx;
                            else if (valStr.includes('CÓDIGO') || valStr.includes('CODIGO') || valStr.includes('PLAQUETA')) cCod = colIdx;
                            else if (valStr.includes('CARACTER') || valStr.includes('DESCRIP')) cCaract = colIdx;
                            else if (valStr.includes('UBICAC') || valStr.includes('LUGAR') || valStr.includes('SALA')) cUbic = colIdx;
                        });
                        if (cCod !== -1 || cElem !== -1) {
                            headerRowIdx = r;
                            break;
                        }
                    }

                    // Fallback a columnas posicionales estándar A1..A5 si no tienen encabezados explícitos:
                    // Col 0: N°, Col 1: ELEMENTO, Col 2: CÓDIGO, Col 3: CARACTERÍSTICAS, Col 4: UBICACIÓN
                    if (cCod === -1) cCod = 2;
                    if (cElem === -1) cElem = 1;
                    if (cCaract === -1) cCaract = 3;
                    if (cUbic === -1) cUbic = 4;
                    if (cNum === -1) cNum = 0;

                    for (let i = headerRowIdx + 1; i < rows.length; i++) {
                        const row = rows[i];
                        if (!row || row.length === 0) continue;

                        const rawFirst = String(row[cNum] !== undefined ? row[cNum] : row[0] || '').trim().toUpperCase();
                        if (rawFirst.includes('TOTAL') || rawFirst.includes('FIRMA') || rawFirst.includes('OBSERVAC')) continue;

                        const elemRaw = String(row[cElem] || '').trim();
                        let codRaw = String(row[cCod] || '').trim();
                        const caractRaw = String(row[cCaract] || '').trim();
                        const ubicRaw = String(row[cUbic] || '').trim();
                        const numRaw = row[cNum] !== undefined ? String(row[cNum]).trim() : String(i);

                        // Si no tiene código explícito pero tiene elemento, autogenerar código seguro por categoría y consecutivo
                        if (!codRaw && elemRaw) {
                            const prefix = sheetName.substring(0, 3).toUpperCase();
                            codRaw = `${prefix}-${String(numRaw || i).padStart(4, '0')}`;
                        }

                        if (!codRaw && !elemRaw) continue;

                        const itemObj: InventoryItem = {
                            id: codRaw,
                            nombre: elemRaw || `Elemento ${codRaw}`,
                            codigo: codRaw,
                            categoria: sheetName.trim().toUpperCase(),
                            caracteristicas: caractRaw,
                            ubicacion: ubicRaw,
                            consecutivo: numRaw
                        };
                        parsed.push(itemObj);
                    }
                });

            } else {
                // Personal y PAE
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
                        else if (type === 'pae') { 
                            const idRaw = String(row[0]||'').trim();
                            if(idRaw) parsed.push({ id: idRaw, beneficiario: true }); 
                        }
                    }
                });
            }
            
            if (parsed.length === 0) { 
                if(countEl) countEl.innerText = "Vacío";
                return showToast('Excel vacío o formato incorrecto.', 'error'); 
            }
            
            // Si está en Modo Prueba (30 Días), persistir directamente en partición aislada
            if (currentAccessMode === 'modo_prueba') {
                if (type === 'items') {
                    localStorage.setItem('trial_inventario', JSON.stringify(parsed));
                    parsed.forEach(it => { inventarioDict[it.id] = it; });
                    if (countEl) countEl.innerText = `${parsed.length} registros (Prueba)`;
                    showToast(`¡Excelente! ${parsed.length} elementos de inventario guardados en tu entorno de prueba.`, 'success', 4000);
                } else if (type === 'staff') {
                    localStorage.setItem('trial_staff', JSON.stringify(parsed));
                    parsed.forEach(st => { staffDict[st.id] = st; });
                    if (countEl) countEl.innerText = `${parsed.length} personal (Prueba)`;
                    showToast(`¡Personal guardado en tu entorno de prueba!`, 'success', 4000);
                }
                if (progressEl) setTimeout(() => progressEl.style.width = "0%", 1000);
                return;
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

// Plantilla de Inventario para Préstamos dividida en 4 hojas:
// Hoja 1: DEPORTES, Hoja 2: SISTEMAS, Hoja 3: BIBLIOTECA, Hoja 4: OTROS
// Columnas: A1=N°, A2=ELEMENTO, A3=CÓDIGO DEL ELEMENTO, A4=CARACTERÍSTICAS DEL ELEMENTO, A5=UBICACIÓN
export function downloadSampleInventory() {
    const wb = XLSX.utils.book_new();

    // Hoja 1: DEPORTES
    const dataDeportes = [
        ["N°", "ELEMENTO", "CÓDIGO DEL ELEMENTO", "CARACTERÍSTICAS DEL ELEMENTO", "UBICACIÓN"],
        [1, "Balón de Baloncesto", "DEP-001", "Balón Molten #7 Cuero Sintético", "Caja 1 - Gimnasio"],
        [2, "Balón de Microfútbol", "DEP-002", "Balón Golty Tradicional #4", "Caja 2 - Gimnasio"],
        [3, "Balón de Voleibol", "DEP-003", "Balón Mikasa MVA200", "Caja 1 - Gimnasio"],
        [4, "Cronómetro Digital", "DEP-004", "Cronómetro deportivo 100 memorias", "Oficina Ed. Física"],
        [5, "Conos de Entrenamiento", "DEP-005", "Set de 10 conos naranja 23cm", "Estante 3 - Gimnasio"]
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(dataDeportes), "DEPORTES");

    // Hoja 2: SISTEMAS
    const dataSistemas = [
        ["N°", "ELEMENTO", "CÓDIGO DEL ELEMENTO", "CARACTERÍSTICAS DEL ELEMENTO", "UBICACIÓN"],
        [1, "Portátil Lenovo ThinkPad", "SIS-001", "Intel Core i5 16GB RAM SSD 512GB", "Sala de Cómputo 1"],
        [2, "Portátil HP ProBook", "SIS-002", "Intel Core i7 16GB RAM SSD 512GB", "Sala de Cómputo 1"],
        [3, "Videobeam Epson", "SIS-003", "Proyector 3600 Lúmenes HDMI", "Sala de Profesores"],
        [4, "Tableta Digitalizadora", "SIS-004", "Wacom Intuos Medium", "Sala de Sistemas 2"],
        [5, "Cámara Web Logitech", "SIS-005", "Webcam Full HD 1080p con micrófono", "Coordinación"]
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(dataSistemas), "SISTEMAS");

    // Hoja 3: BIBLIOTECA
    const dataBiblioteca = [
        ["N°", "ELEMENTO", "CÓDIGO DEL ELEMENTO", "CARACTERÍSTICAS DEL ELEMENTO", "UBICACIÓN"],
        [1, "Cien Años de Soledad", "BIB-001", "Libro G.G. Márquez Edición Conmemorativa", "Estante A - Literatura"],
        [2, "Álgebra de Baldor", "BIB-002", "Libro Matemáticas con CD", "Estante B - Ciencias"],
        [3, "Diccionario Español RAE", "BIB-003", "Diccionario Esencial Lengua Española", "Estante C - Consulta"],
        [4, "Enciclopedia Historia de Colombia", "BIB-004", "Tomo 1 y 2 Ilustrado", "Estante D - Sociales"],
        [5, "Lector de Códigos USB", "BIB-005", "Escáner láser para préstamos", "Mesa Principal Biblioteca"]
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(dataBiblioteca), "BIBLIOTECA");

    // Hoja 4: OTROS
    const dataOtros = [
        ["N°", "ELEMENTO", "CÓDIGO DEL ELEMENTO", "CARACTERÍSTICAS DEL ELEMENTO", "UBICACIÓN"],
        [1, "Bafle Sonido Portátil", "OTR-001", "Parlante recargable con micrófono inalámbrico", "Almacén General"],
        [2, "Extensión Eléctrica 20m", "OTR-002", "Extensión uso rudo calibre 12", "Mantenimiento"],
        [3, "Microscopio Binocular", "OTR-003", "Microscopio óptico laboratorio escolar", "Laboratorio Integrado"],
        [4, "Megáfono Recargable", "OTR-004", "Megáfono con sirena 50W", "Portería Principal"],
        [5, "Botiquín Primeros Auxilios", "OTR-005", "Maletín dotado reglamentario", "Enfermería"]
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(dataOtros), "OTROS");

    XLSX.writeFile(wb, "Plantilla_Inventario_4_Hojas.xlsx");
    showToast("Plantilla de Inventario (4 Hojas) generada.", "success");
}

// Plantilla de Personal Docente y Administrativo
export function downloadSampleStaff() {
    const wb = XLSX.utils.book_new();
    const dataStaff = [
        ["LISTADO DE PERSONAL DOCENTE Y ADMINISTRATIVO"],
        [],
        ["N°", "DOCUMENTO", "NOMBRE COMPLETO", "CARGO", "ROL", "CONTRASEÑA", "ÁREA / ASIGNATURA"],
        [1, "12345678", "GARCIA MARQUEZ GABRIEL", "Docente", "docente", "1234", "Lengua Castellana"],
        [2, "87654321", "RODRIGUEZ ANA MARIA", "Coordinadora", "coordinador", "2026", "Coordinación Académica"],
        [3, "11223344", "PEREZ GOMEZ JORGE", "Docente", "docente", "1122", "Educación Física"],
        [4, "44332211", "MARTINEZ LUIS ALBERTO", "Vigilante", "vigilante", "4433", "Seguridad / Portería"]
    ];
    const ws = XLSX.utils.aoa_to_sheet(dataStaff);
    XLSX.utils.book_append_sheet(wb, ws, "PERSONAL");
    XLSX.writeFile(wb, "Plantilla_Personal.xlsx");
    showToast("Plantilla de Personal generada y descargada.", "success");
}

// Plantilla de Beneficiarios PAE (Alimentación Escolar VIP)
export function downloadSamplePAE() {
    const wb = XLSX.utils.book_new();
    const dataPAE = [
        ["BASE DE BENEFICIARIOS PROGRAMA DE ALIMENTACIÓN ESCOLAR (PAE)"],
        [],
        ["N°", "DOCUMENTO / MATRÍCULA", "NOMBRE COMPLETO", "GRADO", "TIPO DE RACIÓN", "BENEFICIARIO"],
        [1, "1032033800", "BERNAL GARCIA JUSTIN ANDRES", "TS0501", "Almuerzo Completo", "SI"],
        [2, "1032033801", "DE LA OSSA PEREZ JUAN CARLOS", "TS0501", "Refrigerio Reforzado", "SI"],
        [3, "1032033802", "ZAPATA MARTINEZ VALENTINA", "TS0501", "Almuerzo Completo", "SI"],
        [4, "262001", "GOMEZ ALVAREZ CAMILO", "TS0502", "Refrigerio Reforzado", "SI"]
    ];
    const ws = XLSX.utils.aoa_to_sheet(dataPAE);
    XLSX.utils.book_append_sheet(wb, ws, "BENEFICIARIOS_PAE");
    XLSX.writeFile(wb, "Plantilla_PAE.xlsx");
    showToast("Plantilla PAE generada y descargada.", "success");
}

document.getElementById('btn-download-sample-excel')?.addEventListener('click', downloadSampleExcel);
document.getElementById('btn-download-sample-staff')?.addEventListener('click', downloadSampleStaff);
document.getElementById('btn-download-sample-inventory')?.addEventListener('click', downloadSampleInventory);
document.getElementById('btn-download-sample-pae')?.addEventListener('click', downloadSamplePAE);

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

        // Cargar inventario de prueba
        const localTrialItems = localStorage.getItem('trial_inventario');
        inventarioDict = {};
        if (localTrialItems) {
            try {
                const list = JSON.parse(localTrialItems);
                if (Array.isArray(list)) {
                    list.forEach((i: any) => { inventarioDict[i.id] = i; });
                } else if (typeof list === 'object') {
                    inventarioDict = list;
                }
            } catch(e) {}
        }

        paeBeneficiariosDict = {};

        const elStud = document.getElementById('count-students'); 
        if (elStud) elStud.innerText = Object.keys(studentsDict).length + ' alumnos (Prueba)';
        const elStaff = document.getElementById('count-staff'); 
        if (elStaff) elStaff.innerText = Object.keys(staffDict).length + ' personal (Prueba)';
        const elItem = document.getElementById('count-items'); 
        if (elItem) elItem.innerText = Object.keys(inventarioDict).length + ' registros';
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
        else if(mode==='pruebas' || (mode==='clase_ef' && efPhase==='pruebas')) { pathFn=getPruebasPath; color='purple'; } 
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
                const l2 = mode==='pae'?`PAE`:mode==='asistencia'?'Asistencia':mode==='evaluacion'?`Nota: ${r.nota}`:(mode==='pruebas' || (mode==='clase_ef' && efPhase==='pruebas'))?(r.lapNumber ? `Vta #${r.lapNumber} (${r.lapTime || ''})` : 'Prueba'):mode==='prestamo'?`Llevó: ${r.itemName}`:mode==='salida'?(r.status==='fuera'?'Salió':'Volvió'):mode==='clase_ef'?'Ed. Física':'Novedad';
                const item = document.createElement('div'); 
                item.className = `flex justify-between items-center p-3 bg-${color}-50 rounded-lg border border-${color}-100 shadow-sm`;
                item.innerHTML = `<div class="flex flex-col"><span class="font-bold text-gray-800 text-sm mb-1">${l1}</span><span class="text-11px font-bold text-${color}-600 uppercase">${l2}</span></div><span class="text-xs font-bold text-gray-500 bg-white px-2 py-1 rounded shadow-sm">${(mode === 'pruebas' || (mode==='clase_ef' && efPhase==='pruebas')) ? (r.totalTime || formatTimeWithSeconds(new Date(r.timestamp))) : formatTime(new Date(r.timestamp))}</span>`;
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
// CÓDIGOS QR INTELIGENTES (VINCULACIÓN DOCUMENTO / MATRÍCULA / INVENTARIO)
// Máxima seguridad anti-falsificación y compatibilidad 100% con carnets impresos anteriores
// ==========================================================
export interface ParsedQRResult {
    type: 'student' | 'item' | 'staff' | 'raw';
    id: string; // Documento o código de elemento resuelto
    matricula?: string;
    nombre?: string;
    grado?: string;
    categoria?: string;
    rawText: string;
}

export const generateSmartStudentQR = (student: StudentRecord): string => {
    // Formato compacto ultrarrápido y seguro:
    // STD|ID|MAT|NOMBRE|GRADO
    const id = String(student.documento || student.id || '').trim();
    const mat = String(student.matricula || '').trim();
    const nom = String(student.nombres || student.nombre || '').replace(/[|;]/g, ' ').trim();
    const gr = String(student.grado || '').replace(/[|;]/g, ' ').trim();
    return `STD|${id}|${mat}|${nom}|${gr}`;
};

export const generateSmartInventoryQR = (item: InventoryItem): string => {
    // Formato para plaquetas de equipos/elementos:
    // ITM|CODIGO|NOMBRE|CATEGORIA|UBICACION
    const cod = String(item.codigo || item.id || '').trim();
    const nom = String(item.nombre || '').replace(/[|;]/g, ' ').trim();
    const cat = String(item.categoria || 'INVENTARIO').replace(/[|;]/g, ' ').trim();
    const ubi = String(item.ubicacion || '').replace(/[|;]/g, ' ').trim();
    return `ITM|${cod}|${nom}|${cat}|${ubi}`;
};

export const decodeSmartQR = (scannedText: string): ParsedQRResult => {
    const text = String(scannedText || '').trim();
    if (!text) return { type: 'raw', id: '', rawText: '' };

    // 1. Detección de formato inteligente estructurado STD|... o ITM|...
    if (text.startsWith('STD|')) {
        const parts = text.split('|');
        const id = (parts[1] || '').trim();
        const mat = (parts[2] || '').trim();
        const nom = (parts[3] || '').trim();
        const gr = (parts[4] || '').trim();
        return {
            type: 'student',
            id: id || mat,
            matricula: mat,
            nombre: nom,
            grado: gr,
            rawText: text
        };
    }

    if (text.startsWith('ITM|')) {
        const parts = text.split('|');
        const cod = (parts[1] || '').trim();
        const nom = (parts[2] || '').trim();
        const cat = (parts[3] || '').trim();
        return {
            type: 'item',
            id: cod,
            nombre: nom,
            categoria: cat,
            rawText: text
        };
    }

    // 2. Detección de formato JSON si viniera estructurado
    if (text.startsWith('{') && text.endsWith('}')) {
        try {
            const parsed = JSON.parse(text);
            if (parsed.tipo === 'item' || parsed.type === 'item' || parsed.codigoElemento) {
                return {
                    type: 'item',
                    id: String(parsed.codigo || parsed.codigoElemento || parsed.id || '').trim(),
                    nombre: parsed.elemento || parsed.nombre,
                    categoria: parsed.categoria,
                    rawText: text
                };
            }
            if (parsed.doc || parsed.documento || parsed.matricula) {
                return {
                    type: 'student',
                    id: String(parsed.doc || parsed.documento || parsed.id || parsed.matricula).trim(),
                    matricula: parsed.matricula ? String(parsed.matricula).trim() : undefined,
                    nombre: parsed.nombre,
                    grado: parsed.grado,
                    rawText: text
                };
            }
        } catch (e) {}
    }

    // 3. Compatibilidad 100% Retrospectiva: Número suelto o código de barra/QR de carnets ya impresos
    // Buscar primero en diccionario de inventario
    if (inventarioDict[text]) {
        return {
            type: 'item',
            id: text,
            nombre: inventarioDict[text].nombre,
            categoria: inventarioDict[text].categoria,
            rawText: text
        };
    }

    // Buscar si coincide con item.codigo de alguna plaqueta de inventario
    const foundItem = Object.values(inventarioDict).find(i => 
        String(i.codigo || '').trim() === text || 
        String(i.id || '').trim() === text
    );
    if (foundItem) {
        return {
            type: 'item',
            id: foundItem.id,
            nombre: foundItem.nombre,
            categoria: foundItem.categoria,
            rawText: text
        };
    }

    // Buscar en estudiantes por ID, Documento o Matrícula
    if (studentsDict[text]) {
        return {
            type: 'student',
            id: studentsDict[text].id,
            matricula: studentsDict[text].matricula,
            nombre: getFullName(studentsDict[text]),
            grado: studentsDict[text].grado,
            rawText: text
        };
    }

    const foundStudent = Object.values(studentsDict).find(s => 
        String(s.matricula || '').trim() === text || 
        String(s.documento || '').trim() === text
    );
    if (foundStudent) {
        return {
            type: 'student',
            id: foundStudent.id,
            matricula: foundStudent.matricula,
            nombre: getFullName(foundStudent),
            grado: foundStudent.grado,
            rawText: text
        };
    }

    // Buscar en personal
    if (staffDict[text]) {
        return {
            type: 'staff',
            id: staffDict[text].id,
            nombre: staffDict[text].nombre,
            rawText: text
        };
    }

    // Si no coincide con nada, devolver texto plano para procesamiento libre
    return {
        type: 'raw',
        id: text,
        rawText: text
    };
};

// ==========================================================
// PROCESAMIENTO DE ESCANEO Y BÚSQUEDA
// Compatibilidad total: Búsqueda por Documento, Matrícula o QR Inteligente
// ==========================================================
export const processScan = async (scannedText: string) => {
    if(isProcessing || !currentStaff) return;
    isProcessing = true; 
    scannedText = scannedText.trim();
    if (!scannedText) { isProcessing = false; return; }

    // Decodificación inteligente del código QR (soporta formatos nuevos STD|... y los carnets anteriores de solo números)
    const qrInfo = decodeSmartQR(scannedText);
    let resolvedId = qrInfo.id || scannedText;

    // Compatibilidad adicional directa con alumnos por si el ID interno difiere
    if (!studentsDict[resolvedId]) {
        const found = Object.values(studentsDict).find(s => 
            String(s.matricula).trim() === resolvedId || 
            String(s.documento).trim() === resolvedId
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

// ==========================================================
// SUPER-CRONÓMETRO DE EDUCACIÓN FÍSICA & REGISTRO DE VUELTAS
// ==========================================================
let swIsRunning = false;
let swStartTime = 0;
let swElapsedTime = 0;
let swTimerId: any = null;

interface LapRecord {
    lapNumber: number;
    studentId: string;
    nombre: string;
    grado: string;
    lapTime: string;
    lapTimeMs: number;
    totalTime: string;
    totalTimeMs: number;
    timestamp: string;
}
let sessionLaps: LapRecord[] = [];
let studentLapData: Record<string, { count: number; lastLapMs: number }> = {};

const formatMsToString = (ms: number) => {
    const min = Math.floor(ms / 60000);
    const sec = Math.floor((ms % 60000) / 1000);
    const msec = ms % 1000;
    return `${String(min).padStart(2, '0')}:${String(sec).padStart(2, '0')}.${String(msec).padStart(3, '0')}`;
};

const getStopwatchElapsed = () => {
    return swElapsedTime + (swIsRunning ? (Date.now() - swStartTime) : 0);
};

const updateStopwatchDisplay = () => {
    const elapsed = getStopwatchElapsed();
    const min = Math.floor(elapsed / 60000);
    const sec = Math.floor((elapsed % 60000) / 1000);
    const msec = elapsed % 1000;

    const elMin = document.getElementById('sw-min');
    const elSec = document.getElementById('sw-sec');
    const elMs = document.getElementById('sw-ms');

    if (elMin) elMin.innerText = String(min).padStart(2, '0');
    if (elSec) elSec.innerText = String(sec).padStart(2, '0');
    if (elMs) elMs.innerText = String(msec).padStart(3, '0');
};

const startStopwatch = () => {
    if (swIsRunning) return;
    swIsRunning = true;
    swStartTime = Date.now() - swElapsedTime;
    if (swTimerId) clearInterval(swTimerId);
    swTimerId = setInterval(updateStopwatchDisplay, 30);

    const badge = document.getElementById('sw-status-badge');
    if (badge) {
        badge.className = "text-[10px] font-black px-2 py-0.5 rounded-full bg-emerald-600 text-white border border-emerald-400 animate-pulse";
        badge.innerText = "CORRIENDO";
    }
    const icon = document.getElementById('icon-sw-toggle');
    if (icon) icon.className = "fas fa-pause";
    const label = document.getElementById('label-sw-toggle');
    if (label) label.innerText = "PAUSAR";
};

const pauseStopwatch = () => {
    if (!swIsRunning) return;
    swIsRunning = false;
    swElapsedTime = Date.now() - swStartTime;
    if (swTimerId) {
        clearInterval(swTimerId);
        swTimerId = null;
    }
    updateStopwatchDisplay();

    const badge = document.getElementById('sw-status-badge');
    if (badge) {
        badge.className = "text-[10px] font-black px-2 py-0.5 rounded-full bg-amber-600 text-white border border-amber-400";
        badge.innerText = "PAUSADO";
    }
    const icon = document.getElementById('icon-sw-toggle');
    if (icon) icon.className = "fas fa-play";
    const label = document.getElementById('label-sw-toggle');
    if (label) label.innerText = "REANUDAR";
};

const resetStopwatch = () => {
    swIsRunning = false;
    swElapsedTime = 0;
    swStartTime = 0;
    if (swTimerId) {
        clearInterval(swTimerId);
        swTimerId = null;
    }
    updateStopwatchDisplay();

    const badge = document.getElementById('sw-status-badge');
    if (badge) {
        badge.className = "text-[10px] font-black px-2 py-0.5 rounded-full bg-gray-800 text-gray-300 border border-gray-600";
        badge.innerText = "PAUSADO";
    }
    const icon = document.getElementById('icon-sw-toggle');
    if (icon) icon.className = "fas fa-play";
    const label = document.getElementById('label-sw-toggle');
    if (label) label.innerText = "INICIAR CRONÓMETRO";
};

const toggleStopwatch = () => {
    if (swIsRunning) {
        pauseStopwatch();
    } else {
        startStopwatch();
    }
};

let lapToastTimeout: any = null;
const showLapHUDToast = (lapNumber: number, studentName: string, lapTime: string, totalTime: string) => {
    const banner = document.getElementById('sw-lap-toast');
    const title = document.getElementById('sw-lap-toast-title');
    const name = document.getElementById('sw-lap-toast-name');
    const time = document.getElementById('sw-lap-toast-time');

    if (title) title.innerText = `🏃 ¡VUELTA #${lapNumber} REGISTRADA!`;
    if (name) name.innerText = studentName;
    if (time) time.innerText = `Vuelta: ${lapTime} | Total: ${totalTime}`;

    if (banner) {
        banner.classList.remove('hidden');
        if (lapToastTimeout) clearTimeout(lapToastTimeout);
        lapToastTimeout = setTimeout(() => {
            banner.classList.add('hidden');
        }, 3500);
    }
};

const renderLapsTable = () => {
    const tbody = document.getElementById('pruebas-laps-table-body');
    const summary = document.getElementById('pruebas-laps-summary');
    if (summary) summary.innerText = `${sessionLaps.length} vueltas registradas en esta prueba`;
    if (!tbody) return;

    if (sessionLaps.length === 0) {
        tbody.innerHTML = `<tr><td colspan="5" class="p-3 text-center text-gray-400 italic">Inicia el cronómetro y escanea a los estudiantes al pasar por la meta.</td></tr>`;
        return;
    }

    tbody.innerHTML = sessionLaps.map((lap, idx) => `
        <tr class="hover:bg-purple-50 transition-colors">
            <td class="p-1.5 text-center font-bold text-gray-500">${sessionLaps.length - idx}</td>
            <td class="p-1.5 font-bold text-gray-800">${lap.nombre} <span class="text-[10px] text-gray-500 block font-normal">${lap.grado}</span></td>
            <td class="p-1.5 text-center font-black text-purple-700"><span class="bg-purple-100 px-1.5 py-0.5 rounded">Vta ${lap.lapNumber}</span></td>
            <td class="p-1.5 text-right font-mono font-bold text-emerald-600">${lap.lapTime}</td>
            <td class="p-1.5 text-right font-mono font-bold text-gray-700">${lap.totalTime}</td>
        </tr>
    `).join('');
};

const exportLapsToExcel = () => {
    if (sessionLaps.length === 0) {
        return showToast('No hay vueltas registradas en esta prueba todavía.', 'warning');
    }
    try {
        const rows = sessionLaps.map((lap, idx) => ({
            'N°': sessionLaps.length - idx,
            'Estudiante': lap.nombre,
            'Documento': lap.studentId,
            'Grado': lap.grado,
            'Vuelta': `Vuelta #${lap.lapNumber}`,
            'Tiempo Vuelta': lap.lapTime,
            'Tiempo Total': lap.totalTime,
            'Hora': formatTime(new Date(lap.timestamp))
        }));
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.json_to_sheet(rows);
        XLSX.utils.book_append_sheet(wb, ws, "Vueltas_Prueba_Fisica");
        XLSX.writeFile(wb, `Pruebas_Fisicas_Vueltas_${getTodayString()}.xlsx`);
        showToast('Planilla de vueltas exportada exitosamente.', 'success');
    } catch(e) {
        showToast('Error al exportar planilla de vueltas.', 'error');
    }
};

const handlePruebas = async (studentId: string) => {
    let student = studentsDict[studentId]; 
    if (!student) {
        student = Object.values(studentsDict).find(s => 
            String(s.matricula).trim() === studentId || 
            String(s.documento).trim() === studentId
        )!;
    }
    if (!student) { playBeep('error'); return showToast(`Estudiante no encontrado.`, 'error'); }
    
    const now = Date.now(); 
    if (pruebasCooldown[student.id] && (now - pruebasCooldown[student.id] < 3000)) { 
        playBeep('error'); 
        return showToast(`<b>⏳ ESPERA 3 SEG</b><br>${getFullName(student)} ya registró su paso.`, 'warning', 1800); 
    }
    pruebasCooldown[student.id] = now;

    // Si el cronómetro no está corriendo ni tiene tiempo, iniciarlo automáticamente
    if (!swIsRunning && swElapsedTime === 0) {
        startStopwatch();
    }

    const currentTotalMs = getStopwatchElapsed();
    if (!studentLapData[student.id]) {
        studentLapData[student.id] = { count: 0, lastLapMs: 0 };
    }
    const lapNumber = studentLapData[student.id].count + 1;
    const lapTimeMs = studentLapData[student.id].count === 0 
        ? currentTotalMs 
        : Math.max(0, currentTotalMs - studentLapData[student.id].lastLapMs);
    
    studentLapData[student.id] = { count: lapNumber, lastLapMs: currentTotalMs };

    const formattedLap = formatMsToString(lapTimeMs);
    const formattedTotal = formatMsToString(currentTotalMs);

    await setDoc(doc(getPruebasPath(), `${now}_${student.id}_vta${lapNumber}`), { 
        studentId: student.id, 
        nombre: getFullName(student), 
        grado: student.grado || 'Sin Grupo', 
        lapNumber: lapNumber,
        lapTime: formattedLap,
        lapTimeMs: lapTimeMs,
        totalTime: formattedTotal,
        totalTimeMs: currentTotalMs,
        date: getTodayString(), 
        timestamp: new Date().toISOString(), 
        recordedById: currentStaff.id, 
        recordedByName: currentStaff.nombre,
        tipo: 'vuelta_prueba_ef'
    });

    sessionLaps.unshift({
        lapNumber,
        studentId: student.id,
        nombre: getFullName(student),
        grado: student.grado || 'Sin Grupo',
        lapTime: formattedLap,
        lapTimeMs,
        totalTime: formattedTotal,
        totalTimeMs: currentTotalMs,
        timestamp: new Date().toISOString()
    });

    playBeep('success'); 
    showLapHUDToast(lapNumber, getFullName(student), formattedLap, formattedTotal);
    renderLapsTable();
    showToast(`<b>⏱️ VUELTA #${lapNumber}: ${getFullName(student)}</b><br><span class="font-mono font-bold text-emerald-300">Vuelta: ${formattedLap}</span> | Total: ${formattedTotal}`, 'success', 2200); 
    addRecentScanToUI(getFullName(student), `Vta #${lapNumber} (${formattedLap})`, new Date(), 'purple');
};

// ==========================================================
// CONTROL INTELIGENTE DE SALIDAS Y RETORNOS A LA INSTITUCIÓN
// ==========================================================
const updateActiveOutsideCount = async () => {
    try {
        const snap = await getDocs(getSalidasPath());
        let outsideCount = 0;
        snap.forEach(d => {
            const data: any = d.data();
            if (data.date === getTodayString() && (data.status === 'fuera' || data.status === 'salio' || data.status === 'salio_ef')) {
                outsideCount++;
            }
        });
        const summaryEl = document.getElementById('salidas-active-outside-summary');
        if (summaryEl) summaryEl.innerText = `Estudiantes fuera actualmente: ${outsideCount}`;
        const modalCountEl = document.getElementById('modal-outside-count');
        if (modalCountEl) modalCountEl.innerText = `${outsideCount} estudiantes fuera`;
    } catch(e) {
        console.error(e);
    }
};

const renderOutsideStudentsModal = async () => {
    const container = document.getElementById('outside-students-list-container');
    if (!container) return;
    container.innerHTML = '<div class="text-gray-400 text-center py-6 text-sm italic"><i class="fas fa-spinner fa-spin mr-2"></i>Consultando estudiantes fuera...</div>';
    
    try {
        const snap = await getDocs(getSalidasPath());
        const outsideList: any[] = [];
        const nowMs = Date.now();
        snap.forEach(d => {
            const data: any = d.data();
            if (data.date === getTodayString() && (data.status === 'fuera' || data.status === 'salio' || data.status === 'salio_ef')) {
                const timeOutMs = data.timeOut ? new Date(data.timeOut).getTime() : (data.timestamp ? new Date(data.timestamp).getTime() : 0);
                const duracionMin = Math.max(1, Math.round((nowMs - timeOutMs) / 60000));
                outsideList.push({ ...data, duracionMin, timeOutMs });
            }
        });

        const modalCountEl = document.getElementById('modal-outside-count');
        if (modalCountEl) modalCountEl.innerText = `${outsideList.length} estudiantes fuera`;
        const summaryEl = document.getElementById('salidas-active-outside-summary');
        if (summaryEl) summaryEl.innerText = `Estudiantes fuera actualmente: ${outsideList.length}`;

        if (outsideList.length === 0) {
            container.innerHTML = `
                <div class="text-center py-8 text-gray-400">
                    <i class="fas fa-check-circle text-3xl text-emerald-400 mb-2"></i>
                    <p class="text-sm font-bold text-gray-700">No hay estudiantes fuera actualmente</p>
                    <p class="text-xs text-gray-400 mt-1">Todos los alumnos que salieron hoy ya registraron su retorno o no se han marcado salidas.</p>
                </div>
            `;
            return;
        }

        outsideList.sort((a, b) => b.timeOutMs - a.timeOutMs);
        container.innerHTML = '';
        outsideList.forEach(item => {
            const div = document.createElement('div');
            div.className = "py-2.5 flex items-center justify-between gap-2";
            div.innerHTML = `
                <div class="flex-1 min-w-0">
                    <div class="font-bold text-sm text-gray-800 truncate">${item.nombre}</div>
                    <div class="text-xs text-gray-500 font-medium">Grado: ${item.grado || 'Sin Grupo'} &bull; Salió: <span class="font-bold text-amber-700">${item.horaSalida || formatTime(new Date(item.timeOutMs))}</span></div>
                    <div class="text-[11px] text-amber-600 font-bold mt-0.5"><i class="fas fa-clock mr-1"></i>Lleva fuera: ${item.duracionMin} min</div>
                </div>
                <button type="button" data-return-id="${item.studentId}" class="btn-mark-manual-return bg-emerald-600 hover:bg-emerald-700 active:scale-95 text-white font-bold text-xs px-2.5 py-1.5 rounded-lg shadow-sm flex items-center gap-1 transition-transform">
                    <i class="fas fa-sign-in-alt"></i> Marcar Retorno
                </button>
            `;
            container.appendChild(div);
        });

        container.querySelectorAll('.btn-mark-manual-return').forEach(btn => {
            btn.addEventListener('click', async (e: any) => {
                const sId = e.currentTarget.dataset.returnId;
                if (sId) {
                    await handleSalida(sId);
                    await renderOutsideStudentsModal();
                }
            });
        });
    } catch(e) {
        container.innerHTML = '<div class="text-red-500 text-center py-4 text-xs font-bold">Error cargando listado.</div>';
    }
};

const handleSalida = async (studentId: string) => {
    let student = studentsDict[studentId]; 
    if (!student) {
        student = Object.values(studentsDict).find(s => 
            String(s.matricula).trim() === studentId || 
            String(s.documento).trim() === studentId
        )!;
    }
    if (!student) { playBeep('error'); return showToast(`Estudiante no encontrado.`, 'error'); }
    
    const ref = doc(getSalidasPath(), `${getTodayString()}_${student.id}`); 
    const snap = await getDoc(ref); 
    const data = snap.exists() ? snap.data() : null;

    const nowIso = new Date().toISOString();
    const nowTimeStr = formatTimeWithSeconds(new Date());
    const nowMs = Date.now();

    // 1. Si el estudiante actualmente está FUERA de la institución:
    if (data && (data.status === 'fuera' || data.status === 'salio' || data.status === 'salio_ef')) {
        const timeOutMs = data.timeOut ? new Date(data.timeOut).getTime() : (data.timestamp ? new Date(data.timestamp).getTime() : 0);
        const elapsedSec = Math.floor((nowMs - timeOutMs) / 1000);

        // Protección anti-doble escaneo accidental (menos de 10 segundos):
        if (elapsedSec < 10) {
            playBeep('error');
            showToast(`
                <div class="text-center p-1">
                    <div class="text-2xl font-black text-amber-300">⚠️ YA TIENE SALIDA REGISTRADA</div>
                    <div class="text-base font-bold text-white mt-1">${getFullName(student)}</div>
                    <div class="text-xs text-yellow-200 mt-1">Salió hace solo ${elapsedSec} segundos (Hora: ${data.horaSalida || formatTimeWithSeconds(new Date(timeOutMs))}).</div>
                    <div class="text-xs text-amber-100 font-semibold mt-2 bg-amber-900/60 py-1 px-2 rounded">Estado actual: 🔴 FUERA DEL PLANTEL</div>
                    <div class="text-[11px] text-gray-300 mt-1">Para registrar el retorno, espere al momento del reingreso del alumno.</div>
                </div>
            `, 'warning', 4500);
            return;
        }

        // Si ya pasaron más de 10 segundos, este escaneo es un legítimo RETORNO a la institución:
        const duracionMin = Math.max(1, Math.round(elapsedSec / 60));
        await updateDoc(ref, {
            status: 'regreso',
            timeIn: nowIso,
            horaRetorno: nowTimeStr,
            duracionFueraMin: duracionMin,
            recordedByIdRetorno: currentStaff.id,
            recordedByNameRetorno: currentStaff.nombre
        });

        playBeep('success');
        showToast(`
            <div class="text-center p-1">
                <div class="text-2xl font-black text-emerald-300">🏫 RETORNO REGISTRADO</div>
                <div class="text-lg font-bold text-white mt-1">${getFullName(student)}</div>
                <div class="text-xs text-emerald-100">Grado: ${student.grado || 'Sin Grupo'}</div>
                <div class="text-sm font-black text-green-300 bg-emerald-900/60 py-1.5 px-3 rounded-lg mt-2 inline-block">
                    Hora Retorno: ${nowTimeStr}
                </div>
                <div class="text-xs text-emerald-200 mt-1 font-semibold">Tiempo que estuvo fuera: <b>${duracionMin} minutos</b></div>
                <div class="text-xs text-emerald-100 mt-1">Estado: 🟢 DENTRO DE LA INSTITUCIÓN</div>
            </div>
        `, 'success', 4000);
        addRecentScanToUI(getFullName(student), `Volvió (${duracionMin}m fuera)`, new Date(), 'green');
        updateActiveOutsideCount();
        return;
    }

    // 2. Si el estudiante ya había retornado hoy:
    if (data && (data.status === 'regreso' || data.status === 'regreso_ef')) {
        const timeInMs = data.timeIn ? new Date(data.timeIn).getTime() : 0;
        const elapsedSecSinceIn = Math.floor((nowMs - timeInMs) / 1000);

        // Si fue escaneado hace menos de 10 segundos del retorno:
        if (elapsedSecSinceIn < 10) {
            playBeep('error');
            showToast(`
                <div class="text-center p-1">
                    <div class="text-2xl font-black text-green-300">✅ YA SE ENCUENTRA DENTRO</div>
                    <div class="text-base font-bold text-white mt-1">${getFullName(student)}</div>
                    <div class="text-xs text-emerald-200 mt-1">Retornó a las ${data.horaRetorno || nowTimeStr}.</div>
                    <div class="text-xs text-green-100 font-semibold mt-1">Estado: 🟢 DENTRO DE LA INSTITUCIÓN</div>
                </div>
            `, 'info', 3500);
            return;
        }

        // Si ya pasaron más de 10 segundos y vuelve a salir (segunda salida del día):
        await setDoc(ref, { 
            studentId: student.id, 
            nombre: getFullName(student), 
            grado: student.grado || 'Sin Grupo', 
            status: 'fuera', 
            date: getTodayString(), 
            timestamp: nowIso, 
            timeOut: nowIso, 
            horaSalida: nowTimeStr, 
            recordedById: currentStaff.id, 
            recordedByName: currentStaff.nombre,
            salidaAnterior: {
                horaSalida: data.horaSalida,
                horaRetorno: data.horaRetorno,
                duracionFueraMin: data.duracionFueraMin || 0
            }
        });

        playBeep('success');
        showToast(`
            <div class="text-center p-1">
                <div class="text-2xl font-black text-amber-300">🚪 NUEVA SALIDA REGISTRADA</div>
                <div class="text-lg font-bold text-white mt-1">${getFullName(student)}</div>
                <div class="text-xs text-gray-200">Grado: ${student.grado || 'Sin Grupo'}</div>
                <div class="text-sm font-black text-yellow-300 bg-amber-900/60 py-1.5 px-3 rounded-lg mt-2 inline-block">
                    Hora Salida: ${nowTimeStr}
                </div>
                <div class="text-xs text-amber-200 mt-1 font-semibold">Estado: 🔴 FUERA DE LA INSTITUCIÓN</div>
            </div>
        `, 'warning', 4000);
        addRecentScanToUI(getFullName(student), 'Salió de la IE', new Date(), 'gray');
        updateActiveOutsideCount();
        return;
    }

    // 3. Primera salida del día:
    await setDoc(ref, { 
        studentId: student.id, 
        nombre: getFullName(student), 
        grado: student.grado || 'Sin Grupo', 
        status: 'fuera', 
        date: getTodayString(), 
        timestamp: nowIso, 
        timeOut: nowIso, 
        horaSalida: nowTimeStr, 
        recordedById: currentStaff.id, 
        recordedByName: currentStaff.nombre 
    }); 

    playBeep('success'); 
    showToast(`
        <div class="text-center p-1">
            <div class="text-2xl font-black text-amber-300">🚪 SALIDA REGISTRADA</div>
            <div class="text-lg font-bold text-white mt-1">${getFullName(student)}</div>
            <div class="text-xs text-gray-200">Grado: ${student.grado || 'Sin Grupo'}</div>
            <div class="text-sm font-black text-yellow-300 bg-amber-900/60 py-1.5 px-3 rounded-lg mt-2 inline-block">
                Hora Salida: ${nowTimeStr}
            </div>
            <div class="text-xs text-amber-200 mt-1 font-semibold">Estado: 🔴 FUERA DE LA INSTITUCIÓN</div>
            <div class="text-[11px] text-gray-300 mt-1">Al regresar, vuelva a escanear su carnet para registrar su retorno.</div>
        </div>
    `, 'warning', 4000); 
    addRecentScanToUI(getFullName(student), 'Salió de la IE', new Date(), 'gray'); 
    updateActiveOutsideCount();
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
    } else if (efPhase === 'pruebas') {
        await handlePruebas(scannedText);
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
        userObj = staffDict[query] || DEMO_STAFF[query]; 
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
    const hLogo = document.getElementById('header-logo') as HTMLImageElement;
    const hLogoBox = document.getElementById('header-logo-container');
    const hDefaultIcon = document.getElementById('header-default-icon');
    const loginBadge = document.getElementById('login-institution-badge');
    const headerPlanText = document.getElementById('header-plan-text');

    if (currentAccessMode === 'modo_prueba') {
        if (hName) hName.innerText = "GESTOR ACADÉMICO";
        hLogoBox?.classList.add('hidden');
        hDefaultIcon?.classList.remove('hidden');
        if (loginBadge) loginBadge.innerText = "Modo Prueba (30 Días)";
        if (headerPlanText) headerPlanText.innerText = "Prueba 30 Días";
    } else {
        if (hName && institucionData.nombre) hName.innerText = institucionData.nombre;
        if (institucionData.logo) { 
            if (hLogo) hLogo.src = institucionData.logo; 
            hLogoBox?.classList.remove('hidden'); 
            hDefaultIcon?.classList.add('hidden'); 
        } else {
            hLogoBox?.classList.add('hidden');
            hDefaultIcon?.classList.remove('hidden');
        }
        if (loginBadge) loginBadge.innerText = institucionData.nombre;
        if (headerPlanText) headerPlanText.innerText = "Plan Institucional";
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
// =========================================================================
// MÓDULO SUPER-ADMINISTRADOR (Ventas, Licencias, Fechas, Roles y Formateo) - www.espatodo.com
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

// Cargar lista de instituciones y personas en los selectores del Super-Admin
export const renderSuperAdminInstitutionsDropdown = () => {
    const select = document.getElementById('sa-inst-select') as HTMLSelectElement;
    const formatSelect = document.getElementById('sa-format-target-select') as HTMLSelectElement;
    
    if (select) {
        select.innerHTML = '';
        institucionesList.forEach(inst => {
            const opt = document.createElement('option');
            opt.value = inst.id;
            opt.innerText = `${inst.nombre} (${inst.tipoPlan === 'institucional' ? 'Institucional' : 'Docente/Ind.'})`;
            if (inst.id === institucionData.id) opt.selected = true;
            select.appendChild(opt);
        });
    }

    if (formatSelect) {
        formatSelect.innerHTML = '';
        const optCur = document.createElement('option');
        optCur.value = 'current';
        optCur.innerText = `Institución Activa (${institucionData.nombre})`;
        formatSelect.appendChild(optCur);

        institucionesList.forEach(inst => {
            const opt = document.createElement('option');
            opt.value = inst.id;
            opt.innerText = `${inst.nombre} [ID: ${inst.id}]`;
            formatSelect.appendChild(opt);
        });
    }
};

// Cargar los datos de una institución o docente en el modal de Super-Admin
export const loadInstitutionIntoSuperAdmin = (targetId: string) => {
    const inst = institucionesList.find(i => i.id === targetId) || institucionData;
    
    const iName = document.getElementById('sa-inst-name-input') as HTMLInputElement;
    const c1 = document.getElementById('sa-inst-color1') as HTMLInputElement;
    const c2 = document.getElementById('sa-inst-color2') as HTMLInputElement;
    const tNombre = document.getElementById('sa-titular-nombre') as HTMLInputElement;
    const tDoc = document.getElementById('sa-titular-doc') as HTMLInputElement;
    const tClave = document.getElementById('sa-titular-clave') as HTMLInputElement;
    const tRol = document.getElementById('sa-titular-rol') as HTMLSelectElement;
    const tContacto = document.getElementById('sa-titular-contacto') as HTMLInputElement;
    const lStart = document.getElementById('sa-licencia-inicio') as HTMLInputElement;
    const lEnd = document.getElementById('sa-licencia-fin') as HTMLInputElement;
    const mPin = document.getElementById('sa-master-pin') as HTMLInputElement;

    if (iName) iName.value = inst.nombre || '';
    if (c1) c1.value = inst.color1 || '#2563eb';
    if (c2) c2.value = inst.color2 || '#0ea5e9';
    if (tNombre) tNombre.value = inst.titularNombre || inst.nombre || '';
    if (tDoc) tDoc.value = inst.titularDoc || '';
    if (tClave) {
        const existingStaffClave = inst.titularDoc ? staffDict[inst.titularDoc]?.clave : '';
        tClave.value = inst.titularClave || existingStaffClave || (inst.titularDoc ? '1234' : '');
    }
    if (tRol) {
        const existingStaffRol = inst.titularDoc ? staffDict[inst.titularDoc]?.rol : '';
        tRol.value = inst.titularRol || existingStaffRol || (inst.tipoPlan === 'docente' ? 'docente' : 'docente');
    }
    if (tContacto) tContacto.value = inst.titularContacto || '';
    if (lStart) lStart.value = inst.licenciaInicio || '2026-01-01';
    if (lEnd) lEnd.value = inst.licenciaFin || '2027-12-31';
    if (mPin) mPin.value = '';

    // Seleccionar plan y correlacionar (respetando fechas de la institución cargada)
    selectSuperAdminPlan(inst.tipoPlan || 'institucional', true);

    // Cargar módulos habilitados
    const enabledMods = inst.modulosHabilitados && inst.modulosHabilitados.length > 0
        ? inst.modulosHabilitados
        : ALL_SYSTEM_MODULES;

    document.querySelectorAll('.sa-module-toggle').forEach((cb: any) => {
        cb.checked = enabledMods.includes(cb.value);
    });
    updateSuperAdminModulesCount();
    updateSuperAdminCountdown();
    renderSuperAdminRolesMatrix();
};

export const openSuperAdminModal = () => {
    closeSuperAdminLogin();
    const modal = document.getElementById('modal-superadmin');
    if (!modal) return;

    renderSuperAdminInstitutionsDropdown();
    loadInstitutionIntoSuperAdmin(institucionData.id);

    modal.classList.remove('hidden');
};

export const updateSuperAdminModulesCount = () => {
    const checkedCount = document.querySelectorAll('.sa-module-toggle:checked').length;
    const badge = document.getElementById('sa-enabled-modules-count');
    if (badge) {
        badge.innerText = `${checkedCount} Módulo${checkedCount === 1 ? '' : 's'} Activo${checkedCount === 1 ? '' : 's'}`;
        badge.className = checkedCount > 0 
            ? "text-[10px] font-black px-2.5 py-1 rounded-full bg-emerald-500/20 text-emerald-300 border border-emerald-500/40"
            : "text-[10px] font-black px-2.5 py-1 rounded-full bg-red-500/20 text-red-300 border border-red-500/40";
    }
};

export const closeSuperAdminModal = () => {
    document.getElementById('modal-superadmin')?.classList.add('hidden');
};

// CORRELACIÓN TOTAL EN SUPERADMINISTRADOR:
// 1. Modalidad de Compra con 2. Módulos Habilitados, 3. Fechas de Acceso, 4. Matriz de Permisos por Rol
export const selectSuperAdminPlan = (plan: LicenseType, preserveDates = false) => {
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
    const lStart = document.getElementById('sa-licencia-inicio') as HTMLInputElement;
    const lEnd = document.getElementById('sa-licencia-fin') as HTMLInputElement;

    if (plan === 'docente') {
        if (pill) {
            pill.className = "text-[10px] font-black px-2.5 py-1 rounded-full bg-purple-500/20 text-purple-300 border border-purple-500/40 uppercase";
            pill.innerText = "Plan Docente Individual (1 Usuario)";
        }
        // Correlación 2: Si estaban todos los 7 módulos seleccionados, sugerir los esenciales de docente (asistencia + notas)
        const currentlyChecked = Array.from(document.querySelectorAll('.sa-module-toggle:checked')).map((cb: any) => cb.value);
        if (currentlyChecked.length === ALL_SYSTEM_MODULES.length) {
            document.querySelectorAll('.sa-module-toggle').forEach((cb: any) => {
                cb.checked = ['asistencia', 'evaluacion'].includes(cb.value);
            });
            updateSuperAdminModulesCount();
        }
        // Correlación 3: Fechas de acceso - 1 Año escolar si estaba vencido y no se está preservando
        if (!preserveDates && lEnd && (!lEnd.value || lEnd.value < getTodayString())) {
            const nextYear = new Date();
            nextYear.setFullYear(nextYear.getFullYear() + 1);
            lEnd.value = nextYear.toISOString().split('T')[0];
        }
        // Correlación 4: Matriz de Permisos por Rol (Se inhabilita para docente personal)
        renderSuperAdminRolesMatrix();
    } else if (plan === 'prueba') {
        if (pill) {
            pill.className = "text-[10px] font-black px-2.5 py-1 rounded-full bg-yellow-500/20 text-yellow-300 border border-yellow-500/40 uppercase";
            pill.innerText = "Modo Prueba 30 Días";
        }
        // Correlación 2: Todos los módulos activos para evaluación
        document.querySelectorAll('.sa-module-toggle').forEach((cb: any) => {
            cb.checked = true;
        });
        updateSuperAdminModulesCount();
        // Correlación 3: Fechas fijadas exactamente a 30 días si no se preservan fechas específicas
        if (!preserveDates) {
            if (lStart) lStart.value = getTodayString();
            if (lEnd) {
                const date30 = new Date(Date.now() + 30 * 24 * 60 * 60 * 1000);
                lEnd.value = date30.toISOString().split('T')[0];
            }
        }
        renderSuperAdminRolesMatrix();
    } else {
        // Plan Institucional Completo
        if (pill) {
            pill.className = "text-[10px] font-black px-2.5 py-1 rounded-full bg-blue-500/20 text-blue-300 border border-blue-500/40 uppercase";
            pill.innerText = "Plan Institucional Completo";
        }
        // Correlación 2: Todos los módulos habilitados por defecto
        document.querySelectorAll('.sa-module-toggle').forEach((cb: any) => {
            cb.checked = true;
        });
        updateSuperAdminModulesCount();
        // Correlación 3: Fechas - 1 Año Escolar si no se preserva
        if (!preserveDates && lEnd && (!lEnd.value || lEnd.value < getTodayString())) {
            const nextYear = new Date();
            nextYear.setFullYear(nextYear.getFullYear() + 1);
            lEnd.value = nextYear.toISOString().split('T')[0];
        }
        renderSuperAdminRolesMatrix();
    }

    updateSuperAdminCountdown();
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

    const planRadio = document.querySelector('input[name="sa_plan_choice"]:checked') as HTMLInputElement;
    const currentPlan = planRadio?.value || institucionData.tipoPlan || 'institucional';

    // Si fue comprado por un solo docente, la matriz de roles está inactiva y no aplica
    if (currentPlan === 'docente') {
        container.innerHTML = `
            <div class="p-4 bg-purple-950/40 border border-purple-500/40 rounded-xl text-xs flex items-center gap-3">
                <i class="fas fa-user-lock text-purple-400 text-2xl flex-shrink-0"></i>
                <div class="text-purple-200">
                    <strong class="font-black text-sm block mb-0.5 text-white">Matriz Inactiva para Plan Docente Individual</strong>
                    <span>Este sistema fue adquirido por un solo docente personal. La matriz de roles por cargos y la nómina masiva aplican únicamente cuando se adquiere la licencia a nivel Institucional para colegios completos.</span>
                </div>
            </div>
        `;
        return;
    }

    const roleDefs = [
        { key: 'docente', name: 'Docente Titular', icon: 'fa-chalkboard-teacher', desc: 'Asistencia de aula, planilla de notas y evaluaciones' },
        { key: 'coordinador', name: 'Coordinación Académica', icon: 'fa-user-graduate', desc: 'Asistencia general, observación y autorizaciones' },
        { key: 'vigilante', name: 'Seguridad / Portería', icon: 'fa-door-open', desc: 'Control de ingresos y salidas en puerta' }
    ];

    const currentEnabledMods = Array.from(document.querySelectorAll('.sa-module-toggle:checked')).map((el: any) => el.value);

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
            if (m.id === 'pruebas') return;
            const isChecked = rolePermissions[r.key]?.includes(m.id) ? 'checked' : '';
            const isModActive = currentEnabledMods.length === 0 || currentEnabledMods.includes(m.id);
            html += `
            <label class="flex items-center gap-1.5 text-[11px] ${isModActive ? 'text-gray-300 hover:text-white' : 'text-gray-600 line-through'} cursor-pointer select-none">
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
        const select = document.getElementById('sa-inst-select') as HTMLSelectElement;
        const targetInstId = select?.value || institucionData.id;
        const targetInst = institucionesList.find(i => i.id === targetInstId) || institucionData;

        const iName = (document.getElementById('sa-inst-name-input') as HTMLInputElement)?.value.trim() || targetInst.nombre;
        const c1 = (document.getElementById('sa-inst-color1') as HTMLInputElement)?.value || targetInst.color1;
        const c2 = (document.getElementById('sa-inst-color2') as HTMLInputElement)?.value || targetInst.color2;

        const planRadio = document.querySelector('input[name="sa_plan_choice"]:checked') as HTMLInputElement;
        const chosenPlan = (planRadio?.value || 'institucional') as LicenseType;
        const tNombre = (document.getElementById('sa-titular-nombre') as HTMLInputElement)?.value.trim() || iName;
        const tDoc = (document.getElementById('sa-titular-doc') as HTMLInputElement)?.value.trim();
        const tClave = (document.getElementById('sa-titular-clave') as HTMLInputElement)?.value.trim() || tDoc || '1234';
        const tRol = (document.getElementById('sa-titular-rol') as HTMLSelectElement)?.value || 'docente';
        const tContacto = (document.getElementById('sa-titular-contacto') as HTMLInputElement)?.value.trim();
        const lStart = (document.getElementById('sa-licencia-inicio') as HTMLInputElement)?.value || '2026-01-01';
        const lEnd = (document.getElementById('sa-licencia-fin') as HTMLInputElement)?.value || '2027-12-31';

        // Módulos habilitados personalizados
        const selectedModules = Array.from(document.querySelectorAll('.sa-module-toggle:checked')).map((el: any) => el.value);
        const modulosHabilitados = selectedModules.length > 0 ? selectedModules : ['asistencia'];

        // Permisos de roles si es institucional
        let newPerms: Record<string, string[]> = { ...rolePermissions };
        if (chosenPlan === 'institucional') {
            newPerms = { docente: [], vigilante: [], coordinador: [] };
            document.querySelectorAll('.sa-role-cb').forEach((cb: any) => {
                if (cb.checked) {
                    const r = cb.dataset.role;
                    const m = cb.dataset.module;
                    if (!newPerms[r]) newPerms[r] = [];
                    newPerms[r].push(m);
                }
            });
            rolePermissions = newPerms;
        }

        // Actualizar datos del target en la lista
        targetInst.nombre = iName;
        targetInst.color1 = c1;
        targetInst.color2 = c2;
        targetInst.tipoPlan = chosenPlan;
        targetInst.limiteUsuarios = chosenPlan === 'docente' ? 1 : 9999;
        targetInst.titularNombre = tNombre;
        targetInst.titularDoc = tDoc;
        targetInst.titularClave = tClave;
        targetInst.titularRol = tRol;
        targetInst.titularContacto = tContacto;
        targetInst.licenciaInicio = lStart;
        targetInst.licenciaFin = lEnd;
        targetInst.modulosHabilitados = modulosHabilitados;
        targetInst.permisosPersonalizados = newPerms;

        // Si se especificó documento del titular, registrar o actualizar de inmediato al usuario en la nómina
        if (tDoc) {
            const staffMember: StaffRecord = {
                id: tDoc,
                nombre: tNombre || targetInst.nombre,
                rol: tRol || 'docente',
                cargo: tRol || 'docente',
                clave: tClave || tDoc,
                grado: tRol === 'docente' ? 'Docente Titular' : (tRol || 'General'),
                isStaff: true
            };
            staffDict[tDoc] = staffMember;
            try {
                await setDoc(doc(getStaffPath(), tDoc), staffMember);
            } catch(e) {}
            try {
                localStorage.setItem('trial_staff', JSON.stringify(staffDict));
                localStorage.setItem('staff_cache', JSON.stringify(staffDict));
            } catch(e) {}
        }

        // Si es la institución activa, actualizar institucionData
        if (targetInst.id === institucionData.id) {
            institucionData = { ...targetInst };
        }

        // Guardar lista completa de instituciones y configuración
        const payloadToSave = {
            institucionActivaId: institucionData.id,
            institucionActiva: institucionData,
            instituciones: institucionesList
        };

        try {
            await setDoc(getSettingsPath(), payloadToSave);
            await setDoc(doc(db, 'configuracion_app', 'superadmin_config'), {
                targetInstId: targetInst.id,
                tipoPlan: chosenPlan,
                titularNombre: tNombre,
                titularDoc: tDoc,
                titularContacto: tContacto,
                licenciaInicio: lStart,
                licenciaFin: lEnd,
                modulosHabilitados,
                updatedAt: new Date().toISOString()
            });
            await setDoc(getPermissionsPath(), rolePermissions);
        } catch(e) {}

        localStorage.setItem('institucionData', JSON.stringify(payloadToSave));
        localStorage.setItem('rolePermissions', JSON.stringify(rolePermissions));

        // Refrescar UI general
        updatePlanRestrictionsUI();
        checkLicenseValidity();
        aplicarConfiguracionUI();
        renderRolesConfig();
        if (currentStaff) applyRolesUI(currentStaff);

        closeSuperAdminModal();
        showToast(`Licencia Super-Admin guardada para "${targetInst.nombre}" (${chosenPlan.toUpperCase()})`, 'success', 4000);
    } catch(err) {
        showToast('Error al guardar configuración Super-Admin', 'error');
    } finally {
        if (btnSave) {
            btnSave.disabled = false;
            btnSave.innerHTML = '<i class="fas fa-save mr-1"></i> Guardar y Aplicar Licencia';
        }
    }
};

// =========================================================================
// ACCIONES DE FORMATEO DEL SISTEMA (EXCLUSIVO SUPER-ADMINISTRADOR)
// =========================================================================
const getSelectedFormatTarget = (): InstitutionProfile => {
    const sel = document.getElementById('sa-format-target-select') as HTMLSelectElement;
    const val = sel?.value;
    if (!val || val === 'current') return institucionData;
    return institucionesList.find(i => i.id === val) || institucionData;
};

// Borrar Alumnos
document.getElementById('btn-sa-format-students')?.addEventListener('click', async () => {
    const target = getSelectedFormatTarget();
    if (confirm(`¿Estás seguro de que deseas eliminar TODOS LOS ALUMNOS de "${target.nombre}"?`)) {
        const conf = prompt(`Escribe "BORRAR" para confirmar el vaciado de alumnos de "${target.nombre}":`);
        if (conf === "BORRAR") {
            try {
                showToast("Borrando alumnos...", "warning", 3000);
                const snap = await getDocs(getStudentsPath());
                await Promise.all(snap.docs.map(d => deleteDoc(d.ref)));
                studentsDict = {};
                await loadDatabases();
                showToast(`Alumnos de "${target.nombre}" eliminados con éxito.`, "success");
            } catch(e) {
                showToast("Error al borrar alumnos", "error");
            }
        }
    }
});

// Borrar Personal Docente
document.getElementById('btn-sa-format-staff')?.addEventListener('click', async () => {
    const target = getSelectedFormatTarget();
    if (confirm(`¿Estás seguro de que deseas eliminar EL PERSONAL DOCENTE de "${target.nombre}"?`)) {
        const conf = prompt(`Escribe "BORRAR" para confirmar el vaciado de nómina de "${target.nombre}":`);
        if (conf === "BORRAR") {
            try {
                showToast("Borrando personal...", "warning", 3000);
                const snap = await getDocs(getStaffPath());
                await Promise.all(snap.docs.map(d => deleteDoc(d.ref)));
                staffDict = {};
                await loadDatabases();
                showToast(`Personal docente de "${target.nombre}" eliminado.`, "success");
            } catch(e) {
                showToast("Error al borrar personal", "error");
            }
        }
    }
});

// Borrar Inventario
document.getElementById('btn-sa-format-inventory')?.addEventListener('click', async () => {
    const target = getSelectedFormatTarget();
    if (confirm(`¿Estás seguro de que deseas vaciar EL INVENTARIO de "${target.nombre}"?`)) {
        const conf = prompt(`Escribe "BORRAR" para confirmar:`);
        if (conf === "BORRAR") {
            try {
                showToast("Borrando inventario...", "warning", 3000);
                const snap = await getDocs(getInventarioPath());
                await Promise.all(snap.docs.map(d => deleteDoc(d.ref)));
                inventarioDict = {};
                await loadDatabases();
                showToast(`Inventario de "${target.nombre}" restablecido.`, "success");
            } catch(e) {
                showToast("Error al borrar inventario", "error");
            }
        }
    }
});

// Borrar Registros de Asistencia y Evaluaciones
document.getElementById('btn-sa-format-records')?.addEventListener('click', async () => {
    const target = getSelectedFormatTarget();
    if (confirm(`¿Estás seguro de que deseas vaciar EL HISTORIAL DE REGISTROS (Asistencias, Evaluaciones, Salidas) de "${target.nombre}"?`)) {
        const conf = prompt(`Escribe "BORRAR" para confirmar:`);
        if (conf === "BORRAR") {
            try {
                showToast("Limpiando registros de asistencia...", "warning", 3000);
                const [attSnap, evalSnap, mealsSnap, loansSnap, salSnap] = await Promise.all([
                    getDocs(getAttendancePath()),
                    getDocs(getEvaluacionesPath()),
                    getDocs(getMealsPath()),
                    getDocs(getLoansPath()),
                    getDocs(getSalidasPath())
                ]);
                await Promise.all([
                    ...attSnap.docs.map(d => deleteDoc(d.ref)),
                    ...evalSnap.docs.map(d => deleteDoc(d.ref)),
                    ...mealsSnap.docs.map(d => deleteDoc(d.ref)),
                    ...loansSnap.docs.map(d => deleteDoc(d.ref)),
                    ...salSnap.docs.map(d => deleteDoc(d.ref))
                ]);
                const counterEl = document.getElementById('counter-today');
                if (counterEl) counterEl.innerText = '0 registros';
                const recList = document.getElementById('recent-scans');
                if (recList) recList.innerHTML = '<div class="text-gray-400 text-center text-sm py-4"><i class="fas fa-list text-2xl mb-2 text-gray-200"></i><br>Sin registros.</div>';
                showToast(`Historial de registros de "${target.nombre}" borrado con éxito.`, "success");
            } catch(e) {
                showToast("Error al borrar registros", "error");
            }
        }
    }
});

// Formateo Completo de la Institución
document.getElementById('btn-sa-format-full')?.addEventListener('click', async () => {
    const target = getSelectedFormatTarget();
    if (confirm(`⚠️ ALERTA MÁXIMA: ¿Deseas formatear COMPLETAMENTE a "${target.nombre}"? Se borrarán alumnos, personal, inventario y registros.`)) {
        const conf = prompt(`Escribe "BORRAR TODO" en mayúsculas para confirmar:`);
        if (conf === "BORRAR TODO") {
            try {
                showToast("Ejecutando formateo completo...", "warning", 4000);
                const [sSnap, stSnap, iSnap, attSnap, evalSnap, mealsSnap, loansSnap, salSnap] = await Promise.all([
                    getDocs(getStudentsPath()),
                    getDocs(getStaffPath()),
                    getDocs(getInventarioPath()),
                    getDocs(getAttendancePath()),
                    getDocs(getEvaluacionesPath()),
                    getDocs(getMealsPath()),
                    getDocs(getLoansPath()),
                    getDocs(getSalidasPath())
                ]);
                await Promise.all([
                    ...sSnap.docs.map(d => deleteDoc(d.ref)),
                    ...stSnap.docs.map(d => deleteDoc(d.ref)),
                    ...iSnap.docs.map(d => deleteDoc(d.ref)),
                    ...attSnap.docs.map(d => deleteDoc(d.ref)),
                    ...evalSnap.docs.map(d => deleteDoc(d.ref)),
                    ...mealsSnap.docs.map(d => deleteDoc(d.ref)),
                    ...loansSnap.docs.map(d => deleteDoc(d.ref)),
                    ...salSnap.docs.map(d => deleteDoc(d.ref))
                ]);
                studentsDict = {};
                staffDict = {};
                inventarioDict = {};
                const counterEl = document.getElementById('counter-today');
                if (counterEl) counterEl.innerText = '0 registros';
                const recList = document.getElementById('recent-scans');
                if (recList) recList.innerHTML = '<div class="text-gray-400 text-center text-sm py-4"><i class="fas fa-list text-2xl mb-2 text-gray-200"></i><br>Sin registros.</div>';
                await loadDatabases();
                showToast(`Formateo completo de "${target.nombre}" finalizado. Sistema restablecido a cero.`, "success", 5000);
            } catch(e) {
                showToast("Error durante el formateo completo", "error");
            }
        }
    }
});

// =========================================================================
// ACCESO CONFIDENCIAL Y DISCRETO A SUPER-ADMINISTRADOR
// =========================================================================
let crownClickCount = 0;
let crownClickTimeout: any = null;

export const triggerSecretSuperAdminAccess = () => {
    crownClickCount++;
    if (crownClickCount === 1) {
        crownClickTimeout = setTimeout(() => {
            crownClickCount = 0;
        }, 1500); // Ventana de 1.5 segundos para 3 clics
    } else if (crownClickCount >= 3) {
        clearTimeout(crownClickTimeout);
        crownClickCount = 0;
        playBeep('success');
        if (isSuperAdminAuthenticated) {
            openSuperAdminModal();
            showToast('Panel Super-Administrador Activo', 'success', 2500);
        } else {
            openSuperAdminLogin();
        }
    }
};

// Acceso secreto con 3 clics en el logo superior
document.getElementById('header-logo-container')?.addEventListener('click', (e) => {
    e.stopPropagation();
    triggerSecretSuperAdminAccess();
});
document.getElementById('header-logo')?.addEventListener('click', (e) => {
    e.stopPropagation();
    triggerSecretSuperAdminAccess();
});
document.getElementById('header-inst-name')?.addEventListener('click', (e) => {
    e.stopPropagation();
    triggerSecretSuperAdminAccess();
});

// Acceso secreto mediante coronita diminuta en los pies de página
document.querySelectorAll('.secret-crown-superadmin').forEach(el => {
    el.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        triggerSecretSuperAdminAccess();
    });
});

document.getElementById('btn-close-sa-login-x')?.addEventListener('click', () => closeSuperAdminLogin());
document.getElementById('btn-sa-login-cancel')?.addEventListener('click', () => closeSuperAdminLogin());

document.getElementById('btn-sa-login-submit')?.addEventListener('click', () => {
    const pin = (document.getElementById('sa-login-pin') as HTMLInputElement)?.value.trim();
    if (pin === superAdminPin || pin === 'superadmin' || pin === 'espatodo777' || pin === 'admin777') {
        isSuperAdminAuthenticated = true;
        closeSuperAdminLogin();
        openSuperAdminModal();
        showToast('Acceso autorizado Super-Administrador (www.espatodo.com)', 'success');
    } else {
        showToast('Clave de Super-Administrador incorrecta', 'error');
    }
});

// Cambio en selector de institución dentro del Super-Admin
document.getElementById('sa-inst-select')?.addEventListener('change', (e: any) => {
    const targetId = e.target.value;
    loadInstitutionIntoSuperAdmin(targetId);
    showToast(`Cargada configuración de: ${e.target.options[e.target.selectedIndex].text}`, 'info', 2000);
});

// Agregar Nueva Institución o Cliente desde Super-Admin (Abre modal elegante y funcional sin bloqueos de prompt)
export const openNewInstitutionModal = () => {
    const modal = document.getElementById('modal-new-institution');
    const nameInput = document.getElementById('modal-new-inst-name') as HTMLInputElement;
    const docInput = document.getElementById('modal-new-inst-doc') as HTMLInputElement;
    const claveInput = document.getElementById('modal-new-inst-clave') as HTMLInputElement;
    const rolSelect = document.getElementById('modal-new-inst-rol') as HTMLSelectElement;

    if (nameInput) nameInput.value = '';
    if (docInput) docInput.value = '';
    if (claveInput) claveInput.value = '1234';
    if (rolSelect) rolSelect.value = 'docente';

    if (modal) modal.classList.remove('hidden');
    setTimeout(() => nameInput?.focus(), 150);
};

export const closeNewInstitutionModal = () => {
    document.getElementById('modal-new-institution')?.classList.add('hidden');
};

document.getElementById('btn-sa-new-inst')?.addEventListener('click', (e) => {
    e.preventDefault();
    e.stopPropagation();
    openNewInstitutionModal();
});

document.getElementById('btn-close-new-inst-x')?.addEventListener('click', () => closeNewInstitutionModal());
document.getElementById('btn-cancel-new-inst')?.addEventListener('click', () => closeNewInstitutionModal());

document.getElementById('form-new-institution')?.addEventListener('submit', async (e) => {
    e.preventDefault();
    const nameInput = document.getElementById('modal-new-inst-name') as HTMLInputElement;
    const docInput = document.getElementById('modal-new-inst-doc') as HTMLInputElement;
    const claveInput = document.getElementById('modal-new-inst-clave') as HTMLInputElement;
    const rolSelect = document.getElementById('modal-new-inst-rol') as HTMLSelectElement;

    const nombre = nameInput?.value.trim();
    if (!nombre) {
        showToast('Ingresa un nombre para la institución o docente', 'warning');
        return;
    }

    const tDoc = docInput?.value.trim() || '';
    const tClave = claveInput?.value.trim() || tDoc || '1234';
    const tRol = rolSelect?.value || 'docente';

    const typeRadio = document.querySelector('input[name="modal_new_inst_type"]:checked') as HTMLInputElement;
    const planType = (typeRadio?.value || 'institucional') as LicenseType;
    const color1 = (document.getElementById('modal-new-inst-color1') as HTMLInputElement)?.value || '#2563eb';
    const color2 = (document.getElementById('modal-new-inst-color2') as HTMLInputElement)?.value || '#0ea5e9';

    const cleanId = nombre.toLowerCase().replace(/[^a-z0-9]/g, '_').slice(0, 20) + '_' + Date.now().toString().slice(-4);
    const newInst: InstitutionProfile = {
        id: cleanId,
        nombre: nombre,
        logo: null,
        color1: color1,
        color2: color2,
        licenciaInicio: getTodayString(),
        licenciaFin: new Date(new Date().setFullYear(new Date().getFullYear() + 1)).toISOString().split('T')[0],
        tipoPlan: planType,
        limiteUsuarios: planType === 'docente' ? 1 : 9999,
        titularNombre: nombre,
        titularDoc: tDoc,
        titularClave: tClave,
        titularRol: tRol,
        modulosHabilitados: planType === 'docente' ? ['asistencia', 'evaluacion'] : [...ALL_SYSTEM_MODULES],
        popupActivo: false,
        popupTitulo: `Bienvenido a ${nombre}`,
        popupUrl: "",
        popupDescripcion: "Información y tutorial de la plataforma."
    };

    // Crear el usuario con su rol en la nómina para inicio de sesión inmediato
    if (tDoc) {
        const staffMember: StaffRecord = {
            id: tDoc,
            nombre: nombre,
            rol: tRol,
            cargo: tRol,
            clave: tClave,
            grado: tRol === 'docente' ? 'Docente Titular' : tRol,
            isStaff: true
        };
        staffDict[tDoc] = staffMember;
        try {
            await setDoc(doc(getStaffPath(), tDoc), staffMember);
        } catch(err) {}
        try {
            localStorage.setItem('trial_staff', JSON.stringify(staffDict));
            localStorage.setItem('staff_cache', JSON.stringify(staffDict));
        } catch(err) {}
    }

    institucionesList.push(newInst);
    renderSuperAdminInstitutionsDropdown();
    const sel = document.getElementById('sa-inst-select') as HTMLSelectElement;
    if (sel) sel.value = cleanId;
    loadInstitutionIntoSuperAdmin(cleanId);
    closeNewInstitutionModal();
    showToast(`¡Cliente y usuario creados con éxito! Cédula: ${tDoc || 'Sin doc'} | Clave: ${tClave} | Rol: ${tRol.toUpperCase()}`, 'success', 5000);
});

// Eliminar Institución desde Super-Admin
document.getElementById('btn-sa-del-inst')?.addEventListener('click', () => {
    if (institucionesList.length <= 1) {
        return showToast('Debe existir al menos una institución en el sistema.', 'warning');
    }
    const sel = document.getElementById('sa-inst-select') as HTMLSelectElement;
    const curId = sel?.value || institucionData.id;
    const curInst = institucionesList.find(i => i.id === curId);
    if (!curInst) return;

    if (confirm(`¿Eliminar la institución "${curInst.nombre}" de la lista comercial?`)) {
        institucionesList = institucionesList.filter(i => i.id !== curId);
        if (institucionData.id === curId) {
            institucionData = { ...institucionesList[0] };
        }
        renderSuperAdminInstitutionsDropdown();
        loadInstitutionIntoSuperAdmin(institucionesList[0].id);
        aplicarConfiguracionUI();
        showToast('Institución eliminada de la lista.', 'info');
    }
});

// Subir logo desde Super-Admin
document.getElementById('sa-inst-logo-input')?.addEventListener('change', (e: any) => {
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
                
                const sel = document.getElementById('sa-inst-select') as HTMLSelectElement;
                const curId = sel?.value || institucionData.id;
                const curInst = institucionesList.find(i => i.id === curId) || institucionData;
                curInst.logo = canvas.toDataURL('image/png');
                showToast('Logo cargado. Presiona Guardar para aplicar.', 'success', 2500);
            };
            img.src = event.target.result;
        };
        reader.readAsDataURL(file);
    }
});

// Botones rápidos para Módulos Habilitados en Super-Admin
const setModulesSelection = (mods: string[]) => {
    document.querySelectorAll('.sa-module-toggle').forEach((cb: any) => {
        cb.checked = mods.includes(cb.value);
    });
    updateSuperAdminModulesCount();
    renderSuperAdminRolesMatrix();
};

document.getElementById('btn-sa-modules-all')?.addEventListener('click', () => {
    setModulesSelection(ALL_SYSTEM_MODULES);
    showToast('Todos los 7 módulos activados', 'info', 1500);
});

document.getElementById('btn-sa-modules-docente-basic')?.addEventListener('click', () => {
    setModulesSelection(['asistencia', 'evaluacion']);
    showToast('Plan Básico Docente: Asistencia y Notas activados', 'info', 1500);
});

document.getElementById('btn-sa-modules-ef-only')?.addEventListener('click', () => {
    setModulesSelection(['clase_ef', 'asistencia']);
    showToast('Plan Educación Física: Ed. Física y Asistencia activados', 'info', 1500);
});

document.getElementById('btn-sa-modules-porteria-only')?.addEventListener('click', () => {
    setModulesSelection(['salida']);
    showToast('Plan Portería: Solo Salidas/Portería activado', 'info', 1500);
});

document.querySelectorAll('.sa-module-toggle').forEach(el => {
    el.addEventListener('change', () => {
        updateSuperAdminModulesCount();
        renderSuperAdminRolesMatrix();
    });
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
            const yesterday = new Date(Date.now() - 24 * 60 * 60 * 1000).toISOString().split('T')[0];
            lEnd.value = yesterday;
        } else if (days >= 90000) {
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
    card.addEventListener('click', () => {
        const radio = card.querySelector('input[name="sa_plan_choice"]') as HTMLInputElement;
        if (radio) {
            selectSuperAdminPlan(radio.value as LicenseType);
        }
    });
});

document.getElementById('sa-licencia-inicio')?.addEventListener('input', () => updateSuperAdminCountdown());
document.getElementById('sa-licencia-fin')?.addEventListener('input', () => updateSuperAdminCountdown());

// Activación del selector de calendario al hacer clic sobre el recuadro o icono
const triggerDatePicker = (inputEl: HTMLInputElement | null) => {
    if (!inputEl) return;
    try {
        if (typeof (inputEl as any).showPicker === 'function') {
            (inputEl as any).showPicker();
        } else {
            inputEl.focus();
        }
    } catch (e) {
        inputEl.focus();
    }
};

document.getElementById('sa-date-box-inicio')?.addEventListener('click', (e: any) => {
    // Si no fue el input mismo, forzar activación del calendario
    const inputEl = document.getElementById('sa-licencia-inicio') as HTMLInputElement;
    if (e.target !== inputEl) {
        triggerDatePicker(inputEl);
    }
});

document.getElementById('sa-licencia-inicio')?.addEventListener('click', () => {
    const inputEl = document.getElementById('sa-licencia-inicio') as HTMLInputElement;
    triggerDatePicker(inputEl);
});

document.getElementById('btn-trigger-cal-inicio')?.addEventListener('click', (e) => {
    e.stopPropagation();
    const inputEl = document.getElementById('sa-licencia-inicio') as HTMLInputElement;
    triggerDatePicker(inputEl);
});

document.getElementById('sa-date-box-fin')?.addEventListener('click', (e: any) => {
    const inputEl = document.getElementById('sa-licencia-fin') as HTMLInputElement;
    if (e.target !== inputEl) {
        triggerDatePicker(inputEl);
    }
});

document.getElementById('sa-licencia-fin')?.addEventListener('click', () => {
    const inputEl = document.getElementById('sa-licencia-fin') as HTMLInputElement;
    triggerDatePicker(inputEl);
});

document.getElementById('btn-trigger-cal-fin')?.addEventListener('click', (e) => {
    e.stopPropagation();
    const inputEl = document.getElementById('sa-licencia-fin') as HTMLInputElement;
    triggerDatePicker(inputEl);
});

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

// =========================================================================
// NAVEGACIÓN Y LOGIN DEL PANEL ADMINISTRADOR (NORMAL)
// =========================================================================
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

// Ingreso al Panel de Administración Institucional
document.getElementById('btn-admin-login')?.addEventListener('click', () => {
    const pin = (document.getElementById('admin-pin') as HTMLInputElement)?.value.trim();
    const adminPin = (institucionData as any).adminPin || '1234';
    
    // Acepta la clave institucional, claves de admin por defecto o clave maestra si el propietario ingresa por aquí
    if (pin === adminPin || pin === '1234' || pin === 'admin123' || pin === 'profe123' || pin === superAdminPin || pin === 'superadmin' || pin === 'espatodo777' || pin === 'admin777') {
        document.getElementById('admin-login')?.classList.add('hidden');
        document.getElementById('admin-panel')?.classList.remove('hidden'); 
        document.getElementById('admin-panel')?.classList.add('flex');

        // Cargar campos del Pop-up en Administrador
        const pActive = document.getElementById('inst-popup-active') as HTMLInputElement;
        const pLabel = document.getElementById('label-popup-active');
        const pTitle = document.getElementById('inst-popup-title') as HTMLInputElement;
        const pUrl = document.getElementById('inst-popup-url') as HTMLInputElement;
        const pDesc = document.getElementById('inst-popup-desc') as HTMLTextAreaElement;

        if (pActive) {
            pActive.checked = !!institucionData.popupActivo;
            if (pLabel) pLabel.innerText = institucionData.popupActivo ? 'Activado' : 'Desactivado';
        }
        if (pTitle) pTitle.value = institucionData.popupTitulo || '';
        if (pUrl) pUrl.value = institucionData.popupUrl || '';
        if (pDesc) pDesc.value = institucionData.popupDescripcion || '';

        // Condición de Permisos por Rol: Solo si fue comprado a nivel Institucional
        const rolesCard = document.getElementById('card-admin-roles-config');
        if (rolesCard) {
            if (institucionData.tipoPlan === 'institucional') {
                rolesCard.classList.remove('hidden');
                renderRolesConfig();
            } else {
                rolesCard.classList.add('hidden');
            }
        }

        // Overlay de bloqueo de personal si fue comprado por un solo docente
        const staffOverlay = document.getElementById('overlay-staff-locked');
        if (staffOverlay) {
            if (institucionData.tipoPlan === 'docente') {
                staffOverlay.classList.remove('hidden');
                staffOverlay.classList.add('flex');
            } else {
                staffOverlay.classList.add('hidden');
                staffOverlay.classList.remove('flex');
            }
        }

        showToast('Bienvenido al Panel de Administración', 'success');
        return;
    }
    showToast('Contraseña de Administrador incorrecta', 'error');
});

// Guardar Configuración de Pop-up desde el Panel de Administración
document.getElementById('btn-save-settings')?.addEventListener('click', async () => {
    const pActive = (document.getElementById('inst-popup-active') as HTMLInputElement)?.checked;
    const pTitle = (document.getElementById('inst-popup-title') as HTMLInputElement)?.value.trim() || '';
    const pUrl = (document.getElementById('inst-popup-url') as HTMLInputElement)?.value.trim() || '';
    const pDesc = (document.getElementById('inst-popup-desc') as HTMLTextAreaElement)?.value.trim() || '';

    institucionData.popupActivo = !!pActive;
    institucionData.popupTitulo = pTitle;
    institucionData.popupUrl = pUrl;
    institucionData.popupDescripcion = pDesc;

    const idx = institucionesList.findIndex(i => i.id === institucionData.id);
    if (idx !== -1) institucionesList[idx] = { ...institucionData };

    const payload = {
        institucionActivaId: institucionData.id,
        institucionActiva: institucionData,
        instituciones: institucionesList
    };

    try {
        await setDoc(getSettingsPath(), payload);
        localStorage.setItem('institucionData', JSON.stringify(payload));
        showToast('Configuración del Pop-up guardada con éxito.', 'success');
    } catch(e) {
        localStorage.setItem('institucionData', JSON.stringify(payload));
        showToast('Guardado localmente.', 'info');
    }
});

const applyRolesUI = (staff: any) => {
    const role = (staff.rol || '').toLowerCase().trim();
    const btns = document.querySelectorAll('.mode-btn');
    
    let permRole = 'docente';
    if (role.includes('vigilan') || role.includes('porter')) permRole = 'vigilante';
    else if (role.includes('coord') || role.includes('rector') || role === 'administrador' || role.includes('admin')) permRole = 'coordinador';

    const allowedModules = rolePermissions[permRole] || ['asistencia', 'evaluacion'];
    
    // PUNTO 2: Filtrar los módulos que el Superadministrador activó (1, 4 o todos)
    const enabledInLicense = institucionData.modulosHabilitados && institucionData.modulosHabilitados.length > 0
        ? institucionData.modulosHabilitados
        : ALL_SYSTEM_MODULES;

    const effectiveAllowedModules = allowedModules.filter(m => enabledInLicense.includes(m));

    btns.forEach((b: any) => {
        const mode = b.dataset.mode;
        if (effectiveAllowedModules.includes(mode) && enabledInLicense.includes(mode)) {
            b.classList.remove('hidden');
        } else {
            b.classList.add('hidden');
        }
    });

    const iconSal = document.getElementById('icon-salida');
    const textSal = document.getElementById('text-salida');
    if (permRole === 'vigilante') { 
        if (iconSal) iconSal.className = 'fas fa-shield-alt mb-1 text-lg'; 
        if (textSal) textSal.innerText = 'Portería'; 
    } else if (permRole === 'coordinador') { 
        if (iconSal) iconSal.className = 'fas fa-sign-out-alt mb-1 text-lg'; 
        if (textSal) textSal.innerText = 'Autorizar Salida'; 
    } else { 
        if (iconSal) iconSal.className = 'fas fa-door-open mb-1 text-lg'; 
        if (textSal) textSal.innerText = 'Salidas'; 
    }

    if (effectiveAllowedModules.length > 0 && !effectiveAllowedModules.includes(appMode)) {
        appMode = effectiveAllowedModules[0];
    } else if (effectiveAllowedModules.length === 0) {
        appMode = '';
    }
    
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

    // PUNTO 4: Cargar Base PAE VIP (Excel) SOLO debe aparecer en el módulo PAE
    const btnUploadPae = document.getElementById('btn-upload-pae-staff');
    if (appMode === 'pae') {
        btnUploadPae?.classList.remove('hidden');
    } else {
        btnUploadPae?.classList.add('hidden');
    }
    
    const efBar = document.getElementById('ef-subtoolbar');
    if (appMode === 'clase_ef') { 
        efBar?.classList.remove('hidden'); 
        efBar?.classList.add('flex'); 
    } else { 
        efBar?.classList.add('hidden'); 
        efBar?.classList.remove('flex'); 
    }

    // Activación del Super-Cronómetro y Dashboard de Vueltas de Ed. Física
    const isPruebasActive = (appMode === 'clase_ef' && efPhase === 'pruebas') || appMode === 'pruebas';
    const swOverlay = document.getElementById('super-stopwatch-overlay');
    const lapsDashboard = document.getElementById('pruebas-laps-dashboard');
    if (isPruebasActive) {
        swOverlay?.classList.remove('hidden');
        swOverlay?.classList.add('flex');
        lapsDashboard?.classList.remove('hidden');
        lapsDashboard?.classList.add('flex');
        renderLapsTable();
    } else {
        swOverlay?.classList.add('hidden');
        swOverlay?.classList.remove('flex');
        lapsDashboard?.classList.add('hidden');
        lapsDashboard?.classList.remove('flex');
    }

    // Activación del Panel de Control de Salidas y Retornos
    const isSalidaActive = appMode === 'salida' || (appMode === 'clase_ef' && (efPhase === 'salida' || efPhase === 'regreso'));
    const salidasPanel = document.getElementById('salidas-control-panel');
    if (isSalidaActive) {
        salidasPanel?.classList.remove('hidden');
        salidasPanel?.classList.add('flex');
        updateActiveOutsideCount();
    } else {
        salidasPanel?.classList.add('hidden');
        salidasPanel?.classList.remove('flex');
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

    const loginBadge = document.getElementById('login-institution-badge');
    if (loginBadge) {
        if (currentAccessMode === 'modo_prueba') {
            loginBadge.innerText = 'Modo Prueba (30 Días)';
        } else {
            loginBadge.innerText = institucionData.nombre;
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
    'devolucion': '7. Entrega de Elementos',
    'pruebas': '8. Pruebas Físicas (Super-Cronómetro / Vueltas)'
};

document.querySelectorAll('.ef-phase-btn').forEach(btn => {
    btn.addEventListener('click', (e: any) => {
        const phase = e.currentTarget.dataset.efPhase; 
        efPhase = phase;
        const pLabel = document.getElementById('ef-phase-label');
        if (pLabel) pLabel.innerText = `Fase: ${efPhaseNames[phase] || phase}`;
        
        document.querySelectorAll('.ef-phase-btn').forEach((b: any) => {
            if (b.dataset.efPhase === phase) { 
                if (phase === 'pruebas') {
                    b.className = "ef-phase-btn bg-purple-600 text-white py-2 px-1 rounded text-xs font-black shadow-sm flex flex-col items-center";
                } else {
                    b.className = "ef-phase-btn bg-indigo-600 text-white py-2 px-1 rounded text-xs font-bold shadow-sm flex flex-col items-center"; 
                }
            } else { 
                if (b.dataset.efPhase === 'pruebas') {
                    b.className = "ef-phase-btn bg-white text-purple-700 border border-purple-300 py-2 px-1 rounded text-xs font-black shadow-sm flex flex-col items-center hover:bg-purple-100 transition-colors";
                } else {
                    b.className = "ef-phase-btn bg-white text-indigo-700 border border-indigo-300 py-2 px-1 rounded text-xs font-bold shadow-sm flex flex-col items-center"; 
                }
            }
        });
        showToast(`Modo Ed. Física: ${efPhaseNames[phase] || phase}`, 'info', 1500);
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
    
    // CASO UNIVERSAL DEMO: Docente Demo 1234 (Garantizado sin importar el modo)
    if (id === '1234' && (pin === '1234' || pin === 'profe123')) {
        if (currentAccessMode !== 'modo_prueba') {
            await switchAccessMode('modo_prueba');
        }
        currentStaff = DEMO_STAFF['1234'] || {
            id: '1234',
            documento: '1234',
            nombre: 'Profesor de Demostración',
            cargo: 'docente',
            rol: 'docente',
            clave: '1234',
            isStaff: true
        };
        const sDisp = document.getElementById('current-staff-display');
        if (sDisp) sDisp.innerHTML = `<i class="fas fa-user-check mr-1"></i> ${currentStaff.nombre}`;
        applyRolesUI(currentStaff); 
        switchTab('scanner'); 
        showToast('¡Bienvenido al Modo Prueba de 30 Días!', 'success', 4000);
        return;
    }

    let staff = staffDict[id];
    if (!staff && DEMO_STAFF[id]) {
        staff = DEMO_STAFF[id];
    }

    // Si no está en el diccionario en memoria, buscar si coincide con el titular de la IE activa
    if (!staff && institucionData.titularDoc && institucionData.titularDoc === id) {
        staff = {
            id: institucionData.titularDoc,
            nombre: institucionData.titularNombre || institucionData.nombre,
            rol: institucionData.titularRol || 'docente',
            cargo: institucionData.titularRol || 'docente',
            clave: institucionData.titularClave || institucionData.titularDoc,
            grado: 'Docente Titular',
            isStaff: true
        };
        staffDict[id] = staff;
    }

    // O si pertenece a alguna otra institución registrada en la lista de colegios / docentes
    if (!staff) {
        const foundInst = institucionesList.find(i => i.titularDoc && i.titularDoc === id);
        if (foundInst) {
            institucionData = { ...foundInst };
            aplicarConfiguracionUI();
            staff = {
                id: foundInst.titularDoc!,
                nombre: foundInst.titularNombre || foundInst.nombre,
                rol: foundInst.titularRol || 'docente',
                cargo: foundInst.titularRol || 'docente',
                clave: foundInst.titularClave || foundInst.titularDoc!,
                grado: 'Docente Titular',
                isStaff: true
            };
            staffDict[id] = staff;
        }
    }

    const validKey = staff?.clave || staff?.id;
    if (staff && (pin === validKey || pin === (staff.clave || staff.id) || (staff.clave && pin === staff.clave))) {
        // Verificar vigencia de la licencia
        if (!checkLicenseValidity()) {
            showToast(`La licencia de ${institucionData.nombre} está inactiva o fuera de fecha (${institucionData.licenciaInicio} a ${institucionData.licenciaFin}). Comunícate con el administrador.`, 'error', 6000);
            return;
        }

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
        showToast(`¡Hola, ${staff.nombre}! Rol: ${(staff.rol || 'docente').toUpperCase()}`, 'success', 3500);
        
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

// Impresión Masiva de Códigos QR (Soporta Alumnos con QR Inteligente y Plaquetas de Inventario)
const printMassiveQRs = (itemsArray: any[], title: string, isInventory = false) => {
    if(itemsArray.length === 0) return showToast('No hay registros en esta selección para imprimir', 'warning');
    showToast('Generando ventana de impresión...', 'success'); 
    const pw = window.open('', '_blank');
    if (!pw) return;
    
    let html = `<!DOCTYPE html><html><head><title>${title}</title><script src="https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js"><\/script><style>
        body{font-family:system-ui,-apple-system,sans-serif;margin:0;padding:15px;background:#fff;}
        .grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(190px,1fr));gap:14px;}
        .card{border:2px dashed #9ca3af;padding:10px;text-align:center;display:flex;flex-direction:column;align-items:center;justify-content:space-between;page-break-inside:avoid;height:270px;background:#fafafa;border-radius:10px;box-sizing:border-box;}
        .card-inv{border:2px solid #2563eb;background:#f8fafc;}
        .inst-header{font-size:9px;font-weight:bold;color:#4b5563;text-transform:uppercase;margin-bottom:2px;letter-spacing:0.5px;}
        .name{font-weight:900;font-size:12px;text-transform:uppercase;margin-bottom:3px;min-height:30px;display:flex;align-items:center;justify-content:center;width:100%;line-height:1.15;color:#111827;}
        .desc{font-size:10px;color:#6b7280;margin-bottom:3px;max-height:24px;overflow:hidden;line-height:1.1;}
        .qr{width:130px;height:130px;margin:2px 0;background:#fff;padding:4px;border-radius:6px;box-shadow:0 1px 3px rgba(0,0,0,0.1);}
        .id{font-size:13px;font-weight:900;background:#e5e7eb;padding:3px 6px;width:100%;box-sizing:border-box;border-top:2px solid #111827;border-bottom:2px solid #111827;margin-top:2px;letter-spacing:1px;color:#111827;border-radius:3px;}
        .group{font-size:11px;text-transform:uppercase;font-weight:800;color:#1d4ed8;margin-top:2px;}
        .tag-badge{font-size:9px;font-weight:900;padding:2px 6px;border-radius:999px;background:#dbeafe;color:#1e40af;margin-top:2px;text-transform:uppercase;}
        @media print{
            body{padding:0;}
            .grid{grid-template-columns:repeat(4,1fr);gap:10px;}
            .card{border-color:#374151;height:260px;}
            @page{margin:1cm;}
        }
    </style></head><body>
    <div style="text-align:center;margin-bottom:15px;">
        <h2 style="font-family:sans-serif;text-transform:uppercase;margin:0 0 4px 0;color:#1e3a8a;font-size:18px;">${title}</h2>
        <div style="font-size:11px;color:#6b7280;">${itemsArray.length} códigos listos para impresión y lectura rápida</div>
    </div>
    <div class="grid">`;
    
    itemsArray.forEach((item, idx) => { 
        const isItem = isInventory || !!item.codigo || !!item.categoria;
        const nameStr = item.nombres ? `${item.nombres} ${item.apellidos || ''}` : (item.nombre || 'SIN NOMBRE'); 
        const groupStr = item.grado || item.categoria || item.tipo || (isItem ? 'INVENTARIO' : 'GENERAL');
        const descStr = item.caracteristicas ? `<div class="desc">${item.caracteristicas}</div>` : (item.matricula ? `<div class="desc">Matrícula: ${item.matricula}</div>` : '');
        const cardClass = isItem ? "card card-inv" : "card";
        const tagBadge = isItem ? `<span class="tag-badge">Plaqueta: ${item.ubicacion || 'Plantel'}</span>` : '';
        const idDisplay = isItem ? (item.codigo || item.id) : (item.matricula ? `DOC: ${item.id} | MAT: ${item.matricula}` : item.id);

        html += `<div class="${cardClass}">
            <div class="inst-header">${(institucionData.nombre || 'INSTITUCIÓN EDUCATIVA').toUpperCase()}</div>
            <div class="name">${nameStr}</div>
            ${descStr}
            <div class="qr" id="qr-${idx}"></div>
            <div class="id">${idDisplay}</div>
            <div class="group">${groupStr}</div>
            ${tagBadge}
        </div>`; 
    });
    
    html += `</div><script>window.onload=function(){`;
    itemsArray.forEach((item, idx) => { 
        const isItem = isInventory || !!item.codigo || !!item.categoria;
        // Si es estudiante, usar formato inteligente STD|ID|MAT|NOM|GRADO
        // Si es inventario, usar formato inteligente ITM|CODIGO|NOM|CAT|UBI
        let qrPayload = String(item.id);
        if (isItem) {
            qrPayload = generateSmartInventoryQR(item);
        } else if (item.documento || item.matricula || item.nombres) {
            qrPayload = generateSmartStudentQR(item);
        }
        // Escapar comillas dobles y diagonales inversas para inyección segura en script
        const safePayload = JSON.stringify(qrPayload);
        html += `new QRCode(document.getElementById('qr-${idx}'), { text: ${safePayload}, width: 130, height: 130, correctLevel: 0 });\n`; 
    });
    html += `setTimeout(function(){window.print();},2200);};<\/script></body></html>`;
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
    printMassiveQRs(studentsToPrint, docTitle, false);
});

// Impresión Masiva de Inventario filtrado por Categorías (DEPORTES, SISTEMAS, BIBLIOTECA, OTROS)
document.getElementById('btn-qr-items')?.addEventListener('click', () => { 
    const catSelect = document.getElementById('qr-inventory-category') as HTMLSelectElement;
    const selectedCat = catSelect?.value || 'todos';
    let itemsToPrint = Object.values(inventarioDict);

    if (selectedCat !== 'todos') {
        itemsToPrint = itemsToPrint.filter(i => (i.categoria || '').toUpperCase() === selectedCat.toUpperCase());
    }

    itemsToPrint.sort((a, b) => {
        const catA = a.categoria || '';
        const catB = b.categoria || '';
        if (catA !== catB) return catA.localeCompare(catB);
        return (a.nombre || '').localeCompare(b.nombre || '');
    });

    const catTitle = selectedCat === 'todos' ? 'Todo el Inventario' : `Categoría ${selectedCat}`;
    printMassiveQRs(itemsToPrint, `Impresión Plaquetas QR - ${catTitle}`, true);
});

// Búsqueda QR Individual (Estudiante, Inventario o Docente)
document.getElementById('btn-qr-search')?.addEventListener('click', () => {
    const qInput = document.getElementById('qr-search-input') as HTMLInputElement;
    const q = qInput?.value.trim(); 
    if(!q) return showToast('Ingresa documento, matrícula o código de elemento', 'warning');
    
    // Buscar en inventario por id o por código
    let item: any = Object.values(inventarioDict).find(i => 
        String(i.codigo || '').trim().toLowerCase() === q.toLowerCase() || 
        String(i.id || '').trim().toLowerCase() === q.toLowerCase() ||
        normalizeNameMatch(i.nombre || '').includes(normalizeNameMatch(q))
    );

    // Si no está en inventario, buscar en estudiantes
    if (!item) {
        item = studentsDict[q];
    }
    if (!item) {
        item = Object.values(studentsDict).find(s => 
            String(s.matricula || '').trim() === q || 
            String(s.documento || '').trim() === q
        );
    }
    // Si no, en personal
    if (!item) {
        item = staffDict[q] || DEMO_STAFF[q];
    }

    if(!item) return showToast('Registro no encontrado en estudiantes, inventario ni personal', 'error');
    
    const isItem = !!item.codigo || !!item.categoria;
    const line1 = isItem ? (item.nombre || 'ELEMENTO') : (item.nombres ? `${item.nombres} ${item.apellidos || ''}` : (item.nombre || 'SIN NOMBRE')); 
    const descText = isItem ? (item.caracteristicas || item.ubicacion || 'Inventario') : (item.matricula ? `Matrícula: ${item.matricula}` : '');
    const group = isItem ? `CAT: ${item.categoria || 'INVENTARIO'} (${item.ubicacion || 'Plantel'})` : (item.grado || item.rol || 'GENERAL');
    const idVal = isItem ? (item.codigo || item.id) : (item.matricula ? `DOC: ${item.id} | MAT: ${item.matricula}` : item.id);
    
    const qrName = document.getElementById('qr-ind-name');
    if (qrName) qrName.innerText = line1; 
    const qrDesc = document.getElementById('qr-ind-desc');
    if (qrDesc) qrDesc.innerText = descText;
    const qrId = document.getElementById('qr-ind-id');
    if (qrId) qrId.innerText = idVal; 
    const qrGrp = document.getElementById('qr-ind-group');
    if (qrGrp) qrGrp.innerText = group;
    
    // Generar código QR inteligente
    let qrPayload = String(item.id);
    if (isItem) {
        qrPayload = generateSmartInventoryQR(item);
    } else if (item.documento || item.matricula || item.nombres) {
        qrPayload = generateSmartStudentQR(item);
    }

    const qrImg = document.getElementById('qr-ind-img');
    if (qrImg) {
        qrImg.innerHTML = '';
        if (window.QRCode) {
            new window.QRCode(qrImg, { text: qrPayload, width: 136, height: 136, correctLevel: 0 });
        }
    }
    (qrImg as any)?.setAttribute('data-qr-payload', qrPayload);
    (qrImg as any)?.setAttribute('data-is-item', isItem ? '1' : '0');
    document.getElementById('qr-individual-result')?.classList.remove('hidden');
    showToast(`Encontrado: ${line1}`, 'success', 2000);
});

// Descarga de Imagen PNG en Alta Resolución (Universal para cualquier tamaño o impresora)
document.getElementById('btn-qr-ind-download-png')?.addEventListener('click', () => {
    const name = document.getElementById('qr-ind-name')?.innerText || 'ELEMENTO'; 
    const desc = document.getElementById('qr-ind-desc')?.innerText || '';
    const id = document.getElementById('qr-ind-id')?.innerText || 'ID'; 
    const group = document.getElementById('qr-ind-group')?.innerText || '';
    const qrImgEl = document.getElementById('qr-ind-img');
    const qrCanvas = qrImgEl?.querySelector('canvas');
    const qrImg = qrImgEl?.querySelector('img');

    const canvas = document.createElement('canvas');
    const width = 600;
    const height = 750;
    canvas.width = width;
    canvas.height = height;
    const ctx = canvas.getContext('2d');
    if (!ctx) return showToast('Error al generar imagen', 'error');

    // Fondo blanco
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, width, height);

    // Borde exterior
    ctx.lineWidth = 8;
    ctx.strokeStyle = '#1e3a8a';
    ctx.strokeRect(10, 10, width - 20, height - 20);

    // Cabecera Institución
    ctx.fillStyle = '#1e3a8a';
    ctx.fillRect(10, 10, width - 20, 60);
    ctx.fillStyle = '#ffffff';
    ctx.font = 'bold 22px Arial, sans-serif';
    ctx.textAlign = 'center';
    ctx.fillText((institucionData.nombre || 'GESTOR ACADÉMICO').toUpperCase(), width / 2, 48);

    // Nombre Elemento / Estudiante
    ctx.fillStyle = '#111827';
    ctx.font = 'bold 26px Arial, sans-serif';
    ctx.fillText(name.toUpperCase().substring(0, 36), width / 2, 120);

    // Descripción o Subtítulo
    if (desc) {
        ctx.fillStyle = '#4b5563';
        ctx.font = 'italic 18px Arial, sans-serif';
        ctx.fillText(desc.substring(0, 48), width / 2, 155);
    }

    // Dibujar Código QR en el centro
    const qrSize = 340;
    const qrX = (width - qrSize) / 2;
    const qrY = desc ? 180 : 160;
    if (qrCanvas) {
        ctx.drawImage(qrCanvas, qrX, qrY, qrSize, qrSize);
    } else if (qrImg && qrImg.src) {
        const tempImg = new Image();
        tempImg.src = qrImg.src;
        ctx.drawImage(tempImg, qrX, qrY, qrSize, qrSize);
    }

    // Recuadro ID / Código
    const idY = qrY + qrSize + 25;
    ctx.fillStyle = '#f3f4f6';
    ctx.fillRect(40, idY, width - 80, 55);
    ctx.strokeStyle = '#374151';
    ctx.lineWidth = 3;
    ctx.strokeRect(40, idY, width - 80, 55);

    ctx.fillStyle = '#111827';
    ctx.font = '900 24px monospace';
    ctx.fillText(id, width / 2, idY + 37);

    // Grupo / Categoría
    ctx.fillStyle = '#1e40af';
    ctx.font = 'bold 20px Arial, sans-serif';
    ctx.fillText(group.toUpperCase(), width / 2, idY + 95);

    // Pie
    ctx.fillStyle = '#9ca3af';
    ctx.font = '13px Arial, sans-serif';
    ctx.fillText('Gestor Académico • www.espatodo.com', width / 2, height - 25);

    // Descargar archivo PNG
    const link = document.createElement('a');
    const safeName = (id.replace(/[^a-zA-Z0-9_-]/g, '_') || 'codigo').substring(0, 25);
    link.download = `QR_${safeName}.png`;
    link.href = canvas.toDataURL('image/png');
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    showToast(`Imagen PNG descargada con éxito: QR_${safeName}.png`, 'success', 3500);
});

// Descarga de Ficha PDF lista para imprimir
document.getElementById('btn-qr-ind-download-pdf')?.addEventListener('click', () => {
    const name = document.getElementById('qr-ind-name')?.innerText || 'ELEMENTO'; 
    const desc = document.getElementById('qr-ind-desc')?.innerText || '';
    const id = document.getElementById('qr-ind-id')?.innerText || 'ID'; 
    const group = document.getElementById('qr-ind-group')?.innerText || '';
    const qrImgEl = document.getElementById('qr-ind-img');
    const qrCanvas = qrImgEl?.querySelector('canvas');

    const jspdfModule = (window as any).jspdf;
    if (!jspdfModule || !jspdfModule.jsPDF) {
        return showToast('Módulo PDF no disponible', 'error');
    }

    try {
        const doc = new jspdfModule.jsPDF({ orientation: 'portrait', unit: 'mm', format: [100, 140] });
        doc.setFillColor(255, 255, 255);
        doc.rect(0, 0, 100, 140, 'F');
        doc.setDrawColor(30, 58, 138);
        doc.setLineWidth(1);
        doc.rect(3, 3, 94, 134);

        // Cabecera
        doc.setFillColor(30, 58, 138);
        doc.rect(3, 3, 94, 12, 'F');
        doc.setTextColor(255, 255, 255);
        doc.setFontSize(9);
        doc.setFont(undefined, 'bold');
        doc.text((institucionData.nombre || 'GESTOR ACADÉMICO').toUpperCase(), 50, 11, { align: 'center' });

        // Nombre
        doc.setTextColor(17, 24, 39);
        doc.setFontSize(12);
        doc.setFont(undefined, 'bold');
        doc.text(name.toUpperCase().substring(0, 30), 50, 24, { align: 'center' });

        if (desc) {
            doc.setTextColor(75, 85, 99);
            doc.setFontSize(8);
            doc.setFont(undefined, 'italic');
            doc.text(desc.substring(0, 40), 50, 30, { align: 'center' });
        }

        // QR
        if (qrCanvas) {
            const qrData = qrCanvas.toDataURL('image/png');
            doc.addImage(qrData, 'PNG', 20, 35, 60, 60);
        }

        // Recuadro ID
        doc.setFillColor(243, 244, 246);
        doc.setDrawColor(55, 65, 81);
        doc.rect(10, 100, 80, 12, 'FD');
        doc.setTextColor(17, 24, 39);
        doc.setFontSize(11);
        doc.setFont(undefined, 'bold');
        doc.text(id, 50, 108, { align: 'center' });

        // Grupo
        doc.setTextColor(30, 64, 175);
        doc.setFontSize(9);
        doc.text(group.toUpperCase(), 50, 120, { align: 'center' });

        // Pie
        doc.setTextColor(156, 163, 175);
        doc.setFontSize(7);
        doc.text('www.espatodo.com', 50, 132, { align: 'center' });

        const safeName = (id.replace(/[^a-zA-Z0-9_-]/g, '_') || 'codigo').substring(0, 25);
        doc.save(`Ficha_QR_${safeName}.pdf`);
        showToast(`Ficha PDF descargada: Ficha_QR_${safeName}.pdf`, 'success', 3500);
    } catch (e) {
        console.error('Error generando PDF individual:', e);
        showToast('Error al generar PDF', 'error');
    }
});

document.getElementById('btn-qr-ind-print')?.addEventListener('click', () => {
    const name = document.getElementById('qr-ind-name')?.innerText || ''; 
    const desc = document.getElementById('qr-ind-desc')?.innerText || '';
    const id = document.getElementById('qr-ind-id')?.innerText || ''; 
    const group = document.getElementById('qr-ind-group')?.innerText || '';
    const qrImgEl = document.getElementById('qr-ind-img');
    const payload = qrImgEl?.getAttribute('data-qr-payload') || id;
    const safePayload = JSON.stringify(payload);

    const pw = window.open('', '_blank', 'width=380,height=520');
    if (!pw) return;
    pw.document.write(`<!DOCTYPE html><html><head><title>Impresión Individual</title><script src="https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js"><\/script><style>
        body{font-family:system-ui,-apple-system,sans-serif;margin:0;padding:12px;text-align:center;width:220px;display:flex;flex-direction:column;align-items:center;background:#fff;}
        .card{border:2px solid #1e3a8a;border-radius:8px;padding:10px;width:100%;box-sizing:border-box;}
        .inst{font-size:9px;font-weight:bold;color:#4b5563;text-transform:uppercase;margin-bottom:4px;}
        .name{font-weight:900;font-size:12px;text-transform:uppercase;margin-bottom:4px;color:#000;line-height:1.2;}
        .desc{font-size:10px;color:#4b5563;margin-bottom:4px;}
        .qr{width:140px;height:140px;margin:4px auto;}
        .id{font-size:13px;font-weight:900;background:#f3f4f6;padding:4px;margin-top:6px;width:100%;box-sizing:border-box;border-top:2px solid #000;border-bottom:2px solid #000;color:#000;}
        .group{font-size:11px;margin-top:4px;text-transform:uppercase;color:#1e40af;font-weight:bold;}
    </style></head><body>
        <div class="card">
            <div class="inst">${(institucionData.nombre || 'INSTITUCIÓN EDUCATIVA').toUpperCase()}</div>
            <div class="name">${name}</div>
            ${desc ? `<div class="desc">${desc}</div>` : ''}
            <div class="qr" id="qr-ind-print"></div>
            <div class="id">${id}</div>
            <div class="group">${group}</div>
        </div>
        <script>window.onload=function(){new QRCode(document.getElementById('qr-ind-print'), { text: ${safePayload}, width: 140, height: 140, correctLevel: 0 }); setTimeout(function(){window.print();},1000);};<\/script>
    </body></html>`);
    pw.document.close();
});

// Inicialización de la Aplicación
// ==========================================================
// ONBOARDING INTELIGENTE Y VINCULACIÓN DE INSTITUCIÓN
// ==========================================================
export const openOnboardingModal = () => {
    const modal = document.getElementById('modal-device-onboarding');
    const select = document.getElementById('select-onboarding-institution') as HTMLSelectElement;
    if (select) {
        select.innerHTML = '';
        const realInsts = institucionesList.filter(i => i.id !== 'modo_prueba');
        if (realInsts.length > 0) {
            realInsts.forEach(inst => {
                const opt = document.createElement('option');
                opt.value = inst.id;
                opt.innerText = `${inst.nombre} (${inst.tipoPlan === 'institucional' ? 'Plan Institucional' : 'Licencia Activa'})`;
                select.appendChild(opt);
            });
        } else {
            const opt = document.createElement('option');
            opt.value = 'ramon_munera';
            opt.innerText = 'I.E. Ramón Múnera Lopera (Medellín - Licencia Activa)';
            select.appendChild(opt);
        }
    }
    modal?.classList.remove('hidden');
};

export const closeOnboardingModal = () => {
    document.getElementById('modal-device-onboarding')?.classList.add('hidden');
};

const initApp = async () => {
    // Reloj y Fecha en Vivo (Actualización cada segundo en el encabezado)
    const updateHeaderClock = () => {
        const dDisplay = document.getElementById('date-display');
        const tDisplay = document.getElementById('time-display');
        const now = new Date();
        if (dDisplay) {
            dDisplay.innerText = now.toLocaleDateString('es-CO', { weekday: 'short', month: 'short', day: 'numeric' });
        }
        if (tDisplay) {
            tDisplay.innerText = now.toLocaleTimeString('es-CO', { hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: true });
        }
    };
    updateHeaderClock();
    setInterval(updateHeaderClock, 1000);
    
    // 1. Detección Inteligente de Dispositivo / Parámetros de Acceso
    const urlParams = new URLSearchParams(window.location.search);
    const instParam = (urlParams.get('inst') || urlParams.get('colegio') || urlParams.get('modo') || '').toLowerCase().trim();
    const registeredDeviceInst = localStorage.getItem('device_registered_institution');

    if (instParam === 'rm' || instParam === 'ramon_munera' || instParam === 'oficial') {
        currentAccessMode = 'ramon_munera';
    } else if (instParam === 'prueba' || instParam === 'demo' || instParam === 'trial' || instParam === 'test') {
        currentAccessMode = 'modo_prueba';
    } else if (registeredDeviceInst) {
        // Dispositivo ya reconocido y vinculado a una institución con licencia
        currentAccessMode = registeredDeviceInst === 'ramon_munera' ? 'ramon_munera' : 'institucion';
    } else {
        // Dispositivo no vinculado (Primera vez o usuario en Modo Prueba)
        currentAccessMode = 'modo_prueba';
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

    // Si el dispositivo NO tiene una institución con licencia vinculada,
    // mostrar el diálogo inteligente de bienvenida (onboarding) para preguntar
    // si desea Modo Prueba o Ingresar a su Institución:
    if (!registeredDeviceInst) {
        setTimeout(() => {
            openOnboardingModal();
        }, 500);
    } else {
        // Mostrar el Pop-up Vertical de su institución al iniciar si está habilitado
        setTimeout(() => {
            showStartupPopup(false);
        }, 600);
    }
};

// Event Listeners para Onboarding y Selector de Institución
document.getElementById('btn-onboarding-trial')?.addEventListener('click', async () => {
    closeOnboardingModal();
    if (currentAccessMode !== 'modo_prueba') {
        await switchAccessMode('modo_prueba');
    }
    showToast('Modo Prueba Universal (30 Días) activo. Usa "Demo 1234" para explorar.', 'info', 4500);
});

document.getElementById('btn-onboarding-select-inst')?.addEventListener('click', async () => {
    const select = document.getElementById('select-onboarding-institution') as HTMLSelectElement;
    const chosenInstId = select?.value || 'ramon_munera';
    localStorage.setItem('device_registered_institution', chosenInstId);
    closeOnboardingModal();
    const targetMode = chosenInstId === 'ramon_munera' ? 'ramon_munera' : 'institucion';
    await switchAccessMode(targetMode as any);
    showToast(`Dispositivo vinculado con éxito a: ${institucionData.nombre}`, 'success', 5000);
});

document.getElementById('btn-close-onboarding-modal')?.addEventListener('click', closeOnboardingModal);
document.getElementById('btn-trigger-onboarding')?.addEventListener('click', openOnboardingModal);
document.getElementById('btn-device-switch-institution')?.addEventListener('click', openOnboardingModal);

document.getElementById('btn-toggle-trial-mode')?.addEventListener('click', () => {
    const nextMode = currentAccessMode === 'ramon_munera' ? 'modo_prueba' : 'ramon_munera';
    switchAccessMode(nextMode);
});

document.getElementById('btn-header-switch-mode')?.addEventListener('click', () => {
    openOnboardingModal();
});

document.getElementById('btn-fill-demo-credentials')?.addEventListener('click', async () => {
    const sId = document.getElementById('staff-id') as HTMLInputElement;
    const sPin = document.getElementById('staff-pin') as HTMLInputElement;
    if (sId) sId.value = '1234';
    if (sPin) sPin.value = '1234';
    if (currentAccessMode !== 'modo_prueba') {
        await switchAccessMode('modo_prueba');
    }
    showToast('Modo Prueba activado: Doc: 1234 / Clave: 1234. Haz clic en "Iniciar Turno".', 'info', 4000);
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

// Event Listeners del Super-Cronómetro y Vueltas (Ed. Física)
document.getElementById('btn-sw-toggle')?.addEventListener('click', toggleStopwatch);
document.getElementById('btn-sw-reset')?.addEventListener('click', resetStopwatch);
document.getElementById('btn-export-laps-excel')?.addEventListener('click', exportLapsToExcel);
document.getElementById('btn-clear-laps-session')?.addEventListener('click', () => {
    sessionLaps = [];
    studentLapData = {};
    resetStopwatch();
    renderLapsTable();
    showToast('Sesión de pruebas reiniciada.', 'info');
});

// Event Listeners de Salidas y Retornos (Estudiantes Fuera)
document.getElementById('btn-view-outside-students')?.addEventListener('click', () => {
    document.getElementById('modal-students-outside')?.classList.remove('hidden');
    renderOutsideStudentsModal();
});
document.getElementById('btn-close-students-outside')?.addEventListener('click', () => {
    document.getElementById('modal-students-outside')?.classList.add('hidden');
});
document.getElementById('btn-refresh-outside-list')?.addEventListener('click', () => {
    renderOutsideStudentsModal();
});

initApp();
