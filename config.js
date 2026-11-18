/**
 * Configuración Global - Sistema de Gestión de Actas de Grado (IVSECALDAS)
 * Diseñado para despliegue estático en GitHub Pages + Microsoft Graph API (OneDrive)
 */

const APP_CONFIG = {
    // Parámetros de OneDrive / Microsoft Entra ID (Azure AD)
    // Reemplaza 'TU_CLIENT_ID_DE_AZURE' al registrar la app en portal.azure.com
    azure: {
        clientId: localStorage.getItem('ivse_client_id') || 'TU_CLIENT_ID_DE_AZURE',
        tenantId: localStorage.getItem('ivse_tenant_id') || 'common',
        redirectUri: window.location.origin + window.location.pathname.replace(/[^/]*$/, 'index.html'),
        scopes: ['User.Read', 'Files.ReadWrite.All']
    },
    onedrive: {
        ownerEmail: 'aarias@sedcaldas.edu.co',
        folderPath: 'IVSECALDAS',
        excelFileName: 'RegistroActasDeGrado.xlsx'
    },
    sheets: {
        usuarios: 'Usuarios',
        actas: 'Actas',
        municipio: 'Municipio',
        institucion: 'Institucion',
        graduados: 'Graduados'
    }
};

// Datos iniciales basados en la estructura real de RegistroActasDeGrado.xlsx
const INITIAL_DATA = {
    Usuarios: [
        { NOMBRE: 'ADMIN', CLAVE: 'ADMIN2026*', ROL: 'ADMINISTRADOR', ESTADO: 'ACTIVO' },
        { NOMBRE: '17001', CLAVE: '123456789*', ROL: 'OPERADOR', ESTADO: 'ACTIVO' }
    ],
    Actas: [],
    Municipio: [
        'Aguadas', 'Anserma', 'Aranzazu', 'Belalcázar', 'Chinchiná', 'Filadelfia',
        'La Dorada', 'La Merced', 'Manzanares', 'Marmato', 'Marquetalia', 'Marulanda',
        'Neira', 'Norcasia', 'Pácora', 'Palestina', 'Pensilvania', 'Riosucio',
        'Risaralda', 'Salamina', 'Samaná', 'San José', 'Supía', 'Victoria',
        'Villamaría', 'Viterbo'
    ].map(m => ({ MUNICIPIO: m })),
    Institucion: [],
    Graduados: [],
    Pdfs: []
};

/**
 * Capa de Persistencia (Microsoft Graph API + Modo Local Sincronizable)
 */
const DataService = {
    getMode() {
        return localStorage.getItem('ivse_mode') || 'LOCAL'; // 'GRAPH' o 'LOCAL'
    },

    setMode(mode) {
        localStorage.setItem('ivse_mode', mode);
    },

    getAccessToken() {
        return sessionStorage.getItem('msal_access_token') || null;
    },

    initLocalDB() {
        if (!localStorage.getItem('ivse_db')) {
            localStorage.setItem('ivse_db', JSON.stringify(INITIAL_DATA));
        }
        return JSON.parse(localStorage.getItem('ivse_db'));
    },

    saveLocalDB(db) {
        localStorage.setItem('ivse_db', JSON.stringify(db));
    },

    // Obtener todas las tablas (desde OneDrive vía Graph API o desde caché local)
    async loadAllData() {
        const mode = this.getMode();
        const token = this.getAccessToken();

        if (mode === 'GRAPH' && token) {
            try {
                return await this.fetchFromOneDrive(token);
            } catch (err) {
                console.warn('Error leyendo OneDrive, usando respaldo local:', err);
                return this.initLocalDB();
            }
        }
        return this.initLocalDB();
    },

    // Leer el archivo RegistroActasDeGrado.xlsx directamente desde OneDrive usando Microsoft Graph
    async fetchFromOneDrive(token) {
        const folder = APP_CONFIG.onedrive.folderPath;
        const file = APP_CONFIG.onedrive.excelFileName;
        const baseUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/${folder}/${file}:/workbook/worksheets`;

        const headers = {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json'
        };

        const db = this.initLocalDB();
        const sheetNames = ['Usuarios', 'Actas', 'Municipio', 'Institucion', 'Graduados'];

        for (const name of sheetNames) {
            const res = await fetch(`${baseUrl}('${name}')/usedRange`, { headers });
            if (res.ok) {
                const json = await res.json();
                const values = json.values || [];
                if (values.length > 0) {
                    const cols = values[0];
                    db[name] = values.slice(1).map(row => {
                        const obj = {};
                        cols.forEach((col, idx) => {
                            obj[String(col).trim()] = row[idx] !== undefined && row[idx] !== null ? String(row[idx]) : '';
                        });
                        return obj;
                    });
                } else {
                    db[name] = [];
                }
            }
        }
        this.saveLocalDB(db);
        return db;
    },

    // Agregar fila a una hoja de Excel en OneDrive (Graph API)
    async appendRowToOneDrive(sheetName, rowArray) {
        const token = this.getAccessToken();
        if (this.getMode() !== 'GRAPH' || !token) return false;

        const folder = APP_CONFIG.onedrive.folderPath;
        const file = APP_CONFIG.onedrive.excelFileName;
        const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${folder}/${file}:/workbook/worksheets('${sheetName}')/tables/Tabla_${sheetName}/rows/add`;
        
        // Alternativa directa por rango si no hay tabla estructurada en Excel
        const rangeRes = await fetch(
            `https://graph.microsoft.com/v1.0/me/drive/root:/${folder}/${file}:/workbook/worksheets('${sheetName}')/usedRange`,
            { headers: { 'Authorization': `Bearer ${token}` } }
        );
        if (!rangeRes.ok) return false;
        const rangeData = await rangeRes.json();
        const nextRow = (rangeData.rowCount || 1) + 1;
        const endCol = String.fromCharCode(64 + rowArray.length);
        const updateUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/${folder}/${file}:/workbook/worksheets('${sheetName}')/range(address='A${nextRow}:${endCol}${nextRow}')`;

        const res = await fetch(updateUrl, {
            method: 'PATCH',
            headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ values: [rowArray] })
        });
        return res.ok;
    },

    // Sobrescribir hoja completa en OneDrive (útil para ediciones/eliminaciones del Administrador)
    async syncSheetToOneDrive(sheetName, headersList, rowsObjects) {
        const token = this.getAccessToken();
        if (this.getMode() !== 'GRAPH' || !token) return false;

        const folder = APP_CONFIG.onedrive.folderPath;
        const file = APP_CONFIG.onedrive.excelFileName;
        const base = `https://graph.microsoft.com/v1.0/me/drive/root:/${folder}/${file}:/workbook/worksheets('${sheetName}')`;

        // Limpiar contenido previo
        await fetch(`${base}/range(address='A1:Z2000')/clear`, {
            method: 'POST',
            headers: { 'Authorization': `Bearer ${token}` }
        });

        const matrix = [headersList];
        rowsObjects.forEach(item => {
            matrix.push(headersList.map(h => item[h] || ''));
        });

        const endCol = String.fromCharCode(64 + headersList.length);
        const endRow = matrix.length;
        const res = await fetch(`${base}/range(address='A1:${endCol}${endRow}')`, {
            method: 'PATCH',
            headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ values: matrix })
        });
        return res.ok;
    },

    // Listar archivos PDF en la carpeta IVSECALDAS de OneDrive
    async listOneDrivePDFs() {
        const token = this.getAccessToken();
        if (this.getMode() === 'GRAPH' && token) {
            const folder = APP_CONFIG.onedrive.folderPath;
            const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${folder}:/children?$filter=endswith(name,'.pdf')`;
            const res = await fetch(url, {
                headers: { 'Authorization': `Bearer ${token}` }
            });
            if (res.ok) {
                const data = await res.json();
                return (data.value || []).filter(f => f.name.toLowerCase().endsWith('.pdf')).map(f => ({
                    id: f.id,
                    name: f.name,
                    size: f.size,
                    lastModified: f.lastModifiedDateTime,
                    downloadUrl: f['@microsoft.graph.downloadUrl'],
                    webUrl: f.webUrl,
                    source: 'OneDrive'
                }));
            }
        }
        const db = this.initLocalDB();
        return db.Pdfs || [];
    },

    // Subir PDF a la carpeta IVSECALDAS de OneDrive
    async uploadPDFToOneDrive(file, uploadedBy) {
        const token = this.getAccessToken();
        if (this.getMode() === 'GRAPH' && token) {
            const folder = APP_CONFIG.onedrive.folderPath;
            const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${folder}/${encodeURIComponent(file.name)}:/content`;
            const res = await fetch(url, {
                method: 'PUT',
                headers: {
                    'Authorization': `Bearer ${token}`,
                    'Content-Type': 'application/pdf'
                },
                body: file
            });
            if (res.ok) {
                return await res.json();
            }
            throw new Error('No se pudo subir el archivo PDF a OneDrive.');
        } else {
            // Almacenamiento en modo local (Base64 en IndexedDB / LocalStorage para pruebas)
            return new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onload = () => {
                    const db = this.initLocalDB();
                    if (!db.Pdfs) db.Pdfs = [];
                    const record = {
                        id: 'pdf_' + Date.now(),
                        name: file.name,
                        size: file.size,
                        lastModified: new Date().toISOString(),
                        uploadedBy: uploadedBy,
                        dataUrl: reader.result,
                        source: 'Local / GitHub Demo'
                    };
                    db.Pdfs.unshift(record);
                    this.saveLocalDB(db);
                    resolve(record);
                };
                reader.onerror = reject;
                reader.readAsDataURL(file);
            });
        }
    }
};

// Control de Sesión
const AuthSession = {
    getUser() {
        const raw = sessionStorage.getItem('ivse_user');
        return raw ? JSON.parse(raw) : null;
    },
    setUser(userObj) {
        sessionStorage.setItem('ivse_user', JSON.stringify(userObj));
    },
    logout() {
        sessionStorage.removeItem('ivse_user');
        sessionStorage.removeItem('msal_access_token');
        window.location.href = 'index.html';
    },
    requireAuth(allowedRoles = ['ADMINISTRADOR', 'OPERADOR']) {
        const user = this.getUser();
        if (!user) {
            window.location.href = 'index.html';
            return null;
        }
        if (!allowedRoles.includes(user.ROL)) {
            window.location.href = user.ROL === 'ADMINISTRADOR' ? 'admin.html' : 'operador.html';
            return null;
        }
        return user;
    }
};