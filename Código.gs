/**
 * ASISTENCIA PRO v9.0 - CORE ERP ENTERPRISE (PRODUCTOS QUÍMICOS Y SAAS PRE-OPTIMIZADO)
 * Lógica modular del servidor, control RBAC y aprovisionamiento automático.
 */

let _ss = null;
function getSS() {
  if (!_ss) {
    _ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  return _ss;
}
const FOLDER_ID = '12Z67EFRZXU9IFIIW7z0s189B-fHb5zLe';

// Helper estándar para inyectar sub-archivos HTML modularmente
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function doGet(e) {
  try {
    if (e && e.parameter && e.parameter.action === 'getData') {
      const token = e.parameter.token;
      if (!token) throw new Error('Token requerido');
      const user = validateToken(token);
      const data = getMasterData(user);
      data._meta = { newToken: user.newToken };
      return response(data);
    }
    
    const template = HtmlService.createTemplateFromFile('index');
    template.scriptUrl = ScriptApp.getService().getUrl(); // Inyección segura evaluada en servidor
    
    return template.evaluate()
      .setTitle('ERP SPARTAN')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return response({ status: 'error', message: err.message });
  }
}

function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents);
    const result = processPostAction(params);
    return response(result);
  } catch (err) {
    return response({ status: 'error', message: err.message });
  }
}

function processPostAction(params) {
  const { action, token } = params;
  if (action === 'login') {
    return loginUser(params.usuario, params.password);
  }

  const user = validateToken(token);
  if (!hasAccess(user.rol, params.table, action)) {
    throw new Error('Permiso denegado: el rol ' + user.rol + ' no tiene acceso a ' + action + ' en la tabla ' + params.table);
  }

  if (params.table === 'Empleados' && user.rol === 'EMPLEADO') {
    const empId = params.id || params.data.id;
    if (empId !== user.empleadoId) {
      throw new Error('No puedes modificar otro empleado');
    }
  }

  if (params.table === 'Logs' && action !== 'read') {
    throw new Error('No está permitido escribir directamente en la tabla Logs');
  }

  let responseData = {};

  switch (action) {
    case 'create': {
      const resCreate = createRecord(params.table, params.data);
      if (params.table === 'Accesos') {
        writeLog('CREACION_ACCESO', `Se creó el acceso para el usuario: ${params.data.usuario}`, token);
      } else if (params.table === 'Clientes' || params.table === 'Empresas') {
        writeLog(`CREACION_${params.table.toUpperCase()}`, `Se creó el registro con ID: ${resCreate.id}`, token);
      } else {
        writeLog(`CREACION_${params.table.toUpperCase()}`, `Se creó registro con ID: ${resCreate.id}`, token);
      }
      responseData = resCreate;
      break;
    }
    case 'update': {
      const resUpdate = updateRecord(params.table, params.id, params.data);
      if (params.table === 'Accesos') {
        writeLog('MODIFICACION_ACCESO', `Se modificó el acceso del usuario ID: ${params.id}. Campos: ${Object.keys(params.data).join(', ')}`, token);
      } else if (params.table === 'Clientes' || params.table === 'Empresas') {
        writeLog(`MODIFICACION_${params.table.toUpperCase()}`, `Se modificó el registro ID: ${params.id}. Campos: ${Object.keys(params.data).join(', ')}`, token);
      } else {
        writeLog(`MODIFICACION_${params.table.toUpperCase()}`, `Se modificó el registro ID: ${params.id}`, token);
      }
      responseData = resUpdate;
      break;
    }
    case 'delete': {
      const resDelete = deleteRecord(params.table, params.id);
      if (params.table === 'Accesos') {
        writeLog('ELIMINACION_ACCESO', `Se eliminó el acceso ID: ${params.id}`, token);
      } else if (params.table === 'Clientes' || params.table === 'Empresas') {
        writeLog(`ELIMINACION_${params.table.toUpperCase()}`, `Se eliminó el registro ID: ${params.id}`, token);
      } else {
        writeLog(`ELIMINACION_${params.table.toUpperCase()}`, `Se eliminó el registro ID: ${params.id}`, token);
      }
      responseData = resDelete;
      break;
    }
    default: throw new Error('Acción no reconocida');
  }

  responseData.newToken = user.newToken;
  if (!responseData.status) responseData.status = 'success';
  return responseData;
}

function hasAccess(rol, resource, action) {
  const normRol = String(rol).trim().toUpperCase();
  const normResource = String(resource).trim().toLowerCase();
  const normAction = String(action).trim().toLowerCase();

  if (normRol === 'SUPERADMIN' || normRol === 'ADMIN') return true;

  const permissions = {
    'SUPERVISOR': {
      'asistencias': ['read', 'create', 'update'],
      'empleados': ['read'],
      'empresas': ['read'],
      'clientes': ['read', 'create', 'update'],
      'sedes': ['read'],
      'calendario': ['read', 'create', 'update', 'delete'],
      'productos': ['read'],
      'lotes': ['read'],
      'crmhistorial': ['read', 'create', 'update'],
      'pipeline': ['read', 'create', 'update'],
      'documentos': ['read', 'create'],
      'muestras': ['read', 'create', 'update'],
      'incidencias': ['read', 'create', 'update']
    },
    'EMPLEADO': {
      'asistencias': ['read', 'create', 'update'],
      'empleados': ['read', 'update'],
      'empresas': ['read'],
      'clientes': ['read'],
      'sedes': ['read'],
      'calendario': ['read', 'create', 'update'],
      'productos': ['read'],
      'lotes': ['read'],
      'crmhistorial': ['read', 'create'],
      'pipeline': ['read', 'create', 'update'],
      'documentos': ['read'],
      'muestras': ['read', 'create'],
      'incidencias': ['read', 'create']
    },
    'CLIENTE': {
      'asistencias': ['read'],
      'clientes': ['read'],
      'sedes': ['read'],
      'crmhistorial': ['read'],
      'documentos': ['read'],
      'productos': ['read'],
      'incidencias': ['read', 'create'],
      'muestras': ['read']
    }
  };

  if (!permissions[normRol]) return false;
  if (!permissions[normRol][normResource]) return false;
  return permissions[normRol][normResource].indexOf(normAction) > -1;
}

function filterRowsByRole(tableName, dataRows, headers, user) {
  const normTable = tableName.toLowerCase();
  const rol = user.rol;
  const empId = String(user.empleadoId).trim();

  if (rol === 'SUPERADMIN' || rol === 'ADMIN') return dataRows;

  const empIdIdx = headers.indexOf('empleadoId');
  const cliIdIdx = headers.indexOf('clienteId');
  const idIdx = headers.indexOf('id');
  const empAsgIdx = headers.indexOf('empleadosAsignados');

  if (rol === 'SUPERVISOR') {
    if (normTable === 'accesos' || normTable === 'logs') return [];
    return dataRows;
  }

  if (rol === 'EMPLEADO') {
    if (normTable === 'accesos' || normTable === 'logs') return [];
    if (normTable === 'empleados') {
      if (idIdx === -1) return dataRows;
      return dataRows.filter(r => String(r[idIdx]).trim() === empId);
    }
    if (normTable === 'asistencias') {
      if (empIdIdx === -1) return dataRows;
      return dataRows.filter(r => String(r[empIdIdx]).trim() === empId);
    }
    if (normTable === 'calendario') {
      if (empIdIdx === -1) return dataRows;
      return dataRows.filter(r => String(r[empIdIdx]).trim() === empId);
    }
    if (normTable === 'crmhistorial') {
      if (empIdIdx === -1) return dataRows;
      return dataRows.filter(r => String(r[empIdIdx]).trim() === empId);
    }
    if (normTable === 'muestras') {
      if (empIdIdx === -1) return dataRows;
      return dataRows.filter(r => String(r[empIdIdx]).trim() === empId);
    }
    if (normTable === 'incidencias') {
      if (empIdIdx === -1) return dataRows;
      return dataRows.filter(r => String(r[empIdIdx]).trim() === empId);
    }

    let assignedClientIds = [];
    try {
      const clientRows = getCachedSheetValues('Clientes');
      if (clientRows && clientRows.length > 1) {
        const cHeaders = clientRows[0].map(h => toCamelCase(h));
        const cIdIdx = cHeaders.indexOf('id');
        const cEmpAsgIdx = cHeaders.indexOf('empleadosAsignados');
        if (cIdIdx !== -1 && cEmpAsgIdx !== -1) {
          clientRows.slice(1).forEach(r => {
            const assigned = r[cEmpAsgIdx] ? String(r[cEmpAsgIdx]).split(',').map(id => id.trim()) : [];
            if (assigned.indexOf(empId) > -1) {
              assignedClientIds.push(String(r[cIdIdx]).trim());
            }
          });
        }
      }
    } catch(e) {}

    if (normTable === 'clientes') {
      if (idIdx === -1 || empAsgIdx === -1) return dataRows;
      return dataRows.filter(r => {
        const assigned = r[empAsgIdx] ? String(r[empAsgIdx]).split(',').map(id => id.trim()) : [];
        return assigned.indexOf(empId) > -1;
      });
    }
    if (normTable === 'empresas') {
      if (idIdx === -1 || empAsgIdx === -1) return dataRows;
      return dataRows.filter(r => {
        const assigned = r[empAsgIdx] ? String(r[empAsgIdx]).split(',').map(id => id.trim()) : [];
        return assigned.indexOf(empId) > -1;
      });
    }
    if (normTable === 'sedes') {
      let authSedes = [];
      try {
        const empRows = getCachedSheetValues('Empleados');
        if (empRows && empRows.length > 1) {
          const eHeaders = empRows[0].map(h => toCamelCase(h));
          const eIdIdx = eHeaders.indexOf('id');
          const eSedesIdx = eHeaders.indexOf('sedesAutorizadas');
          if (eIdIdx !== -1 && eSedesIdx !== -1) {
            const myRow = empRows.slice(1).find(r => String(r[eIdIdx]).trim() === empId);
            if (myRow && myRow[eSedesIdx]) {
              authSedes = String(myRow[eSedesIdx]).split(',').map(id => id.trim());
            }
          }
        }
      } catch(e) {}
      if (idIdx === -1) return dataRows;
      return dataRows.filter(r => authSedes.indexOf(String(r[idIdx]).trim()) > -1);
    }
    if (normTable === 'pipeline' || normTable === 'documentos') {
      if (cliIdIdx === -1) return dataRows;
      return dataRows.filter(r => assignedClientIds.indexOf(String(r[cliIdIdx]).trim()) > -1);
    }
    return dataRows;
  }

  if (rol === 'CLIENTE') {
    const clientLinkId = empId;
    if (normTable === 'empleados' || normTable === 'accesos' || normTable === 'logs') return [];
    if (normTable === 'clientes') {
      if (idIdx === -1) return dataRows;
      return dataRows.filter(r => String(r[idIdx]).trim() === clientLinkId);
    }
    if (normTable === 'empresas') {
      let clientCompanyId = '';
      try {
        const clientRows = getCachedSheetValues('Clientes');
        if (clientRows && clientRows.length > 1) {
          const cHeaders = clientRows[0].map(h => toCamelCase(h));
          const cIdIdx = cHeaders.indexOf('id');
          const cCompIdx = cHeaders.indexOf('empresaId');
          if (cIdIdx !== -1 && cCompIdx !== -1) {
            const myRow = clientRows.slice(1).find(r => String(r[cIdIdx]).trim() === clientLinkId);
            if (myRow) clientCompanyId = String(myRow[cCompIdx]).trim();
          }
        }
      } catch(e) {}
      if (idIdx === -1) return dataRows;
      return dataRows.filter(r => String(r[idIdx]).trim() === clientCompanyId);
    }
    if (normTable === 'sedes') {
      if (cliIdIdx === -1) return dataRows;
      return dataRows.filter(r => String(r[cliIdIdx]).trim() === clientLinkId);
    }
    if (normTable === 'asistencias' || normTable === 'calendario' || normTable === 'crmhistorial' || normTable === 'pipeline' || normTable === 'documentos' || normTable === 'muestras' || normTable === 'incidencias') {
      if (cliIdIdx === -1) return dataRows;
      return dataRows.filter(r => String(r[cliIdIdx]).trim() === clientLinkId);
    }
    return dataRows;
  }
  return dataRows;
}

function getCachedSheetValues(sheetName) {
  const SS = getSS();
  const cacheKey = 'sheet_' + sheetName;
  try {
    const cachedData = getLargeCache(cacheKey);
    if (cachedData) return cachedData;
  } catch(e) {
    Logger.log('Error al leer caché para ' + sheetName + ': ' + e.message);
  }
  
  const sheet = SS.getSheetByName(sheetName);
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  try {
    const expiration = (sheetName === 'Accesos') ? 60 : 3600;
    setLargeCache(cacheKey, values, expiration);
  } catch(e) {
    Logger.log('Error al escribir caché para ' + sheetName + ': ' + e.message);
  }
  return values;
}

function invalidateSheetCache(sheetName) {
  const cacheKey = 'sheet_' + sheetName;
  try {
    removeLargeCache(cacheKey);
  } catch(e) {
    Logger.log('Error al invalidar caché para ' + sheetName + ': ' + e.message);
  }
}

function setLargeCache(key, data, expirationSec) {
  const cache = CacheService.getScriptCache();
  const json = JSON.stringify(data);
  const chunkSize = 90 * 1024;
  const numChunks = Math.ceil(json.length / chunkSize);
  cache.put(key + '_meta', JSON.stringify({ chunks: numChunks }), expirationSec);
  for (let i = 0; i < numChunks; i++) {
    cache.put(key + '_chunk_' + i, json.substring(i * chunkSize, (i + 1) * chunkSize), expirationSec);
  }
}

function getLargeCache(key) {
  const cache = CacheService.getScriptCache();
  const metaStr = cache.get(key + '_meta');
  if (!metaStr) return null;
  try {
    const meta = JSON.parse(metaStr);
    let json = '';
    for (let i = 0; i < meta.chunks; i++) {
      const chunk = cache.get(key + '_chunk_' + i);
      if (!chunk) return null;
      json += chunk;
    }
    return JSON.parse(json);
  } catch(e) {
    return null;
  }
}

function removeLargeCache(key) {
  const cache = CacheService.getScriptCache();
  const metaStr = cache.get(key + '_meta');
  if (metaStr) {
    try {
      const meta = JSON.parse(metaStr);
      cache.remove(key + '_meta');
      for (let i = 0; i < meta.chunks; i++) {
        cache.remove(key + '_chunk_' + i);
      }
    } catch(e) {}
  }
}

function loginUser(usuario, password) {
  const sheet = getSheetAndCreateIfMissing('Accesos');
  if (!sheet) throw new Error('Hoja "Accesos" no encontrada');
  const data = getCachedSheetValues('Accesos');
  if (data.length < 2) throw new Error('No hay usuarios registrados');

  const headers = data[0].map(h => toCamelCase(h));
  const rows = data.slice(1);
  const uIdx = headers.indexOf('usuario');
  const pIdx = headers.indexOf('passwordHash');
  const sIdx = headers.indexOf('estado');
  const empIdx = headers.indexOf('empleadoId');
  const rolIdx = headers.indexOf('rol');
  const tokenIdx = headers.indexOf('token');
  const expIdx = headers.indexOf('fechaExpiracion');
  const uAccIdx = headers.indexOf('ultimoAcceso');
  if (uIdx === -1 || pIdx === -1) throw new Error('Estructura de hoja "Accesos" incorrecta');

  const hashedPass = computeHash(password);
  const rowIndex = rows.findIndex(r => r[uIdx] === usuario && String(r[pIdx]) === hashedPass);
  if (rowIndex === -1) return { status: 'error', message: 'Usuario o contraseña incorrectos' };

  const userRow = rows[rowIndex];
  if (userRow[sIdx] !== 'ACTIVO') return { status: 'error', message: 'Cuenta desactivada' };
  
  const token = Utilities.getUuid();
  const expiration = new Date(Date.now() + 2 * 60 * 60 * 1000);
  if (tokenIdx !== -1) sheet.getRange(rowIndex + 2, tokenIdx + 1).setValue(token);
  if (expIdx !== -1) sheet.getRange(rowIndex + 2, expIdx + 1).setValue(expiration);
  if (uAccIdx !== -1) sheet.getRange(rowIndex + 2, uAccIdx + 1).setValue(new Date());
  
  invalidateSheetCache('Accesos');

  let nombre = 'Administrador';
  const empId = userRow[empIdx];
  if (empId && empId !== '0') {
    const empSheet = getSheetAndCreateIfMissing('Empleados');
    if (empSheet) {
      const empData = getCachedSheetValues('Empleados');
      const empHeaders = empData[0].map(h => toCamelCase(h));
      const empRows = empData.slice(1);
      const empRow = empRows.find(r => r[empHeaders.indexOf('id')] == empId);
      if (empRow) nombre = empRow[empHeaders.indexOf('nombre')] || 'Empleado';
    }
  }

  const rolVal = (userRow[rolIdx] || 'EMPLEADO').toString().trim().toUpperCase();
  const userObj = { rol: rolVal, empleadoId: empId || '0' };
  const db = getMasterData(userObj);
  
  return {
    status: 'success',
    usuario: usuario,
    nombre: nombre,
    role: rolVal,
    empId: empId || '0',
    token: token,
    db: db
  };
}

function validateToken(token) {
  if (!token) throw new Error('Token no proporcionado');
  const sheet = getSheetAndCreateIfMissing('Accesos');
  if (!sheet) throw new Error('Hoja "Accesos" no encontrada');
  
  const data = getCachedSheetValues('Accesos');
  if (data.length < 2) throw new Error('No hay usuarios registrados');
  const headers = data[0].map(h => toCamelCase(h));
  const rows = data.slice(1);
  const tIdx = headers.indexOf('token');
  const sIdx = headers.indexOf('estado');
  const expIdx = headers.indexOf('fechaExpiracion');
  const empIdx = headers.indexOf('empleadoId');
  const rolIdx = headers.indexOf('rol');
  const uAccIdx = headers.indexOf('ultimoAcceso');
  if (tIdx === -1) throw new Error('Columna "token" no encontrada');
  const userRow = rows.find(r => r[tIdx] === token);
  if (!userRow) throw new Error('Token inválido');
  if (userRow[sIdx] !== 'ACTIVO') throw new Error('Cuenta desactivada');

  if (expIdx !== -1) {
    const expDate = new Date(userRow[expIdx]);
    if (expDate < new Date()) {
      sheet.getRange(rows.indexOf(userRow) + 2, tIdx + 1).setValue('');
      invalidateSheetCache('Accesos');
      throw new Error('Sesión expirada, inicie sesión nuevamente');
    }
  }

  let rotatedToken = token;
  const lastAccessVal = userRow[uAccIdx];
  const lastAccess = lastAccessVal ? new Date(lastAccessVal) : new Date(0);
  const now = new Date();
  const rowIndex = rows.indexOf(userRow) + 2;

  if (now.getTime() - lastAccess.getTime() > 5 * 60 * 1000) {
    rotatedToken = Utilities.getUuid();
    sheet.getRange(rowIndex, tIdx + 1).setValue(rotatedToken);
    sheet.getRange(rowIndex, expIdx + 1).setValue(new Date(Date.now() + 2 * 60 * 60 * 1000));
    sheet.getRange(rowIndex, uAccIdx + 1).setValue(now);
    invalidateSheetCache('Accesos');
  } else {
    sheet.getRange(rowIndex, uAccIdx + 1).setValue(now);
    invalidateSheetCache('Accesos');
  }

  return {
    empleadoId: userRow[empIdx] || '0',
    rol: (userRow[rolIdx] || 'EMPLEADO').toString().trim().toUpperCase(),
    newToken: rotatedToken
  };
}

function getMasterData(user) {
  const SS = getSS();
  const tables = ['Asistencias', 'Empleados', 'Empresas', 'Clientes', 'Calendario', 'Accesos', 'Sedes', 'Logs', 'Productos', 'Lotes', 'CrmHistorial', 'Pipeline', 'Documentos', 'Muestras', 'Incidencias'];
  const db = {};
  const limit = 1000;
  const timezone = SS.getSpreadsheetTimeZone();
  tables.forEach(t => {
    let sheet = getSheetAndCreateIfMissing(t);
    try {
      let changed = false;
      if (t === 'Empresas') {
        let rowsTemp = sheet.getDataRange().getValues();
        let headerRow = rowsTemp[0];
        let normalizedHeaders = headerRow.map(h => toCamelCase(h));
        if (!normalizedHeaders.includes('sede')) {
          sheet.getRange(1, headerRow.length + 1).setValue('Sede');
          changed = true;
        }
        headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        normalizedHeaders = headerRow.map(h => toCamelCase(h));
        if (!normalizedHeaders.includes('empleadosAsignados')) {
          sheet.getRange(1, headerRow.length + 1).setValue('Empleados Asignados');
          changed = true;
        }
      } else if (t === 'Clientes') {
        let rowsTemp = sheet.getDataRange().getValues();
        let headerRow = rowsTemp[0];
        let normalizedHeaders = headerRow.map(h => toCamelCase(h));
        if (!normalizedHeaders.includes('contacto')) {
          sheet.getRange(1, headerRow.length + 1).setValue('Contacto');
          changed = true;
        }
        headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        normalizedHeaders = headerRow.map(h => toCamelCase(h));
        if (!normalizedHeaders.includes('estado')) {
          sheet.getRange(1, headerRow.length + 1).setValue('Estado');
          changed = true;
        }
        headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        normalizedHeaders = headerRow.map(h => toCamelCase(h));
        if (!normalizedHeaders.includes('empleadosAsignados')) {
          sheet.getRange(1, headerRow.length + 1).setValue('Empleados Asignados');
          changed = true;
        }
      } else if (t === 'Asistencias') {
        let rowsTemp = sheet.getDataRange().getValues();
        const headerRow = rowsTemp[0];
        const normalizedHeaders = headerRow.map(h => toCamelCase(h));
        if (!normalizedHeaders.includes('sedeId')) {
          sheet.getRange(1, headerRow.length + 1).setValue('Sede Id');
          changed = true;
        }
      } else if (t === 'Empleados') {
        let rowsTemp = sheet.getDataRange().getValues();
        const headerRow = rowsTemp[0];
        const normalizedHeaders = headerRow.map(h => toCamelCase(h));
        if (!normalizedHeaders.includes('sedesAutorizadas')) {
          sheet.getRange(1, headerRow.length + 1).setValue('Sedes Autorizadas');
          changed = true;
        }
      } else if (t === 'Sedes') {
        let rowsTemp = sheet.getDataRange().getValues();
        let headerRow = rowsTemp[0];
        let normalizedHeaders = headerRow.map(h => toCamelCase(h));
        const expectedHeaders = ['Id', 'Nombre', 'Empresa Id', 'Cliente Id', 'Direccion', 'Distrito', 'Ciudad', 'Referencia', 'Estado', 'Observaciones'];
        const expectedNormalized = expectedHeaders.map(h => toCamelCase(h));
        expectedNormalized.forEach((eh, index) => {
          if (!normalizedHeaders.includes(eh)) {
            sheet.getRange(1, sheet.getLastColumn() + 1).setValue(expectedHeaders[index]);
            changed = true;
          }
        });
      } else if (t === 'Accesos') {
        let rowsTemp = sheet.getDataRange().getValues();
        const headerRow = rowsTemp[0];
        const normalizedHeaders = headerRow.map(h => toCamelCase(h));
        if (!normalizedHeaders.includes('fechaModificacion')) {
          sheet.getRange(1, headerRow.length + 1).setValue('Fecha Modificacion');
          changed = true;
        }
      }
      if (changed) {
        SpreadsheetApp.flush();
        invalidateSheetCache(t);
      }
    } catch (migrationErr) {
      Logger.log('Advertencia de migración en tabla ' + t + ': ' + migrationErr.message);
    }

    let rows = getCachedSheetValues(t);
    if (rows.length < 1) { db[t.toLowerCase()] = []; return; }
    
    const headers = rows[0].map(h => toCamelCase(h));
    let dataRows = rows.slice(1);
    dataRows = filterRowsByRole(t, dataRows, headers, user);
    if (dataRows.length > limit) dataRows = dataRows.slice(0, limit);

    db[t.toLowerCase()] = dataRows.map(row => {
      const obj = {};
      headers.forEach((h, i) => {
        let val = row[i];
        if (t.toLowerCase() === 'accesos' && (h === 'passwordHash' || h === 'token')) {
          val = '********';
        }
        if (val instanceof Date) {
          const year = val.getFullYear();
          if (year === 1899 || year === 1900) {
            val = Utilities.formatDate(val, timezone, "HH:mm");
          } else {
            const timeStr = Utilities.formatDate(val, timezone, "HH:mm:ss.SSS");
            if (timeStr === "00:00:00.000") {
              val = Utilities.formatDate(val, timezone, "yyyy-MM-dd");
            } else {
              val = Utilities.formatDate(val, timezone, "yyyy-MM-dd'T'HH:mm:ss");
            }
          }
        }
        obj[h] = val;
      });
      return obj;
    });
  });
  return db;
}

function getIdColumn(table) {
  return (table.toLowerCase() === 'accesos') ? 'usuarioId' : 'id';
}

function createRecord(table, data) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) {
    throw new Error('No se pudo adquirir el bloqueo de escritura para creación (tiempo excedido)');
  }
  try {
    const sheet = getSheetAndCreateIfMissing(table);
    if (!sheet) throw new Error('Tabla no encontrada: ' + table);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const cleanHeaders = headers.map(h => toCamelCase(h));
    const idKey = getIdColumn(table);
    const idIdx = cleanHeaders.findIndex(h => h.toLowerCase() === idKey.toLowerCase());
    if (idIdx !== -1) {
      const actualIdKey = cleanHeaders[idIdx];
      if (!data[actualIdKey]) {
        const prefix = (table.toLowerCase() === 'accesos') ? 'USR-' : (table.toLowerCase() === 'productos') ? 'PROD-' : (table.toLowerCase() === 'lotes') ? 'LOTE-' : 'ID-';
        data[actualIdKey] = prefix + Utilities.getUuid().split('-')[0].toUpperCase();
      }
    }

    if (table.toLowerCase() === 'accesos') {
      const username = String(data.usuario || '').trim().toLowerCase();
      if (!username) throw new Error('El nombre de usuario es obligatorio');
      const uIdx = cleanHeaders.findIndex(h => h.toLowerCase() === 'usuario');
      if (uIdx !== -1) {
        const accRows = getCachedSheetValues('Accesos');
        const accDataRows = accRows.slice(1);
        const exists = accDataRows.some(r => String(r[uIdx]).trim().toLowerCase() === username);
        if (exists) throw new Error('El usuario ya existe');
      }
      if (data.password) {
        data.passwordHash = computeHash(data.password);
        delete data.password;
      }
      if (data.empleadoId) {
        const empRows = getCachedSheetValues('Empleados').slice(1);
        const exists = empRows.some(r => r[0] == data.empleadoId);
        if (!exists && data.empleadoId !== '0') throw new Error('El colaborador seleccionado no existe');
      }
      data.fechaCreacion = new Date();
      data.fechaModificacion = '';
    }

    if (table.toLowerCase() === 'asistencias') {
      if (!data.empresaId && data.empleadoId) {
        const empRows = getCachedSheetValues('Empleados').slice(1);
        const empHeaders = getCachedSheetValues('Empleados')[0].map(h => toCamelCase(h));
        const empRow = empRows.find(r => r[empHeaders.indexOf('id')] == data.empleadoId);
        if (empRow) data.empresaId = empRow[empHeaders.indexOf('empresaId')];
      }
      if (!data.empresaId && data.clienteId) {
        const cliRows = getCachedSheetValues('Clientes').slice(1);
        const cliHeaders = getCachedSheetValues('Clientes')[0].map(h => toCamelCase(h));
        const cliRow = cliRows.find(r => r[cliHeaders.indexOf('id')] == data.clienteId);
        if (cliRow) data.empresaId = cliRow[cliHeaders.indexOf('empresaId')];
      }
    }

    processBase64Fields(table, data[idKey], data);
    if (table === 'Asistencias' && (data.actividad || data.realizado || data.pendiente)) {
      const obs = {
        actividad: data.actividad || '',
        realizado: data.realizado || '',
        pendiente: data.pendiente || '',
        recomendaciones: data.recomendaciones || ''
      };
      if (data.observaciones) {
        try { const old = JSON.parse(data.observaciones); Object.assign(obs, old); } catch(e) {}
      }
      data.observaciones = JSON.stringify(obs);
      delete data.actividad; delete data.realizado; delete data.pendiente; delete data.recomendaciones;
    }

    const newRow = headers.map(h => {
      const key = toCamelCase(h);
      let val = data[key] !== undefined ? data[key] : '';
      if (typeof val === 'string' && val.startsWith('=')) val = "'" + val;
      return val;
    });
    sheet.appendRow(newRow);
    SpreadsheetApp.flush();
    invalidateSheetCache(table);
    return { status: 'success', id: data[idKey] || 'Creado' };
  } finally { lock.releaseLock(); }
}

function updateRecord(table, id, data) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) {
    throw new Error('No se pudo adquirir el bloqueo de escritura para actualización (tiempo excedido)');
  }
  try {
    const sheet = getSheetAndCreateIfMissing(table);
    if (!sheet) throw new Error('Tabla no encontrada: ' + table);
    const rows = sheet.getDataRange().getValues();
    if (rows.length < 2) throw new Error('No hay registros');
    const headers = rows[0].map(h => toCamelCase(h));
    const idKey = getIdColumn(table);
    const idIdx = headers.findIndex(h => h.toLowerCase() === idKey.toLowerCase());
    if (idIdx === -1) throw new Error('La tabla no tiene columna ID: ' + idKey);
    const rowIndex = rows.findIndex((r, i) => i > 0 && String(r[idIdx]).trim() === String(id).trim());
    if (rowIndex === -1) throw new Error('Registro no encontrado con ID: ' + id);

    processBase64Fields(table, id, data);
    if (table.toLowerCase() === 'accesos') {
      if (data.usuario) {
        const username = String(data.usuario).trim().toLowerCase();
        const uIdx = headers.findIndex(h => h.toLowerCase() === 'usuario');
        if (uIdx !== -1) {
          const duplicateIndex = rows.findIndex((r, i) => i > 0 && String(r[idIdx]).trim() !== String(id).trim() && String(r[uIdx]).trim().toLowerCase() === username);
          if (duplicateIndex !== -1) throw new Error('El nombre de usuario ya está en uso por otro acceso');
        }
      }
      if (data.password) {
        data.passwordHash = computeHash(data.password);
        delete data.password;
      }
      data.fechaModificacion = new Date();
    }

    const rowRange = sheet.getRange(rowIndex + 1, 1, 1, sheet.getLastColumn());
    const rowValues = rowRange.getValues()[0];
    Object.keys(data).forEach(key => {
      const colIndex = headers.indexOf(key);
      if (colIndex > -1 && key !== idKey) {
        let val = data[key];
        if (typeof val === 'string' && val.startsWith('=')) val = "'" + val;
        rowValues[colIndex] = val;
      }
    });
    rowRange.setValues([rowValues]);
    SpreadsheetApp.flush();
    invalidateSheetCache(table);
    return { status: 'success' };
  } finally { lock.releaseLock(); }
}

function checkRelationsBeforeDelete(table, id) {
  const SS = getSS();
  const normTable = String(table).toLowerCase();
  if (normTable === 'empresas') {
    const cliSheet = SS.getSheetByName('Clientes');
    if (cliSheet) {
      const cliData = cliSheet.getDataRange().getValues().slice(1);
      const cliHeaders = cliSheet.getDataRange().getValues()[0].map(h => toCamelCase(h));
      const cIdx = cliHeaders.indexOf('empresaId');
      if (cIdx !== -1 && cliData.some(r => String(r[cIdx]).trim() == String(id).trim())) {
        throw new Error('No se puede eliminar la empresa: Existen clientes vinculados a ella.');
      }
    }
    const empSheet = SS.getSheetByName('Empleados');
    if (empSheet) {
      const empData = empSheet.getDataRange().getValues().slice(1);
      const empHeaders = empSheet.getDataRange().getValues()[0].map(h => toCamelCase(h));
      const eIdx = empHeaders.indexOf('empresaId');
      if (eIdx !== -1 && empData.some(r => String(r[eIdx]).trim() == String(id).trim())) {
        throw new Error('No se puede eliminar la empresa: Existen colaboradores vinculados a ella.');
      }
    }
    const sedesSheet = SS.getSheetByName('Sedes');
    if (sedesSheet) {
      const sedesData = sedesSheet.getDataRange().getValues().slice(1);
      const sedesHeaders = sedesSheet.getDataRange().getValues()[0].map(h => toCamelCase(h));
      const sIdx = sedesHeaders.indexOf('empresaId');
      if (sIdx !== -1 && sedesData.some(r => String(r[sIdx]).trim() == String(id).trim())) {
        throw new Error('No se puede eliminar la empresa: Existen sedes vinculadas a ella.');
      }
    }
  }
  
  if (normTable === 'clientes') {
    const sedesSheet = SS.getSheetByName('Sedes');
    if (sedesSheet) {
      const sedesData = sedesSheet.getDataRange().getValues().slice(1);
      const sedesHeaders = sedesSheet.getDataRange().getValues()[0].map(h => toCamelCase(h));
      const sIdx = sedesHeaders.indexOf('clienteId');
      if (sIdx !== -1 && sedesData.some(r => String(r[sIdx]).trim() == String(id).trim())) {
        throw new Error('No se puede eliminar el cliente: Existen sedes vinculadas a él.');
      }
    }
    const calSheet = SS.getSheetByName('Calendario');
    if (calSheet) {
      const calData = calSheet.getDataRange().getValues().slice(1);
      const calHeaders = calSheet.getDataRange().getValues()[0].map(h => toCamelCase(h));
      const cIdx = calHeaders.indexOf('clienteId');
      if (cIdx !== -1 && calData.some(r => String(r[cIdx]).trim() == String(id).trim())) {
        throw new Error('No se puede eliminar el cliente: Existen actividades en el calendario vinculadas.');
      }
    }
    const asisSheet = SS.getSheetByName('Asistencias');
    if (asisSheet) {
      const asisData = asisSheet.getDataRange().getValues().slice(1);
      const asisHeaders = asisSheet.getDataRange().getValues()[0].map(h => toCamelCase(h));
      const aIdx = asisHeaders.indexOf('clienteId');
      if (aIdx !== -1 && asisData.some(r => String(r[aIdx]).trim() == String(id).trim())) {
        throw new Error('No se puede eliminar el cliente: Existen asistencias vinculadas.');
      }
    }
  }
}

function deleteRecord(table, id) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) {
    throw new Error('No se pudo adquirir el bloqueo de escritura para eliminación (tiempo excedido)');
  }
  try {
    const sheet = getSheetAndCreateIfMissing(table);
    if (!sheet) throw new Error('Tabla no encontrada: ' + table);
    checkRelationsBeforeDelete(table, id);
    const rows = sheet.getDataRange().getValues();
    if (rows.length < 2) throw new Error('No hay registros');
    const headers = rows[0].map(h => toCamelCase(h));
    const idKey = getIdColumn(table);
    const idIdx = headers.findIndex(h => h.toLowerCase() === idKey.toLowerCase());
    if (idIdx === -1) throw new Error('La tabla no tiene columna ID: ' + idKey);
    const rowIndex = rows.findIndex((r, i) => i > 0 && r[idIdx] == id);
    if (rowIndex === -1) throw new Error('Registro no encontrado con ID: ' + id);
    
    sheet.deleteRow(rowIndex + 1);
    SpreadsheetApp.flush();
    invalidateSheetCache(table);
    return { status: 'success' };
  } finally { lock.releaseLock(); }
}

function toCamelCase(str) {
  if (!str) return '';
  return str.toString()
    .replace(/([a-z])([A-Z])/g, '$1 $2')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[-_]/g, ' ')
    .replace(/[^a-zA-Z0-9 ]/g, '')
    .split(' ')
    .filter(w => w.length > 0)
    .map((w, i) => i === 0 ? w.toLowerCase() : w.charAt(0).toUpperCase() + w.slice(1).toLowerCase())
    .join('');
}

function computeHash(input) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, input);
  return digest.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

function processBase64Fields(table, id, data) {
  const photoKeys = ['foto', 'fotoEntrada', 'fotoSalida'];
  photoKeys.forEach(key => {
    if (data[key] && data[key].length > 100 && !data[key].startsWith('http') && !data[key].startsWith('Drive:')) {
      const folderName = (table.toLowerCase() === 'empleados') ? 'Colaboradores' : 'Asistencias';
      const suffix = key === 'foto' ? '' : '_' + key;
      const filename = (id || 'photo') + suffix + '.jpg';
      try {
        const uploadResult = uploadPhoto(data[key], filename, folderName);
        data[key] = uploadResult;
      } catch (uploadError) {
        Logger.log('Error al subir la foto en processBase64Fields: ' + uploadError.message);
        delete data[key];
      }
    }
  });
}

function uploadPhoto(base64, filename, folderName) {
  try {
    const blob = Utilities.newBlob(Utilities.base64Decode(base64), 'image/jpeg', filename);
    let rootFolder;
    try { rootFolder = DriveApp.getFolderById(FOLDER_ID); } catch (folderError) {
      rootFolder = DriveApp.getRootFolder();
    }
    let targetFolder = rootFolder;
    if (folderName) {
      try {
        const subfolders = rootFolder.getFoldersByName(folderName);
        if (subfolders.hasNext()) targetFolder = subfolders.next();
        else targetFolder = rootFolder.createFolder(folderName);
      } catch (subfolderErr) { targetFolder = rootFolder; }
    }
    try {
      const existing = targetFolder.getFilesByName(filename);
      while (existing.hasNext()) { existing.next().setTrashed(true); }
    } catch (trashError) {}
    
    const file = targetFolder.createFile(blob);
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {
      try { file.setSharing(DriveApp.Access.DOMAIN, DriveApp.Permission.VIEW); } catch(d) {}
    }
    return `http://googleusercontent.com/profile/picture/${file.getId()}`; // CORREGIDO: Interpolación limpia
  } catch (e) { throw new Error('Error al subir la foto: ' + e.message); }
}

function getSheetAndCreateIfMissing(table) {
  const SS = getSS();
  let sheet = SS.getSheetByName(table);
  if (sheet && table === 'Asistencias') {
    try { migrateAsistenciasToUnified(sheet); } catch(err) { Logger.log(err.message); }
  }
  if (!sheet) {
    sheet = SS.insertSheet(table);
    let headers = [];
    if (table === 'Sedes') headers = ['Id', 'Nombre', 'Empresa Id', 'Cliente Id', 'Direccion', 'Distrito', 'Ciudad', 'Referencia', 'Estado', 'Observaciones'];
    else if (table === 'Asistencias') headers = ['Id', 'Empleado Id', 'Empresa Id', 'Cliente Id', 'Sede Id', 'Fecha', 'Hora Entrada', 'Hora Salida', 'Lat Entrada', 'Lng Entrada', 'Lat Salida', 'Lng Salida', 'Foto Entrada', 'Foto Salida', 'Observaciones Entrada', 'Observaciones Salida', 'Estado', 'Horas Trabajadas'];
    else if (table === 'Empleados') headers = ['Id', 'Dni', 'Nombre', 'Cargo', 'Telefono', 'Correo', 'Estado', 'Foto', 'Sedes Autorizadas'];
    else if (table === 'Empresas') headers = ['Id', 'Razon Social', 'Ruc', 'Telefono', 'Direccion', 'Correo', 'Sede', 'Empleados Asignados'];
    else if (table === 'Clientes') headers = ['Id', 'Nombre', 'Ruc Empresa', 'Contacto', 'Telefono', 'Direccion', 'Correo', 'Estado', 'Empresa Id', 'Empleados Asignados'];
    else if (table === 'Calendario') headers = ['Id', 'Fecha', 'Hora Inicio', 'Hora Fin', 'Empleado Id', 'Cliente Id', 'Estado', 'Prioridad', 'Observaciones'];
    else if (table === 'Accesos') headers = ['Usuario Id', 'Usuario', 'Password Hash', 'Empleado Id', 'Rol', 'Estado', 'Fecha Creacion', 'Ultimo Acceso', 'Token', 'Fecha Expiracion', 'Fecha Modificacion'];
    else if (table === 'Logs') headers = ['Id', 'Fecha', 'Usuario Admin', 'Accion', 'Detalle'];
    else if (table === 'Productos') headers = ['Id', 'Nombre', 'Categoria', 'Marca', 'Unidad Medida', 'Stock General', 'Precio Base', 'Estado'];
    else if (table === 'Lotes') headers = ['Id', 'Producto Id', 'Numero Lote', 'Fecha Fabricacion', 'Fecha Vencimiento', 'Stock Lote', 'Certificado Calidad Link', 'Estado'];
    else if (table === 'CrmHistorial') headers = ['Id', 'Cliente Id', 'Empleado Id', 'Tipo', 'Detalle', 'Fecha Hora', 'Adjunto Link'];
    else if (table === 'Pipeline') headers = ['Id', 'Cliente Id', 'Empleado Id', 'Estado', 'Valor Estimado', 'Fecha Probable Cierre', 'Observaciones'];
    else if (table === 'Documentos') headers = ['Id', 'Cliente Id', 'Nombre', 'Tipo', 'Drive File Id', 'Fecha Carga', 'Cargado Por'];
    else if (table === 'Muestras') headers = ['Id', 'Cliente Id', 'Producto Id', 'Lote Id', 'Cantidad', 'Unidad Medida', 'Fecha Entrega', 'Estado', 'Observaciones'];
    else if (table === 'Incidencias') headers = ['Id', 'Cliente Id', 'Empleado Id', 'Titulo', 'Descripcion', 'Tipo', 'Estado', 'Fecha Creacion', 'Fecha Resolucion', 'Resolucion Detalle'];
    else headers = ['Id', 'Nombre'];
    sheet.appendRow(headers);
    SpreadsheetApp.flush(); invalidateSheetCache(table);
  }
  return sheet;
}

function response(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function writeLog(action, detail, token) {
  try {
    let adminName = 'Sistema';
    if (token) {
      try {
        const user = validateToken(token);
        if (user.empleadoId && user.empleadoId !== '0') {
          const empData = getCachedSheetValues('Empleados');
          const empHeaders = empData[0].map(h => toCamelCase(h));
          const empRow = empData.slice(1).find(r => r[empHeaders.indexOf('id')] == user.empleadoId);
          if (empRow) adminName = empRow[empHeaders.indexOf('nombre')] || 'Admin';
        } else if (user.rol === 'ADMIN') { adminName = 'Administrador'; }
      } catch (tokenErr) {}
    }
    const sheet = getSheetAndCreateIfMissing('Logs');
    if (sheet) {
      const logId = 'LOG-' + Utilities.getUuid().split('-')[0].toUpperCase();
      sheet.appendRow([logId, new Date(), adminName, action, detail]);
      SpreadsheetApp.flush(); invalidateSheetCache('Logs');
    }
  } catch(e) {}
}

function migrateAsistenciasToUnified(sheet) {
  const SS = getSS();
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 1 || !rows[0].map(h => toCamelCase(h)).includes('tipo')) return;
  const headers = rows[0].map(h => toCamelCase(h));
  const dataRows = rows.slice(1);
  const empIdx = headers.indexOf('empleadoId');
  const fechaIdx = headers.indexOf('fecha');
  const tipoIdx = headers.indexOf('tipo');
  const horaIdx = headers.indexOf('hora');
  const latIdx = headers.indexOf('lat');
  const lngIdx = headers.indexOf('lng');
  const fotoIdx = headers.indexOf('foto');
  const obsIdx = headers.indexOf('observaciones');
  const cliIdx = headers.indexOf('clienteId');
  const empCompIdx = headers.indexOf('empresaId');
  const sedeIdx = headers.indexOf('sedeId');

  const groups = {};
  dataRows.forEach(row => {
    const empId = String(row[empIdx]).trim();
    let fecha = row[fechaIdx];
    if (fecha instanceof Date) fecha = Utilities.formatDate(fecha, SS.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
    else fecha = String(fecha).trim();
    if (!empId || !fecha) return;
    const key = empId + '_' + fecha;
    if (!groups[key]) groups[key] = [];
    groups[key].push(row);
  });

  const newRows = [];
  Object.keys(groups).forEach(key => {
    const groupRows = groups[key];
    const entradaRow = groupRows.find(r => String(r[tipoIdx]).toUpperCase() === 'ENTRADA');
    const salidaRow = groupRows.find(r => String(r[tipoIdx]).toUpperCase() === 'SALIDA');
    const baseRow = entradaRow || salidaRow || groupRows[0];
    const id = baseRow[headers.indexOf('id')] || ('ID-' + Utilities.getUuid().split('-')[0].toUpperCase());
    let fecha = baseRow[fechaIdx];
    if (fecha instanceof Date) fecha = Utilities.formatDate(fecha, SS.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
    
    let horaEntrada = '', latEntrada = '', lngEntrada = '', fotoEntrada = '', obsEntrada = '';
    if (entradaRow) {
      horaEntrada = entradaRow[horaIdx] instanceof Date ? Utilities.formatDate(entradaRow[horaIdx], SS.getSpreadsheetTimeZone(), 'HH:mm:ss') : entradaRow[horaIdx];
      latEntrada = entradaRow[latIdx]; lngEntrada = entradaRow[lngIdx]; fotoEntrada = entradaRow[fotoIdx];
      let obsText = entradaRow[obsIdx] || '';
      try { obsEntrada = JSON.parse(obsText).actividad || obsText; } catch(e) { obsEntrada = obsText; }
    }
    
    let horaSalida = '', latSalida = '', lngSalida = '', fotoSalida = '', obsSalida = '';
    if (salidaRow) {
      horaSalida = salidaRow[horaIdx] instanceof Date ? Utilities.formatDate(salidaRow[horaIdx], SS.getSpreadsheetTimeZone(), 'HH:mm:ss') : salidaRow[horaIdx];
      latSalida = salidaRow[latIdx]; lngSalida = salidaRow[lngIdx]; fotoSalida = salidaRow[fotoIdx];
      let obsText = salidaRow[obsIdx] || '';
      try { obsSalida = obsText; JSON.parse(obsText); } catch(e) {
        obsSalida = JSON.stringify({ realizado: obsText, observaciones: '', pendiente: '', recomendaciones: '' });
      }
    }
    
    const estado = (horaEntrada && horaSalida) ? 'COMPLETO' : 'ABIERTO';
    let horasTrabajadas = '';
    if (horaEntrada && horaSalida) {
      try {
        const parseTime = (str) => { var p = String(str).split(':'); return parseInt(p[0], 10) * 60 + parseInt(p[1], 10); };
        const diff = parseTime(horaSalida) - parseTime(horaEntrada);
        if (diff >= 0) horasTrabajadas = Math.floor(diff / 60) + 'h ' + (diff % 60) + 'm';
      } catch(e) {}
    }
    newRows.push([id, baseRow[empIdx], baseRow[empCompIdx], baseRow[cliIdx], baseRow[sedeIdx], fecha, horaEntrada, horaSalida, latEntrada, lngEntrada, latSalida, lngSalida, fotoEntrada, fotoSalida, obsEntrada, obsSalida, estado, horasTrabajadas]);
  });

  sheet.clear();
  const newHeaders = ['Id', 'Empleado Id', 'Empresa Id', 'Cliente Id', 'Sede Id', 'Fecha', 'Hora Entrada', 'Hora Salida', 'Lat Entrada', 'Lng Entrada', 'Lat Salida', 'Lng Salida', 'Foto Entrada', 'Foto Salida', 'Observaciones Entrada', 'Observaciones Salida', 'Estado', 'Horas Trabajadas'];
  sheet.appendRow(newHeaders);
  if (newRows.length > 0) sheet.getRange(2, 1, newRows.length, newHeaders.length).setValues(newRows);
  SpreadsheetApp.flush(); invalidateSheetCache('Asistencias');
}

function apiGet(parameter) {
  try {
    const token = parameter.token; if (!token) throw new Error('Token requerido');
    const user = validateToken(token); const data = getMasterData(user);
    data._meta = { newToken: user.newToken }; return data;
  } catch (err) { return { status: 'error', message: err.message }; }
}

function apiPost(params) {
  try { return processPostAction(params); } catch (err) { return { status: 'error', message: err.message }; }
}