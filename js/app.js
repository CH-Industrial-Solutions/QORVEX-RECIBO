/* QORVEX Recibo - interfaz y acciones */
(function () {
  'use strict';

  var QR = window.QR;
  var XLSX = window.XLSX;
  var ARCHIVO = 'QORVEX_Recibo.xlsx';
  var RAIZ = document.getElementById('raiz');
  var ESTADOS_FACTURA = ['Pendiente', 'Recibida', 'Conciliada'];
  var TIPOS = ['Descargados', 'Recolectados', 'Ubicados', 'Cargados'];
  var COLOR_ESTATUS = {
    'Completo': '#E2F0D9', 'Parcial': '#FFF2CC', 'Ubicado': '#DDEBF7', 'Recolectar': '#FBE5D6',
    'Por Revisar': '#EDEDED', 'Sin Información': '#F8D7D7', 'Solicitar caja': '#D9E2F3', 'Caja Solicitada': '#E4DFEC'
  };
  // Cinta de menu: grupos y botones. m = modos (roles) que ven el boton; sin m, lo ven todos.
  var CAP_SUP = ['Supervisor', 'Administrador'];
  var CAP_OPS = ['Capturista', 'Supervisor', 'Administrador'];
  var CINTAS = [
    { id: 'Inicio', items: [{ t: 'inicio', n: 'Inicio' }, { t: 'ops', n: 'Últimas operaciones' }] },
    { id: 'Operación', items: [{ t: 'piso', n: 'Piso', m: CAP_OPS }, { t: 'rampas', n: 'Rampas', m: CAP_OPS },
      { t: 'facturas', n: 'Facturas', m: CAP_OPS }, { t: 'cortos', n: 'Cortos', m: CAP_SUP }, { t: 'notas', n: 'Notas', m: CAP_SUP }] },
    { id: 'Análisis', items: [{ t: 'kpi', n: 'KPIs' }] },
    { id: 'Sistema', items: [{ t: 'config', n: 'Configuración', m: ['Administrador'] }] }
  ];

  var S = {
    pantalla: 'carga', dir: null, fh: null, mtimes: {}, reg: null, data: null,
    lock: null, miBloqueo: null, motivoBloqueo: '', setup: false, escribirTodo: false,
    usuario: null, rol: null, encargado: '', tab: 'inicio', cinta: 'Inicio', msg: null,
    turno: '', turnoAuto: true, tema: 'negro',
    snap: null, cortosModo: 'primera', preguntaActualizar: false,
    f: {
      turno: 'Todos', sup: 'Todos', linea: 'Todas', estatus: 'Todos', buscar: '', bf: '',
      ops: { turno: 'Todos', modulo: 'Todos', usuario: 'Todos' }, kpiDia: '', kpiMes: '', entDia: '', entTurno: ''
    }
  };

  // ---------- utilidades ----------
  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }
  function p2(n) { return (n < 10 ? '0' : '') + n; }
  function fmtF(v) {
    if (v instanceof Date && !isNaN(v)) return p2(v.getDate()) + '/' + p2(v.getMonth() + 1) + '/' + v.getFullYear() + ' ' + p2(v.getHours()) + ':' + p2(v.getMinutes());
    return v == null ? '' : String(v);
  }
  function fmtD(v) {
    if (typeof v === 'number' && v > 30000 && v < 80000) {
      var d = new Date(Math.round((v - 25569) * 86400000));
      return p2(d.getUTCDate()) + '/' + p2(d.getUTCMonth() + 1) + '/' + d.getUTCFullYear();
    }
    if (v instanceof Date && !isNaN(v)) return p2(v.getDate()) + '/' + p2(v.getMonth() + 1) + '/' + v.getFullYear();
    return v == null ? '' : String(v);
  }
  function fmtH(h) { return h == null || isNaN(h) ? '—' : h.toFixed(1) + ' h'; }
  function hh(v) { var m = QR.minutosHora(v); return m == null ? '' : p2(Math.floor(m / 60)) + ':' + p2(m % 60); }
  function aInput(d) { return d.getFullYear() + '-' + p2(d.getMonth() + 1) + '-' + p2(d.getDate()) + 'T' + p2(d.getHours()) + ':' + p2(d.getMinutes()); }
  function aFecha(d) { return d.getFullYear() + '-' + p2(d.getMonth() + 1) + '-' + p2(d.getDate()); }
  function aMes(d) { return d.getFullYear() + '-' + p2(d.getMonth() + 1); }
  function deFecha(s) { var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s || ''); return m ? new Date(+m[1], +m[2] - 1, +m[3]) : null; }
  function deMes(s) { var m = /^(\d{4})-(\d{2})$/.exec(s || ''); return m ? { anio: +m[1], mes: +m[2] - 1 } : null; }
  function deInputDT(s) { var d = s ? new Date(s) : null; return d && !isNaN(d) ? d : null; }
  function inicioDia(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
  function finDia(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate(), 23, 59, 59, 999); }
  function tiempo(v) { if (v instanceof Date) return v.getTime(); var t = Date.parse(v); return isNaN(t) ? 0 : t; }
  function P() { return QR.parametros(S.data.Parametros); }
  function esSup() { return S.rol === 'Supervisor' || S.rol === 'Administrador'; }
  function listaTurnos() { return S.data.Turnos.map(function (t) { return t.Turno; }); }
  function turnoObj(nombre) { return S.data.Turnos.filter(function (t) { return t.Turno === nombre; })[0] || null; }
  function resolverTurno() { return QR.turnoEn(new Date(), S.data.Turnos) || listaTurnos()[0] || ''; }
  function nombreDe(noEmp) {
    var u = S.data.Usuarios.filter(function (x) { return x.NoEmpleado === noEmp; })[0];
    return u ? u.Nombre : noEmp;
  }
  function claseSem(e) { return { 'En tiempo': 'sem-ok', 'Alerta': 'sem-alerta', 'Demora': 'sem-demora' }[e] || 'sem-vacia'; }
  function opts(lista, sel) {
    return lista.map(function (v) { return '<option value="' + esc(v) + '"' + (String(v) === String(sel) ? ' selected' : '') + '>' + esc(v) + '</option>'; }).join('');
  }
  function metaDe(t) {
    var m = S.data.Metas.filter(function (x) { return x.Turno === t; })[0];
    return m ? { Descargados: +m.Descargados || 0, Recolectados: +m.Recolectados || 0, Ubicados: +m.Ubicados || 0, Cargados: +m.Cargados || 0 }
      : { Descargados: 0, Recolectados: 0, Ubicados: 0, Cargados: 0 };
  }
  function sumaMeta(m) { return m.Descargados + m.Recolectados + m.Ubicados + m.Cargados; }
  function cuentaTipos(lista) {
    var c = { Descargados: 0, Recolectados: 0, Ubicados: 0, Cargados: 0 };
    lista.forEach(function (e) { if (c[e.tipo] != null) c[e.tipo]++; });
    return c;
  }
  function suma(c) { return TIPOS.reduce(function (s, k) { return s + c[k]; }, 0); }
  function pct(real, meta) { return meta ? Math.round(real / meta * 100) + '%' : '—'; }
  function celdasTipos(c) { return TIPOS.map(function (k) { return '<td class="num">' + c[k] + '</td>'; }).join(''); }
  function kpi(n, t, cls) { return '<div class="kpi ' + (cls || '') + '"><b>' + n + '</b><span>' + t + '</span></div>'; }
  function snapActivo() {
    var a = S.data.Cortos.filter(function (r) { return r.Activo === 'Sí'; })[0];
    return a ? String(a.Snapshot) : null;
  }
  function maxSnap() {
    var m = 0;
    S.data.CortosSnapshots.concat(S.data.Cortos, S.data.CortosHistorico || []).forEach(function (r) { var n = +r.Snapshot; if (n > m) m = n; });
    return m;
  }
  function logH(ev) {
    var fila = Object.assign({
      FechaEvento: new Date(), ID: '', Bahia: '', Caja: '', Evento: '', EstatusAnt: '', EstatusNvo: '',
      Responsable: '', NoEmpleado: S.usuario.NoEmpleado, NombreUsuario: S.usuario.Nombre, Turno: S.turno, Detalle: '',
      Encargado: encargadoActual(), Unidad: ''
    }, ev);
    S.data.Historial.push(fila);
    // La fila va al registro de la computadora que tiene el bloqueo
    var destino = S.lock === 'cortos' ? S.reg.regS : S.lock === 'rampas' ? S.reg.regC : null;
    if (destino) destino.push(fila);
  }
  // Encargado del turno: el supervisor elegido al entrar (o cambiado en la barra superior)
  function encargadoActual() { return S.encargado || (S.usuario ? S.usuario.NoEmpleado : ''); }
  // Encargados posibles: supervisores y administradores activos
  function supervisores() { return S.data.Usuarios.filter(function (u) { return (u.Rol === 'Supervisor' || u.Rol === 'Administrador') && u.Activo === 'Sí'; }); }
  function nombreEnc(no) { return no ? nombreDe(no) : 'Sin encargado'; }
  function opcionesEncargado(sel, conVacio) {
    return (conVacio ? '<option value="">— Elige encargado —</option>' : '') + supervisores().map(function (u) {
      return '<option value="' + esc(u.NoEmpleado) + '"' + (u.NoEmpleado === sel ? ' selected' : '') + '>' + esc(u.Nombre) + '</option>';
    }).join('');
  }

  // ---------- turnos, jornadas y dias laborables ----------
  // Eventos de productividad con su dia de jornada (turno nocturno: madrugada cuenta en el dia anterior)
  function eventosDeRango(desde, hasta) {
    var ini = new Date(desde.getFullYear(), desde.getMonth(), desde.getDate() - 1);
    var fin = new Date(hasta.getFullYear(), hasta.getMonth(), hasta.getDate() + 1, 23, 59, 59, 999);
    var a = inicioDia(desde), b = inicioDia(hasta);
    return QR.eventosProductividad(S.data, ini, fin).map(function (e) {
      return { tipo: e.tipo, fecha: e.fecha, turno: e.turno || '', noEmp: e.noEmp, enc: e.enc || '', dia: QR.diaJornada(e.fecha, S.data.Turnos, e.turno).dia };
    }).filter(function (e) { return e.dia >= a && e.dia <= b; });
  }
  function jornadaHoy() { return QR.diaJornada(new Date(), S.data.Turnos, S.turno).dia; }
  function trabajaTurno(dia, turno) {
    return QR.turnosDelDia(QR.nombreDia(dia), S.data.DiasLaborales).indexOf(turno) >= 0 && !QR.esFestivo(dia, S.data.Festivos);
  }
  function metaDelDia(turno, dia) { return trabajaTurno(dia, turno) ? metaDe(turno) : { Descargados: 0, Recolectados: 0, Ubicados: 0, Cargados: 0 }; }

  // ---------- acceso a la carpeta y al libro ----------
  function kvAbrir() {
    return new Promise(function (res, rej) {
      var r = indexedDB.open('qorvex-recibo', 1);
      r.onupgradeneeded = function () { r.result.createObjectStore('kv'); };
      r.onsuccess = function () { res(r.result); };
      r.onerror = function () { rej(r.error); };
    });
  }
  function kvGet(k) {
    return kvAbrir().then(function (db) {
      return new Promise(function (res, rej) {
        var q = db.transaction('kv').objectStore('kv').get(k);
        q.onsuccess = function () { res(q.result || null); };
        q.onerror = function () { rej(q.error); };
      });
    });
  }
  function kvSet(k, v) {
    return kvAbrir().then(function (db) {
      return new Promise(function (res, rej) {
        var t = db.transaction('kv', 'readwrite');
        t.objectStore('kv').put(v, k);
        t.oncomplete = function () { res(); };
        t.onerror = function () { rej(t.error); };
      });
    });
  }
  function hashPass(noEmp, pass) {
    var bytes = new TextEncoder().encode(noEmp + ':' + pass);
    return crypto.subtle.digest('SHA-256', bytes).then(function (buf) {
      return Array.prototype.map.call(new Uint8Array(buf), function (b) { return (b < 16 ? '0' : '') + b.toString(16); }).join('');
    });
  }

  // ---------- archivos compartidos, bloqueos y sesion ----------
  // Identidad de esta computadora (reconoce su propio bloqueo)
  var PC_ID = (function () {
    try {
      var v = localStorage.getItem('qr-pc');
      if (!v) { v = Math.random().toString(36).slice(2) + Date.now().toString(36); localStorage.setItem('qr-pc', v); }
      return v;
    } catch (e) { return 'pc-' + Math.random().toString(36).slice(2); }
  })();
  // Pestana -> area que la escribe
  var AREA_TAB = { piso: 'rampas', rampas: 'rampas', facturas: 'rampas', cortos: 'cortos', notas: 'cortos', config: 'config' };
  // Ultima actividad real (clic, tecla, desplazar, escribir). La recarga automatica no cuenta.
  var ultimaAct = Date.now();
  ['pointerdown', 'keydown', 'scroll', 'input', 'change', 'touchstart'].forEach(function (ev) {
    document.addEventListener(ev, function () { ultimaAct = Date.now(); }, true);
  });
  function enFoco() { var el = document.activeElement; return !!(el && RAIZ.contains(el) && /^(INPUT|SELECT|TEXTAREA)$/.test(el.tagName)); }

  // Quien puede escribir cada area: rampas (capturista o administrador), cortos (supervisor), config (administrador)
  function puedeEscribir(area) {
    if (S.setup) return true;
    if (area === 'rampas') return S.lock === 'rampas' && (S.rol === 'Capturista' || S.rol === 'Administrador');
    if (area === 'config') return S.lock === 'rampas' && S.rol === 'Administrador';
    if (area === 'cortos') return S.lock === 'cortos' && S.rol === 'Supervisor';
    return false;
  }
  function motivoLectura() {
    if (S.rol === 'Consulta') return 'Solo lectura: tu modo es consulta.';
    if (S.lock) return 'Esta parte no se edita en este modo.';
    return 'Solo lectura: el bloqueo lo tiene ' + (S.motivoBloqueo || 'otra computadora') + '. Toma el bloqueo cuando quede libre.';
  }
  // Archivos que esta computadora escribe: los de su bloqueo
  function archivosPropios() {
    if (S.setup || S.escribirTodo) return Object.keys(QR.ARCHIVOS);
    if (S.lock === 'rampas') return ['rampas', 'regC'];
    if (S.lock === 'cortos') return ['cortos', 'regS'];
    return [];
  }
  // Hojas de cada archivo, desde la memoria de trabajo
  function hojasDe(k) {
    var o = {};
    if (k === 'rampas') QR.ARCHIVOS.rampas.hojas.forEach(function (n) { o[n] = S.data[n] || []; });
    else if (k === 'cortos') { o.Cortos = S.data.Cortos || []; o.Notas = S.data.Notas || []; }
    else if (k === 'regC') o.Historial = S.reg.regC;
    else if (k === 'regS') {
      o.Historial = S.reg.regS; o.CortosLog = S.data.CortosLog || [];
      o.CortosVersiones = S.data.CortosSnapshots || []; o.CortosHistorico = S.data.CortosHistorico || [];
    }
    return o;
  }

  var cola = Promise.resolve();
  // Escribe solo los archivos de esta computadora (su bloqueo)
  function guardar(forzar) {
    var p = cola.then(async function () {
      var lista = archivosPropios();
      for (var i = 0; i < lista.length; i++) {
        var k = lista[i], fh = S.fh[k];
        if (!forzar) {
          var f0 = await fh.getFile();
          if (f0.lastModified !== S.mtimes[k]) throw new Error('el archivo ' + QR.ARCHIVOS[k].archivo + ' cambió en otra computadora; se recargará');
        }
        var bytes = QR.aLibro(XLSX, hojasDe(k), QR.ARCHIVOS[k].hojas);
        var w = await fh.createWritable();
        await w.write(new Blob([bytes]));
        await w.close();
        S.mtimes[k] = (await fh.getFile()).lastModified;
      }
    });
    cola = p.catch(function () {});
    return p;
  }

  // Bloqueo: archivo de texto en la carpeta que dice quien escribe. Se renueva cada minuto.
  async function leerBloqueo(k) {
    var fh = await S.dir.getFileHandle(QR.BLOQUEOS[k], { create: true });
    var buf = await (await fh.getFile()).arrayBuffer();
    try { var o = JSON.parse(new TextDecoder().decode(buf)); return o && o.pc ? o : null; } catch (e) { return null; }
  }
  async function escribirBloqueo(k, obj) {
    var fh = await S.dir.getFileHandle(QR.BLOQUEOS[k], { create: true });
    var w = await fh.createWritable();
    await w.write(new Blob([JSON.stringify(obj || {})]));
    await w.close();
  }
  async function tomarBloqueo(k, forzar) {
    var ahora = new Date();
    var actual = await leerBloqueo(k);
    if (!forzar && QR.bloqueoOcupado(actual, PC_ID, ahora)) return { ok: false, de: actual };
    var mio = { pc: PC_ID, usuario: S.usuario.Nombre, modo: S.rol, desde: ahora.toISOString(), latido: ahora.toISOString() };
    await escribirBloqueo(k, mio);
    var verif = await leerBloqueo(k);
    if (!verif || verif.pc !== PC_ID) return { ok: false, de: verif };
    return { ok: true, mio: mio };
  }
  async function liberarBloqueo() {
    var k = S.lock;
    S.lock = null; S.miBloqueo = null;
    if (!k || !S.dir) return;
    try { var a = await leerBloqueo(k); if (a && a.pc === PC_ID) await escribirBloqueo(k, {}); } catch (e) { /* sin red: el bloqueo caduca solo */ }
  }
  // Renueva el latido. Si otra computadora tomo el bloqueo, esta pasa a solo lectura.
  async function latir() {
    if (!S.lock || !S.miBloqueo || !S.dir) return;
    var k = S.lock;
    try {
      var ahora = new Date();
      var actual = await leerBloqueo(k);
      if (actual && actual.pc !== PC_ID && !QR.latidoVencido(actual, ahora)) {
        S.lock = null; S.miBloqueo = null;
        S.motivoBloqueo = actual.usuario + ' (' + actual.modo + ')';
        S.msg = { tipo: 'info', texto: 'Otra computadora tomó el bloqueo. Ahora estás en solo lectura.' };
        return render();
      }
      S.miBloqueo.latido = ahora.toISOString();
      await escribirBloqueo(k, S.miBloqueo);
    } catch (e) { /* se reintenta en el siguiente latido */ }
  }
  function estadoBloqueo() {
    if (S.lock) return '<b>Editando: ' + (S.lock === 'rampas' ? 'Rampas y bahías' : 'Cortos') + '</b>';
    return 'Solo lectura' + (S.motivoBloqueo ? ' · bloqueado por ' + esc(S.motivoBloqueo) : '') +
      (QR.bloqueoDeModo(S.rol) ? ' <button class="sec peq" data-act="tomarBloqueo">Tomar bloqueo</button>' : '');
  }
  // Entra en un modo: toma el bloqueo que le corresponde o queda en solo lectura
  async function entrarEnModo(modo) {
    await liberarBloqueo();
    S.rol = modo; S.motivoBloqueo = ''; ultimaAct = Date.now();
    var k = QR.bloqueoDeModo(modo);
    if (!k) return;
    var r = await tomarBloqueo(k, false);
    if (r.ok) {
      S.lock = k; S.miBloqueo = r.mio;
      if (k === 'rampas' && migrar()) await guardarSiCambio();
    } else {
      S.motivoBloqueo = r.de ? r.de.usuario + ' (' + r.de.modo + ')' : 'otra computadora';
    }
  }
  // Entrada desde el login: valida el modo y el encargado
  async function entrarDesdeLogin(modo) {
    if ((QR.MODOS[S.usuario.Rol] || []).indexOf(modo) < 0) { S.msg = { tipo: 'err', texto: 'Tu usuario no puede entrar en ese modo.' }; return render(); }
    if (modo === 'Supervisor') S.encargado = S.usuario.NoEmpleado;
    if (modo === 'Capturista') {
      var sel = RAIZ.querySelector('#encEntrada');
      var enc = sel ? sel.value : '';
      if (!enc) { S.msg = { tipo: 'err', texto: 'Elige el encargado del turno para entrar como capturista.' }; return render(); }
      S.encargado = enc;
    }
    S.msg = null;
    await entrarEnModo(modo);
    S.pantalla = 'app'; S.tab = 'inicio'; S.cinta = 'Inicio';
    S.turnoAuto = true; S.turno = resolverTurno();
    if (!S.lock && S.motivoBloqueo) S.msg = { tipo: 'info', texto: 'Entras en solo lectura: el bloqueo lo tiene ' + S.motivoBloqueo + '.' };
    render();
  }
  // Sin actividad durante 15 minutos: libera el bloqueo y pide solo la contrasena
  async function revisarInactividad() {
    if (S.pantalla !== 'app') return;
    if (Date.now() - ultimaAct < QR.MINUTOS_INACTIVIDAD * 60000) return;
    await liberarBloqueo();
    S.pantalla = 'bloqueo';
    S.msg = { tipo: 'info', texto: 'Sesión bloqueada por inactividad. Escribe tu contraseña para seguir.' };
    render();
  }

  // Lectura de archivos
  // Hojas de un archivo; null si esta vacio
  async function leerArchivo(k) {
    var f = await S.fh[k].getFile();
    S.mtimes[k] = f.lastModified;
    var buf = await f.arrayBuffer();
    return buf.byteLength ? QR.desdeLibro(XLSX, buf, QR.ARCHIVOS[k].hojas) : null;
  }
  function vacioDe(k) {
    var v = QR.datosVacios(), o = {};
    QR.ARCHIVOS[k].hojas.forEach(function (n) { o[n] = v[n] || []; });
    return o;
  }
  // Hojas -> memoria de trabajo. CortosVersiones entra como CortosSnapshots; el historial va aparte.
  function hojasAMemoria(o) {
    Object.keys(o).forEach(function (n) {
      if (n === 'Historial') return;
      S.data[n === 'CortosVersiones' ? 'CortosSnapshots' : n] = o[n] || [];
    });
  }
  function refrescarHistorial() { S.data.Historial = S.reg.regC.concat(S.reg.regS); }
  function aplicarArchivos(a) {
    S.reg = { regC: (a.regC.Historial || []).slice(), regS: (a.regS.Historial || []).slice() };
    S.data = {};
    hojasAMemoria(a.rampas); hojasAMemoria(a.cortos); hojasAMemoria(a.regS);
    refrescarHistorial();
  }
  // Carga los cuatro archivos. Devuelve los que no existian o estaban vacios.
  async function cargarTodo() {
    var nuevo = {}, sinDatos = [];
    for (var k of Object.keys(QR.ARCHIVOS)) {
      var v = await leerArchivo(k);
      if (!v) { sinDatos.push(k); v = vacioDe(k); }
      nuevo[k] = v;
    }
    aplicarArchivos(nuevo);
    migrar();
    return sinDatos;
  }
  async function recargar() {
    await cargarTodo();
    if (S.usuario) {
      S.usuario = S.data.Usuarios.filter(function (u) { return u.NoEmpleado === S.usuario.NoEmpleado; })[0] || null;
    }
  }
  // Cada minuto relee los archivos que no son de esta computadora (lo que ven entrega y KPIs)
  function refrescarLectura() {
    if (!S.fh || !S.data) return Promise.resolve();
    var p = cola.then(async function () {
      var propios = archivosPropios();
      for (var k of Object.keys(QR.ARCHIVOS)) {
        if (propios.indexOf(k) >= 0) continue;
        var v = await leerArchivo(k);
        if (!v) continue;
        if (k === 'regC' || k === 'regS') S.reg[k] = (v.Historial || []).slice();
        hojasAMemoria(v);
      }
      refrescarHistorial();
      if (S.usuario) S.usuario = S.data.Usuarios.filter(function (u) { return u.NoEmpleado === S.usuario.NoEmpleado; })[0] || S.usuario;
      if (S.pantalla === 'app' && !enFoco()) render();
    }).catch(function () {});
    cola = p;
    return p;
  }
  // Bases creadas con la version anterior: completa dias laborables y horarios sin tocar lo demas
  function migrar() {
    var cambio = false;
    if (!S.data.DiasLaborales || !S.data.DiasLaborales.length) {
      S.data.DiasLaborales = QR.DEFAULTS.DiasLaborales.map(function (r) { return { Dia: r[0], Turnos: r[1] }; });
      cambio = true;
    }
    if (!S.data.Festivos) { S.data.Festivos = []; cambio = true; }
    if (!S.data.EstatusRampas || !S.data.EstatusRampas.length) {
      S.data.EstatusRampas = QR.DEFAULTS.EstatusRampas.map(function (r) { return { Estatus: r[0] }; });
      cambio = true;
    }
    S.data.Rampas.forEach(function (r) {
      if (QR.texto(r.Estatus) === '') { r.Estatus = QR.texto(r.Caja) ? 'Por Descargar' : 'Vacía'; cambio = true; }
    });
    S.data.Turnos.forEach(function (t) {
      var d = QR.DEFAULTS.Turnos.filter(function (x) { return x[0] === t.Turno; })[0];
      if (d && (t.Inicio === '' || t.Inicio == null) && (t.Fin === '' || t.Fin == null)) {
        t.Inicio = d[1]; t.Fin = d[2]; cambio = true;
      }
    });
    return cambio;
  }
  async function guardarSiCambio() { try { await guardar(true); } catch (e) { /* se guardara en la siguiente accion */ } }
  // Base de una sola pieza anterior (QORVEX_Recibo.xlsx), si existe. Nunca se modifica.
  async function leerLegado() {
    var fh;
    try { fh = await S.dir.getFileHandle(ARCHIVO, { create: false }); } catch (e) { return null; }
    var buf = await (await fh.getFile()).arrayBuffer();
    return buf.byteLength ? QR.desdeLibro(XLSX, buf) : null;
  }
  async function abrirDatos() {
    S.fh = {}; S.mtimes = {};
    for (var k of Object.keys(QR.ARCHIVOS)) {
      S.fh[k] = await S.dir.getFileHandle(QR.ARCHIVOS[k].archivo, { create: true });
    }
    var sinDatos = await cargarTodo();
    if (sinDatos.indexOf('rampas') >= 0) {
      // Primera vez en esta carpeta: separa la base anterior si existe, o crea los archivos vacios
      var legado = await leerLegado();
      if (legado) {
        aplicarArchivos(QR.repartirLegacy(legado));
        migrar();
        S.msg = { tipo: 'ok', texto: 'Separé la base en los cuatro archivos. ' + ARCHIVO + ' no se modificó.' };
      } else {
        S.msg = { tipo: 'ok', texto: 'Creé los archivos nuevos en la carpeta elegida.' };
      }
      S.escribirTodo = true;
      try { await guardar(true); } finally { S.escribirTodo = false; }
    }
    S.setup = S.data.Usuarios.length === 0;
    S.pantalla = S.setup ? 'crear' : 'login';
  }
  async function elegirCarpetaDatos() {
    var dir = await window.showDirectoryPicker({ mode: 'readwrite' });
    S.dir = dir;
    await kvSet('dirDatos', dir);
    await abrirDatos();
  }
  async function iniciar() {
    if (!window.showDirectoryPicker || !window.indexedDB) {
      S.pantalla = 'carpeta';
      S.msg = { tipo: 'err', texto: 'Este navegador no permite recordar carpetas. Abre la app con Chrome o Edge.' };
      return render();
    }
    var guardado = null;
    try { guardado = await kvGet('dirDatos'); } catch (e) { guardado = null; }
    if (!guardado) { S.pantalla = 'carpeta'; return render(); }
    S.dir = guardado;
    var permiso = await S.dir.queryPermission({ mode: 'readwrite' });
    if (permiso === 'granted') {
      try { await abrirDatos(); } catch (e) { S.msg = { tipo: 'err', texto: 'No pude abrir ' + ARCHIVO + ': ' + e.message }; S.pantalla = 'carpeta'; }
    } else { S.pantalla = 'permiso'; }
    render();
  }

  // ---------- ejecutar acciones ----------
  async function ejecutar(fn, okMsg) {
    try {
      S.msg = null;
      if (!puedeEscribir(AREA_TAB[S.tab] || null)) throw motivoLectura();
      await fn();
      if (okMsg) S.msg = { tipo: 'ok', texto: okMsg };
      await guardar();
    } catch (e) {
      if (typeof e === 'string') {
        S.msg = { tipo: 'err', texto: e };
      } else {
        console.error(e);
        S.msg = { tipo: 'err', texto: 'No se guardó: ' + (e && e.message ? e.message : e) + '. Revisa que el Excel no esté abierto.' };
        try { await recargar(); } catch (_) { /* sin archivo */ }
      }
    }
    render();
  }
  async function intentar(fn) {
    try { S.msg = null; await fn(); }
    catch (e) {
      if (e && e.name === 'AbortError') return render();
      S.msg = { tipo: 'err', texto: typeof e === 'string' ? e : 'Error: ' + (e && e.message ? e.message : e) };
    }
    render();
  }

  // ---------- dominio: piso ----------
  function pisoPorId(id) {
    var r = S.data.Piso.filter(function (x) { return String(x.ID) === String(id); })[0];
    if (!r) throw 'No encuentro esa caja en piso.';
    return r;
  }
  function ocupadosBahia(b) { return S.data.Piso.filter(function (r) { return r.Bahia === b; }).length; }
  function cajaEnPiso(c) { return S.data.Piso.some(function (r) { return r.Caja === c; }); }
  function cajaEnRampa(c) { return S.data.Rampas.some(function (r) { return r.Caja === c; }); }
  function validaCaja(c) {
    var v = QR.texto(c);
    if (!v) throw 'Captura el número de caja o placa.';
    return v;
  }
  function registrarEnPiso(o) {
    var caja = validaCaja(o.caja);
    var max = P().max_cajas_bahia;
    if (!S.data.Bahias.some(function (b) { return b.Bahia === o.bahia; })) throw 'Bahía no válida.';
    if (ocupadosBahia(o.bahia) >= max) throw 'La bahía ' + o.bahia + ' ya tiene el máximo de ' + max + ' cajas.';
    if (cajaEnPiso(caja)) throw 'La caja ' + caja + ' ya está en piso.';
    var fecha = o.fecha || new Date();
    var resp = QR.responsableDe(o.estatus, S.data.Estatus);
    var id = QR.nextId(S.data.Piso.concat(S.data.Historial), 'ID');
    var prov = (S.data.Bahias.filter(function (b) { return b.Bahia === o.bahia; })[0] || {}).Proveedor || '';
    S.data.Piso.push({
      ID: id, Bahia: o.bahia, Caja: caja, Proveedor: prov, Pallets: o.pallets || '', Unidad: o.unidad || 'Caja',
      Estatus: o.estatus, FechaDescarga: fecha, Responsable: resp, InicioResponsable: fecha,
      NoEmpleado: S.usuario.NoEmpleado, Turno: S.turno
    });
    logH({ ID: id, Bahia: o.bahia, Caja: caja, Evento: 'Registro', EstatusNvo: o.estatus, Responsable: resp, Detalle: o.detalle || '', Unidad: o.unidad || 'Caja' });
    return id;
  }
  function cambiarEstatusPiso(id, nuevo) {
    var r = pisoPorId(id);
    if (!S.data.Estatus.some(function (e) { return e.Estatus === nuevo; })) throw 'Estatus no válido.';
    if (r.Estatus === nuevo) return;
    var antes = r.Estatus, respAnt = r.Responsable, respNva = QR.responsableDe(nuevo, S.data.Estatus);
    r.Estatus = nuevo; r.Responsable = respNva;
    if (respNva !== respAnt) r.InicioResponsable = new Date();
    logH({ ID: r.ID, Bahia: r.Bahia, Caja: r.Caja, Evento: 'Cambio estatus', EstatusAnt: antes, EstatusNvo: nuevo, Responsable: respNva });
  }
  function corregirInicioPiso(id, nueva) {
    var r = pisoPorId(id);
    var antes = r.InicioResponsable;
    r.InicioResponsable = nueva;
    logH({ ID: r.ID, Bahia: r.Bahia, Caja: r.Caja, Evento: 'Correccion hora resp.', EstatusAnt: r.Estatus, EstatusNvo: r.Estatus, Responsable: r.Responsable,
      Detalle: 'Inicio resp.: ' + fmtF(antes) + ' -> ' + fmtF(nueva) });
  }
  function sacarDePiso(id, evento) {
    var r = pisoPorId(id);
    logH({ ID: r.ID, Bahia: r.Bahia, Caja: r.Caja, Evento: evento, EstatusAnt: r.Estatus,
      EstatusNvo: evento === 'Ubicado' ? 'Ubicado' : 'Cargada', Responsable: r.Responsable });
    S.data.Piso = S.data.Piso.filter(function (x) { return x !== r; });
  }

  // ---------- dominio: rampas ----------
  function rampaPorCodigo(c) {
    var r = S.data.Rampas.filter(function (x) { return x.Rampa === c; })[0];
    if (!r) throw 'Rampa no encontrada.';
    return r;
  }
  function estatusRampas() { return S.data.EstatusRampas.map(function (e) { return e.Estatus; }); }
  function estatusRampaValido(e) { return estatusRampas().indexOf(e) >= 0; }
  function asignarRampa(c, o) {
    var r = rampaPorCodigo(c);
    var caja = validaCaja(o.caja);
    if (QR.texto(r.Caja)) throw 'La rampa ' + c + ' ya tiene caja. Vacíala primero.';
    if (cajaEnRampa(caja) || cajaEnPiso(caja)) throw 'La caja ' + caja + ' ya está en rampa o en piso.';
    var est = QR.texto(o.estatus);
    if (!est || est === 'Vacía' || !estatusRampaValido(est)) throw 'Elige un estatus de rampa válido.';
    if (!o.fecha) throw 'Captura la fecha y hora de llegada.';
    r.Caja = caja; r.Proveedor = QR.texto(o.proveedor) || 'Insourcing'; r.Estatus = est; r.FechaLlegada = o.fecha;
    r.NoEmpleado = S.usuario.NoEmpleado;
    logH({ Bahia: c, Caja: caja, Evento: 'Rampa asignada', EstatusNvo: est, Detalle: r.Proveedor });
    if (S.abierto) S.abierto['r' + c] = false;
  }
  function cambiarEstatusRampa(c, nuevo) {
    var r = rampaPorCodigo(c);
    if (!estatusRampaValido(nuevo)) throw 'Estatus de rampa no válido.';
    if (nuevo === 'Vacía') return vaciarRampa(c);
    if (!QR.texto(r.Caja)) throw 'La rampa ' + c + ' está vacía. Asigna una caja primero.';
    if (r.Estatus === nuevo) return;
    logH({ Bahia: c, Caja: r.Caja, Evento: 'Rampa estatus', EstatusAnt: r.Estatus, EstatusNvo: nuevo });
    r.Estatus = nuevo; r.NoEmpleado = S.usuario.NoEmpleado;
  }
  function vaciarRampa(c) {
    var r = rampaPorCodigo(c);
    if (!QR.texto(r.Caja)) throw 'La rampa ya está vacía.';
    logH({ Bahia: c, Caja: r.Caja, Evento: 'Rampa vaciada', EstatusAnt: r.Estatus, EstatusNvo: 'Vacía', Detalle: '' });
    r.Caja = ''; r.FechaLlegada = ''; r.Estatus = 'Vacía'; r.NoEmpleado = S.usuario.NoEmpleado;
  }
  function pasarRampaABahia(c, o) {
    var r = rampaPorCodigo(c);
    if (!QR.texto(r.Caja)) throw 'La rampa ' + c + ' está vacía.';
    var caja = r.Caja;
    var id = registrarEnPiso({ bahia: o.bahia, caja: caja, pallets: o.pallets, unidad: o.unidad, estatus: o.estatus, fecha: o.fecha, detalle: 'Desde rampa ' + c });
    r.Caja = ''; r.FechaLlegada = ''; r.Estatus = 'Vacía'; r.NoEmpleado = S.usuario.NoEmpleado;
    logH({ ID: id, Bahia: c, Caja: caja, Evento: 'Rampa vaciada', EstatusNvo: 'Vacía', Detalle: 'Pasó a bahía ' + o.bahia });
  }

  // ---------- dominio: cortos ----------
  function editarCorto(key, campo, valor) {
    var snap = snapActivo();
    if (S.snap && String(S.snap) !== String(snap)) throw 'Estás viendo una versión archivada: solo lectura.';
    if (QR.CAT.camposTrabajoCortos.indexOf(campo) < 0) throw 'Ese campo no se edita aquí.';
    if (campo === 'Estatus' && QR.CAT.estatusCortos.indexOf(valor) < 0) throw 'Estatus no válido.';
    var row = S.data.Cortos.filter(function (r) { return r.Activo === 'Sí' && r.Key === key; })[0];
    if (!row) throw 'No encuentro ese ítem.';
    var antes = Object.assign({}, row);
    row[campo] = valor;
    row.NoEmpleado = S.usuario.NoEmpleado;
    row.FechaActualiz = new Date();
    QR.cambiosCortos(antes, row, row.Snapshot, S.usuario.NoEmpleado, S.turno, new Date(), encargadoActual())
      .forEach(function (c) { S.data.CortosLog.push(c); });
  }
  async function importarArchivoCortos(file) {
    var buf = await file.arrayBuffer();
    var wb = XLSX.read(new Uint8Array(buf), { type: 'array', cellDates: true });
    var visible = function (n) {
      var s = wb.Workbook && wb.Workbook.Sheets ? wb.Workbook.Sheets.filter(function (x) { return x.name === n; })[0] : null;
      return !(s && s.Hidden);
    };
    var nombres = wb.SheetNames.filter(function (n) { return /^\s*\d+\s*$/.test(n) && visible(n); });
    if (!nombres.length) nombres = wb.SheetNames.filter(visible);
    var hojas = nombres.map(function (n) {
      return { name: n.trim(), rows: XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, defval: '', raw: true }) };
    });
    var modo = S.cortosModo;
    var anteriores = S.data.Cortos.filter(function (r) { return r.Activo === 'Sí'; });
    var previas = modo === 'actualizar' ? anteriores : [];
    var snap = maxSnap() + 1;
    var res = QR.importarCortos(hojas, previas, snap, S.usuario.NoEmpleado, new Date());
    if (!res.rows.length) throw 'El archivo no tiene ítems: revisa que las hojas tengan la columna "Item".';
    var previasPorKey = {};
    previas.forEach(function (r) { previasPorKey[r.Key] = r; });
    var ahoraImp = new Date();
    res.rows.forEach(function (r) {
      var prev = previasPorKey[r.Key];
      if (prev) {
        QR.cambiosCortos(prev, r, snap, S.usuario.NoEmpleado, S.turno, ahoraImp, encargadoActual())
          .forEach(function (c) { S.data.CortosLog.push(c); });
      }
    });
    // La versión anterior pasa al registro (CortosHistorico); la hoja Cortos queda solo con el reporte nuevo
    anteriores.forEach(function (r) { r.Activo = 'No'; S.data.CortosHistorico.push(Object.assign({}, r)); });
    S.data.Cortos = res.rows.slice();
    S.data.CortosSnapshots.push({ Snapshot: snap, FechaCarga: new Date(), Archivo: file.name, Filas: res.rows.length, NoEmpleado: S.usuario.NoEmpleado });
    S.snap = null;
    S.preguntaActualizar = false;
    S.msg = { tipo: 'ok', texto: 'Cortos cargados: ' + res.rows.length + ' ítems (' + hojas.map(function (h) { return h.name; }).join(', ') + ')' +
      (previas.length ? '. Versión anterior guardada para consulta; ' + res.arrastrados + ' ítems conservaron su estatus y folio.' : '.') };
  }

  // ---------- dominio: facturas, notas ----------
  // La factura se identifica por su caja / placa (no hay numero de factura aparte).
  function agregarFactura(o) {
    var caja = QR.texto(o.CajaPlaca);
    if (!caja) throw 'Captura la caja o placa.';
    if (S.data.Facturas.some(function (f) { return f.CajaPlaca === caja && f.Eliminada !== 'Sí'; })) throw 'Esa caja / placa ya tiene una factura registrada.';
    var estatus = o.Estatus || 'Pendiente';
    S.data.Facturas.push({ NumFactura: caja, CajaPlaca: caja, Proveedor: QR.texto(o.Proveedor),
      FechaRecepcion: deFecha(o.FechaRecepcion) || new Date(), Estatus: estatus,
      Comentario: QR.texto(o.Comentario), Eliminada: '', NoEmpleado: S.usuario.NoEmpleado, FechaRegistro: new Date() });
    logH({ Evento: 'Factura registrada', Caja: caja, EstatusNvo: estatus });
  }
  function facturaEn(i) {
    var f = S.data.Facturas[i];
    if (!f) throw 'Factura no encontrada.';
    return f;
  }
  function cambiarEstatusFactura(i, nuevo) {
    var f = facturaEn(i);
    if (f.Eliminada === 'Sí') throw 'La factura está eliminada; restaurala para cambiar su estatus.';
    if (ESTADOS_FACTURA.indexOf(nuevo) < 0) throw 'Estatus de factura no válido.';
    if (f.Estatus === nuevo) return;
    logH({ Evento: 'Factura estatus', Caja: f.CajaPlaca, EstatusAnt: f.Estatus, EstatusNvo: nuevo });
    f.Estatus = nuevo;
  }
  // Eliminar no borra la fila: queda marcada como Deleted y se puede restaurar.
  function marcarFactura(i, borrar) {
    var f = facturaEn(i);
    if (borrar === (f.Eliminada === 'Sí')) return;
    f.Eliminada = borrar ? 'Sí' : '';
    logH({ Evento: borrar ? 'Factura eliminada' : 'Factura restaurada', Caja: f.CajaPlaca, EstatusNvo: f.Estatus });
  }
  function agregarNota(o) {
    var caja = validaCaja(o.Caja);
    var nota = { Fecha: new Date(), Carrier: QR.texto(o.Carrier), Caja: caja, Ubicacion: QR.texto(o.Ubicacion),
      Rampa: QR.texto(o.Rampa), Afectacion: QR.texto(o.Afectacion), Comentario: QR.texto(o.Comentario), NoEmpleado: S.usuario.NoEmpleado };
    S.data.Notas.push(nota);
    logH({ Evento: 'Nota registrada', Caja: caja, Detalle: [nota.Carrier, nota.Ubicacion, nota.Rampa, nota.Afectacion].filter(Boolean).join(' · ') });
  }

  // ---------- criticidad y operaciones ----------
  // Cajas activas ordenadas de mayor a menor tiempo (piso: el mayor entre tiempo en piso y tiempo del responsable)
  function criticas() {
    var ahora = new Date(), p = P(), out = [];
    S.data.Piso.forEach(function (r) {
      var hp = QR.horasDesde(r.FechaDescarga, ahora), hr = QR.horasDesde(r.InicioResponsable, ahora);
      var h = Math.max(hp || 0, hr || 0);
      out.push({ caja: r.Caja, lugar: 'Bahía ' + r.Bahia, estatus: r.Estatus, horas: h, estado: QR.estadoTiempo(h, p.umbral_alerta_h, p.umbral_demora_h),
        detalle: r.Estatus + ' · ' + r.Responsable + ' ' + fmtH(hr) + ' · en piso ' + fmtH(hp) });
    });
    S.data.Rampas.filter(function (r) { return QR.texto(r.Caja) !== ''; }).forEach(function (r) {
      var h = QR.horasDesde(r.FechaLlegada, ahora) || 0;
      out.push({ caja: r.Caja, lugar: 'Rampa ' + r.Rampa, estatus: r.Estatus, horas: h, estado: QR.estadoTiempo(h, p.umbral_alerta_h, p.umbral_demora_h),
        detalle: 'Llegada ' + fmtF(r.FechaLlegada) });
    });
    return out.sort(function (a, b) { return b.horas - a.horas; });
  }
  function listaOperaciones() {
    var out = [];
    S.data.Historial.forEach(function (h) {
      out.push({ f: h.FechaEvento, turno: h.Turno || '', modulo: QR.moduloDeEvento(h.Evento), usuario: h.NombreUsuario || nombreDe(h.NoEmpleado),
        evento: h.Evento, lugar: h.Bahia, caja: h.Caja, detalle: [h.EstatusAnt && h.EstatusNvo ? h.EstatusAnt + ' → ' + h.EstatusNvo : h.EstatusNvo, h.Detalle].filter(Boolean).join(' · ') });
    });
    S.data.CortosLog.forEach(function (c) {
      out.push({ f: c.Fecha, turno: c.Turno || '', modulo: 'Cortos', usuario: nombreDe(c.NoEmpleado), evento: 'Cortos · ' + c.Campo,
        lugar: c.Linea, caja: c.CajaPlacas, detalle: (c.ValorAnt || '—') + ' → ' + (c.ValorNvo || '—') + ' (item ' + c.Item + ')' });
    });
    return out.sort(function (a, b) { return tiempo(b.f) - tiempo(a.f); });
  }

  // ---------- pantallas de sesion ----------
  function caja(inner) {
    return '<div class="pantalla"><div class="login"><div class="marca"><img src="assets/ch_logo.png" alt="CH Industrial Solutions"></div>' +
      (S.msg ? '<div class="msg ' + S.msg.tipo + '">' + esc(S.msg.texto) + '</div>' : '') + inner + '</div></div>';
  }
  function pantallaCarga() { return caja('<div class="muted">Cargando QORVEX Recibo…</div>'); }
  function pantallaCarpeta() {
    return caja('<div class="card"><h2>QORVEX Recibo</h2>' +
      '<p class="muted">Elige la carpeta compartida (SharePoint) donde están los archivos de QORVEX Recibo, o donde quieres crearlos. ' +
      'La app la recuerda. Si existe <b>' + ARCHIVO + '</b> de una versión anterior, lo separa en cuatro archivos sin modificarlo.</p>' +
      '<button style="width:100%" data-act="elegirCarpeta">Elegir carpeta de datos</button></div>');
  }
  function pantallaPermiso() {
    return caja('<div class="card"><h2>QORVEX Recibo</h2>' +
      '<p class="muted">Carpeta recordada: <b>' + esc(S.dir && S.dir.name) + '</b>. Dale permiso para continuar.</p>' +
      '<div class="btns"><button style="width:100%" data-act="permiso">Permitir acceso a la carpeta</button>' +
      '<button class="sec" style="width:100%" data-act="elegirCarpeta">Elegir otra carpeta</button></div></div>');
  }
  function pantallaCrear() {
    return caja('<form class="card" data-form="primer-supervisor"><h2>Primer administrador</h2>' +
      '<p class="muted">La base está vacía. Crea la cuenta de administrador inicial; después podrás agregar usuarios en Configuración.</p>' +
      '<label>No. de empleado <input name="noEmp" required></label>' +
      '<label>Nombre a mostrar <input name="nombre" required></label>' +
      '<label>Contraseña (mínimo 4) <input name="pass" type="password" required></label>' +
      '<label>Confirmar contraseña <input name="pass2" type="password" required></label>' +
      '<button type="submit" style="width:100%">Crear y continuar</button></form>');
  }
  function pantallaLogin() {
    return caja('<form class="card" data-form="login"><h2>Iniciar sesión</h2>' +
      '<label>No. de empleado <input name="noEmp" required autocomplete="username"></label>' +
      '<label>Contraseña <input name="pass" type="password" required autocomplete="current-password"></label>' +
      '<button type="submit" style="width:100%">Entrar</button></form>' +
      '<div class="muted" style="margin-top:10px">Carpeta: ' + esc(S.dir && S.dir.name) + '</div>');
  }
  function pantallaRol() {
    var modos = QR.MODOS[S.usuario.Rol] || ['Consulta'];
    var etiq = { Supervisor: 'Entrar como supervisor', Capturista: 'Entrar como capturista',
      Administrador: 'Entrar como administrador', Consulta: 'Entrar en consulta' };
    var sinEnc = supervisores().length === 0;
    return caja('<div class="card"><h2>Hola, ' + esc(S.usuario.Nombre) + '</h2>' +
      '<p class="muted">¿Cómo quieres entrar hoy?</p>' +
      (modos.indexOf('Capturista') >= 0
        ? '<label>Encargado del turno (obligatorio para capturista)<select id="encEntrada">' +
          opcionesEncargado(S.usuario.Rol === 'Supervisor' ? S.usuario.NoEmpleado : '', true) + '</select></label>' +
          (sinEnc ? '<div class="msg err" style="margin-top:8px">No hay supervisores activos. Pide a un supervisor que te dé de alta.</div>' : '')
        : '') +
      '<div class="btns" style="margin-top:12px">' +
      modos.map(function (m) {
        return '<button ' + (m === 'Consulta' ? 'class="sec" ' : '') + 'style="width:100%" data-act="rol" data-rol="' + m + '">' + etiq[m] + '</button>';
      }).join('') +
      '<button class="sec" style="width:100%" data-act="salir">Cancelar</button></div></div>');
  }
  function pantallaBloqueo() {
    return caja('<form class="card" data-form="desbloquear"><h2>Sesión bloqueada</h2>' +
      '<p class="muted">Hola, ' + esc(S.usuario.Nombre) + '. Escribe tu contraseña para seguir donde estabas.</p>' +
      '<label>Contraseña <input name="pass" type="password" required autocomplete="current-password"></label>' +
      '<button type="submit" style="width:100%">Continuar</button>' +
      '<button type="button" class="sec" style="width:100%;margin-top:8px" data-act="salir">Cerrar sesión</button></form>');
  }

  // ---------- pantalla principal: cabecera, cinta y contenido ----------
  function itemsVisibles(g) { return g.items.filter(function (i) { return !i.m || i.m.indexOf(S.rol) >= 0; }); }
  function cintaVisible() { return CINTAS.filter(function (g) { return itemsVisibles(g).length > 0; }); }

  function pantallaApp() {
    var grupos = cintaVisible();
    if (!grupos.some(function (g) { return g.id === S.cinta; })) S.cinta = grupos[0].id;
    var grupo = grupos.filter(function (g) { return g.id === S.cinta; })[0];
    var visibles = itemsVisibles(grupo);
    if (!visibles.some(function (i) { return i.t === S.tab; })) S.tab = visibles.length ? visibles[0].t : 'inicio';
    var cont = VISTAS[S.tab] || tabInicio;
    var to = turnoObj(S.turno);
    var temaOtro = S.tema === 'negro' ? 'blanco' : 'negro';
    return '<header class="top"><div class="marca"><img src="assets/ch_logo.png" alt="CH Industrial Solutions"></div>' +
      '<div><div class="titulo">QORVEX Recibo</div><div class="turno-info">' +
        (S.turno ? 'Turno ' + esc(S.turno) + (to ? ' · ' + hh(to.Inicio) + '–' + hh(to.Fin) : '') + (S.turnoAuto ? ' · automático' : ' · manual') : 'Sin turnos configurados') +
      '</div></div>' +
      '<div class="sesion"><label>Turno <select data-chg="turno"><option value="auto"' + (S.turnoAuto ? ' selected' : '') + '>Auto (reloj)</option>' +
        listaTurnos().map(function (t) { return '<option value="' + esc(t) + '"' + (!S.turnoAuto && t === S.turno ? ' selected' : '') + '>' + esc(t) + '</option>'; }).join('') +
      '</select></label>' +
      '<label>Encargado <select data-chg="encargado">' + opcionesEncargado(S.encargado, false) + '</select></label>' +
      '<span>' + esc(S.usuario.Nombre) + ' · <b>' + esc(S.rol) + '</b> · ' + estadoBloqueo() + '</span>' +
      '<button class="sec peq" data-act="tema">Tema ' + temaOtro + '</button>' +
      '<button class="sec peq" data-act="cambiarCarpeta">Cambiar carpeta</button>' +
      '<button class="sec peq" data-act="salir">Salir</button></div></header>' +
      '<nav class="cinta" aria-label="Menú"><div class="cinta-pestanas">' + grupos.map(function (g) {
        return '<button data-act="cinta" data-cinta="' + esc(g.id) + '" class="' + (g.id === S.cinta ? 'activa' : '') + '">' + esc(g.id) + '</button>';
      }).join('') + '</div><div class="cinta-botones">' + visibles.map(function (i) {
        return '<button class="boton-cinta ' + (S.tab === i.t ? 'activa' : '') + '" data-act="tab" data-tab="' + i.t + '">' + esc(i.n) + '</button>';
      }).join('') + '</div></nav>' +
      '<main>' + (S.msg ? '<div class="msg ' + S.msg.tipo + '">' + esc(S.msg.texto) + '</div>' : '') + cont() + '</main>';
  }

  // ---------- inicio ----------
  function tabInicio() {
    var ahora = new Date(), p = P();
    var piso = S.data.Piso;
    var rampas = S.data.Rampas.filter(function (r) { return QR.texto(r.Caja) !== ''; });
    var resps = [];
    S.data.Estatus.forEach(function (e) { if (resps.indexOf(e.Responsable) < 0) resps.push(e.Responsable); });
    var demora = piso.filter(function (r) { return QR.estadoTiempo(QR.horasDesde(r.InicioResponsable, ahora), p.umbral_alerta_h, p.umbral_demora_h) === 'Demora'; }).length;
    var cri = criticas();
    var top = cri.slice(0, 3);
    var porEstatus = S.data.Estatus.map(function (e) {
      var n = piso.filter(function (r) { return r.Estatus === e.Estatus; }).length;
      return '<div class="kpi" style="border-top-color:' + (COLOR_ESTATUS[e.Estatus] || '#BFBFBF') + '"><b>' + n + '</b><span>' + esc(e.Estatus) + '</span></div>';
    }).join('');
    var bigCard = top.length
      ? '<div style="margin-bottom:10px"><span class="chip ' + claseSem(top[0].estado) + '">' + esc(top[0].estado) + '</span> <b style="font-size:20px">' + esc(top[0].caja) + '</b> · ' + esc(top[0].lugar) +
        ' · <b class="num-h" style="color:#C00000">' + fmtH(top[0].horas) + '</b><div class="muted">' + esc(top[0].detalle) + '</div></div>'
      : '<div class="muted">Sin cajas activas.</div>';
    var listaTop = top.length > 1 ? '<div class="bloque-sub">Top 3</div><ol class="lista-criticas">' + top.map(function (c) {
      return '<li><span class="chip ' + claseSem(c.estado) + '">' + esc(c.estado) + '</span><b>' + esc(c.caja) + '</b><span>' + esc(c.lugar) + '</span>' +
        '<span class="muted">' + esc(c.detalle) + '</span><b style="margin-left:auto">' + fmtH(c.horas) + '</b></li>';
    }).join('') + '</ol>' : '';
    return '<div class="kpis">' +
      kpi(piso.length, 'Cajas en piso') +
      kpi(rampas.length, 'Cajas en rampa') +
      resps.map(function (rp) { return kpi(piso.filter(function (r) { return r.Responsable === rp; }).length, 'Responsable: ' + esc(rp)); }).join('') +
      kpi(demora, 'Responsable en demora', demora ? 'critica' : '') +
      '</div>' +
      '<div class="dos">' +
      '<div class="card"><h2>Cajas por estatus</h2><div class="kpis" style="grid-template-columns:repeat(auto-fill,minmax(130px,1fr));margin:0">' + porEstatus + '</div></div>' +
      '<div class="card"><h2>Caja más crítica</h2>' + bigCard + listaTop + '</div>' +
      '</div>' +
      '<div class="dos">' +
      '<div class="card"><h2>Rampas</h2><div class="muted">' + (rampas.length ? rampas.map(function (r) { return esc(r.Rampa) + ' · ' + esc(r.Caja) + ' · ' + esc(r.Estatus); }).join('<br>') : 'Sin cajas en rampa.') + '</div></div>' +
      '<div class="card"><h2>Turno activo</h2><div>' + (S.turno ? '<b>' + esc(S.turno) + '</b>' : 'Sin turnos configurados') + '</div>' +
        '<div class="muted">' + (S.turnoAuto ? 'Detectado por el reloj. Cambia el turno en la barra superior para fijarlo.' : 'Turno fijado a mano. Elige "Auto (reloj)" para volver al horario.') + '</div></div>' +
      '</div>';
  }

  // ---------- operaciones ----------
  function tabOperaciones() {
    var todas = listaOperaciones();
    var f = S.f.ops;
    var usuarios = [];
    todas.forEach(function (o) { if (usuarios.indexOf(o.usuario) < 0) usuarios.push(o.usuario); });
    usuarios.sort();
    var lista = todas.filter(function (o) {
      return (f.turno === 'Todos' || o.turno === f.turno) && (f.modulo === 'Todos' || o.modulo === f.modulo) && (f.usuario === 'Todos' || o.usuario === f.usuario);
    }).slice(0, 300);
    var filas = lista.map(function (o) {
      return '<tr><td>' + esc(fmtF(o.f)) + '</td><td>' + esc(o.turno) + '</td><td>' + esc(o.modulo) + '</td><td>' + esc(o.usuario) + '</td>' +
        '<td>' + esc(o.evento) + '</td><td>' + esc(o.lugar) + '</td><td><b>' + esc(o.caja) + '</b></td><td>' + esc(o.detalle) + '</td></tr>';
    }).join('');
    return '<div class="card"><div class="row">' +
      '<label>Turno <select data-chg="ops" data-f="turno"><option>Todos</option>' + opts(listaTurnos(), f.turno) + '</select></label>' +
      '<label>Módulo <select data-chg="ops" data-f="modulo"><option>Todos</option>' + opts(['Piso', 'Rampas', 'Facturas', 'Cortos', 'Notas'], f.modulo) + '</select></label>' +
      '<label>Usuario <select data-chg="ops" data-f="usuario"><option>Todos</option>' + opts(usuarios, f.usuario) + '</select></label>' +
      '<div class="muted" style="margin-left:auto">Mostrando ' + lista.length + ' de ' + todas.length + '</div></div>' +
      '<div class="tabla-wrap" style="max-height:620px"><table><thead><tr><th>Fecha</th><th>Turno</th><th>Módulo</th><th>Usuario</th><th>Evento</th>' +
      '<th>Bahía/Rampa/Línea</th><th>Caja</th><th>Detalle</th></tr></thead><tbody>' +
      (filas || '<tr><td colspan="8" class="muted">Sin operaciones con este filtro.</td></tr>') + '</tbody></table></div></div>';
  }

  // ---------- piso (capturista y supervisor) ----------
  function tabPiso() {
    var p = P(), ahora = new Date();
    var bahias = S.data.Bahias.map(function (b) { return b.Bahia; });
    var estatus = S.data.Estatus.map(function (e) { return e.Estatus; });
    var piso = S.data.Piso.slice().sort(function (a, b) {
      return a.Bahia < b.Bahia ? -1 : a.Bahia > b.Bahia ? 1 : tiempo(a.FechaDescarga) - tiempo(b.FechaDescarga);
    });
    var hoy = cuentaTipos(eventosDeRango(jornadaHoy(), jornadaHoy()));
    var filas = piso.map(function (r) {
      var hp = QR.horasDesde(r.FechaDescarga, ahora), hr = QR.horasDesde(r.InicioResponsable, ahora);
      var ep = QR.estadoTiempo(hp, p.umbral_alerta_h, p.umbral_demora_h);
      var er = QR.estadoTiempo(hr, p.umbral_alerta_h, p.umbral_demora_h);
      var k = 'p' + r.ID, abierto = !!(S.abierto && S.abierto[k]);
      var fila = '<tr><td>' + esc(r.ID) + '</td><td>' + esc(r.Bahia) + '</td><td><b>' + esc(r.Caja) + '</b></td>' +
        '<td>' + esc(r.Proveedor) + '</td><td>' + esc(r.Pallets) + ' ' + esc(r.Unidad) + '</td>' +
        '<td><select class="estatus-sel" data-chg="piso-estatus" data-id="' + esc(r.ID) + '" style="background:' + (COLOR_ESTATUS[r.Estatus] || '#fff') + ';min-width:130px">' + opts(estatus, r.Estatus) + '</select></td>' +
        '<td>' + esc(r.Responsable) + '</td>' +
        '<td><span class="chip ' + claseSem(ep) + '">' + fmtH(hp) + '</span></td>' +
        '<td><span class="chip ' + claseSem(er) + '">' + fmtH(hr) + '</span></td>' +
        '<td>' + esc(r.Turno) + '</td>' +
        '<td>' + botonMas(k, abierto) + '</td></tr>';
      if (!abierto) return fila;
      return fila + '<tr class="fila-extra"><td colspan="11"><div class="extra-acciones">' +
        '<div class="row" style="flex-wrap:nowrap">' +
        '<button class="sec peq" data-act="corregir" data-id="' + esc(r.ID) + '">Hora resp.</button>' +
        '<button class="sec peq" data-act="ubicar" data-id="' + esc(r.ID) + '">Ubicar</button>' +
        '<button class="sec peq" data-act="cargar" data-id="' + esc(r.ID) + '">Cargar</button></div></div></td></tr>';
    }).join('');
    var demora = piso.filter(function (r) { return QR.estadoTiempo(QR.horasDesde(r.InicioResponsable, ahora), p.umbral_alerta_h, p.umbral_demora_h) === 'Demora'; }).length;
    return '<div class="kpis">' +
      kpi(piso.length, 'Cajas en piso') +
      kpi(hoy.Descargados, 'Descargadas en la jornada') +
      kpi(hoy.Ubicados, 'Ubicadas en la jornada') +
      kpi(hoy.Cargados, 'Cargadas en la jornada') +
      kpi(demora, 'Responsable en demora (&gt;' + p.umbral_demora_h + ' h)', demora ? 'critica' : '') +
      '</div>' +
      '<div class="card"><h2>Registrar caja en bahía</h2>' +
      '<form data-form="piso-alta" class="row">' +
      '<label>Bahía <select name="bahia">' + opts(bahias, '') + '</select></label>' +
      '<label>Caja / placa <input name="caja" required autocomplete="off"></label>' +
      '<label>Pallets <input name="pallets" type="number" min="0"></label>' +
      '<label>Unidad <select name="unidad">' + opts(QR.CAT.unidades, 'Caja') + '</select></label>' +
      '<label>Estatus <select name="estatus">' + opts(estatus, 'Por Validar') + '</select></label>' +
      '<label>Descarga (fecha y hora) <input name="fecha" type="datetime-local" value="' + aInput(ahora) + '"></label>' +
      '<button type="submit">Registrar</button></form>' +
      '<div class="muted" style="margin-top:8px">Máximo ' + p.max_cajas_bahia + ' cajas por bahía. El responsable se toma del estatus. Turno activo: ' + esc(S.turno) + '.</div></div>' +
      '<div class="card"><h2>Cajas en piso</h2><div class="tabla-wrap"><table class="compacta"><thead><tr>' +
      '<th>ID</th><th>Bahía</th><th>Caja</th><th>Proveedor</th><th>Pallets</th><th>Estatus</th><th>Responsable</th>' +
      '<th>Hrs piso</th><th>Hrs resp.</th><th>Turno</th><th>Acciones</th></tr></thead><tbody>' +
      (filas || '<tr><td colspan="11" class="muted">Sin cajas en piso.</td></tr>') + '</tbody></table></div></div>';
  }

  // Boton que muestra u oculta las acciones extra de una fila o tarjeta
  function botonMas(k, abierto) {
    return '<button class="sec peq" data-act="expandir" data-k="' + esc(k) + '">' + (abierto ? '▲ Menos' : '▼ Más') + '</button>';
  }

  // ---------- rampas ----------
  function tabRampas() {
    var p = P(), ahora = new Date();
    var ocupadas = S.data.Rampas.filter(function (r) { return QR.texto(r.Caja) !== ''; });
    var asignables = estatusRampas().filter(function (e) { return e !== 'Vacía'; });
    var tarjetas = S.data.Rampas.map(function (r) {
      var vacia = QR.texto(r.Caja) === '';
      if (vacia) {
        var kv = 'r' + r.Rampa, abiertoV = !!(S.abierto && S.abierto[kv]);
        return '<div class="rampa vacia"><div class="muted">' + esc(r.Rampa) + '</div><div class="caja">VACÍA</div>' +
          '<div style="margin-top:6px">' + botonMas(kv, abiertoV) + '</div>' +
          (abiertoV ? '<form data-form="rampa-asignar" data-rampa="' + esc(r.Rampa) + '" class="pila">' +
          '<label>Caja / placa <input name="caja" autocomplete="off" required></label>' +
          '<label>Proveedor <input name="proveedor" value="Insourcing"></label>' +
          '<label>Estatus <select name="estatus">' + opts(asignables, 'Por Descargar') + '</select></label>' +
          '<label>Llegada <input name="fecha" type="datetime-local" value="' + aInput(ahora) + '" required></label>' +
          '<button type="submit">Asignar a rampa</button></form>' : '') + '</div>';
      }
      var h = QR.horasDesde(r.FechaLlegada, ahora);
      var est = QR.estadoTiempo(h, p.umbral_alerta_h, p.umbral_demora_h);
      var cls = est === 'En tiempo' ? 'ok' : est === 'Alerta' ? 'alerta' : 'demora';
      var k = 'r' + r.Rampa, abierto = !!(S.abierto && S.abierto[k]);
      return '<div class="rampa ' + cls + '"><div class="muted">' + esc(r.Rampa) + ' · ' + esc(r.Proveedor || '') + '</div>' +
        '<div class="caja">' + esc(r.Caja) + '</div>' +
        '<div><span class="chip ' + (r.Estatus === 'Por Descargar' ? 'sem-alerta' : 'sem-ok') + '">' + esc(r.Estatus || '—') + '</span></div>' +
        '<div class="muted">Llegada ' + esc(fmtF(r.FechaLlegada)) + ' · ' + fmtH(h) + '</div>' +
        '<div style="margin-top:6px">' + botonMas(k, abierto) + '</div>' +
        (abierto ? '<form data-form="rampa-estatus" data-rampa="' + esc(r.Rampa) + '" class="pila">' +
          '<label>Estatus <select name="estatus">' + opts(estatusRampas(), r.Estatus) + '</select></label>' +
          '<div class="row" style="flex-wrap:nowrap"><button class="peq">Cambiar estatus</button>' +
          '<button type="button" class="sec peq" data-act="rampaVaciar" data-rampa="' + esc(r.Rampa) + '">Vaciar</button></div></form>' : '') + '</div>';
    }).join('');
    var demora = ocupadas.filter(function (r) { return QR.estadoTiempo(QR.horasDesde(r.FechaLlegada, ahora), p.umbral_alerta_h, p.umbral_demora_h) === 'Demora'; }).length;
    var masCritica = ocupadas.slice().sort(function (a, b) { return (QR.horasDesde(b.FechaLlegada, ahora) || 0) - (QR.horasDesde(a.FechaLlegada, ahora) || 0); })[0];
    var porEstatus = asignables.map(function (e) {
      return kpi(ocupadas.filter(function (r) { return r.Estatus === e; }).length, esc(e));
    }).join('');
    return '<div class="kpis">' +
      kpi(ocupadas.length, 'Rampas con caja') +
      kpi(S.data.Rampas.length - ocupadas.length, 'Rampas vacías') +
      kpi(demora, 'Rampas en demora', demora ? 'critica' : '') +
      kpi(masCritica ? esc(masCritica.Rampa) + ' · ' + fmtH(QR.horasDesde(masCritica.FechaLlegada, ahora)) : '—', 'Rampa más crítica') +
      porEstatus +
      '</div>' +
      '<div class="card"><h2>Rampas</h2><div class="grid-rampas">' + tarjetas + '</div></div>' +
      '<div class="card"><h2>Pasar caja de rampa a bahía</h2>' +
      (ocupadas.length ? '<form data-form="rampa-pasar" class="row">' +
        '<label>Rampa <select name="rampa">' + opts(ocupadas.map(function (r) { return r.Rampa; }), '') + '</select></label>' +
        '<label>Bahía <select name="bahia">' + opts(S.data.Bahias.map(function (b) { return b.Bahia; }), '') + '</select></label>' +
        '<label>Estatus <select name="estatus">' + opts(S.data.Estatus.map(function (e) { return e.Estatus; }), 'Por Validar') + '</select></label>' +
        '<label>Pallets <input name="pallets" type="number" min="0"></label>' +
        '<label>Unidad <select name="unidad">' + opts(QR.CAT.unidades, 'Caja') + '</select></label>' +
        '<label>Descarga <input name="fecha" type="datetime-local" value="' + aInput(ahora) + '"></label>' +
        '<button type="submit">Pasar a bahía</button></form>'
        : '<div class="muted">No hay cajas en rampa.</div>') + '</div>';
  }

  // ---------- cortos (supervisor) ----------
  function tabCortos() {
    var act = snapActivo();
    var ver = S.snap || act;
    var solo = !!ver && ver !== act;
    var snaps = S.data.CortosSnapshots.slice().sort(function (a, b) { return +b.Snapshot - +a.Snapshot; });
    var info = snaps.filter(function (s) { return String(s.Snapshot) === String(ver); })[0];
    var filas = S.data.Cortos.concat(S.data.CortosHistorico || []).filter(function (r) { return String(r.Snapshot) === String(ver); });
    var lineas = filas.map(function (r) { return r.Linea; }).filter(function (v, i, a) { return a.indexOf(v) === i; }).sort();
    var q = S.f.buscar.toUpperCase();
    var vista = filas.filter(function (r) {
      return (S.f.linea === 'Todas' || r.Linea === S.f.linea) &&
        (S.f.estatus === 'Todos' || r.Estatus === S.f.estatus) &&
        (!q || (r.Item + ' ' + r.CajaPlacas + ' ' + r.Descr + ' ' + r.Proveedor).toUpperCase().indexOf(q) >= 0);
    });
    var total = filas.length, comp = filas.filter(function (r) { return r.Estatus === 'Completo'; }).length;
    var cnt = function (e) { return filas.filter(function (r) { return r.Estatus === e; }).length; };
    var conteo = QR.CAT.estatusCortos.map(function (e) {
      return '<span class="chip estatus-sel" style="background:' + COLOR_ESTATUS[e] + '">' + esc(e) + ': ' + cnt(e) + '</span>';
    }).join(' ');

    var filasHtml = vista.map(function (r) {
      var dis = solo || !puedeEscribir('cortos') ? ' disabled' : '';
      var sel = ' class="estatus-sel" style="background:' + (COLOR_ESTATUS[r.Estatus] || '#fff') + '" data-chg="cortos" data-key="' + esc(r.Key) + '" data-campo="Estatus"' + dis;
      return '<tr><td>' + esc(r.Linea) + '</td><td>' + esc(r.Item) + '</td><td>' + esc(r.Descr) + '</td><td>' + esc(r.Proveedor) + '</td>' +
        '<td>' + esc(fmtD(r.Fecha)) + '</td><td>' + esc(r.Prioridad) + '</td><td>' + esc(r.Hora) + '</td><td>' + esc(r.CajaPlacas) + '</td>' +
        '<td><input' + dis + ' style="width:86px" value="' + esc(r.BahiaRampa) + '" data-chg="cortos" data-key="' + esc(r.Key) + '" data-campo="BahiaRampa"></td>' +
        '<td><select' + sel + '>' + opts(QR.CAT.estatusCortos, r.Estatus) + '</select></td>' +
        '<td><input' + dis + ' style="width:58px" value="' + esc(r.QtyEntrega) + '" data-chg="cortos" data-key="' + esc(r.Key) + '" data-campo="QtyEntrega"></td>' +
        '<td><input' + dis + ' style="width:80px" value="' + esc(r.Folio) + '" data-chg="cortos" data-key="' + esc(r.Key) + '" data-campo="Folio"></td>' +
        '<td><input' + dis + ' style="width:220px" value="' + esc(r.ComentariosRecibo) + '" data-chg="cortos" data-key="' + esc(r.Key) + '" data-campo="ComentariosRecibo"></td>' +
        '<td><input' + dis + ' style="width:96px" value="' + esc(r.TipoCaja) + '" data-chg="cortos" data-key="' + esc(r.Key) + '" data-campo="TipoCaja"></td></tr>';
    }).join('');

    var dash = total
      ? '<div class="kpis">' + kpi(total, 'Ítems') + kpi(comp, 'Completos') + kpi(total ? Math.round(comp / total * 100) + '%' : '0%', 'Avance') +
        kpi(cnt('Recolectar'), 'Por recolectar') + kpi(cnt('Sin Información'), 'Sin información', cnt('Sin Información') ? 'critica' : '') +
        kpi(cnt('Solicitar caja') + cnt('Caja Solicitada'), 'Solicitar / solicitada') + '</div>' +
        '<div class="card" style="padding:10px 14px">' + conteo + '</div>'
      : '';

    return '<div class="card"><div class="row"><div><h2>Reporte de cortos</h2>' +
      '<div class="muted">' + (info ? 'Versión #' + esc(info.Snapshot) + ' · cargada ' + esc(fmtF(info.FechaCarga)) + ' · ' + esc(info.Archivo) + ' · ' + info.Filas + ' ítems' +
        (solo ? ' · ARCHIVADA (solo lectura)' : ' · activa') : 'Sin cortos cargados') + '</div></div>' +
      '<div class="row" style="margin-left:auto;justify-content:flex-end">' +
      (act ? '<button class="sec" data-act="cortosPreguntar">Actualizar cortos</button>' : '<button data-act="cortosPrimera">Cargar cortos</button>') +
      (snaps.length ? '<label>Ver versión <select data-chg="snap">' + snaps.map(function (s) {
        return '<option value="' + esc(s.Snapshot) + '"' + (String(s.Snapshot) === String(ver) ? ' selected' : '') + '>#' + esc(s.Snapshot) + ' · ' + esc(fmtD(s.FechaCarga)) + ' · ' + esc(s.Archivo) + (String(s.Snapshot) === String(act) ? ' (activa)' : '') + '</option>';
      }).join('') + '</select></label>' : '') + '</div></div>' +
      (S.preguntaActualizar ? '<div class="msg info" style="margin-top:12px"><b>¿Quieres actualizar los cortos?</b> La versión actual se guarda para consulta y métricas. ' +
        '<div class="row" style="margin-top:8px"><button data-act="cortosSeleccionar">Sí, seleccionar archivo nuevo</button><button class="sec" data-act="cortosNo">No</button></div></div>' : '') +
      '<input type="file" id="archivoCortos" accept=".xlsx,.xlsm" hidden data-chg="archivoCortos"></div>' +
      dash +
      '<div class="card"><div class="row" style="margin-bottom:10px">' +
      '<label>Línea <select data-chg="filtro" data-f="linea"><option' + (S.f.linea === 'Todas' ? ' selected' : '') + '>Todas</option>' + opts(lineas, S.f.linea) + '</select></label>' +
      '<label>Estatus <select data-chg="filtro" data-f="estatus"><option>Todos</option>' + opts(QR.CAT.estatusCortos, S.f.estatus) + '</select></label>' +
      '<label>Buscar item / caja <input data-chg="filtro" data-f="buscar" value="' + esc(S.f.buscar) + '"></label></div>' +
      '<div class="tabla-wrap"><table><thead><tr><th>Línea</th><th>Item</th><th>Descripción</th><th>Proveedor</th><th>Fecha</th><th>Prio</th><th>Hora</th><th>Caja/Placas</th>' +
      '<th>Bahía/Rampa</th><th>Estatus</th><th>Qty</th><th>Folio</th><th>Comentarios recibo</th><th>Tipo caja</th></tr></thead><tbody>' +
      (filasHtml || '<tr><td colspan="14" class="muted">Sin ítems con este filtro.</td></tr>') + '</tbody></table></div>' +
      '<div class="muted" style="margin-top:8px">Las columnas de Item a Caja/Placas son del cliente (solo lectura). Las de Bahía/Rampa a Tipo de caja son de trabajo.</div></div>';
  }

  // ---------- facturas (capturista y supervisor) ----------
  function tabFacturas() {
    var q = String(S.f.bf || '').toUpperCase();
    var hoyS = aFecha(new Date());
    var todas = S.data.Facturas.map(function (f, i) { return { f: f, i: i }; });
    var activas = todas.filter(function (o) { return o.f.Eliminada !== 'Sí'; });
    var lista = todas.filter(function (o) {
      return !q || (o.f.CajaPlaca + ' ' + o.f.Proveedor).toUpperCase().indexOf(q) >= 0;
    });
    var cnt = function (e) { return activas.filter(function (o) { return o.f.Estatus === e; }).length; };
    var dash = '<div class="kpis">' + kpi(activas.length, 'Facturas activas') + kpi(cnt('Pendiente'), 'Pendientes') +
      kpi(cnt('Recibida'), 'Recibidas') + kpi(cnt('Conciliada'), 'Conciliadas') +
      kpi(activas.filter(function (o) { return aFecha(o.f.FechaRegistro instanceof Date ? o.f.FechaRegistro : new Date(0)) === hoyS; }).length, 'Registradas hoy') +
      kpi(todas.length - activas.length, 'Eliminadas') + '</div>';
    var filas = lista.slice().reverse().map(function (o) {
      var f = o.f, borrada = f.Eliminada === 'Sí';
      var accion = borrada
        ? '<span class="chip" style="background:#C00000;color:#fff">Deleted</span> <button class="sec peq" data-act="facturaRestaurar" data-i="' + o.i + '">Restaurar</button>'
        : '<button class="sec peq" data-act="facturaEliminar" data-i="' + o.i + '">Eliminar</button>';
      return '<tr class="' + (borrada ? 'borrada' : '') + '"><td><b>' + esc(f.CajaPlaca) + '</b></td><td>' + esc(f.Proveedor) + '</td>' +
        '<td>' + esc(fmtD(f.FechaRecepcion)) + '</td><td>' + esc(f.Comentario) + '</td>' +
        '<td><select class="estatus-sel" data-chg="factura" data-i="' + o.i + '"' + (borrada ? ' disabled' : '') + '>' + opts(ESTADOS_FACTURA, f.Estatus) + '</select></td>' +
        '<td>' + esc(nombreDe(f.NoEmpleado)) + '</td><td>' + accion + '</td></tr>';
    }).join('');
    return dash + '<div class="card"><h2>Registrar factura</h2>' +
      '<form data-form="factura-alta" class="row">' +
      '<label>Caja / placa <input name="CajaPlaca" required></label>' +
      '<label>Proveedor <input name="Proveedor"></label>' +
      '<label>Fecha recepción <input name="FechaRecepcion" type="date" value="' + aFecha(new Date()) + '"></label>' +
      '<label>Estatus <select name="Estatus">' + opts(ESTADOS_FACTURA, 'Pendiente') + '</select></label>' +
      '<label>Comentario <input name="Comentario"></label>' +
      '<button type="submit">Guardar</button></form></div>' +
      '<div class="card"><div class="row" style="margin-bottom:10px"><label>Buscar <input data-chg="filtro" data-f="bf" value="' + esc(S.f.bf || '') + '"></label></div>' +
      '<div class="tabla-wrap"><table class="compacta"><thead><tr><th>Caja / placa</th><th>Proveedor</th><th>Recepción</th><th>Comentario</th><th>Estatus</th><th>Registró</th><th>Acción</th></tr></thead><tbody>' +
      (filas || '<tr><td colspan="7" class="muted">Sin facturas.</td></tr>') + '</tbody></table></div></div>';
  }

  // ---------- notas (supervisor) ----------
  function tabNotas() {
    var hoyS = aFecha(new Date());
    var cajasDist = S.data.Notas.map(function (n) { return n.Caja; }).filter(function (v, i, a) { return a.indexOf(v) === i; }).length;
    var porCarrier = {};
    S.data.Notas.forEach(function (n) { var c = n.Carrier || 'Sin carrier'; porCarrier[c] = (porCarrier[c] || 0) + 1; });
    var topC = Object.keys(porCarrier).sort(function (a, b) { return porCarrier[b] - porCarrier[a]; })[0];
    var dash = '<div class="kpis">' + kpi(S.data.Notas.length, 'Notas') +
      kpi(S.data.Notas.filter(function (n) { return aFecha(n.Fecha instanceof Date ? n.Fecha : new Date(0)) === hoyS; }).length, 'Notas hoy') +
      kpi(cajasDist, 'Cajas solicitadas / distintas') + kpi(topC ? esc(topC) + ' (' + porCarrier[topC] + ')' : '—', 'Carrier con más notas') + '</div>';
    var filas = S.data.Notas.slice().reverse().map(function (n) {
      return '<tr><td>' + esc(fmtF(n.Fecha)) + '</td><td>' + esc(n.Carrier) + '</td><td><b>' + esc(n.Caja) + '</b></td><td>' + esc(n.Ubicacion) +
        '</td><td>' + esc(n.Rampa) + '</td><td>' + esc(n.Afectacion) + '</td><td>' + esc(n.Comentario) + '</td><td>' + esc(nombreDe(n.NoEmpleado)) + '</td></tr>';
    }).join('');
    return dash + '<div class="card"><h2>Nueva nota</h2>' +
      '<p class="muted">Cajas solicitadas y avisos de carrier, por ejemplo: HG · 1445 · CHARQUEÑO · R749 · 13:30.</p>' +
      '<form data-form="nota-alta" class="row">' +
      '<label>Carrier <input name="Carrier"></label>' +
      '<label>Caja <input name="Caja" required></label>' +
      '<label>Ubicación <input name="Ubicacion"></label>' +
      '<label>Rampa <input name="Rampa"></label>' +
      '<label>Afectación / hora <input name="Afectacion"></label>' +
      '<label>Comentario <input name="Comentario"></label>' +
      '<button type="submit">Agregar</button></form></div>' +
      '<div class="card"><div class="tabla-wrap"><table><thead><tr><th>Fecha</th><th>Carrier</th><th>Caja</th><th>Ubicación</th><th>Rampa</th><th>Afectación</th><th>Comentario</th><th>Registró</th></tr></thead><tbody>' +
      (filas || '<tr><td colspan="8" class="muted">Sin notas.</td></tr>') + '</tbody></table></div></div>';
  }

  // ---------- KPIs ----------
  function bloqueMiTurno() {
    var turno = S.turno, dia = jornadaHoy();
    if (!turno) return '<div class="card"><h2>Mi turno hoy</h2><div class="muted">Configura los turnos en Configuración.</div></div>';
    var ev = eventosDeRango(dia, dia);
    var yo = ev.filter(function (e) { return e.noEmp === S.usuario.NoEmpleado && e.turno === turno; });
    var equipo = ev.filter(function (e) { return e.turno === turno; });
    var cy = cuentaTipos(yo), ce = cuentaTipos(equipo);
    var meta = metaDelDia(turno, dia), mt = sumaMeta(meta);
    var aviso = trabajaTurno(dia, turno) ? '' : '<div class="msg info">Hoy no es día laborable para el turno ' + esc(turno) + ' (festivo o fuera de los días configurados): la meta es 0.</div>';
    var fila = function (etq, c) {
      var r = suma(c);
      return '<tr><td><b>' + etq + '</b></td>' + celdasTipos(c) + '<td class="num"><b>' + r + '</b></td></tr>';
    };
    return '<div class="card"><h2>Mi turno hoy · ' + esc(turno) + '</h2>' +
      '<div class="muted" style="margin-bottom:10px">Jornada del ' + esc(aFecha(dia)) + ' · se actualiza al cambiar el turno.</div>' + aviso +
      '<div class="tabla-wrap"><table><thead><tr><th>Concepto</th><th class="num">Descargados</th><th class="num">Recolectados</th><th class="num">Ubicados</th><th class="num">Cargados</th><th class="num">Total</th></tr></thead><tbody>' +
      fila('Yo', cy) + fila('Turno (equipo)', ce) +
      '<tr><td><b>Meta del turno</b></td>' + TIPOS.map(function (k) { return '<td class="num">' + meta[k] + '</td>'; }).join('') + '<td class="num">' + mt + '</td></tr>' +
      '<tr><td><b>% logrado (equipo)</b></td>' + TIPOS.map(function (k) { return '<td class="num">' + pct(ce[k], meta[k]) + '</td>'; }).join('') +
      '<td class="num">' + pct(suma(ce), mt) + '</td></tr>' +
      '</tbody></table></div></div>';
  }
  function bloqueReporteDiario() {
    var dia = deFecha(S.f.kpiDia) || inicioDia(new Date());
    var ev = eventosDeRango(dia, dia);
    var festivo = QR.esFestivo(dia, S.data.Festivos);
    var porUsuario = {};
    ev.forEach(function (e) { var k = e.enc || ''; (porUsuario[k] = porUsuario[k] || []).push(e); });
    var filasU = Object.keys(porUsuario).sort(function (a, b) { return nombreEnc(a).localeCompare(nombreEnc(b)); }).map(function (no) {
      var lst = porUsuario[no], c = cuentaTipos(lst), real = suma(c);
      var conTurno = {};
      lst.forEach(function (e) { if (e.turno) conTurno[e.turno] = (conTurno[e.turno] || 0) + 1; });
      var tps = Object.keys(conTurno);
      var m = tps.reduce(function (s, t) { return s + sumaMeta(metaDelDia(t, dia)); }, 0);
      return '<tr><td><b>' + esc(nombreEnc(no)) + '</b></td><td>' + esc(tps.join(', ') || '—') + '</td>' + celdasTipos(c) +
        '<td class="num"><b>' + real + '</b></td><td class="num">' + m + '</td><td class="num">' + pct(real, m) + '</td></tr>';
    }).join('');
    var metaTotal = 0;
    var filasT = S.data.Turnos.map(function (t) {
      var lst = ev.filter(function (e) { return e.turno === t.Turno; }), c = cuentaTipos(lst), real = suma(c);
      var m = sumaMeta(metaDelDia(t.Turno, dia)); metaTotal += m;
      return '<tr><td><b>' + esc(t.Turno) + '</b></td><td>' + (trabajaTurno(dia, t.Turno) ? 'Sí' : 'No') + '</td>' + celdasTipos(c) +
        '<td class="num"><b>' + real + '</b></td><td class="num">' + m + '</td><td class="num">' + pct(real, m) + '</td></tr>';
    }).join('');
    var general = cuentaTipos(ev);
    return '<div class="card"><div class="row"><div><h2>Reporte diario</h2><div class="muted">Jornada del ' + esc(aFecha(dia)) +
      (festivo ? ' · FESTIVO: no cuenta en metas' : '') + '</div></div>' +
      '<label style="margin-left:auto">Día <input type="date" data-chg="kpi" data-f="kpiDia" value="' + esc(aFecha(dia)) + '"></label></div></div>' +
      '<div class="dos"><div class="card"><h2>General por turno</h2><div class="tabla-wrap"><table><thead><tr><th>Turno</th><th>Laborable</th>' +
      '<th class="num">Desc.</th><th class="num">Recol.</th><th class="num">Ubic.</th><th class="num">Carg.</th><th class="num">Total</th><th class="num">Meta</th><th class="num">%</th></tr></thead><tbody>' +
      filasT + '<tr><td><b>Total</b></td><td></td>' + celdasTipos(general) + '<td class="num"><b>' + suma(general) + '</b></td><td class="num">' + metaTotal + '</td><td class="num">' + pct(suma(general), metaTotal) + '</td></tr>' +
      '</tbody></table></div></div>' +
      '<div class="card"><h2>Por encargado (supervisor del turno)</h2><div class="tabla-wrap"><table><thead><tr><th>Encargado</th><th>Turno</th>' +
      '<th class="num">Desc.</th><th class="num">Recol.</th><th class="num">Ubic.</th><th class="num">Carg.</th><th class="num">Total</th><th class="num">Meta</th><th class="num">%</th></tr></thead><tbody>' +
      (filasU || '<tr><td colspan="9" class="muted">Sin movimientos ese día.</td></tr>') + '</tbody></table></div></div></div>';
  }
  function bloqueMensual() {
    var ahora = new Date();
    var mm = deMes(S.f.kpiMes) || { anio: ahora.getFullYear(), mes: ahora.getMonth() };
    var desde = new Date(mm.anio, mm.mes, 1), hasta = new Date(mm.anio, mm.mes + 1, 0);
    var ev = eventosDeRango(desde, hasta);
    var DL = S.data.DiasLaborales, FE = S.data.Festivos;
    var dias = QR.diasPorTurno(desde, hasta, DL, FE), nc = QR.diasNoContados(desde, hasta, DL, FE);
    var totalMeta = 0, totalC = { Descargados: 0, Recolectados: 0, Ubicados: 0, Cargados: 0 };
    var filasT = S.data.Turnos.map(function (t) {
      var n = dias[t.Turno] || 0, m = metaDe(t.Turno);
      var metaT = sumaMeta(m) * n; totalMeta += metaT;
      var c = cuentaTipos(ev.filter(function (e) { return e.turno === t.Turno; }));
      TIPOS.forEach(function (k) { totalC[k] += c[k]; });
      return '<tr><td><b>' + esc(t.Turno) + '</b></td><td class="num">' + n + '</td><td class="num">' + metaT + '</td>' + celdasTipos(c) +
        '<td class="num"><b>' + suma(c) + '</b></td><td class="num">' + pct(suma(c), metaT) + '</td></tr>';
    }).join('');
    var sinTurno = cuentaTipos(ev.filter(function (e) { return !listaTurnos().includes(e.turno); }));
    var filasSinT = suma(sinTurno) ? '<tr><td class="muted">Sin turno válido</td><td></td><td></td>' + celdasTipos(sinTurno) + '<td class="num">' + suma(sinTurno) + '</td><td></td></tr>' : '';
    var usuarios = {};
    ev.forEach(function (e) { var k = e.enc || ''; (usuarios[k] = usuarios[k] || []).push(e); });
    var filasU = Object.keys(usuarios).sort(function (a, b) { return nombreEnc(a).localeCompare(nombreEnc(b)); }).map(function (no) {
      var c = cuentaTipos(usuarios[no]);
      return '<tr><td><b>' + esc(nombreEnc(no)) + '</b></td>' + celdasTipos(c) + '<td class="num"><b>' + suma(c) + '</b></td></tr>';
    }).join('');
    var totalReal = suma(totalC);
    return '<div class="card"><div class="row"><div><h2>Métrico mensual</h2><div class="muted">' +
      'Días del mes: ' + hasta.getDate() + ' · Festivos excluidos: ' + nc.festivos + ' · Días no laborables excluidos: ' + nc.noLaborables +
      ' · La meta = meta por turno × días laborables del turno.</div></div>' +
      '<label style="margin-left:auto">Mes <input type="month" data-chg="kpi" data-f="kpiMes" value="' + esc(aMes(desde)) + '"></label></div></div>' +
      '<div class="dos"><div class="card"><h2>Por turno</h2><div class="tabla-wrap"><table><thead><tr><th>Turno</th><th class="num">Días</th><th class="num">Meta</th>' +
      '<th class="num">Desc.</th><th class="num">Recol.</th><th class="num">Ubic.</th><th class="num">Carg.</th><th class="num">Total</th><th class="num">%</th></tr></thead><tbody>' +
      filasT + filasSinT + '<tr><td><b>Total</b></td><td></td><td class="num">' + totalMeta + '</td>' + celdasTipos(totalC) +
      '<td class="num"><b>' + totalReal + '</b></td><td class="num">' + pct(totalReal, totalMeta) + '</td></tr></tbody></table></div></div>' +
      '<div class="card"><h2>Por encargado</h2><div class="tabla-wrap"><table><thead><tr><th>Encargado</th><th class="num">Desc.</th><th class="num">Recol.</th>' +
      '<th class="num">Ubic.</th><th class="num">Carg.</th><th class="num">Total</th></tr></thead><tbody>' +
      (filasU || '<tr><td colspan="6" class="muted">Sin movimientos en el mes.</td></tr>') + '</tbody></table></div></div></div>';
  }
  // Entrega de turno: foto del turno para imprimir (movimientos, activo en bahias y rampas, cajas con mas tiempo)
  function bloqueEntrega() {
    var dia = deFecha(S.f.entDia) || inicioDia(new Date());
    var sel = S.f.entTurno || 'Todos';
    if (sel !== 'Todos' && listaTurnos().indexOf(sel) < 0) sel = 'Todos';
    var ev = eventosDeRango(dia, dia).filter(function (e) { return sel === 'Todos' || e.turno === sel; });
    var n = function (t) { return ev.filter(function (e) { return e.tipo === t; }).length; };
    var meta = { Descargados: 0, Recolectados: 0, Ubicados: 0, Cargados: 0 };
    (sel === 'Todos' ? listaTurnos() : [sel]).forEach(function (t) {
      var mt = metaDelDia(t, dia);
      TIPOS.forEach(function (k) { meta[k] += mt[k]; });
    });
    var encs = [];
    ev.forEach(function (e) { var k = e.enc || ''; if (encs.indexOf(k) < 0) encs.push(k); });
    var filas = [
      ['Cajas descargadas', n('Descargados'), meta.Descargados],
      ['De ellas, camionetas descargadas', n('Camionetas'), ''],
      ['Procesadas (cambios de estatus)', n('Procesadas'), ''],
      ['Ubicadas', n('Ubicados'), meta.Ubicados],
      ['Recolectadas (cortos completos)', n('Recolectados'), meta.Recolectados],
      ['Cargadas', n('Cargados'), meta.Cargados]
    ].map(function (f) {
      return '<tr><td>' + esc(f[0]) + '</td><td class="num"><b>' + f[1] + '</b></td><td class="num">' + f[2] + '</td></tr>';
    }).join('');
    var bahias = S.data.Bahias.map(function (b) {
      return { b: b.Bahia, cajas: S.data.Piso.filter(function (r) { return r.Bahia === b.Bahia; }) };
    }).filter(function (x) { return x.cajas.length; });
    var bahiasHtml = bahias.length
      ? '<div class="tabla-wrap"><table class="compacta"><thead><tr><th>Bahía</th><th class="num">Cajas</th><th>Caja (estatus)</th></tr></thead><tbody>' +
        bahias.map(function (x) {
          return '<tr><td><b>' + esc(x.b) + '</b></td><td class="num">' + x.cajas.length + '</td><td>' +
            x.cajas.map(function (r) { return esc(r.Caja) + ' <span class="muted">(' + esc(r.Estatus) + ')</span>'; }).join('<br>') + '</td></tr>';
        }).join('') + '</tbody></table></div>'
      : '<div class="muted">Sin cajas en bahías.</div>';
    var ahora = new Date();
    var rampas = S.data.Rampas.filter(function (r) { return QR.texto(r.Caja) !== ''; });
    var rampasHtml = rampas.length
      ? '<div class="tabla-wrap"><table class="compacta"><thead><tr><th>Rampa</th><th>Caja</th><th>Estatus</th><th>Llegada</th></tr></thead><tbody>' +
        rampas.map(function (r) {
          return '<tr><td><b>' + esc(r.Rampa) + '</b></td><td>' + esc(r.Caja) + '</td><td>' + esc(r.Estatus) + '</td><td>' +
            esc(fmtF(r.FechaLlegada)) + ' · ' + fmtH(QR.horasDesde(r.FechaLlegada, ahora)) + '</td></tr>';
        }).join('') + '</tbody></table></div>'
      : '<div class="muted">Sin cajas en rampa.</div>';
    var top = criticas().slice(0, 5);
    var topHtml = top.length
      ? '<ol class="lista-criticas">' + top.map(function (c) {
        return '<li><span class="chip ' + claseSem(c.estado) + '">' + esc(c.estado) + '</span><b>' + esc(c.caja) + '</b><span>' + esc(c.lugar) + '</span>' +
          '<span class="muted">' + esc(c.estatus) + '</span><b style="margin-left:auto">' + fmtH(c.horas) + '</b></li>';
      }).join('') + '</ol>'
      : '<div class="muted">Sin cajas activas.</div>';
    return '<div class="card" id="entrega"><div class="row"><div><h2>Entrega de turno</h2>' +
      '<div class="muted">Jornada del ' + esc(aFecha(dia)) + ' · Turno: ' + esc(sel === 'Todos' ? 'todos' : sel) +
      ' · Encargado: ' + esc(encs.map(nombreEnc).join(', ') || '—') + '</div>' +
      '<div class="muted">Movimientos de la jornada; el activo y las cajas con más tiempo son al momento de consultar.</div></div>' +
      '<div class="row no-imp" style="margin-left:auto;justify-content:flex-end">' +
      '<label>Día <input type="date" data-chg="kpi" data-f="entDia" value="' + esc(aFecha(dia)) + '"></label>' +
      '<label>Turno <select data-chg="kpi" data-f="entTurno"><option value="Todos"' + (sel === 'Todos' ? ' selected' : '') + '>Todos</option>' +
      opts(listaTurnos(), sel) + '</select></label>' +
      '<button class="sec" type="button" data-act="actualizar">Actualizar datos</button>' +
      '<button class="sec" type="button" data-act="imprimir">Imprimir entrega</button></div></div>' +
      '<div class="dos" style="margin-top:12px"><div><h3>Movimientos</h3><div class="tabla-wrap"><table><thead><tr><th>Concepto</th><th class="num">Real</th><th class="num">Meta</th></tr></thead><tbody>' +
      filas + '</tbody></table></div></div>' +
      '<div><h3>Nota: cajas con más tiempo</h3>' + topHtml + '</div></div>' +
      '<div class="dos" style="margin-top:12px"><div><h3>Activo en bahías (' + S.data.Piso.length + ')</h3>' + bahiasHtml + '</div>' +
      '<div><h3>Activo en rampas (' + rampas.length + ')</h3>' + rampasHtml + '</div></div></div>';
  }
  function tabKpi() {
    if (!S.f.kpiDia) S.f.kpiDia = aFecha(new Date());
    if (!S.f.kpiMes) S.f.kpiMes = aMes(new Date());
    if (!S.f.entDia) S.f.entDia = aFecha(new Date());
    if (!S.f.entTurno) S.f.entTurno = S.turno || 'Todos';
    return bloqueEntrega() + bloqueMiTurno() + (esSup() ? bloqueReporteDiario() + bloqueMensual() : '');
  }

  // ---------- configuracion (supervisor) ----------
  function tabConfig() {
    var p = P();
    var turnos = S.data.Turnos;
    var usuarios = S.data.Usuarios.map(function (u) {
      return '<tr><td>' + esc(u.NoEmpleado) + '</td><td>' + esc(u.Nombre) + '</td><td>' + esc(u.Rol) + '</td>' +
        '<td>' + (u.Activo === 'Sí' ? 'Activo' : 'Inactivo') + '</td>' +
        '<td><form data-form="cfg-password" data-id="' + esc(u.NoEmpleado) + '" class="row" style="flex-wrap:nowrap">' +
        '<input name="pass" type="password" placeholder="Nueva contraseña" style="width:150px"><button class="sec peq">Cambiar</button></form>' +
        '<button class="sec peq" style="margin-top:4px" data-act="usuarioToggle" data-id="' + esc(u.NoEmpleado) + '">' + (u.Activo === 'Sí' ? 'Desactivar' : 'Activar') + '</button></td></tr>';
    }).join('');
    var filasT = turnos.map(function (t, i) {
      var m = metaDe(t.Turno);
      return '<tr><td><b>' + esc(t.Turno) + '</b><input type="hidden" name="t_' + i + '" value="' + esc(t.Turno) + '"></td>' +
        '<td><input type="time" name="ini_' + i + '" value="' + esc(hh(t.Inicio)) + '" style="width:110px"></td>' +
        '<td><input type="time" name="fin_' + i + '" value="' + esc(hh(t.Fin)) + '" style="width:110px"></td>' +
        TIPOS.map(function (k) { return '<td><input type="number" min="0" name="' + k + '_' + i + '" value="' + m[k] + '" style="width:72px"></td>'; }).join('') +
        '<td><button class="sec peq" type="button" data-act="turnoDel" data-id="' + esc(t.Turno) + '">Quitar</button></td></tr>';
    }).join('');
    var matriz = QR.DIAS.map(function (dia, i) {
      var sel = QR.turnosDelDia(dia, S.data.DiasLaborales);
      return '<tr><td><b>' + esc(dia) + '</b></td>' + turnos.map(function (t, j) {
        return '<td><input type="checkbox" name="d_' + i + '_' + j + '"' + (sel.indexOf(t.Turno) >= 0 ? ' checked' : '') + '></td>';
      }).join('') + '</tr>';
    }).join('');
    var festivos = S.data.Festivos.slice().sort(function (a, b) { return tiempo(a.Fecha) - tiempo(b.Fecha); }).map(function (f) {
      return '<tr><td>' + esc(fmtD(f.Fecha)) + '</td><td>' + esc(f.Descripcion) + '</td>' +
        '<td><button class="sec peq" data-act="festivoDel" data-id="' + esc(QR.claveFecha(f.Fecha)) + '">Quitar</button></td></tr>';
    }).join('');
    var bahias = S.data.Bahias.map(function (b) { return b.Bahia + ',' + b.Proveedor; }).join('\n');
    var estatus = S.data.Estatus.map(function (e) { return e.Estatus + ',' + e.Responsable; }).join('\n');
    return '<div class="dos"><div>' +
      '<div class="card"><h2>Turnos, horarios y metas</h2>' +
      '<p class="muted">Horario de inicio y fin en formato 24 horas. Un turno puede cruzar medianoche (ej. Noche 22:00 a 06:00).</p>' +
      '<form data-form="cfg-turnos"><div class="tabla-wrap"><table><thead><tr><th>Turno</th><th>Inicio</th><th>Fin</th>' +
      '<th>Desc.</th><th>Recol.</th><th>Ubic.</th><th>Carg.</th><th></th></tr></thead><tbody>' +
      (filasT || '<tr><td colspan="8" class="muted">Sin turnos.</td></tr>') + '</tbody></table></div>' +
      '<div class="row" style="margin-top:10px"><button type="submit">Guardar turnos y metas</button></div></form>' +
      '<form data-form="cfg-turno-nuevo" class="row" style="margin-top:12px">' +
      '<label>Nuevo turno <input name="turno" required></label>' +
      '<label>Inicio <input name="inicio" type="time" required></label>' +
      '<label>Fin <input name="fin" type="time" required></label>' +
      '<button class="sec" type="submit">Agregar turno</button></form></div>' +
      '<div class="card"><h2>Usuarios</h2><div class="tabla-wrap"><table><thead><tr><th>No. empleado</th><th>Nombre</th><th>Rol</th><th>Estado</th><th>Acciones</th></tr></thead><tbody>' +
      (usuarios || '<tr><td colspan="5" class="muted">Sin usuarios.</td></tr>') + '</tbody></table></div>' +
      '<form data-form="cfg-usuario" class="row" style="margin-top:12px">' +
      '<label>No. empleado <input name="noEmp" required></label><label>Nombre <input name="nombre" required></label>' +
      '<label>Rol <select name="rol">' + opts(QR.CAT.roles, 'Capturista') + '</select></label>' +
      '<label>Contraseña <input name="pass" type="password" required></label><button type="submit">Agregar usuario</button></form></div>' +
      '</div><div>' +
      '<div class="card"><h2>Días laborables</h2><p class="muted">Marca en qué turnos se trabaja cada día. Los días sin turno no cuentan en el métrico.</p>' +
      '<form data-form="cfg-dias"><div class="tabla-wrap"><table><thead><tr><th>Día</th>' +
      turnos.map(function (t) { return '<th>' + esc(t.Turno) + '</th>'; }).join('') + '</tr></thead><tbody>' + matriz + '</tbody></table></div>' +
      '<div class="row" style="margin-top:10px"><button type="submit">Guardar días laborables</button></div></form></div>' +
      '<div class="card"><h2>Festivos</h2><p class="muted">Los festivos no cuentan en las metas ni en el métrico mensual.</p>' +
      '<div class="tabla-wrap"><table><thead><tr><th>Fecha</th><th>Descripción</th><th></th></tr></thead><tbody>' +
      (festivos || '<tr><td colspan="3" class="muted">Sin festivos registrados.</td></tr>') + '</tbody></table></div>' +
      '<form data-form="cfg-festivo" class="row" style="margin-top:12px"><label>Fecha <input name="fecha" type="date" required></label>' +
      '<label>Descripción <input name="descripcion"></label><button class="sec" type="submit">Agregar festivo</button></form></div>' +
      '<div class="card"><h2>Parámetros</h2><form data-form="cfg-parametros" class="row">' +
      '<label>Alerta (h) <input name="umbral_alerta_h" type="number" min="0" value="' + p.umbral_alerta_h + '"></label>' +
      '<label>Demora (h) <input name="umbral_demora_h" type="number" min="0" value="' + p.umbral_demora_h + '"></label>' +
      '<label>Máx. cajas por bahía <input name="max_cajas_bahia" type="number" min="1" value="' + p.max_cajas_bahia + '"></label>' +
      '<button type="submit">Guardar</button></form></div>' +
      '<div class="card"><h2>Bahías</h2><form data-form="cfg-bahias"><label>Una por línea: Bahía,Proveedor</label><textarea name="lista">' + esc(bahias) + '</textarea>' +
      '<div class="row" style="margin-top:8px"><button type="submit">Guardar bahías</button></div></form></div>' +
      '<div class="card"><h2>Estatus de rampa</h2><form data-form="cfg-estatus-rampa"><label>Uno por línea. "Vacía" se agrega sola y vacía la rampa al elegirlo.</label><textarea name="lista">' + esc(S.data.EstatusRampas.map(function (e) { return e.Estatus; }).join('\n')) + '</textarea>' +
      '<div class="row" style="margin-top:8px"><button type="submit">Guardar estatus de rampa</button></div></form></div>' +
      '<div class="card"><h2>Estatus de piso</h2><form data-form="cfg-estatus"><label>Una por línea: Estatus,Responsable</label><textarea name="lista">' + esc(estatus) + '</textarea>' +
      '<div class="row" style="margin-top:8px"><button type="submit">Guardar estatus</button></div></form></div>' +
      '</div></div>';
  }

  var VISTAS = { inicio: tabInicio, ops: tabOperaciones, piso: tabPiso, rampas: tabRampas, facturas: tabFacturas,
    cortos: tabCortos, notas: tabNotas, kpi: tabKpi, config: tabConfig };

  // ---------- render ----------
  // Con un error se conservan los valores que la persona ya escribio (no las contrasenas)
  function claveCampo(f, el) { return (f.dataset.form || '') + '#' + (f.dataset.id || f.dataset.rampa || '') + '|' + el.name; }
  function capturarFormularios() {
    var v = {};
    RAIZ.querySelectorAll('form').forEach(function (f) {
      f.querySelectorAll('input[name],select[name],textarea[name]').forEach(function (el) {
        if (el.type === 'password') return;
        v[claveCampo(f, el)] = el.type === 'checkbox' ? el.checked : el.value;
      });
    });
    return v;
  }
  function restaurarFormularios(v) {
    RAIZ.querySelectorAll('form').forEach(function (f) {
      f.querySelectorAll('input[name],select[name],textarea[name]').forEach(function (el) {
        var k = claveCampo(f, el);
        if (el.type === 'password' || !Object.prototype.hasOwnProperty.call(v, k)) return;
        if (el.type === 'checkbox') el.checked = !!v[k]; else el.value = v[k];
      });
    });
  }
  function render() {
    var html;
    var conservar = S.msg && S.msg.tipo === 'err' && RAIZ.querySelector('form') ? capturarFormularios() : null;
    switch (S.pantalla) {
      case 'carpeta': html = pantallaCarpeta(); break;
      case 'permiso': html = pantallaPermiso(); break;
      case 'crear': html = pantallaCrear(); break;
      case 'login': html = pantallaLogin(); break;
      case 'rol': html = pantallaRol(); break;
      case 'bloqueo': html = pantallaBloqueo(); break;
      case 'app': html = pantallaApp(); break;
      default: html = pantallaCarga();
    }
    RAIZ.innerHTML = html;
    if (conservar) restaurarFormularios(conservar);
  }

  // Turno automatico: se revisa cada 30 s. No se redibuja si la persona esta escribiendo en un campo.
  function tic() {
    if (!S.data || S.pantalla !== 'app') return;
    var el = document.activeElement;
    if (el && RAIZ.contains(el) && /^(INPUT|SELECT|TEXTAREA)$/.test(el.tagName)) return;
    var antes = S.turno;
    if (S.turnoAuto) S.turno = resolverTurno();
    if (S.turno !== antes || S.tab === 'inicio' || S.tab === 'piso' || S.tab === 'rampas' || S.tab === 'kpi') render();
  }

  // ---------- eventos ----------
  var ACC = {
    elegirCarpeta: function () { intentar(elegirCarpetaDatos); },
    permiso: function () {
      intentar(async function () {
        var p = await S.dir.requestPermission({ mode: 'readwrite' });
        if (p === 'granted') await abrirDatos();
        else { S.pantalla = 'carpeta'; S.msg = { tipo: 'err', texto: 'Sin permiso para la carpeta.' }; }
      });
    },
    cambiarCarpeta: function () {
      liberarBloqueo().then(function () { S.fh = null; S.pantalla = 'carpeta'; S.msg = null; S.usuario = null; render(); });
    },
    salir: function () {
      liberarBloqueo().then(function () {
        S.usuario = null; S.rol = null; S.encargado = ''; S.motivoBloqueo = '';
        S.pantalla = S.data && S.data.Usuarios.length ? 'login' : 'carga'; S.msg = null; render();
      });
    },
    rol: function (el) { entrarDesdeLogin(el.dataset.rol); },
    tomarBloqueo: function () {
      var k = QR.bloqueoDeModo(S.rol);
      if (!k) return;
      if (!confirm('¿Tomar el bloqueo? La otra computadora pasará a solo lectura.')) return;
      intentar(async function () {
        var r = await tomarBloqueo(k, true);
        if (!r.ok) throw 'No pude tomar el bloqueo. Intenta de nuevo.';
        S.lock = k; S.miBloqueo = r.mio; S.motivoBloqueo = '';
        if (k === 'rampas' && migrar()) await guardarSiCambio();
        S.msg = { tipo: 'ok', texto: 'Tienes el bloqueo: ya puedes editar.' };
      });
    },
    actualizar: function () { intentar(refrescarLectura); },
    tab: function (el) { S.tab = el.dataset.tab; S.msg = null; render(); },
    cinta: function (el) {
      S.cinta = el.dataset.cinta;
      var g = cintaVisible().filter(function (x) { return x.id === S.cinta; })[0];
      var vis = g ? itemsVisibles(g) : [];
      if (vis.length && !vis.some(function (i) { return i.t === S.tab; })) S.tab = vis[0].t;
      S.msg = null; render();
    },
    tema: function () {
      S.tema = S.tema === 'negro' ? 'blanco' : 'negro';
      document.documentElement.setAttribute('data-tema', S.tema);
      try { localStorage.setItem('qr-tema', S.tema); } catch (e) { /* sin persistencia: el tema dura esta sesion */ }
      render();
    },
    ubicar: function (el) { if (confirm('¿Ubicar esta caja? Sale del piso y queda en el historial.')) ejecutar(function () { sacarDePiso(el.dataset.id, 'Ubicado'); }, 'Caja ubicada.'); },
    cargar: function (el) { if (confirm('¿Marcar esta caja como cargada? Sale del piso y queda en el historial.')) ejecutar(function () { sacarDePiso(el.dataset.id, 'Cargada'); }, 'Caja cargada.'); },
    corregir: function (el) {
      var r = S.data.Piso.filter(function (x) { return String(x.ID) === String(el.dataset.id); })[0];
      var txt = window.prompt('Nueva hora de inicio del responsable (aaaa-mm-dd hh:mm):', r ? aInput(r.InicioResponsable instanceof Date ? r.InicioResponsable : new Date()).replace('T', ' ') : '');
      if (txt === null) return;
      var nueva = deInputDT(txt.trim().replace(' ', 'T'));
      if (!nueva) { S.msg = { tipo: 'err', texto: 'Formato de fecha no válido.' }; return render(); }
      ejecutar(function () { corregirInicioPiso(el.dataset.id, nueva); }, 'Hora corregida.');
    },
    expandir: function (el) {
      S.abierto = S.abierto || {};
      S.abierto[el.dataset.k] = !S.abierto[el.dataset.k];
      render();
    },
    facturaEliminar: function (el) {
      if (confirm('¿Eliminar esta factura? Quedará marcada como Deleted y se puede restaurar.')) ejecutar(function () { marcarFactura(+el.dataset.i, true); }, 'Factura eliminada.');
    },
    facturaRestaurar: function (el) { ejecutar(function () { marcarFactura(+el.dataset.i, false); }, 'Factura restaurada.'); },
    rampaVaciar: function (el) { if (confirm('¿Vaciar la rampa ' + el.dataset.rampa + '?')) ejecutar(function () { vaciarRampa(el.dataset.rampa); }, 'Rampa vaciada.'); },
    cortosPrimera: function () {
      if (!puedeEscribir('cortos')) { S.msg = { tipo: 'err', texto: motivoLectura() }; return render(); }
      S.cortosModo = 'primera'; S.preguntaActualizar = false; var i = document.getElementById('archivoCortos'); if (i) i.click();
    },
    cortosPreguntar: function () { S.preguntaActualizar = !S.preguntaActualizar; render(); },
    cortosNo: function () { S.preguntaActualizar = false; render(); },
    cortosSeleccionar: function () {
      if (!puedeEscribir('cortos')) { S.msg = { tipo: 'err', texto: motivoLectura() }; return render(); }
      S.cortosModo = 'actualizar'; S.preguntaActualizar = false; var i = document.getElementById('archivoCortos'); if (i) i.click();
    },
    usuarioToggle: function (el) {
      ejecutar(function () {
        var u = S.data.Usuarios.filter(function (x) { return x.NoEmpleado === el.dataset.id; })[0];
        if (!u) throw 'Usuario no encontrado.';
        var activo = u.Activo === 'Sí';
        if (activo && u.Rol === 'Supervisor' && S.data.Usuarios.filter(function (x) { return x.Rol === 'Supervisor' && x.Activo === 'Sí'; }).length <= 1) throw 'Debe quedar al menos un supervisor activo.';
        u.Activo = activo ? 'No' : 'Sí';
      }, 'Usuario actualizado.');
    },
    turnoDel: function (el) {
      ejecutar(function () {
        if (S.data.Turnos.length <= 1) throw 'Debe quedar al menos un turno.';
        var nombre = el.dataset.id;
        S.data.Turnos = S.data.Turnos.filter(function (t) { return t.Turno !== nombre; });
        S.data.Metas = S.data.Metas.filter(function (m) { return m.Turno !== nombre; });
        S.data.DiasLaborales.forEach(function (d) {
          d.Turnos = QR.turnosDelDia(d.Dia, S.data.DiasLaborales).filter(function (t) { return t !== nombre; }).join(',');
        });
        refrescarTurnoTrasCambio();
      }, 'Turno quitado.');
    },
    imprimir: function () { window.print(); },
    festivoDel: function (el) {
      ejecutar(function () {
        S.data.Festivos = S.data.Festivos.filter(function (f) { return QR.claveFecha(f.Fecha) !== el.dataset.id; });
      }, 'Festivo quitado.');
    }
  };

  function refrescarTurnoTrasCambio() {
    if (S.turnoAuto) S.turno = resolverTurno();
    else if (listaTurnos().indexOf(S.turno) < 0) { S.turnoAuto = true; S.turno = resolverTurno(); }
  }

  var FORMAS = {
    'primer-supervisor': function (f, d) {
      var noEmp = QR.texto(d.get('noEmp')), pass = String(d.get('pass') || '');
      if (!noEmp) return ejecutar(function () { throw 'Captura el número de empleado.'; });
      if (pass.length < 4) return ejecutar(function () { throw 'La contraseña debe tener al menos 4 caracteres.'; });
      if (pass !== String(d.get('pass2'))) return ejecutar(function () { throw 'Las contraseñas no coinciden.'; });
      hashPass(noEmp, pass).then(function (h) {
        return ejecutar(function () {
          S.data.Usuarios.push({ NoEmpleado: noEmp, Nombre: QR.texto(d.get('nombre')), Rol: 'Administrador', PassHash: h, Activo: 'Sí' });
        }, 'Administrador creado. Ya puedes entrar.').then(function () { S.setup = false; S.pantalla = 'login'; render(); });
      });
    },
    desbloquear: function (f, d) {
      hashPass(S.usuario.NoEmpleado, String(d.get('pass') || '')).then(function (h) {
        if (h !== S.usuario.PassHash) { S.msg = { tipo: 'err', texto: 'Contraseña incorrecta.' }; return render(); }
        S.pantalla = 'app'; S.msg = null;
        return entrarEnModo(S.rol).then(function () { render(); });
      });
    },
    login: function (f, d) {
      var noEmp = QR.texto(d.get('noEmp')), pass = String(d.get('pass') || '');
      hashPass(noEmp, pass).then(function (h) {
        var u = S.data.Usuarios.filter(function (x) { return x.NoEmpleado === noEmp && x.Activo === 'Sí'; })[0];
        if (!u || u.PassHash !== h) { S.msg = { tipo: 'err', texto: 'Número de empleado o contraseña incorrectos.' }; return render(); }
        S.usuario = u; S.pantalla = 'rol'; S.msg = null; render();
      });
    },
    'piso-alta': function (f, d) {
      ejecutar(function () {
        registrarEnPiso({ bahia: d.get('bahia'), caja: d.get('caja'), pallets: d.get('pallets'), unidad: d.get('unidad'),
          estatus: d.get('estatus'), fecha: deInputDT(d.get('fecha')) });
      }, 'Caja registrada.');
    },
    'piso-estatus': function (f, d) {
      ejecutar(function () { cambiarEstatusPiso(f.dataset.id, d.get('estatus')); }, 'Estatus actualizado.');
    },
    'rampa-asignar': function (f, d) {
      ejecutar(function () {
        asignarRampa(f.dataset.rampa, { caja: d.get('caja'), proveedor: d.get('proveedor'), estatus: d.get('estatus'),
          fecha: deInputDT(d.get('fecha')) });
      }, 'Caja asignada a la rampa.');
    },
    'rampa-estatus': function (f, d) {
      ejecutar(function () { cambiarEstatusRampa(f.dataset.rampa, d.get('estatus')); }, 'Estatus de rampa actualizado.');
    },
    'cfg-estatus-rampa': function (f, d) {
      ejecutar(function () {
        var lista = String(d.get('lista') || '').split('\n').map(function (l) { return l.trim(); }).filter(Boolean);
        lista = lista.filter(function (e, i) { return lista.indexOf(e) === i; });
        if (lista.indexOf('Vacía') < 0) lista.push('Vacía');
        if (lista.length < 2) throw 'Deja al menos un estatus además de Vacía.';
        S.data.Rampas.forEach(function (r) {
          if (QR.texto(r.Caja) && lista.indexOf(r.Estatus) < 0) throw 'El estatus "' + r.Estatus + '" está en uso en la rampa ' + r.Rampa + '.';
        });
        S.data.EstatusRampas = lista.map(function (e) { return { Estatus: e }; });
      }, 'Estatus de rampa guardados.');
    },
    'rampa-pasar': function (f, d) {
      ejecutar(function () {
        pasarRampaABahia(d.get('rampa'), { bahia: d.get('bahia'), estatus: d.get('estatus'), pallets: d.get('pallets'),
          unidad: d.get('unidad'), fecha: deInputDT(d.get('fecha')) });
      }, 'Caja pasada de rampa a bahía.');
    },
    'factura-alta': function (f, d) {
      ejecutar(function () {
        agregarFactura({ CajaPlaca: d.get('CajaPlaca'), Proveedor: d.get('Proveedor'),
          FechaRecepcion: d.get('FechaRecepcion'), Estatus: d.get('Estatus'), Comentario: d.get('Comentario') });
      }, 'Factura registrada.');
    },
    'nota-alta': function (f, d) {
      ejecutar(function () {
        agregarNota({ Carrier: d.get('Carrier'), Caja: d.get('Caja'), Ubicacion: d.get('Ubicacion'), Rampa: d.get('Rampa'),
          Afectacion: d.get('Afectacion'), Comentario: d.get('Comentario') });
      }, 'Nota agregada.');
    },
    'cfg-turnos': function (f, d) {
      ejecutar(function () {
        var nT = [], nM = [];
        S.data.Turnos.forEach(function (t, i) {
          var ini = QR.texto(d.get('ini_' + i)), fin = QR.texto(d.get('fin_' + i));
          if (QR.minutosHora(ini) == null || QR.minutosHora(fin) == null) throw 'Horario no válido en ' + t.Turno + ' (usa HH:mm).';
          if (QR.minutosHora(ini) === QR.minutosHora(fin)) throw 'Inicio y fin no pueden ser iguales en ' + t.Turno + '.';
          nT.push({ Turno: t.Turno, Inicio: ini, Fin: fin });
          var m = { Turno: t.Turno };
          TIPOS.forEach(function (k) {
            var v = Number(d.get(k + '_' + i));
            if (isNaN(v) || v < 0) throw 'Meta no válida en ' + t.Turno + '.';
            m[k] = v;
          });
          nM.push(m);
        });
        S.data.Turnos = nT; S.data.Metas = nM;
        refrescarTurnoTrasCambio();
      }, 'Turnos y metas guardados.');
    },
    'cfg-turno-nuevo': function (f, d) {
      ejecutar(function () {
        var t = QR.texto(d.get('turno'));
        if (!t) throw 'Captura el nombre del turno.';
        if (listaTurnos().indexOf(t) >= 0) throw 'Ese turno ya existe.';
        var ini = QR.texto(d.get('inicio')), fin = QR.texto(d.get('fin'));
        if (QR.minutosHora(ini) == null || QR.minutosHora(fin) == null) throw 'Captura el horario de inicio y fin.';
        if (QR.minutosHora(ini) === QR.minutosHora(fin)) throw 'Inicio y fin no pueden ser iguales.';
        S.data.Turnos.push({ Turno: t, Inicio: ini, Fin: fin });
        S.data.Metas.push({ Turno: t, Descargados: 9, Recolectados: 9, Ubicados: 12, Cargados: 5 });
        refrescarTurnoTrasCambio();
      }, 'Turno agregado. Marca sus días laborables abajo.');
    },
    'cfg-dias': function (f, d) {
      ejecutar(function () {
        var nombres = listaTurnos();
        S.data.DiasLaborales = QR.DIAS.map(function (dia, i) {
          var marcados = nombres.filter(function (t, j) { return d.get('d_' + i + '_' + j) === 'on'; });
          return { Dia: dia, Turnos: marcados.join(',') };
        });
        refrescarTurnoTrasCambio();
      }, 'Días laborables guardados.');
    },
    'cfg-festivo': function (f, d) {
      ejecutar(function () {
        var fecha = deFecha(d.get('fecha'));
        if (!fecha) throw 'Captura la fecha del festivo.';
        var clave = QR.claveFecha(fecha);
        if (S.data.Festivos.some(function (x) { return QR.claveFecha(x.Fecha) === clave; })) throw 'Ese festivo ya está registrado.';
        S.data.Festivos.push({ Fecha: fecha, Descripcion: QR.texto(d.get('descripcion')) });
      }, 'Festivo agregado.');
    },
    'cfg-usuario': function (f, d) {
      var noEmp = QR.texto(d.get('noEmp')), pass = String(d.get('pass') || '');
      if (!noEmp) return ejecutar(function () { throw 'Falta el número de empleado.'; });
      if (pass.length < 4) return ejecutar(function () { throw 'La contraseña debe tener al menos 4 caracteres.'; });
      if (S.data.Usuarios.some(function (u) { return u.NoEmpleado === noEmp; })) return ejecutar(function () { throw 'Ese número ya existe.'; });
      hashPass(noEmp, pass).then(function (h) {
        ejecutar(function () {
          S.data.Usuarios.push({ NoEmpleado: noEmp, Nombre: QR.texto(d.get('nombre')), Rol: d.get('rol'), PassHash: h, Activo: 'Sí' });
        }, 'Usuario agregado.');
      });
    },
    'cfg-password': function (f, d) {
      var pass = String(d.get('pass') || '');
      if (pass.length < 4) return ejecutar(function () { throw 'La contraseña debe tener al menos 4 caracteres.'; });
      hashPass(f.dataset.id, pass).then(function (h) {
        ejecutar(function () {
          var u = S.data.Usuarios.filter(function (x) { return x.NoEmpleado === f.dataset.id; })[0];
          if (!u) throw 'Usuario no encontrado.';
          u.PassHash = h;
        }, 'Contraseña cambiada.');
      });
    },
    'cfg-parametros': function (f, d) {
      ejecutar(function () {
        ['umbral_alerta_h', 'umbral_demora_h', 'max_cajas_bahia'].forEach(function (k) {
          var v = Number(d.get(k));
          if (isNaN(v) || v < 0) throw 'Valor no válido en ' + k + '.';
          var fila = S.data.Parametros.filter(function (x) { return x.Clave === k; })[0];
          if (fila) fila.Valor = v; else S.data.Parametros.push({ Clave: k, Valor: v });
        });
      }, 'Parámetros guardados.');
    },
    'cfg-bahias': function (f, d) {
      ejecutar(function () {
        var filas = String(d.get('lista') || '').split('\n').map(function (l) { return l.trim(); }).filter(Boolean).map(function (l) {
          var p = l.split(','); return { Bahia: p[0].trim(), Proveedor: (p[1] || 'Insourcing').trim() };
        });
        if (!filas.length) throw 'Deja al menos una bahía.';
        S.data.Piso.forEach(function (r) {
          if (!filas.some(function (b) { return b.Bahia === r.Bahia; })) throw 'La bahía ' + r.Bahia + ' tiene cajas en piso; no se puede quitar.';
        });
        S.data.Bahias = filas;
      }, 'Bahías guardadas.');
    },
    'cfg-estatus': function (f, d) {
      ejecutar(function () {
        var filas = String(d.get('lista') || '').split('\n').map(function (l) { return l.trim(); }).filter(Boolean).map(function (l) {
          var p = l.split(','); return { Estatus: p[0].trim(), Responsable: (p[1] || 'Recibo').trim() };
        });
        if (!filas.length) throw 'Deja al menos un estatus.';
        S.data.Piso.forEach(function (r) {
          if (!filas.some(function (e) { return e.Estatus === r.Estatus; })) throw 'El estatus "' + r.Estatus + '" está en uso en piso.';
        });
        S.data.Estatus = filas;
      }, 'Estatus guardados.');
    }
  };

  document.addEventListener('click', function (ev) {
    var el = ev.target.closest('[data-act]');
    if (!el || el.tagName === 'FORM') return;
    var fn = ACC[el.dataset.act];
    if (fn) { ev.preventDefault(); fn(el, ev); }
  });
  document.addEventListener('submit', function (ev) {
    var f = ev.target;
    if (!f.dataset || !f.dataset.form) return;
    ev.preventDefault();
    var fn = FORMAS[f.dataset.form];
    if (fn) fn(f, new FormData(f));
  });
  document.addEventListener('change', function (ev) {
    var el = ev.target;
    var c = el.dataset && el.dataset.chg;
    if (!c) return;
    if (c === 'turno') {
      if (el.value === 'auto') { S.turnoAuto = true; S.turno = resolverTurno(); }
      else { S.turnoAuto = false; S.turno = el.value; }
      render();
    }
    else if (c === 'filtro') { S.f[el.dataset.f] = el.value; render(); }
    else if (c === 'ops') { S.f.ops[el.dataset.f] = el.value; render(); }
    else if (c === 'kpi') { S.f[el.dataset.f] = el.value; render(); }
    else if (c === 'encargado') { S.encargado = el.value; render(); }
    else if (c === 'snap') { S.snap = el.value === snapActivo() ? null : el.value; render(); }
    else if (c === 'cortos') {
      ejecutar(function () { editarCorto(el.dataset.key, el.dataset.campo, el.value); }, null);
    }
    else if (c === 'piso-estatus') {
      ejecutar(function () { cambiarEstatusPiso(el.dataset.id, el.value); }, 'Estatus actualizado.');
    }
    else if (c === 'factura') {
      ejecutar(function () { cambiarEstatusFactura(+el.dataset.i, el.value); }, null);
    }
    else if (c === 'archivoCortos') {
      var file = el.files && el.files[0];
      el.value = '';
      if (!file) return;
      ejecutar(function () { return importarArchivoCortos(file); }, null);
    }
  });

  // ---------- arranque ----------
  try { S.tema = document.documentElement.getAttribute('data-tema') === 'blanco' ? 'blanco' : 'negro'; } catch (e) { S.tema = 'negro'; }
  setInterval(tic, 30000);
  setInterval(revisarInactividad, 30000);
  setInterval(latir, 60000);
  setInterval(refrescarLectura, 60000);
  render();
  iniciar();
})();
