/* QORVEX Recibo - nucleo de datos y reglas (sin DOM, probable en Node) */
(function (root, factory) {
  var api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.QR = api;
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // ---- Esquema del libro QORVEX_Recibo.xlsx (una hoja por tabla) ----
  var SHEETS = {
    Usuarios: ['NoEmpleado', 'Nombre', 'Rol', 'PassHash', 'Activo'],
    Turnos: ['Turno', 'Inicio', 'Fin'],
    Metas: ['Turno', 'Descargados', 'Recolectados', 'Ubicados', 'Cargados'],
    DiasLaborales: ['Dia', 'Turnos'],
    Festivos: ['Fecha', 'Descripcion'],
    Parametros: ['Clave', 'Valor'],
    Bahias: ['Bahia', 'Proveedor'],
    Estatus: ['Estatus', 'Responsable'],
    Rampas: ['Rampa', 'Caja', 'Proveedor', 'Estatus', 'FechaLlegada', 'NoEmpleado'],
    EstatusRampas: ['Estatus'],
    Piso: ['ID', 'Bahia', 'Caja', 'Proveedor', 'Pallets', 'Unidad', 'Estatus',
           'FechaDescarga', 'Responsable', 'InicioResponsable', 'NoEmpleado', 'Turno'],
    Historial: ['FechaEvento', 'ID', 'Bahia', 'Caja', 'Evento', 'EstatusAnt', 'EstatusNvo',
                'Responsable', 'NoEmpleado', 'NombreUsuario', 'Turno', 'Detalle', 'Encargado', 'Unidad'],
    Facturas: ['NumFactura', 'CajaPlaca', 'Proveedor', 'FechaRecepcion', 'Estatus',
               'Comentario', 'Eliminada', 'NoEmpleado', 'FechaRegistro'],
    CortosSnapshots: ['Snapshot', 'FechaCarga', 'Archivo', 'Filas', 'NoEmpleado'],
    Cortos: ['Snapshot', 'Activo', 'Key', 'Linea', 'Area', 'WC', 'Item', 'Descr', 'Proveedor',
             'LSJob', 'Unid', 'Fecha', 'Off', 'Wip', 'Prioridad', 'Hora', 'EnviarHOT',
             'CajaPlacas', 'BahiaRampa', 'Estatus', 'QtyEntrega', 'Folio', 'ComentariosRecibo',
             'TipoCaja', 'NoEmpleado', 'FechaActualiz'],
    CortosLog: ['Fecha', 'Snapshot', 'Linea', 'Item', 'CajaPlacas', 'Campo', 'ValorAnt',
                'ValorNvo', 'NoEmpleado', 'Turno', 'Encargado'],
    Notas: ['Fecha', 'Carrier', 'Caja', 'Ubicacion', 'Rampa', 'Afectacion', 'Comentario', 'NoEmpleado']
  };
  // Versiones archivadas de cortos (van al registro); misma forma que las hojas de cortos
  SHEETS.CortosVersiones = SHEETS.CortosSnapshots.slice();
  SHEETS.CortosHistorico = SHEETS.Cortos.slice();

  // Campos que siempre se guardan como texto (llaves y codigos)
  var TEXT_FIELDS = ['Caja', 'CajaPlaca', 'CajaPlacas', 'Item', 'Bahia', 'BahiaRampa', 'Rampa',
    'NoEmpleado', 'Estatus', 'Folio', 'Turno', 'Key', 'Linea', 'Prioridad', 'Clave', 'Valor',
    'PassHash', 'Activo', 'Responsable', 'Proveedor', 'Descr', 'Comentario', 'Comentarios',
    'ComentariosRecibo', 'Afectacion', 'Ubicacion', 'Carrier', 'Rol', 'Nombre', 'Evento',
    'EstatusAnt', 'EstatusNvo', 'Detalle', 'Unidad', 'Unid', 'Hora', 'EnviarHOT', 'Off', 'Wip',
    'LSJob', 'WC', 'Area', 'TipoCaja', 'QtyEntrega', 'Snapshot', 'Archivo', 'Campo', 'ValorAnt',
    'ValorNvo', 'NumFactura', 'NombreUsuario', 'Pallets', 'Dia', 'Turnos', 'Descripcion', 'Eliminada', 'Encargado'];

  var DEFAULTS = {
    Turnos: [['Día', '06:00', '14:00'], ['Tarde', '14:00', '22:00'], ['Noche', '22:00', '06:00']],
    Metas: [['Día', 9, 9, 12, 5], ['Tarde', 9, 9, 12, 5], ['Noche', 9, 9, 12, 5]],
    DiasLaborales: [['Lunes', 'Día,Tarde,Noche'], ['Martes', 'Día,Tarde,Noche'], ['Miércoles', 'Día,Tarde,Noche'],
      ['Jueves', 'Día,Tarde,Noche'], ['Viernes', 'Día,Tarde,Noche'], ['Sábado', 'Día,Tarde'], ['Domingo', '']],
    Festivos: [],
    Parametros: [['umbral_alerta_h', 8], ['umbral_demora_h', 24], ['max_cajas_bahia', 100]],
    Bahias: [['B743', 'Insourcing'], ['B744', 'Insourcing'], ['B745', 'Insourcing'],
      ['B746', 'Insourcing'], ['B746A', 'Insourcing'], ['B746B', 'Insourcing'], ['B747', 'Insourcing'],
      ['B748', 'Insourcing'], ['B749', 'Insourcing'], ['B750', 'Insourcing'], ['B751', 'Insourcing'],
      ['B752', 'Insourcing'], ['B753', 'Insourcing'], ['B754', 'Insourcing'], ['B755', 'Insourcing']],
    Estatus: [['Por Validar', 'Recibo'], ['Por Etiquetar', 'Recibo'],
      ['Problemas de ASN', 'Recibo'], ['Por Ubicar', 'Putaway']],
    // "Vacía" es el estatus de una rampa sin caja; elegirlo en una rampa con caja la vacía.
    EstatusRampas: [['Por Descargar'], ['Retornable'], ['Por Retirarse'], ['Vacía']],
    Rampas: ['R743', 'R744', 'R745', 'R746', 'R746A', 'R746B', 'R747', 'R748', 'R749', 'R750',
      'R751', 'R752', 'R753', 'R754', 'R755'].map(function (r) { return [r, '', 'Insourcing', 'Vacía', '', '']; })
  };

  var DIAS = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo'];

  var CAT = {
    estatusCortos: ['Completo', 'Parcial', 'Ubicado', 'Recolectar', 'Por Revisar',
      'Sin Información', 'Solicitar caja', 'Caja Solicitada'],
    prioridades: ['0', '1', '2', '3'],
    unidades: ['Caja', 'Camioneta'],
    roles: ['Administrador', 'Supervisor', 'Capturista', 'Consulta'],
    camposTrabajoCortos: ['BahiaRampa', 'Estatus', 'QtyEntrega', 'Folio', 'ComentariosRecibo', 'TipoCaja']
  };

  // Encabezado del archivo de cortos del cliente -> campo interno
  var CORTOS_MAP = {
    area: 'Area', wc: 'WC', item: 'Item', descr: 'Descr', proveedor: 'Proveedor',
    lsjob: 'LSJob', unid: 'Unid', fecha: 'Fecha', off: 'Off', wip: 'Wip',
    prioridad: 'Prioridad', hora: 'Hora', enviarahot: 'EnviarHOT', cajaplacas: 'CajaPlacas',
    bahiaorampa: 'BahiaRampa', rampaobahia: 'BahiaRampa', estatus: 'Estatus',
    qtyentrega: 'QtyEntrega', qty: 'QtyEntrega', folio: 'Folio',
    comentariosrecibo: 'ComentariosRecibo', comentarios: 'ComentariosRecibo', tipodecaja: 'TipoCaja'
  };

  // ---- Utilidades ----
  function norm(h) {
    return String(h == null ? '' : h).normalize('NFD').replace(/[̀-ͯ]/g, '')
      .toLowerCase().replace(/[^a-z0-9]/g, '');
  }

  function texto(v) {
    if (v == null) return '';
    if (v instanceof Date) return v.toISOString();
    return String(v).trim();
  }

  function p2(n) { return (n < 10 ? '0' : '') + n; }

  function claveFecha(v) {
    if (v instanceof Date && !isNaN(v)) return v.getFullYear() + '-' + p2(v.getMonth() + 1) + '-' + p2(v.getDate());
    if (typeof v === 'string') return v.slice(0, 10);
    return '';
  }

  function nombreDia(d) { return DIAS[(d.getDay() + 6) % 7]; }

  function horasDesde(fecha, ahora) {
    if (!(fecha instanceof Date) || isNaN(fecha.getTime())) return null;
    return (ahora.getTime() - fecha.getTime()) / 3600000;
  }

  // Semaforo: En tiempo / Alerta / Demora (umbrales desde Parametros)
  function estadoTiempo(horas, alerta, demora) {
    if (horas == null || isNaN(horas)) return 'Sin caja';
    if (horas < alerta) return 'En tiempo';
    if (horas < demora) return 'Alerta';
    return 'Demora';
  }

  function nextId(rows, campo) {
    var max = 0;
    rows.forEach(function (r) { var n = Number(r[campo]); if (n > max) max = n; });
    return max + 1;
  }

  function parametros(paramRows) {
    var p = { umbral_alerta_h: 8, umbral_demora_h: 24, max_cajas_bahia: 100 };
    (paramRows || []).forEach(function (r) {
      var n = Number(r.Valor);
      if (r.Clave && !isNaN(n) && r.Valor !== '') p[r.Clave] = n;
    });
    return p;
  }

  function responsableDe(estatus, catalogoEstatus) {
    var f = (catalogoEstatus || []).filter(function (e) { return e.Estatus === estatus; })[0];
    return f ? f.Responsable : 'Recibo';
  }

  // ---- Turnos, dias laborables y festivos ----
  // Hora "HH:mm" (tambien acepta fraccion de dia de Excel o Date) -> minutos del dia
  function minutosHora(v) {
    if (typeof v === 'number' && v >= 0 && v < 1) return Math.round(v * 1440);
    if (v instanceof Date && !isNaN(v)) return v.getHours() * 60 + v.getMinutes();
    var m = /^\s*(\d{1,2}):(\d{2})/.exec(String(v == null ? '' : v));
    return m ? (+m[1]) * 60 + (+m[2]) : null;
  }

  // Turno que corresponde a un momento (soporta turnos que cruzan medianoche)
  function turnoEn(momento, turnos) {
    var m = momento.getHours() * 60 + momento.getMinutes();
    for (var i = 0; i < (turnos || []).length; i++) {
      var t = turnos[i];
      var s = minutosHora(t.Inicio), e = minutosHora(t.Fin);
      if (s == null || e == null || s === e) continue;
      if (s < e ? (m >= s && m < e) : (m >= s || m < e)) return t.Turno;
    }
    return '';
  }

  function turnosDelDia(nombre, diasLab) {
    var f = (diasLab || []).filter(function (d) { return d.Dia === nombre; })[0];
    return f && f.Turnos ? String(f.Turnos).split(',').map(function (s) { return s.trim(); }).filter(Boolean) : [];
  }

  // Dia de jornada de un evento: un turno que cruza medianoche (p. ej. Noche 22:00-06:00)
  // cuenta sus horas de madrugada en el dia en que inicio el turno.
  function diaJornada(fecha, turnos, nombreTurno) {
    var d = new Date(fecha.getFullYear(), fecha.getMonth(), fecha.getDate());
    var t = (turnos || []).filter(function (x) { return x.Turno === nombreTurno; })[0];
    if (t) {
      var s = minutosHora(t.Inicio), e = minutosHora(t.Fin);
      var m = fecha.getHours() * 60 + fecha.getMinutes();
      if (s != null && e != null && s > e && m < e) d.setDate(d.getDate() - 1);
    }
    return { turno: nombreTurno || '', dia: d };
  }

  function esFestivo(fecha, festivos) {
    var k = claveFecha(fecha);
    return (festivos || []).some(function (f) { return claveFecha(f.Fecha) === k; });
  }

  function diaInicio(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }

  // Cuantos dias cuenta cada turno en el rango (festivos y dias no laborables no cuentan)
  function diasPorTurno(desde, hasta, diasLab, festivos) {
    var out = {};
    for (var d = diaInicio(desde); d <= hasta; d = new Date(d.getFullYear(), d.getMonth(), d.getDate() + 1)) {
      if (esFestivo(d, festivos)) continue;
      turnosDelDia(nombreDia(d), diasLab).forEach(function (t) { out[t] = (out[t] || 0) + 1; });
    }
    return out;
  }

  // Dias del rango que no se cuentan, separados por motivo
  function diasNoContados(desde, hasta, diasLab, festivos) {
    var r = { festivos: 0, noLaborables: 0 };
    for (var d = diaInicio(desde); d <= hasta; d = new Date(d.getFullYear(), d.getMonth(), d.getDate() + 1)) {
      if (esFestivo(d, festivos)) r.festivos++;
      else if (!turnosDelDia(nombreDia(d), diasLab).length) r.noLaborables++;
    }
    return r;
  }

  function nombreMes(anio, mes) {
    return ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre',
      'Octubre', 'Noviembre', 'Diciembre'][mes] + ' ' + anio;
  }

  // ---- Cortos ----
  function llaveCortos(linea, item, caja, ocurrencia) {
    return [linea, item, caja, ocurrencia].join('|');
  }

  // Lee las hojas del archivo del cliente (una hoja por linea: 011, 016...)
  // hojas: [{ name, rows: arreglos (fila 0 = encabezados) }]
  // previas: filas activas anteriores (con campos de trabajo)
  function importarCortos(hojas, previas, snapshot, noEmpleado, ahora) {
    var previasPorKey = {};
    (previas || []).forEach(function (r) { previasPorKey[r.Key] = r; });
    var salida = [], arrastrados = 0, conteo = {};

    hojas.forEach(function (hoja) {
      var enc = hoja.rows[0] || [];
      var mapa = {};
      enc.forEach(function (h, i) { var c = CORTOS_MAP[norm(h)]; if (c && mapa[c] == null) mapa[c] = i; });
      if (mapa.Item == null) return; // hoja sin columna Item: se omite
      var linea = String(hoja.name).trim();

      hoja.rows.slice(1).forEach(function (fila) {
        var item = texto(fila[mapa.Item]);
        if (!item) return;
        var caja = texto(mapa.CajaPlacas != null ? fila[mapa.CajaPlacas] : '');
        var base = linea + '|' + item + '|' + caja;
        conteo[base] = (conteo[base] || 0) + 1;
        var key = llaveCortos(linea, item, caja, conteo[base]);
        var prev = previasPorKey[key];

        var reg = { Snapshot: snapshot, Activo: 'Sí', Key: key, Linea: linea, NoEmpleado: noEmpleado,
          FechaActualiz: ahora };
        Object.keys(mapa).forEach(function (campo) {
          if (campo === 'Item' || campo === 'CajaPlacas') return;
          reg[campo] = fila[mapa[campo]] == null ? '' : fila[mapa[campo]];
        });
        reg.Item = item;
        reg.CajaPlacas = caja;

        // Campos de trabajo: si el archivo nuevo no trae valor, se conserva el anterior
        if (prev) {
          var uso = false;
          CAT.camposTrabajoCortos.forEach(function (c) {
            if (texto(reg[c]) === '' && texto(prev[c]) !== '') { reg[c] = prev[c]; uso = true; }
          });
          if (uso) arrastrados++;
        }
        CAT.camposTrabajoCortos.forEach(function (c) { if (reg[c] == null) reg[c] = ''; });
        salida.push(reg);
      });
    });

    return { rows: salida, arrastrados: arrastrados };
  }

  // Diferencias de campos de trabajo entre dos registros -> filas de bitacora
  function cambiosCortos(anterior, nuevo, snapshot, noEmpleado, turno, ahora, encargado) {
    var out = [];
    CAT.camposTrabajoCortos.forEach(function (c) {
      var a = texto(anterior[c]), n = texto(nuevo[c]);
      if (a !== n) {
        out.push({ Fecha: ahora, Snapshot: snapshot, Linea: nuevo.Linea, Item: nuevo.Item,
          CajaPlacas: nuevo.CajaPlacas, Campo: c, ValorAnt: a, ValorNvo: n,
          NoEmpleado: noEmpleado, Turno: turno, Encargado: encargado || noEmpleado });
      }
    });
    return out;
  }

  // Eventos para productividad (dentro del rango)
  // Descargados = Registro de piso; Ubicados = Ubicado; Cargados = Cargada;
  // Recolectados = cortos que pasan a Completo.
  function eventosProductividad(data, desde, hasta) {
    var out = [];
    var en = function (f) { return f instanceof Date && f >= desde && f <= hasta; };
    var mapaHist = { 'Registro': 'Descargados', 'Ubicado': 'Ubicados', 'Cargada': 'Cargados', 'Cambio estatus': 'Procesadas' };
    (data.Historial || []).forEach(function (h) {
      var tipo = mapaHist[h.Evento];
      if (!tipo || !en(h.FechaEvento)) return;
      var enc = h.Encargado || h.NoEmpleado;
      out.push({ tipo: tipo, fecha: h.FechaEvento, turno: h.Turno, noEmp: h.NoEmpleado, enc: enc, nombre: h.NombreUsuario });
      // Una descarga de camioneta cuenta tambien en "Camionetas" (subconjunto de Descargados)
      if (h.Evento === 'Registro' && h.Unidad === 'Camioneta') out.push({ tipo: 'Camionetas', fecha: h.FechaEvento, turno: h.Turno, noEmp: h.NoEmpleado, enc: enc, nombre: h.NombreUsuario });
    });
    (data.CortosLog || []).forEach(function (c) {
      if (c.Campo === 'Estatus' && c.ValorNvo === 'Completo' && en(c.Fecha)) {
        out.push({ tipo: 'Recolectados', fecha: c.Fecha, turno: c.Turno, noEmp: c.NoEmpleado, enc: c.Encargado || c.NoEmpleado, nombre: c.NoEmpleado });
      }
    });
    return out;
  }

  // Modulo al que pertenece un evento del historial (para Ultimas operaciones)
  function moduloDeEvento(evento) {
    var e = String(evento || '');
    if (e.indexOf('Factura') === 0) return 'Facturas';
    if (e.indexOf('Nota') === 0) return 'Notas';
    if (e.indexOf('Rampa') === 0) return 'Rampas';
    return 'Piso';
  }

  // ---- Datos iniciales y libro ----
  function datosVacios() {
    var d = {};
    Object.keys(SHEETS).forEach(function (n) {
      var campos = SHEETS[n];
      d[n] = (DEFAULTS[n] || []).map(function (arr) {
        var o = {};
        campos.forEach(function (c, i) { o[c] = arr[i] == null ? '' : arr[i]; });
        return o;
      });
    });
    return d;
  }

  // Objeto de datos -> bytes de un .xlsx (XLSX = libreria SheetJS)
  function aLibro(XLSX, data, hojas) {
    var wb = XLSX.utils.book_new();
    (hojas || Object.keys(SHEETS)).forEach(function (n) {
      var campos = SHEETS[n];
      var aoa = [campos].concat((data[n] || []).map(function (r) {
        return campos.map(function (c) { return r[c] == null ? '' : r[c]; });
      }));
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa, { cellDates: true }), n);
    });
    return XLSX.write(wb, { bookType: 'xlsx', type: 'array', cellDates: true, dateNF: 'dd/mm/yyyy hh:mm' });
  }

  // Bytes de un .xlsx -> objeto de datos
  function desdeLibro(XLSX, buf, hojas) {
    var wb = XLSX.read(new Uint8Array(buf), { type: 'array', cellDates: true });
    var d = {};
    (hojas || Object.keys(SHEETS)).forEach(function (n) {
      var campos = SHEETS[n];
      var ws = wb.Sheets[n];
      if (!ws) { d[n] = []; return; }
      d[n] = XLSX.utils.sheet_to_json(ws, { defval: '', raw: true }).map(function (r) {
        var o = {};
        campos.forEach(function (c) {
          var v = r[c];
          if (v === undefined) v = '';
          o[c] = TEXT_FIELDS.indexOf(c) >= 0 ? texto(v) : v;
        });
        return o;
      }).filter(function (o) { return campos.some(function (c) { return o[c] !== ''; }); });
    });
    return d;
  }

  // ---- Archivos de la carpeta compartida (SharePoint) ----
  // Cada archivo tiene un solo escritor a la vez: el que tiene su bloqueo. Los demas solo lo leen.
  var ARCHIVOS = {
    rampas: { archivo: 'Rampas_Bahias.xlsx', hojas: ['Usuarios', 'Turnos', 'Metas', 'DiasLaborales', 'Festivos',
      'Parametros', 'Bahias', 'Estatus', 'Rampas', 'EstatusRampas', 'Piso', 'Facturas'] },
    cortos: { archivo: 'Cortos.xlsx', hojas: ['Cortos', 'Notas'] },
    regC: { archivo: 'Registro_Capturista.xlsx', hojas: ['Historial'] },
    regS: { archivo: 'Registro_Supervisor.xlsx', hojas: ['Historial', 'CortosLog', 'CortosVersiones', 'CortosHistorico'] }
  };
  // Archivo de bloqueo por escritor: rampas (capturista o administrador) y cortos (supervisor)
  var BLOQUEOS = { rampas: 'Bloqueo_Rampas_Bahias.lock', cortos: 'Bloqueo_Cortos.lock' };
  // Modos de entrada permitidos por rol de usuario
  var MODOS = {
    Administrador: ['Administrador', 'Supervisor', 'Capturista'],
    Supervisor: ['Supervisor', 'Capturista'],
    Capturista: ['Capturista'],
    Consulta: ['Consulta']
  };
  var MINUTOS_INACTIVIDAD = 15;
  var MINUTOS_VENCE_LATIDO = 3;

  // Bloqueo que toma cada modo (null = no escribe nada)
  function bloqueoDeModo(modo) {
    if (modo === 'Supervisor') return 'cortos';
    if (modo === 'Capturista' || modo === 'Administrador') return 'rampas';
    return null;
  }
  // Un bloqueo sin latido reciente (computadora apagada o sin red) ya no cuenta
  function latidoVencido(bloqueo, ahora) {
    var t = bloqueo && bloqueo.latido ? Date.parse(bloqueo.latido) : NaN;
    return isNaN(t) || ahora.getTime() - t > MINUTOS_VENCE_LATIDO * 60000;
  }
  // Bloqueo vigente de otra computadora
  function bloqueoOcupado(bloqueo, miPc, ahora) {
    return !!(bloqueo && bloqueo.pc && bloqueo.pc !== miPc && !latidoVencido(bloqueo, ahora));
  }
  // Reparte una base de una sola pieza (QORVEX_Recibo.xlsx) entre los cuatro archivos.
  // El historial completo se pasa al registro del capturista; la bitacora de cortos al del supervisor.
  function repartirLegacy(d) {
    var tomar = function (hojas) { var o = {}; hojas.forEach(function (n) { o[n] = d[n] || []; }); return o; };
    var activos = (d.Cortos || []).filter(function (r) { return r.Activo === 'Sí'; });
    var historico = (d.Cortos || []).filter(function (r) { return r.Activo !== 'Sí'; });
    return {
      rampas: tomar(ARCHIVOS.rampas.hojas),
      cortos: { Cortos: activos, Notas: d.Notas || [] },
      regC: { Historial: d.Historial || [] },
      regS: { Historial: [], CortosLog: d.CortosLog || [], CortosVersiones: d.CortosSnapshots || [], CortosHistorico: historico }
    };
  }

  return {
    datosVacios: datosVacios, aLibro: aLibro, desdeLibro: desdeLibro,
    ARCHIVOS: ARCHIVOS, BLOQUEOS: BLOQUEOS, MODOS: MODOS, MINUTOS_INACTIVIDAD: MINUTOS_INACTIVIDAD,
    bloqueoDeModo: bloqueoDeModo, latidoVencido: latidoVencido, bloqueoOcupado: bloqueoOcupado, repartirLegacy: repartirLegacy,
    SHEETS: SHEETS, TEXT_FIELDS: TEXT_FIELDS, DEFAULTS: DEFAULTS, CAT: CAT, CORTOS_MAP: CORTOS_MAP, DIAS: DIAS,
    norm: norm, texto: texto, claveFecha: claveFecha, nombreDia: nombreDia, nombreMes: nombreMes,
    horasDesde: horasDesde, estadoTiempo: estadoTiempo, nextId: nextId, parametros: parametros,
    responsableDe: responsableDe, minutosHora: minutosHora, turnoEn: turnoEn, turnosDelDia: turnosDelDia,
    esFestivo: esFestivo, diaJornada: diaJornada, diasPorTurno: diasPorTurno, diasNoContados: diasNoContados,
    llaveCortos: llaveCortos, importarCortos: importarCortos, cambiosCortos: cambiosCortos,
    eventosProductividad: eventosProductividad, moduloDeEvento: moduloDeEvento
  };
});
