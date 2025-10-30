// Mostrar menú en la hoja de cálculo
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Adelantos')
    .addItem('Iniciar solicitud', 'mostrarFormulario')
    .addToUi();
}

// Renderizar formulario desde HTML
function doGet() {
  return HtmlService.createHtmlOutputFromFile('SolicitudAdelantoApp')
    .setTitle('Solicitud de Adelanto')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Mostrar formulario como modal en hoja
function mostrarFormulario() {
  const html = HtmlService.createHtmlOutputFromFile('SolicitudAdelantoApp')
    .setWidth(400)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Solicitud de Adelanto');
}

// Verificar si el usuario está registrado
function verificarUsuario(ci) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Usuarios");
  const datos = hoja.getDataRange().getValues();

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][7] == ci) {
      return {
        ok: true,
        datos: {
          habilitado: datos[i][0],
          empresa: datos[i][1],
          funcionario: datos[i][2],
          nombre1: datos[i][3],
          nombre2: datos[i][4],
          apellido1: datos[i][5],
          apellido2: datos[i][6],
          ci: datos[i][7],
          cel: datos[i][8],
          email: datos[i][9],
          secuencia: datos[i][11],
          depto: datos[i][12],
          sucursal: datos[i][13],
          seccion: datos[i][14],
          sucursal_nombre: datos[i][16]
        }
      };
    }
  }
  return { ok: false };
}

// Registrar una solicitud de adelanto
function registrarSolicitud(usuario, celularIngresado, monto) {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const hojaUsuarios = libro.getSheetByName("Usuarios");
  const fechaActual = new Date();
  const dia = fechaActual.getDate();

if ((dia >= 9 && dia <= 10) || (dia >= 19 && dia <= 20)) {
  return { 
    ok: false, 
    mensaje: "❌ Fuera de fecha habilitada para solicitar adelantos." 
  };
}

  const mes = fechaActual.toLocaleString('default', { month: 'long' });
  const anio = fechaActual.getFullYear();
  const nombreHoja = `${mes.charAt(0).toUpperCase() + mes.slice(1)} ${anio}`;

  if (Number(monto) > 30000) {
    return { ok: false, mensaje: "❌ El importe no puede superar los $30.000." };
  }

  if (usuario.habilitado !== true && usuario.habilitado !== "TRUE") {
    return { ok: false, mensaje: "❌ Usted no está habilitado para solicitar adelantos. Contacte a Recursos Humanos." };
  }

  let hojaSolicitudes = libro.getSheetByName(nombreHoja);
  if (!hojaSolicitudes) {
    hojaSolicitudes = libro.insertSheet(nombreHoja);
    hojaSolicitudes.clearContents();
    hojaSolicitudes.appendRow([
      'Nombre1', 'Nombre2', 'Apellido1', 'Apellido2', 'CI', 'CEL', 'EMAIL', 'FECHA',
      'Empresa', 'FUNCIONARIO', 'CONCEPTO', 'SECUENCIA', 'DEPARTAMENTO',
      'SUCURSAL', 'SECCION', 'IMPORTE', 'SUCURSAL_NOMBRE', 'CONFIRMACION'
    ]);
  }

  const datos = hojaUsuarios.getDataRange().getValues();

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][7] == usuario.ci) {
      const celReal = datos[i][8]?.toString().replace(/\D/g, '');
      const celIngresado = celularIngresado?.toString().replace(/\D/g, '');

      if (celIngresado === celReal) {
        let secuenciaValor = "";
        if (dia <= 9) {
          secuenciaValor = "12";
        } else if (dia <= 19) {
          secuenciaValor = "13";
        } else {
          secuenciaValor = "50";
        }

        const fechaFormateada = Utilities.formatDate(fechaActual, Session.getScriptTimeZone(), "yyyy/MM/dd");

        hojaSolicitudes.appendRow([
          datos[i][3], datos[i][4], datos[i][5], datos[i][6], datos[i][7],
          datos[i][8], datos[i][9], fechaFormateada, datos[i][1], datos[i][2],
          "15", secuenciaValor, datos[i][12], datos[i][13], datos[i][14],
          monto, datos[i][16], false
        ]);

        const fila = hojaSolicitudes.getLastRow();
        hojaSolicitudes.getRange(fila, 18).insertCheckboxes();

        return { ok: true, mensaje: "✅ Solicitud registrada con éxito." };
      } else {
        return { ok: false, mensaje: "❌ El celular no coincide. Contacte a Recursos Humanos." };
      }
    }
  }

  return { ok: false, mensaje: "❌ CI no encontrado en el sistema." };
}

// Registrar un nuevo usuario
function registrarNuevoUsuario(datos) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Usuarios");
  const data = hoja.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][7] == datos.ci) {
      return { ok: false, mensaje: "❌ El número de documento ya existe." };
    }
    if (data[i][8] == datos.cel) {
      return { ok: false, mensaje: "❌ El número de celular ya está registrado." };
    }
    if (data[i][9] == datos.email) {
      return { ok: false, mensaje: "❌ El email ya está registrado." };
    }
  }

  const fechaRegistro = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

  hoja.appendRow([
    '', datos.empresa || '', datos.funcionario || '', datos.nombre1, datos.nombre2,
    datos.apellido1, datos.apellido2, datos.ci, "'" + datos.cel, datos.email,
    '15', '', '999', '999', '99999', '', datos.sucursal || '', fechaRegistro
  ]);

  const ultimaFila = hoja.getLastRow();
  const celda = hoja.getRange(ultimaFila, 1);
  celda.insertCheckboxes();
  celda.setValue(true);

  return { ok: true, mensaje: "✅ Usuario registrado con éxito." };
}

// Obtener historial de solicitudes
function obtenerHistorial(ci) {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const hojas = libro.getSheets();
  let historial = [];

  hojas.forEach(hoja => {
    const nombre = hoja.getName();
    if (/^[A-Z][a-z]+ \d{4}$/.test(nombre)) {
      const datos = hoja.getDataRange().getValues();
      for (let i = 1; i < datos.length; i++) {
        if (datos[i][4] == ci) {
          historial.push({
            fecha: new Date(datos[i][7]),
            monto: datos[i][15],  // ahora sí es IMPORTE
            empresa: datos[i][8],
            funcionario: datos[i][9],
            nombre1: datos[i][0],
            nombre2: datos[i][1],
            apellido1: datos[i][2],
            apellido2: datos[i][3],
            cel: datos[i][5],
            email: datos[i][6],
            secuencia: datos[i][11],
            depto: datos[i][12],
            sucursal: datos[i][13],
            seccion: datos[i][14],
            sucursal_nombre: datos[i][16],
            hoja: nombre
          });
        }
      }
    }
  });

  historial.sort((a, b) => b.fecha - a.fecha);
  const ultimas4 = historial.slice(0, 4);
  ultimas4.forEach(item => {
    item.fecha = Utilities.formatDate(item.fecha, Session.getScriptTimeZone(), "yyyy/MM/dd");
  });

  return ultimas4;
}
