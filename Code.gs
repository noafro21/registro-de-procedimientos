// --- Funciones principales del menú ---
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Registro Médico")
    // Procedimientos
    .addItem("Registrar procedimiento", "mostrarFormularioProcedimientos")
    .addItem(
      "Generar reporte de Procedimientos",
      "mostrarReportePagoProcedimientos"
    )
    .addSeparator()
    // Ultrasonidos
    .addItem("Registrar ultrasonido", "mostrarFormularioUltrasonido")
    .addItem("Generar reporte de Radiología", "mostrarReportePagoUltrasonido")
    .addSeparator()
    // Horas Extras
    .addItem("Registrar horas extras", "mostrarFormularioHorasExtras")
    .addItem("Generar reporte de Horas Extras", "mostrarReportePagoHorasExtras")
    .addSeparator()
    // Biopsias
    .addItem("Registrar biopsia", "mostrarRegistroBiopsias")
    .addItem("Buscar y reportar biopsias", "mostrarBuscarBiopsias")
    .addSeparator()
    .addItem("Inicializar checkboxes de biopsias", "inicializarCheckboxes")
    .addToUi();
}

/**
 * Función para obtener el HTML de una página específica
 * Esta función es llamada desde el cliente en el entorno de desarrollo
 * @param {string} pageName - Nombre de la página a cargar
 * @returns {string} HTML de la página solicitada
 */
function getPageHtml(pageName) {
  Logger.log(
    "DEBUG (Server): getPageHtml() ejecutada para la página: " + pageName
  );

  // Verificar si pageName es undefined o null y establecer un valor por defecto
  if (!pageName) {
    Logger.log(
      "ADVERTENCIA: pageName es undefined o null, usando 'mainMenu' como valor predeterminado"
    );
    pageName = "mainMenu";
  }

  let template;
  let pageTitle = "Registro Médico"; // Título por defecto

  try {
    switch (pageName) {
      case "mainMenu":
      case "mainMenuContent": // Para el botón de regresar al menú principal
        template = HtmlService.createTemplateFromFile("mainMenu");
        pageTitle = "Menú Principal";
        break;
      case "formularioProcedimientos":
        template = HtmlService.createTemplateFromFile(
          "formularioProcedimientos"
        );
        pageTitle = "Registrar Procedimiento";
        break;
      case "reportePagoProcedimientos":
        template = HtmlService.createTemplateFromFile(
          "reportePagoProcedimientos"
        );
        pageTitle = "Reporte de Procedimientos";
        break;
      case "formularioUltrasonido":
        template = HtmlService.createTemplateFromFile("formularioUltrasonido");
        pageTitle = "Registrar Ultrasonido";
        break;
      case "reportePagoUltrasonido":
        template = HtmlService.createTemplateFromFile("reportePagoUltrasonido");
        pageTitle = "Reporte de Radiología";
        break;
      case "formularioHorasExtras":
        template = HtmlService.createTemplateFromFile("formularioHorasExtras");
        pageTitle = "Registrar Horas Extras";
        break;
      case "reportePagoHorasExtras":
        template = HtmlService.createTemplateFromFile("reportePagoHorasExtras");
        pageTitle = "Reporte de Horas Extras";
        break;
      case "registroBiopsias": // Nuevo
        template = HtmlService.createTemplateFromFile("registroBiopsias");
        pageTitle = "Registro de Biopsias";
        break;
      case "buscarBiopsias": // Nuevo
        template = HtmlService.createTemplateFromFile("buscarBiopsiasReg");
        pageTitle = "Búsqueda y Reporte de Biopsias";
        break;
      default:
        Logger.log(
          "ERROR (Server): Página no reconocida en getPageHtml(): " + pageName
        );
        // En lugar de lanzar un error, cargar el menú principal
        template = HtmlService.createTemplateFromFile("mainMenu");
        pageTitle = "Menú Principal";
        break;
    }

    const htmlOutput = template.evaluate().setTitle(pageTitle);
    return htmlOutput.getContent();
  } catch (error) {
    Logger.log(
      'ERROR (Server): Error al obtener HTML para la página "' +
        pageName +
        '": ' +
        error.message
    );
    throw new Error("No se pudo cargar la página solicitada: " + error.message);
  }
}

/**
 * Función que se ejecuta cuando la aplicación web es accedida.
 * Sirve diferentes páginas HTML basadas en el parámetro 'page' en la URL.
 * Si no hay parámetro 'page', muestra el menú principal.
 * @param {Object} e Objeto de evento de la solicitud HTTP GET.
 * @returns {HtmlOutput} Contenido HTML de la página.
 */
function doGet(e) {
  Logger.log(
    "DEBUG (Server): doGet() ejecutada con parámetros: " + JSON.stringify(e)
  );

  e = e || {};
  const page = e.parameter ? e.parameter.page : null;

  Logger.log('DEBUG (Server): Parámetro "page" recibido: ' + page);

  let template;
  let pageTitle = "Registro Médico";

  try {
    switch (page) {
      case "formularioProcedimientos":
        template = HtmlService.createTemplateFromFile(
          "formularioProcedimientos"
        );
        pageTitle = "Registrar Procedimiento";
        break;
      case "reportePagoProcedimientos":
        template = HtmlService.createTemplateFromFile(
          "reportePagoProcedimientos"
        );
        pageTitle = "Reporte de Procedimientos";
        break;
      case "formularioUltrasonido":
        template = HtmlService.createTemplateFromFile("formularioUltrasonido");
        pageTitle = "Registrar Ultrasonido";
        break;
      case "reportePagoUltrasonido":
        template = HtmlService.createTemplateFromFile("reportePagoUltrasonido");
        pageTitle = "Reporte de Radiología";
        break;
      case "formularioHorasExtras":
        template = HtmlService.createTemplateFromFile("formularioHorasExtras");
        pageTitle = "Registrar Horas Extras";
        break;
      case "reportePagoHorasExtras":
        template = HtmlService.createTemplateFromFile("reportePagoHorasExtras");
        pageTitle = "Reporte de Horas Extras";
        break;
      case "registroBiopsias": // Nuevo
        template = HtmlService.createTemplateFromFile("registroBiopsias");
        pageTitle = "Registro de Biopsias";
        break;
      case "buscarBiopsias": // Nuevo
        template = HtmlService.createTemplateFromFile("buscarBiopsiasReg");
        pageTitle = "Búsqueda y Reporte de Biopsias";
        break;
      default:
        template = HtmlService.createTemplateFromFile("mainMenu");
        pageTitle = "Menú Principal";
        break;
    }

    Logger.log(
      "DEBUG (Server): Cargando archivo HTML: " + (page || "mainMenu") + ".html"
    );

    const htmlOutput = template
      .evaluate()
      .setTitle(pageTitle)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    return htmlOutput;
  } catch (error) {
    Logger.log(
      'ERROR (Server): Error al cargar la página "' +
        page +
        '": ' +
        error.message
    );

    return HtmlService.createHtmlOutput(
      `
      <h1>Error al cargar la página</h1>
      <p>Ha ocurrido un problema al intentar mostrar esta sección.</p>
      <p>Mensaje de error: ${error.message}</p>
      <p>Por favor, contacte al administrador del sistema.</p>
      <button onclick="window.top.location.href='${ScriptApp.getService().getUrl()}'">Volver al Menú Principal</button>
    `
    ).setTitle("Error");
  }
}

/**
 * Función para incluir el contenido de otros archivos HTML dentro de una plantilla.
 * Este contenido es siempre un FRAGMENTO HTML (sin <html>, <head>, <body>).
 * @param {string} filename - El nombre base del archivo HTML a incluir (sin extensión .html).
 * @returns {string} El contenido HTML del archivo.
 */
function include(filename) {
  Logger.log(
    "DEBUG (Server): Intentando incluir el archivo: " + filename + ".html"
  );

  // Verificar si filename es undefined o null
  if (!filename) {
    Logger.log(
      "ERROR (Server): Nombre de archivo indefinido o nulo en la función include()"
    );
    return "<!-- ERROR: Nombre de archivo indefinido o nulo -->";
  }

  try {
    // === MENÚ PRINCIPAL ===
    if (filename === "stylesMainMenu") filename = "stylesMainMenu";
    if (filename === "scriptMainMenu") filename = "scriptMainMenu";

    // === FORMULARIOS DE PROCEDIMIENTOS ===
    if (filename === "stylesFormulario")
      filename = "stylesFormularioProcedimientos";
    if (filename === "scriptFormulario")
      filename = "scriptFormularioProcedimientos";

    // === ULTRASONIDO ===
    if (filename === "stylesUltrasonido")
      filename = "stylesFormularioUltrasonido";
    if (filename === "scriptUltrasonido")
      filename = "scriptFormularioUltrasonido";

    // === HORAS EXTRAS ===
    if (filename === "stylesHorasExtras") filename = "stylesHorasExtras";
    if (filename === "scriptHorasExtras") filename = "scriptHorasExtras";

    // === REPORTE COSTOS DE PROCEDIMIENTOS ===
    if (filename === "stylesCostos")
      filename = "stylesReportePagoProcedimientos";
    if (filename === "scriptCostos")
      filename = "scriptReportePagoProcedimientos";

    // === BOTÓN REGRESAR ===
    if (filename === "scriptBotonRegresar") filename = "scriptBotonRegresar";

    // === BIOPSIAS ===
    if (filename === "stylesRegistroBiopsias")
      filename = "stylesRegistroBiopsias";
    if (filename === "scriptRegistroBiopsias")
      filename = "scriptRegistroBiopsias";
    if (filename === "stylesBuscarBiopsias") filename = "stylesBuscarBiopsias";
    if (filename === "scriptBuscarBiopsias") filename = "scriptBuscarBiopsias";

    // === RETORNO DEL CONTENIDO HTML ===
    return HtmlService.createHtmlOutputFromFile(
      filename + ".html"
    ).getContent();
  } catch (e) {
    Logger.log(
      "ERROR (Server): Falló la inclusión de " +
        filename +
        ".html. Es posible que el archivo no exista o tenga un nombre incorrecto. Mensaje: " +
        e.message
    );
    return `<!-- ERROR: No se pudo incluir ${filename}.html: ${e.message} -->`;
  }
}

// --- Funciones de utilidad general ---

/**
 * Función global para normalizar nombres
 * @param {string} str - Cadena a normalizar
 * @returns {string} Cadena normalizada
 */
function normalizarNombre(str) {
  return String(str || "").replace(/\s+/g, " ").trim().toUpperCase();
}

/**
 * Obtiene una lista de nombres de personal de la hoja 'Personal', filtrando por columnas específicas.
 * @param {number[]} columnIndexes Un array de índices de columnas (base 0) de la hoja 'Personal' a incluir.
 * @returns {string[]} Un array de nombres de personal únicos y ordenados.
 */
function obtenerPersonalFiltrado(columnIndexes) {
  if (!Array.isArray(columnIndexes)) {
    columnIndexes = [0, 1, 2];
    Logger.log(
      "ADVERTENCIA: columnIndexes no recibido, usando [0,1,2] por defecto"
    );
  }

  try {
    Logger.log(
      "obtenerPersonalFiltrado llamado con columnIndexes: " + columnIndexes
    );

    // Verificar que columnIndexes sea un array válido
    if (!columnIndexes || !Array.isArray(columnIndexes)) {
      Logger.log("ERROR: columnIndexes no es un array válido");
      return []; // Devolver array vacío en lugar de lanzar error
    }

    const hoja =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Personal");
    if (!hoja) {
      Logger.log("ERROR: La hoja 'Personal' no existe.");
      throw new Error("❌ La hoja 'Personal' no existe.");
    }

    Logger.log("Hoja 'Personal' encontrada");
    const datos = hoja.getDataRange().getValues();
    Logger.log("Datos obtenidos de la hoja 'Personal'. Filas: " + datos.length);

    const personal = new Set();
    for (let i = 1; i < datos.length; i++) {
      // Iterar desde la segunda fila (ignorar encabezado)
      const fila = datos[i];
      columnIndexes.forEach((colIndex) => {
        // Verificar que el índice de columna sea válido
        if (colIndex >= 0 && colIndex < fila.length) {
          const nombre = fila[colIndex];
          if (nombre && typeof nombre === "string" && nombre.trim()) {
            personal.add(nombre.trim());
            Logger.log(
              "Añadido personal: " +
                nombre.trim() +
                " de la columna " +
                colIndex
            );
          }
        } else {
          Logger.log(
            "ADVERTENCIA: Índice de columna fuera de rango: " + colIndex
          );
        }
      });
    }

    const resultado = [...personal].sort();
    Logger.log("Nombres de personal filtrados: " + resultado.join(", "));

    return resultado;
  } catch (error) {
    Logger.log("ERROR en obtenerPersonalFiltrado: " + error.message);
    // En lugar de propagar el error, devolver un array vacío
    return [];
  }
}

/**
 * Obtiene el email asociado a un nombre de personal en la hoja 'Personal'.
 * @param {string} nombrePersonal El nombre del personal a buscar.
 * @returns {string} El email del personal o una cadena vacía si no se encuentra.
 */
function obtenerEmailParaPersonal(nombrePersonal) {
  const hojaPersonal =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Personal");
  if (!hojaPersonal) {
    Logger.log("❌ La hoja 'Personal' no existe al intentar obtener el email.");
    return "";
  }
  const datos = hojaPersonal.getDataRange().getValues();
  const emailColIndex = 6; // Columna G (índice 6)

  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    for (let col = 0; col <= 5; col++) {
      // Búsqueda en columnas de nombres (A-F)
      if (fila[col] && fila[col].trim() === nombrePersonal) {
        return fila[emailColIndex] ? String(fila[emailColIndex]).trim() : "";
      }
    }
  }
  return "";
}

/**
 * Genera un PDF a partir de contenido HTML y lo envía por correo electrónico.
 * @param {string} htmlContent El contenido HTML del reporte (solo el body).
 * @param {string} styleFileName El nombre del archivo de estilo HTML (.html) a incrustar.
 * @param {string} pageTitle El título de la página del PDF.
 * @param {string} recipientEmail La dirección de correo electrónico del destinatario.
 * @param {string} subject El asunto del correo electrónico.
 * @param {string} fileName El nombre del archivo PDF (ej. "Reporte_Mes.pdf").
 * @returns {object} Un objeto indicando el éxito y un mensaje.
 */
function enviarReportePDF(
  htmlContent,
  styleFileName,
  pageTitle,
  recipientEmail,
  subject,
  fileName
) {
  try {
    const styleContent =
      HtmlService.createHtmlOutputFromFile(styleFileName).getContent();
    const fullHtml = `
      <!DOCTYPE html>
      <html>
      <head>
        <title>${pageTitle}</title>
        <style>${styleContent}</style>
      </head>
      <body>
        ${htmlContent}
      </body>
      </html>
    `;
    const htmlOutput = HtmlService.createHtmlOutput(fullHtml);
    const pdfBlob = htmlOutput.getAs(MimeType.PDF).setName(fileName);

    MailApp.sendEmail({
      to: recipientEmail,
      subject: subject,
      htmlBody:
        "Adjunto su reporte en formato PDF.<br><br>Este correo fue generado automáticamente. Por favor, no responda a este email.",
      attachments: [pdfBlob],
    });

    Logger.log("✅ Reporte PDF enviado con éxito a: " + recipientEmail);
    return { success: true, message: "✅ Reporte enviado con éxito." };
  } catch (e) {
    Logger.log("Error al enviar el reporte PDF: " + e.message);
    return {
      success: false,
      message:
        "❌ Error al enviar el reporte: " +
        e.message +
        ". Por favor, revise la dirección de correo y los permisos de su script.",
    };
  }
}

/**
 * Genera un PDF a partir de contenido HTML y lo devuelve como una cadena Base64.
 * Esto es útil para descargar el PDF directamente desde el cliente.
 * @param {string} htmlContent El contenido HTML del reporte (solo el body).
 * @param {string} fileName El nombre deseado para el archivo PDF.
 * @param {string} styleFileName El nombre del archivo de estilo HTML (.html) a incrustar.
 * @param {string} pageTitle El título de la página del PDF.
 * @returns {string} El contenido del PDF en formato Base64.
 */
function generarPdfParaDescarga(
  htmlContent,
  fileName,
  styleFileName,
  pageTitle = "Reporte"
) {
  try {
    // Verificar parámetros
    if (!htmlContent) {
      throw new Error("El contenido HTML no puede estar vacío");
    }

    if (!fileName) {
      fileName = "Reporte.pdf";
    }

    if (!styleFileName) {
      styleFileName = "stylesReportePagoProcedimientos"; // Estilo por defecto
    }

    // Obtener el contenido del archivo de estilo
    let styleContent = "";
    try {
      styleContent =
        HtmlService.createHtmlOutputFromFile(styleFileName).getContent();
    } catch (styleError) {
      Logger.log(
        "Advertencia: No se pudo cargar el archivo de estilo: " +
          styleError.message
      );
      // Continuar sin estilos si no se puede cargar el archivo
    }

    // Construir el HTML completo
    const fullHtml = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="UTF-8">
        <title>${pageTitle}</title>
        <style>${styleContent}</style>
      </head>
      <body>
        ${htmlContent}
      </body>
      </html>
    `;

    Logger.log("DEBUG (Server): Generando PDF para: " + fileName);

    // Crear el PDF
    const htmlOutput = HtmlService.createHtmlOutput(fullHtml);
    const pdfBlob = htmlOutput.getAs(MimeType.PDF).setName(fileName);

    Logger.log("✅ PDF generado para descarga: " + fileName);

    // Convertir a Base64
    return Utilities.base64Encode(pdfBlob.getBytes());
  } catch (e) {
    Logger.log("Error al generar PDF para descarga: " + e.message);
    throw new Error("❌ Error al generar el PDF: " + e.message);
  }
}

// --- Funciones para el Módulo de Procedimientos ---
function mostrarFormularioProcedimientos() {
  try {
    Logger.log("Iniciando mostrarFormularioProcedimientos()");

    // Crear el HTML desde la plantilla
    const html = HtmlService.createTemplateFromFile("formularioProcedimientos")
      .evaluate()
      .setWidth(900) // Aumentar el ancho para mejor visualización
      .setHeight(800) // Aumentar la altura para mostrar más contenido
      .setSandboxMode(HtmlService.SandboxMode.IFRAME); // Asegurar que se cargue en un iframe

    // Mostrar el diálogo
    SpreadsheetApp.getUi().showModalDialog(html, "Registro de Procedimientos");

    Logger.log("Formulario de procedimientos mostrado correctamente");
  } catch (error) {
    Logger.log(
      "ERROR al mostrar formulario de procedimientos: " + error.message
    );
    SpreadsheetApp.getUi().alert(
      "Error al mostrar el formulario: " + error.message
    );
  }
}

/**
 * Agrega un registro al final de la hoja "RegistrosProcedimientos".
 * Método tradicional que no altera el orden de los registros existentes.
 * @param {Object} data - Datos del registro.
 * @returns {boolean} - Verdadero si se guardó correctamente.
 */
function guardarRegistroCompletoArriba(data) {
  try {
    Logger.log(
      "Insertando registro en la hoja 'RegistrosProcedimientos' (arriba): " +
        JSON.stringify(data)
    );

    if (!data || typeof data !== "object") {
      throw new Error("No se recibieron datos válidos para el registro.");
    }

    const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      "RegistrosProcedimientos"
    );
    if (!hoja) {
      throw new Error("La hoja 'RegistrosProcedimientos' no existe.");
    }

    // Validar y obtener la fecha
    const fechaStr = data.fecha || "";
    if (!fechaStr || typeof fechaStr !== "string") {
      throw new Error(
        "El campo de fecha es obligatorio y debe tener formato YYYY-MM-DD."
      );
    }
    let fechaRegistro;
    try {
      let [year, month, day] = fechaStr.split("-").map(Number);
      fechaRegistro = new Date(Date.UTC(year, month - 1, day, 12, 0, 0));
    } catch (e) {
      throw new Error("Formato de fecha inválido: " + e.message);
    }

    // Preparar la fila para añadir (ajusta los campos según tu estructura)
    const fila = [
      fechaRegistro,
      data.personal || "",
      data.consulta_regular || 0,
      data.consulta_medismart || 0,
      data.consulta_higado || 0,
      data.consulta_pylori || 0,
      data.gastro_regular || 0,
      data.gastro_medismart || 0,
      data.colono_regular || 0,
      data.colono_medismart || 0,
      data.gastocolono_regular || 0,
      data.gastocolono_medismart || 0,
      data.recto_regular || 0,
      data.recto_medismart || 0,
      data.asa_fria || 0,
      data.asa_fria2 || 0,
      data.asa_caliente || 0,
      data.dictamen || 0,
      data.comentario || "",
    ];

    // Agregar el registro al final de la hoja (método tradicional)
    const ultimaFila = hoja.getLastRow();
    hoja.getRange(ultimaFila + 1, 1, 1, fila.length).setValues([fila]);

    // Formatear la celda de fecha
    hoja.getRange(ultimaFila + 1, 1).setNumberFormat("dd/MM/yyyy");

    Logger.log("✅ Registro agregado al final de la hoja con éxito.");
    return true;
  } catch (error) {
    Logger.log("ERROR al insertar registro: " + error.message);
    throw new Error("Error al guardar el registro: " + error.message);
  }
}

function mostrarReportePagoProcedimientos() {
  const html = HtmlService.createTemplateFromFile("reportePagoProcedimientos")
    .evaluate()
    .setWidth(900)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, "Reporte de Costos");
}

function obtenerTipoDePersonal(nombre, hojaPersonal) {
  if (!hojaPersonal) {
    hojaPersonal =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Personal");
    if (!hojaPersonal) {
      Logger.log("❌ La hoja 'Personal' no existe.");
      return null;
    }
  }
  const datos = hojaPersonal.getDataRange().getValues();
  const nombreNorm = String(nombre || "")
    .trim()
    .toUpperCase();
  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    for (let col = 0; col <= 5; col++) {
      const celdaNorm = String(fila[col] || "")
        .trim()
        .toUpperCase();
      if (celdaNorm === nombreNorm) {
        switch (col) {
          case 0:
            return "Doctor";
          case 1:
            return "Anestesiólogo";
          case 2:
            return "Técnico";
          case 3:
            return "Radiólogo";
          case 4:
            return "Enfermero";
          case 5:
            return "Secretaria";
        }
      }
    }
  }
  return null;
}

// ...existing code...

function calcularCostos(nombre, mes, anio, quincena) {
  if (!nombre || !mes || !anio || !quincena) {
    Logger.log("❌ Parámetros incompletos en calcularCostos");
    return {
      lv: {},
      sab: {},
      totales: { subtotal: 0, iva: 0, total_con_iva: 0 },
    };
  }

  Logger.log(
    "🚀 INICIANDO calcularCostos: nombre=" +
      nombre +
      ", mes=" +
      mes +
      ", anio=" +
      anio +
      ", quincena=" +
      quincena
  );
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaRegistrosProcedimientos = ss.getSheetByName("RegistrosProcedimientos");
  const hojaPrecios = ss.getSheetByName("Precios");
  const hojaPersonal = ss.getSheetByName("Personal");

  if (!hojaRegistrosProcedimientos || !hojaPrecios || !hojaPersonal) {
    throw new Error(
      "Faltan hojas necesarias: RegistrosProcedimientos, Precios o Personal"
    );
  }

  const datos = hojaRegistrosProcedimientos.getDataRange().getValues();
  const preciosDatos = hojaPrecios.getDataRange().getValues();
  
  Logger.log(`📊 Total de filas en RegistrosProcedimientos: ${datos.length}`);
  Logger.log(`🔍 Nombre normalizado para búsqueda: "${normalizarNombre(nombre)}"`);
  
  // Mostrar muestra de nombres en las primeras 10 filas
  Logger.log(`📋 MUESTRA DE NOMBRES EN REGISTROS:`);
  for (let i = 1; i <= Math.min(10, datos.length - 1); i++) {
    const nombreFila = String(datos[i][1] || "").trim();
    const fechaFila = datos[i][0];
    Logger.log(`   Fila ${i + 1}: "${nombreFila}" - ${fechaFila instanceof Date ? fechaFila.toDateString() : fechaFila}`);
  }
  
  const preciosPorProcedimiento = {};
  
  preciosDatos.slice(1).forEach((fila) => {
    preciosPorProcedimiento[fila[0]] = {
      doctorLV: fila[1] || 0,
      doctorSab: fila[2] || 0,
      anest: fila[3] || 0,
      tecnico: fila[4] || 0,
    };
  });

  function normalizarNombre(str) {
    return String(str || "")
      .replace(/\s+/g, " ")
      .trim()
      .toUpperCase();
  }

  // ✅ FUNCIÓN MEJORADA: Búsqueda más flexible pero precisa
  function esLaMismaPersona(nombreSeleccionado, nombreEnRegistro) {
    // Manejar valores nulos/undefined
    if (!nombreSeleccionado || !nombreEnRegistro) {
      return false;
    }
    
    const seleccionadoNorm = normalizarNombre(nombreSeleccionado);
    const registroNorm = normalizarNombre(nombreEnRegistro);
    
    // 1. Coincidencia exacta (preferida)
    if (seleccionadoNorm === registroNorm) {
      Logger.log(`✅ COINCIDENCIA EXACTA: "${seleccionadoNorm}"`);
      return true;
    }
    
    // 2. Uno contiene al otro (para nombres parciales)
    if (seleccionadoNorm.includes(registroNorm) || registroNorm.includes(seleccionadoNorm)) {
      Logger.log(`✅ COINCIDENCIA PARCIAL: "${seleccionadoNorm}" ↔ "${registroNorm}"`);
      return true;
    }
    
    // 3. Coincidencia por palabras clave (para nombres compuestos)
    const palabrasSeleccionado = seleccionadoNorm.split(' ').filter(p => p.length > 2);
    const palabrasRegistro = registroNorm.split(' ').filter(p => p.length > 2);
    
    if (palabrasSeleccionado.length > 0 && palabrasRegistro.length > 0) {
      const coincidencias = palabrasSeleccionado.filter(p1 => 
        palabrasRegistro.some(p2 => p1.includes(p2) || p2.includes(p1))
      );
      
      // Si al menos 50% de las palabras coinciden
      if (coincidencias.length >= Math.min(palabrasSeleccionado.length, palabrasRegistro.length) * 0.5) {
        Logger.log(`✅ COINCIDENCIA POR PALABRAS: "${seleccionadoNorm}" ↔ "${registroNorm}"`);
        return true;
      }
    }
    
    return false;
  }

  // Obtener el tipo de personal
  const tipoPersonal = obtenerTipoDePersonal(nombre, hojaPersonal);
  
  if (!tipoPersonal) {
    Logger.log("❌ El nombre '" + nombre + "' no está registrado en la hoja Personal.");
    return {
      lv: {},
      sab: {},
      totales: { subtotal: 0, iva: 0, total_con_iva: 0 },
    };
  }

  Logger.log(`🏥 Tipo de personal: "${tipoPersonal}"`);

  const mapeoProcedimientos = {
    consulta_regular: "Consulta Regular",
    consulta_medismart: "Consulta MediSmart",
    consulta_higado: "Consulta Hígado",
    consulta_pylori: "Consulta H. Pylori",
    gastro_regular: "Gastroscopía Regular",
    gastro_medismart: "Gastroscopía MediSmart",
    colono_regular: "Colonoscopía Regular",
    colono_medismart: "Colonoscopía MediSmart",
    gastocolono_regular: "Gastrocolonoscopía Regular",
    gastocolono_medismart: "Gastrocolonoscopía MediSmart",
    recto_regular: "Rectoscopía Regular",
    recto_medismart: "Rectoscopía MediSmart",
    asa_fria: "Proc. Terapéutico Fría Menor",
    asa_fria2: "Proc. Terapéutico Fría Mayor",
    asa_caliente: "Proc. Terapéutico Térmica",
    dictamen: "Dictamen",
  };

  const resumen = {
    lv: {},
    sab: {},
    totales: { subtotal: 0, iva: 0, total_con_iva: 0 },
  };

  Object.keys(mapeoProcedimientos).forEach((key) => {
    resumen.lv[key] = {
      nombre: mapeoProcedimientos[key],
      cantidad: 0,
      costo_unitario: 0,
      subtotal: 0,
      iva: 0,
      total_con_iva: 0,
    };
    resumen.sab[key] = {
      nombre: mapeoProcedimientos[key],
      cantidad: 0,
      costo_unitario: 0,
      subtotal: 0,
      iva: 0,
      total_con_iva: 0,
    };
  });

  const iva = 0.04;
  let registrosEncontrados = 0;
  let registrosRechazados = 0;
  let filasConFechaInvalida = 0;

  Logger.log(`📋 Buscando registros para: "${nombre}" en ${mes}/${anio}, quincena: ${quincena}`);

  // ✅ PROCESAR CADA FILA CON MANEJO DE ERRORES ROBUSTO
  datos.slice(1).forEach((fila, idx) => {
    const numeroFila = idx + 2; // +2 porque idx empieza en 0 y saltamos encabezados
    
    try {
      // Validar y procesar fecha
      let fecha = fila[0];
      
      if (!fecha) {
        Logger.log(`⚠️ Fila ${numeroFila}: Fecha vacía, saltando...`);
        registrosRechazados++;
        return;
      }
      
      // Convertir a Date si no lo es
      if (!(fecha instanceof Date)) {
        fecha = new Date(fecha);
      }
      
      if (isNaN(fecha.getTime())) {
        Logger.log(`❌ Fila ${numeroFila}: Fecha inválida: ${fila[0]}`);
        filasConFechaInvalida++;
        return;
      }
      
      const filaMes = fecha.getMonth() + 1;
      const filaAnio = fecha.getFullYear();
      const filaDia = fecha.getDate();
      const personaEnRegistro = fila[1];

      // ✅ VERIFICAR CADA CONDICIÓN
      const esPersonaCoincide = esLaMismaPersona(nombre, personaEnRegistro);
      const esMismoMes = Number(filaMes) === Number(mes);
      const esMismoAnio = Number(filaAnio) === Number(anio);

      // Log solo si hay coincidencia de persona (para reducir spam)
      if (esPersonaCoincide) {
        Logger.log(`📅 Fila ${numeroFila}: PERSONA COINCIDE - "${personaEnRegistro}", Fecha: ${fecha.toDateString()}`);
        Logger.log(`   🔍 Mes (${filaMes} === ${mes}): ${esMismoMes}, Año (${filaAnio} === ${anio}): ${esMismoAnio}`);
      }

      if (!esPersonaCoincide || !esMismoMes || !esMismoAnio) {
        if (esPersonaCoincide) {
          Logger.log(`   ❌ Rechazado por fecha`);
        }
        registrosRechazados++;
        return;
      }

      // ✅ VERIFICAR QUINCENA
      if (quincena === "1-15" && filaDia > 15) {
        Logger.log(`   ❌ Día ${filaDia} no está en quincena 1-15`);
        registrosRechazados++;
        return;
      }
      if (quincena === "16-31" && filaDia <= 15) {
        Logger.log(`   ❌ Día ${filaDia} no está en quincena 16-31`);
        registrosRechazados++;
        return;
      }

      registrosEncontrados++;
      Logger.log(`   ✅ REGISTRO VÁLIDO #${registrosEncontrados} - Procesando procedimientos...`);

      const esSabado = fecha.getDay() === 6;
      let colIndex = 2;
      let procedimientosEnFila = 0;
      
      for (const key of Object.keys(mapeoProcedimientos)) {
        const cantidad = Number(fila[colIndex]) || 0;
        if (cantidad > 0) {
          procedimientosEnFila++;
          const procNombre = mapeoProcedimientos[key];
          let costo = 0;

          const precios = preciosPorProcedimiento[procNombre] || {};
          if (tipoPersonal === "Doctor") {
            costo = esSabado ? precios.doctorSab : precios.doctorLV;
          } else if (tipoPersonal === "Anestesiólogo") {
            costo = precios.anest;
          } else if (tipoPersonal === "Técnico") {
            costo = precios.tecnico;
          }

          const subtotal = cantidad * costo;
          const ivaMonto = subtotal * iva;
          const total = subtotal + ivaMonto;

          const targetResumen = esSabado ? resumen.sab : resumen.lv;
          targetResumen[key].cantidad += cantidad;
          targetResumen[key].costo_unitario = costo;
          targetResumen[key].subtotal += subtotal;
          targetResumen[key].iva += ivaMonto;
          targetResumen[key].total_con_iva += total;

          resumen.totales.subtotal += subtotal;
          resumen.totales.iva += ivaMonto;
          resumen.totales.total_con_iva += total;

          Logger.log(`     💰 ${procNombre}: ${cantidad} x ${costo} = $${subtotal}`);
        }
        colIndex++;
      }
      
      if (procedimientosEnFila === 0) {
        Logger.log(`   ⚠️ Fila válida pero sin procedimientos registrados`);
      }
      
    } catch (error) {
      Logger.log(`❌ Error procesando fila ${numeroFila}: ${error.message}`);
      registrosRechazados++;
    }
  });
  
  Logger.log(`\n🎯 RESUMEN FINAL:`);
  Logger.log(`   ✅ Registros VÁLIDOS encontrados: ${registrosEncontrados}`);
  Logger.log(`   ❌ Registros RECHAZADOS: ${registrosRechazados}`);
  Logger.log(`   📅 Filas con fecha inválida: ${filasConFechaInvalida}`);
  Logger.log(`   💰 Total calculado: $${resumen.totales.total_con_iva.toFixed(2)}`);
  
  if (registrosEncontrados === 0) {
    Logger.log(`\n🔍 EJECUTANDO DIAGNÓSTICO COMPLETO...`);
    
    // Ejecutar diagnóstico completo
    const diagnostico = diagnosticoCompleto(nombre, mes, anio);
    
    Logger.log(`\n📋 DIAGNÓSTICO COMPLETO COMPLETADO`);
    Logger.log(`Problemas identificados: ${diagnostico.problemas ? diagnostico.problemas.length : 0}`);
    
    if (diagnostico.problemas && diagnostico.problemas.length > 0) {
      Logger.log("\n❌ PROBLEMAS ENCONTRADOS:");
      diagnostico.problemas.forEach(problema => {
        Logger.log("   " + problema);
      });
    }
    
    // Mostrar recomendaciones si existen
    if (diagnostico.recomendaciones && diagnostico.recomendaciones.length > 0) {
      Logger.log("\n💡 RECOMENDACIONES:");
      diagnostico.recomendaciones.forEach(recomendacion => {
        Logger.log("   " + recomendacion);
      });
    }
    
    // Mostrar coincidencias de nombres encontradas
    if (diagnostico.analisisRegistros && diagnostico.analisisRegistros.coincidenciasNombre.length > 0) {
      Logger.log("\n💡 NOMBRES SIMILARES ENCONTRADOS:");
      diagnostico.analisisRegistros.coincidenciasNombre.slice(0, 5).forEach(c => {
        Logger.log(`   Fila ${c.fila}: "${c.nombre}" (${c.fecha})`);
      });
    }
  }
  
  return resumen;
}

// ...existing code...

// --- Funciones Wrapper para Enviar/Descargar Reportes de Procedimientos ---
function enviarReporteCostos(
  htmlContent,
  emailRecipient,
  reportTitle,
  subject,
  fileName
) {
  return enviarReportePDF(
    htmlContent,
    "stylesReportePagoProcedimientos",
    reportTitle,
    emailRecipient,
    subject,
    fileName
  );
}

/**
 * Función de diagnóstico completo para identificar problemas de búsqueda
 * @param {string} nombrePersonal - Nombre del personal a diagnosticar
 * @param {number} mes - Mes a buscar
 * @param {number} anio - Año a buscar
 * @param {string} quincena - Quincena a buscar
 * @returns {Object} Información detallada de diagnóstico
 */
/**
 * Función simplificada para diagnosticar desde el frontend
 * @param {string} nombrePersonal - Nombre del personal
 * @returns {Object} Resultado del diagnóstico
 */
/**
 * Función de emergencia para diagnóstico inmediato - llamar desde el frontend
 * @param {string} nombrePersonal - Nombre a buscar
 * @param {number} mes - Mes a buscar  
 * @param {number} anio - Año a buscar
 * @returns {Object} Información de diagnóstico inmediato
 */
/**
 * Función consolidada de diagnóstico completo del sistema
 * @param {string} nombrePersonal - Nombre a buscar
 * @param {number} mes - Mes (opcional, para análisis específico)
 * @param {number} anio - Año (opcional, para análisis específico)
 * @returns {Object} Diagnóstico completo del sistema
 */
function diagnosticoCompleto(nombrePersonal, mes = null, anio = null) {
  try {
    Logger.log("� DIAGNÓSTICO COMPLETO DEL SISTEMA");
    Logger.log(`Analizando: "${nombrePersonal}"`);
    if (mes && anio) {
      Logger.log(`Período específico: ${mes}/${anio}`);
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaRegistros = ss.getSheetByName("RegistrosProcedimientos");
    const hojaPersonal = ss.getSheetByName("Personal");
    
    if (!hojaRegistros || !hojaPersonal) {
      return { error: "Faltan hojas necesarias (RegistrosProcedimientos o Personal)" };
    }
    
    const datos = hojaRegistros.getDataRange().getValues();
    const datosPersonal = hojaPersonal.getDataRange().getValues();
    
    // Normalización de nombres
    const nombreNormalizado = normalizarNombre(nombrePersonal);
    
    // 1. ANÁLISIS DE ESTRUCTURA
    const estructura = {
      totalFilasRegistros: datos.length - 1,
      totalFilasPersonal: datosPersonal.length - 1,
      encabezadosRegistros: datos[0],
      encabezadosPersonal: datosPersonal[0]
    };
    
    // 2. BÚSQUEDA EN HOJA PERSONAL
    const resultadoPersonal = {
      encontrado: false,
      columna: -1,
      fila: -1,
      nombreExacto: "",
      tipo: ""
    };
    
    for (let i = 1; i < datosPersonal.length; i++) {
      for (let j = 0; j < 3 && j < datosPersonal[i].length; j++) {
        const nombrePersonalHoja = normalizarNombre(datosPersonal[i][j]);
        if (nombrePersonalHoja === nombreNormalizado || 
            nombrePersonalHoja.includes(nombreNormalizado) || 
            nombreNormalizado.includes(nombrePersonalHoja)) {
          resultadoPersonal.encontrado = true;
          resultadoPersonal.columna = j;
          resultadoPersonal.fila = i;
          resultadoPersonal.nombreExacto = String(datosPersonal[i][j] || "").trim();
          resultadoPersonal.tipo = obtenerTipoDePersonal(nombrePersonal, hojaPersonal);
          break;
        }
      }
      if (resultadoPersonal.encontrado) break;
    }
    
    // 3. ANÁLISIS DE REGISTROS
    const analisisRegistros = {
      coincidenciasNombre: [],
      registrosPorPeriodo: [],
      combinacionNombreFecha: 0,
      ejemplosDatos: [],
      fechaInicio: null,
      fechaFin: null,
      procedimientosUnicos: new Set(),
      totalRegistrosPersonal: 0
    };
    
    // Buscar en registros
    for (let i = 1; i < datos.length; i++) {
      const nombreRegistro = normalizarNombre(datos[i][1]);
      const fechaRegistro = datos[i][0];
      
      // Verificar coincidencia de nombre
      if (nombreRegistro === nombreNormalizado || 
          nombreRegistro.includes(nombreNormalizado) || 
          nombreNormalizado.includes(nombreRegistro)) {
        
        analisisRegistros.totalRegistrosPersonal++;
        analisisRegistros.coincidenciasNombre.push({
          fila: i + 1,
          nombre: String(datos[i][1] || "").trim(),
          fecha: fechaRegistro instanceof Date ? fechaRegistro.toLocaleDateString() : String(fechaRegistro)
        });
        
        // Análisis de fechas
        if (fechaRegistro) {
          const fecha = fechaRegistro instanceof Date ? fechaRegistro : new Date(fechaRegistro);
          if (!isNaN(fecha.getTime())) {
            if (!analisisRegistros.fechaInicio || fecha < analisisRegistros.fechaInicio) {
              analisisRegistros.fechaInicio = fecha;
            }
            if (!analisisRegistros.fechaFin || fecha > analisisRegistros.fechaFin) {
              analisisRegistros.fechaFin = fecha;
            }
            
            // Si se especificó período, verificar coincidencia
            if (mes && anio) {
              const mesRegistro = fecha.getMonth() + 1;
              const anioRegistro = fecha.getFullYear();
              
              if (mesRegistro === Number(mes) && anioRegistro === Number(anio)) {
                analisisRegistros.combinacionNombreFecha++;
                analisisRegistros.registrosPorPeriodo.push({
                  fila: i + 1,
                  fecha: fecha.toLocaleDateString(),
                  nombre: String(datos[i][1] || "").trim()
                });
              }
            }
          }
        }
        
        // Análizar procedimientos
        for (let col = 2; col < datos[i].length; col++) {
          if (datos[i][col] && datos[i][col] > 0) {
            const procedimiento = datos[0][col];
            analisisRegistros.procedimientosUnicos.add(procedimiento);
            
            if (analisisRegistros.ejemplosDatos.length < 5) {
              analisisRegistros.ejemplosDatos.push({
                fila: i + 1,
                fecha: fechaRegistro instanceof Date ? fechaRegistro.toLocaleDateString() : String(fechaRegistro),
                procedimiento: procedimiento,
                cantidad: datos[i][col]
              });
            }
          }
        }
      }
    }
    
    // 4. MUESTREO GENERAL DE DATOS
    const muestreoGeneral = {
      primeros10Nombres: [],
      fechasDelPeriodo: 0,
      ejemplosFechasPeriodo: []
    };
    
    for (let i = 1; i <= Math.min(10, datos.length - 1); i++) {
      const nombre = String(datos[i][1] || "").trim();
      if (nombre) {
        muestreoGeneral.primeros10Nombres.push({
          fila: i + 1,
          nombre: nombre,
          normalizado: normalizarNombre(nombre)
        });
      }
    }
    
    if (mes && anio) {
      for (let i = 1; i < datos.length; i++) {
        const fechaCelda = datos[i][0];
        if (fechaCelda) {
          const fecha = fechaCelda instanceof Date ? fechaCelda : new Date(fechaCelda);
          if (!isNaN(fecha.getTime())) {
            const mesRegistro = fecha.getMonth() + 1;
            const anioRegistro = fecha.getFullYear();
            
            if (mesRegistro === Number(mes) && anioRegistro === Number(anio)) {
              muestreoGeneral.fechasDelPeriodo++;
              if (muestreoGeneral.ejemplosFechasPeriodo.length < 5) {
                muestreoGeneral.ejemplosFechasPeriodo.push({
                  fila: i + 1,
                  fecha: fecha.toLocaleDateString(),
                  nombre: String(datos[i][1] || "").trim()
                });
              }
            }
          }
        }
      }
    }
    
    // RESULTADO CONSOLIDADO
    const resultado = {
      timestamp: new Date().toLocaleString(),
      nombreBuscado: nombrePersonal,
      nombreNormalizado: nombreNormalizado,
      periodoAnalizado: mes && anio ? `${mes}/${anio}` : "Todos los registros",
      
      estructura: estructura,
      personalEnHojaPersonal: resultadoPersonal,
      analisisRegistros: {
        ...analisisRegistros,
        fechaInicio: analisisRegistros.fechaInicio ? analisisRegistros.fechaInicio.toLocaleDateString() : null,
        fechaFin: analisisRegistros.fechaFin ? analisisRegistros.fechaFin.toLocaleDateString() : null,
        procedimientosUnicos: Array.from(analisisRegistros.procedimientosUnicos)
      },
      muestreoGeneral: muestreoGeneral,
      
      // DIAGNÓSTICO FINAL
      problemas: [],
      recomendaciones: []
    };
    
    // Identificar problemas
    if (!resultadoPersonal.encontrado) {
      resultado.problemas.push("Personal no encontrado en hoja Personal");
      resultado.recomendaciones.push("Verificar ortografía del nombre o agregarlo a la hoja Personal");
    }
    
    if (analisisRegistros.totalRegistrosPersonal === 0) {
      resultado.problemas.push("No hay registros para este personal");
      resultado.recomendaciones.push("Verificar que el personal tenga procedimientos registrados");
    }
    
    if (mes && anio && analisisRegistros.combinacionNombreFecha === 0) {
      resultado.problemas.push(`No hay registros para ${mes}/${anio}`);
      resultado.recomendaciones.push("Verificar el período de búsqueda o que existan registros en ese mes");
    }
    
    if (muestreoGeneral.fechasDelPeriodo === 0 && mes && anio) {
      resultado.problemas.push(`No hay registros generales para ${mes}/${anio}`);
      resultado.recomendaciones.push("El período especificado no tiene registros en todo el sistema");
    }
    
    Logger.log("📊 DIAGNÓSTICO COMPLETADO");
    Logger.log(`Problemas identificados: ${resultado.problemas.length}`);
    Logger.log(`Registros del personal: ${analisisRegistros.totalRegistrosPersonal}`);
    
    return resultado;
    
  } catch (error) {
    Logger.log("❌ Error en diagnóstico completo: " + error.message);
    return {
      error: error.message,
      nombreBuscado: nombrePersonal,
      timestamp: new Date().toLocaleString()
    };
  }
}





/**
 * Función de prueba para diagnosticar problemas de búsqueda
 * Ejecutar manualmente en Apps Script para debuggear
 */
function testearBusquedaProblemas() {
  Logger.log("🧪 INICIANDO PRUEBA DE DIAGNÓSTICO");
  
  // Cambiar estos valores por un caso real problemático
  const nombrePrueba = "NOMBRE_EJEMPLO"; // Cambiar por un nombre que esté dando problemas
  const mesPrueba = 12; // Cambiar por el mes que estás probando
  const anioPrueba = 2024; // Cambiar por el año que estás probando
  
  Logger.log(`Probando con: "${nombrePrueba}", ${mesPrueba}/${anioPrueba}`);
  
  try {
    // 1. Ejecutar diagnóstico completo
    const diagnostico = diagnosticoCompleto(nombrePrueba, mesPrueba, anioPrueba);
    
    Logger.log("\n🔍 RESULTADOS DEL DIAGNÓSTICO:");
    Logger.log("==========================================");
    
    // 2. Mostrar información de estructura
    Logger.log(`📊 ESTRUCTURA DE DATOS:`);
    Logger.log(`   - Filas en RegistrosProcedimientos: ${diagnostico.estructura.totalFilasRegistros}`);
    Logger.log(`   - Filas en Personal: ${diagnostico.estructura.totalFilasPersonal}`);
    
    // 3. Mostrar si el personal está en la hoja Personal
    Logger.log(`\n👤 PERSONAL EN HOJA PERSONAL:`);
    if (diagnostico.personalEnHojaPersonal.encontrado) {
      Logger.log(`   ✅ Encontrado: "${diagnostico.personalEnHojaPersonal.nombreExacto}"`);
      Logger.log(`   📍 Columna: ${diagnostico.personalEnHojaPersonal.columna}, Fila: ${diagnostico.personalEnHojaPersonal.fila}`);
      Logger.log(`   🏥 Tipo: ${diagnostico.personalEnHojaPersonal.tipo}`);
    } else {
      Logger.log(`   ❌ NO encontrado en hoja Personal`);
    }
    
    // 4. Mostrar análisis de registros
    Logger.log(`\n📋 ANÁLISIS DE REGISTROS:`);
    Logger.log(`   - Total registros del personal: ${diagnostico.analisisRegistros.totalRegistrosPersonal}`);
    Logger.log(`   - Coincidencias de nombre: ${diagnostico.analisisRegistros.coincidenciasNombre.length}`);
    Logger.log(`   - Período específico: ${diagnostico.analisisRegistros.combinacionNombreFecha} registros`);
    
    if (diagnostico.analisisRegistros.coincidenciasNombre.length > 0) {
      Logger.log(`\n📝 PRIMERAS COINCIDENCIAS DE NOMBRE:`);
      diagnostico.analisisRegistros.coincidenciasNombre.slice(0, 5).forEach(c => {
        Logger.log(`   Fila ${c.fila}: "${c.nombre}" (${c.fecha})`);
      });
    }
    
    // 5. Mostrar muestra general de datos
    Logger.log(`\n📈 MUESTRA GENERAL:`);
    Logger.log(`   - Total registros en período: ${diagnostico.muestreoGeneral.fechasDelPeriodo}`);
    
    if (diagnostico.muestreoGeneral.primeros10Nombres.length > 0) {
      Logger.log(`\n👥 PRIMEROS 10 NOMBRES EN LA HOJA:`);
      diagnostico.muestreoGeneral.primeros10Nombres.forEach(n => {
        Logger.log(`   Fila ${n.fila}: "${n.nombre}" → "${n.normalizado}"`);
      });
    }
    
    // 6. Mostrar problemas identificados
    if (diagnostico.problemas.length > 0) {
      Logger.log(`\n❌ PROBLEMAS IDENTIFICADOS:`);
      diagnostico.problemas.forEach(problema => {
        Logger.log(`   - ${problema}`);
      });
    }
    
    // 7. Mostrar recomendaciones
    if (diagnostico.recomendaciones.length > 0) {
      Logger.log(`\n💡 RECOMENDACIONES:`);
      diagnostico.recomendaciones.forEach(recomendacion => {
        Logger.log(`   - ${recomendacion}`);
      });
    }
    
    // 8. Probar también calcularCostos
    Logger.log(`\n🧮 PROBANDO CALCULAR COSTOS...`);
    const resultado = calcularCostos(nombrePrueba, mesPrueba, anioPrueba, "1-15");
    
    Logger.log(`💰 RESULTADO CALCULAR COSTOS:`);
    Logger.log(`   - Subtotal: $${resultado.totales.subtotal}`);
    Logger.log(`   - IVA: $${resultado.totales.iva}`);
    Logger.log(`   - Total: $${resultado.totales.total_con_iva}`);
    
    return {
      diagnostico: diagnostico,
      calculoCostos: resultado
    };
    
  } catch (error) {
    Logger.log(`❌ ERROR EN PRUEBA: ${error.message}`);
    Logger.log(`Stack: ${error.stack}`);
    return { error: error.message };
  }
}

/**
 * Función de prueba rápida para ver todos los nombres únicos en registros
 */
function verNombresEnRegistros() {
  Logger.log("📋 OBTENIENDO TODOS LOS NOMBRES EN REGISTROS...");
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("RegistrosProcedimientos");
    
    if (!hoja) {
      Logger.log("❌ No se encontró la hoja RegistrosProcedimientos");
      return;
    }
    
    const datos = hoja.getDataRange().getValues();
    const nombresUnicos = new Set();
    
    // Revisar columna B (índice 1) desde fila 2
    for (let i = 1; i < datos.length; i++) {
      const nombre = String(datos[i][1] || "").trim();
      if (nombre) {
        nombresUnicos.add(nombre);
      }
    }
    
    const lista = Array.from(nombresUnicos).sort();
    
    Logger.log(`📊 TOTAL DE NOMBRES ÚNICOS: ${lista.length}`);
    Logger.log(`📝 LISTA COMPLETA:`);
    
    lista.forEach((nombre, index) => {
      Logger.log(`   ${index + 1}. "${nombre}"`);
    });
    
    return lista;
    
  } catch (error) {
    Logger.log(`❌ ERROR: ${error.message}`);
    return [];
  }
}

function descargarReporteCostosPDF(htmlContent, reportTitle, fileName) {
  return generarPdfParaDescarga(
    htmlContent,
    "stylesReportePagoProcedimientos",
    reportTitle,
    fileName
  );
}

// --- Funciones para el Módulo de Ultrasonido ---
function mostrarFormularioUltrasonido() {
  const html = HtmlService.createTemplateFromFile("formularioUltrasonido")
    .evaluate()
    .setWidth(900)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, "Registro de Ultrasonidos");
}

function guardarRegistroUltrasonido(data) {
  if (!data || typeof data !== "object") {
    throw new Error(
      "❌ No se recibieron datos válidos para el registro de ultrasonido."
    );
  }

  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    "RegistroUltrasonidos"
  );
  if (!hoja) {
    throw new Error("❌ La hoja 'RegistroUltrasonidos' no existe.");
  }

  let [year, month, day] = data.fecha.split("-").map(Number);
  let fechaRegistro = new Date(Date.UTC(year, month - 1, day, 12, 0, 0));

  const fila = [
    fechaRegistro,
    data.radiologo || "",
    data.abdomen_completo || 0,
    data.abdomen_superior || 0,
    data.mama || 0,
    data.testiculos || 0,
    data.tracto_urinario || 0,
    data.tejidos_blandos || 0,
    data.tiroides || 0,
    data.articular || 0,
    data.pelvico || 0,
    data.doppler || 0,
  ];

  hoja.appendRow(fila);
  Logger.log("✅ Registro de ultrasonido guardado con éxito.");
}

function mostrarReportePagoUltrasonido() {
  const html = HtmlService.createTemplateFromFile("reportePagoUltrasonido")
    .evaluate()
    .setWidth(900)
    .setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(
    html,
    "Reporte de Pagos de Radiología"
  );
}

function obtenerRadiologos() {
  const hojaPersonal =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Personal");
  if (!hojaPersonal) {
    throw new Error("❌ La hoja 'Personal' no existe para obtener radiólogos.");
  }
  const datos = hojaPersonal.getDataRange().getValues();
  const radiologos = new Set();
  const columnaRadiologosIndex = 3; // Columna D

  for (let i = 1; i < datos.length; i++) {
    const nombre = datos[i][columnaRadiologosIndex];
    if (nombre && typeof nombre === "string" && nombre.trim()) {
      radiologos.add(nombre.trim());
    }
  }
  Logger.log(
    "Nombres de radiólogos obtenidos: " + [...radiologos].sort().join(", ")
  );
  return [...radiologos].sort();
}

function calcularPagosRadiologia(radiologo, mes, anio, quincena) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaRegistros = ss.getSheetByName("RegistroUltrasonidos");
  const hojaPrecios = ss.getSheetByName("Precios");

  if (!hojaRegistros || !hojaPrecios) {
    throw new Error(
      "Faltan hojas necesarias: 'RegistroUltrasonidos' o 'Precios'."
    );
  }

  const datosRegistros = hojaRegistros.getDataRange().getValues();
  const preciosDatos = hojaPrecios.getDataRange().getValues();

  const preciosPorUltrasonido = {};
  Logger.log("--- Leyendo precios de la hoja 'Precios' ---");
  preciosDatos.slice(1).forEach((fila, index) => {
    const tipo = fila[6]; // Columna G
    const costo = fila[7]; // Columna H
    if (tipo && typeof costo === "number") {
      preciosPorUltrasonido[tipo] = costo;
      Logger.log(`Fila ${index + 2}: Tipo "${tipo}" -> Costo ${costo}`);
    } else {
      Logger.log(
        `Fila ${
          index + 2
        }: No se pudo leer el precio. Tipo: "${tipo}", Costo: "${costo}" (¿es un número?)`
      );
    }
  });
  Logger.log(
    "Precios cargados (objeto): " + JSON.stringify(preciosPorUltrasonido)
  );
  Logger.log("-----------------------------------------");

  const mapeoUltrasonidos = {
    abdomen_completo: "Abdomen Completo",
    abdomen_superior: "Abdomen Superior",
    mama: "Mama",
    testiculos: "Testículos",
    tracto_urinario: "Tracto Urinario",
    tejidos_blandos: "Tejidos Blandos",
    tiroides: "Tiroides",
    articular: "Articular",
    pelvico: "Pélvico",
    doppler: "Doppler",
  };

  const resumen = {};
  let totalGananciaSinIVA = 0;
  let totalIVA = 0;
  let totalGananciaConIVA = 0;
  const IVA_PERCENTAGE = 0.04;

  Object.keys(mapeoUltrasonidos).forEach((key) => {
    resumen[key] = {
      nombre: mapeoUltrasonidos[key],
      cantidad: 0,
      costo_unitario: preciosPorUltrasonido[mapeoUltrasonidos[key]] || 0,
      ganancia_sin_iva: 0,
      iva_monto: 0,
      ganancia_con_iva: 0,
    };
  });

  Logger.log("--- Procesando registros de ultrasonido ---");
  datosRegistros.slice(1).forEach((fila) => {
    const fecha = new Date(fila[0]);
    const filaRadiologo = fila[1];
    const filaMes = fecha.getMonth() + 1;
    const filaAnio = fecha.getFullYear();
    const filaDia = fecha.getDate();

    if (filaRadiologo !== radiologo || filaMes !== mes || filaAnio !== anio)
      return;
    if (quincena === "1-15" && filaDia > 15) return;
    if (quincena === "16-31" && filaDia <= 15) return;

    const sheetColumnOrder = [
      "Abdomen Completo",
      "Abdomen Superior",
      "Mama",
      "Testículos",
      "Tracto Urinario",
      "Tejidos Blandos",
      "Tiroides",
      "Articular",
      "Pélvico",
      "Doppler",
    ];

    let currentSheetColIndex = 2;
    for (const sheetColName of sheetColumnOrder) {
      const key = Object.keys(mapeoUltrasonidos).find(
        (k) => mapeoUltrasonidos[k] === sheetColName
      );

      if (key) {
        const cantidad = Number(fila[currentSheetColIndex]) || 0;
        if (cantidad > 0) {
          const tipoUltrasonido = mapeoUltrasonidos[key];
          let costoUnitario = preciosPorUltrasonido[tipoUltrasonido];

          if (typeof costoUnitario === "undefined" || costoUnitario === null) {
            Logger.log(
              `🚨 Advertencia: No se encontró costo para "${tipoUltrasonido}". Cantidad: ${cantidad}. Estableciendo costo a 0.`
            );
            costoUnitario = 0;
          } else {
            Logger.log(
              `Procesando: ${tipoUltrasonido}, Cantidad: ${cantidad}, Costo unitario: ${costoUnitario}`
            );
          }

          const gananciaSinIVA = cantidad * costoUnitario;
          const ivaMonto = gananciaSinIVA * IVA_PERCENTAGE;
          const gananciaConIVA = gananciaSinIVA + ivaMonto;

          resumen[key].cantidad += cantidad;
          resumen[key].ganancia_sin_iva += gananciaSinIVA;
          resumen[key].iva_monto += ivaMonto;
          resumen[key].ganancia_con_iva += gananciaConIVA;
          resumen[key].costo_unitario = costoUnitario;

          totalGananciaSinIVA += gananciaSinIVA;
          totalIVA += ivaMonto;
          totalGananciaConIVA += gananciaConIVA;
        }
      }
      currentSheetColIndex++;
    }
  });
  Logger.log("--- Fin del procesamiento de registros ---");
  return { resumen, totalGananciaSinIVA, totalIVA, totalGananciaConIVA };
}

// --- Funciones Wrapper para Enviar/Descargar Reportes de Ultrasonido ---
function enviarReporteUltrasonido(
  htmlContent,
  emailRecipient,
  reportTitle,
  subject,
  fileName
) {
  return enviarReportePDF(
    htmlContent,
    "stylesReportePagoUltrasonido",
    reportTitle,
    emailRecipient,
    subject,
    fileName
  );
}

function descargarReporteUltrasonidoPDF(htmlContent, reportTitle, fileName) {
  return generarPdfParaDescarga(
    htmlContent,
    "stylesReportePagoUltrasonido",
    reportTitle,
    fileName
  );
}

// --- Funciones para el Módulo de Horas Extras ---
function mostrarFormularioHorasExtras() {
  const html = HtmlService.createTemplateFromFile("formularioHorasExtras")
    .evaluate()
    .setWidth(400)
    .setHeight(450);
  SpreadsheetApp.getUi().showModalDialog(html, "Registro de Horas Extras");
}

/**
 * Inserta un registro en la hoja "Horas Extras" en la fila 2 (debajo del encabezado).
 * Los registros anteriores se desplazan hacia abajo.
 * @param {Object} data - Datos del registro de horas extras.
 * @returns {boolean} - Verdadero si se guardó correctamente.
 */
function guardarHorasExtrasArriba(data) {
  Logger.log(
    "guardarHorasExtrasArriba llamado con datos: " + JSON.stringify(data)
  );

  if (!data || typeof data !== "object") {
    Logger.log(
      "ERROR: No se recibieron datos válidos para el registro de horas extras."
    );
    throw new Error(
      "❌ No se recibieron datos válidos para el registro de horas extras."
    );
  }

  // Validar datos
  if (!data.fecha || !data.trabajador || !data.cantidad_horas) {
    Logger.log("ERROR: Datos incompletos para el registro de horas extras.");
    throw new Error("❌ Datos incompletos. Todos los campos son obligatorios.");
  }

  // Validar que la cantidad de horas sea un número positivo
  if (typeof data.cantidad_horas !== "number" || data.cantidad_horas <= 0) {
    Logger.log("ERROR: La cantidad de horas debe ser un número positivo.");
    throw new Error("❌ La cantidad de horas debe ser un número positivo.");
  }

  // Obtener la hoja de cálculo
  const hoja =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("HorasExtras");
  if (!hoja) {
    Logger.log("ERROR: La hoja 'Horas Extras' no existe.");
    throw new Error("❌ La hoja 'Horas Extras' no existe.");
  }

  try {
    // Convertir la fecha de string a objeto Date
    let [year, month, day] = data.fecha.split("-").map(Number);
    let fechaRegistro = new Date(Date.UTC(year, month - 1, day, 12, 0, 0));

    // Preparar la fila para añadir
    const fila = [fechaRegistro, data.trabajador, data.cantidad_horas];

    // Insertar una nueva fila en la posición 2 (debajo del encabezado)
    hoja.insertRowAfter(1);

    // Copiar formato de la fila 3 (antigua fila 2) a la nueva fila 2
    hoja
      .getRange(3, 1, 1, hoja.getLastColumn())
      .copyTo(hoja.getRange(2, 1, 1, hoja.getLastColumn()), {
        formatOnly: true,
      });

    // Añadir los datos en la nueva fila 2
    hoja.getRange(2, 1, 1, fila.length).setValues([fila]);

    // Formatear la celda de fecha
    hoja.getRange(2, 1).setNumberFormat("dd/MM/yyyy");

    // Formatear la celda de cantidad de horas
    hoja.getRange(2, 3).setNumberFormat("0.0");

    Logger.log("✅ Registro de horas extras insertado en la fila 2 con éxito.");
    return true;
  } catch (error) {
    Logger.log("ERROR al guardar horas extras: " + error.message);
    throw new Error("❌ Error al guardar el registro: " + error.message);
  }
}

function mostrarReportePagoHorasExtras() {
  const html = HtmlService.createTemplateFromFile("reportePagoHorasExtras")
    .evaluate()
    .setWidth(700)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, "Reporte de Horas Extras");
}

function calcularReporteHorasExtras(trabajador, mes, anio, quincena) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaHorasExtras = ss.getSheetByName("Horas Extras");

  if (!hojaHorasExtras) {
    throw new Error("Faltan hojas necesarias: 'HorasExtras'.");
  }

  const datosHorasExtras = hojaHorasExtras.getDataRange().getValues();

  const resumenHorasExtras = [];
  let totalHoras = 0;

  Logger.log("--- Procesando registros de Horas Extras ---");
  datosHorasExtras.slice(1).forEach((fila) => {
    const fecha = new Date(fila[0]);
    const filaTrabajador = fila[1];
    const cantidadHoras = Number(fila[2]) || 0;
    const filaMes = fecha.getMonth() + 1;
    const filaAnio = fecha.getFullYear();
    const filaDia = fecha.getDate();

    if (filaTrabajador !== trabajador || filaMes !== mes || filaAnio !== anio)
      return;
    if (quincena === "1-15" && filaDia > 15) return;
    if (quincena === "16-31" && filaDia <= 15) return;

    resumenHorasExtras.push({
      fecha: Utilities.formatDate(
        fecha,
        ss.getSpreadsheetTimeZone(),
        "dd/MM/yyyy"
      ),
      trabajador: filaTrabajador,
      cantidad_horas: cantidadHoras,
    });
    totalHoras += cantidadHoras;

    Logger.log(
      `Fecha: ${fecha}, Trabajador: ${filaTrabajador}, Horas: ${cantidadHoras}`
    );
  });
  Logger.log("--- Fin del procesamiento de registros de Horas Extras ---");

  return { resumen: resumenHorasExtras, totalHoras: totalHoras };
}

// --- Funciones para el Módulo de Biopsias ---

/**
 * Muestra el formulario para registrar biopsias.
 */
function mostrarRegistroBiopsias() {
  const html = HtmlService.createTemplateFromFile("registroBiopsias")
    .evaluate()
    .setWidth(800)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, "Registro de Biopsias");
}

/**
 * Guarda un registro de biopsia en la hoja "Biopsias".
 * @param {Object} data - Datos del registro de biopsia.
 * @returns {boolean} - Verdadero si se guardó correctamente.
 */
function guardarRegistroBiopsia(data) {
  try {
    Logger.log("Guardando registro de biopsia: " + JSON.stringify(data));

    if (!data || typeof data !== "object") {
      throw new Error(
        "No se recibieron datos válidos para el registro de biopsia."
      );
    }

    const hoja =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RegistroBiopsias");
    if (!hoja) {
      throw new Error("La hoja 'RegistroBiopsias' no existe.");
    }

    // Validar y obtener la fecha (acepta fecha_toma o fechaToma)
    const fechaStr = data.fechaToma || data.fecha_toma || data.fecha || "";
    if (!fechaStr || typeof fechaStr !== "string") {
      throw new Error(
        "El campo de fecha es obligatorio y debe tener formato YYYY-MM-DD."
      );
    }
    let fechaToma;
    try {
      let [year, month, day] = fechaStr.split("-").map(Number);
      fechaToma = new Date(Date.UTC(year, month - 1, day, 12, 0, 0));
    } catch (e) {
      throw new Error("Formato de fecha inválido: " + e.message);
    }

    // Preparar la fila para añadir
    const fila = [
      fechaToma, // Fecha Toma
      false, // Recibida (por defecto)
      false, // Enviada (por defecto)
      data.cedula || "",
      data.telefono || "",
      data.nombre_cliente || "",
      Number(data.frascos_gastro) || 0,
      Number(data.frascos_colon) || 0,
      data.medico || "",
      data.comentario || "",
    ];

    // Insertar una nueva fila en la posición 2 (debajo del encabezado)
    hoja.insertRowAfter(1);

    // Copiar formato de la fila 3 (antigua fila 2) a la nueva fila 2
    hoja
      .getRange(3, 1, 1, hoja.getLastColumn())
      .copyTo(hoja.getRange(2, 1, 1, hoja.getLastColumn()), {
        formatOnly: true,
      });

    hoja.getRange(2, 1, 1, fila.length).setValues([fila]);

    // Formatear la celda de fecha
    hoja.getRange(2, 1).setNumberFormat("yyyy-mm-dd");

    // Convertir las celdas de recibida y enviada en checkboxes
    hoja.getRange(2, 2, 1, 2).insertCheckboxes();

    Logger.log("✅ Registro de biopsia insertado en la fila 2 con éxito.");
    return true;
  } catch (error) {
    Logger.log("ERROR al guardar biopsia: " + error.message);
    throw new Error("Error al guardar el registro: " + error.message);
  }
}

/**
 * Muestra la interfaz de búsqueda y reporte de biopsias.
 */
function mostrarBuscarBiopsias() {
  const html = HtmlService.createTemplateFromFile("buscarBiopsiasReg")
    .evaluate()
    .setWidth(900)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(
    html,
    "Búsqueda y Reporte de Biopsias"
  );
}

/**
 * Busca biopsias según el tipo y valor de búsqueda.
 * @param {string} searchType - Tipo de búsqueda: "fecha", "cedula" o "nombre".
 * @param {string} searchValue - Valor a buscar.
 * @returns {Array} - Resultados de la búsqueda.
 */
function buscarBiopsiasServidorV2(searchType, searchValue) {
  Logger.log("typeof searchType: " + typeof searchType);
  Logger.log("typeof searchValue: " + typeof searchValue);
  Logger.log("searchType: " + searchType);
  Logger.log("searchValue: " + searchValue);
  Logger.log("DEBUG: searchType recibido: " + searchType);
  Logger.log("DEBUG: searchValue recibido: " + searchValue);
  Logger.log("INICIO buscarBiopsiasServidor");
  Logger.log(
    "Parámetros recibidos: searchType=%s, searchValue=%s",
    searchType,
    searchValue
  );

  if (!searchType || !searchValue) {
    Logger.log("ERROR: searchType o searchValue no definidos");
    return [];
  }

  try {
    Logger.log(`Iniciando búsqueda: tipo=${searchType}, valor=${searchValue}`);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("RegistroBiopsias");

    if (!hoja) {
      Logger.log("ERROR: La hoja 'Biopsias' no existe");
      return [];
    }

    const datos = hoja.getDataRange().getValues();
    Logger.log("Primeras filas de datos: " + JSON.stringify(datos.slice(0, 3)));

    // Normalización para búsquedas robustas
    const normalizar = (str) =>
      String(str || "")
        .toLowerCase()
        .replace(/\s+/g, " ")
        .trim();

    const datosFiltrados = [];

    // Justo antes del for principal en buscarBiopsiasServidor
    Logger.log("Listado de cédulas en la hoja:");
    for (let i = 1; i < datos.length; i++) {
      Logger.log("Fila " + (i + 1) + ": " + JSON.stringify(datos[i][3]));
    }

    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      let coincide = false;

      if (searchType === "fecha") {
        let fechaHojaStr = "";
        if (fila[0] instanceof Date) {
          fechaHojaStr = Utilities.formatDate(
            fila[0],
            Session.getScriptTimeZone(),
            "yyyy-MM-dd"
          );
        } else if (typeof fila[0] === "string") {
          let fechaTmp = new Date(fila[0]);
          fechaHojaStr = Utilities.formatDate(
            fechaTmp,
            Session.getScriptTimeZone(),
            "yyyy-MM-dd"
          );
        }
        let fechaBusquedaStr = searchValue;
        if (searchValue instanceof Date) {
          fechaBusquedaStr = Utilities.formatDate(
            searchValue,
            Session.getScriptTimeZone(),
            "yyyy-MM-dd"
          );
        }
        Logger.log(
          `Comparando: hoja=${fechaHojaStr} vs busqueda=${fechaBusquedaStr}`
        );
        if (fechaHojaStr === fechaBusquedaStr) {
          coincide = true;
        }
      } else if (searchType === "cedula") {
        // Normaliza: convierte a string, elimina espacios y caracteres no numéricos
        const normalizarCedula = (str) => String(str || "").replace(/\D/g, "");
        const cedulaHoja = normalizarCedula(fila[3]);
        const cedulaBusqueda = normalizarCedula(searchValue);
        Logger.log(
          `Comparando cédula: hoja=${cedulaHoja} vs busqueda=${cedulaBusqueda}`
        );
        Logger.log(
          `Fila ${
            i + 1
          }: cedulaHoja=${cedulaHoja}, cedulaBusqueda=${cedulaBusqueda}`
        );
        if (cedulaHoja === cedulaBusqueda) {
          coincide = true;
        }
      } else if (searchType === "nombre") {
        const nombreHoja = normalizar(fila[5]);
        const nombreBusqueda = normalizar(searchValue);
        if (nombreHoja.includes(nombreBusqueda)) {
          coincide = true;
        }
      }

      if (coincide) {
        const filaConIndice = [...fila, i + 1];
        datosFiltrados.push(filaConIndice);
      }
    }

    Logger.log(
      `FIN buscarBiopsiasServidor. Resultados encontrados: ${JSON.stringify(
        datosFiltrados
      )}`
    );
    return datosFiltrados || [];
  } catch (error) {
    Logger.log(`ERROR buscarBiopsiasServidor: ${error.message}`);
    return [];
  }
}

/**
 * Actualiza el estado de una biopsia (Recibida o Enviada).
 * @param {number} rowIndex - Índice de la fila a actualizar.
 * @param {string} column - Columna a actualizar: "recibida" o "enviada".
 * @param {boolean} value - Nuevo valor para la celda.
 */
function actualizarEstadoBiopsia(rowIndex, column, value) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("RegistroBiopsias");

    if (!sheet) {
      throw new Error("La hoja 'RegistroBiopsias' no existe.");
    }

    // Determinar la columna a actualizar
    let columnIndex;
    if (column === "recibida") {
      columnIndex = 2; // Columna B (índice 1 + 1)
    } else if (column === "enviada") {
      columnIndex = 3; // Columna C (índice 2 + 1)
    } else {
      throw new Error("Columna no válida para actualizar.");
    }

    // Actualizar el valor
    sheet.getRange(rowIndex, columnIndex).setValue(value);

    // Si se actualizó la columna "enviada", aplicar formato condicional
    if (column === "enviada") {
      aplicarFormatoCondicionalFilaCompleta(rowIndex);
    }

    return true;
  } catch (error) {
    Logger.log(`ERROR en actualizarEstadoBiopsia: ${error.message}`);
    throw new Error(`Error al actualizar estado: ${error.message}`);
  }
}

/**
 * Aplica formato condicional a toda la fila cuando la columna C está marcada.
 * @param {number} rowIndex - Índice de la fila a formatear.
 */
function aplicarFormatoCondicionalFilaCompleta(rowIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Biopsias");

  if (!sheet) {
    throw new Error("La hoja 'RegistroBiopsias' no existe.");
  }

  // Obtener el valor de la celda C (Enviada al cliente)
  const cellCValue = sheet.getRange(rowIndex, 3).getValue();

  if (cellCValue === true) {
    // Si la celda C está marcada (TRUE para checkbox)
    // Aplicar fondo amarillo a toda la fila
    sheet
      .getRange(rowIndex, 1, 1, sheet.getLastColumn())
      .setBackground("#FFFF00"); // Amarillo
  } else {
    // Quitar el fondo si no está marcada (restaurar al color por defecto o Blanco Roto)
    sheet
      .getRange(rowIndex, 1, 1, sheet.getLastColumn())
      .setBackground("#EFEEEA"); // Blanco roto
  }
}

/**
 * Inicializa los checkbox en las columnas B y C.
 * Se debe llamar esta función una vez, manualmente o desde un trigger,
 * para convertir las celdas existentes en checkboxes.
 */
function inicializarCheckboxes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("RegistroBiopsias");

  if (!sheet) {
    throw new Error("La hoja 'Biopsias' no existe.");
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return; // No hay datos para procesar (asumiendo encabezados en fila 1)

  // Columna B (Recibida por el centro) - índice 1 + 1
  sheet.getRange(2, 2, lastRow - 1, 1).insertCheckboxes();

  // Columna C (Enviada al cliente) - índice 2 + 1
  sheet.getRange(2, 3, lastRow - 1, 1).insertCheckboxes();

  // Aplicar formato condicional inicial para filas que ya tienen la columna C marcada
  for (let i = 2; i <= lastRow; i++) {
    aplicarFormatoCondicionalFilaCompleta(i);
  }
}

/**
 * Obtiene biopsias filtradas por mes y año.
 * @param {number} mes - Mes (1-12).
 * @param {number} anio - Año (ej. 2023).
 * @returns {Array} - Resultados filtrados.
 */
function obtenerBiopsiasPorMesAnio(mes, anio) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("RegistroBiopsias");

    if (!hoja) {
      throw new Error("La hoja 'RegistroBiopsias' no existe.");
    }

    const datos = hoja.getDataRange().getValues();
    const datosFiltrados = [];

    // Si no hay datos (solo encabezado), retornar array vacío
    if (datos.length <= 1) {
      return [];
    }

    // Iterar desde la segunda fila (índice 1) para ignorar encabezados
    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      const fechaCelda = fila[0];

      if (fechaCelda instanceof Date) {
        const fechaMes = fechaCelda.getMonth() + 1; // getMonth() devuelve 0-11
        const fechaAnio = fechaCelda.getFullYear();

        if (fechaMes === mes && fechaAnio === anio) {
          datosFiltrados.push(fila);
        }
      }
    }

    return datosFiltrados;
  } catch (error) {
    Logger.log(`ERROR en obtenerBiopsiasPorMesAnio: ${error.message}`);
    throw new Error(
      `Error al obtener biopsias por mes y año: ${error.message}`
    );
  }
}

/**
 * Genera un PDF con el reporte de biopsias.
 * @param {string} htmlContent - Contenido HTML del reporte.
 * @param {string} reportTitle - Título del reporte.
 * @returns {string} - URL del PDF generado.
 */
function generarPdfBiopsias(htmlContent, reportTitle) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const spreadsheetId = spreadsheet.getId();
    const spreadsheetName = spreadsheet.getName();

    // Crear una hoja temporal para el PDF
    const tempSheet = spreadsheet.insertSheet(`Temp_${new Date().getTime()}`);

    // Configurar la hoja temporal con el contenido HTML
    tempSheet.getRange(1, 1).setValue(reportTitle);
    tempSheet
      .getRange(2, 1)
      .setValue('=IMPORTHTML("' + htmlContent + '", "table", 1)');

    // Configurar opciones de exportación
    const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=pdf&gid=${tempSheet.getSheetId()}&portrait=true&fitw=true&gridlines=false&printtitle=false&sheetnames=false&pagenum=false&attachment=false`;

    const options = {
      headers: {
        Authorization: "Bearer " + ScriptApp.getOAuthToken(),
      },
    };

    // Obtener el PDF
    const response = UrlFetchApp.fetch(url, options);
    const pdf = response.getBlob().setName("Reporte_Biopsias.pdf");

    // Guardar en Drive y obtener URL
    const folder = DriveApp.createFolder("Reportes_Biopsias");
    const file = folder.createFile(pdf);
    const fileUrl = file.getUrl();

    // Eliminar la hoja temporal
    spreadsheet.deleteSheet(tempSheet);

    return fileUrl;
  } catch (error) {
    Logger.log(`ERROR en generarPdfBiopsias: ${error.message}`);
    throw new Error(`Error al generar PDF: ${error.message}`);
  }
}

// Función para ejecutar una vez e inicializar los checkboxes
// Puedes ejecutarla manualmente desde el editor de Apps Script (Run > inicializarCheckboxes)
// o agregarla a un disparador para que se ejecute al abrir la hoja si lo prefieres.
// ScriptApp.newTrigger('inicializarCheckboxes')
//     .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
//     .onOpen()
//     .create();

// ================== FUNCIONES DE DEBUGGING ==================

/**
 * Función para verificar datos específicos de un personal
 * @param {string} nombrePersonal - Nombre del personal a verificar
 * @returns {Object} Información detallada sobre el personal
 */

/**
 * Función para obtener todo el personal disponible en RegistrosProcedimientos
 * @returns {Array} Lista completa de personal único
 */
function obtenerPersonalCompleto() {
  try {
    Logger.log("🔍 Obteniendo personal completo de RegistrosProcedimientos");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaRegistros = ss.getSheetByName("RegistrosProcedimientos");
    
    if (!hojaRegistros) {
      throw new Error("La hoja 'RegistrosProcedimientos' no existe");
    }
    
    const datos = hojaRegistros.getDataRange().getValues();
    const personalSet = new Set();
    
    // Revisar columna B (Personal) desde la fila 2
    for (let i = 1; i < datos.length; i++) {
      const personal = datos[i][1]; // Columna B
      if (personal && typeof personal === 'string' && personal.trim()) {
        personalSet.add(personal.trim());
      }
    }
    
    const resultado = Array.from(personalSet).sort();
    Logger.log("✅ Personal encontrado: " + resultado.length + " personas");
    Logger.log("Lista: " + resultado.join(", "));
    
    return resultado;
    
  } catch (error) {
    Logger.log("❌ Error en obtenerPersonalCompleto: " + error.message);
    return [];
  }
}

/**
 * Función para verificar la estructura de la hoja RegistrosProcedimientos
 * @returns {Object} Información sobre la estructura
 */
function verificarEstructuraHoja() {
  try {
    Logger.log("🔍 Verificando estructura de RegistrosProcedimientos");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("RegistrosProcedimientos");
    
    if (!hoja) {
      throw new Error("La hoja 'RegistrosProcedimientos' no existe");
    }
    
    const rango = hoja.getDataRange();
    const datos = rango.getValues();
    const encabezados = datos[0];
    
    const resultado = {
      nombreHoja: hoja.getName(),
      totalFilas: rango.getNumRows(),
      totalColumnas: rango.getNumColumns(),
      encabezados: encabezados,
      filasConDatos: datos.length - 1 // Sin contar encabezados
    };
    
    Logger.log("✅ Estructura verificada: " + JSON.stringify(resultado));
    return resultado;
    
  } catch (error) {
    Logger.log("❌ Error en verificarEstructuraHoja: " + error.message);
    return {
      error: error.message
    };
  }
}

/**
 * Función mejorada para buscar personal por nombre con múltiples estrategias
 * @param {string} nombreBuscar - Nombre a buscar
 * @returns {Object} Resultados de búsqueda con coincidencias exactas y parciales
 */
function buscarPersonalPorNombre(nombreBuscar) {
  try {
    Logger.log("🔍 Iniciando búsqueda de personal: " + nombreBuscar);
    
    if (!nombreBuscar || nombreBuscar.trim() === '') {
      return {
        coincidenciasExactas: [],
        coincidenciasParciales: [],
        sugerencias: []
      };
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaRegistros = ss.getSheetByName("RegistrosProcedimientos");
    const hojaPersonal = ss.getSheetByName("Personal");
    
    if (!hojaRegistros || !hojaPersonal) {
      throw new Error("No se encontraron las hojas necesarias");
    }
    
    function normalizarNombre(str) {
      return String(str || "").replace(/\s+/g, " ").trim().toUpperCase();
    }
    
    const nombreBuscarNorm = normalizarNombre(nombreBuscar);
    
    // Obtener todos los nombres de ambas hojas
    const nombresEncontrados = new Set();
    
    // Buscar en hoja Personal (columnas A, B, C, D)
    const datosPersonal = hojaPersonal.getDataRange().getValues();
    for (let i = 1; i < datosPersonal.length; i++) {
      const fila = datosPersonal[i];
      for (let j = 0; j < 4 && j < fila.length; j++) {
        const nombre = String(fila[j] || "").trim();
        if (nombre && nombre.length > 0) {
          nombresEncontrados.add(nombre);
        }
      }
    }
    
    // Buscar en hoja RegistrosProcedimientos (columna B)
    const datosRegistros = hojaRegistros.getDataRange().getValues();
    for (let i = 1; i < datosRegistros.length; i++) {
      const nombre = String(datosRegistros[i][1] || "").trim();
      if (nombre && nombre.length > 0) {
        nombresEncontrados.add(nombre);
      }
    }
    
    const todosLosNombres = Array.from(nombresEncontrados);
    
    // Categorizar coincidencias
    const coincidenciasExactas = [];
    const coincidenciasParciales = [];
    const sugerencias = [];
    
    todosLosNombres.forEach(nombre => {
      const nombreNorm = normalizarNombre(nombre);
      
      // Coincidencia exacta
      if (nombreNorm === nombreBuscarNorm) {
        coincidenciasExactas.push(nombre);
        return;
      }
      
      // Coincidencia parcial (uno contiene al otro)
      if (nombreNorm.includes(nombreBuscarNorm) || nombreBuscarNorm.includes(nombreNorm)) {
        coincidenciasParciales.push(nombre);
        return;
      }
      
      // Sugerencias (palabras similares)
      const palabrasBusqueda = nombreBuscarNorm.split(' ').filter(p => p.length > 2);
      const palabrasNombre = nombreNorm.split(' ').filter(p => p.length > 2);
      
      if (palabrasBusqueda.length > 0 && palabrasNombre.length > 0) {
        const coincidencias = palabrasBusqueda.filter(p1 => 
          palabrasNombre.some(p2 => p1.includes(p2) || p2.includes(p1))
        );
        
        if (coincidencias.length >= Math.min(palabrasBusqueda.length, palabrasNombre.length) * 0.4) {
          sugerencias.push(nombre);
        }
      }
    });
    
    const resultado = {
      coincidenciasExactas: coincidenciasExactas.slice(0, 10),
      coincidenciasParciales: coincidenciasParciales.slice(0, 10),
      sugerencias: sugerencias.slice(0, 5)
    };
    
    Logger.log("Búsqueda completada: " + JSON.stringify(resultado));
    return resultado;
    
  } catch (error) {
    Logger.log("Error en buscarPersonalPorNombre: " + error.message);
    throw new Error("Error al buscar personal: " + error.message);
  }
}
