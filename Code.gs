// ================= MEJORAS IMPLEMENTADAS =================
// Configuración centralizada de columnas para evitar confusiones
const COLUMNAS_PERSONAL = {
  DOCTOR: 0, // Columna A
  ANESTESIOLOGO: 1, // Columna B (CONFIRMADO)
  TECNICO: 2, // Columna C
  RADIOLOGO: 3, // Columna D
  ENFERMERO: 4, // Columna E
  SECRETARIA: 5, // Columna F
  EMAIL: 6, // Columna G
};

const TIPOS_PERSONAL = {
  0: "Doctor",
  1: "Anestesiólogo",
  2: "Técnico",
  3: "Radiólogo",
  4: "Enfermero",
  5: "Secretaria",
};

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
  return String(str || "")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
}

/**
 * FUNCIÓN DE PRUEBA: Verificar acceso a hoja Personal
 * Esta función ayuda a diagnosticar problemas de carga de personal
 */
function probarAccesoPersonal() {
  try {
    Logger.log("🧪 INICIANDO PRUEBA DE ACCESO A PERSONAL");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log(`📊 Spreadsheet obtenido: ${ss.getName()}`);

    const hojas = ss.getSheets();
    Logger.log(
      `📋 Hojas disponibles: ${hojas.map((h) => h.getName()).join(", ")}`
    );

    const hojaPersonal = ss.getSheetByName("Personal");
    if (!hojaPersonal) {
      Logger.log("❌ ERROR: Hoja 'Personal' no encontrada");
      return {
        error: "Hoja Personal no encontrada",
        hojas: hojas.map((h) => h.getName()),
      };
    }

    Logger.log("✅ Hoja 'Personal' encontrada");

    const datos = hojaPersonal.getDataRange().getValues();
    Logger.log(`📊 Datos leídos: ${datos.length} filas`);

    if (datos.length > 0) {
      Logger.log(`📋 Encabezados: ${datos[0].join(" | ")}`);

      if (datos.length > 1) {
        Logger.log("📝 Primeras 3 filas de datos:");
        for (let i = 1; i < Math.min(4, datos.length); i++) {
          Logger.log(`   Fila ${i + 1}: ${datos[i].slice(0, 7).join(" | ")}`);
        }
      } else {
        Logger.log("⚠️ Solo hay encabezados, no hay datos");
      }
    }

    Logger.log("🧪 PRUEBA COMPLETADA");
    return {
      success: true,
      filas: datos.length,
      encabezados: datos[0] || [],
      primeraFila: datos[1] || [],
    };
  } catch (error) {
    Logger.log(`❌ ERROR en probarAccesoPersonal: ${error.message}`);
    Logger.log(`Stack trace: ${error.stack}`);
    return { error: error.message, stack: error.stack };
  }
}

/**
 * Obtiene una lista de nombres de personal de la hoja 'Personal', filtrando por columnas específicas.
 * @param {number[]} columnIndexes Un array de índices de columnas (base 0) de la hoja 'Personal' a incluir.
 * @returns {string[]} Un array de nombres de personal únicos y ordenados.
 */
function obtenerPersonalFiltrado(columnIndexes) {
  Logger.log("🔍 INICIANDO obtenerPersonalFiltrado()");
  Logger.log(
    `📥 Parámetros recibidos: columnIndexes = ${JSON.stringify(columnIndexes)}`
  );

  if (!Array.isArray(columnIndexes)) {
    columnIndexes = [0, 1, 2];
    Logger.log(
      "⚠️ ADVERTENCIA: columnIndexes no recibido, usando [0,1,2] por defecto"
    );
  }

  try {
    Logger.log(
      "🔧 obtenerPersonalFiltrado llamado con columnIndexes: " + columnIndexes
    );

    // Verificar que columnIndexes sea un array válido
    if (!columnIndexes || !Array.isArray(columnIndexes)) {
      Logger.log("❌ ERROR: columnIndexes no es un array válido");
      return []; // Devolver array vacío en lugar de lanzar error
    }

    const hoja =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Personal");
    if (!hoja) {
      Logger.log("❌ ERROR: La hoja 'Personal' no existe.");
      throw new Error("❌ La hoja 'Personal' no existe.");
    }

    Logger.log("✅ Hoja 'Personal' encontrada");
    const datos = hoja.getDataRange().getValues();
    Logger.log(
      `📊 Datos obtenidos de la hoja 'Personal'. Filas: ${datos.length}`
    );

    if (datos.length <= 1) {
      Logger.log("⚠️ No hay datos en la hoja Personal o solo hay encabezados");
      return [];
    }

    // Mostrar encabezados para debugging
    Logger.log(`📋 Encabezados: ${datos[0].join(", ")}`);

    const personal = new Set();
    for (let i = 1; i < datos.length; i++) {
      // Iterar desde la segunda fila (ignorar encabezado)
      const fila = datos[i];
      Logger.log(
        `   Procesando fila ${i + 1}: ${JSON.stringify(fila.slice(0, 7))}`
      ); // Solo mostrar primeras 7 columnas

      columnIndexes.forEach((colIndex) => {
        // Verificar que el índice de columna sea válido
        if (colIndex >= 0 && colIndex < fila.length) {
          const nombre = fila[colIndex];
          Logger.log(
            `     Columna ${colIndex}: "${nombre}" (tipo: ${typeof nombre})`
          );

          if (nombre && typeof nombre === "string" && nombre.trim()) {
            const nombreLimpio = nombre.trim();
            personal.add(nombreLimpio);
            Logger.log(
              `     ✅ Añadido personal: "${nombreLimpio}" de la columna ${colIndex}`
            );
          }
        } else {
          Logger.log(
            `     ⚠️ ADVERTENCIA: Índice de columna fuera de rango: ${colIndex} (fila tiene ${fila.length} columnas)`
          );
        }
      });
    }

    const resultado = [...personal].sort();
    Logger.log(`🎯 RESULTADO: ${resultado.length} personas encontradas`);
    Logger.log("📋 Nombres de personal filtrados: " + resultado.join(", "));

    return resultado;
  } catch (error) {
    Logger.log("❌ ERROR en obtenerPersonalFiltrado: " + error.message);
    Logger.log("Stack trace: " + error.stack);
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
 * VERSIÓN OPTIMIZADA PARA EVITAR ERROR 413 (Request too large)
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
    Logger.log("🔧 generarPdfParaDescarga - Iniciando validaciones");
    Logger.log(
      `📝 HTML Content recibido: ${
        htmlContent ? "SÍ (length: " + htmlContent.length + ")" : "NO/VACÍO"
      }`
    );
    Logger.log(`📁 Filename: ${fileName}`);
    Logger.log(`🎨 Style filename: ${styleFileName}`);
    Logger.log(`📋 Page title: ${pageTitle}`);

    // Verificar parámetros ANTES de procesar
    if (
      !htmlContent ||
      typeof htmlContent !== "string" ||
      htmlContent.trim().length === 0
    ) {
      Logger.log(
        "❌ ERROR CRÍTICO: htmlContent está vacío, nulo o no es string"
      );
      Logger.log(`   Tipo recibido: ${typeof htmlContent}`);
      Logger.log(`   Valor: ${htmlContent}`);
      throw new Error("El contenido HTML no puede estar vacío");
    }

    // Log del contenido para debugging (primeros 200 caracteres)
    Logger.log(
      `📄 Muestra del contenido HTML: ${htmlContent.substring(0, 200)}...`
    );

    if (!fileName) {
      fileName = "Reporte.pdf";
    }

    // OPTIMIZACIÓN 1: Reducir el tamaño del contenido HTML
    const htmlContentOptimizado = optimizarHtmlParaPdf(htmlContent);
    Logger.log(
      `📏 HTML optimizado: ${htmlContentOptimizado.length} caracteres`
    );

    // OPTIMIZACIÓN 2: Usar estilos CSS mínimos y compactos para PDF
    const estilosMinimos = `
      <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { 
          font-family: Arial, sans-serif; 
          font-size: 10px; 
          line-height: 1.3; 
          color: #333; 
          background: white; 
          padding: 15px; 
          max-width: 100%; 
        }
        table { 
          width: 100%; 
          border-collapse: collapse; 
          margin: 8px 0; 
          font-size: 9px; 
        }
        th, td { 
          border: 1px solid #ddd; 
          padding: 4px 3px; 
          text-align: left; 
          vertical-align: top; 
          word-wrap: break-word; 
        }
        th { 
          background-color: #f5f5f5; 
          font-weight: bold; 
          font-size: 8px; 
          text-align: center; 
        }
        h1, h2, h3 { 
          color: #273F4F; 
          margin: 10px 0 5px 0; 
          page-break-after: avoid; 
        }
        h2 { 
          font-size: 14px; 
          border-bottom: 2px solid #FE7743; 
          padding-bottom: 3px; 
          text-align: center; 
        }
        h3 { 
          font-size: 12px; 
          color: #FE7743; 
        }
        .tabla-detalle { font-size: 8px; }
        .tabla-detalle th { padding: 3px 2px; font-size: 7px; }
        .tabla-detalle td { padding: 3px 2px; }
        .resumen-totales { 
          background-color: #f8f9fa; 
          padding: 10px; 
          border-radius: 5px; 
          border: 1px solid #e0e0e0; 
          margin: 10px 0; 
        }
        .currency { 
          text-align: right; 
          font-weight: bold; 
          color: #2e7d32; 
        }
        .no-records { 
          text-align: center; 
          padding: 20px; 
          color: #666; 
        }
        @media print {
          body { font-size: 9px; padding: 10px; }
          table { font-size: 8px; }
          th, td { padding: 2px 1px; font-size: 7px; }
        }
      </style>
    `;

    // Construir el HTML completo con estilos optimizados
    const fullHtml = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="UTF-8">
        <title>${pageTitle}</title>
        ${estilosMinimos}
      </head>
      <body>
        ${htmlContentOptimizado}
      </body>
      </html>
    `;

    Logger.log(`🔧 DEBUG: Generando PDF para: ${fileName}`);
    Logger.log(`📏 DEBUG: Tamaño del HTML: ${fullHtml.length} caracteres`);

    // OPTIMIZACIÓN 3: Validar tamaño antes de generar PDF
    if (fullHtml.length > 50000) {
      // 50KB límite
      Logger.log(
        "⚠️ ADVERTENCIA: HTML muy grande, aplicando compresión adicional"
      );
      return generarPdfComprimido(htmlContentOptimizado, fileName, pageTitle);
    }

    // Crear el PDF
    const htmlOutput = HtmlService.createHtmlOutput(fullHtml);
    const pdfBlob = htmlOutput.getAs(MimeType.PDF).setName(fileName);

    Logger.log("✅ PDF generado para descarga: " + fileName);

    // Convertir a Base64
    return Utilities.base64Encode(pdfBlob.getBytes());
  } catch (e) {
    Logger.log("❌ Error al generar PDF para descarga: " + e.message);

    // FALLBACK: Intentar con versión ultra-comprimida
    if (e.message.includes("413") || e.message.includes("large")) {
      Logger.log("🔄 Intentando generar PDF comprimido como fallback...");
      return generarPdfComprimido(htmlContent, fileName, pageTitle);
    }

    throw new Error("❌ Error al generar el PDF: " + e.message);
  }
}

/**
 * FUNCIÓN AUXILIAR: Optimiza el contenido HTML para reducir su tamaño
 */
function optimizarHtmlParaPdf(htmlContent) {
  try {
    // Remover comentarios HTML
    let htmlOptimizado = htmlContent.replace(/<!--[\s\S]*?-->/g, "");

    // Remover espacios en blanco excesivos
    htmlOptimizado = htmlOptimizado.replace(/\s+/g, " ");

    // Remover atributos de estilo inline innecesarios
    htmlOptimizado = htmlOptimizado.replace(
      /style="[^"]*display:\s*none[^"]*"/g,
      ""
    );

    // Compactar atributos class repetidos
    htmlOptimizado = htmlOptimizado.replace(
      /class="([^"]*)\s+\1"/g,
      'class="$1"'
    );

    // Remover divs vacíos
    htmlOptimizado = htmlOptimizado.replace(/<div[^>]*>\s*<\/div>/g, "");

    Logger.log(
      `HTML optimizado: ${htmlContent.length} → ${htmlOptimizado.length} caracteres`
    );

    return htmlOptimizado;
  } catch (error) {
    Logger.log("⚠️ Error optimizando HTML, usando original: " + error.message);
    return htmlContent;
  }
}

/**
 * FUNCIÓN AUXILIAR: Genera PDF ultra-comprimido como último recurso
 */
function generarPdfComprimido(htmlContent, fileName, pageTitle) {
  try {
    Logger.log("🔧 Generando PDF ultra-comprimido...");

    // Extraer solo el contenido esencial (tablas y texto principal)
    let contenidoEsencial = htmlContent;

    // Remover elementos no esenciales
    contenidoEsencial = contenidoEsencial.replace(
      /<button[^>]*>.*?<\/button>/gi,
      ""
    );
    contenidoEsencial = contenidoEsencial.replace(
      /<script[^>]*>.*?<\/script>/gi,
      ""
    );
    contenidoEsencial = contenidoEsencial.replace(
      /<style[^>]*>.*?<\/style>/gi,
      ""
    );
    contenidoEsencial = contenidoEsencial.replace(/onclick="[^"]*"/gi, "");
    contenidoEsencial = contenidoEsencial.replace(/class="[^"]*"/gi, "");
    contenidoEsencial = contenidoEsencial.replace(/id="[^"]*"/gi, "");

    // HTML mínimo
    const htmlMinimo = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="UTF-8">
        <title>${pageTitle}</title>
        <style>
          body { font-family: Arial; font-size: 9px; margin: 10px; }
          table { width: 100%; border-collapse: collapse; margin: 5px 0; font-size: 8px; }
          th, td { border: 1px solid #ddd; padding: 2px; }
          th { background: #f0f0f0; font-weight: bold; text-align: center; }
          h2 { font-size: 12px; color: #333; text-align: center; margin: 8px 0; }
          h3 { font-size: 10px; margin: 5px 0; }
          .currency { text-align: right; font-weight: bold; }
        </style>
      </head>
      <body>
        ${contenidoEsencial}
      </body>
      </html>
    `;

    Logger.log(`PDF comprimido: ${htmlMinimo.length} caracteres`);

    const htmlOutput = HtmlService.createHtmlOutput(htmlMinimo);
    const pdfBlob = htmlOutput.getAs(MimeType.PDF).setName(fileName);

    Logger.log("✅ PDF ultra-comprimido generado exitosamente");

    return Utilities.base64Encode(pdfBlob.getBytes());
  } catch (error) {
    Logger.log("❌ Error incluso con PDF comprimido: " + error.message);
    throw new Error(
      "❌ No se pudo generar el PDF. El reporte es demasiado grande. Intente con un período más pequeño."
    );
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
  try {
    if (!hojaPersonal) {
      hojaPersonal =
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Personal");
      if (!hojaPersonal) {
        Logger.log("❌ La hoja 'Personal' no existe.");
        return null;
      }
    }

    const datos = hojaPersonal.getDataRange().getValues();
    const nombreNormalizado = normalizarNombreMejorado(nombre);

    Logger.log(`🔍 obtenerTipoDePersonal: Buscando "${nombreNormalizado}"`);
    Logger.log(`📊 Revisando ${datos.length} filas en Personal`);

    // Buscar en todas las columnas de personal (A-F)
    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];

      for (let col = 0; col <= 5; col++) {
        const valorCelda = String(fila[col] || "").trim();

        if (
          valorCelda &&
          sonLaMismaPersonaMejorado(nombreNormalizado, valorCelda)
        ) {
          const tipoEncontrado = TIPOS_PERSONAL[col];
          Logger.log(
            `✅ ENCONTRADO en fila ${
              i + 1
            }, columna ${col}: "${valorCelda}" -> ${tipoEncontrado}`
          );
          return tipoEncontrado;
        }
      }
    }

    Logger.log(`❌ No se encontró "${nombreNormalizado}" en la hoja Personal`);
    return null;
  } catch (error) {
    Logger.log(`❌ Error en obtenerTipoDePersonal: ${error.message}`);
    return null;
  }
}

function calcularCostos(nombre, mes, anio, quincena) {
  if (!nombre || !mes || !anio || !quincena) {
    Logger.log("❌ Parámetros incompletos en calcularCostos");
    return {
      lv: {},
      sab: {},
      totales: { subtotal: 0, iva: 0, total_con_iva: 0 },
      detalleRegistros: [], // ✅ AGREGADO: incluir detalleRegistros siempre
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
  const hojaRegistrosProcedimientos = ss.getSheetByName(
    "RegistrosProcedimientos"
  );
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
  Logger.log(
    `🔍 Nombre normalizado para búsqueda: "${normalizarNombre(nombre)}"`
  );

  // Mostrar muestra de nombres en las primeras 10 filas
  Logger.log(`📋 MUESTRA DE NOMBRES EN REGISTROS:`);
  for (let i = 1; i <= Math.min(10, datos.length - 1); i++) {
    const nombreFila = String(datos[i][1] || "").trim();
    const fechaFila = datos[i][0];
    Logger.log(
      `   Fila ${i + 1}: "${nombreFila}" - ${
        fechaFila instanceof Date ? fechaFila.toDateString() : fechaFila
      }`
    );
  }

  const preciosPorProcedimiento = {};

  Logger.log(`📊 PROCESANDO HOJA DE PRECIOS:`);

  preciosDatos.slice(1).forEach((fila, index) => {
    const nombreProcedimiento = fila[0];
    Logger.log(
      `   Fila ${index + 2}: "${nombreProcedimiento}" -> DoctorLV: ${
        fila[1]
      }, DoctorSab: ${fila[2]}, Anest: ${fila[3]}, Tecnico: ${fila[4]}`
    );

    preciosPorProcedimiento[nombreProcedimiento] = {
      doctorLV: fila[1] || 0,
      doctorSab: fila[2] || 0,
      anest: fila[3] || 0,
      tecnico: fila[4] || 0,
    };
  });

  Logger.log(`🔍 VERIFICANDO PRECIOS PARA PROCEDIMIENTOS FRIA:`);
  Logger.log(
    `   "Proc. Terapéutico Fría Menor": ${JSON.stringify(
      preciosPorProcedimiento["Proc. Terapéutico Fría Menor"]
    )}`
  );
  Logger.log(
    `   "Proc. Terapéutico Fría Mayor": ${JSON.stringify(
      preciosPorProcedimiento["Proc. Terapéutico Fría Mayor"]
    )}`
  );

  Logger.log(`📋 TODOS LOS PROCEDIMIENTOS DISPONIBLES EN PRECIOS:`);
  Object.keys(preciosPorProcedimiento).forEach((nombre) => {
    if (
      nombre.toLowerCase().includes("fría") ||
      nombre.toLowerCase().includes("fria") ||
      nombre.toLowerCase().includes("terapeutico") ||
      nombre.toLowerCase().includes("terapéutico")
    ) {
      Logger.log(`   🔍 ENCONTRADO RELACIONADO: "${nombre}"`);
    }
  });

  function normalizarNombre(str) {
    return String(str || "")
      .replace(/\s+/g, " ")
      .trim()
      .toUpperCase();
  }

  // ✅ FUNCIÓN CORREGIDA: Búsqueda PRECISA para evitar falsos positivos
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

    // 2. MODIFICADO: Solo permitir coincidencia si el nombre seleccionado está completo en el registro
    // Para evitar que "ANEST MANUEL" coincida con "ANEST NICOLE"
    if (registroNorm === seleccionadoNorm) {
      Logger.log(`✅ COINCIDENCIA COMPLETA: "${seleccionadoNorm}"`);
      return true;
    }

    // 3. NUEVO: Búsqueda estricta por nombres únicos
    // Solo coincide si el nombre seleccionado contiene TODAS las palabras del registro
    const palabrasSeleccionado = seleccionadoNorm
      .split(" ")
      .filter((p) => p.length > 1);
    const palabrasRegistro = registroNorm
      .split(" ")
      .filter((p) => p.length > 1);

    // Si el registro tiene menos palabras que la selección, verificar coincidencia completa
    if (palabrasRegistro.length <= palabrasSeleccionado.length) {
      const todasLasPalabrasCoinciden = palabrasRegistro.every((palabraReg) =>
        palabrasSeleccionado.some((palabraSel) => palabraSel === palabraReg)
      );

      if (todasLasPalabrasCoinciden) {
        Logger.log(
          `✅ COINCIDENCIA COMPLETA DE PALABRAS: "${seleccionadoNorm}" ↔ "${registroNorm}"`
        );
        return true;
      }
    }

    // 4. NUEVO: Verificación estricta si ambos nombres tienen exactamente las mismas palabras clave
    if (
      palabrasSeleccionado.length === palabrasRegistro.length &&
      palabrasSeleccionado.length >= 2
    ) {
      const coincidenciasExactas = palabrasSeleccionado.filter((palabraSel) =>
        palabrasRegistro.includes(palabraSel)
      );

      if (coincidenciasExactas.length === palabrasSeleccionado.length) {
        Logger.log(
          `✅ COINCIDENCIA EXACTA DE TODAS LAS PALABRAS: "${seleccionadoNorm}" ↔ "${registroNorm}"`
        );
        return true;
      }
    }

    // Si ninguna condición se cumple, NO hay coincidencia
    Logger.log(`❌ NO COINCIDE: "${seleccionadoNorm}" ≠ "${registroNorm}"`);
    return false;
  }

  // Obtener el tipo de personal
  const tipoPersonal = obtenerTipoDePersonal(nombre, hojaPersonal);

  Logger.log(
    `🏥 Resultado de obtenerTipoDePersonal("${nombre}"): "${tipoPersonal}"`
  );

  if (!tipoPersonal) {
    Logger.log(
      "❌ El nombre '" + nombre + "' no está registrado en la hoja Personal."
    );
    Logger.log("🔍 EJECUTANDO DIAGNÓSTICO ADICIONAL...");

    // Diagnóstico adicional para entender por qué no se encuentra
    const datosPersonalDebug = hojaPersonal.getDataRange().getValues();
    Logger.log(
      `📊 Debug - Total filas en Personal: ${datosPersonalDebug.length}`
    );

    // Buscar manualmente
    let encontradoManual = false;
    for (let i = 1; i < datosPersonalDebug.length; i++) {
      for (let col = 0; col <= 5; col++) {
        const valorCelda = String(datosPersonalDebug[i][col] || "")
          .trim()
          .toUpperCase();
        const nombreNormDebug = String(nombre || "")
          .trim()
          .toUpperCase();

        if (valorCelda === nombreNormDebug) {
          encontradoManual = true;
          Logger.log(
            `🎯 ENCONTRADO MANUALMENTE en fila ${i + 1}, columna ${col}: "${
              datosPersonalDebug[i][col]
            }"`
          );
          Logger.log(
            `🔍 Comparación: "${valorCelda}" === "${nombreNormDebug}" = ${
              valorCelda === nombreNormDebug
            }`
          );
        }
      }
    }

    if (!encontradoManual) {
      Logger.log(`❌ CONFIRMADO: "${nombre}" no existe en la hoja Personal`);
    }

    return {
      lv: {},
      sab: {},
      totales: { subtotal: 0, iva: 0, total_con_iva: 0 },
      detalleRegistros: [], // ✅ AGREGADO: incluir detalleRegistros siempre
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
    detalleRegistros: [], // ✅ NUEVO: Array para almacenar detalles de cada registro
  };

  Object.keys(mapeoProcedimientos).forEach((key) => {
    const nombreTecnico = mapeoProcedimientos[key];
    // Los nombres del mapeo ya son los nombres finales para mostrar
    const nombreParaMostrar = nombreTecnico;

    resumen.lv[key] = {
      nombre: nombreParaMostrar,
      cantidad: 0,
      costo_unitario: 0,
      subtotal: 0,
      iva: 0,
      total_con_iva: 0,
    };
    resumen.sab[key] = {
      nombre: nombreParaMostrar,
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

  Logger.log(
    `📋 Buscando registros para: "${nombre}" en ${mes}/${anio}, quincena: ${quincena}`
  );

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
        Logger.log(
          `📅 Fila ${numeroFila}: PERSONA COINCIDE - "${personaEnRegistro}", Fecha: ${fecha.toDateString()}`
        );
        Logger.log(
          `   🔍 Mes (${filaMes} === ${mes}): ${esMismoMes}, Año (${filaAnio} === ${anio}): ${esMismoAnio}`
        );
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
      Logger.log(
        `   ✅ REGISTRO VÁLIDO #${registrosEncontrados} - Procesando procedimientos...`
      );

      const esSabado = fecha.getDay() === 6;
      let colIndex = 2;
      let procedimientosEnFila = 0;

      // ✅ NUEVO: Objeto para capturar detalles de este registro
      const detalleRegistro = {
        fila: numeroFila,
        fecha: fecha.toLocaleDateString("es-ES", {
          weekday: "long",
          year: "numeric",
          month: "long",
          day: "numeric",
        }),
        fechaOriginal: fecha,
        persona: personaEnRegistro,
        esSabado: esSabado,
        procedimientos: [],
        subtotalRegistro: 0,
      };

      for (const key of Object.keys(mapeoProcedimientos)) {
        const cantidad = Number(fila[colIndex]) || 0;
        if (cantidad > 0) {
          procedimientosEnFila++;
          const procNombre = mapeoProcedimientos[key];

          // 🎨 Función para humanizar nombres de procedimientos para mostrar al usuario
          const humanizarNombreProcedimiento = (nombre) => {
            // Ya no necesitamos transformaciones porque los nombres del mapeo son los finales
            return nombre;
          };

          const nombreParaMostrar = humanizarNombreProcedimiento(procNombre);

          // 🐛 DEBUG ESPECÍFICO para procedimientos FRÍA
          if (key === "asa_fria" || key === "asa_fria2") {
            Logger.log(`🔍 PROCESANDO ${key}:`);
            Logger.log(`   Nombre técnico: "${procNombre}"`);
            Logger.log(`   Nombre para mostrar: "${nombreParaMostrar}"`);
            Logger.log(`   Cantidad: ${cantidad}`);
            Logger.log(
              `   Precios disponibles: ${JSON.stringify(
                preciosPorProcedimiento[procNombre]
              )}`
            );
            Logger.log(`   Tipo personal: ${tipoPersonal}`);
            Logger.log(`   Es sábado: ${esSabado}`);
          }

          const precios = preciosPorProcedimiento[procNombre] || {};
          let costo = 0;

          if (tipoPersonal === "Doctor") {
            costo = esSabado ? precios.doctorSab || 0 : precios.doctorLV || 0;
          } else if (tipoPersonal === "Anestesiólogo") {
            costo = precios.anest || 0;
          } else if (tipoPersonal === "Técnico") {
            costo = precios.tecnico || 0;
          }

          // 🐛 DEBUG ESPECÍFICO para procedimientos FRÍA
          if (key === "asa_fria" || key === "asa_fria2") {
            Logger.log(`   Costo calculado: ${costo}`);
            Logger.log(`   precios.doctorLV: ${precios.doctorLV}`);
            Logger.log(`   precios.doctorSab: ${precios.doctorSab}`);
            Logger.log(`   precios.anest: ${precios.anest}`);
            Logger.log(`   precios.tecnico: ${precios.tecnico}`);
          }

          // Validar que el costo sea un número válido
          if (isNaN(costo) || costo === undefined || costo === null) {
            Logger.log(
              `⚠️ ADVERTENCIA: Precio inválido para ${procNombre} (${tipoPersonal}). Estableciendo a 0.`
            );
            costo = 0;
          }

          Logger.log(
            `💰 DEBUG: ${procNombre} para ${tipoPersonal} - Precio: ${costo} (esSabado: ${esSabado})`
          );

          const subtotal = cantidad * costo;
          const ivaMonto = subtotal * iva;
          const total = subtotal + ivaMonto;

          // ✅ NUEVO: Agregar procedimiento al detalle del registro
          detalleRegistro.procedimientos.push({
            nombre: nombreParaMostrar, // 🎨 Usar nombre humanizado para mostrar
            cantidad: cantidad,
            costoUnitario: costo,
            subtotal: subtotal,
            iva: ivaMonto,
            total: total,
          });
          detalleRegistro.subtotalRegistro += subtotal; // ✅ CORREGIDO: Sumar solo el subtotal (sin IVA)

          const targetResumen = esSabado ? resumen.sab : resumen.lv;
          targetResumen[key].cantidad += cantidad;
          targetResumen[key].costo_unitario = costo;
          targetResumen[key].subtotal += subtotal;
          targetResumen[key].iva += ivaMonto;
          targetResumen[key].total_con_iva += total;

          resumen.totales.subtotal += subtotal;
          resumen.totales.iva += ivaMonto;
          resumen.totales.total_con_iva += total;

          Logger.log(
            `     💰 ${procNombre}: ${cantidad} x ${costo} = $${subtotal}`
          );
        }
        colIndex++;
      }

      if (procedimientosEnFila === 0) {
        Logger.log(`   ⚠️ Fila válida pero sin procedimientos registrados`);
      } else {
        // ✅ NUEVO: Solo agregar registros que tienen procedimientos
        resumen.detalleRegistros.push(detalleRegistro);
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
  Logger.log(
    `   💰 Total calculado: $${resumen.totales.total_con_iva.toFixed(2)}`
  );
  Logger.log(
    `   📋 Detalles de registros capturados: ${resumen.detalleRegistros.length}`
  );

  // ✅ NUEVO: Asegurar que el objeto es JSON válido antes de retornarlo
  try {
    const resumenParaRetornar = {
      lv: resumen.lv || {},
      sab: resumen.sab || {},
      totales: resumen.totales || { subtotal: 0, iva: 0, total_con_iva: 0 },
      detalleRegistros: resumen.detalleRegistros || [],
    };

    // Validar que se puede serializar
    const jsonTest = JSON.stringify(resumenParaRetornar);
    Logger.log(`✅ Objeto JSON válido: ${jsonTest.length} caracteres`);
    Logger.log(
      `🔍 Estructura final a retornar: ${JSON.stringify(
        resumenParaRetornar,
        null,
        2
      )}`
    );

    return resumenParaRetornar;
  } catch (error) {
    Logger.log(`❌ Error preparando objeto para retorno: ${error.message}`);

    // Retornar objeto simple y seguro
    return {
      lv: {},
      sab: {},
      totales: { subtotal: 0, iva: 0, total_con_iva: 0 },
      detalleRegistros: [],
    };
  }
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
 * Función de prueba específica para verificar costos con un nombre conocido
 * Ejecutar manualmente en Apps Script para diagnosticar
 */
function pruebaEspecifica() {
  Logger.log("🧪 === PRUEBA ESPECÍFICA DE COSTOS ===");

  // CAMBIAR ESTOS VALORES POR DATOS REALES DE TU SISTEMA
  const nombrePrueba = "NOMBRE_REAL_AQUÍ"; // Cambiar por un nombre que sepas que existe
  const mesPrueba = 7; // Cambiar por el mes que estés probando
  const anioPrueba = 2025; // Cambiar por el año que estés probando
  const quincenaPrueba = "1-15"; // o "16-31" o "completo"

  try {
    Logger.log(
      `🔍 Probando con: "${nombrePrueba}", ${mesPrueba}/${anioPrueba}, ${quincenaPrueba}`
    );

    // 1. Verificar tipo de personal
    const tipoPersonal = obtenerTipoDePersonal(nombrePrueba);
    Logger.log(`🏥 Tipo de personal: "${tipoPersonal}"`);

    // 2. Verificar precios
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaPrecios = ss.getSheetByName("Precios");
    const preciosDatos = hojaPrecios.getDataRange().getValues();

    // Construir diccionario de precios
    const preciosPorProcedimiento = {};
    preciosDatos.slice(1).forEach((fila) => {
      preciosPorProcedimiento[fila[0]] = {
        doctorLV: fila[1] || 0,
        doctorSab: fila[2] || 0,
        anest: fila[3] || 0,
        tecnico: fila[4] || 0,
      };
    });

    Logger.log(
      `📊 Precios cargados: ${
        Object.keys(preciosPorProcedimiento).length
      } procedimientos`
    );

    // 3. Probar un procedimiento específico
    const procedimientoPrueba = "Consulta MediSmart";
    const precios = preciosPorProcedimiento[procedimientoPrueba];

    if (precios) {
      Logger.log(
        `💰 Precios para "${procedimientoPrueba}": ${JSON.stringify(precios)}`
      );

      // Probar el cálculo
      let costoUnitario = 0;
      if (tipoPersonal === "Doctor") {
        costoUnitario = precios.doctorLV; // Para L-V
        Logger.log(`👨‍⚕️ Doctor L-V: ${costoUnitario}`);
      } else if (tipoPersonal === "Anestesiólogo") {
        costoUnitario = precios.anest;
        Logger.log(`💉 Anestesiólogo: ${costoUnitario}`);
      } else if (tipoPersonal === "Técnico") {
        costoUnitario = precios.tecnico;
        Logger.log(`🔧 Técnico: ${costoUnitario}`);
      } else {
        Logger.log(`❌ Tipo de personal no reconocido: "${tipoPersonal}"`);
      }

      Logger.log(`💰 Costo unitario final: ${costoUnitario}`);
    } else {
      Logger.log(`❌ No se encontraron precios para "${procedimientoPrueba}"`);
    }

    // 4. Llamar a la función principal
    Logger.log("🔄 Ejecutando obtenerDetalleRegistrosSimplificado...");
    const resultado = obtenerDetalleRegistrosSimplificado(
      nombrePrueba,
      mesPrueba,
      anioPrueba,
      quincenaPrueba
    );

    Logger.log(
      `📋 Resultado: ${resultado.detalleRegistros.length} registros encontrados`
    );

    if (resultado.detalleRegistros.length > 0) {
      const primerRegistro = resultado.detalleRegistros[0];
      Logger.log(
        `🔍 Primer registro: ${JSON.stringify(primerRegistro, null, 2)}`
      );
    }
  } catch (error) {
    Logger.log(`❌ Error en prueba específica: ${error.message}`);
    Logger.log(`❌ Stack trace: ${error.stack}`);
  }
}

/**
 * Función de diagnóstico rápido para verificar precios y tipos de personal
 * Ejecutar esta función manualmente en Apps Script para diagnosticar
 */
function diagnosticoRapido() {
  Logger.log("🔍 === DIAGNÓSTICO RÁPIDO ===");

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. Verificar hoja de precios
    const hojaPrecios = ss.getSheetByName("Precios");
    if (!hojaPrecios) {
      Logger.log("❌ No existe la hoja Precios");
      return;
    }

    const preciosDatos = hojaPrecios.getDataRange().getValues();
    Logger.log(`📊 Hoja Precios tiene ${preciosDatos.length} filas`);
    Logger.log(`📋 Encabezados: ${preciosDatos[0].join(", ")}`);

    // Mostrar algunos precios
    for (let i = 1; i < Math.min(6, preciosDatos.length); i++) {
      const fila = preciosDatos[i];
      Logger.log(
        `   ${i}: "${fila[0]}" -> Doctor L-V: ${fila[1]}, Doctor Sáb: ${fila[2]}, Anest: ${fila[3]}, Técnico: ${fila[4]}`
      );
    }

    // 2. Verificar hoja de personal
    const hojaPersonal = ss.getSheetByName("Personal");
    if (!hojaPersonal) {
      Logger.log("❌ No existe la hoja Personal");
      return;
    }

    const personalDatos = hojaPersonal.getDataRange().getValues();
    Logger.log(`👥 Hoja Personal tiene ${personalDatos.length} filas`);
    Logger.log(`📋 Encabezados: ${personalDatos[0].join(", ")}`);

    // Mostrar algunos registros de personal
    for (let i = 1; i < Math.min(6, personalDatos.length); i++) {
      const fila = personalDatos[i];
      Logger.log(
        `   ${i}: Doctor: "${fila[0]}", Anest: "${fila[1]}", Técnico: "${fila[2]}"`
      );
    }

    // 3. Probar obtenerTipoDePersonal con algunos nombres
    const nombresParaProbar = ["MANUEL", "NICOLE", "CARLOS", "MARIA"]; // Cambiar por nombres reales

    nombresParaProbar.forEach((nombre) => {
      const tipo = obtenerTipoDePersonal(nombre, hojaPersonal);
      Logger.log(`🔍 obtenerTipoDePersonal("${nombre}") = "${tipo}"`);
    });

    // 4. Verificar mapeo de tipos
    Logger.log(`🏥 TIPOS_PERSONAL: ${JSON.stringify(TIPOS_PERSONAL)}`);

    Logger.log("✅ Diagnóstico completado - revisa los logs");
  } catch (error) {
    Logger.log(`❌ Error en diagnóstico: ${error.message}`);
  }
}

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
      return {
        error: "Faltan hojas necesarias (RegistrosProcedimientos o Personal)",
      };
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
      encabezadosPersonal: datosPersonal[0],
    };

    // 2. BÚSQUEDA EN HOJA PERSONAL
    const resultadoPersonal = {
      encontrado: false,
      columna: -1,
      fila: -1,
      nombreExacto: "",
      tipo: "",
    };

    for (let i = 1; i < datosPersonal.length; i++) {
      for (let j = 0; j < 3 && j < datosPersonal[i].length; j++) {
        const nombrePersonalHoja = normalizarNombre(datosPersonal[i][j]);
        if (
          nombrePersonalHoja === nombreNormalizado ||
          nombrePersonalHoja.includes(nombreNormalizado) ||
          nombreNormalizado.includes(nombrePersonalHoja)
        ) {
          resultadoPersonal.encontrado = true;
          resultadoPersonal.columna = j;
          resultadoPersonal.fila = i;
          resultadoPersonal.nombreExacto = String(
            datosPersonal[i][j] || ""
          ).trim();
          resultadoPersonal.tipo = obtenerTipoDePersonal(
            nombrePersonal,
            hojaPersonal
          );
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
      totalRegistrosPersonal: 0,
    };

    // Buscar en registros
    for (let i = 1; i < datos.length; i++) {
      const nombreRegistro = normalizarNombre(datos[i][1]);
      const fechaRegistro = datos[i][0];

      // Verificar coincidencia de nombre
      if (
        nombreRegistro === nombreNormalizado ||
        nombreRegistro.includes(nombreNormalizado) ||
        nombreNormalizado.includes(nombreRegistro)
      ) {
        analisisRegistros.totalRegistrosPersonal++;
        analisisRegistros.coincidenciasNombre.push({
          fila: i + 1,
          nombre: String(datos[i][1] || "").trim(),
          fecha:
            fechaRegistro instanceof Date
              ? fechaRegistro.toLocaleDateString()
              : String(fechaRegistro),
        });

        // Análisis de fechas
        if (fechaRegistro) {
          const fecha =
            fechaRegistro instanceof Date
              ? fechaRegistro
              : new Date(fechaRegistro);
          if (!isNaN(fecha.getTime())) {
            if (
              !analisisRegistros.fechaInicio ||
              fecha < analisisRegistros.fechaInicio
            ) {
              analisisRegistros.fechaInicio = fecha;
            }
            if (
              !analisisRegistros.fechaFin ||
              fecha > analisisRegistros.fechaFin
            ) {
              analisisRegistros.fechaFin = fecha;
            }

            // Si se especificó período, verificar coincidencia
            if (mes && anio) {
              const mesRegistro = fecha.getMonth() + 1;
              const anioRegistro = fecha.getFullYear();

              if (
                mesRegistro === Number(mes) &&
                anioRegistro === Number(anio)
              ) {
                analisisRegistros.combinacionNombreFecha++;
                analisisRegistros.registrosPorPeriodo.push({
                  fila: i + 1,
                  fecha: fecha.toLocaleDateString(),
                  nombre: String(datos[i][1] || "").trim(),
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
                fecha:
                  fechaRegistro instanceof Date
                    ? fechaRegistro.toLocaleDateString()
                    : String(fechaRegistro),
                procedimiento: procedimiento,
                cantidad: datos[i][col],
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
      ejemplosFechasPeriodo: [],
    };

    for (let i = 1; i <= Math.min(10, datos.length - 1); i++) {
      const nombre = String(datos[i][1] || "").trim();
      if (nombre) {
        muestreoGeneral.primeros10Nombres.push({
          fila: i + 1,
          nombre: nombre,
          normalizado: normalizarNombre(nombre),
        });
      }
    }

    if (mes && anio) {
      for (let i = 1; i < datos.length; i++) {
        const fechaCelda = datos[i][0];
        if (fechaCelda) {
          const fecha =
            fechaCelda instanceof Date ? fechaCelda : new Date(fechaCelda);
          if (!isNaN(fecha.getTime())) {
            const mesRegistro = fecha.getMonth() + 1;
            const anioRegistro = fecha.getFullYear();

            if (mesRegistro === Number(mes) && anioRegistro === Number(anio)) {
              muestreoGeneral.fechasDelPeriodo++;
              if (muestreoGeneral.ejemplosFechasPeriodo.length < 5) {
                muestreoGeneral.ejemplosFechasPeriodo.push({
                  fila: i + 1,
                  fecha: fecha.toLocaleDateString(),
                  nombre: String(datos[i][1] || "").trim(),
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
        fechaInicio: analisisRegistros.fechaInicio
          ? analisisRegistros.fechaInicio.toLocaleDateString()
          : null,
        fechaFin: analisisRegistros.fechaFin
          ? analisisRegistros.fechaFin.toLocaleDateString()
          : null,
        procedimientosUnicos: Array.from(
          analisisRegistros.procedimientosUnicos
        ),
      },
      muestreoGeneral: muestreoGeneral,

      // DIAGNÓSTICO FINAL
      problemas: [],
      recomendaciones: [],
    };

    // Identificar problemas
    if (!resultadoPersonal.encontrado) {
      resultado.problemas.push("Personal no encontrado en hoja Personal");
      resultado.recomendaciones.push(
        "Verificar ortografía del nombre o agregarlo a la hoja Personal"
      );
    }

    if (analisisRegistros.totalRegistrosPersonal === 0) {
      resultado.problemas.push("No hay registros para este personal");
      resultado.recomendaciones.push(
        "Verificar que el personal tenga procedimientos registrados"
      );
    }

    if (mes && anio && analisisRegistros.combinacionNombreFecha === 0) {
      resultado.problemas.push(`No hay registros para ${mes}/${anio}`);
      resultado.recomendaciones.push(
        "Verificar el período de búsqueda o que existan registros en ese mes"
      );
    }

    if (muestreoGeneral.fechasDelPeriodo === 0 && mes && anio) {
      resultado.problemas.push(
        `No hay registros generales para ${mes}/${anio}`
      );
      resultado.recomendaciones.push(
        "El período especificado no tiene registros en todo el sistema"
      );
    }

    Logger.log("📊 DIAGNÓSTICO COMPLETADO");
    Logger.log(`Problemas identificados: ${resultado.problemas.length}`);
    Logger.log(
      `Registros del personal: ${analisisRegistros.totalRegistrosPersonal}`
    );

    return resultado;
  } catch (error) {
    Logger.log("❌ Error en diagnóstico completo: " + error.message);
    return {
      error: error.message,
      nombreBuscado: nombrePersonal,
      timestamp: new Date().toLocaleString(),
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
    const diagnostico = diagnosticoCompleto(
      nombrePrueba,
      mesPrueba,
      anioPrueba
    );

    Logger.log("\n🔍 RESULTADOS DEL DIAGNÓSTICO:");
    Logger.log("==========================================");

    // 2. Mostrar información de estructura
    Logger.log(`📊 ESTRUCTURA DE DATOS:`);
    Logger.log(
      `   - Filas en RegistrosProcedimientos: ${diagnostico.estructura.totalFilasRegistros}`
    );
    Logger.log(
      `   - Filas en Personal: ${diagnostico.estructura.totalFilasPersonal}`
    );

    // 3. Mostrar si el personal está en la hoja Personal
    Logger.log(`\n👤 PERSONAL EN HOJA PERSONAL:`);
    if (diagnostico.personalEnHojaPersonal.encontrado) {
      Logger.log(
        `   ✅ Encontrado: "${diagnostico.personalEnHojaPersonal.nombreExacto}"`
      );
      Logger.log(
        `   📍 Columna: ${diagnostico.personalEnHojaPersonal.columna}, Fila: ${diagnostico.personalEnHojaPersonal.fila}`
      );
      Logger.log(`   🏥 Tipo: ${diagnostico.personalEnHojaPersonal.tipo}`);
    } else {
      Logger.log(`   ❌ NO encontrado en hoja Personal`);
    }

    // 4. Mostrar análisis de registros
    Logger.log(`\n📋 ANÁLISIS DE REGISTROS:`);
    Logger.log(
      `   - Total registros del personal: ${diagnostico.analisisRegistros.totalRegistrosPersonal}`
    );
    Logger.log(
      `   - Coincidencias de nombre: ${diagnostico.analisisRegistros.coincidenciasNombre.length}`
    );
    Logger.log(
      `   - Período específico: ${diagnostico.analisisRegistros.combinacionNombreFecha} registros`
    );

    if (diagnostico.analisisRegistros.coincidenciasNombre.length > 0) {
      Logger.log(`\n📝 PRIMERAS COINCIDENCIAS DE NOMBRE:`);
      diagnostico.analisisRegistros.coincidenciasNombre
        .slice(0, 5)
        .forEach((c) => {
          Logger.log(`   Fila ${c.fila}: "${c.nombre}" (${c.fecha})`);
        });
    }

    // 5. Mostrar muestra general de datos
    Logger.log(`\n📈 MUESTRA GENERAL:`);
    Logger.log(
      `   - Total registros en período: ${diagnostico.muestreoGeneral.fechasDelPeriodo}`
    );

    if (diagnostico.muestreoGeneral.primeros10Nombres.length > 0) {
      Logger.log(`\n👥 PRIMEROS 10 NOMBRES EN LA HOJA:`);
      diagnostico.muestreoGeneral.primeros10Nombres.forEach((n) => {
        Logger.log(`   Fila ${n.fila}: "${n.nombre}" → "${n.normalizado}"`);
      });
    }

    // 6. Mostrar problemas identificados
    if (diagnostico.problemas.length > 0) {
      Logger.log(`\n❌ PROBLEMAS IDENTIFICADOS:`);
      diagnostico.problemas.forEach((problema) => {
        Logger.log(`   - ${problema}`);
      });
    }

    // 7. Mostrar recomendaciones
    if (diagnostico.recomendaciones.length > 0) {
      Logger.log(`\n💡 RECOMENDACIONES:`);
      diagnostico.recomendaciones.forEach((recomendacion) => {
        Logger.log(`   - ${recomendacion}`);
      });
    }

    // 8. Probar también calcularCostos
    Logger.log(`\n🧮 PROBANDO CALCULAR COSTOS...`);
    const resultado = calcularCostos(
      nombrePrueba,
      mesPrueba,
      anioPrueba,
      "1-15"
    );

    Logger.log(`💰 RESULTADO CALCULAR COSTOS:`);
    Logger.log(`   - Subtotal: $${resultado.totales.subtotal}`);
    Logger.log(`   - IVA: $${resultado.totales.iva}`);
    Logger.log(`   - Total: $${resultado.totales.total_con_iva}`);

    return {
      diagnostico: diagnostico,
      calculoCostos: resultado,
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
/**
 * Función de diagnóstico específica para problemas con hoja Personal
 * @param {string} nombrePersonal - Nombre a verificar
 * @returns {Object} Información detallada sobre la ubicación del personal
 */
function diagnosticarHojaPersonal(nombrePersonal) {
  Logger.log("🔍 DIAGNÓSTICO ESPECÍFICO HOJA PERSONAL");
  Logger.log(`Verificando: "${nombrePersonal}"`);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaPersonal = ss.getSheetByName("Personal");

    if (!hojaPersonal) {
      return { error: "Hoja Personal no encontrada" };
    }

    const datos = hojaPersonal.getDataRange().getValues();
    const encabezados = datos[0] || [];

    Logger.log(`📊 Estructura de hoja Personal:`);
    Logger.log(`   - Total filas: ${datos.length}`);
    Logger.log(`   - Encabezados: ${JSON.stringify(encabezados)}`);

    const nombreNormalizado = normalizarNombre(nombrePersonal);

    const resultado = {
      nombreBuscado: nombrePersonal,
      nombreNormalizado: nombreNormalizado,
      encontrado: false,
      ubicaciones: [],
      encabezados: encabezados,
      totalFilas: datos.length - 1,
      muestraCompleta: [],
    };

    // Buscar en todas las celdas
    for (let fila = 1; fila < datos.length; fila++) {
      const filaData = datos[fila];

      // Muestra completa (primeras 10 filas)
      if (fila <= 10) {
        resultado.muestraCompleta.push({
          fila: fila + 1,
          datos: filaData.map((celda, colIndex) => ({
            columna: String.fromCharCode(65 + colIndex),
            indice: colIndex,
            valor: String(celda || "").trim(),
            normalizado: normalizarNombre(celda),
          })),
        });
      }

      for (let col = 0; col < filaData.length; col++) {
        const valorCelda = String(filaData[col] || "").trim();
        const valorNormalizado = normalizarNombre(valorCelda);

        if (
          valorNormalizado === nombreNormalizado ||
          valorNormalizado.includes(nombreNormalizado) ||
          nombreNormalizado.includes(valorNormalizado)
        ) {
          resultado.encontrado = true;

          let tipoPersonal = "Desconocido";
          switch (col) {
            case 0:
              tipoPersonal = "Doctor";
              break;
            case 1:
              tipoPersonal = "Anestesiólogo";
              break;
            case 2:
              tipoPersonal = "Técnico";
              break;
            case 3:
              tipoPersonal = "Radiólogo";
              break;
            case 4:
              tipoPersonal = "Enfermero";
              break;
            case 5:
              tipoPersonal = "Secretaria";
              break;
          }

          resultado.ubicaciones.push({
            fila: fila + 1,
            columna: String.fromCharCode(65 + col),
            indiceColumna: col,
            valorOriginal: valorCelda,
            valorNormalizado: valorNormalizado,
            tipoPersonal: tipoPersonal,
            coincidencia:
              valorNormalizado === nombreNormalizado ? "EXACTA" : "PARCIAL",
          });

          Logger.log(
            `✅ Encontrado en fila ${fila + 1}, columna ${String.fromCharCode(
              65 + col
            )}: "${valorCelda}" (${tipoPersonal})`
          );
        }
      }
    }

    // Verificar qué devuelve obtenerPersonalFiltrado
    const personalFiltrado = obtenerPersonalFiltrado([0, 1, 2]);
    resultado.enPersonalFiltrado = personalFiltrado.includes(nombrePersonal);
    resultado.personalFiltradoCompleto = personalFiltrado;

    // Verificar qué devuelve obtenerTipoDePersonal
    const tipoDetectado = obtenerTipoDePersonal(nombrePersonal, hojaPersonal);
    resultado.tipoDetectado = tipoDetectado;

    Logger.log(`📋 RESUMEN:`);
    Logger.log(`   - Encontrado: ${resultado.encontrado}`);
    Logger.log(`   - Ubicaciones: ${resultado.ubicaciones.length}`);
    Logger.log(`   - En personal filtrado: ${resultado.enPersonalFiltrado}`);
    Logger.log(`   - Tipo detectado: ${resultado.tipoDetectado}`);

    return resultado;
  } catch (error) {
    Logger.log(`❌ Error en diagnóstico: ${error.message}`);
    return {
      error: error.message,
      nombreBuscado: nombrePersonal,
    };
  }
}

/**
 * Función para verificar específicamente datos de anestesiólogos
 * @returns {Object} Lista completa de anestesiólogos encontrados
 */
function verificarAnestesiologos() {
  Logger.log("🔍 VERIFICANDO TODOS LOS ANESTESIÓLOGOS");

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaPersonal = ss.getSheetByName("Personal");

    if (!hojaPersonal) {
      return { error: "Hoja Personal no encontrada" };
    }

    const datos = hojaPersonal.getDataRange().getValues();
    const anestesiologos = [];

    // Buscar en columna B (índice 1) que corresponde a Anestesiólogos
    for (let fila = 1; fila < datos.length; fila++) {
      const valorCelda = String(datos[fila][1] || "").trim();

      if (valorCelda) {
        anestesiologos.push({
          fila: fila + 1,
          nombre: valorCelda,
          normalizado: normalizarNombre(valorCelda),
          tipo: obtenerTipoDePersonal(valorCelda, hojaPersonal),
          enPersonalFiltrado: obtenerPersonalFiltrado([0, 1, 2]).includes(
            valorCelda
          ),
        });

        Logger.log(
          `Anestesiólogo encontrado en fila ${fila + 1}: "${valorCelda}"`
        );
      }
    }

    const resultado = {
      totalAnestesiologos: anestesiologos.length,
      anestesiologos: anestesiologos,
      personalFiltradoCompleto: obtenerPersonalFiltrado([0, 1, 2]),
    };

    Logger.log(`📊 Total anestesiólogos encontrados: ${anestesiologos.length}`);

    return resultado;
  } catch (error) {
    Logger.log(`❌ Error verificando anestesiólogos: ${error.message}`);
    return { error: error.message };
  }
}

/**
 * Función de prueba para verificar un anestesiólogo específico
 * Ejecuta esta función directamente en el editor de Apps Script
 */
function probarAnestesiologo() {
  // Cambia este nombre por el del anestesiólogo que estás probando
  const nombreAPrueba = "Nombre del Anestesiólogo"; // ⬅️ CAMBIA ESTE NOMBRE

  Logger.log("=".repeat(50));
  Logger.log("🧪 PRUEBA ESPECÍFICA DE ANESTESIÓLOGO");
  Logger.log("=".repeat(50));

  // 1. Diagnóstico específico
  const diagnostico = diagnosticarHojaPersonal(nombreAPrueba);
  Logger.log("1️⃣ DIAGNÓSTICO ESPECÍFICO:");
  Logger.log(JSON.stringify(diagnostico, null, 2));

  // 2. Verificar todos los anestesiólogos
  const todosAnestesiologos = verificarAnestesiologos();
  Logger.log("\n2️⃣ TODOS LOS ANESTESIÓLOGOS:");
  Logger.log(JSON.stringify(todosAnestesiologos, null, 2));

  // 3. Probar calcularCostos
  Logger.log("\n3️⃣ PRUEBA DE CALCULAR COSTOS:");
  try {
    const resultado = calcularCostos(nombreAPrueba, 12, 2024, 1); // Diciembre 2024, primera quincena
    Logger.log(
      `Resultado calcularCostos: ${JSON.stringify(resultado, null, 2)}`
    );
  } catch (error) {
    Logger.log(`❌ Error en calcularCostos: ${error.message}`);
  }

  Logger.log("=".repeat(50));
  Logger.log("✅ Prueba completada. Revisa los logs arriba.");
  Logger.log("=".repeat(50));
}

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
  try {
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

    const ultimaFila = hoja.getLastRow();
    const filaDestino = ultimaFila + 1;
    hoja.getRange(filaDestino, 1, 1, fila.length).setValues([fila]);

    // Formatear la celda de fecha
    hoja.getRange(filaDestino, 1).setNumberFormat("dd/MM/yyyy");

    Logger.log(
      "✅ Registro de ultrasonido guardado con éxito en la fila " + filaDestino
    );

    return {
      success: true,
      fila: filaDestino,
      message: "Registro guardado exitosamente",
    };
  } catch (error) {
    Logger.log("❌ Error al guardar registro de ultrasonido: " + error.message);
    throw new Error("Error al guardar el registro: " + error.message);
  }
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
  const detalleRegistros = []; // ✅ NUEVO: Array para detalles por registro
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
  datosRegistros.slice(1).forEach((fila, indiceRealFila) => {
    const fecha = new Date(fila[0]);
    const filaRadiologo = fila[1];
    const filaMes = fecha.getMonth() + 1;
    const filaAnio = fecha.getFullYear();
    const filaDia = fecha.getDate();

    if (filaRadiologo !== radiologo || filaMes !== mes || filaAnio !== anio)
      return;
    if (quincena === "1-15" && filaDia > 15) return;
    if (quincena === "16-31" && filaDia <= 15) return;

    // ✅ NUEVO: Crear objeto de detalle para este registro
    const registroDetalle = {
      fecha: fecha.toLocaleDateString("es-CR", {
        weekday: "long",
        year: "numeric",
        month: "long",
        day: "numeric",
      }),
      fechaOriginal: fecha.toISOString(),
      persona: filaRadiologo,
      fila: indiceRealFila + 2, // +2 porque slice(1) y las filas empiezan en 1
      ultrasonidos: [],
      subtotalRegistro: 0,
    };

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

          // Actualizar resumen agregado
          resumen[key].cantidad += cantidad;
          resumen[key].ganancia_sin_iva += gananciaSinIVA;
          resumen[key].iva_monto += ivaMonto;
          resumen[key].ganancia_con_iva += gananciaConIVA;
          resumen[key].costo_unitario = costoUnitario;

          // ✅ NUEVO: Agregar al detalle del registro
          registroDetalle.ultrasonidos.push({
            nombre: tipoUltrasonido,
            cantidad: cantidad,
            costoUnitario: costoUnitario,
            subtotal: gananciaSinIVA,
            iva: ivaMonto,
            total: gananciaConIVA,
          });

          registroDetalle.subtotalRegistro += gananciaSinIVA;

          totalGananciaSinIVA += gananciaSinIVA;
          totalIVA += ivaMonto;
          totalGananciaConIVA += gananciaConIVA;
        }
      }
      currentSheetColIndex++;
    }

    // ✅ NUEVO: Solo agregar el registro si tiene ultrasonidos
    if (registroDetalle.ultrasonidos.length > 0) {
      detalleRegistros.push(registroDetalle);
    }
  });
  Logger.log("--- Fin del procesamiento de registros ---");

  // ✅ NUEVO: Devolver respuesta estructurada con detalle de registros
  return {
    success: true,
    resumen,
    totalGananciaSinIVA,
    totalIVA,
    totalGananciaConIVA,
    detalleRegistros, // ✅ NUEVO: Detalle por registro
    metadatos: {
      radiologo: radiologo,
      periodo: `${mes}/${anio}`,
      quincena: quincena,
      totalRegistros: detalleRegistros.length,
    },
  };
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
    fileName,
    "stylesReportePagoUltrasonido",
    reportTitle
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
function guardarRegistroHorasExtras(data) {
  Logger.log(
    "guardarRegistroHorasExtras llamado con datos: " + JSON.stringify(data)
  );

  if (!data || typeof data !== "object") {
    Logger.log(
      "ERROR: No se recibieron datos válidos para el registro de horas extras."
    );
    throw new Error(
      "❌ No se recibieron datos válidos para el registro de horas extras."
    );
  }

  // Validar datos obligatorios
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
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    "RegistroHorasExtras"
  );
  if (!hoja) {
    Logger.log("ERROR: La hoja 'RegistroHorasExtras' no existe.");
    throw new Error("❌ La hoja 'RegistroHorasExtras' no existe.");
  }

  try {
    // Convertir la fecha de string a objeto Date
    let [year, month, day] = data.fecha.split("-").map(Number);
    let fechaRegistro = new Date(Date.UTC(year, month - 1, day, 12, 0, 0));

    // Preparar la fila para añadir - SOLO los campos básicos necesarios
    const fila = [
      fechaRegistro,
      data.trabajador,
      data.cantidad_horas,
      data.comentarios || "",
    ];

    // Agregar al final de los datos (orden cronológico natural)
    const ultimaFila = hoja.getLastRow();
    hoja.getRange(ultimaFila + 1, 1, 1, fila.length).setValues([fila]);

    // Formatear las celdas
    hoja.getRange(ultimaFila + 1, 1).setNumberFormat("dd/MM/yyyy"); // Fecha
    hoja.getRange(ultimaFila + 1, 3).setNumberFormat("0.0"); // Horas
    Logger.log(
      "✅ Registro de horas extras guardado con éxito en orden cronológico."
    );
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

function calcularPagosHorasExtras(trabajador, mes, anio, quincena) {
  Logger.log(
    `🚀 INICIANDO calcularPagosHorasExtras: trabajador=${trabajador}, mes=${mes}, anio=${anio}, quincena=${quincena}`
  );

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaHorasExtras = ss.getSheetByName("RegistroHorasExtras");

  if (!hojaHorasExtras) {
    throw new Error("Faltan hojas necesarias: 'RegistroHorasExtras'.");
  }

  const datosHorasExtras = hojaHorasExtras.getDataRange().getValues();
  const detalleRegistros = [];

  let totalHoras = 0;
  let totalSinImpuesto = 0;
  let totalConImpuesto = 0;
  let totalImpuesto = 0;

  Logger.log("--- Procesando registros de Horas Extras ---");

  datosHorasExtras.slice(1).forEach((fila, indiceRealFila) => {
    const numeroFila = indiceRealFila + 2; // +2 porque slice(1) quita header y los índices empiezan en 0

    try {
      const fecha = new Date(fila[0]);
      const filaTrabajador = fila[1];
      const cantidadHoras = Number(fila[2]) || 0;
      const tarifaPorHora = Number(fila[3]) || 0;
      const aplicaImpuesto =
        fila[4] === true || fila[4] === "TRUE" || fila[4] === 1;
      const porcentajeImpuesto = Number(fila[5]) || 0;
      const comentarios = fila[6] || "";

      const filaMes = fecha.getMonth() + 1;
      const filaAnio = fecha.getFullYear();
      const filaDia = fecha.getDate();

      // Verificar coincidencias exactas
      if (
        filaTrabajador !== trabajador ||
        filaMes !== parseInt(mes) ||
        filaAnio !== parseInt(anio)
      ) {
        return;
      }

      // Verificar quincena
      if (quincena === "1-15" && filaDia > 15) return;
      if (quincena === "16-31" && filaDia <= 15) return;

      // Calcular montos
      const subtotal = cantidadHoras * tarifaPorHora;
      const impuesto = aplicaImpuesto
        ? subtotal * (porcentajeImpuesto / 100)
        : 0;
      const totalRegistro = subtotal + impuesto;

      const registro = {
        fila: numeroFila,
        fecha: Utilities.formatDate(
          fecha,
          ss.getSpreadsheetTimeZone(),
          "dd/MM/yyyy"
        ),
        fechaCompleta: Utilities.formatDate(
          fecha,
          ss.getSpreadsheetTimeZone(),
          "EEEE, dd 'de' MMMM 'de' yyyy"
        ),
        trabajador: filaTrabajador,
        cantidadHoras: cantidadHoras,
        tarifaPorHora: tarifaPorHora,
        subtotal: subtotal,
        aplicaImpuesto: aplicaImpuesto,
        impuesto: impuesto,
        total: totalRegistro,
        comentarios: comentarios,
      };

      detalleRegistros.push(registro);

      totalHoras += cantidadHoras;
      totalSinImpuesto += subtotal;
      totalImpuesto += impuesto;
      totalConImpuesto += totalRegistro;

      Logger.log(
        `✅ Registro procesado - Fila ${numeroFila}: ${fecha.toDateString()}, ${cantidadHoras}h, $${subtotal}, Impuesto: ${
          aplicaImpuesto ? "$" + impuesto.toFixed(2) : "No"
        }`
      );
    } catch (error) {
      Logger.log(`❌ Error procesando fila ${numeroFila}: ${error.message}`);
    }
  });

  Logger.log("--- Fin del procesamiento de registros de Horas Extras ---");

  // Calcular totales finales
  const totales = {
    totalHoras: totalHoras,
    subtotal: totalSinImpuesto,
    impuesto: totalImpuesto,
    total_con_impuesto: totalConImpuesto,
  };

  Logger.log(
    `📊 TOTALES: ${totalHoras} horas, Subtotal: $${totalSinImpuesto.toFixed(
      2
    )}, Impuesto: $${totalImpuesto.toFixed(
      2
    )}, Total: $${totalConImpuesto.toFixed(2)}`
  );

  return {
    detalleRegistros: detalleRegistros,
    totales: totales,
    metadatos: {
      trabajador: trabajador,
      periodo: `${mes}/${anio}`,
      quincena: quincena,
      numeroRegistros: detalleRegistros.length,
      fechaGeneracion: new Date().toISOString(),
    },
  };
}

/**
 * Nueva función que calcula pagos de horas extras con parámetros de cálculo personalizados
 * @param {string} trabajador - Nombre del trabajador
 * @param {number} mes - Mes para el reporte
 * @param {number} anio - Año para el reporte
 * @param {string} quincena - Quincena ("completo", "1-15", "16-31")
 * @param {Object} parametrosCalculo - Parámetros personalizados de cálculo
 * @returns {Object} Resultado del cálculo con parámetros personalizados
 */
function calcularPagosHorasExtrasConParametros(
  trabajador,
  mes,
  anio,
  quincena,
  parametrosCalculo
) {
  Logger.log(
    `🚀 INICIANDO calcularPagosHorasExtrasConParametros: trabajador=${trabajador}, mes=${mes}, anio=${anio}, quincena=${quincena}`
  );
  Logger.log(`📊 Parámetros de cálculo: ${JSON.stringify(parametrosCalculo)}`);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaHorasExtras = ss.getSheetByName("RegistroHorasExtras");

  if (!hojaHorasExtras) {
    throw new Error("Faltan hojas necesarias: 'RegistroHorasExtras'.");
  }

  // Extraer parámetros de cálculo con valores por defecto
  const tarifaPorHoraPersonalizada = parametrosCalculo.tarifaPorHora || 3500;
  const aplicaImpuestoPersonalizado =
    parametrosCalculo.aplicaImpuesto !== undefined
      ? parametrosCalculo.aplicaImpuesto
      : true;
  const porcentajeImpuestoPersonalizado =
    parametrosCalculo.porcentajeImpuesto || 13;

  Logger.log(`💰 Usando tarifa personalizada: ₡${tarifaPorHoraPersonalizada}`);
  Logger.log(`📋 Aplicar impuesto: ${aplicaImpuestoPersonalizado}`);
  Logger.log(`📊 Porcentaje de impuesto: ${porcentajeImpuestoPersonalizado}%`);

  const datosHorasExtras = hojaHorasExtras.getDataRange().getValues();
  const detalleRegistros = [];

  let totalHoras = 0;
  let totalSinImpuesto = 0;
  let totalConImpuesto = 0;
  let totalImpuesto = 0;

  Logger.log(
    "--- Procesando registros de Horas Extras con parámetros personalizados ---"
  );

  datosHorasExtras.slice(1).forEach((fila, indiceRealFila) => {
    const numeroFila = indiceRealFila + 2; // +2 porque slice(1) quita header y los índices empiezan en 0

    try {
      const fecha = new Date(fila[0]);
      const filaTrabajador = fila[1];
      const cantidadHoras = Number(fila[2]) || 0;
      // SOLO comentarios de la estructura simplificada (columna 4)
      const comentarios = fila[3] || "";

      const filaMes = fecha.getMonth() + 1;
      const filaAnio = fecha.getFullYear();
      const filaDia = fecha.getDate();

      // Verificar coincidencias exactas
      if (
        filaTrabajador !== trabajador ||
        filaMes !== parseInt(mes) ||
        filaAnio !== parseInt(anio)
      ) {
        return;
      }

      // Verificar quincena
      if (quincena === "1-15" && filaDia > 15) return;
      if (quincena === "16-31" && filaDia <= 15) return;

      // Calcular montos usando parámetros personalizados
      const subtotal = cantidadHoras * tarifaPorHoraPersonalizada;
      const impuesto = aplicaImpuestoPersonalizado
        ? subtotal * (porcentajeImpuestoPersonalizado / 100)
        : 0;
      const totalRegistro = subtotal + impuesto;

      const registro = {
        fila: numeroFila,
        fecha: Utilities.formatDate(
          fecha,
          ss.getSpreadsheetTimeZone(),
          "dd/MM/yyyy"
        ),
        fechaCompleta: Utilities.formatDate(
          fecha,
          ss.getSpreadsheetTimeZone(),
          "EEEE, dd 'de' MMMM 'de' yyyy"
        ),
        trabajador: filaTrabajador,
        cantidadHoras: cantidadHoras,
        tarifaPorHora: tarifaPorHoraPersonalizada, // Usar tarifa personalizada
        subtotal: subtotal,
        aplicaImpuesto: aplicaImpuestoPersonalizado, // Usar configuración personalizada
        porcentajeImpuesto: porcentajeImpuestoPersonalizado, // Usar porcentaje personalizado
        impuesto: impuesto,
        total: totalRegistro,
        comentarios: comentarios,
      };

      detalleRegistros.push(registro);

      totalHoras += cantidadHoras;
      totalSinImpuesto += subtotal;
      totalImpuesto += impuesto;
      totalConImpuesto += totalRegistro;

      Logger.log(
        `✅ Registro procesado - Fila ${numeroFila}: ${fecha.toDateString()}, ${cantidadHoras}h, $${subtotal}, Impuesto: ${
          aplicaImpuestoPersonalizado ? "$" + impuesto.toFixed(2) : "No"
        }`
      );
    } catch (error) {
      Logger.log(`❌ Error procesando fila ${numeroFila}: ${error.message}`);
    }
  });

  Logger.log(
    "--- Fin del procesamiento de registros de Horas Extras con parámetros personalizados ---"
  );

  // Calcular totales finales
  const totales = {
    totalHoras: totalHoras,
    subtotal: totalSinImpuesto,
    impuesto: totalImpuesto,
    total_con_impuesto: totalConImpuesto,
  };

  Logger.log(
    `📊 TOTALES PERSONALIZADOS: ${totalHoras} horas, Subtotal: $${totalSinImpuesto.toFixed(
      2
    )}, Impuesto: $${totalImpuesto.toFixed(
      2
    )}, Total: $${totalConImpuesto.toFixed(2)}`
  );

  return {
    detalleRegistros: detalleRegistros,
    totales: totales,
    parametrosUsados: {
      tarifaPorHora: tarifaPorHoraPersonalizada,
      aplicaImpuesto: aplicaImpuestoPersonalizado,
      porcentajeImpuesto: porcentajeImpuestoPersonalizado,
    },
    metadatos: {
      trabajador: trabajador,
      periodo: `${mes}/${anio}`,
      quincena: quincena,
      numeroRegistros: detalleRegistros.length,
      fechaGeneracion: new Date().toISOString(),
      calculoPersonalizado: true,
    },
  };
}

// --- Funciones adicionales para Horas Extras ---

/**
 * Obtiene la lista de trabajadores que han registrado horas extras
 */
function obtenerTrabajadoresHorasExtras() {
  try {
    const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      "RegistroHorasExtras"
    );
    if (!hoja) {
      throw new Error("❌ La hoja 'RegistroHorasExtras' no existe.");
    }

    const datos = hoja.getDataRange().getValues();
    const trabajadores = new Set();

    // Obtener nombres únicos de trabajadores de la columna B (índice 1)
    for (let i = 1; i < datos.length; i++) {
      const trabajador = datos[i][1];
      if (trabajador && typeof trabajador === "string" && trabajador.trim()) {
        trabajadores.add(trabajador.trim());
      }
    }

    const resultado = [...trabajadores].sort();
    Logger.log(
      "Trabajadores con horas extras encontrados: " + resultado.join(", ")
    );
    return resultado;
  } catch (error) {
    Logger.log("ERROR en obtenerTrabajadoresHorasExtras: " + error.message);
    return [];
  }
}

/**
 * Busca registros de horas extras por criterios específicos
 */
function buscarRegistrosHorasExtras(tipoBusqueda, valorBusqueda) {
  Logger.log(
    `🔍 Buscando registros de horas extras: ${tipoBusqueda} = ${valorBusqueda}`
  );

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("RegistroHorasExtras");

    if (!hoja) {
      throw new Error("❌ La hoja 'RegistroHorasExtras' no existe.");
    }

    const datos = hoja.getDataRange().getValues();
    const resultados = [];

    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      const fecha = new Date(fila[0]);
      const trabajador = fila[1];
      const cantidadHoras = fila[2];
      const tarifaPorHora = fila[3] || 0;
      const aplicaImpuesto = fila[4] || false;
      const comentarios = fila[5] || "";

      let coincide = false;

      switch (tipoBusqueda) {
        case "fecha":
          const fechaBusqueda = new Date(valorBusqueda);
          coincide = fecha.toDateString() === fechaBusqueda.toDateString();
          break;
        case "trabajador":
          coincide =
            trabajador &&
            trabajador.toLowerCase().includes(valorBusqueda.toLowerCase());
          break;
        case "mes":
          const [mes, anio] = valorBusqueda.split("/");
          coincide =
            fecha.getMonth() + 1 === parseInt(mes) &&
            fecha.getFullYear() === parseInt(anio);
          break;
        default:
          Logger.log(`❌ Tipo de búsqueda no válido: ${tipoBusqueda}`);
          return [];
      }

      if (coincide) {
        const subtotal = cantidadHoras * tarifaPorHora;
        const impuesto = aplicaImpuesto ? subtotal * 0.13 : 0;
        const total = subtotal + impuesto;

        resultados.push({
          fila: i + 1,
          fecha: Utilities.formatDate(
            fecha,
            ss.getSpreadsheetTimeZone(),
            "dd/MM/yyyy"
          ),
          trabajador: trabajador,
          cantidadHoras: cantidadHoras,
          tarifaPorHora: tarifaPorHora,
          subtotal: subtotal,
          aplicaImpuesto: aplicaImpuesto,
          impuesto: impuesto,
          total: total,
          comentarios: comentarios,
        });
      }
    }

    Logger.log(
      `✅ Búsqueda completada: ${resultados.length} registros encontrados`
    );
    return resultados;
  } catch (error) {
    Logger.log(`❌ Error en búsqueda: ${error.message}`);
    throw new Error(`Error al buscar registros: ${error.message}`);
  }
}

/**
 * Actualiza un registro existente de horas extras
 */
function actualizarRegistroHorasExtras(fila, datosActualizados) {
  Logger.log(
    `🔄 Actualizando registro en fila ${fila}: ${JSON.stringify(
      datosActualizados
    )}`
  );

  try {
    const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      "RegistroHorasExtras"
    );

    if (!hoja) {
      throw new Error("❌ La hoja 'RegistroHorasExtras' no existe.");
    }

    // Validar que la fila existe
    if (fila < 2 || fila > hoja.getLastRow()) {
      throw new Error(`❌ Fila ${fila} no válida.`);
    }

    // Preparar los datos actualizados
    const fecha = new Date(datosActualizados.fecha);
    const trabajador = datosActualizados.trabajador;
    const cantidadHoras = parseFloat(datosActualizados.cantidad_horas);
    const tarifaPorHora = parseFloat(datosActualizados.tarifa_por_hora) || 0;
    const aplicaImpuesto = datosActualizados.aplica_impuesto || false;
    const comentarios = datosActualizados.comentarios || "";

    // Actualizar la fila
    hoja.getRange(fila, 1).setValue(fecha);
    hoja.getRange(fila, 2).setValue(trabajador);
    hoja.getRange(fila, 3).setValue(cantidadHoras);
    hoja.getRange(fila, 4).setValue(tarifaPorHora);
    hoja.getRange(fila, 5).setValue(aplicaImpuesto);
    hoja.getRange(fila, 6).setValue(comentarios);

    // Aplicar formato
    hoja.getRange(fila, 1).setNumberFormat("dd/MM/yyyy");
    hoja.getRange(fila, 3).setNumberFormat("0.0");
    hoja.getRange(fila, 4).setNumberFormat("$#,##0.00");

    Logger.log(`✅ Registro actualizado exitosamente en fila ${fila}`);
    return true;
  } catch (error) {
    Logger.log(`❌ Error actualizando registro: ${error.message}`);
    throw new Error(`Error al actualizar registro: ${error.message}`);
  }
}

/**
 * Elimina un registro de horas extras
 */
function eliminarRegistroHorasExtras(fila) {
  Logger.log(`🗑️ Eliminando registro en fila ${fila}`);

  try {
    const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      "RegistroHorasExtras"
    );

    if (!hoja) {
      throw new Error("❌ La hoja 'RegistroHorasExtras' no existe.");
    }

    // Validar que la fila existe y no es el header
    if (fila < 2 || fila > hoja.getLastRow()) {
      throw new Error(`❌ Fila ${fila} no válida.`);
    }

    // Eliminar la fila
    hoja.deleteRow(fila);

    Logger.log(`✅ Registro eliminado exitosamente de fila ${fila}`);
    return true;
  } catch (error) {
    Logger.log(`❌ Error eliminando registro: ${error.message}`);
    throw new Error(`Error al eliminar registro: ${error.message}`);
  }
}

/**
 * Funciones Wrapper para Enviar/Descargar Reportes de Horas Extras
 */
function enviarReporteHorasExtras(
  htmlContent,
  emailRecipient,
  reportTitle,
  subject,
  fileName
) {
  return enviarReportePDF(
    htmlContent,
    emailRecipient,
    reportTitle,
    subject,
    fileName
  );
}

/**
 * Función específica para generar PDF de horas extras
 * @param {string} htmlContent - Contenido HTML del reporte
 * @param {string} fileName - Nombre del archivo PDF
 * @param {string} reportTitle - Título del reporte
 * @returns {string} PDF en formato Base64
 */
function descargarReporteHorasExtrasPDF(htmlContent, fileName, reportTitle) {
  Logger.log("🔧 descargarReporteHorasExtrasPDF iniciado");
  Logger.log(
    `📝 HTML Content length: ${htmlContent ? htmlContent.length : "NULL"}`
  );
  Logger.log(`📁 Filename: ${fileName}`);
  Logger.log(`📋 Report Title: ${reportTitle}`);

  try {
    // Validar que el contenido HTML no esté vacío
    if (!htmlContent || htmlContent.trim() === "") {
      Logger.log("❌ Error: htmlContent está vacío");
      throw new Error(
        "El contenido HTML para el PDF está vacío. Por favor, genere el reporte primero."
      );
    }

    // Validar parámetros
    if (!fileName) {
      fileName = "Reporte_Horas_Extras.pdf";
      Logger.log(`📁 Filename establecido por defecto: ${fileName}`);
    }

    if (!reportTitle) {
      reportTitle = "Reporte de Horas Extras";
      Logger.log(`📋 Report title establecido por defecto: ${reportTitle}`);
    }

    // Llamar a la función de generación de PDF con parámetros específicos para horas extras
    const pdfBase64 = generarPdfParaDescarga(
      htmlContent,
      fileName,
      "stylesReportePagoHorasExtras", // Archivo de estilos específico
      reportTitle
    );

    Logger.log("✅ PDF de horas extras generado exitosamente");
    return pdfBase64;
  } catch (error) {
    Logger.log(`❌ Error en descargarReporteHorasExtrasPDF: ${error.message}`);
    throw new Error(
      `Error al generar el PDF de horas extras: ${error.message}`
    );
  }
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
 * Función de prueba directa para biopsias - versión que funciona como el test
 * @param {Object} data - Datos del registro de biopsia.
 * @returns {boolean} - Verdadero si se guardó correctamente.
 */
function testGuardarBiopsia(data) {
  try {
    Logger.log("🧪 [TEST] Iniciando testGuardarBiopsia");
    Logger.log("🔍 [TEST] Data recibida: " + JSON.stringify(data));

    const hoja =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RegistroBiopsias");
    if (!hoja) {
      throw new Error("Hoja RegistroBiopsias no encontrada");
    }

    // Validaciones básicas
    if (!data || !data.nombre_cliente || !data.medico || !data.fecha_toma) {
      throw new Error("Faltan datos obligatorios");
    }

    // Procesar fecha
    let fechaProcesada;
    try {
      const [year, month, day] = data.fecha_toma.split("-").map(Number);
      fechaProcesada = new Date(Date.UTC(year, month - 1, day, 12, 0, 0));
    } catch (e) {
      fechaProcesada = data.fecha_toma; // Usar como string si falla el procesamiento
    }

    // Construir fila
    const fila = [
      fechaProcesada,
      false, // recibida
      false, // enviada
      data.cedula || "",
      data.telefono || "",
      data.nombre_cliente,
      Number(data.frascos_gastro) || 0,
      Number(data.frascos_colon) || 0,
      data.medico,
      data.comentario || "",
    ];

    const ultimaFila = hoja.getLastRow();
    const nuevaFila = ultimaFila + 1;

    Logger.log("💾 [TEST] Insertando en fila " + nuevaFila);
    hoja.getRange(nuevaFila, 1, 1, fila.length).setValues([fila]);

    // Aplicar formatos básicos
    hoja.getRange(nuevaFila, 1).setNumberFormat("yyyy-mm-dd");
    hoja.getRange(nuevaFila, 2, 1, 2).insertCheckboxes();

    Logger.log("✅ [TEST] Guardado exitoso en fila " + nuevaFila);
    return true;
  } catch (error) {
    Logger.log("❌ [TEST] Error: " + error.message);
    throw new Error("Error en test: " + error.message);
  }
}

/**
 * Guarda un registro de biopsia en la hoja "Biopsias".
 * @param {Object} data - Datos del registro de biopsia.
 * @returns {boolean} - Verdadero si se guardó correctamente.
 */
function guardarRegistroBiopsia(data) {
  try {
    Logger.log("🧪 Iniciando función guardarRegistroBiopsia");
    Logger.log("📊 Tipo de dato recibido: " + typeof data);
    Logger.log("🔍 Valor de data: " + String(data));
    Logger.log(
      "🧪 Guardando registro de biopsia (JSON): " + JSON.stringify(data)
    );

    if (data === undefined) {
      throw new Error(
        "Los datos están undefined. Posible problema en el frontend."
      );
    }

    if (data === null) {
      throw new Error("Los datos están null. Posible problema en el frontend.");
    }

    if (!data || typeof data !== "object") {
      throw new Error(
        `No se recibieron datos válidos para el registro de biopsia. Tipo: ${typeof data}, Valor: ${String(
          data
        )}`
      );
    }

    const hoja =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RegistroBiopsias");
    if (!hoja) {
      throw new Error("La hoja 'RegistroBiopsias' no existe.");
    }

    Logger.log("📊 Hoja encontrada. Filas actuales: " + hoja.getLastRow());

    // Validar datos obligatorios
    if (!data.nombre_cliente || data.nombre_cliente.trim() === "") {
      throw new Error("El nombre del paciente es obligatorio.");
    }

    if (!data.medico || data.medico === "") {
      throw new Error("El médico es obligatorio.");
    }

    // Validar y procesar fecha de toma
    if (!data.fecha_toma) {
      throw new Error("La fecha de toma es obligatoria.");
    }
    const fechaToma = procesarFecha(data.fecha_toma);
    Logger.log("📅 Fecha procesada: " + fechaToma);

    // Limpiar cédula de guiones y espacios (puede ser cédula, pasaporte o DIMEX)
    let cedulaLimpia = "";
    if (data.cedula && data.cedula.trim() !== "") {
      cedulaLimpia = data.cedula.replace(/[-\s]/g, "").trim();
      Logger.log(
        "🆔 Cédula original: '" +
          data.cedula +
          "' -> Limpia: '" +
          cedulaLimpia +
          "'"
      );
    }

    // Preparar la fila para añadir (estructura simplificada de 10 columnas)
    const fila = [
      fechaToma, // A - Fecha Toma
      false, // B - Recibida (checkbox) - por defecto false
      false, // C - Enviada (checkbox) - por defecto false
      cedulaLimpia, // D - Cédula (sin guiones)
      data.telefono || "", // E - Teléfono
      data.nombre_cliente.trim(), // F - Nombre Cliente
      Number(data.frascos_gastro) || 0, // G - Frascos Gastro
      Number(data.frascos_colon) || 0, // H - Frascos Colon
      data.medico, // I - Médico
      data.comentario || "", // J - Comentario
    ];

    Logger.log("🔍 Datos de la fila a insertar: " + JSON.stringify(fila));
    Logger.log("📏 Longitud de la fila: " + fila.length + " columnas");

    // Obtener la última fila con datos para agregar el nuevo registro al final
    const ultimaFila = hoja.getLastRow();
    const nuevaFilaNum = ultimaFila + 1;

    Logger.log("📊 Última fila con datos: " + ultimaFila);
    Logger.log("➕ Nuevo registro se agregará en la fila: " + nuevaFilaNum);

    // Copiar formato de la fila anterior si existe (para mantener consistencia)
    try {
      if (ultimaFila > 1) {
        Logger.log("🎨 Iniciando copia de formato...");
        hoja
          .getRange(ultimaFila, 1, 1, Math.min(10, hoja.getLastColumn()))
          .copyTo(hoja.getRange(nuevaFilaNum, 1, 1, 10), {
            formatOnly: true,
          });
        Logger.log(
          "🎨 Formato copiado de fila " + ultimaFila + " a fila " + nuevaFilaNum
        );
      }
    } catch (formatError) {
      Logger.log(
        "⚠️ Error copiando formato (continuando): " + formatError.message
      );
    }

    // Insertar los datos en la nueva fila al final
    Logger.log("💾 Insertando datos en la fila " + nuevaFilaNum + "...");
    hoja.getRange(nuevaFilaNum, 1, 1, fila.length).setValues([fila]);
    Logger.log("✅ Datos insertados en la fila " + nuevaFilaNum);

    // Formatear la celda de fecha
    Logger.log("📅 Aplicando formato de fecha...");
    hoja.getRange(nuevaFilaNum, 1).setNumberFormat("yyyy-mm-dd");
    Logger.log("📅 Formato de fecha aplicado");

    // Convertir las celdas de recibida y enviada en checkboxes
    Logger.log("☑️ Insertando checkboxes...");
    hoja.getRange(nuevaFilaNum, 2, 1, 2).insertCheckboxes();
    Logger.log("✅ Checkboxes insertados en columnas B y C");

    // Verificar que los datos se guardaron correctamente
    Logger.log("🔍 Verificando datos guardados...");
    const filaGuardada = hoja.getRange(nuevaFilaNum, 1, 1, 10).getValues()[0];
    Logger.log(
      "🔍 Verificación - Datos guardados en fila " +
        nuevaFilaNum +
        ": " +
        JSON.stringify(filaGuardada)
    );

    Logger.log(
      "✅ Registro de biopsia guardado exitosamente en la fila " +
        nuevaFilaNum +
        "."
    );
    return true;
  } catch (error) {
    Logger.log("❌ ERROR al guardar biopsia: " + error.message);
    throw new Error("Error al guardar el registro: " + error.message);
  }
}

/**
 * Guarda un registro de biopsia con parámetros individuales (soluciona problemas de serialización).
 * @param {string} fechaToma - Fecha de toma en formato YYYY-MM-DD.
 * @param {string} cedula - Cédula, pasaporte o DIMEX del paciente.
 * @param {string} telefono - Teléfono del paciente.
 * @param {string} nombreCliente - Nombre completo del paciente.
 * @param {number} frascosGastro - Cantidad de frascos gastro.
 * @param {number} frascosColon - Cantidad de frascos colon.
 * @param {string} medico - Nombre del médico.
 * @param {string} comentario - Comentarios adicionales.
 * @returns {boolean} - Verdadero si se guardó correctamente.
 */
function guardarRegistroBiopsiaConParametros(
  fechaToma,
  cedula,
  telefono,
  nombreCliente,
  frascosGastro,
  frascosColon,
  medico,
  comentario
) {
  try {
    Logger.log(
      "🧪 [PARAMS] Iniciando función guardarRegistroBiopsiaConParametros"
    );
    Logger.log("📊 [PARAMS] Parámetros recibidos:");
    Logger.log("   fechaToma: " + fechaToma);
    Logger.log("   cedula: " + cedula);
    Logger.log("   telefono: " + telefono);
    Logger.log("   nombreCliente: " + nombreCliente);
    Logger.log("   frascosGastro: " + frascosGastro);
    Logger.log("   frascosColon: " + frascosColon);
    Logger.log("   medico: " + medico);
    Logger.log("   comentario: " + comentario);

    const hoja =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RegistroBiopsias");
    if (!hoja) {
      throw new Error("La hoja 'RegistroBiopsias' no existe.");
    }

    // Validar datos obligatorios
    if (!nombreCliente || nombreCliente.trim() === "") {
      throw new Error("El nombre del paciente es obligatorio.");
    }

    if (!medico || medico === "") {
      throw new Error("El médico es obligatorio.");
    }

    if (!fechaToma) {
      throw new Error("La fecha de toma es obligatoria.");
    }

    const fechaTomaProcessed = procesarFecha(fechaToma);
    const cedulaLimpia = cedula ? cedula.replace(/[-\s]/g, "").trim() : "";

    // Preparar la fila para añadir
    const fila = [
      fechaTomaProcessed, // A - Fecha Toma
      false, // B - Recibida (checkbox)
      false, // C - Enviada (checkbox)
      cedulaLimpia, // D - Cédula
      telefono || "", // E - Teléfono
      nombreCliente.trim(), // F - Nombre Cliente
      Number(frascosGastro) || 0, // G - Frascos Gastro
      Number(frascosColon) || 0, // H - Frascos Colon
      medico, // I - Médico
      comentario || "", // J - Comentario
    ];

    const ultimaFila = hoja.getLastRow();
    const nuevaFilaNum = ultimaFila + 1;

    Logger.log("💾 [PARAMS] Insertando en fila " + nuevaFilaNum);

    // Insertar datos
    hoja.getRange(nuevaFilaNum, 1, 1, fila.length).setValues([fila]);

    // Aplicar formato
    hoja.getRange(nuevaFilaNum, 1).setNumberFormat("yyyy-mm-dd");
    hoja.getRange(nuevaFilaNum, 2, 1, 2).insertCheckboxes();

    Logger.log("✅ [PARAMS] Registro guardado en fila " + nuevaFilaNum);
    return true;
  } catch (error) {
    Logger.log("❌ [PARAMS] ERROR: " + error.message);
    throw new Error("Error al guardar: " + error.message);
  }
}

/**
 * Versión simplificada para debugging - Guarda un registro sin formato
 * @param {Object} data - Datos del registro de biopsia.
 * @returns {boolean} - Verdadero si se guardó correctamente.
 */
function guardarRegistroBiopsiaSimple(data) {
  try {
    Logger.log("🧪 [SIMPLE] Iniciando función guardarRegistroBiopsiaSimple");
    Logger.log("🔍 [SIMPLE] Datos recibidos: " + JSON.stringify(data));

    const hoja =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RegistroBiopsias");
    if (!hoja) {
      throw new Error("La hoja 'RegistroBiopsias' no existe.");
    }

    // Validar datos básicos
    if (!data.nombre_cliente || !data.medico) {
      throw new Error("Faltan datos obligatorios.");
    }

    const fechaToma = procesarFecha(data.fecha_toma);
    const cedulaLimpia = data.cedula
      ? data.cedula.replace(/[-\s]/g, "").trim()
      : "";

    const fila = [
      fechaToma,
      false,
      false,
      cedulaLimpia,
      data.telefono || "",
      data.nombre_cliente.trim(),
      Number(data.frascos_gastro) || 0,
      Number(data.frascos_colon) || 0,
      data.medico,
      data.comentario || "",
    ];

    const ultimaFila = hoja.getLastRow();
    const nuevaFilaNum = ultimaFila + 1;

    Logger.log("💾 [SIMPLE] Insertando datos en fila " + nuevaFilaNum);
    hoja.getRange(nuevaFilaNum, 1, 1, fila.length).setValues([fila]);

    Logger.log("✅ [SIMPLE] Registro guardado exitosamente");
    return true;
  } catch (error) {
    Logger.log("❌ [SIMPLE] ERROR: " + error.message);
    throw new Error("Error simple: " + error.message);
  }
}

/**
 * Función de debugging ultra-simple - Solo inserta texto sin procesar
 * @param {Object} data - Datos del registro de biopsia.
 * @returns {boolean} - Verdadero si se guardó correctamente.
 */
function guardarRegistroBiopsiaUltraSimple(data) {
  try {
    Logger.log("🧪 [ULTRA-SIMPLE] Iniciando función test");
    Logger.log("🔍 [ULTRA-SIMPLE] Datos: " + JSON.stringify(data));

    const hoja =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RegistroBiopsias");
    if (!hoja) {
      throw new Error("La hoja 'RegistroBiopsias' no existe.");
    }

    Logger.log("📊 [ULTRA-SIMPLE] Hoja encontrada");

    const ultimaFila = hoja.getLastRow();
    const nuevaFilaNum = ultimaFila + 1;

    Logger.log("💾 [ULTRA-SIMPLE] Insertando en fila " + nuevaFilaNum);

    // Datos ultra-simples sin procesamiento
    const fila = [
      data.fecha_toma, // Como string
      "NO", // Como texto simple
      "NO", // Como texto simple
      data.cedula,
      data.telefono,
      data.nombre_cliente,
      data.frascos_gastro,
      data.frascos_colon,
      data.medico,
      data.comentario,
    ];

    hoja.getRange(nuevaFilaNum, 1, 1, fila.length).setValues([fila]);

    Logger.log("✅ [ULTRA-SIMPLE] Completado");
    return true;
  } catch (error) {
    Logger.log("❌ [ULTRA-SIMPLE] ERROR: " + error.message);
    throw new Error("[ULTRA-SIMPLE] Error: " + error.message);
  }
}

/**
 * Procesa una fecha en formato string y la convierte a objeto Date
 * @param {string} fechaStr - Fecha en formato YYYY-MM-DD
 * @returns {Date} - Objeto Date
 */
function procesarFecha(fechaStr) {
  if (!fechaStr || typeof fechaStr !== "string") {
    throw new Error("Formato de fecha inválido");
  }

  try {
    const [year, month, day] = fechaStr.split("-").map(Number);
    return new Date(Date.UTC(year, month - 1, day, 12, 0, 0));
  } catch (e) {
    throw new Error("Formato de fecha inválido: " + e.message);
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
      Logger.log("ERROR: La hoja 'RegistroBiopsias' no existe");
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
        // Normaliza: convierte a string, elimina espacios, guiones y otros caracteres especiales
        // pero conserva letras y números para soportar pasaportes
        const normalizarIdentificacion = (str) =>
          String(str || "")
            .replace(/[-\s]/g, "")
            .toUpperCase();
        const identificacionHoja = normalizarIdentificacion(fila[3]);
        const identificacionBusqueda = normalizarIdentificacion(searchValue);
        Logger.log(
          `Comparando identificación: hoja=${identificacionHoja} vs busqueda=${identificacionBusqueda}`
        );
        Logger.log(
          `Fila ${
            i + 1
          }: identificacionHoja=${identificacionHoja}, identificacionBusqueda=${identificacionBusqueda}`
        );
        if (identificacionHoja === identificacionBusqueda) {
          coincide = true;
        }
      } else if (searchType === "nombre") {
        const nombreHoja = normalizar(fila[5]); // Columna F - nombre_cliente
        const nombreBusqueda = normalizar(searchValue);
        if (nombreHoja.includes(nombreBusqueda)) {
          coincide = true;
        }
      } else if (searchType === "estado") {
        // Determinar estado basado en los checkboxes únicamente
        let estadoMuestra = "registrada";
        const recibida = fila[1]; // Checkbox recibida (columna B)
        const enviada = fila[2]; // Checkbox enviada (columna C)

        if (enviada) {
          estadoMuestra = "enviada";
        } else if (recibida) {
          estadoMuestra = "recibida";
        }

        Logger.log(
          `Fila ${
            i + 1
          }: estado calculado=${estadoMuestra}, buscado=${searchValue}`
        );
        if (estadoMuestra === searchValue) {
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
 * Función mejorada para buscar biopsias con múltiples opciones
 * @param {string} searchType - Tipo de búsqueda: "fecha", "cedula", "nombre", "estado", "mes"
 * @param {string} searchValue - Valor a buscar
 * @param {string} searchValue2 - Valor adicional (para búsquedas de mes/año)
 * @returns {Array} - Resultados de la búsqueda con índice de fila
 */
function buscarBiopsiasServidorMejorado(searchType, searchValue, searchValue2) {
  try {
    // Asegurar que los parámetros no sean undefined
    searchType = searchType || "";
    searchValue = searchValue || "";
    searchValue2 = searchValue2 || null;

    Logger.log(
      `🔍 BÚSQUEDA: tipo=${searchType}, valor1=${searchValue}, valor2=${searchValue2}`
    );
    Logger.log(
      `🔍 BÚSQUEDA TIPOS: tipo=${typeof searchType}, valor1=${typeof searchValue}, valor2=${typeof searchValue2}`
    );

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("RegistroBiopsias");

    if (!hoja) {
      Logger.log("❌ ERROR: La hoja 'RegistroBiopsias' no existe");
      return [];
    }

    const datos = hoja.getDataRange().getValues();
    Logger.log(`📊 Total de filas: ${datos.length}`);

    if (datos.length <= 1) {
      Logger.log("ℹ️ No hay datos para buscar");
      return [];
    }

    const resultados = [];

    // Funciones auxiliares
    const normalizarTexto = (str) =>
      String(str || "")
        .toLowerCase()
        .trim();
    const normalizarIdentificacion = (str) =>
      String(str || "")
        .replace(/[-\s]/g, "")
        .toUpperCase();
    const formatearFecha = (fecha) => {
      if (fecha instanceof Date) {
        return Utilities.formatDate(
          fecha,
          Session.getScriptTimeZone(),
          "yyyy-MM-dd"
        );
      } else if (typeof fecha === "string") {
        const fechaTmp = new Date(fecha);
        return Utilities.formatDate(
          fechaTmp,
          Session.getScriptTimeZone(),
          "yyyy-MM-dd"
        );
      }
      return "";
    };

    // Procesar cada fila (saltar encabezado)
    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      let coincide = false;

      Logger.log(`🔍 Procesando fila ${i} con tipo: ${searchType}`);

      switch (searchType) {
        case "fecha":
          const fechaHoja = formatearFecha(fila[0]);
          const fechaBusqueda = searchValue;
          Logger.log(`📅 Comparando: ${fechaHoja} vs ${fechaBusqueda}`);
          coincide = fechaHoja === fechaBusqueda;
          break;

        case "mes":
          // Búsqueda por mes y año específico
          if (fila[0] instanceof Date || typeof fila[0] === "string") {
            const fecha = new Date(fila[0]);
            const mesHoja = fecha.getMonth() + 1; // getMonth() es 0-based
            const anioHoja = fecha.getFullYear();
            const mesBusqueda = parseInt(searchValue);
            const anioBusqueda = parseInt(searchValue2);

            Logger.log(
              `📅 MES: Comparando ${mesHoja}/${anioHoja} vs ${mesBusqueda}/${anioBusqueda}`
            );
            coincide = mesHoja === mesBusqueda && anioHoja === anioBusqueda;
          }
          break;

        case "cedula":
          const identificacionHoja = normalizarIdentificacion(fila[3]);
          const identificacionBusqueda = normalizarIdentificacion(searchValue);
          Logger.log(
            `🆔 CEDULA DETALLE: Fila ${i}, Hoja="${fila[3]}" -> "${identificacionHoja}", Búsqueda="${searchValue}" -> "${identificacionBusqueda}"`
          );
          coincide = identificacionHoja === identificacionBusqueda;
          if (coincide) {
            Logger.log(`🎯 CEDULA MATCH encontrado en fila ${i + 1}`);
          }
          break;

        case "nombre":
          const nombreHoja = normalizarTexto(fila[5]); // Columna F - nombre_cliente
          const nombreBusqueda = normalizarTexto(searchValue);
          Logger.log(
            `👤 Comparando: "${nombreHoja}" incluye "${nombreBusqueda}"`
          );
          coincide = nombreHoja.includes(nombreBusqueda);
          break;

        case "estado":
          // Determinar estado basado en checkboxes
          let estadoMuestra = "registrada";
          const recibida = Boolean(fila[1]); // Columna B
          const enviada = Boolean(fila[2]); // Columna C

          if (recibida && enviada) {
            estadoMuestra = "completada";
          } else if (enviada) {
            estadoMuestra = "enviada";
          } else if (recibida) {
            estadoMuestra = "recibida";
          }

          Logger.log(
            `📋 Estado calculado: ${estadoMuestra} vs buscado: ${searchValue}`
          );
          coincide = estadoMuestra === searchValue;
          break;

        default:
          Logger.log(`⚠️ Tipo de búsqueda no reconocido: ${searchType}`);
          break;
      }

      if (coincide) {
        // Agregar número de fila al final para edición posterior
        const filaConIndice = [...fila, i + 1];
        resultados.push(filaConIndice);
        Logger.log(`✅ Coincidencia encontrada en fila ${i + 1}`);
      }
    }

    Logger.log(`🎯 Resultados encontrados: ${resultados.length}`);
    return resultados;
  } catch (error) {
    Logger.log(`❌ ERROR en búsqueda: ${error.message}`);
    return [];
  }
}

/**
 * Actualiza un registro completo de biopsia
 * @param {number} fila - Número de fila a actualizar
 * @param {Object} datos - Datos actualizados
 * @returns {boolean} - true si se actualizó correctamente
 */
function actualizarRegistroBiopsia(fila, datos) {
  try {
    Logger.log(`📝 ACTUALIZAR: Fila ${fila}, datos: ${JSON.stringify(datos)}`);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("RegistroBiopsias");

    if (!hoja) {
      throw new Error("Hoja RegistroBiopsias no encontrada");
    }

    // Validar datos obligatorios
    if (!datos.fecha_toma || !datos.nombre_cliente || !datos.medico) {
      throw new Error("Faltan datos obligatorios");
    }

    // Crear la fila de datos actualizada
    const filaActualizada = [
      datos.fecha_toma, // A: fecha_toma
      Boolean(datos.recibida), // B: recibida
      Boolean(datos.enviada), // C: enviada
      datos.cedula || "", // D: cedula
      datos.telefono || "", // E: telefono
      datos.nombre_cliente, // F: nombre_cliente
      Number(datos.frascos_gastro) || 0, // G: frascos_gastro
      Number(datos.frascos_colon) || 0, // H: frascos_colon
      datos.medico, // I: medico
      datos.comentario || "", // J: comentario
    ];

    // Actualizar la fila
    hoja
      .getRange(fila, 1, 1, filaActualizada.length)
      .setValues([filaActualizada]);

    // Aplicar formatos
    hoja.getRange(fila, 1).setNumberFormat("yyyy-mm-dd");
    hoja.getRange(fila, 2, 1, 2).insertCheckboxes();

    // Aplicar coloración si ambos checkboxes están marcados
    if (datos.recibida && datos.enviada) {
      hoja.getRange(fila, 1, 1, 8).setBackground("#FFFF99"); // Amarillo
    } else {
      hoja.getRange(fila, 1, 1, 8).setBackground(null); // Sin color
    }

    Logger.log(
      `✅ ACTUALIZAR: Registro actualizado correctamente en fila ${fila}`
    );
    return true;
  } catch (error) {
    Logger.log(`❌ ACTUALIZAR: Error: ${error.message}`);
    throw error;
  }
}

/**
 * Actualiza solo el estado de los checkboxes de una biopsia
 * @param {number} fila - Número de fila a actualizar
 * @param {boolean} recibida - Estado del checkbox "Recibida"
 * @param {boolean} enviada - Estado del checkbox "Enviada"
 * @returns {boolean} - true si se actualizó correctamente
 */
function actualizarEstadosBiopsia(fila, recibida, enviada) {
  try {
    Logger.log(
      `📋 ESTADOS: Fila ${fila}, recibida: ${recibida}, enviada: ${enviada}`
    );

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("RegistroBiopsias");

    if (!hoja) {
      throw new Error("Hoja RegistroBiopsias no encontrada");
    }

    // Actualizar checkboxes
    hoja.getRange(fila, 2).setValue(Boolean(recibida)); // Columna B
    hoja.getRange(fila, 3).setValue(Boolean(enviada)); // Columna C

    // Aplicar coloración si ambos están marcados
    if (recibida && enviada) {
      hoja.getRange(fila, 1, 1, 8).setBackground("#FFFF99"); // Amarillo
      Logger.log(`🎨 ESTADOS: Aplicada coloración amarilla a fila ${fila}`);
    } else {
      hoja.getRange(fila, 1, 1, 8).setBackground(null); // Sin color
      Logger.log(`🎨 ESTADOS: Removida coloración de fila ${fila}`);
    }

    Logger.log(
      `✅ ESTADOS: Estados actualizados correctamente en fila ${fila}`
    );
    return true;
  } catch (error) {
    Logger.log(`❌ ESTADOS: Error: ${error.message}`);
    throw error;
  }
}

/**
 * Obtiene los datos de una fila específica para edición
 * @param {number} fila - Número de fila a obtener
 * @returns {Object} - Datos de la fila
 */
function obtenerDatosBiopsia(fila) {
  try {
    Logger.log(`📖 OBTENER: Fila ${fila}`);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("RegistroBiopsias");

    if (!hoja) {
      throw new Error("Hoja RegistroBiopsias no encontrada");
    }

    const datos = hoja.getRange(fila, 1, 1, 10).getValues()[0];

    const resultado = {
      fecha_toma: datos[0],
      recibida: Boolean(datos[1]),
      enviada: Boolean(datos[2]),
      cedula: datos[3] || "",
      telefono: datos[4] || "",
      nombre_cliente: datos[5] || "",
      frascos_gastro: Number(datos[6]) || 0,
      frascos_colon: Number(datos[7]) || 0,
      medico: datos[8] || "",
      comentario: datos[9] || "",
      fila: fila,
    };

    Logger.log(`✅ OBTENER: Datos obtenidos: ${JSON.stringify(resultado)}`);
    return resultado;
  } catch (error) {
    Logger.log(`❌ OBTENER: Error: ${error.message}`);
    throw error;
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
  const sheet = ss.getSheetByName("RegistroBiopsias");

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
    throw new Error("La hoja 'RegistroBiopsias' no existe.");
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

// ================== FUNCIONES DE BIOPSIAS (NUEVAS) ==================

/**
 * Muestra el formulario para registrar biopsias
 */
function mostrarRegistroBiopsias() {
  const html = HtmlService.createTemplateFromFile("registroBiopsias")
    .evaluate()
    .setWidth(800)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, "Registro de Biopsias");
}

/**
 * Función simple y limpia para guardar biopsias
 * @param {Object} datos - Datos del formulario
 * @returns {boolean} - true si se guardó correctamente
 */
function guardarBiopsia(datos) {
  try {
    Logger.log("🧪 NUEVO: Guardando biopsia");
    Logger.log("📊 NUEVO: Datos recibidos: " + JSON.stringify(datos));

    // Verificar que recibimos datos
    if (!datos) {
      throw new Error("No se recibieron datos");
    }

    // Obtener la hoja
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log("📊 NUEVO: Spreadsheet obtenido: " + spreadsheet.getName());

    const hoja = spreadsheet.getSheetByName("RegistroBiopsias");
    if (!hoja) {
      Logger.log("❌ NUEVO: Hoja RegistroBiopsias no encontrada");
      throw new Error("Hoja RegistroBiopsias no encontrada");
    }

    Logger.log("📊 NUEVO: Hoja encontrada: " + hoja.getName());

    // Validar datos obligatorios
    if (!datos.fecha_toma || !datos.nombre_cliente || !datos.medico) {
      const error =
        "Faltan datos obligatorios: " +
        (!datos.fecha_toma ? "fecha_toma " : "") +
        (!datos.nombre_cliente ? "nombre_cliente " : "") +
        (!datos.medico ? "medico " : "");
      Logger.log("❌ NUEVO: " + error);
      throw new Error(error);
    }

    // Crear la fila de datos
    const fila = [
      datos.fecha_toma, // A: fecha_toma
      false, // B: recibida
      false, // C: enviada
      datos.cedula || "", // D: cedula
      datos.telefono || "", // E: telefono
      datos.nombre_cliente, // F: nombre_cliente
      Number(datos.frascos_gastro) || 0, // G: frascos_gastro
      Number(datos.frascos_colon) || 0, // H: frascos_colon
      datos.medico, // I: medico
      datos.comentario || "", // J: comentario
    ];

    Logger.log("📊 NUEVO: Fila a insertar: " + JSON.stringify(fila));

    // Insertar en la última fila
    const ultimaFila = hoja.getLastRow();
    const nuevaFila = ultimaFila + 1;

    Logger.log("📊 NUEVO: Insertando en fila: " + nuevaFila);

    hoja.getRange(nuevaFila, 1, 1, fila.length).setValues([fila]);
    Logger.log("📊 NUEVO: Datos insertados correctamente");

    // Aplicar formatos
    hoja.getRange(nuevaFila, 1).setNumberFormat("yyyy-mm-dd");
    hoja.getRange(nuevaFila, 2, 1, 2).insertCheckboxes();
    Logger.log("📊 NUEVO: Formatos aplicados");

    Logger.log("✅ NUEVO: Guardado exitoso en fila " + nuevaFila);
    return {
      success: true,
      message: "Biopsia guardada correctamente",
      fila: nuevaFila,
    };
  } catch (error) {
    Logger.log("❌ NUEVO: Error completo: " + error.toString());
    Logger.log("❌ NUEVO: Stack: " + error.stack);
    throw new Error("Error al guardar biopsia: " + error.message);
  }
}

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
      if (personal && typeof personal === "string" && personal.trim()) {
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
      filasConDatos: datos.length - 1, // Sin contar encabezados
    };

    Logger.log("✅ Estructura verificada: " + JSON.stringify(resultado));
    return resultado;
  } catch (error) {
    Logger.log("❌ Error en verificarEstructuraHoja: " + error.message);
    return {
      error: error.message,
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

    if (!nombreBuscar || nombreBuscar.trim() === "") {
      return {
        coincidenciasExactas: [],
        coincidenciasParciales: [],
        sugerencias: [],
      };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaRegistros = ss.getSheetByName("RegistrosProcedimientos");
    const hojaPersonal = ss.getSheetByName("Personal");

    if (!hojaRegistros || !hojaPersonal) {
      throw new Error("No se encontraron las hojas necesarias");
    }

    function normalizarNombre(str) {
      return String(str || "")
        .replace(/\s+/g, " ")
        .trim()
        .toUpperCase();
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

    todosLosNombres.forEach((nombre) => {
      const nombreNorm = normalizarNombre(nombre);

      // Coincidencia exacta
      if (nombreNorm === nombreBuscarNorm) {
        coincidenciasExactas.push(nombre);
        return;
      }

      // Coincidencia parcial (uno contiene al otro)
      if (
        nombreNorm.includes(nombreBuscarNorm) ||
        nombreBuscarNorm.includes(nombreNorm)
      ) {
        coincidenciasParciales.push(nombre);
        return;
      }

      // Sugerencias (palabras similares)
      const palabrasBusqueda = nombreBuscarNorm
        .split(" ")
        .filter((p) => p.length > 2);
      const palabrasNombre = nombreNorm.split(" ").filter((p) => p.length > 2);

      if (palabrasBusqueda.length > 0 && palabrasNombre.length > 0) {
        const coincidencias = palabrasBusqueda.filter((p1) =>
          palabrasNombre.some((p2) => p1.includes(p2) || p2.includes(p1))
        );

        if (
          coincidencias.length >=
          Math.min(palabrasBusqueda.length, palabrasNombre.length) * 0.4
        ) {
          sugerencias.push(nombre);
        }
      }
    });

    const resultado = {
      coincidenciasExactas: coincidenciasExactas.slice(0, 10),
      coincidenciasParciales: coincidenciasParciales.slice(0, 10),
      sugerencias: sugerencias.slice(0, 5),
    };

    Logger.log("Búsqueda completada: " + JSON.stringify(resultado));
    return resultado;
  } catch (error) {
    Logger.log("Error en buscarPersonalPorNombre: " + error.message);
    throw new Error("Error al buscar personal: " + error.message);
  }
}

/**
 * Función de diagnóstico específica para "Anest Manuel"
 */
function diagnosticarAnestManuel() {
  Logger.log("🔍 DIAGNÓSTICO ESPECÍFICO PARA 'ANEST MANUEL'");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaPersonal = ss.getSheetByName("Personal");

  if (!hojaPersonal) {
    Logger.log("❌ La hoja 'Personal' no existe.");
    return;
  }

  const datos = hojaPersonal.getDataRange().getValues();
  const nombreBuscado = "ANEST MANUEL";

  Logger.log(`📋 Buscando "${nombreBuscado}" en hoja Personal...`);
  Logger.log(`📊 Total de filas: ${datos.length}`);

  // Mostrar toda la columna B (índice 1) donde deberían estar los anestesiólogos
  Logger.log(`\n🏥 CONTENIDO COMPLETO DE COLUMNA B (ANESTESIÓLOGOS):`);
  for (let i = 0; i < datos.length; i++) {
    const valorB = String(datos[i][1] || "").trim();
    const valorBNorm = valorB.toUpperCase();
    if (valorB !== "") {
      Logger.log(
        `   Fila ${
          i + 1
        }: "${valorB}" -> normalizado: "${valorBNorm}" | Coincide: ${
          valorBNorm === nombreBuscado
        }`
      );
    }
  }

  // Buscar en todas las columnas por si acaso
  Logger.log(`\n🔍 BÚSQUEDA EN TODAS LAS COLUMNAS:`);
  let encontrado = false;

  for (let i = 0; i < datos.length; i++) {
    const fila = datos[i];
    for (let col = 0; col <= 5; col++) {
      const valor = String(fila[col] || "").trim();
      const valorNorm = valor.toUpperCase();

      if (valorNorm.includes("MANUEL") || valorNorm.includes("ANEST")) {
        Logger.log(
          `   Fila ${i + 1}, Col ${col}: "${valor}" -> "${valorNorm}"`
        );
        if (valorNorm === nombreBuscado) {
          encontrado = true;
          Logger.log(`   ✅ ENCONTRADO EN FILA ${i + 1}, COLUMNA ${col}!`);
        }
      }
    }
  }

  if (!encontrado) {
    Logger.log(`❌ NO SE ENCONTRÓ "${nombreBuscado}" en ninguna celda`);
  }

  // Probar la función obtenerTipoDePersonal
  Logger.log(`\n🧪 PROBANDO obtenerTipoDePersonal("Anest Manuel"):`);
  const resultado = obtenerTipoDePersonal("Anest Manuel", hojaPersonal);
  Logger.log(`📋 Resultado: ${resultado}`);

  // Probar también calcularCostos
  Logger.log(
    `\n🧮 PROBANDO calcularCostos("Anest Manuel", 7, 2025, "completo"):`
  );
  try {
    const resultadoCostos = calcularCostos("Anest Manuel", 7, 2025, "completo");
    Logger.log(`💰 Resultado calcularCostos:`);
    Logger.log(`   - Subtotal: $${resultadoCostos.totales.subtotal}`);
    Logger.log(`   - IVA: $${resultadoCostos.totales.iva}`);
    Logger.log(`   - Total: $${resultadoCostos.totales.total_con_iva}`);
    Logger.log(`   - Es null?: ${resultadoCostos === null}`);
  } catch (error) {
    Logger.log(`❌ Error en calcularCostos: ${error.message}`);
  }

  return {
    encontrado: encontrado,
    tipoPersonal: resultado,
  };
}

/**
 * Función para obtener solo anestesiólogos para reportes específicos
 */
function obtenerAnestesiologos() {
  try {
    const hoja =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Personal");
    if (!hoja) {
      throw new Error("❌ La hoja 'Personal' no existe.");
    }

    const datos = hoja.getDataRange().getValues();
    const anestesiologos = new Set();

    // Solo columna 1 (Anestesiólogos)
    for (let i = 1; i < datos.length; i++) {
      const nombre = datos[i][1]; // Columna B
      if (nombre && typeof nombre === "string" && nombre.trim()) {
        anestesiologos.add(nombre.trim());
      }
    }

    const resultado = [...anestesiologos].sort();
    Logger.log("Anestesiólogos encontrados: " + resultado.join(", "));
    return resultado;
  } catch (error) {
    Logger.log("ERROR en obtenerAnestesiologos: " + error.message);
    return [];
  }
}

/**
 * Función para obtener solo doctores para reportes específicos
 */
function obtenerDoctores() {
  try {
    Logger.log("🔍 INICIANDO obtenerDoctores()");

    const hoja =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Personal");
    if (!hoja) {
      Logger.log("❌ La hoja 'Personal' no existe.");
      throw new Error("❌ La hoja 'Personal' no existe.");
    }

    Logger.log("✅ Hoja 'Personal' encontrada");
    const datos = hoja.getDataRange().getValues();
    Logger.log(`📊 Datos obtenidos. Total filas: ${datos.length}`);

    if (datos.length <= 1) {
      Logger.log("⚠️ No hay datos en la hoja Personal o solo hay encabezados");
      return [];
    }

    const doctores = new Set();

    // Solo columna 0 (Doctores)
    for (let i = 1; i < datos.length; i++) {
      const nombre = datos[i][0]; // Columna A
      Logger.log(`   Fila ${i + 1}: "${nombre}" (tipo: ${typeof nombre})`);

      if (nombre && typeof nombre === "string" && nombre.trim()) {
        const nombreLimpio = nombre.trim();
        doctores.add(nombreLimpio);
        Logger.log(`   ✅ Agregado doctor: "${nombreLimpio}"`);
      } else {
        Logger.log(`   ⏭️ Saltando fila ${i + 1}: valor inválido`);
      }
    }

    const resultado = [...doctores].sort();
    Logger.log(`🎯 RESULTADO: ${resultado.length} doctores encontrados`);
    Logger.log("📋 Lista final: " + resultado.join(", "));
    return resultado;
  } catch (error) {
    Logger.log("❌ ERROR en obtenerDoctores: " + error.message);
    Logger.log("Stack trace: " + error.stack);
    return [];
  }
}

/**
 * Función para obtener solo técnicos para reportes específicos
 */
function obtenerTecnicos() {
  try {
    const hoja =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Personal");
    if (!hoja) {
      throw new Error("❌ La hoja 'Personal' no existe.");
    }

    const datos = hoja.getDataRange().getValues();
    const tecnicos = new Set();

    // Solo columna 2 (Técnicos)
    for (let i = 1; i < datos.length; i++) {
      const nombre = datos[i][2]; // Columna C
      if (nombre && typeof nombre === "string" && nombre.trim()) {
        tecnicos.add(nombre.trim());
      }
    }

    const resultado = [...tecnicos].sort();
    Logger.log("Técnicos encontrados: " + resultado.join(", "));
    return resultado;
  } catch (error) {
    Logger.log("ERROR en obtenerTecnicos: " + error.message);
    return [];
  }
}

// ================= FUNCIONES MEJORADAS AGREGADAS =================

/**
 * Función mejorada para normalizar nombres - elimina caracteres especiales y unifica formato
 */
function normalizarNombreMejorado(nombre) {
  if (!nombre) return "";

  return String(nombre)
    .trim()
    .replace(/\s+/g, " ") // Múltiples espacios a uno solo
    .toUpperCase()
    .replace(/[ÁÀÄÂ]/g, "A")
    .replace(/[ÉÈËÊ]/g, "E")
    .replace(/[ÍÌÏÎ]/g, "I")
    .replace(/[ÓÒÖÔ]/g, "O")
    .replace(/[ÚÙÜÛ]/g, "U")
    .replace(/Ñ/g, "N");
}

/**
 * Función mejorada para comparar nombres de personal
 */
function sonLaMismaPersonaMejorado(nombre1, nombre2) {
  if (!nombre1 || !nombre2) return false;

  const norm1 = normalizarNombreMejorado(nombre1);
  const norm2 = normalizarNombreMejorado(nombre2);

  // 1. Coincidencia exacta
  if (norm1 === norm2) {
    return true;
  }

  // 2. Uno contiene al otro (para nombres parciales como "MANUEL" vs "ANEST MANUEL")
  if (norm1.includes(norm2) || norm2.includes(norm1)) {
    return true;
  }

  // 3. Comparación por palabras clave (para nombres compuestos)
  const palabras1 = norm1.split(" ").filter((p) => p.length > 2);
  const palabras2 = norm2.split(" ").filter((p) => p.length > 2);

  if (palabras1.length > 0 && palabras2.length > 0) {
    const coincidencias = palabras1.filter((p1) =>
      palabras2.some((p2) => p1.includes(p2) || p2.includes(p1))
    );

    // Si al menos 50% de las palabras coinciden
    const porcentajeCoincidencia =
      coincidencias.length / Math.min(palabras1.length, palabras2.length);
    return porcentajeCoincidencia >= 0.5;
  }

  return false;
}

/**
 * Función auxiliar para verificar si una fecha está en el rango solicitado
 */
function estaEnRangoFechaMejorado(fecha, mes, anio, quincena) {
  try {
    if (!(fecha instanceof Date) || isNaN(fecha.getTime())) {
      return false;
    }

    const anioFecha = fecha.getFullYear();
    const mesFecha = fecha.getMonth() + 1; // getMonth() devuelve 0-11
    const diaFecha = fecha.getDate();

    // Verificar año y mes
    if (anioFecha !== parseInt(anio) || mesFecha !== parseInt(mes)) {
      return false;
    }

    // Verificar quincena
    switch (quincena) {
      case "1-15":
        return diaFecha >= 1 && diaFecha <= 15;
      case "16-31":
        return diaFecha >= 16;
      case "completo":
        return true;
      default:
        Logger.log(`⚠️ Quincena no reconocida: ${quincena}`);
        return false;
    }
  } catch (error) {
    Logger.log(`❌ Error en estaEnRangoFechaMejorado: ${error.message}`);
    return false;
  }
}

// ===== NUEVA FUNCIÓN DE DIAGNÓSTICO PARA FILTRADO ESPECÍFICO =====
function diagnosticarFiltradoPersonal(nombreSeleccionado) {
  Logger.log(
    `\n🔍 === DIAGNÓSTICO DE FILTRADO PARA: "${nombreSeleccionado}" ===`
  );

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaRegistros = ss.getSheetByName("RegistrosProcedimientos");

  if (!hojaRegistros) {
    Logger.log("❌ No se encontró la hoja RegistrosProcedimientos");
    return;
  }

  const datos = hojaRegistros.getDataRange().getValues();
  const nombreNormalizado = normalizarNombreMejorado(nombreSeleccionado);

  Logger.log(`📋 Nombre normalizado: "${nombreNormalizado}"`);
  Logger.log(`📊 Total de filas a revisar: ${datos.length - 1}`);

  // Función interna para pruebas
  function normalizarNombrePrueba(str) {
    return String(str || "")
      .replace(/\s+/g, " ")
      .trim()
      .toUpperCase();
  }

  function esLaMismaPersonaPrueba(nombreSel, nombreReg) {
    if (!nombreSel || !nombreReg) return false;

    const selNorm = normalizarNombrePrueba(nombreSel);
    const regNorm = normalizarNombrePrueba(nombreReg);

    // Solo coincidencia exacta
    if (selNorm === regNorm) return true;

    const palabrasSel = selNorm.split(" ").filter((p) => p.length > 1);
    const palabrasReg = regNorm.split(" ").filter((p) => p.length > 1);

    // Solo si el registro tiene las mismas palabras o menos que la selección
    if (palabrasReg.length <= palabrasSel.length) {
      return palabrasReg.every((palabraReg) =>
        palabrasSel.some((palabraSel) => palabraSel === palabraReg)
      );
    }

    return false;
  }

  let coincidenciasEncontradas = 0;
  let falsosPositivos = [];

  // Revisar cada fila
  for (let i = 1; i < datos.length; i++) {
    const nombreEnRegistro = String(datos[i][1] || "").trim();

    if (nombreEnRegistro) {
      const coincide = esLaMismaPersonaPrueba(
        nombreSeleccionado,
        nombreEnRegistro
      );

      if (coincide) {
        coincidenciasEncontradas++;
        const fecha = datos[i][0];
        Logger.log(
          `✅ COINCIDENCIA #${coincidenciasEncontradas}: "${nombreEnRegistro}" en fila ${
            i + 1
          } (${fecha})`
        );

        // Verificar si es un falso positivo
        if (
          nombreEnRegistro.toUpperCase() !== nombreSeleccionado.toUpperCase()
        ) {
          const esValidoComoVariante = nombreEnRegistro
            .toUpperCase()
            .includes(nombreSeleccionado.toUpperCase());
          if (!esValidoComoVariante) {
            falsosPositivos.push({
              fila: i + 1,
              nombre: nombreEnRegistro,
              fecha: fecha,
            });
          }
        }
      }
    }
  }

  Logger.log(`\n📊 RESUMEN DEL DIAGNÓSTICO:`);
  Logger.log(`   ✅ Total de coincidencias: ${coincidenciasEncontradas}`);
  Logger.log(`   ⚠️ Posibles falsos positivos: ${falsosPositivos.length}`);

  if (falsosPositivos.length > 0) {
    Logger.log(`\n⚠️ FALSOS POSITIVOS DETECTADOS:`);
    falsosPositivos.forEach((fp) => {
      Logger.log(`   Fila ${fp.fila}: "${fp.nombre}" (${fp.fecha})`);
    });
  }

  if (coincidenciasEncontradas === 0) {
    Logger.log(`\n🔍 BÚSQUEDA ALTERNATIVA (nombres similares):`);
    for (let i = 1; i < Math.min(datos.length, 20); i++) {
      const nombreReg = String(datos[i][1] || "")
        .trim()
        .toUpperCase();
      if (
        nombreReg.includes("ANEST") ||
        nombreReg.includes(nombreNormalizado.split(" ")[0])
      ) {
        Logger.log(`   Fila ${i + 1}: "${datos[i][1]}" (${datos[i][0]})`);
      }
    }
  }

  return {
    totalCoincidencias: coincidenciasEncontradas,
    falsosPositivos: falsosPositivos.length,
    detallesFalsosPositivos: falsosPositivos,
  };
}

// ===== FUNCIÓN SIMPLE DE PRUEBA =====
function probarConexion() {
  Logger.log("🧪 Función de prueba ejecutada correctamente");

  const objetoPrueba = {
    mensaje: "Conexión exitosa",
    timestamp: new Date().toISOString(),
    test: true,
    detalleRegistros: [
      {
        fila: 1,
        fecha: "Prueba",
        procedimientos: [],
      },
    ],
  };

  // Verificar que es JSON válido
  try {
    const jsonString = JSON.stringify(objetoPrueba);
    Logger.log(`✅ JSON válido: ${jsonString.length} caracteres`);
    return objetoPrueba;
  } catch (error) {
    Logger.log(`❌ Error en JSON: ${error.message}`);
    return { error: "Error en serialización JSON" };
  }
}

// ===== FUNCIÓN DE PRUEBA PARA calcularCostos =====
function probarCalcularCostos() {
  Logger.log("🧪 Probando calcularCostos con datos de prueba...");

  try {
    const resultado = calcularCostos("Test User", 1, 2024, "1-15");
    Logger.log("✅ calcularCostos ejecutado sin errores");
    Logger.log("Estructura del resultado:");
    Logger.log(JSON.stringify(resultado, null, 2));
    return resultado;
  } catch (error) {
    Logger.log("❌ Error en calcularCostos: " + error.message);
    Logger.log("Stack trace: " + error.stack);
    throw error;
  }
}

// ===== NUEVA FUNCIÓN: Probar con datos reales =====
function probarConDatosReales() {
  Logger.log("🧪 Probando calcularCostos con datos reales...");

  // Usar nombres que vemos en los logs
  const nombresReales = [
    "Dra Ivannia Chavarria Soto",
    "Anest Manuel",
    "Anest Nicole",
  ];

  nombresReales.forEach((nombre) => {
    try {
      Logger.log(`\n🔍 Probando con: "${nombre}"`);
      const resultado = calcularCostos(nombre, 7, 2025, "completo");
      Logger.log(`✅ Resultado para ${nombre}:`);
      Logger.log(
        `   - Registros encontrados: ${resultado.detalleRegistros.length}`
      );
      Logger.log(`   - Total: $${resultado.totales.total_con_iva}`);
      Logger.log(
        `   - Estructura JSON válida: ${JSON.stringify(resultado).length} chars`
      );
    } catch (error) {
      Logger.log(`❌ Error con ${nombre}: ${error.message}`);
    }
  });

  return "Prueba completada - revisa los logs";
}

// ===== SOLUCIÓN: calcularCostos SIN detalleRegistros para evitar límites =====
function calcularCostosSimple(nombre, mes, anio, quincena) {
  Logger.log(
    `🚀 INICIANDO calcularCostosSimple: nombre=${nombre}, mes=${mes}, anio=${anio}, quincena=${quincena}`
  );

  try {
    // Llamar a la función original para obtener el resultado completo
    const resultadoCompleto = calcularCostos(nombre, mes, anio, quincena);

    // Crear una versión sin detalleRegistros para evitar límites de tamaño
    const resultadoSimple = {
      lv: resultadoCompleto.lv,
      sab: resultadoCompleto.sab,
      totales: resultadoCompleto.totales,
      // NO incluir detalleRegistros para evitar problema de tamaño
      metadatos: {
        numeroRegistros: resultadoCompleto.detalleRegistros
          ? resultadoCompleto.detalleRegistros.length
          : 0,
        nombre: nombre,
        periodo: `${mes}/${anio}`,
        quincena: quincena,
      },
    };

    Logger.log(
      `✅ Resultado simple generado - ${
        JSON.stringify(resultadoSimple).length
      } caracteres`
    );
    return resultadoSimple;
  } catch (error) {
    Logger.log(`❌ Error en calcularCostosSimple: ${error.message}`);
    throw error;
  }
}

// ===== FUNCIÓN SEPARADA OPTIMIZADA para obtener solo los detalles de registros =====
function obtenerDetalleRegistros(nombre, mes, anio, quincena) {
  Logger.log(
    `🔍 OBTENIENDO detalles OPTIMIZADO para: ${nombre}, ${mes}/${anio}, ${quincena}`
  );

  try {
    // ===== OBTENER SHEETS =====
    const sheetRegistros = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      "RegistrosProcedimientos"
    );
    const registrosData = sheetRegistros.getDataRange().getValues();

    Logger.log(
      `📊 Total de filas en RegistrosProcedimientos: ${registrosData.length}`
    );

    // ===== NORMALIZAR NOMBRE =====
    const nombreNormalizado = nombre.toUpperCase();
    Logger.log(`🔍 Nombre normalizado para búsqueda: "${nombreNormalizado}"`);

    const detalleRegistros = [];
    let registrosEncontrados = 0;

    // ===== BUSCAR SOLO REGISTROS QUE COINCIDAN =====
    for (let i = 1; i < registrosData.length; i++) {
      const fila = registrosData[i];
      const fechaRegistro = fila[0]; // La fecha está en columna 0
      const nombrePersona = fila[1]; // El nombre está en columna 1

      if (!nombrePersona || !fechaRegistro) continue;

      // Normalizar nombre del registro
      const nombreRegistroNormalizado = nombrePersona.toString().toUpperCase();

      // Verificar coincidencia exacta del nombre
      if (nombreRegistroNormalizado !== nombreNormalizado) continue;

      // Verificar fecha
      const fecha = new Date(fechaRegistro);
      if (
        fecha.getMonth() + 1 !== parseInt(mes) ||
        fecha.getFullYear() !== parseInt(anio)
      )
        continue;

      // Verificar quincena si no es "completo"
      if (quincena !== "completo") {
        const dia = fecha.getDate();
        if (quincena === "1-15" && dia > 15) continue;
        if (quincena === "16-31" && dia <= 15) continue;
      }

      registrosEncontrados++;
      Logger.log(
        `✅ REGISTRO VÁLIDO #${registrosEncontrados} - Fila ${
          i + 1
        }: ${nombrePersona}, ${fecha.toDateString()}`
      );

      // Solo crear el detalle básico para este registro
      const esSabado = fecha.getDay() === 6;
      const fechaFormateada = fecha.toLocaleDateString("es-ES", {
        weekday: "long",
        year: "numeric",
        month: "long",
        day: "numeric",
      });

      detalleRegistros.push({
        fila: i + 1,
        fecha: fechaFormateada,
        fechaOriginal: fecha.toISOString(),
        persona: nombrePersona,
        esSabado: esSabado,
        // Simplificar procedimientos para reducir tamaño
        procedimientosResumen: `${registrosEncontrados} registro(s) procesado(s)`,
      });
    }

    const resultado = {
      detalleRegistros: detalleRegistros,
      metadatos: {
        nombre: nombre,
        periodo: `${mes}/${anio}`,
        quincena: quincena,
        totalRegistros: detalleRegistros.length,
      },
    };

    const tamanoJson = JSON.stringify(resultado).length;
    Logger.log(
      `✅ Detalles OPTIMIZADOS obtenidos - ${detalleRegistros.length} registros, ${tamanoJson} caracteres`
    );

    return resultado;
  } catch (error) {
    Logger.log(`❌ Error obteniendo detalles: ${error.message}`);
    throw error;
  }
}

// ===== FUNCIÓN MÍNIMA para probar comunicación de detalles =====
function obtenerDetalleRegistrosMinimo(nombre, mes, anio, quincena) {
  Logger.log(
    `🔍 OBTENIENDO detalles MÍNIMO para: ${nombre}, ${mes}/${anio}, ${quincena}`
  );

  try {
    const resultado = {
      detalleRegistros: [
        {
          fila: 1,
          fecha: "Prueba de comunicación",
          persona: nombre,
          esSabado: false,
        },
      ],
      metadatos: {
        nombre: nombre,
        periodo: `${mes}/${anio}`,
        quincena: quincena,
        totalRegistros: 1,
        mensaje: "Función de prueba - comunicación exitosa",
      },
    };

    const tamanoJson = JSON.stringify(resultado).length;
    Logger.log(`✅ Detalles MÍNIMOS generados - ${tamanoJson} caracteres`);

    return resultado;
  } catch (error) {
    Logger.log(`❌ Error en detalles mínimos: ${error.message}`);
    throw error;
  }
}

// ===== FUNCIÓN HÍBRIDA: Obtener detalles usando datos del resumen =====
function obtenerDetalleRegistrosHibrido(nombre, mes, anio, quincena) {
  Logger.log(
    `🔄 OBTENIENDO detalles HÍBRIDO para: ${nombre}, ${mes}/${anio}, ${quincena}`
  );

  try {
    // Obtener el resumen simple primero
    const resumenSimple = calcularCostosSimple(nombre, mes, anio, quincena);

    if (
      !resumenSimple ||
      !resumenSimple.metadatos ||
      resumenSimple.metadatos.numeroRegistros === 0
    ) {
      Logger.log(
        `⚠️ No hay registros para ${nombre} en el período especificado`
      );
      return {
        detalleRegistros: [],
        metadatos: {
          nombre: nombre,
          periodo: `${mes}/${anio}`,
          quincena: quincena,
          totalRegistros: 0,
          mensaje: "No se encontraron registros",
        },
      };
    }

    // Obtener solo la información básica de las filas de registros
    const sheetRegistros = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      "RegistrosProcedimientos"
    );
    const registrosData = sheetRegistros.getDataRange().getValues();

    const nombreNormalizado = nombre.toUpperCase();
    const detalleRegistros = [];
    let registrosEncontrados = 0;

    // Buscar solo las filas que coincidan para obtener fechas y información básica
    for (let i = 1; i < registrosData.length; i++) {
      const fila = registrosData[i];
      const nombrePersona = fila[0];
      const fechaRegistro = fila[1];

      if (!nombrePersona || !fechaRegistro) continue;

      const nombreRegistroNormalizado = nombrePersona.toString().toUpperCase();
      if (nombreRegistroNormalizado !== nombreNormalizado) continue;

      const fecha = new Date(fechaRegistro);
      if (
        fecha.getMonth() + 1 !== parseInt(mes) ||
        fecha.getFullYear() !== parseInt(anio)
      )
        continue;

      if (quincena !== "completo") {
        const dia = fecha.getDate();
        if (quincena === "1-15" && dia > 15) continue;
        if (quincena === "16-31" && dia <= 15) continue;
      }

      registrosEncontrados++;
      const esSabado = fecha.getDay() === 6;
      const fechaFormateada = fecha.toLocaleDateString("es-ES", {
        weekday: "long",
        year: "numeric",
        month: "long",
        day: "numeric",
      });

      // Crear información básica de procedimientos desde el resumen
      const procedimientosBasicos = [];
      const datosCategoria = esSabado ? resumenSimple.sab : resumenSimple.lv;

      // Solo incluir los procedimientos más importantes para reducir tamaño
      const procedimientosImportantes = [
        "gastro_regular",
        "gastro_medismart",
        "gastocolono_regular",
        "consulta_medismart",
      ];

      procedimientosImportantes.forEach((key) => {
        const item = datosCategoria[key];
        if (item && item.cantidad > 0) {
          procedimientosBasicos.push({
            nombre: item.nombre,
            cantidad: Math.round(item.cantidad / registrosEncontrados), // Aproximar por registro
            costoUnitario: item.costo_unitario,
            total: Math.round(item.total_con_iva / registrosEncontrados), // Aproximar por registro
          });
        }
      });

      detalleRegistros.push({
        fila: i + 1,
        fecha: fechaFormateada,
        fechaOriginal: fecha.toISOString(),
        persona: nombrePersona,
        esSabado: esSabado,
        procedimientos: procedimientosBasicos,
        subtotalRegistro: Math.round(
          resumenSimple.totales.total_con_iva / registrosEncontrados
        ),
      });
    }

    const resultado = {
      detalleRegistros: detalleRegistros,
      metadatos: {
        nombre: nombre,
        periodo: `${mes}/${anio}`,
        quincena: quincena,
        totalRegistros: detalleRegistros.length,
        mensaje: "Detalles híbridos generados correctamente",
      },
    };

    const tamanoJson = JSON.stringify(resultado).length;
    Logger.log(
      `✅ Detalles HÍBRIDOS obtenidos - ${detalleRegistros.length} registros, ${tamanoJson} caracteres`
    );

    return resultado;
  } catch (error) {
    Logger.log(`❌ Error en detalles híbridos: ${error.message}`);
    throw error;
  }
}

// ===== FUNCIÓN SIMPLIFICADA: Obtener detalles básicos con lógica idéntica a calcularCostos =====
// 🔍 FUNCIÓN DE DEBUGGING PARA ENTENDER EL PROBLEMA CON DETALLES
function diagnosticarDetalleRegistros(nombre, mes, anio, quincena) {
  Logger.log(
    `🔍 DIAGNÓSTICO DETALLADO para: ${nombre}, ${mes}/${anio}, ${quincena}`
  );

  try {
    const sheetRegistros = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      "RegistrosProcedimientos"
    );
    const registrosData = sheetRegistros.getDataRange().getValues();

    Logger.log(
      `📊 Total de filas en RegistrosProcedimientos: ${registrosData.length}`
    );

    const nombreNormalizado = nombre.toUpperCase();
    Logger.log(`🔍 Nombre normalizado: "${nombreNormalizado}"`);

    let registrosEncontrados = 0;
    let registrosConFechaCorrecta = 0;
    let registrosConQuincenaCorrecta = 0;
    let registrosConNombreCoincidente = 0;
    let ejemplosNombres = [];

    // Analizar cada fila para diagnóstico
    for (let i = 1; i < registrosData.length; i++) {
      const fila = registrosData[i];
      const fechaRegistro = fila[0]; // La fecha está en columna 0
      const nombrePersona = fila[1]; // El nombre está en columna 1

      // DEBUG: Mostrar las primeras filas para entender la estructura
      if (i <= 3) {
        Logger.log(`🔍 FILA ${i}: [${fila.join(", ")}]`);
        Logger.log(`   Fecha (col 0): ${fechaRegistro}`);
        Logger.log(`   Nombre (col 1): ${nombrePersona}`);
      }

      if (!nombrePersona || !fechaRegistro) continue;

      // Guardar ejemplos de nombres para comparación
      if (ejemplosNombres.length < 10) {
        ejemplosNombres.push(nombrePersona.toString());
      }

      const nombreRegistroNormalizado = nombrePersona.toString().toUpperCase();

      // Verificar nombre - USAR EXACTAMENTE LA MISMA LÓGICA QUE LA FUNCIÓN ORIGINAL
      if (nombreRegistroNormalizado === nombreNormalizado) {
        registrosConNombreCoincidente++;
        Logger.log(
          `✅ NOMBRE COINCIDENTE EXACTO en fila ${i + 1}: "${nombrePersona}"`
        );

        // Verificar fecha
        const fecha = new Date(fechaRegistro);
        if (
          fecha.getMonth() + 1 === parseInt(mes) &&
          fecha.getFullYear() === parseInt(anio)
        ) {
          registrosConFechaCorrecta++;
          Logger.log(
            `✅ FECHA CORRECTA en fila ${i + 1}: ${fecha.toDateString()}`
          );

          // Verificar quincena
          const dia = fecha.getDate();
          let quincenaCorrecta = false;
          if (quincena === "completo") {
            quincenaCorrecta = true;
          } else if (quincena === "1-15" && dia <= 15) {
            quincenaCorrecta = true;
          } else if (quincena === "16-31" && dia > 15) {
            quincenaCorrecta = true;
          }

          if (quincenaCorrecta) {
            registrosConQuincenaCorrecta++;
            registrosEncontrados++;
            Logger.log(
              `✅ REGISTRO COMPLETO en fila ${
                i + 1
              }: ${nombrePersona}, ${fecha.toDateString()}`
            );
          } else {
            Logger.log(
              `❌ QUINCENA INCORRECTA en fila ${
                i + 1
              }: día ${dia}, quincena buscada: ${quincena}`
            );
          }
        } else {
          Logger.log(
            `❌ FECHA INCORRECTA en fila ${
              i + 1
            }: ${fecha.toDateString()}, buscando: ${mes}/${anio}`
          );
        }
      }
    }

    const resultado = {
      nombreBuscado: nombre,
      nombreNormalizado: nombreNormalizado,
      periodo: `${mes}/${anio}`,
      quincena: quincena,
      totalFilas: registrosData.length - 1, // Excluir header
      registrosConNombreCoincidente: registrosConNombreCoincidente,
      registrosConFechaCorrecta: registrosConFechaCorrecta,
      registrosConQuincenaCorrecta: registrosConQuincenaCorrecta,
      registrosEncontrados: registrosEncontrados,
      ejemplosNombres: ejemplosNombres,
      mensaje:
        registrosEncontrados > 0
          ? `✅ Encontrados ${registrosEncontrados} registros válidos`
          : `❌ No se encontraron registros válidos. Problema: ${
              registrosConNombreCoincidente === 0
                ? "NOMBRE no coincide exactamente"
                : registrosConFechaCorrecta === 0
                ? "FECHA incorrecta"
                : "QUINCENA incorrecta"
            }`,
    };

    Logger.log(
      `📋 RESULTADO DIAGNÓSTICO: ${JSON.stringify(resultado, null, 2)}`
    );
    return resultado;
  } catch (error) {
    Logger.log(`❌ ERROR en diagnóstico: ${error.toString()}`);
    return { error: error.toString() };
  }
}

function obtenerDetalleRegistrosSimplificado(nombre, mes, anio, quincena) {
  Logger.log(
    `🔄 OBTENIENDO detalles SIMPLIFICADO para: ${nombre}, ${mes}/${anio}, ${quincena}`
  );

  try {
    // Usar exactamente la misma lógica que calcularCostos
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetRegistros = ss.getSheetByName("RegistrosProcedimientos");
    const hojaPrecios = ss.getSheetByName("Precios");
    const hojaPersonal = ss.getSheetByName("Personal");

    if (!sheetRegistros || !hojaPrecios || !hojaPersonal) {
      throw new Error(
        "Faltan hojas necesarias: RegistrosProcedimientos, Precios o Personal"
      );
    }

    const registrosData = sheetRegistros.getDataRange().getValues();
    const preciosDatos = hojaPrecios.getDataRange().getValues();

    Logger.log(
      `📊 Total de filas en RegistrosProcedimientos: ${registrosData.length}`
    );

    // Construir diccionario de precios EXACTO como en calcularCostos
    const preciosPorProcedimiento = {};
    Logger.log(`📊 Procesando ${preciosDatos.length} filas de precios...`);

    preciosDatos.slice(1).forEach((fila, index) => {
      const procedimiento = fila[0];
      const precios = {
        doctorLV: fila[1] || 0,
        doctorSab: fila[2] || 0,
        anest: fila[3] || 0,
        tecnico: fila[4] || 0,
      };

      preciosPorProcedimiento[procedimiento] = precios;
      Logger.log(
        `💰 Precios para "${procedimiento}": ${JSON.stringify(precios)}`
      );
    });

    Logger.log(
      `📋 Diccionario de precios creado con ${
        Object.keys(preciosPorProcedimiento).length
      } procedimientos`
    );

    // Obtener el tipo de personal EXACTO como en calcularCostos
    const tipoPersonal = obtenerTipoDePersonal(nombre, hojaPersonal);
    Logger.log(`🏥 Tipo de personal para ${nombre}: ${tipoPersonal}`);

    if (!tipoPersonal) {
      Logger.log(`❌ No se encontró tipo de personal para ${nombre}`);
      return {
        detalleRegistros: [],
        metadatos: {
          nombre: nombre,
          periodo: `${mes}/${anio}`,
          quincena: quincena,
          totalRegistros: 0,
          mensaje: `No se encontró tipo de personal para ${nombre}`,
        },
      };
    }

    // Mapeo de procedimientos EXACTO como en calcularCostos
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

    // Función de normalización EXACTA como en calcularCostos
    function normalizarNombre(str) {
      return String(str || "")
        .replace(/\s+/g, " ")
        .trim()
        .toUpperCase();
    }

    // Función de coincidencia EXACTA como en calcularCostos
    function esLaMismaPersona(nombreSeleccionado, nombreEnRegistro) {
      if (!nombreSeleccionado || !nombreEnRegistro) {
        return false;
      }

      const seleccionadoNorm = normalizarNombre(nombreSeleccionado);
      const registroNorm = normalizarNombre(nombreEnRegistro);

      // Coincidencia exacta
      if (seleccionadoNorm === registroNorm) {
        Logger.log(`✅ COINCIDENCIA EXACTA: "${seleccionadoNorm}"`);
        return true;
      }

      Logger.log(`❌ NO COINCIDE: "${seleccionadoNorm}" ≠ "${registroNorm}"`);
      return false;
    }

    const detalleRegistros = [];
    let registrosEncontrados = 0;

    // Buscar registros con EXACTAMENTE la misma lógica que calcularCostos
    Logger.log(
      `📋 Buscando registros para: "${nombre}" en ${mes}/${anio}, quincena: ${quincena}`
    );

    registrosData.slice(1).forEach((fila, idx) => {
      const numeroFila = idx + 2; // +2 porque idx empieza en 0 y saltamos encabezados

      try {
        // Validar fecha - EXACTO como en calcularCostos
        let fecha = fila[0];
        if (!fecha) {
          return;
        }

        if (!(fecha instanceof Date)) {
          fecha = new Date(fecha);
        }

        if (isNaN(fecha.getTime())) {
          return;
        }

        const filaMes = fecha.getMonth() + 1;
        const filaAnio = fecha.getFullYear();
        const filaDia = fecha.getDate();
        const personaEnRegistro = fila[1];

        // Verificar coincidencias - EXACTO como en calcularCostos
        const esPersonaCoincide = esLaMismaPersona(nombre, personaEnRegistro);
        const esMismoMes = Number(filaMes) === Number(mes);
        const esMismoAnio = Number(filaAnio) === Number(anio);

        if (!esPersonaCoincide || !esMismoMes || !esMismoAnio) {
          return;
        }

        // Verificar quincena - EXACTO como en calcularCostos
        if (quincena === "1-15" && filaDia > 15) {
          return;
        }
        if (quincena === "16-31" && filaDia <= 15) {
          return;
        }

        registrosEncontrados++;
        Logger.log(
          `✅ REGISTRO VÁLIDO #${registrosEncontrados} - Procesando procedimientos...`
        );

        const esSabado = fecha.getDay() === 6;
        const fechaFormateada = fecha.toLocaleDateString("es-ES", {
          weekday: "long",
          year: "numeric",
          month: "long",
          day: "numeric",
        });

        const procedimientos = [];
        let subtotalRegistro = 0;
        let colIndex = 2;

        // Procesar cada procedimiento EXACTO como en calcularCostos
        Object.keys(mapeoProcedimientos).forEach((key) => {
          const cantidad = Number(fila[colIndex]) || 0;
          if (cantidad > 0) {
            const procNombre = mapeoProcedimientos[key];
            const precios = preciosPorProcedimiento[procNombre] || {};
            let costo = 0;

            // Calcular costo EXACTO como en calcularCostos
            if (tipoPersonal === "Doctor") {
              costo = esSabado ? precios.doctorSab || 0 : precios.doctorLV || 0;
            } else if (tipoPersonal === "Anestesiólogo") {
              costo = precios.anest || 0;
            } else if (tipoPersonal === "Técnico") {
              costo = precios.tecnico || 0;
            }

            if (isNaN(costo) || costo === undefined || costo === null) {
              Logger.log(
                `⚠️ ADVERTENCIA: Precio inválido para ${procNombre} (${tipoPersonal}). Estableciendo a 0.`
              );
              costo = 0;
            }

            const subtotal = cantidad * costo;
            const ivaMonto = subtotal * 0.04;
            const total = subtotal + ivaMonto;

            procedimientos.push({
              nombre: procNombre,
              cantidad: cantidad,
              costoUnitario: costo,
              subtotal: subtotal,
              iva: ivaMonto,
              total: total,
            });
            subtotalRegistro += subtotal; // ✅ CORREGIDO: Sumar solo el subtotal (sin IVA)

            Logger.log(
              `💰 ${procNombre}: ${cantidad} x ${costo} = ${subtotal} + IVA ${ivaMonto} = ${total}`
            );
          }
          colIndex++;
        });

        if (procedimientos.length > 0) {
          detalleRegistros.push({
            fila: numeroFila,
            fecha: fechaFormateada,
            fechaOriginal: fecha.toISOString(),
            persona: personaEnRegistro,
            esSabado: esSabado,
            procedimientos: procedimientos,
            subtotalRegistro: subtotalRegistro,
            comentario: fila[18] || "",
          });
        }
      } catch (error) {
        Logger.log(`❌ Error procesando fila ${numeroFila}: ${error.message}`);
      }
    });

    const resultado = {
      detalleRegistros: detalleRegistros,
      metadatos: {
        nombre: nombre,
        periodo: `${mes}/${anio}`,
        quincena: quincena,
        totalRegistros: detalleRegistros.length,
        mensaje: `Encontrados ${detalleRegistros.length} registros con cálculos reales`,
      },
    };

    Logger.log(
      `✅ Detalles SIMPLIFICADOS obtenidos - ${detalleRegistros.length} registros`
    );

    return resultado;
  } catch (error) {
    Logger.log(`❌ Error en detalles simplificados: ${error.message}`);
    throw error;
  }
}

// ===== FUNCIÓN DE COMPARACIÓN: Encontrar diferencias entre calcularCostos y búsqueda directa =====
function compararLogicas(nombre, mes, anio, quincena) {
  Logger.log(
    `🔍 COMPARANDO lógicas para: ${nombre}, ${mes}/${anio}, ${quincena}`
  );

  try {
    // MÉTODO 1: Usar calcularCostos (que SÍ funciona)
    Logger.log("📊 MÉTODO 1: Ejecutando calcularCostos...");
    const resultadoCalcular = calcularCostos(nombre, mes, anio, quincena);
    const registrosEnCalcular = resultadoCalcular.detalleRegistros
      ? resultadoCalcular.detalleRegistros.length
      : 0;
    Logger.log(`✅ calcularCostos encontró: ${registrosEnCalcular} registros`);

    // MÉTODO 2: Buscar directamente (que NO funciona)
    Logger.log("🔍 MÉTODO 2: Búsqueda directa...");
    const sheetRegistros = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      "RegistrosProcedimientos"
    );
    const registrosData = sheetRegistros.getDataRange().getValues();

    const nombreNormalizado = nombre.toUpperCase();
    Logger.log(`🔤 Nombre normalizado: "${nombreNormalizado}"`);

    let registrosDirectos = 0;
    let muestraDatos = [];

    // Mostrar muestra de los primeros 10 registros para comparar
    Logger.log("📋 MUESTRA DE DATOS EN LA HOJA:");
    for (let i = 1; i < Math.min(11, registrosData.length); i++) {
      const fila = registrosData[i];
      const nombrePersona = fila[0];
      const fechaRegistro = fila[1];

      Logger.log(`   Fila ${i + 1}: "${nombrePersona}" - ${fechaRegistro}`);

      if (nombrePersona) {
        const nombreNorm = nombrePersona.toString().toUpperCase();
        Logger.log(`      Normalizado: "${nombreNorm}"`);
        Logger.log(
          `      ¿Coincide con "${nombreNormalizado}"? ${
            nombreNorm === nombreNormalizado
          }`
        );

        if (nombreNorm === nombreNormalizado) {
          Logger.log(`      ⭐ COINCIDENCIA ENCONTRADA!`);
          muestraDatos.push({
            fila: i + 1,
            nombre: nombrePersona,
            fecha: fechaRegistro,
          });
        }
      }
    }

    // Buscar en toda la hoja
    for (let i = 1; i < registrosData.length; i++) {
      const fila = registrosData[i];
      const nombrePersona = fila[0];
      const fechaRegistro = fila[1];

      if (!nombrePersona || !fechaRegistro) continue;

      const nombreRegistroNormalizado = nombrePersona.toString().toUpperCase();
      if (nombreRegistroNormalizado !== nombreNormalizado) continue;

      const fecha = new Date(fechaRegistro);
      if (
        fecha.getMonth() + 1 !== parseInt(mes) ||
        fecha.getFullYear() !== parseInt(anio)
      )
        continue;

      if (quincena !== "completo") {
        const dia = fecha.getDate();
        if (quincena === "1-15" && dia > 15) continue;
        if (quincena === "16-31" && dia <= 15) continue;
      }

      registrosDirectos++;
      Logger.log(
        `✅ REGISTRO DIRECTO #${registrosDirectos} - Fila ${
          i + 1
        }: ${nombrePersona}, ${fecha.toDateString()}`
      );
    }

    Logger.log(`\n🎯 COMPARACIÓN FINAL:`);
    Logger.log(
      `   📊 calcularCostos encontró: ${registrosEnCalcular} registros`
    );
    Logger.log(
      `   🔍 Búsqueda directa encontró: ${registrosDirectos} registros`
    );
    Logger.log(
      `   📋 Coincidencias de nombre en muestra: ${muestraDatos.length}`
    );

    if (muestraDatos.length > 0) {
      Logger.log(`   📋 Datos de coincidencias:`);
      muestraDatos.forEach((item) => {
        Logger.log(`      Fila ${item.fila}: "${item.nombre}" - ${item.fecha}`);
      });
    }

    return {
      calcularCostos: registrosEnCalcular,
      busquedaDirecta: registrosDirectos,
      coincidenciasNombre: muestraDatos.length,
      muestra: muestraDatos,
    };
  } catch (error) {
    Logger.log(`❌ Error en comparación: ${error.message}`);
    Logger.log(`Stack: ${error.stack}`);
    throw error;
  }
}

// ===== FUNCIÓN DE VERIFICACIÓN: Chequear parámetros exactos =====
function verificarParametros(nombre, mes, anio, quincena) {
  Logger.log(`🔍 VERIFICANDO parámetros recibidos:`);
  Logger.log(`   nombre: "${nombre}" (tipo: ${typeof nombre})`);
  Logger.log(`   mes: "${mes}" (tipo: ${typeof mes})`);
  Logger.log(`   anio: "${anio}" (tipo: ${typeof anio})`);
  Logger.log(`   quincena: "${quincena}" (tipo: ${typeof quincena})`);

  // Conversiones
  const mesInt = parseInt(mes);
  const anioInt = parseInt(anio);
  Logger.log(`   mes convertido: ${mesInt} (tipo: ${typeof mesInt})`);
  Logger.log(`   anio convertido: ${anioInt} (tipo: ${typeof anioInt})`);

  // Probar con estos parámetros exactos en calcularCostos
  Logger.log(`\n🧪 Probando calcularCostos con parámetros exactos...`);
  try {
    const resultado = calcularCostos(nombre, mesInt, anioInt, quincena);
    Logger.log(
      `✅ calcularCostos exitoso: ${
        resultado.detalleRegistros ? resultado.detalleRegistros.length : 0
      } registros`
    );

    // Mostrar algunos detalles
    if (resultado.detalleRegistros && resultado.detalleRegistros.length > 0) {
      Logger.log(`📋 Muestra de registros encontrados por calcularCostos:`);
      resultado.detalleRegistros.slice(0, 3).forEach((reg, index) => {
        Logger.log(
          `   ${index + 1}. Fila ${reg.fila}: ${reg.persona} - ${reg.fecha}`
        );
      });
    }
  } catch (error) {
    Logger.log(`❌ Error en calcularCostos: ${error.message}`);
  }

  return {
    parametros: {
      nombre: nombre,
      mes: mes,
      anio: anio,
      quincena: quincena,
      mesInt: mesInt,
      anioInt: anioInt,
    },
    tipos: {
      nombre: typeof nombre,
      mes: typeof mes,
      anio: typeof anio,
      quincena: typeof quincena,
    },
  };
}

// ===== FUNCIÓN PARA PROBAR EL DIAGNÓSTICO =====
function probarDiagnosticoFiltrado() {
  Logger.log("🧪 === PRUEBA DE DIAGNÓSTICO DE FILTRADO ===");

  // Probar con diferentes nombres
  const nombresPrueba = [
    "Anest Manuel",
    "Anest Nicole",
    "ANEST MANUEL",
    "ANEST NICOLE",
  ];

  nombresPrueba.forEach((nombre) => {
    diagnosticarFiltradoPersonal(nombre);
  });
}

/**
 * Función de prueba para diagnosticar problemas de comunicación frontend-backend
 */
function testBusquedaBiopsias(tipo, valor, valor2) {
  Logger.log(`🧪 TEST: Recibidos ${arguments.length} argumentos`);
  Logger.log(`🧪 TEST: tipo='${tipo}', valor='${valor}', valor2='${valor2}'`);
  Logger.log(
    `🧪 TEST: Tipos - tipo:${typeof tipo}, valor:${typeof valor}, valor2:${typeof valor2}`
  );

  return {
    argumentos: arguments.length,
    tipo: tipo,
    valor: valor,
    valor2: valor2,
    mensaje: "Función de prueba ejecutada correctamente",
  };
}

/**
 * Función para debugging manual - ejecutar desde el editor de Apps Script
 */
function debugBusquedaCedula() {
  Logger.log("🧪 INICIO DEBUG - Búsqueda por cédula");

  try {
    const resultado = buscarBiopsiasServidorMejorado(
      "cedula",
      "801180052",
      null
    );
    Logger.log(
      `🧪 RESULTADO DEBUG: ${resultado.length} resultados encontrados`
    );

    if (resultado.length > 0) {
      Logger.log("🧪 PRIMER RESULTADO:");
      Logger.log(JSON.stringify(resultado[0]));
    }

    return resultado;
  } catch (error) {
    Logger.log(`🧪 ERROR DEBUG: ${error.message}`);
    Logger.log(`🧪 STACK: ${error.stack}`);
    return [];
  }
}

// Nueva función idéntica para evitar problemas de caché
function buscarBiopsiasServidor_v2(parametros) {
  try {
    Logger.log(
      "🚀 buscarBiopsiasServidor_v2 - Iniciando con parámetros:",
      parametros
    );

    // Extraer parámetros del objeto
    const {
      type: searchType,
      value: searchValue,
      value2: searchValue2,
    } = parametros || {};

    Logger.log(
      `🔍 Tipo de búsqueda: ${searchType}, Valor: ${searchValue}, Valor2: ${searchValue2}`
    );

    // Funciones auxiliares (igual que en la función original que funciona)
    const normalizarTexto = (str) =>
      String(str || "")
        .toLowerCase()
        .trim();
    const normalizarIdentificacion = (str) =>
      String(str || "")
        .replace(/[-\s]/g, "")
        .toUpperCase();
    const formatearFecha = (fecha) => {
      if (!fecha) return "";
      const fechaObj = new Date(fecha);
      if (isNaN(fechaObj.getTime())) return "";
      return fechaObj.toISOString().split("T")[0];
    };

    // Usar openById() para acceso desde el frontend web
    const ss = SpreadsheetApp.openById(
      "1NjqsT9ApcCb9bpkNY2K01Z9_YRkxoSlPYD52Ku0dGS8"
    );
    const hoja = ss.getSheetByName("RegistroBiopsias");

    if (!hoja) {
      Logger.log("❌ Error: Hoja 'RegistroBiopsias' no encontrada");
      return [];
    }

    // Usar getDataRange() como en la función original
    const datos = hoja.getDataRange().getValues();
    Logger.log(`📊 Total de filas: ${datos.length}`);

    const resultados = [];

    // Procesar cada fila (saltando el encabezado) - como en la función original
    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      let coincide = false;

      Logger.log(`🔍 Procesando fila ${i} con tipo: ${searchType}`);

      switch (searchType) {
        case "cedula":
          const identificacionHoja = normalizarIdentificacion(fila[3]);
          const identificacionBusqueda = normalizarIdentificacion(searchValue);
          Logger.log(
            `🆔 CEDULA DETALLE: Fila ${i}, Hoja="${fila[3]}" -> "${identificacionHoja}", Búsqueda="${searchValue}" -> "${identificacionBusqueda}"`
          );
          coincide = identificacionHoja === identificacionBusqueda;
          if (coincide) {
            Logger.log(`🎯 CEDULA MATCH encontrado en fila ${i + 1}`);
          }
          break;

        case "nombre":
          const nombreHoja = normalizarTexto(fila[5]);
          const nombreBusqueda = normalizarTexto(searchValue);
          Logger.log(
            `👤 NOMBRE: Comparando "${nombreHoja}" incluye "${nombreBusqueda}"`
          );
          coincide = nombreHoja.includes(nombreBusqueda);
          break;

        case "fecha":
          const fechaHoja = formatearFecha(fila[0]);
          const fechaBusqueda = searchValue;
          Logger.log(
            `📅 FECHA: Comparando "${fechaHoja}" vs "${fechaBusqueda}"`
          );
          coincide = fechaHoja === fechaBusqueda;
          break;

        case "mes":
          if (fila[0] instanceof Date || typeof fila[0] === "string") {
            const fecha = new Date(fila[0]);
            const mesHoja = fecha.getMonth() + 1;
            const anioHoja = fecha.getFullYear();
            const mesBusqueda = parseInt(searchValue);
            const anioBusqueda = parseInt(searchValue2);
            Logger.log(
              `📅 MES: Comparando ${mesHoja}/${anioHoja} vs ${mesBusqueda}/${anioBusqueda}`
            );
            coincide = mesHoja === mesBusqueda && anioHoja === anioBusqueda;
          }
          break;

        case "estado":
          const estadoHoja = String(fila[9] || "")
            .toLowerCase()
            .trim();
          const estadoBusqueda = searchValue.toLowerCase().trim();
          Logger.log(
            `📋 ESTADO: Comparando "${estadoHoja}" vs "${estadoBusqueda}"`
          );
          coincide = estadoHoja === estadoBusqueda;
          break;
      }

      if (coincide) {
        Logger.log(`✅ Match encontrado en fila ${i + 1}`);
        // Usar el mismo formato que la función original
        resultados.push({
          fila: i + 1, // Número de fila real
          fecha: fila[0],
          numeroRequisicion: fila[1],
          codigoProcedimiento: fila[2],
          identificacion: fila[3],
          tipoIdentificacion: fila[4],
          nombreCliente: fila[5],
          precioProcedimiento: fila[6],
          tipoProcedimiento: fila[7],
          sede: fila[8],
          estado: fila[9],
        });
      }
    }

    Logger.log(`🎯 Total de resultados encontrados: ${resultados.length}`);
    return resultados;
  } catch (error) {
    Logger.log("❌ Error en buscarBiopsiasServidor_v2:", error.toString());
    return [];
  }
}

// Función de prueba para verificar comunicación
function testBusquedaCedula() {
  const parametros = {
    type: "cedula",
    value: "801180052",
    value2: null,
  };

  console.log("🧪 TEST: Probando buscarBiopsiasServidor_v2");
  console.log("🧪 TEST: Parámetros:", parametros);

  const resultado = buscarBiopsiasServidor_v2(parametros);

  console.log("🧪 TEST: Resultado:", resultado);
  console.log("🧪 TEST: Cantidad de resultados:", resultado.length);

  return resultado;
}

// Función súper simple para probar comunicación
function testComunicacion(parametros) {
  Logger.log("🧪 testComunicacion recibió:", parametros);
  return {
    mensaje: "Comunicación exitosa",
    parametrosRecibidos: parametros,
    timestamp: new Date().toISOString(),
  };
}

// Función de prueba directa con los mismos parámetros del frontend
function buscarBiopsiasServidor_v3(parametros) {
  Logger.log("🚀 V3 - Parámetros recibidos:", JSON.stringify(parametros));

  try {
    // Verificar que recibimos un objeto
    if (typeof parametros !== "object" || parametros === null) {
      Logger.log("❌ V3 - Parámetros no es un objeto válido");
      return [];
    }

    const { type, value, value2 } = parametros;
    Logger.log(
      `🔍 V3 - Extracted: type=${type}, value=${value}, value2=${value2}`
    );

    // Probar acceso a spreadsheet
    Logger.log("🔗 V3 - Intentando acceder al spreadsheet...");
    const ss = SpreadsheetApp.openById(
      "1NjqsT9ApcCb9bpkNY2K01Z9_YRkxoSlPYD52Ku0dGS8"
    );
    Logger.log("✅ V3 - Spreadsheet obtenido correctamente");

    const hoja = ss.getSheetByName("RegistroBiopsias");
    if (!hoja) {
      Logger.log("❌ V3 - Hoja 'RegistroBiopsias' no encontrada");
      return [];
    }
    Logger.log("✅ V3 - Hoja encontrada");

    const datos = hoja.getDataRange().getValues();
    Logger.log(`📊 V3 - Total de filas obtenidas: ${datos.length}`);

    // Devolver info básica para verificar
    return [
      {
        debug: true,
        totalFilas: datos.length,
        parametros: parametros,
        primeraFila: datos.length > 1 ? datos[1] : null,
      },
    ];
  } catch (error) {
    Logger.log("❌ V3 - Error:", error.toString());
    return [
      {
        error: true,
        mensaje: error.toString(),
      },
    ];
  }
}

// Función para revisar datos en la hoja
function revisarDatosHoja() {
  try {
    const ss = SpreadsheetApp.openById(
      "1NjqsT9ApcCb9bpkNY2K01Z9_YRkxoSlPYD52Ku0dGS8"
    );
    const hoja = ss.getSheetByName("RegistroBiopsias");

    if (!hoja) {
      console.log("❌ Hoja no encontrada");
      return;
    }

    const ultimaFila = hoja.getLastRow();
    console.log(`📊 Última fila: ${ultimaFila}`);

    if (ultimaFila <= 1) {
      console.log("📋 No hay datos");
      return;
    }

    // Leer las primeras 5 filas para ver la estructura
    const rango = hoja.getRange(1, 1, Math.min(ultimaFila, 6), 10);
    const datos = rango.getValues();

    console.log("📋 Encabezados:", datos[0]);

    for (let i = 1; i < datos.length; i++) {
      console.log(`📋 Fila ${i + 1}:`, datos[i]);
      if (datos[i][3]) {
        // Si hay cédula
        console.log(`🆔 Cédula en fila ${i + 1}: "${datos[i][3]}"`);
      }
    }

    return datos;
  } catch (error) {
    console.log("❌ Error:", error.toString());
  }
}
