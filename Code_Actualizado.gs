// ================= CÓDIGO ACTUALIZADO Y MEJORADO =================
// Versión consolidada que reemplaza el Code.gs original
// Elimina funciones duplicadas y soluciona problemas con búsqueda de personal

// --- CONFIGURACIÓN DE COLUMNAS CONFIRMADA ---
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

// --- FUNCIONES PRINCIPALES DEL MENÚ (CONSERVADAS) ---
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

function getPageHtml(pageName) {
  Logger.log(
    "DEBUG (Server): getPageHtml() ejecutada para la página: " + pageName
  );

  if (!pageName) {
    Logger.log(
      "ADVERTENCIA: pageName es undefined o null, usando 'mainMenu' como valor predeterminado"
    );
    pageName = "mainMenu";
  }

  let template;
  let pageTitle = "Registro Médico";

  try {
    switch (pageName) {
      case "mainMenu":
      case "mainMenuContent":
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
      case "registroBiopsias":
        template = HtmlService.createTemplateFromFile("registroBiopsias");
        pageTitle = "Registro de Biopsias";
        break;
      case "buscarBiopsias":
        template = HtmlService.createTemplateFromFile("buscarBiopsiasReg");
        pageTitle = "Búsqueda y Reporte de Biopsias";
        break;
      default:
        Logger.log(
          "ERROR (Server): Página no reconocida en getPageHtml(): " + pageName
        );
        template = HtmlService.createTemplateFromFile("mainMenu");
        pageTitle = "Menú Principal";
        break;
    }

    const htmlOutput = template.evaluate();
    htmlOutput.setTitle(pageTitle);
    htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    Logger.log(
      "DEBUG (Server): HTML generado para " +
        pageName +
        " con título: " +
        pageTitle
    );
    return htmlOutput;
  } catch (error) {
    Logger.log("ERROR (Server): " + error.toString());
    const errorTemplate = HtmlService.createTemplate(
      "<p>Error: No se pudo cargar la página solicitada.</p>"
    );
    return errorTemplate.evaluate().setTitle("Error");
  }
}

function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (error) {
    Logger.log("ERROR al incluir " + filename + ": " + error.toString());
    return "<!-- Error al cargar " + filename + " -->";
  }
}

// ================= FUNCIONES DE PERSONAL MEJORADAS =================

/**
 * Función unificada para obtener personal - REEMPLAZA TODAS LAS FUNCIONES DUPLICADAS
 * @param {number[]} columnIndexes - Array de índices de columnas a incluir
 * @returns {string[]} Array de nombres de personal únicos y ordenados
 */
function obtenerPersonalFiltrado(columnIndexes) {
  try {
    // Validar y establecer valores por defecto
    if (!Array.isArray(columnIndexes)) {
      columnIndexes = [0, 1, 2]; // Doctor, Anestesiólogo, Técnico por defecto
      Logger.log(
        "ADVERTENCIA: columnIndexes no válido, usando [0,1,2] por defecto"
      );
    }

    Logger.log(
      `🔍 obtenerPersonalFiltrado - Columnas solicitadas: ${columnIndexes}`
    );

    const hoja =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Personal");
    if (!hoja) {
      Logger.log("❌ La hoja 'Personal' no existe.");
      throw new Error("❌ La hoja 'Personal' no existe.");
    }

    const datos = hoja.getDataRange().getValues();
    const personal = new Set();

    Logger.log(`📊 Procesando ${datos.length} filas de la hoja Personal`);

    // Procesar desde la fila 2 (saltar encabezados)
    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];

      columnIndexes.forEach((colIndex) => {
        if (colIndex >= 0 && colIndex < fila.length) {
          const nombre = fila[colIndex];
          if (nombre && typeof nombre === "string" && nombre.trim()) {
            const nombreLimpio = nombre.trim();
            personal.add(nombreLimpio);

            Logger.log(
              `   ✅ Agregado: "${nombreLimpio}" (Columna ${colIndex} - ${
                TIPOS_PERSONAL[colIndex] || "Desconocido"
              })`
            );
          }
        } else {
          Logger.log(`   ⚠️ Índice de columna fuera de rango: ${colIndex}`);
        }
      });
    }

    const resultado = [...personal].sort();
    Logger.log(`✅ Total de personal encontrado: ${resultado.length}`);
    Logger.log(`📋 Lista: ${resultado.join(", ")}`);

    return resultado;
  } catch (error) {
    Logger.log(`❌ Error en obtenerPersonalFiltrado: ${error.message}`);
    return [];
  }
}

/**
 * Función mejorada para normalizar nombres - elimina caracteres especiales y unifica formato
 */
function normalizarNombre(nombre) {
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
 * Función mejorada para determinar el tipo de personal
 * @param {string} nombre - Nombre del personal a buscar
 * @param {Object} hojaPersonal - Hoja de Personal (opcional)
 * @returns {string|null} Tipo de personal o null si no se encuentra
 */
function obtenerTipoDePersonal(nombre, hojaPersonal = null) {
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
    const nombreNormalizado = normalizarNombre(nombre);

    Logger.log(`🔍 obtenerTipoDePersonal: Buscando "${nombreNormalizado}"`);
    Logger.log(`📊 Revisando ${datos.length} filas en Personal`);

    // Buscar en todas las columnas de personal (A-F)
    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];

      for (let col = 0; col <= 5; col++) {
        const valorCelda = String(fila[col] || "").trim();

        if (valorCelda && sonLaMismaPersona(nombreNormalizado, valorCelda)) {
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

/**
 * Función mejorada para comparar nombres de personal
 */
function sonLaMismaPersona(nombre1, nombre2) {
  if (!nombre1 || !nombre2) return false;

  const norm1 = normalizarNombre(nombre1);
  const norm2 = normalizarNombre(nombre2);

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
 * Función para obtener el email de un personal
 */
function obtenerEmailParaPersonal(nombrePersonal) {
  try {
    const hojaPersonal =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Personal");
    if (!hojaPersonal) {
      Logger.log(
        "❌ La hoja 'Personal' no existe al intentar obtener el email."
      );
      return "";
    }

    const datos = hojaPersonal.getDataRange().getValues();
    const emailColIndex = COLUMNAS_PERSONAL.EMAIL; // Columna G (índice 6)
    const nombreNormalizado = normalizarNombre(nombrePersonal);

    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];

      // Buscar en todas las columnas de nombres (A-F)
      for (let col = 0; col <= 5; col++) {
        const valorCelda = String(fila[col] || "").trim();
        if (valorCelda && sonLaMismaPersona(nombreNormalizado, valorCelda)) {
          const email = fila[emailColIndex];
          Logger.log(`📧 Email encontrado para ${nombrePersonal}: ${email}`);
          return email ? String(email).trim() : "";
        }
      }
    }

    Logger.log(`❌ No se encontró email para: ${nombrePersonal}`);
    return "";
  } catch (error) {
    Logger.log(`❌ Error al obtener email: ${error.message}`);
    return "";
  }
}

// ================= FUNCIÓN DE CÁLCULO DE COSTOS MEJORADA =================

function calcularCostos(nombre, mes, anio, quincena) {
  try {
    // Validar parámetros
    if (!nombre || !mes || !anio || !quincena) {
      Logger.log("❌ Parámetros incompletos en calcularCostos");
      Logger.log(
        `Recibido: nombre="${nombre}", mes="${mes}", anio="${anio}", quincena="${quincena}"`
      );
      return {
        lv: {},
        sab: {},
        totales: { subtotal: 0, iva: 0, total_con_iva: 0 },
      };
    }

    Logger.log(
      `🚀 INICIANDO calcularCostos: nombre="${nombre}", mes=${mes}, anio=${anio}, quincena="${quincena}"`
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

    // Obtener el tipo de personal
    const tipoPersonal = obtenerTipoDePersonal(nombre, hojaPersonal);
    Logger.log(`🏥 Tipo de personal para "${nombre}": "${tipoPersonal}"`);

    if (!tipoPersonal) {
      Logger.log(
        `❌ El nombre '${nombre}' no está registrado en la hoja Personal.`
      );
      return {
        lv: {},
        sab: {},
        totales: { subtotal: 0, iva: 0, total_con_iva: 0 },
      };
    }

    // Crear mapa de precios
    const preciosPorProcedimiento = {};
    preciosDatos.slice(1).forEach((fila) => {
      if (fila[0]) {
        // Si hay nombre de procedimiento
        preciosPorProcedimiento[fila[0]] = {
          doctorLV: fila[1] || 0,
          doctorSab: fila[2] || 0,
          anest: fila[3] || 0,
          tecnico: fila[4] || 0,
        };
      }
    });

    // Configurar mapeo de procedimientos
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

    // Inicializar resumen
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
    const nombreNormalizado = normalizarNombre(nombre);

    // Procesar registros
    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      const fechaRegistro = fila[0];
      const personalRegistro = String(fila[1] || "").trim();

      // Verificar si es la persona correcta
      if (
        !personalRegistro ||
        !sonLaMismaPersona(nombreNormalizado, personalRegistro)
      ) {
        continue;
      }

      // Validar fecha
      if (!fechaRegistro || !(fechaRegistro instanceof Date)) {
        continue;
      }

      // Verificar si la fecha está en el rango solicitado
      if (!estaEnRangoFecha(fechaRegistro, mes, anio, quincena)) {
        continue;
      }

      registrosEncontrados++;
      Logger.log(
        `✅ Registro ${registrosEncontrados} encontrado en fila ${
          i + 1
        }: ${personalRegistro} - ${fechaRegistro.toDateString()}`
      );

      // Determinar si es sábado
      const esSabado = fechaRegistro.getDay() === 6;
      const tipoReporte = esSabado ? "sab" : "lv";

      // Procesar cada tipo de procedimiento
      Object.keys(mapeoProcedimientos).forEach((key, index) => {
        const columnaIndex = index + 2; // Las columnas de procedimientos empiezan en la C (índice 2)
        const cantidad = parseInt(fila[columnaIndex]) || 0;

        if (cantidad > 0) {
          const nombreProcedimiento = mapeoProcedimientos[key];
          const precios = preciosPorProcedimiento[nombreProcedimiento];

          if (precios) {
            let costoUnitario = 0;

            // Determinar el costo según el tipo de personal y día
            switch (tipoPersonal) {
              case "Doctor":
                costoUnitario = esSabado ? precios.doctorSab : precios.doctorLV;
                break;
              case "Anestesiólogo":
                costoUnitario = precios.anest;
                break;
              case "Técnico":
                costoUnitario = precios.tecnico;
                break;
              default:
                costoUnitario = 0;
                Logger.log(
                  `⚠️ Tipo de personal no reconocido para cálculo de costos: ${tipoPersonal}`
                );
            }

            if (costoUnitario > 0) {
              const subtotal = cantidad * costoUnitario;
              const ivaCalculado = subtotal * iva;
              const totalConIva = subtotal + ivaCalculado;

              // Actualizar resumen
              resumen[tipoReporte][key].cantidad += cantidad;
              resumen[tipoReporte][key].costo_unitario = costoUnitario;
              resumen[tipoReporte][key].subtotal += subtotal;
              resumen[tipoReporte][key].iva += ivaCalculado;
              resumen[tipoReporte][key].total_con_iva += totalConIva;

              // Actualizar totales generales
              resumen.totales.subtotal += subtotal;
              resumen.totales.iva += ivaCalculado;
              resumen.totales.total_con_iva += totalConIva;

              Logger.log(
                `💰 ${nombreProcedimiento} (${tipoReporte.toUpperCase()}): ${cantidad} x ₡${costoUnitario} = ₡${subtotal}`
              );
            }
          } else {
            Logger.log(
              `⚠️ No se encontraron precios para: ${nombreProcedimiento}`
            );
          }
        }
      });
    }

    Logger.log(
      `📊 RESUMEN FINAL: ${registrosEncontrados} registros procesados`
    );
    Logger.log(`💰 Total: ₡${resumen.totales.total_con_iva.toFixed(2)}`);

    return resumen;
  } catch (error) {
    Logger.log(`❌ Error en calcularCostos: ${error.message}`);
    Logger.log(`Stack trace: ${error.stack}`);
    throw new Error(`Error al calcular costos: ${error.message}`);
  }
}

/**
 * Función auxiliar para verificar si una fecha está en el rango solicitado
 */
function estaEnRangoFecha(fecha, mes, anio, quincena) {
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
    Logger.log(`❌ Error en estaEnRangoFecha: ${error.message}`);
    return false;
  }
}

// ================= FUNCIONES ESPECÍFICAS PARA COMPATIBILIDAD =================

/**
 * Estas funciones mantienen compatibilidad con el código existente
 * pero ahora usan las funciones mejoradas internamente
 */

function obtenerDoctores() {
  return obtenerPersonalFiltrado([COLUMNAS_PERSONAL.DOCTOR]);
}

function obtenerAnestesiologos() {
  return obtenerPersonalFiltrado([COLUMNAS_PERSONAL.ANESTESIOLOGO]);
}

function obtenerTecnicos() {
  return obtenerPersonalFiltrado([COLUMNAS_PERSONAL.TECNICO]);
}

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
    return resultado;
  } catch (error) {
    Logger.log("❌ Error en obtenerPersonalCompleto: " + error.message);
    return [];
  }
}

// ================= FUNCIONES DE DIAGNÓSTICO MEJORADAS =================

function diagnosticarFechasPersonal(nombrePersonal) {
  try {
    Logger.log(`🔍 DIAGNÓSTICO DE FECHAS PARA: "${nombrePersonal}"`);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaRegistros = ss.getSheetByName("RegistrosProcedimientos");

    if (!hojaRegistros) {
      throw new Error("La hoja 'RegistrosProcedimientos' no existe");
    }

    const datos = hojaRegistros.getDataRange().getValues();
    const nombreNormalizado = normalizarNombre(nombrePersonal);

    let totalRegistros = 0;
    let registrosValidosFecha = 0;
    let registrosInvalidosFecha = 0;
    let primerRegistro = null;
    let ultimoRegistro = null;

    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      const personalFila = String(fila[1] || "").trim();
      const fechaFila = fila[0];

      if (personalFila && sonLaMismaPersona(nombreNormalizado, personalFila)) {
        totalRegistros++;

        if (fechaFila instanceof Date && !isNaN(fechaFila.getTime())) {
          registrosValidosFecha++;

          if (!primerRegistro || fechaFila < primerRegistro) {
            primerRegistro = fechaFila;
          }
          if (!ultimoRegistro || fechaFila > ultimoRegistro) {
            ultimoRegistro = fechaFila;
          }
        } else {
          registrosInvalidosFecha++;
        }
      }
    }

    const resultado = {
      nombreBuscado: nombrePersonal,
      totalRegistros: totalRegistros,
      registrosValidosFecha: registrosValidosFecha,
      registrosInvalidosFecha: registrosInvalidosFecha,
      primerRegistro: primerRegistro ? primerRegistro.toDateString() : null,
      ultimoRegistro: ultimoRegistro ? ultimoRegistro.toDateString() : null,
    };

    Logger.log("✅ Diagnóstico completado: " + JSON.stringify(resultado));
    return resultado;
  } catch (error) {
    Logger.log("❌ Error en diagnóstico: " + error.message);
    throw error;
  }
}

function buscarPersonalPorNombre(nombreBuscar) {
  try {
    Logger.log("🔍 Iniciando búsqueda de personal: " + nombreBuscar);

    if (!nombreBuscar || nombreBuscar.trim() === "") {
      return {
        coincidenciasExactas: [],
        coincidenciasParciales: [],
      };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaPersonal = ss.getSheetByName("Personal");

    if (!hojaPersonal) {
      throw new Error("No se encontró la hoja Personal");
    }

    const nombreNormalizado = normalizarNombre(nombreBuscar);
    const datosPersonal = hojaPersonal.getDataRange().getValues();

    const coincidenciasExactas = [];
    const coincidenciasParciales = [];

    for (let i = 1; i < datosPersonal.length; i++) {
      const fila = datosPersonal[i];

      for (let j = 0; j <= 5; j++) {
        // Columnas A-F
        const nombre = String(fila[j] || "").trim();
        if (!nombre) continue;

        const nombreCeldaNorm = normalizarNombre(nombre);

        if (nombreCeldaNorm === nombreNormalizado) {
          coincidenciasExactas.push(nombre);
        } else if (sonLaMismaPersona(nombreNormalizado, nombre)) {
          coincidenciasParciales.push(nombre);
        }
      }
    }

    const resultado = {
      coincidenciasExactas: [...new Set(coincidenciasExactas)],
      coincidenciasParciales: [...new Set(coincidenciasParciales)],
    };

    Logger.log("Búsqueda completada: " + JSON.stringify(resultado));
    return resultado;
  } catch (error) {
    Logger.log("Error en buscarPersonalPorNombre: " + error.message);
    throw error;
  }
}

function verificarDatosPersonal(nombrePersonal) {
  try {
    Logger.log(`🔍 Verificando datos para: ${nombrePersonal}`);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaRegistros = ss.getSheetByName("RegistrosProcedimientos");

    if (!hojaRegistros) {
      throw new Error("La hoja 'RegistrosProcedimientos' no existe");
    }

    const datos = hojaRegistros.getDataRange().getValues();
    const nombreNormalizado = normalizarNombre(nombrePersonal);

    let totalRegistros = 0;
    let fechaInicio = null;
    let fechaFin = null;
    const tiposProcedimientos = new Set();

    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      const personalFila = String(fila[1] || "").trim();
      const fechaFila = fila[0];

      if (personalFila && sonLaMismaPersona(nombreNormalizado, personalFila)) {
        totalRegistros++;

        if (fechaFila instanceof Date) {
          if (!fechaInicio || fechaFila < fechaInicio) {
            fechaInicio = fechaFila;
          }
          if (!fechaFin || fechaFila > fechaFin) {
            fechaFin = fechaFila;
          }
        }

        // Contar tipos de procedimientos (columnas 2 en adelante)
        for (let j = 2; j < fila.length; j++) {
          if (fila[j] && parseInt(fila[j]) > 0) {
            tiposProcedimientos.add(j);
          }
        }
      }
    }

    const resultado = {
      totalRegistros: totalRegistros,
      fechaInicio: fechaInicio ? fechaInicio.toDateString() : null,
      fechaFin: fechaFin ? fechaFin.toDateString() : null,
      tiposProcedimientos: tiposProcedimientos.size,
    };

    Logger.log("Verificación completada: " + JSON.stringify(resultado));
    return resultado;
  } catch (error) {
    Logger.log("Error en verificarDatosPersonal: " + error.message);
    throw error;
  }
}

// ================= FUNCIONES PARA PDF Y EMAIL (CONSERVADAS) =================

function enviarReporteCostos(htmlContent, emailRecipient, subject, fileName) {
  try {
    Logger.log(`📧 Enviando reporte a: ${emailRecipient}`);

    // Crear PDF
    const blob = Utilities.newBlob(htmlContent, "text/html", "temp.html");
    const pdfBlob = blob.getAs("application/pdf");
    pdfBlob.setName(fileName);

    // Enviar email
    GmailApp.sendEmail(
      emailRecipient,
      subject,
      "Adjunto encontrará el reporte solicitado.",
      {
        attachments: [pdfBlob],
      }
    );

    Logger.log("✅ Reporte enviado exitosamente");
    return { success: true, message: "✅ Reporte enviado exitosamente" };
  } catch (error) {
    Logger.log(`❌ Error al enviar reporte: ${error.message}`);
    return { success: false, message: `❌ Error al enviar: ${error.message}` };
  }
}

function generarPdfParaDescarga(htmlContent, fileName, styleSheet, pageTitle) {
  try {
    Logger.log(`📄 Generando PDF para descarga: ${fileName}`);

    const blob = Utilities.newBlob(htmlContent, "text/html", "temp.html");
    const pdfBlob = blob.getAs("application/pdf");

    // Convertir a base64
    const base64 = Utilities.base64Encode(pdfBlob.getBytes());

    Logger.log("✅ PDF generado exitosamente");
    return base64;
  } catch (error) {
    Logger.log(`❌ Error al generar PDF: ${error.message}`);
    throw error;
  }
}

// ================= FUNCIONES DE MOSTRAR PÁGINAS (CONSERVADAS) =================

function mostrarFormularioProcedimientos() {
  const htmlOutput = getPageHtml("formularioProcedimientos");
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Registrar Procedimiento");
}

function mostrarReportePagoProcedimientos() {
  const htmlOutput = getPageHtml("reportePagoProcedimientos");
  SpreadsheetApp.getUi().showModalDialog(
    htmlOutput,
    "Reporte de Procedimientos"
  );
}

function mostrarFormularioUltrasonido() {
  const htmlOutput = getPageHtml("formularioUltrasonido");
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Registrar Ultrasonido");
}

function mostrarReportePagoUltrasonido() {
  const htmlOutput = getPageHtml("reportePagoUltrasonido");
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Reporte de Radiología");
}

function mostrarFormularioHorasExtras() {
  const htmlOutput = getPageHtml("formularioHorasExtras");
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Registrar Horas Extras");
}

function mostrarReportePagoHorasExtras() {
  const htmlOutput = getPageHtml("reportePagoHorasExtras");
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Reporte de Horas Extras");
}

function mostrarRegistroBiopsias() {
  const htmlOutput = getPageHtml("registroBiopsias");
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Registro de Biopsias");
}

function mostrarBuscarBiopsias() {
  const htmlOutput = getPageHtml("buscarBiopsias");
  SpreadsheetApp.getUi().showModalDialog(
    htmlOutput,
    "Búsqueda y Reporte de Biopsias"
  );
}
