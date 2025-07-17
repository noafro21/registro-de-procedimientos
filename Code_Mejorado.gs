// ================= VERSIÓN MEJORADA DEL CÓDIGO =================
// Este archivo contiene las funciones consolidadas y mejoradas

// --- CONFIGURACIÓN DE COLUMNAS ---
const COLUMNAS_PERSONAL = {
  DOCTOR: 0, // Columna A
  ANESTESIOLOGO: 1, // Columna B
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

// ================= FUNCIONES PRINCIPALES MEJORADAS =================

/**
 * Función unificada para obtener personal por tipo o tipos específicos
 * @param {number[]} columnasIndices - Array de índices de columnas a incluir
 * @param {string} tipoPersonal - Tipo específico de personal (opcional)
 * @returns {string[]} Array de nombres de personal únicos y ordenados
 */
function obtenerPersonalPorTipo(
  columnasIndices = [0, 1, 2],
  tipoPersonal = null
) {
  try {
    Logger.log(
      `🔍 obtenerPersonalPorTipo - Columnas: ${columnasIndices}, Tipo: ${tipoPersonal}`
    );

    // Validar parámetros
    if (!Array.isArray(columnasIndices)) {
      columnasIndices = [0, 1, 2];
    }

    const hoja =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Personal");
    if (!hoja) {
      throw new Error("❌ La hoja 'Personal' no existe.");
    }

    const datos = hoja.getDataRange().getValues();
    const personal = new Set();

    Logger.log(`📊 Procesando ${datos.length} filas de la hoja Personal`);

    // Procesar desde la fila 2 (saltar encabezados)
    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];

      columnasIndices.forEach((colIndex) => {
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

    return resultado;
  } catch (error) {
    Logger.log(`❌ Error en obtenerPersonalPorTipo: ${error.message}`);
    return [];
  }
}

/**
 * Función consolidada para obtener personal específico por tipo
 * @param {string} tipo - Tipo de personal ("doctor", "anestesiologo", "tecnico", etc.)
 * @returns {string[]} Array de nombres del tipo especificado
 */
function obtenerPersonalEspecifico(tipo) {
  const tipoLower = tipo.toLowerCase();
  let columnaIndex;

  switch (tipoLower) {
    case "doctor":
    case "doctores":
      columnaIndex = COLUMNAS_PERSONAL.DOCTOR;
      break;
    case "anestesiologo":
    case "anestesiologos":
      columnaIndex = COLUMNAS_PERSONAL.ANESTESIOLOGO;
      break;
    case "tecnico":
    case "tecnicos":
      columnaIndex = COLUMNAS_PERSONAL.TECNICO;
      break;
    case "radiologo":
    case "radiologos":
      columnaIndex = COLUMNAS_PERSONAL.RADIOLOGO;
      break;
    case "enfermero":
    case "enfermeros":
      columnaIndex = COLUMNAS_PERSONAL.ENFERMERO;
      break;
    case "secretaria":
    case "secretarias":
      columnaIndex = COLUMNAS_PERSONAL.SECRETARIA;
      break;
    default:
      Logger.log(`❌ Tipo de personal no reconocido: ${tipo}`);
      return [];
  }

  return obtenerPersonalPorTipo([columnaIndex], tipo);
}

/**
 * Función mejorada para determinar el tipo de personal
 * @param {string} nombre - Nombre del personal a buscar
 * @param {Object} hojaPersonal - Hoja de Personal (opcional)
 * @returns {string|null} Tipo de personal o null si no se encuentra
 */
function obtenerTipoDePersonalMejorado(nombre, hojaPersonal = null) {
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
    const nombreNormalizado = normalizarNombreCompleto(nombre);

    Logger.log(`🔍 Buscando tipo para: "${nombreNormalizado}"`);
    Logger.log(`📊 Revisando ${datos.length} filas en Personal`);

    // Buscar en todas las columnas de personal
    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];

      for (let col = 0; col <= 5; col++) {
        // Columnas A-F
        const valorCelda = String(fila[col] || "").trim();

        if (valorCelda && sonElMismoPersonal(nombreNormalizado, valorCelda)) {
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
    Logger.log(`❌ Error en obtenerTipoDePersonalMejorado: ${error.message}`);
    return null;
  }
}

/**
 * Función mejorada para normalizar nombres
 * @param {string} nombre - Nombre a normalizar
 * @returns {string} Nombre normalizado
 */
function normalizarNombreCompleto(nombre) {
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
 * Función mejorada para comparar si dos nombres se refieren a la misma persona
 * @param {string} nombre1 - Primer nombre
 * @param {string} nombre2 - Segundo nombre
 * @returns {boolean} true si se refieren a la misma persona
 */
function sonElMismoPersonal(nombre1, nombre2) {
  if (!nombre1 || !nombre2) return false;

  const norm1 = normalizarNombreCompleto(nombre1);
  const norm2 = normalizarNombreCompleto(nombre2);

  // 1. Coincidencia exacta
  if (norm1 === norm2) {
    return true;
  }

  // 2. Uno contiene al otro (para nombres parciales)
  if (norm1.includes(norm2) || norm2.includes(norm1)) {
    return true;
  }

  // 3. Comparación por palabras clave
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
 * Función de búsqueda universal de personal
 * @param {string} nombreBuscar - Nombre a buscar
 * @param {string} tipoPersonal - Tipo específico a buscar (opcional)
 * @returns {Object} Resultados de búsqueda detallados
 */
function buscarPersonalUniversal(nombreBuscar, tipoPersonal = null) {
  try {
    Logger.log(
      `🔍 Búsqueda universal: "${nombreBuscar}", tipo: ${tipoPersonal}`
    );

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaPersonal = ss.getSheetByName("Personal");
    const hojaRegistros = ss.getSheetByName("RegistrosProcedimientos");

    if (!hojaPersonal || !hojaRegistros) {
      throw new Error("No se encontraron las hojas necesarias");
    }

    const nombreNormalizado = normalizarNombreCompleto(nombreBuscar);
    const resultados = {
      coincidenciasExactas: [],
      coincidenciasParciales: [],
      porTipo: {},
      sugerencias: [],
      estadisticas: {
        totalBusquedas: 0,
        fuentesRevisadas: [],
      },
    };

    // Inicializar contadores por tipo
    Object.values(TIPOS_PERSONAL).forEach((tipo) => {
      resultados.porTipo[tipo] = [];
    });

    // Buscar en hoja Personal
    const datosPersonal = hojaPersonal.getDataRange().getValues();
    resultados.estadisticas.fuentesRevisadas.push(
      `Personal (${datosPersonal.length} filas)`
    );

    for (let i = 1; i < datosPersonal.length; i++) {
      const fila = datosPersonal[i];

      for (let col = 0; col <= 5; col++) {
        const nombre = String(fila[col] || "").trim();
        if (!nombre) continue;

        resultados.estadisticas.totalBusquedas++;

        const tipoActual = TIPOS_PERSONAL[col];

        // Filtrar por tipo si se especificó
        if (
          tipoPersonal &&
          tipoActual.toLowerCase() !== tipoPersonal.toLowerCase()
        ) {
          continue;
        }

        if (sonElMismoPersonal(nombreNormalizado, nombre)) {
          const resultado = {
            nombre: nombre,
            tipo: tipoActual,
            fuente: "Personal",
            fila: i + 1,
            columna: col,
            coincidenciaExacta:
              normalizarNombreCompleto(nombre) === nombreNormalizado,
          };

          if (resultado.coincidenciaExacta) {
            resultados.coincidenciasExactas.push(resultado);
          } else {
            resultados.coincidenciasParciales.push(resultado);
          }

          resultados.porTipo[tipoActual].push(resultado);
        }
      }
    }

    // Buscar en RegistrosProcedimientos para sugerencias adicionales
    const datosRegistros = hojaRegistros.getDataRange().getValues();
    resultados.estadisticas.fuentesRevisadas.push(
      `Registros (${datosRegistros.length} filas)`
    );

    const nombresEnRegistros = new Set();
    for (let i = 1; i < datosRegistros.length; i++) {
      const nombre = String(datosRegistros[i][1] || "").trim();
      if (nombre) {
        nombresEnRegistros.add(nombre);
      }
    }

    // Agregar sugerencias de registros
    Array.from(nombresEnRegistros).forEach((nombre) => {
      if (sonElMismoPersonal(nombreNormalizado, nombre)) {
        resultados.sugerencias.push({
          nombre: nombre,
          fuente: "RegistrosProcedimientos",
          coincidenciaExacta:
            normalizarNombreCompleto(nombre) === nombreNormalizado,
        });
      }
    });

    Logger.log(
      `✅ Búsqueda completada: ${resultados.coincidenciasExactas.length} exactas, ${resultados.coincidenciasParciales.length} parciales`
    );

    return resultados;
  } catch (error) {
    Logger.log(`❌ Error en buscarPersonalUniversal: ${error.message}`);
    throw error;
  }
}

/**
 * Función de diagnóstico universal para cualquier personal
 * @param {string} nombrePersonal - Nombre del personal a diagnosticar
 * @returns {Object} Diagnóstico completo
 */
function diagnosticarPersonalUniversal(nombrePersonal) {
  try {
    Logger.log(`🔍 DIAGNÓSTICO UNIVERSAL PARA: "${nombrePersonal}"`);

    const busqueda = buscarPersonalUniversal(nombrePersonal);
    const tipo = obtenerTipoDePersonalMejorado(nombrePersonal);

    const diagnostico = {
      nombreBuscado: nombrePersonal,
      nombreNormalizado: normalizarNombreCompleto(nombrePersonal),
      tipoEncontrado: tipo,
      busqueda: busqueda,
      recomendaciones: [],
      problemasEncontrados: [],
    };

    // Analizar resultados y generar recomendaciones
    if (busqueda.coincidenciasExactas.length === 0) {
      diagnostico.problemasEncontrados.push(
        "No se encontraron coincidencias exactas"
      );

      if (busqueda.coincidenciasParciales.length > 0) {
        diagnostico.recomendaciones.push(
          "Se encontraron coincidencias parciales, revisar nombres similares"
        );
      } else {
        diagnostico.recomendaciones.push(
          "Verificar que el nombre esté escrito correctamente"
        );
        diagnostico.recomendaciones.push(
          "Revisar que el personal esté registrado en la hoja Personal"
        );
      }
    }

    if (busqueda.coincidenciasExactas.length > 1) {
      diagnostico.problemasEncontrados.push(
        "Se encontraron múltiples coincidencias exactas"
      );
      diagnostico.recomendaciones.push(
        "Hay nombres duplicados, consolidar registros"
      );
    }

    // Verificar consistencia entre hojas
    if (
      busqueda.sugerencias.length > 0 &&
      busqueda.coincidenciasExactas.length === 0
    ) {
      diagnostico.problemasEncontrados.push(
        "Existe en RegistrosProcedimientos pero no en Personal"
      );
      diagnostico.recomendaciones.push(
        "Agregar el personal a la hoja Personal"
      );
    }

    Logger.log(`📋 Diagnóstico completado para: ${nombrePersonal}`);

    return diagnostico;
  } catch (error) {
    Logger.log(`❌ Error en diagnosticarPersonalUniversal: ${error.message}`);
    return {
      nombreBuscado: nombrePersonal,
      error: error.message,
      recomendaciones: ["Contactar al administrador del sistema"],
    };
  }
}

// ================= FUNCIONES DE COMPATIBILIDAD =================
// Estas funciones mantienen la compatibilidad con el código existente

function obtenerPersonalFiltrado(columnIndexes) {
  return obtenerPersonalPorTipo(columnIndexes);
}

function obtenerDoctores() {
  return obtenerPersonalEspecifico("doctor");
}

function obtenerAnestesiologos() {
  return obtenerPersonalEspecifico("anestesiologo");
}

function obtenerTecnicos() {
  return obtenerPersonalEspecifico("tecnico");
}

function obtenerTipoDePersonal(nombre, hojaPersonal) {
  return obtenerTipoDePersonalMejorado(nombre, hojaPersonal);
}

// ================= FUNCIONES DE UTILIDAD ADICIONALES =================

/**
 * Función para validar la integridad de la hoja Personal
 * @returns {Object} Reporte de validación
 */
function validarHojaPersonal() {
  try {
    const hoja =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Personal");
    if (!hoja) {
      return { error: "La hoja 'Personal' no existe" };
    }

    const datos = hoja.getDataRange().getValues();
    const reporte = {
      totalFilas: datos.length,
      filasConDatos: datos.length - 1,
      estadisticasPorColumna: {},
      problemasEncontrados: [],
      resumen: "",
    };

    // Analizar cada columna
    Object.entries(TIPOS_PERSONAL).forEach(([col, tipo]) => {
      const colIndex = parseInt(col);
      const nombres = new Set();
      let vacios = 0;

      for (let i = 1; i < datos.length; i++) {
        const valor = String(datos[i][colIndex] || "").trim();
        if (valor) {
          nombres.add(valor);
        } else {
          vacios++;
        }
      }

      reporte.estadisticasPorColumna[tipo] = {
        columna: colIndex,
        nombresUnicos: nombres.size,
        celdasVacias: vacios,
        nombres: Array.from(nombres).sort(),
      };
    });

    // Detectar problemas
    Object.entries(reporte.estadisticasPorColumna).forEach(([tipo, stats]) => {
      if (stats.nombresUnicos === 0) {
        reporte.problemasEncontrados.push(`No hay ${tipo}s registrados`);
      }
    });

    reporte.resumen = `Hoja Personal validada: ${reporte.totalFilas} filas, ${reporte.problemasEncontrados.length} problemas encontrados`;

    return reporte;
  } catch (error) {
    return { error: `Error al validar hoja Personal: ${error.message}` };
  }
}

/**
 * Función para mostrar estadísticas completas del personal
 */
function mostrarEstadisticasPersonal() {
  const validacion = validarHojaPersonal();

  Logger.log("================= ESTADÍSTICAS DE PERSONAL =================");
  Logger.log(validacion.resumen);
  Logger.log("");

  if (validacion.error) {
    Logger.log(`❌ ERROR: ${validacion.error}`);
    return;
  }

  Object.entries(validacion.estadisticasPorColumna).forEach(([tipo, stats]) => {
    Logger.log(`📊 ${tipo.toUpperCase()}:`);
    Logger.log(
      `   Columna: ${stats.columna} (${String.fromCharCode(
        65 + stats.columna
      )})`
    );
    Logger.log(`   Personas únicas: ${stats.nombresUnicos}`);
    Logger.log(`   Celdas vacías: ${stats.celdasVacias}`);
    if (stats.nombres.length > 0) {
      Logger.log(`   Nombres: ${stats.nombres.join(", ")}`);
    }
    Logger.log("");
  });

  if (validacion.problemasEncontrados.length > 0) {
    Logger.log("⚠️ PROBLEMAS ENCONTRADOS:");
    validacion.problemasEncontrados.forEach((problema) => {
      Logger.log(`   - ${problema}`);
    });
  }

  return validacion;
}
