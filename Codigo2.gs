/**
 * URL del Logo
 */
const LOGO_URL = "https://centrodigestivointegral.com/wp-content/uploads/2022/11/CDI_Logo-Full-color.png"; 

/**
 * Sirve la página web principal.
 * @return {HtmlOutput}
 */
function doGet() {
  // Configura el título y sirve el archivo Comprobante.html
  return HtmlService.createTemplateFromFile('Comprobante')
      .evaluate()
      .setTitle('Generador de Comprobantes de Asistencia')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

/**
 * Incluye el contenido de archivos HTML parciales (como CSS o JS).
 * Esto permite separar el código y mantenerlo organizado.
 * Se usa para incluir StylesComprobante.html y ScriptComprobante.html
 * @param {string} filename El nombre del archivo a incluir (sin la extensión .html).
 * @return {string} El contenido HTML del archivo.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Genera el texto del comprobante formateado.
 * @param {Object} formData Un objeto con los datos del formulario.
 * @return {string} El HTML formateado y estilizado del comprobante.
 */
function generarComprobante(formData) {
  const { 
    nombre, 
    cedula, 
    tipoCita, 
    especialista, 
    horaInicio, 
    horaFin, 
    fecha,
    generarAcompanante,
    acompanianteNombre, 
    acompanianteCedula 
  } = formData;
  
  // Formatear la fecha (que llega como YYYY-MM-DD) a DD/MM/YYYY
  const [year, month, day] = fecha.split('-');
  const fechaFormateada = `${day}/${month}/${year}`;
  
  let cuerpoComprobante;
  
  // ** MODIFICACIÓN CLAVE: Comprobar si el valor es el string 'true' **
  if (generarAcompanante === 'true') {
    // Caso 1: Comprobante para el ACOMPAÑANTE
    const nombreAcomp = acompanianteNombre.toUpperCase();
    
    cuerpoComprobante = `
      <p>Por medio de la presente se hace constar que el/la señor(a) <strong>${nombreAcomp}</strong>, identificado(a) con cédula <strong>${acompanianteCedula}</strong>, asistió a nuestras instalaciones **como acompañante** del paciente <strong>${nombre.toUpperCase()}</strong>, cédula <strong>${cedula}</strong>, en su cita programada.</p>
      
      <p>Los detalles de la cita a la que asistió son los siguientes:</p>
      
      <div style="margin-left: 20px;">
          <p class="mb-1"><strong>Fecha:</strong> ${fechaFormateada}</p>
          <p class="mb-1"><strong>Tipo de Cita:</strong> ${tipoCita}</p>
          <p class="mb-1"><strong>Especialista:</strong> ${especialista}</p>
          <p class="mb-1"><strong>Horario:</strong> de ${horaInicio} a ${horaFin}</p>
      </div>
      
      <p class="mt-4">La asistencia se registró en el <strong>Centro Digestivo Integral</strong>, Sabana Sur.</p>
    `;
    
  } else {
    // Caso 2: Comprobante para el PACIENTE (estructura original)
    cuerpoComprobante = `
      <p>Por medio de la presente se hace constar que el/la paciente <strong>${nombre.toUpperCase()}</strong>, identificado(a) con cédula <strong>${cedula}</strong>, asistió a una cita en nuestro centro.</p>
      
      <p>Los detalles de la cita son los siguientes:</p>
      
      <div style="margin-left: 20px;">
          <p class="mb-1"><strong>Fecha:</strong> ${fechaFormateada}</p>
          <p class="mb-1"><strong>Tipo de Cita:</strong> ${tipoCita}</p>
          <p class="mb-1"><strong>Especialista:</strong> ${especialista}</p>
          <p class="mb-1"><strong>Horario:</strong> de ${horaInicio} a ${horaFin}</p>
      </div>
      
      <p class="mt-4">La asistencia se registró en el <strong>Centro Digestivo Integral</strong>, Sabana Sur.</p>
    `;
  }
  
  // Contenido base del comprobante que envuelve el cuerpo condicional
  const content = `
    <!-- ENCABEZADO ALINEADO A LA IZQUIERDA CON LOGO -->
    <div class="text-start mb-4" style="text-align: left;">
        <img src="${LOGO_URL}" alt="Logo del Centro" style="max-width: 150px; height: auto; ">
    </div>
    
    <h3 class="mb-5" style="font-family: Arial, sans-serif; text-align: left;"><strong>Comprobante de Asistencia</strong></h3>
    
    <!-- CUERPO DEL TEXTO (TODO ALINEADO A LA IZQUIERDA Y MEJOR ORGANIZADO) -->
    <div style="font-size: 12pt; font-family: Arial, sans-serif; text-align: left;">
        ${cuerpoComprobante}
    </div>
    
    <div class="mt-5 mb-5" style="font-size: 12pt; font-family: Arial, sans-serif; text-align: left;">
        <p>Quedamos atentos a cualquier consulta.</p>
    </div>
    
    <!-- PIE DE PÁGINA CON INFORMACIÓN DE CONTACTO E ICONOS -->
    <div class="d-flex justify-content-between align-items-end pt-5" style="border-top: 1px solid #ccc; margin-top: 40px; font-size: 10pt; font-family: Arial, sans-serif;">
        <!-- LADO IZQUIERDO: Contacto con Iconos (usaremos HTML entities por seguridad) -->
        <div style="text-align: left;">
            <p class="mb-1"><span style="color:#007bff;">&#9990;</span> +506 2102-0846</p> <!-- Teléfono -->
            <p class="mb-1"><span style="color:#007bff;">&#9990;</span> +506 8658-4968</p> <!-- Whatsapp -->
            <p class="mb-1"><span style="color:#dc3545;">&#9993;</span> info@centrodigestivointegral.com</p> <!-- Email -->
            <p class="mb-0"><span style="color:#28a745;">&#127968;</span> Sabana Sur, Costa Rica</p> <!-- Dirección -->
        </div>
        <!-- LADO DERECHO: Logo Pequeño -->
        <div style="text-align: right;">
            <img src="${LOGO_URL}" alt="Logo Footer" style="max-width: 80px; height: auto;">
        </div>
    </div>
  `;

  // Estilo completo con Marca de Agua
  const htmlOutput = `
    <div style="width: 100%; max-width: 650px; margin: 0 auto; padding: 20px; position: relative; height: auto;">
      <!-- MARCA DE AGUA -->
      <img src="${LOGO_URL}" alt="Marca de Agua" style="
        position: absolute; 
        top: 50%; 
        left: 50%; 
        transform: translate(-50%, -50%) rotate(-45deg); 
        opacity: 0.1; 
        width: 300px; 
        height: auto; 
        pointer-events: none;
      ">
      ${content}
    </div>
  `;

  return htmlOutput;
}