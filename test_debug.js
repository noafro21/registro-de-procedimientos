// Simulando el test en Apps Script
function testLocal() {
  console.log("Testing buscarBiopsiasServidor_v2 locally");
  
  // Función de normalización
  const normalizarIdentificacion = (str) =>
    String(str || "")
      .replace(/[-\s]/g, "")
      .toUpperCase();
  
  // Test de normalización
  const test1 = normalizarIdentificacion("8-0118-0052");
  const test2 = normalizarIdentificacion("801180052");
  const test3 = normalizarIdentificacion("8 0118 0052");
  
  console.log("🧪 Normalizaciones:");
  console.log(`"8-0118-0052" -> "${test1}"`);
  console.log(`"801180052" -> "${test2}"`);
  console.log(`"8 0118 0052" -> "${test3}"`);
  
  console.log("¿Son iguales?", test1 === test2, test2 === test3);
}

testLocal();
