// Variable global para almacenar el nombre del plan desde Excel
let nombrePlanDesdeExcel = "";

// Consulta de código desde Excel y actualización visual
async function buscarValor() {
  const codigoSeleccionado = document.getElementById("codigo").value;
  const resultadoBox = document.getElementById("resultado");

  if (!codigoSeleccionado) {
    resultadoBox.textContent = "Selecciona un código para ver el resultado";
    resultadoBox.style.borderLeftColor = "#ccc";
    precioPlanDesdeExcel = "";
    return;
  }

  try {
    const response = await fetch("data.xlsx");
    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });
    const hoja = workbook.Sheets[workbook.SheetNames[0]];
    const datos = XLSX.utils.sheet_to_json(hoja, { header: 1 });

    const fila = datos.find(row => row[0] === codigoSeleccionado);
    nameDesdeExcel = fila ? fila[1] : "";
    vPlanDesdeExcel = fila ? fila[2] : "";
    promoDesdeExcel = fila ? fila[3] : "";
    mesesDesdeExcel = fila ? fila[4] : "";
    detaDesdeExcel = fila ? fila[5] : "";

    const detalles = `${nameDesdeExcel}\n\n${detaDesdeExcel}`;
    
    resultadoBox.textContent = detalles;
    resultadoBox.style.borderLeftColor = fila ? "#007bff" : "#dc3545";

    // if (fila) {
    //   document.getElementById("valorPlan").value = valor;
    // }
  } catch (error) {
    console.error("Error al leer el archivo Excel:", error);
    resultadoBox.textContent = "Error al cargar datos";
    resultadoBox.style.borderLeftColor = "#dc3545";
    vPlanDesdeExcel = "";
  }
}

// 1) Generar y guardar .docx en IndexedDB
document.getElementById("contractForm").addEventListener("submit", async (e) => {
  e.preventDefault();

  const data = {
    NOMBRE: document.getElementById("nombre").value,
    DIRECCION: document.getElementById("direccion").value,
    PLAN: nameDesdeExcel,
    VALOR_PLAN: vPlanDesdeExcel,
    VALOR_PROMO: promoDesdeExcel,
    DURACION: mesesDesdeExcel,
    CICLO: document.getElementById("ciclo").value,
    FECHA: document.getElementById("fecha").value,
  };

  try {
    const content = await loadFile("contrato_template.docx");
    const zip = new PizZip(content);
    const doc = new window.docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
      delimiters: { start: "<<", end: ">>" },
    });

    doc.render(data);
    const blob = doc.getZip().generate({
      type: "blob",
      mimeType:
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });

    await saveContrato(blob);
    console.log("DOCX generado y guardado");
    document.getElementById("preview").innerHTML =
      "<p>Contrato generado y guardado. Pulsa “Visualizar contrato”.</p>";
  } catch (err) {
    console.error("Error generando .docx:", err);
    alert("Error generando contrato. Revisa la consola.");
  }
});

// 2) Renderizar y exportar a PDF (multipágina)
document.getElementById("visualizarButton").addEventListener("click", async () => {
  try {
    const blob = await getContrato();
    if (!blob) {
      alert("No hay contrato generado.");
      return;
    }

    const archivo = new File([blob], "Contrato.docx", { type: blob.type });
    const container = document.getElementById("preview");
    container.innerHTML = "";
    await window.docx.renderAsync(archivo, container);
    console.log("Contrato renderizado en pantalla");

    const imgs = container.querySelectorAll("img");
    if (imgs.length > 1) imgs[1].remove();
    const hdr = container.querySelector("div");
    if (hdr) Object.assign(hdr.style, { margin: "0", padding: "0", float: "none", display: "block" });
    const first = container.firstElementChild;
    if (first) Object.assign(first.style, { margin: "0", padding: "0" });
    Object.assign(container.style, { margin: "0", padding: "0" });

    await new Promise(requestAnimationFrame);

    const capture = document.getElementById("pdf-capture");
    const allImgs = capture.querySelectorAll("img");
    allImgs.forEach(img => (img.crossOrigin = "anonymous"));
    await Promise.all(
      Array.from(allImgs).map(
        img =>
          new Promise(resolve => {
            if (img.complete) return resolve();
            img.onload = resolve;
            img.onerror = resolve;
          })
      )
    );

    console.log("Iniciando html2canvas...");
    const canvas = await html2canvas(capture, {
      scale: 2,
      useCORS: true,
      allowTaint: false,
      scrollX: 0,
      scrollY: -window.scrollY,
      width: capture.offsetWidth,
      height: capture.scrollHeight,
    });
    console.log("Canvas capturado:", canvas.width, "×", canvas.height);

    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF({ unit: "mm", format: "letter", orientation: "portrait" });
    const pageW = pdf.internal.pageSize.getWidth();
    const pageH = pdf.internal.pageSize.getHeight();
    const margin = 5;
    const pdfW = pageW - margin * 2;
    const pdfH = pageH - margin * 2;
    const pxPerMm = canvas.width / pdfW;
    const pagePxH = Math.floor(pdfH * pxPerMm);

    const imgData = canvas.toDataURL("image/jpeg", 1.0);
    let renderedH = 0;
    let pageCount = 0;

    while (renderedH < canvas.height) {
      const fragH = Math.min(pagePxH, canvas.height - renderedH);
      const pageCanvas = document.createElement("canvas");
      pageCanvas.width = canvas.width;
      pageCanvas.height = fragH;
      pageCanvas.getContext("2d").drawImage(
        canvas,
        0,
        renderedH,
        canvas.width,
        fragH,
        0,
        0,
        canvas.width,
        fragH
      );

      const fragImg = pageCanvas.toDataURL("image/jpeg", 1.0);
      if (pageCount > 0) pdf.addPage();
      pdf.addImage(fragImg, "JPEG", margin, margin, pdfW, (fragH / canvas.width) * pdfW);

      renderedH += fragH;
      pageCount++;
    }

    pdf.save("Contrato.pdf");
    console.log("PDF generado en", pageCount, "páginas");
  } catch (err) {
    console.error("Error exportando PDF:", err);
    alert("Error exportando PDF. Revisa la consola.");
  }
});

// Helper para cargar plantilla .docx
function loadFile(url) {
  return new Promise((resolve, reject) => {
    window.PizZipUtils.getBinaryContent(url, (err, data) =>
      err ? reject(err) : resolve(data)
    );
  });
}
