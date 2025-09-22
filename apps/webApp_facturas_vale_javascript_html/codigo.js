// Variable para almacenar los datos consolidados de todas las facturas
let allData = [];

// Función para procesar un solo archivo XML
function processXmlContent(xmlContent, zipFileName) {
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(xmlContent, "text/xml");
    const raiz = xmlDoc.documentElement;

    // Verificar si el documento XML es válido
    if (!raiz || raiz.tagName === 'parsererror') {
        console.warn(`Archivo XML corrupto en el ZIP: ${zipFileName}`);
        return null;
    }

    // Definir los espacios de nombres del XML
    const namespaces = {
        cfdi: 'http://www.sat.gob.mx/cfd/4',
        tfd: 'http://www.sat.gob.mx/TimbreFiscalDigital',
    };

    // Helper para buscar elementos con espacios de nombres
    const findElement = (root, tag, ns) => {
        const fullTag = ns ? `${ns.prefix}:${tag}` : tag;
        return root.getElementsByTagName(fullTag)[0] || null;
    };
    
    // Extraer datos clave del XML
    const complemento = raiz.querySelector('Complemento > TimbreFiscalDigital') || findElement(raiz, 'TimbreFiscalDigital', { prefix: 'tfd', uri: namespaces.tfd });
    
    if (!complemento) {
        return null; // No es una factura válida sin timbre
    }

    const tipoDeComprobante = raiz.getAttribute('TipoDeComprobante');
    const impuestosNode = raiz.querySelector('Impuestos') || findElement(raiz, 'Impuestos', { prefix: 'cfdi', uri: namespaces.cfdi });

    if (!['I', 'E'].includes(tipoDeComprobante) || !impuestosNode) {
        return null; // No es un comprobante de Ingreso/Egreso con impuestos
    }
    
    const emisor = raiz.querySelector('Emisor') || findElement(raiz, 'Emisor', { prefix: 'cfdi', uri: namespaces.cfdi });

    const serie = raiz.getAttribute('Serie');
    const folio = raiz.getAttribute('Folio');
    const uuid = complemento.getAttribute('UUID');
    
    // Construir el número de factura
    let factura = uuid.slice(-12);
    if (folio) {
        factura = serie ? `${serie} ${folio}` : folio;
    }

    // Extraer impuestos de forma segura
    const ivaTrasladadoNode = impuestosNode.querySelector('Traslados > Traslado') || findElement(impuestosNode, 'Traslado', { prefix: 'cfdi', uri: namespaces.cfdi });
    const retencionNode = impuestosNode.querySelector('Retenciones > Retencion') || findElement(impuestosNode, 'Retencion', { prefix: 'cfdi', uri: namespaces.cfdi });
    
    const iva = ivaTrasladadoNode ? ivaTrasladadoNode.getAttribute('Importe') : 0;
    const retencion = retencionNode ? retencionNode.getAttribute('Importe') : 0;
    
    // Retornar un objeto con los datos extraídos
    return {
        FACTURA: factura,
        FECHA: complemento.getAttribute('FechaTimbrado'),
        RFC: emisor.getAttribute('Rfc'),
        NOMBRE: emisor.getAttribute('Nombre'),
        SUBTOTAL: parseFloat(raiz.getAttribute('SubTotal')) || 0,
        IVA: parseFloat(iva) || 0,
        RETENCION: parseFloat(retencion) || 0,
        TOTAL: parseFloat(raiz.getAttribute('Total')) || 0,
        UUID: uuid,
        REFERENCIA: `${emisor.getAttribute('Rfc')}*${uuid.slice(0, 8)}*${zipFileName.split(".zip")[0]  }`,
        ORIGEN_CARPETA: zipFileName,
    };
}

// Función principal para procesar los archivos ZIP
async function processZipFiles(files) {
    const statusDiv = document.getElementById('status');
    const spinner = document.getElementById('spinner');
    const downloadButton = document.getElementById('download-button');

    allData = [];
    statusDiv.textContent = `Procesando ${files.length} archivos ZIP...`;
    spinner.style.display = 'block';
    downloadButton.style.display = 'none';

    for (const file of files) {
        try {
            const zip = await JSZip.loadAsync(file);
            const zipFileName = file.name;
            const xmlPromises = [];

            zip.forEach((relativePath, zipEntry) => {
                if (relativePath.toLowerCase().endsWith('.xml') && !zipEntry.dir) {
                    // Crea una promesa para leer el contenido del XML
                    xmlPromises.push(zipEntry.async('string').then(xmlContent => {
                        return processXmlContent(xmlContent, zipFileName);
                    }));
                }
            });

            const results = await Promise.all(xmlPromises);
            const validResults = results.filter(result => result !== null);
            allData.push(...validResults);

        } catch (error) {
            console.error(`Error al procesar el archivo ${file.name}:`, error);
            statusDiv.textContent = `Error al procesar ${file.name}. El archivo podría estar dañado.`;
            spinner.style.display = 'none';
            return;
        }
    }

    spinner.style.display = 'none';

    if (allData.length > 0) {
        statusDiv.textContent = `Procesamiento completado. Se extrajeron ${allData.length} facturas.`;
        downloadButton.style.display = 'block';
    } else {
        statusDiv.textContent = 'No se encontraron facturas válidas en los archivos seleccionados.';
        downloadButton.style.display = 'none';
    }
}

// Función para convertir los datos a formato CSV.
function convertToCSV(data) {
    if (data.length === 0) {
        return '';
    }
    const headers = Object.keys(data[0]);
    const csvRows = [
        headers.join(','),
        ...data.map(row => headers.map(header => {
            const value = row[header];
            // Manejar valores que contienen comas o comillas
            return `"${String(value).replace(/"/g, '""')}"`;
        }).join(','))
    ];
    return csvRows.join('\n');
}

// Lógica principal de la aplicación web.
document.addEventListener('DOMContentLoaded', () => {
    const fileInput = document.getElementById('fileInput');
    const downloadButton = document.getElementById('download-button');

    fileInput.addEventListener('change', (event) => {
        const files = event.target.files;
        if (files.length > 0) {
            processZipFiles(files);
        }
    });

    downloadButton.addEventListener('click', () => {
        if (allData.length === 0) {
            alert('No hay datos para descargar.');
            return;
        }
        
        const csvContent = convertToCSV(allData);
        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'reporte_facturas.csv';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    });
});