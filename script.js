const inputArchivo = document.getElementById('cargarImagen');
const zonaPegado = document.getElementById('zonaPegado');
const resultado = document.getElementById('resultado');
const botonExtraer = document.getElementById('extraerTexto');
const botonExportar = document.getElementById('exportarTexto');

let imagenActual = null;
let textoExtraido = '';

// Función para cargar imagen en objeto HTMLImageElement
function cargarImagen(src) {
    return new Promise((resolve, reject) => {
        const img = new Image();
        img.onload = () => resolve(img);
        img.onerror = reject;
        img.src = src;
    });
}

// Preprocesado optimizado para mejorar precisión del OCR
async function preprocesarImagen(src) {
    const imagen = await cargarImagen(src);

    // Escalado 4x para mejorar DPI efectivo 
    const escala = 4;
    const ancho = Math.max(1, Math.floor(imagen.naturalWidth * escala));
    const alto = Math.max(1, Math.floor(imagen.naturalHeight * escala));

    const canvas = document.createElement('canvas');
    canvas.width = ancho;
    canvas.height = alto;
    const ctx = canvas.getContext('2d');

    // Dibujar con re-muestreo de alta calidad
    ctx.imageSmoothingEnabled = true;
    ctx.imageSmoothingQuality = 'high';
    ctx.drawImage(imagen, 0, 0, ancho, alto);

    const imageData = ctx.getImageData(0, 0, ancho, alto);
    const data = imageData.data;

    // Parámetros más agresivos para mejor distinción de caracteres
    const contraste = 1.5; // Aumentado de 1.3 a 1.5 para mejor distinción
    const brillo = 12; // Aumentado de 8 a 12 para resaltar más el texto
    const umbralAdaptativo = true;

    // Calcular umbral adaptativo más preciso
    let umbral = 128;
    if (umbralAdaptativo) {
        let suma = 0;
        let count = 0;
        let valores = [];
        
        // Recopilar todos los valores de gris para análisis estadístico
        for (let i = 0; i < data.length; i += 4) {
            const gris = 0.299 * data[i] + 0.587 * data[i + 1] + 0.114 * data[i + 2];
            if (gris > 30 && gris < 220) { // Rango más amplio
                valores.push(gris);
                suma += gris;
                count++;
            }
        }
        
        if (count > 0) {
            // Usar percentil 75 para umbral más preciso
            valores.sort((a, b) => a - b);
            const percentil75 = valores[Math.floor(valores.length * 0.75)];
            umbral = Math.min(190, Math.max(110, percentil75 + 10));
        }
    }

    // Aplicar preprocesado pixel por pixel con filtro de suavizado
    for (let i = 0; i < data.length; i += 4) {
        const r = data[i];
        const g = data[i + 1];
        const b = data[i + 2];

        // Convertir a escala de grises
        let gris = 0.299 * r + 0.587 * g + 0.114 * b;

        // Aplicar contraste y brillo más agresivos
        gris = (gris - 128) * contraste + 128 + brillo;
        gris = Math.max(0, Math.min(255, gris));

        // Binarización con umbral adaptativo
        const valorBinario = gris >= umbral ? 255 : 0;

        // Aplicar el valor procesado
        data[i] = valorBinario;
        data[i + 1] = valorBinario;
        data[i + 2] = valorBinario;
        data[i + 3] = 255; // Alpha channel
    }

    ctx.putImageData(imageData, 0, 0);
    return canvas.toDataURL('image/png');
}

// Función para crear y descargar documento de Word
async function exportarAWord(texto) {
    try {
        // Verificar que la librería esté disponible
        if (typeof docx === 'undefined') {
            // Fallback: exportar como archivo de texto que se puede abrir en Word
            return exportarComoTexto(texto);
        }

        // Crear documento usando la librería docx
        const doc = new docx.Document({
            sections: [{
                properties: {},
                children: [
                    new docx.Paragraph({
                        children: [
                            new docx.TextRun({
                                text: "Texto Extraído por OCR",
                                bold: true,
                                size: 28
                            })
                        ],
                        alignment: docx.AlignmentType.CENTER,
                        spacing: {
                            after: 200
                        }
                    }),
                    new docx.Paragraph({
                        children: [
                            new docx.TextRun({
                                text: texto,
                                size: 24,
                                font: "Arial"
                            })
                        ],
                        spacing: {
                            line: 360
                        }
                    })
                ]
            }]
        });

        // Generar y descargar el archivo
        const blob = await docx.Packer.toBlob(doc);
        const url = window.URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = `texto_extraido_${new Date().toISOString().slice(0, 19).replace(/:/g, '-')}.docx`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        window.URL.revokeObjectURL(url);
        
        resultado.textContent = 'Documento Word exportado exitosamente!';
    } catch (error) {
        console.error('Error al exportar a Word:', error);
        // Fallback: exportar como texto
        resultado.textContent = 'Exportando como archivo de texto...';
        exportarComoTexto(texto);
    }
}

// Función alternativa para exportar como archivo de texto
function exportarComoTexto(texto) {
    try {
        const contenido = `TEXTO EXTRAÍDO POR OCR\n\n${texto}`;
        const blob = new Blob([contenido], { type: 'text/plain;charset=utf-8' });
        const url = window.URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = `texto_extraido_${new Date().toISOString().slice(0, 19).replace(/:/g, '-')}.txt`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        window.URL.revokeObjectURL(url);
        
        resultado.textContent = 'Archivo de texto exportado exitosamente! (Se puede abrir en Word)';
    } catch (error) {
        console.error('Error al exportar como texto:', error);
        resultado.textContent = `Error al exportar: ${error.message}`;
    }
}

// Función para mostrar imagen en la zona de pegado
function mostrarImagen(imagenSrc) {
    // Limpiar la zona de pegado
    zonaPegado.innerHTML = '';
    
    // Crear elemento de imagen
    const img = document.createElement('img');
    img.src = imagenSrc;
    img.alt = 'Imagen cargada';
    
    // Agregar la imagen a la zona de pegado
    zonaPegado.appendChild(img);
    
    // Agregar clase para ocultar instrucciones
    zonaPegado.classList.add('con-imagen');
    
    // Guardar referencia de la imagen
    imagenActual = imagenSrc;
    
    // Habilitar botón de extraer
    botonExtraer.disabled = false;
    botonExportar.disabled = true; // Deshabilitar exportar hasta extraer texto
}

// Función para procesar archivo de imagen
function procesarArchivo(archivo) {
    if (archivo && archivo.type.startsWith('image/')) {
        const lector = new FileReader();
        lector.onload = (ev) => {
            mostrarImagen(ev.target.result);
        };
        lector.readAsDataURL(archivo);
    } else {
        alert('Por favor selecciona un archivo de imagen válido');
    }
}

// 1. Carga de archivo mediante botón
inputArchivo.addEventListener('change', (e) => {
    const archivo = e.target.files[0];
    if (archivo) {
        procesarArchivo(archivo);
    }
});

// 2. Pegado desde portapapeles
zonaPegado.addEventListener('paste', (e) => {
    e.preventDefault();
    
    const items = e.clipboardData.items;
    for (let item of items) {
        if (item.type.indexOf('image') !== -1) {
            const archivo = item.getAsFile();
            procesarArchivo(archivo);
            break;
        }
    }
});

// 3. Arrastrar y soltar
zonaPegado.addEventListener('dragover', (e) => {
    e.preventDefault();
    zonaPegado.classList.add('drag-over');
});

zonaPegado.addEventListener('dragleave', (e) => {
    e.preventDefault();
    zonaPegado.classList.remove('drag-over');
});

zonaPegado.addEventListener('drop', (e) => {
    e.preventDefault();
    zonaPegado.classList.remove('drag-over');
    
    const archivos = e.dataTransfer.files;
    if (archivos.length > 0) {
        procesarArchivo(archivos[0]);
    }
});

// 4. Hacer la zona de pegado clickeable para abrir selector de archivos
zonaPegado.addEventListener('click', () => {
    if (!imagenActual) {
        inputArchivo.click();
    }
});

// 5. Extraer texto con preprocesado optimizado
botonExtraer.addEventListener('click', async () => {
    if (!imagenActual) {
        alert('No hay imagen seleccionada');
        return;
    }
    
    // Deshabilitar botón durante el proceso
    botonExtraer.disabled = true;
    botonExtraer.textContent = 'Extrayendo...';
    
    resultado.textContent = 'Preprocesando imagen para mejor precisión...';
    
    try {
        // Aplicar preprocesado optimizado
        const imagenPreprocesada = await preprocesarImagen(imagenActual);
        
        resultado.textContent = 'Reconociendo texto con alta precisión...';
        
        // Configuración optimizada de Tesseract para mejor precisión
        const { data } = await Tesseract.recognize(imagenPreprocesada, 'spa+eng', {
            // Parámetros optimizados para precisión y distinción de caracteres
            tessedit_pageseg_mode: 6, // PSM 6: bloque uniforme (mejor para párrafos)
            user_defined_dpi: 400, // Aumentado de 300 a 400 para mejor resolución
            preserve_interword_spaces: '1',
            tessedit_char_whitelist: 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyzÁÉÍÓÚáéíóúÑñ0123456789.,;:!?()[]{}"\'-_@#$%&*+=<>/\\|~`',
            tessedit_do_invert: '0', // No invertir colores
            textord_heavy_nr: '1', // Mejor detección de líneas
            textord_min_linesize: '2.5', // Aumentado para mejor detección
            edges_max_children: '50', // Aumentado para mejor detección de bordes
            edges_children_per_grandchild: '15', // Optimización de detección
            edges_children_count_limit: '60', // Aumentado para mejor detección
            edges_min_children: '5', // Aumentado para mejor detección
            edges_max_children_per_grandchild: '20', // Optimización adicional
            // Parámetros específicos para mejor distinción de caracteres
            tessedit_ocr_engine_mode: '3', // Modo neural network LSTM
            lstm_use_matrix: '1', // Usar matriz LSTM para mejor precisión
            lstm_choice_mode: '2', // Modo de elección más estricto
            lstm_use_1d: '1', // Usar LSTM 1D para mejor reconocimiento
            // Parámetros de calidad
            textord_min_linesize: '2.5', // Tamaño mínimo de línea más estricto
            textord_heavy_nr: '1', // Detección pesada de líneas
            textord_min_xheight: '8', // Altura mínima de caracteres
            textord_old_xheight: '0', // Usar nuevo algoritmo de altura
            textord_min_blob_height: '8', // Altura mínima de blobs
            textord_min_blob_width: '3' // Ancho mínimo de blobs
        }, {
            // Logger para mostrar progreso en tiempo real
            logger: m => {
                if (m.status === 'recognizing text') {
                    const porcentaje = Math.round(m.progress * 100);
                    resultado.textContent = `Reconociendo texto... ${porcentaje}%`;
                    botonExtraer.textContent = `Extrayendo (${porcentaje}%)`;
                } else if (m.status === 'loading tesseract core') {
                    resultado.textContent = 'Cargando motor OCR...';
                } else if (m.status === 'loading language traineddata') {
                    resultado.textContent = 'Cargando idioma español...';
                } else if (m.status === 'initializing tesseract') {
                    resultado.textContent = 'Inicializando Tesseract...';
                } else if (m.status === 'loading image') {
                    resultado.textContent = 'Cargando imagen...';
                } else if (m.status === 'recognizing text') {
                    const porcentaje = Math.round(m.progress * 100);
                    resultado.textContent = `Reconociendo texto... ${porcentaje}%`;
                    botonExtraer.textContent = `Extrayendo (${porcentaje}%)`;
                }
            }
        });

        textoExtraido = data.text.trim();
        if (textoExtraido) {
            // Limpieza mínima para conservar estructura del párrafo
            const limpio = textoExtraido
                .replace(/[\u2010\u2011\u2012\u2013\u2014]/g, '-') // Normalizar guiones
                .replace(/[ \t]+/g, ' ') // Múltiples espacios a uno solo
                .replace(/\n\s+/g, '\n') // Eliminar espacios al inicio de líneas
                .replace(/\n{3,}/g, '\n\n') // Máximo dos saltos de línea consecutivos
                .trim();
            
            resultado.textContent = limpio;
            botonExportar.disabled = false; // Habilitar botón de exportar
        } else {
            resultado.textContent = "No se detectó texto en la imagen";
            botonExportar.disabled = true;
        }
    } catch (err) {
        resultado.textContent = `Error al extraer texto: ${err.message}`;
        botonExportar.disabled = true;
    } finally {
        // Restaurar botón
        botonExtraer.disabled = false;
        botonExtraer.textContent = 'Extraer texto';
    }
});

// 6. Exportar texto a Word
botonExportar.addEventListener('click', () => {
    if (textoExtraido) {
        botonExportar.disabled = true;
        botonExportar.textContent = 'Exportando...';
        exportarAWord(textoExtraido).finally(() => {
            botonExportar.disabled = false;
            botonExportar.textContent = 'Exportar texto';
        });
    } else {
        alert('No hay texto para exportar. Primero extrae el texto de una imagen.');
    }
});

// Inicialización: deshabilitar botones hasta que haya imagen y texto
botonExtraer.disabled = true;
botonExportar.disabled = true;













