const fs = require('fs');
const path = require('path');
const csv = require('csv-parser');
const PizZip = require('pizzip');
const { exec } = require('child_process');
const util = require('util');

const execPromise = util.promisify(exec);

const libreofficePath = `"C:\\Program Files (x86)\\LibreOffice\\program\\soffice"`;

// Función para leer el archivo CSV y obtener los datos
const readCsv = (csvPath) => {
    return new Promise((resolve, reject) => {
        const results = [];
        fs.createReadStream(csvPath)
            .pipe(csv())
            .on('data', (data) => results.push(data))
            .on('end', () => resolve(results))
            .on('error', (err) => reject(err));
    });
};

// Función para cargar el contenido del archivo de Word
const loadWordContent = (filePath) => {
    const content = fs.readFileSync(filePath, 'binary');
    const zip = new PizZip(content);
    const doc = zip.file("word/document.xml").asText();
    return { zip, doc };
};

// Función para reemplazar texto en el contenido del documento
const replaceTextInContent = (content, searchValue, replaceValue) => {
    const regex = new RegExp(searchValue, 'g');
    return content.replace(regex, replaceValue);
};

// Función para generar documentos con los reemplazos
const generateDocuments = (templatePath, csvData, outputDir) => {
    const { zip, doc } = loadWordContent(templatePath);
    const generatedFiles = [];

    csvData.forEach((row) => {
        let modifiedContent = doc;

        // Reemplaza los valores específicos en el contenido del documento
        modifiedContent = replaceTextInContent(modifiedContent, '2000720', row['Referencia Material']);
        modifiedContent = replaceTextInContent(modifiedContent, 'FAR AWAY F', row['Identificador Substancia']);
        modifiedContent = replaceTextInContent(modifiedContent, 'Especificaciones técnicas', 'Ficha Técnica');
        modifiedContent = replaceTextInContent(modifiedContent, 'MEZCLA DE:', 'MEZCLA DE: ' + row['mezcla']);
        console.log("esta es la mezcla", row['mezcla']);

        // Actualiza el contenido en el zip
        zip.file("word/document.xml", modifiedContent);

        const buf = zip.generate({ type: 'nodebuffer' });
        const outputFilePath = path.join(outputDir, `${row['Referencia Material']}.docx`);

        fs.writeFileSync(outputFilePath, buf);
        console.log(`Generated document: ${outputFilePath}`);
        generatedFiles.push(outputFilePath);
    });

    return generatedFiles;
};

// Función para convertir un archivo de Word a PDF
const convertToPdf = async (docxFilePath, outputDir) => {
    const command = `${libreofficePath} --headless --convert-to pdf "${docxFilePath}" --outdir "${outputDir}"`;
    try {
        const { stdout, stderr } = await execPromise(command);
        console.log(`Archivo convertido a PDF: ${docxFilePath}`);
        if (stderr) {
            console.error(`Error en la conversión de ${docxFilePath}: ${stderr}`);
        }
    } catch (error) {
        console.error(`Error al convertir el archivo ${docxFilePath}:`, error);
        return docxFilePath;
    }
    return null;
};

// Función para convertir todos los archivos de Word a PDF secuencialmente
const convertFilesToPdf = async (files, outputDir) => {
    const failedFiles = [];
    for (const file of files) {
        const failedFile = await convertToPdf(file, outputDir);
        if (failedFile) {
            failedFiles.push(failedFile);
        }
    }
    return failedFiles;
};

const main = async () => {
    const templatePath = 'documento.docx'; // Ruta del archivo de Word
    const csvPath = 'filas.csv'; // Ruta del archivo CSV
    const outputDir = './output'; // Directorio de salida para archivos Word
    const outputPdfDir = './outputpdfs'; // Directorio de salida para PDFs

    // Crear los directorios de salida si no existen
    if (!fs.existsSync(outputDir)){
        fs.mkdirSync(outputDir);
    }
    if (!fs.existsSync(outputPdfDir)){
        fs.mkdirSync(outputPdfDir);
    }

    try {
        const csvData = await readCsv(csvPath);
        const generatedFiles = generateDocuments(templatePath, csvData, outputDir);
        const failedFiles = await convertFilesToPdf(generatedFiles, outputPdfDir);

        if (failedFiles.length > 0) {
            console.log('Los siguientes archivos no se pudieron convertir a PDF:');
            failedFiles.forEach(file => console.log(file));
        } else {
            console.log('Todos los archivos se convirtieron correctamente a PDF.');
        }
    } catch (error) {
        console.error("Error processing files:", error);
    }
};

main();
