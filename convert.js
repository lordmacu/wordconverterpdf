const { exec } = require('child_process');
const fs = require('fs');
const path = require('path');
const util = require('util');

const execPromise = util.promisify(exec);

const libreofficePath = `"C:\\Program Files (x86)\\LibreOffice\\program\\soffice"`;
const inputDir = "C:\\Users\\cgarcia\\Desktop\\New folder (2)\\output";
const outputDir = "C:\\Users\\cgarcia\\Desktop\\New folder (2)\\outputpdfs";

// Crear el directorio de salida si no existe
if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
}

// Función para convertir un archivo
const convertFile = async (file) => {
    const inputFilePath = path.join(inputDir, file);
    const command = `${libreofficePath} --headless --convert-to pdf "${inputFilePath}" --outdir "${outputDir}"`;

    try {
        const { stdout, stderr } = await execPromise(command);
        console.log(`Archivo convertido: ${file}`);
        if (stderr) {
            console.error(`Error en la conversión de ${file}: ${stderr}`);
        }
        return null; // No error
    } catch (error) {
        console.error(`Error al convertir el archivo ${file}:`, error);
        return file; // Return the file that failed
    }
};

// Función para convertir archivos secuencialmente
const convertFilesSequentially = async (files) => {
    const failedFiles = [];
    for (const file of files) {
        const failedFile = await convertFile(file);
        if (failedFile) {
            failedFiles.push(failedFile);
        }
    }

    // Mostrar archivos que no se pudieron convertir
    if (failedFiles.length > 0) {
        console.log('Los siguientes archivos no se pudieron convertir:');
        failedFiles.forEach(file => console.log(file));
    } else {
        console.log('Todos los archivos se convirtieron correctamente.');
    }
};

// Leer todos los archivos del directorio de entrada
fs.readdir(inputDir, (err, files) => {
    if (err) {
        console.error('Error leyendo el directorio de entrada:', err);
        return;
    }

    // Filtrar solo archivos que no sean carpetas
    files = files.filter(file => fs.lstatSync(path.join(inputDir, file)).isFile());

    // Convertir archivos de forma secuencial
    convertFilesSequentially(files);
});
