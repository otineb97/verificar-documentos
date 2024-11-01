const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

const checkDocuments = async (directory) => {
    let goodFiles = 0;
    let badFiles = 0;
    let totalFiles = 0;

    const walkDirectory = async (dir) => {
        const files = fs.readdirSync(dir);
        for (const file of files) {
            const filePath = path.join(dir, file);

            if (fs.statSync(filePath).isDirectory()) {
                await walkDirectory(filePath);
            } else if (file.endsWith('.xls') || file.endsWith('.xlsx')) {
                totalFiles++;
                try {
                    const workbook = xlsx.readFile(filePath);
                    goodFiles++;
                } catch (error) {
                    console.error(`Error con ${filePath}: ${error.message}`);
                    badFiles++;
                }
            }
        }
    };

    await walkDirectory(directory);
    return { totalFiles, goodFiles, badFiles };
};

// Reemplaza 'tu_carpeta' con la ruta a tu carpeta
const directoryPath = 'C:\\Users\\Admin\\Documents\\Documentos_excel\\2024';

checkDocuments(directoryPath).then(({ totalFiles, goodFiles, badFiles }) => {
    console.log(`Total de archivos analizados: ${totalFiles}`);
    console.log(`Archivos buenos: ${goodFiles}`);
    console.log(`Archivos malos: ${badFiles}`);
}).catch(error => {
    console.error('Ocurrió un error al analizar los documentos:', error);
});
