const fs = require('fs');
const path = require('path');
const officeparser = require('officeparser');

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
            } else if (file.endsWith('.doc')) {
                totalFiles++;
                try {
                    await new Promise((resolve, reject) => {
                        officeparser.parseOffice(filePath, (err, data) => {
                            if (err) {
                                console.error(`Error con ${filePath}: ${err.message}`);
                                badFiles++;
                                reject(err);
                            } else {
                                goodFiles++;
                                resolve(data);
                            }
                        });
                    });
                } catch (error) {
                    // Manejo de errores ya hecho en el callback
                }
            }
        }
    };

    await walkDirectory(directory);
    return { totalFiles, goodFiles, badFiles };
};

// Reemplaza 'tu_carpeta' con la ruta a tu carpeta
const directoryPath = 'C:\\Users\\Admin\\Documents\\Documentos_word\\2024';

checkDocuments(directoryPath).then(({ totalFiles, goodFiles, badFiles }) => {
    console.log(`Total de archivos analizados: ${totalFiles}`);
    console.log(`Archivos buenos: ${goodFiles}`);
    console.log(`Archivos malos: ${badFiles}`);
}).catch(error => {
    console.error('Ocurrió un error al analizar los documentos:', error);
});
