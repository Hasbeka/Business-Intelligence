import path from "path";
import fs from "fs/promises";
import Papa from "papaparse";


export const readCsv = async (filename: string): Promise<string[][]> => {
    const filePath = path.join(process.cwd(), 'data', filename);
    const fileContent = await fs.readFile(filePath, 'utf-8');

    return new Promise((resolve, reject) => {
        Papa.parse(fileContent, {
            header: true,
            dynamicTyping: true,
            skipEmptyLines: true,
            delimitersToGuess: [',', '\t', '|', ';'],
            complete: (results: any) => resolve(results.data),
            error: (error: any) => reject(error)
        });
    });

}

export async function readAllCSVs() {
    const dataDir = path.join(process.cwd(), 'data');
    const files = await fs.readdir(dataDir);
    const csvFiles = files.filter(file => file.endsWith('.csv'));

    const results = await Promise.all(
        csvFiles.map(async (file) => ({
            name: path.parse(file).name,
            data: await readCsv(file)
        }))
    );

    return Object.fromEntries(results.map(r => [r.name, r.data]));
}