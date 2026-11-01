import { readdir } from 'node:fs/promises';
import { join } from 'node:path';

const root = 'dist';
const forbidden = /\.(?:xls|xlsx|xlsm|xlsb)$/i;

async function walk(directory) {
  const entries = await readdir(directory, { withFileTypes: true });
  const files = [];
  for (const entry of entries) {
    const path = join(directory, entry.name);
    if (entry.isDirectory()) files.push(...await walk(path));
    else files.push(path);
  }
  return files;
}

const files = await walk(root);
const spreadsheets = files.filter((path) => forbidden.test(path));
if (spreadsheets.length) {
  console.error(`El artefacto público contiene libros de Excel: ${spreadsheets.join(', ')}`);
  process.exitCode = 1;
} else {
  console.log('OK: el artefacto de producción no incluye libros de datos .xls/.xlsx.');
}
