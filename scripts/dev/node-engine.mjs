/* The formula engine in Node, from the copy shipped in assets/engine/, for checks that need no browser. */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
const DIR = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../../assets/engine/');
const bg = await import(path.join(DIR, 'formualizer_wasm_bg.js'));
const { instance } = await WebAssembly.instantiate(fs.readFileSync(path.join(DIR, 'formualizer_wasm_bg.wasm')), { './formualizer_wasm_bg.js': bg });
bg.__wbg_set_wasm(instance.exports);
if (instance.exports.__wbindgen_start) instance.exports.__wbindgen_start();
export const api = { Workbook: bg.Workbook };
