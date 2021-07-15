"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.write = exports.read = void 0;
/**
 *
 * @param {String} area - Fetch area "A1:D4"
 * @param {String} sheetname - sheetname
 * @param {String} filename - Excel Worksheet
 * @returns Array
 */
function read(area, sheetname, filename) {
    return __awaiter(this, void 0, void 0, function* () {
        const book = xlsx_1.readFile(filename);
        const ws = book.Sheets[sheetname];
        let arr = [];
        let decodeRange = yield getdecodeRange(area);
        for (let colIdx = decodeRange.s.c, m = 0; colIdx <= decodeRange.e.c; colIdx++, m++) {
            arr[colIdx] = [];
            for (let rowIdx = decodeRange.s.r, n = 0; rowIdx <= decodeRange.e.r; rowIdx++, n++) {
                // セルのアドレスを取得する
                let address = yield getencodeRange({ r: rowIdx, c: colIdx });
                let cell = ws[address];
                let k;
                if (typeof cell == "undefined" || typeof cell.v == "undefined")
                    k = "";
                else if (!isNaN(cell.v))
                    k = Math.round(cell.v * 1000) / 1000;
                else
                    k = cell.v;
                arr[m][n] = k;
            }
        }
        return arr;
    });
}
exports.read = read;
const xlsx_1 = require("xlsx");
/**
 *
 * @param {Array} data data to write in
 * @param {String} area area to write in, eg. A1:D3
 * @param {String}  sheetname sheetname to write
 * @param {String} filename The book of worksheet
 */
function write(data, area, sheetname, filename) {
    return __awaiter(this, void 0, void 0, function* () {
        const book = xlsx_1.readFile(filename);
        const ws = book.Sheets[sheetname];
        const decodeRange = yield getdecodeRange(area);
        for (let colIdx = decodeRange.s.c, m = 0; colIdx <= decodeRange.e.c; colIdx++, m++) {
            for (let rowIdx = decodeRange.s.r, n = 0; rowIdx <= decodeRange.e.r; rowIdx++, n++) {
                const address = yield getencodeRange({ r: rowIdx, c: colIdx });
                if (!data[m][n]) { }
                else if (isNaN(data[m][n])) {
                    ws[address] = {
                        t: 'f',
                        f: data[m][n]
                    };
                }
                else {
                    ws[address] = {
                        t: 'n',
                        v: data[m][n]
                    };
                }
            }
        }
        book.Sheets[sheetname] = ws;
        xlsx_1.writeFile(book, filename);
        return 0;
    });
}
exports.write = write;
/**
 *
 * @param {String} range input of Conversion
 * @returns - { s: { c: start col, r: start row }, e: { c: end col , r: end row } }

 */
function getdecodeRange(range) {
    return __awaiter(this, void 0, void 0, function* () {
        return xlsx_1.utils.decode_range(range);
    });
}
function getencodeRange(a) {
    return __awaiter(this, void 0, void 0, function* () {
        return xlsx_1.utils.encode_cell(a);
    });
}
