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
const filename = 'data.xlsx';
let sheetname;
const area = "B3:H120";
require("./module");
const module_1 = require("./module");
//--------settings
const Settings = {
    Expectedtile: 3,
    Yonma: true
};
//
function sevenpair(phase) {
    return __awaiter(this, void 0, void 0, function* () {
        let round = 0, prob = null, remains, player_num;
        if (Settings.Yonma)
            remains = 122, sheetname = 'Sheet2', player_num = 4;
        else
            remains = 94, sheetname = 'Sheet1', player_num = 3;
        while (phase >= 0) {
            do {
                if (round > 118)
                    return;
                prob = Settings.Expectedtile * (2 * phase + 1) / (remains);
                if (!phase)
                    prob = player_num * 2 * (2 * phase + 1) / (remains);
                remains--;
                round++;
            } while (prob < (Math.random()));
            phase--;
        }
        return round - 1;
    });
}
//-----
let data = new Array(7);
function main() {
    return __awaiter(this, void 0, void 0, function* () {
        for (let n = 0; n < 7; n++) {
            data[n] = new Array(117).fill(0);
            for (let m = 0; m < 100000; m++) {
                const k = yield sevenpair(n);
                if (k)
                    data[n][k]++;
                //else throw new Error("no respond");
            }
        }
        console.log(data);
        module_1.write(data, area, sheetname, filename);
    });
}
main();
