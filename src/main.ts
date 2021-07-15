const filename = 'data.xlsx'
let sheetname:string
const area = "B3:H120"
import { write } from './module'
//--------settings
const Settings = {
    Expectedtile: 3,
    Yonma:true
}
//
async function sevenpair(phase:number) {
    let round=0, prob=null, remains, player_num
    if (Settings.Yonma) remains = 122, sheetname = 'Sheet2', player_num=4
    else remains = 94, sheetname = 'Sheet1', player_num=3
    while(phase>=0){
        do {
            if(round>118) return
            prob = Settings.Expectedtile*(2 * phase + 1) / (remains)
            if(!phase) prob = player_num*2*(2 * phase + 1) / (remains)
            remains--
            round++
        } while (prob<(Math.random()))
        phase--
    }    
    return round - 1
}
//-----
let data = new Array(7)

async function main() {
    for (let n = 0; n < 7; n++) {
        data[n] = new Array(117).fill(0);
        for (let m = 0; m < 100000; m++) {
            const k = await sevenpair(n)
            if (k) data[n][k]++
            //else throw new Error("no respond");
        }

    }
    console.log(data);
    write(data, area, sheetname, filename)
}
main()
