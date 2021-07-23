//*** グローバル変数 ***
//開いているシート情報
const bk = SpreadsheetApp.getActiveSpreadsheet();
const all_sh_cnt = bk.getNumSheets();

//アクティブシート情報
const act_sh = bk.getActiveSheet();
const act_sh_cnt = act_sh.getIndex();

//GSS の起動時に実行
function onOpen() {

    const myMenu = [
        {name: "ホームへ", functionName: "activate_first_sheet"},
        {name: "1日進む", functionName: "move_next_sheet"},
        {name: "1日戻る", functionName: "move_prev_sheet"},
        {name: "最新日へ", functionName: "activate_last_sheet"},
    ];
    
    SpreadsheetApp.getActiveSpreadsheet().addMenu("追加メニュー",myMenu); //メニューを追加
}

//アクティブシート切り替え
function activate_first_sheet() {
    bk.getSheets()[0].activate(); //配列なので0始まり
}

function move_next_sheet() {
    const nxt_sh = bk.getSheets()[act_sh_cnt]; //0始まりなのでアクティブセルのインデックスが次のセル

    if (act_sh_cnt + 1 <= all_sh_cnt){
        nxt_sh.activate();
    }
}

function move_prev_sheet() {
    const prev_sh = bk.getSheets()[act_sh_cnt - 2]; //0始まりなので2減らす必要がある。

    if (act_sh_cnt - 2 >= 0){
        prev_sh.activate();
    }
}

function activate_last_sheet() {
    bk.getSheets()[all_sh_cnt - 1].activate();
}