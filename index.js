const fs = require('fs');
const XLSX = require('xlsx');

fs.unlink(`${process.cwd()}/that.xlsx`, () => { });

let workbook = XLSX.readFile('this.xlsx');
var new_document = XLSX.utils.book_new();
let deleted = {};
let no_match = {};

workbook.SheetNames.forEach(function (sheetName) {
    deleted[sheetName] = [];
    no_match[sheetName] = [];
    let XL_row_object = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
    let global_obj = {};
    let finall_arr = [];
    XL_row_object.forEach(function (obj) {
        if (!!obj["id"]) {
            global_obj[obj["id"]] = true;
        } else if (!parseInt(obj["extra_ids"], 10) && !parseInt(obj["ids"], 10)) {
            console.log(JSON.stringify(obj, 0, 2), "this obj not taken to global");
        }
    })
    let to_keep_count = Object.keys(global_obj).length;
    let count = 0;
    let pushed = 0;
    let deleted_count = 0;

    XL_row_object.forEach(function (obj) {
        count++;
        if (!!obj["extra_ids"] && !!(global_obj[obj["extra_ids"]])) {
            global_obj[obj["extra_ids"]] = false;
            obj['id'] = obj["extra_ids"];
            finall_arr.push(obj);
            pushed++;
        } else {
            deleted[sheetName].push(obj);
            deleted_count++;
        }
    })
    // console.log(JSON.stringify(global_obj));
    for (id in global_obj) {
        if (!!global_obj[id]) {
            console.log("no match for ", id);
            no_match[sheetName].push({ id: id, sheetName: sheetName });
        }
    }

    let new_sheet = XLSX.utils.json_to_sheet(finall_arr);
    XLSX.utils.book_append_sheet(new_document, new_sheet, sheetName);

    console.log();
    console.log(`-sheetName: ${sheetName}`);
    console.log(`-all rows count: ${count}`);
    console.log(`-to_keep_count: ${to_keep_count}`);
    console.log(`-deleted_count: ${deleted_count}`);
    console.log(`-pushed: ${pushed}`);
    console.log(`-no_match count: ${no_match[sheetName].length}`);
    console.log("------------------------------------------------------------");
})
// create the deleted sheet
var all_Deleted = [];
var all_NOmatch = [];
workbook.SheetNames.forEach(function (sheetName) {
    all_Deleted = all_Deleted.concat(deleted[sheetName].concat([{ sheetName: sheetName }]));
    all_NOmatch = all_NOmatch.concat(no_match[sheetName]);
})
let deleted_sheet = XLSX.utils.json_to_sheet(all_Deleted);
XLSX.utils.book_append_sheet(new_document, deleted_sheet, "deleted_sheet");
let no_match_sheet = XLSX.utils.json_to_sheet(all_NOmatch);
XLSX.utils.book_append_sheet(new_document, no_match_sheet, "no_match_sheet");

// write the xlsx document
XLSX.writeFile(new_document, `${process.cwd()}/that.xlsx`);
