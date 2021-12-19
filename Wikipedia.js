const request = require("request");
const cheerio = require("cheerio");
const xlsx = require("xlsx");
const fs = require("fs");
const path = require("path");

const url = "https://en.wikipedia.org/wiki/A";
const headers = {};
headers['X-Requested-With'] = 'XMLHttpRequest';
headers['Referer-Policy'] = 'no-referrer-when-downgrade';
headers['referer'] = url;
headers['User-Agent'] = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.110 Safari/537.36';
headers['Access-Control-Allow-Origin'] = '*';
headers['Access-Control-Allow-Methods'] = '*';

const Wiki = path.join(__dirname, "Wikipedia");
dirCreate(Wiki);

request.get({ uri: url, encoding: 'binary', headers: headers }, cb);

function cb(error, response, html) {
    if (error) {
        console.error('Error: ', error);
    } else {
        extractHtml(html);
    }
}

function extractHtml(html) {
    let $ = cheerio.load(html);
    let data = $(".mw-body-content.mw-content-ltr .mw-parser-output");
    let pgroup = data.find("p");
    //let ligroup = data.find("li");
    let ulgroup = data.find("ul");
    let History = $(pgroup[5]).text() + $(pgroup[6]).text() + $(pgroup[7]).text() + $(pgroup[8]).text() + $(pgroup[9]).text() + $(pgroup[10]).text() + $(pgroup[11]).text();
    //let Writing = $(ligroup[70]).text() + $(ligroup[71]).text() + $(ligroup[72]).text() + $(ligroup[73]).text() + $(ligroup[74]).text() + $(ligroup[75]).text() + $(ligroup[76]).text(); 
    let Writing = $(pgroup[12]).text() + $(ulgroup[17]).text() + $(pgroup[13]).text() + $(pgroup[14]).text() + $(pgroup[15]).text() + $(pgroup[16]).text() + $(ulgroup[18]).text();
    let Uses = $(pgroup[17]).text() + $(pgroup[18]).text() + $(pgroup[19]).text() + $(pgroup[20]).text() + $(pgroup[21]).text() + $(pgroup[22]).text();
    console.log("HISTORY:" + "\n" + History + "\n");
    console.log("USES IN WRITING SYSTEMS:" + "\n" + Writing + "\n");
    console.log("OTHER USES:" + "\n" + Uses + "\n");

    processData(History, Writing, Uses);
}

function processData(History, Writing, Uses) {
    let name = "A.xlsx";
    let Page = path.join(Wiki, "Wikipedia.xlsx");
    let content = excelReader(Page, "A");
    let wikiObj = {
        History,
        Writing,
        Uses
    }
    content.push(wikiObj);
    excelWriter(Page, content, name);
}

function dirCreate(filePath) {
    if (fs.existsSync(filePath) == false) {
        fs.mkdirSync(filePath);
    }
}

function excelWriter(filePath, json, sheetName) {
    let newWB = xlsx.utils.book_new();
    let newWS = xlsx.utils.json_to_sheet(json);
    xlsx.utils.book_append_sheet(newWB, newWS, sheetName);
    xlsx.writeFile(newWB, filePath);
}

function excelReader(filePath, sheetName) {
    if(fs.existsSync(filePath) == false) {
        return [];
    }
    let wb = xlsx.readFile(filePath);
    let excelData = wb.Sheets[sheetName];
    let ans = xlsx.utils.sheet_to_json(excelData);
    return ans;
}