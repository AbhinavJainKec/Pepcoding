const request = require("request");
const cheerio = require("cheerio");
const xlsx = require("xlsx");
const fs = require("fs");
const path = require("path");

const url = "https://collegedunia.com/btech/chennai-colleges";
const headers = {};
headers['X-Requested-With'] = 'XMLHttpRequest';
headers['Referer-Policy'] = 'no-referrer-when-downgrade';
headers['referer'] = url;
headers['User-Agent'] = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.110 Safari/537.36';
headers['Access-Control-Allow-Origin'] = '*';
headers['Access-Control-Allow-Methods'] = '*';

//let data = require("./Data.json");
const CollegePath = path.join(__dirname, "Chennai_Colleges");
dirCreate(CollegePath);

request.get({ uri: url, encoding: 'binary', headers: headers }, cb);

function cb(error, response, html) {
    if (error) {
        console.error('Error: ', error);
    } else {
        //console.log(html);
        extractHtml(html);
    }
}

function extractHtml(html) {
    let $ = cheerio.load(html);
    let collegeName = $(".jsx-3025402055.jsx-3200366885.jsx-1185381397 .jsx-765939686.automate_client_img_snippet .jsx-765939686.clg-name-address");
    for (let i = 0; i < 6; i++) {
        let href = $(collegeName[i]).find("a").attr("href");
        let link = "https://collegedunia.com" + href;
        let name = $(collegeName[i]).text();
        console.log(link);
        collegePage(link, name);
    }
}

function collegePage(link, name) {

    request.get({ uri: link, encoding: 'binary', headers: headers }, cb);
    function cb(error, response, link) {
        if (error) {
            console.error('Error: ', error);
        } else {
            extractCollegePage(link, name);
        }
    }
}

function extractCollegePage(link, name) {
    console.log(name);
    let $ = cheerio.load(link);
    let highlights = $(".cdcms_college_highlights");
    let course = $(".jsx-2675951502.fees-info.table-data");
    //console.log($(highlights).find("a").attr("href"));
    let summary = $(highlights).find("p");
    summary = $(summary[0]).text() + $(summary[1]).text().trim();
    console.log(summary);
    let table = $(course).find("tr");
    let courseName = [];
    let fees = [];
    let eligibility = [];
    //console.log(table.length);
    for (let i = 0; i < table.length; i++) {
        let allCol = $(table[i]).find("td");
        if(allCol) {
            courseName[i] = $(allCol[0]).text().trim().replace("   ", " ").replace("\n", " ");
            fees[i] = $(allCol[1]).text().trim().replace("   ", " ").replace("\n", " ");
            eligibility[i] = $(allCol[2]).text().trim().replace("   ", " ").replace("\n", " ");
            console.log(courseName);
            console.log(fees);
            console.log(eligibility);
        }
    }
    processData(name, summary, courseName, fees, eligibility);
}

function processData(name, summary, courseName, fees, eligibility) {
    let initial = name.split(' ');
    let collegename = initial[0] + ".xlsx";
    let college = path.join(CollegePath, "ChennaiColleges.xlsx");
    let content = excelReader(college, name);
    let collegeObj = {
        name,
        summary,
        courseName,
        fees,
        eligibility
    }
    content.push(collegeObj);
    excelWriter(college, content, collegename);
}

function dirCreate(filePath) {
    if (fs.existsSync(filePath) == false) {
        fs.mkdirSync(filePath);
    }
}

function excelWriter(filePath, json, sheetName) {
    let newWB = xlsx.readFile(filePath);
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