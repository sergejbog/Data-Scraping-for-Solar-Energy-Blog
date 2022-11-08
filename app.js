const https = require("https");
const Excel = require("exceljs");
const parser = require("node-html-parser");
const cheerio = require("cheerio");
const fs = require("fs");
const { isNumberObject } = require("util/types");

const companyData = JSON.parse(JSON.stringify(require("./dataPartial2.json")));
console.log();

// const xpath = require("xpath-html");

function httpRequest(url, headers = {}) {
   return new Promise(function (resolve, reject) {
      var req = https.get(url, { headers }, function (res) {
         if (res.statusCode < 200 || res.statusCode >= 300) {
            return reject(new Error("statusCode=" + res.statusCode));
         }
         var body = [];
         res.on("data", function (chunk) {
            body.push(chunk);
         });
         res.on("end", function () {
            try {
               body = Buffer.concat(body).toString();
            } catch (e) {
               reject(e);
            }
            resolve(body);
         });
      });
      req.on("error", function (err) {
         reject(err);
      });
      req.end();
   });
}
function sleep(ms) {
   return new Promise((resolve) => setTimeout(resolve, ms));
}

function parseData($) {
   let year = 2021;
   let insideData = {};
   $("table > tbody table tbody tr").each((_, elem) => {
      let leftArr = [];
      if ($(elem).find("h3").text().includes("STATISTICS")) {
         year = parseInt($(elem).find("h3").text());
      }
      $(elem)
         .find(".label")
         .each((i, innerElem) => {
            // if (
            //    $(innerElem)
            //       .text()
            //       .includes(`${year - 1}`)
            // )
            //    year--;
            leftArr.push($(innerElem).text());
            //    console.log($(innerElem).text());
            if (insideData[year] == undefined) {
               insideData[year] = {
                  [$(innerElem).text()]: 1,
               };
            } else {
               insideData[year][$(innerElem).text()] = 1;
            }
         });
      let i = 0;
      $(elem)
         .find("strong, b")
         .each((_, innerElem) => {
            //    console.log($(innerElem).text());
            insideData[year][leftArr[i]] = $(innerElem).text();
            i++;
         });

      let isItNext = false;
      let countryName;
      $(elem)
         .find("td")
         .each((_, innerElem) => {
            if (isItNext) {
               insideData[year][countryName] = $(innerElem).text().trim().toLowerCase();
               isItNext = false;
            }
            countryName = $(innerElem).text().trim().toLowerCase();
            if (states.includes(countryName)) {
               //   console.log($(innerElem).text());
               isItNext = true;
            }
         });
   });
   return insideData;
}

async function dataByCompany(response, isCountry) {
   let root = parser.parse(response);
   let thead = root.querySelectorAll("table.table thead th").map((elem) => {
      return elem.innerText;
   });
   let tableArr = root.querySelectorAll("tbody tr");

   for (let i = 0; i < tableArr.length; i++) {
      let tableElements = tableArr[i].querySelectorAll("td");
      //   if ((isCountry || companyNamesArr.includes(tableElements[1].innerText.trim().toLowerCase())) && !tableElements[3].innerText.trim().includes("Res")) continue;
      //   if (companyNamesArr.includes(tableElements[1].innerText.trim().toLowerCase())) continue;
      //   if (!tableElements[3].innerText.trim().includes("Res")) continue;
      //   console.log("da");

      let moreDetailsLink = tableArr[i].querySelector("a").getAttribute("href");
      let detailsResponse = await httpRequest(moreDetailsLink);
      let $ = cheerio.load(detailsResponse);
      //   let detailsRootXPath = xpath.fromPageSource(detailsResponse).findElement("//*[contains(text(), 'with love')]");
      let insideData = parseData($);
      //   console.log(insideData);

      await sleep(200);

      companyNamesArr.push(tableElements[1].innerText.trim().toLowerCase());

      if (isCountry) {
         companyData[tableElements[1].innerText.trim()] = {
            [thead[3]]: tableElements[3].innerText.trim(),
            [thead[4]]: tableElements[4].innerText.trim(),
            [thead[5]]: tableElements[5].innerText.trim(),
            [thead[6]]: tableElements[6].innerText.trim(),
            detailsLink: moreDetailsLink,
            ...insideData,
         };
      } else {
         companyData[tableElements[1].innerText.trim()] = {
            [thead[0]]: tableElements[0].innerText.trim(),
            [thead[2]]: tableElements[2].innerText.trim(),
            [thead[3]]: tableElements[3].innerText.trim(),
            [thead[4]]: tableElements[4].innerText.trim(),
            [thead[5]]: tableElements[5].innerText.trim(),
            [thead[6]]: tableElements[6].innerText.trim(),
            detailsLink: moreDetailsLink,
            ...insideData,
         };
      }

      console.log(tableElements[1].innerText.trim());
   }
}

// const companyData = {};
// let companyNamesArr = Object.keys(companyData).map((elem) => {
//    return elem.trim().toLowerCase();
// });
companyNamesArr = [];

console.log(companyNamesArr);
const states = Buffer.from(fs.readFileSync("states.txt"))
   .toString()
   .split("\n")
   .map((elem) => {
      return elem.trim().replace("\r", "").toLowerCase();
   });

async function main() {
   let response = await httpRequest("https://www.solarpowerworldonline.com/2022-top-residential-solar-contractors/#");
   await dataByCompany(response, false);

   fs.writeFileSync("dataPartial2.json", JSON.stringify(companyData));
   console.log(Object.keys(companyData).length);

   //    response = await httpRequest("https://www.solarpowerworldonline.com/2022-top-solar-contractors-by-state/");
   //    root = parser.parse(response);
   //    let allLinks = root.querySelectorAll(".entry-content div > a");

   //    for (let i = 0; i < allLinks.length; i++) {
   //       let hrefLink = allLinks[i].getAttribute("href");
   //       let countryDetailsLink = await httpRequest(hrefLink);
   //       await dataByCompany(countryDetailsLink, true);
   //    }

   //    fs.writeFileSync("dataAll.json", JSON.stringify(companyData));
   //    console.log(Object.keys(companyData).length);

   //    console.log(companyData);
}

// main();

const workbook = new Excel.Workbook();
const worksheet = workbook.addWorksheet("My Sheet");
let columnsArr = [
   { header: "Company", key: "company" },
   { header: "Resi Rank", key: "Resi Rank" },
   { header: "Overall Rank", key: "Overall Rank" },
   { header: "HQ State", key: "HQ State" },
   { header: "Primary Service", key: "Primary Service" },
   { header: "Link", key: "detailsLink" },
];

let otherKeys = ["Resi Rank", "Overall Rank", "HQ State", "Primary Service", "detailsLink"];

let years = new Set();
let allHeaders = new Set();
Object.keys(companyData).forEach((element) => {
   let insideData = companyData[element];
   Object.keys(insideData).forEach((el) => {
      //   console.log(Number.isInteger(parseInt(el)));
      if (Number.isInteger(parseInt(el))) {
         years.add(parseInt(el));
         Object.keys(insideData[el]).forEach((elem) => {
            allHeaders.add(`${el} ${elem}`);
         });
      }
   });
});

const sortedYears = Array.from(years).sort((a, b) => b - a);
const sortedAllHeaders = Array.from(allHeaders).sort().reverse();

sortedAllHeaders.forEach((header) => {
   columnsArr.push({ header: header, key: header });
});

// sortedYears.forEach((year) => {
//    //    console.log(year);
//    columnsArr.push({ header: year, key: year });
// });

worksheet.columns = columnsArr;

Object.keys(companyData).forEach((element) => {
   let insideData = companyData[element];
   let rowToAdd = { company: element };
   Object.keys(insideData).forEach((el) => {
      if (Number.isInteger(parseInt(el))) {
         Object.keys(insideData[el]).forEach((elem) => {
            rowToAdd[`${el} ${elem}`] = insideData[el][elem];
         });
      } else if (otherKeys.includes(el)) {
         rowToAdd[el] = insideData[el];
      }
   });
   worksheet.addRow(rowToAdd);
});

workbook.xlsx.writeFile("export2.xlsx");

// worksheet.addRow({
//    date: `${year}/${dateString}`,
//    ca: statesToRead[0],
//    sf: statesToRead[1],
//    TX: statesToRead[2],
//    dallas: statesToRead[3],
//    houston: statesToRead[4],
//    sa: statesToRead[5],
//    fl: statesToRead[6],
//    miami: statesToRead[7],
//    tampa: statesToRead[8],
//    orlando: statesToRead[9],
//    jacksonville: statesToRead[10],
// });
