var express = require('express');
var fs      = require('fs');
var request = require('request');
var cheerio = require('cheerio');
var app     = express();
var async = require('async');
const testFolder = './tests/';
import XLSX from 'xlsx';

app.get('/download', function(req, res) {
  fs.readdir('xlsx', (err, files) => {
    files.forEach(file => {
      readXlsx(file);
    });
  })
  return res.send(200)
})

function readXlsx(filename) {
  const workbook = XLSX.readFile('xlsx/'+filename);
  // 获取 Excel 中所有表名
  const sheetNames = workbook.SheetNames; // 返回 ['sheet1', 'sheet2']
  // 根据表名获取对应某张表
  const worksheet = workbook.Sheets[sheetNames[0]];
  const headers = {};
  const values = {};
  const data = {};
  const datas = {};
  const keys = Object.keys(worksheet);
  const cols = ["B", "K", "L", "M"];
  keys
    // 过滤以 ! 开头的 key
    .filter(k => k[0] !== '!')
    // 遍历所有单元格
    .forEach(k => {
        // 如 A11 中的 A
        let col = k.substring(0, 1);
        // 如 A11 中的 11
        let row = parseInt(k.substring(1));

        if ((cols.indexOf(col) > -1) && row > 7 && row < 25 ) {

          // 当前单元格的值
          let value = worksheet[k].v;
          data.suburb = filename.substring(0, filename.length-8).toLowerCase()
          if (col === 'B') {
            // headers[row] = value;
            data.crimeType = value;
            // return;
          }
          if (col === 'K') {
            // headers[row] = value;
            data.incident = value;
            // return;
          }
          if (col === 'L') {
            // headers[row] = value;
            data.rate = value;
            // return;
          }
          if (col === 'M') {
            // headers[row] = value;
            data.trending = value;
            // return;
          }
          console.log(JSON.stringify(data))
          // parseFloat(value)/100000).toFixed(2)
          // if (!data[headers[row]]) {
          //     data[headers[row]] = {};
          // }
          // console.log(data[headers[row]])
          // data[headers[row]][col] = value;
        }
    });
//     {
//   "albury": {
//     "Murder^": {
//       "K": 0,
//       "L": 0,
//       "M": "nc**"
//     },
//     "Assault - domestic violence related": {
//       "K": 268,
//       "L": 524.6466,
//       "M": "Stable"
//     },
//     "Assault - non-domestic violence related": {
//       "K": 330,
//       "L": 646.0201,
//       "M": "Stable"
//     },
//     "Sexual assault": {
//       "K": 50,
//       "L": 97.8818,
//       "M": "Stable"
//     },
//     "Indecent assault, act of indecency and other sexual offences": {
//       "K": 78,
//       "L": 152.6957,
//       "M": "Stable"
//     },
//     "Robbery without a weapon": {
//       "K": 8,
//       "L": 15.6611,
//       "M": "nc**"
//     },
//     "Robbery with a firearm": {
//       "K": 0,
//       "L": 0,
//       "M": "nc**"
//     },
//     "Robbery with a weapon not a firearm": {
//       "K": 5,
//       "L": 9.7882,
//       "M": "nc**"
//     },
//     "Break and enter dwelling": {
//       "K": 309,
//       "L": 604.9098,
//       "M": "Stable"
//     },
//     "Break and enter non-dwelling": {
//       "K": 133,
//       "L": 260.3657,
//       "M": "Stable"
//     },
//     "Motor vehicle theft": {
//       "K": 118,
//       "L": 231.0011,
//       "M": "Stable"
//     },
//     "Steal from motor vehicle": {
//       "K": 362,
//       "L": 708.6645,
//       "M": "Stable"
//     },
//     "Steal from retail store": {
//       "K": 246,
//       "L": 481.5786,
//       "M": "Stable"
//     },
//     "Steal from dwelling": {
//       "K": 208,
//       "L": 407.1884,
//       "M": "Stable"
//     },
//     "Steal from person": {
//       "K": 38,
//       "L": 74.3902,
//       "M": "Stable"
//     },
//     "Fraud": {
//       "K": 300,
//       "L": 587.291,
//       "M": 0.796
//     },
//     "Malicious damage to property": {
//       "K": 750,
//       "L": 1468.2276,
//       "M": "Stable"
//     }
//   }
// }
    // datas[filename.substring(0, filename.length-8).toLowerCase()] = data;
    // fs.appendFile('output.json', JSON.stringify(datas), function(err){
    //   console.log('File successfully written! - Check your project directory for the output.json file');
    // })
    // return datas;
}

app.get('/crime', function(req, res) {

    // {
    // "Murder^": {
    //   "K": 0,
    //   "L": 0,
    //   "M": "nc**"
    // },
    // "Assault - domestic violence related": {
    //   "K": 115,
    //   "L": 258.4386,
    //   "M": "Stable"
    // },
    // return res.send(data);
})

app.listen('8081')
console.log('Magic happens on port 8081');
exports = module.exports = app;
