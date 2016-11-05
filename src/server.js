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
    // readXlsx(files[0]);
  })
  return res.send(200)
})

function readXlsx(filename) {
  console.log(filename)
  const workbook = XLSX.readFile('xlsx/'+filename);
  // 获取 Excel 中所有表名
  const sheetNames = workbook.SheetNames; // 返回 ['sheet1', 'sheet2']
  // 根据表名获取对应某张表
  const worksheet = workbook.Sheets[sheetNames[0]];
  const headers = {};
  const values = {};
  var data = {};
  var datas = [];
  var prevRow = 0;
  var prevCol = 'A';
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
        let value = worksheet[k].v;
          // console.log(row + ": " + col + " => " + value)
          data.suburb = filename.substring(0, filename.length-8).toLowerCase()

          if (col === 'B') {
            data.crimeType = value;
          }
          if (col === 'K') {
            data.incident = value;
          }
          if (col === 'L') {
            data.rate = (parseFloat(value)/1000).toFixed(2);
          }
          if (col === 'M') {
            data.trending = value;
            prevRow = row;
          }
          // console.log("data is now " + JSON.stringify(data) + ", row is " + row + ", prevrow is " + prevRow)
          if (row === prevRow) {
            // console.log('change a new row')
            fs.appendFile('output.json', JSON.stringify(data)+'\n', function(err){
              console.log('File successfully written! - Check your project directory for the output.json file');
            })
            datas.push(data);
            data = {};
          }
        }
      }
  );
    // console.log(JSON.stringify(datas))

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
