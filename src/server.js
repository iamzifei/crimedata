var express = require('express');
var fs      = require('fs');
var request = require('request');
var cheerio = require('cheerio');
var app     = express();
var async = require('async');
import XLSX from 'xlsx';
const workbook = XLSX.readFile('xlsx/ashfieldlga.xlsx');
// 获取 Excel 中所有表名
const sheetNames = workbook.SheetNames; // 返回 ['sheet1', 'sheet2']
// 根据表名获取对应某张表
const worksheet = workbook.Sheets[sheetNames[0]];

function httpGet(uid, callback) {
  var prefix = 'http://intranet.ourspace.int.corp.sun/searchcentre/pages/peopleresults.aspx?k=';
  var url = prefix+uid;
  request(url, function(error, response, html){
    if(!error){
      var $ = cheerio.load(html);
      var email;
      $('#EmailField>a').filter(function(){
        var data = $(this);
        email = data.text().trim();
      })
    }

    fs.appendFile('output.txt', uid+' = '+uid+' <'+email+'>\n', function(err){
      console.log('File successfully written! - Check your project directory for the output.json file');
    })
  })
}

app.get('/crime', function(req, res) {
  const headers = {};
  const values = {};
  const data = {};
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
          if (col === 'B') {
            headers[row] = value;
            return;
          }
          if (!data[headers[row]]) {
              data[headers[row]] = {};
          }
          console.log(data[headers[row]])
          data[headers[row]][col] = value;
        }
    });
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
    return res.send(data);

})

app.get('/scrape', function(req, res){
  // Let's scrape Anchorman 2
  var uids = ['U257349','U311887','U337815','U361443','U361445','U368462',
  'u257349','u260028','u303972','u304754','u308313',
  'u311887','u312940','u314866','u316620','u316790',
  'u321498','u328565','u328707','u329262','u334913',
  'u336963','u337096','u337815','u343873','u344129',
  'u344942','u359202','u360294','u361443','u363139','u363907'];

  async.map(uids, httpGet, function (err, res){
  if (err) return console.log(err);
    console.log(res);
  });
  // for(var i = 0; i < uids.length; i ++ ){
  //
  //   // request(url, function(error, response, html){
  //   //   if(!error){
  //   //     var $ = cheerio.load(html);
  //   //     var email;
  //   //     $('#EmailField>a').filter(function(){
  //   //       var data = $(this);
  //   //       email = data.text().trim();
  //   //     })
  //   //   }
  //   //
  //   //   fs.appendFile('output.txt', uids[i]+' = '+uids[i]+' <'+email+'>\n', function(err){
  //   //     console.log('File successfully written! - Check your project directory for the output.json file');
  //   //   })
  //   //
  //   //   res.send('Check your console!')
  //   // })
  // }
})

app.listen('8081')
console.log('Magic happens on port 8081');
exports = module.exports = app;
