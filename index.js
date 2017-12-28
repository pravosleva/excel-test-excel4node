"use strict";//http://jsman.ru/express/
var http = require('http'),
  fs = require('fs'),
  url = require('url'),
  port = 8080,
  host = '127.0.0.1',
  express = require('express');
var url = require('url'),
  excel4node = require('excel4node');
var app = express();

let offersFolder = './docs/output';

app.set('views', __dirname + '/views');
app.set('view engine', 'jade');
app.get('/', function(req, res){res.render('index', {title: 'Home'})});
app.get('/xlCreate', function(req, res){
  let urlParsed = url.parse(req.url, true),
    { dirName, fileName } = urlParsed.query;

  var wb = new excel4node.Workbook();
  // Add Worksheets to the workbook
  var ws = wb.addWorksheet('Sheet 1');
  // Create a reusable style
  var style = wb.createStyle({
    font: {
      color: '#FF0800',
      size: 12,
    },
    numberFormat: '$#,##0.00; ($#,##0.00); -'
  });
  // Set value of cell A1 to 100 as a number type styled with paramaters of style
  ws.cell(1,2).number(100).style(style);
  //ws.column(1).setWidth(50);

  // --- --- Border test
  ws.cell(2,3).string(`Border test`).style({
    border: {
      left: { style:`thin`, color:`#141493` },
      bottom: { style:`medium` }
    }
  });
  // --- ---

  // --- --- Img test:
  // Arguments: row, column, {offsetY, offsetX} (in pixels optional)
  //ws.Image('./docs/imgs/test.jpg', ws.Image.ONE_CELL).Position(3, 1, 10, 40).Size(255, 50);
  //var img = ws.Image('./docs/imgs/test.jpg');
  //img.Position(1,1, 0, 0);
  ws.addImage({
    path: './docs/imgs/test.jpg',
    type: 'picture',
    position: {
      type: 'absoluteAnchor',
      x: '1cm',
      y: '1in'
    }
  });
  // --- ---

  fs.mkdir(`${offersFolder}/${dirName}`,function(e){
    /*
      For local test `./docs/output/${dirName}`
      But /output should be exist
    */
    if(!e || (e && e.code === 'EEXIST')){
      //do something with contents
    } else {
      console.log(e);
    }
  });

  wb.write(`${offersFolder}/${dirName}/${fileName}`, function (err, stats) {
    if(err){
      console.error(err);
      res.writeHead(200, {"Content-Type": "text/html; charset=UTF-8"});
      res.end(`<strong>Fuck up!</strong><br /><code>${err.message}</code>`);
    }else{
      res.writeHead(200, {"Content-Type": "text/html; charset=UTF-8"});
      res.end("<strong>DONE</strong>");
    }
  }); // Writes the file ExcelFile.xlsx to the process.cwd();
});

var server = app.listen(port, host);
console.log("Express server running on http://%s:%s", host, port);
// Test:
// http://localhost:8080/xlCreate?fileName=test0001.xlsx&dirName=0000
