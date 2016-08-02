var officegen = require('../lib/index.js');
var OfficeChart = require('../lib/officechart.js');
var _ = require('lodash');
var async = require('async');

var fs = require('fs');
var path = require('path');

var pptx = officegen('pptx');

var slide;
var pObj;

pptx.on('finalize', function (written) {
  console.log('Finish to create a PowerPoint file.\nTotal bytes created: ' + written + '\n');

  // clear the temporatory files
});

pptx.on('error', function (err) {
  console.log(err);
});

pptx.setDocTitle('Sample PPTX Document');


// this shows how one can get the base XML and modify it directly
var chart0 = new OfficeChart({
  title: 'Dynamically generated',
  renderType: 'bar',
  overlap: 50,
  gapWidth: 25,
  valAxisTitle: "Da Value Axis",
  catAxisTitle: "Da Cat Axis",
  catAxisReverseOrder: true,
  valAxisCrossAtMaxCategory: true,
  valAxisMajorGridlines: true,
  valAxisMinorGridlines: true,
  data: [
    {
      name: 'Income',
      labels: ['2005', '2006', '2007', '2008', '2009'],
      values: [23.5, 26.2, 30.1, 29.5, 24.6],
      // schemeColor: 'accent1'
      // color: 'ff0000',
      xml: {
        "c:spPr": {
          "a:solidFill": {
            "a:schemeClr": { "@val": "accent1"}
          },
          "a:ln": {
            "a:solidFill": {
              "a:schemeClr": { "@val": "tx1"}
            }
          }
        }
      }
    },
    {
      name: 'Expense',
      labels: ['2005', '2006', '2007', '2008', '2009'],
      values: [18.1, 22.8, 23.9, 25.1, 25],
      // color: '00ff00',
      // schemeColor: 'bg2'
      xml: {
        "c:spPr": {
          "a:solidFill": {
            "a:schemeClr": { "@val": "bg2"}
          },
          "a:ln": {
            "a:solidFill": {
              "a:schemeClr": { "@val": "tx1"}
            }
          }
        }
      }
    }
  ],
  fontSize: "1200", // equivalent to specifying the xml below
  xml: {
      "c:txPr": {
        "a:bodyPr": {},
        "a:listStyle": {},
        "a:p": {
          "a:pPr": {
            "a:defRPr": {
              "@sz": "1200"
            }
          },
          "a:endParaRPr": {
            "@lang": "en-US"
          }
        }
      }
    }
});

var chartsData = [
  chart0
];


function generateOneChart(chartInfo, callback) {

  slide = pptx.makeNewSlide();
  slide.name = 'OfficeChart slide';
  slide.back = 'ffffff';
  slide.addChart(chartInfo, callback, callback);
}

function generateCharts(callback) {
  async.each(chartsData, generateOneChart, callback);
}


function finalize() {
  var out = fs.createWriteStream('out_charts.pptx');

  out.on('error', function (err) {
    console.log(err);
  });

  pptx.generate(out);
}

async.series([
  generateCharts    // new
], finalize);