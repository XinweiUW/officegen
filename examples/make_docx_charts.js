var officegen = require('../lib/index.js');
var OfficeChart = require('../lib/officechart.js');
var _ = require('lodash');
var async = require ( 'async' );

var fs = require('fs');
var path = require('path');

var docx = officegen ( 'docx' );

// Remove this comment in case of debugging Officegen:
// officegen.setVerboseMode ( true );
docx.on ( 'finalize', function ( written ) {
			console.log ( 'Finish to create Word file.\nTotal bytes created: ' + written + '\n' );
		});

docx.on ( 'error', function ( err ) {
			console.log ( err );
		});

var pObj;



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

  pObj = docx.createP();
  pObj.addChart(chartInfo, callback, callback);
}

function generateCharts(callback) {
  async.each(chartsData, generateOneChart, callback);
}


function finalize() {
  var out = fs.createWriteStream('make_docx_charts.docx');

  out.on('error', function (err) {
    console.log(err);
  });

  docx.generate(out);
}

async.series([
  generateCharts    // new
], finalize);