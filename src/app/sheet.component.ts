import { Component } from '@angular/core';
import { OnInit } from '@angular/core/src/metadata';

import * as XLSX from 'xlsx';

type AOA = any[][];

@Component({
  selector: 'app-sheet',
  templateUrl: './sheet.component.html'
})
export class SheetJSComponent implements OnInit {
  extractedData: AOA = [];

  load() {
    var url =
      'https://docs.google.com/spreadsheets/d/1Wxf5xIOUmc0XiEP3zTgXPncjlpMxZY0V/edit?usp=sharing&ouid=117560585306630778365&rtpof=true&sd=true';
    var oReq = new XMLHttpRequest();
    oReq.open('GET', url, true);
    oReq.responseType = 'arraybuffer';
    var readyData = [];
    oReq.onload = function(e) {
      var arraybuffer = oReq.response;

      /* convert data to binary string */
      var data = new Uint8Array(arraybuffer);
      var arr = new Array();
      for (var i = 0; i != data.length; ++i)
        arr[i] = String.fromCharCode(data[i]);
      var bstr = arr.join('');

      /* Call XLSX */
      var workbook = XLSX.read(bstr, { type: 'binary' });

      /* DO SOMETHING WITH workbook HERE */
      var first_sheet_name = workbook.SheetNames[0];
      /* Get worksheet */
      var worksheet = workbook.Sheets[first_sheet_name];
      let x = XLSX.utils.sheet_to_json(worksheet, { raw: false });
      const dataArray = [];
      x.forEach(object => {
        delete object['__EMPTY'];
        delete object['__rowNum__'];
        if (Object.keys(object).length !== 0) {
          dataArray.push(object);
        }
      });
      const dataValues = [];
      dataArray.forEach((data, index) => {
        if (index > 0) {
          const values = [];
          for (let value in data) {
            values.push(data[value]);
          }
          readyData.push(values);
        }
      });
      console.log('dataValues : ', readyData);
    };
    this.extractedData = readyData;
    oReq.send();
  }
  ngOnInit() {
    this.load();
  }
}
