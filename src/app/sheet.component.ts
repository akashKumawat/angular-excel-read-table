import { Component } from '@angular/core';
import { OnInit } from '@angular/core/src/metadata';

import * as XLSX from 'xlsx';

type AOA = any[][];

@Component({
  selector: 'app-sheet',
  templateUrl: './sheet.component.html'
})
export class SheetJSComponent implements OnInit {
  data: AOA = [[1, 2], [3, 4]];
  wopts: XLSX.WritingOptions = { bookType: 'xlsx', type: 'array' };
  fileName: string = 'SheetJS.xlsx';

  onFileChange(evt: any) {
    /* wire up file reader */
    // const target: DataTransfer = <DataTransfer>(evt.target);
    // if (target.files.length !== 1) throw new Error('Cannot use multiple files');
    // const reader: FileReader = new FileReader();
    // reader.onload = (e: any) => {
    //   /* read workbook */
    //   const bstr: string = e.target.result;
    //   const wb: XLSX.WorkBook = XLSX.read(bstr, { type: 'binary' });
    //   /* grab first sheet */
    //   const wsname: string = wb.SheetNames[0];
    //   const ws: XLSX.WorkSheet = wb.Sheets[wsname];
    //   /* save data */
    //   this.data = <AOA>(XLSX.utils.sheet_to_json(ws, { header: 1 }));
    //   console.log(this.data);
    // };
    // reader.readAsBinaryString(target.files[0]);
  }

  load() {
    var url =
      'https://docs.google.com/spreadsheets/d/1Wxf5xIOUmc0XiEP3zTgXPncjlpMxZY0V/edit?usp=sharing&ouid=117560585306630778365&rtpof=true&sd=true';
    var oReq = new XMLHttpRequest();
    oReq.open('GET', url, true);
    oReq.responseType = 'arraybuffer';

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
      console.log(x);
    };

    oReq.send();
  }
  ngOnInit() {
    this.load();
  }
  export(): void {
    /* generate worksheet */
    const ws: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet(this.data);

    /* generate workbook and add the worksheet */
    const wb: XLSX.WorkBook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

    /* save to file */
    XLSX.writeFile(wb, this.fileName);
  }
}
