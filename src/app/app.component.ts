import { Component } from '@angular/core';
import { ExcelService } from './services/excel.service';
import * as Excel from "exceljs/dist/exceljs.min.js";
import * as ExcelProper from "exceljs";
import * as FileSaver from 'file-saver';

@Component({
  selector: 'my-app',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {

  workbook: ExcelProper.Workbook = new Excel.Workbook();
  blobType: string = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';

  constructor(private excelService: ExcelService) {
    let sheet = this.workbook.addWorksheet('My Sheet', {
    });

/*TITLE*/
sheet.mergeCells('C1', 'J2');
sheet.getCell('C1').value = 'Client List'

/*Column headers*/
const row = 4;
sheet.getRow(row).values = ['id', 'name', 'image'];

/*Define your column keys because this is what you use to insert your data according to your columns, they're column A, B, C, D respectively being idClient, Name, Tel, and Adresse.
So, it's pretty straight forward */
sheet.columns = [
    { key: 'id' },
    { key: 'name' },
    { key: 'image' },
];

/*Let's say you stored your data in an array called arrData. Let's say that your arrData looks like this */
let arrData = [
{ id: 1.1, name: 'child1' },
{ id: 1.2, name: 'child1' },
{ id: 1, name: 'Parent1' },
{ id: 2.1, name: 'child2' },
{ id: 2.1, name: 'child2' },
{ id: 2.1, name: 'child2' },
{ id: 2, name: 'Parent2' }
];


/* Now we use the keys we defined earlier to insert your data by iterating through arrData and calling worksheet.addRow()
*/
arrData.forEach(function(item, index) {
  sheet.addRow({
     id: item.id,
     name: item.name,
  })
  if(Number.isInteger(item.id) || item.id === 0){
    sheet.getRow(row + 1 +index).outlineLevel = 0
  }else{
    sheet.getRow(row + 1 + index).outlineLevel = 1
  }
})

  }
  exportAsXLSX(): void {
    this.workbook.xlsx.writeBuffer().then(data => {
      const blob = new Blob([data], { type: this.blobType });
      console.log(data);
      // this.excelService.exportAsExcelFile(data, 'sample');
      FileSaver.saveAs(blob, 'data');
    });

  }
}
