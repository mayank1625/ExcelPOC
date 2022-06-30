import { Component } from '@angular/core';
import * as XLSX from 'xlsx';
import * as FileSaver from 'file-saver';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  title = 'excel-poc';
  formula = "";
  excelFile: any = null;
  excelBufferArr: any
  excelDataArr: any;
  transformedExcel: any[] = [];

  private formulaCompute() {
    if(!)
  }

  public concat(destinationHeader: string) {
    const paramsIdxStart = this.formula.indexOf("CONCAT(") + 7;
    const paramIdxEnd = this.formula.indexOf(")", paramsIdxStart);
    
    const params = this.formula.substring(paramsIdxStart, paramIdxEnd);
    const [p1, p2] = params.split("+");

    if(!p1 || !p2) return;

    this.excelDataArr.map((ele:any, i: number) => {
      this.transformedExcel[i] = this.transformedExcel[i] ? {...this.transformedExcel[i], [destinationHeader]: `${ele[p1]} ${ele[p2]}`} : {[destinationHeader]: `${ele[p1]} ${ele[p2]}`}
    });
  }
 
  public convert(): void {
    if(!this.formula) return
    this.concat("Name")


    const worksheet: XLSX.WorkSheet = XLSX.utils.json_to_sheet(this.transformedExcel);
    const workbook: XLSX.WorkBook = { Sheets: { 'data': worksheet }, SheetNames: ['data'] };
    const excelBuffer: any = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    const data: Blob = new Blob([excelBuffer], {type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8'});
    FileSaver.saveAs(data, 'transformed' + '_export_' + new  Date().getTime() + '.xlsx');
}

  public addExcel(e: any): void {
    this.excelFile = e.target.files[0];
    let fileReader = new FileReader();
    fileReader.readAsArrayBuffer(this.excelFile);
    fileReader.onload = e => {
      this.excelBufferArr = fileReader.result;
      console.log(this.excelBufferArr)
      const data = new Uint8Array(this.excelBufferArr);    
      const arr = new Array();    
      for(let i = 0; i != data.length; ++i) arr[i] = String.fromCharCode(data[i]);    
      const bstr = arr.join("");    
      const workbook = XLSX.read(bstr, {type:"binary"});    
      const first_sheet_name = workbook.SheetNames[0];    
      const worksheet = workbook.Sheets[first_sheet_name];    
      console.log(XLSX.utils.sheet_to_json(worksheet,{raw:true}));    
      const arraylist = XLSX.utils.sheet_to_json(worksheet,{raw:true});     
      this.excelDataArr = arraylist;    
      console.log(arraylist) 
    }
  }
}
