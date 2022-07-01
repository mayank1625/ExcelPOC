import { Component } from '@angular/core';
import * as XLSX from 'xlsx';
import * as FileSaver from 'file-saver';
import * as EXCEL from "@formulajs/formulajs";

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  title = 'excel-poc';
  formula = "CONCATENATE($FirstName$, $LastName$)";
  excelFile: any = null;
  excelBufferArr: any
  excelDataArr: any;
  transformedExcel: any[] = [];


  private formulaCompute(row: number): unknown {
    let formula = this.formula
    const formulaArr = this.formula.split("")
    let execution = [];
    
    let start = 0
    let startArg = false

    formulaArr.forEach((char, idx) => {
      if(char === "(") {
        execution.push("FUN START")
        execution.push(formula.substring(start, idx))
        start = idx + 1;
        return
      }
      if(char === '$') {
        if(!startArg) {
          start = idx;
          startArg = true;
          return
        }
        execution.push(this.excelDataArr[row][formula.substring(start+1, idx)])
        startArg = false;
        return
      }

      if(char === '"') {
        if(!startArg) {
          start = idx;
          startArg = true;
          return
        }
        execution.push(formula.substring(start+1, idx))
        startArg = false;
        return
      }

      if(char === ")") {
        execution.push("FUN END")
      }
    });


    console.log(execution)

    const execute = (idx = 0) => {
      if(execution[idx] !== "FUN START") return;
      const fun = execution[idx + 1];
      let last_idx = 0
      const args = []
      for(let i = idx + 2; i < execution.length; i++) {
        if(execution[i] !== 'FUN START' && execution[i] !== 'FUN END') {
          args.push(execution[i])
          continue
        }


        if(execution[i] === 'FUN START') {
          execute(i)
          continue
        }

        if(execution[i] === 'FUN END') {

          last_idx = i;
          break
        }
      }

      execution.splice(idx, last_idx - idx + 1, EXCEL[fun] ? EXCEL[fun](args) : '' )

    }

    execute();

    console.log(execution)
    return execution.length > 1 ? "INVALID FORMULA" : execution[0]
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
    //this.concat("Name")

    const destinationHeader= 'name'

    this.excelDataArr.map((ele:any, i: number) => {
      this.transformedExcel[i] = this.transformedExcel[i] ? {...this.transformedExcel[i], [destinationHeader]: `${this.formulaCompute(i)}`} : {[destinationHeader]: `${this.formulaCompute(i)}`}
    });

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
