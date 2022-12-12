import { Component } from '@angular/core';
import { FormBuilder } from '@angular/forms';
import * as XLSX from 'xlsx';

type AOA = any[][];

@Component({
  selector: 'app-sheet',
  templateUrl: './sheet.component.html',
})
export class SheetJSComponent {
  data: AOA = [
    [1, 2],
    [3, 4],
  ];

  wopts: XLSX.WritingOptions = { bookType: 'xlsx', type: 'array' };
  fileName: string = 'SheetJS.xlsx';
  valueAwal: number = 0;
  valueAkhir: number = 0;
  valueSelisih: number = 0;
  indexData: number = 0;
  dataKredit: [] = [];
  dataDebit: [] = [];
  isValueNegatif = false;
  changeValueAwal(evt: any) {
    this.valueAwal = evt;
  }
  changeValueAkhir(evt: any) {
    this.valueAkhir = evt;
    this.valueSelisih = this.valueAkhir - this.valueAwal;
    if (this.valueSelisih < 0) {
      this.valueSelisih = Math.abs(this.valueSelisih);
      this.isValueNegatif = true;
    }
  }
  onFileChange(evt: any) {
    /* wire up file reader */
    const target: DataTransfer = <DataTransfer>evt.target;
    if (target.files.length !== 1) throw new Error('Cannot use multiple files');
    const reader: FileReader = new FileReader();
    reader.onload = (e: any) => {
      /* read workbook */
      const bstr: string = e.target.result;
      const wb: XLSX.WorkBook = XLSX.read(bstr, { type: 'binary' });

      /* grab first sheet */
      const wsname: string = wb.SheetNames[0];
      const ws: XLSX.WorkSheet = wb.Sheets[wsname];

      /* save data */
      this.data = <AOA>XLSX.utils.sheet_to_json(ws, { header: 1 });

      for (let ii = 0; ii < this.data.length; ii++) {
        let dataTempK = Number(this.data[ii][4]);
        let dataTempKCode = this.data[ii][2];
        if (dataTempK && dataTempK <= this.valueSelisih) {
          this.dataKredit.push([dataTempKCode, dataTempK]);
        }
        let dataTempD = Number(this.data[ii][3]);
        let dataTempDKode = this.data[ii][2];
        if (dataTempD && dataTempD <= this.valueSelisih) {
          this.dataDebit.push([dataTempDKode, dataTempD]);
        }
      }
      const dataKreditTempor = this.dataKredit;
      const dataDebitTempor = this.dataDebit;
      for (let jj = 0; jj < this.dataKredit.length; jj++) {
        for (let jjd = 0; jjd < this.dataDebit.length; jjd++) {
          if (
            this.dataKredit[jj][0] === this.dataDebit[jjd][0] &&
            this.dataKredit[jj][1] === this.dataDebit[jjd][1]
          ) {
            console.log(this.dataKredit[jj][0]);
            console.log(this.dataDebit[jjd][0]);
            dataKreditTempor.splice(jj, 1);
            dataDebitTempor.splice(jjd, 1);
          }
        }
      }
      // console.log(this.dataKredit);
      // console.log(this.dataDebit);
    };
    reader.readAsBinaryString(target.files[0]);
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
