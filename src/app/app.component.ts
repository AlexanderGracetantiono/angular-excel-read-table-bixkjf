import { Component } from '@angular/core';

import * as XLSX from 'xlsx';
const { read, write, utils } = XLSX;

@Component({
  selector: 'my-app',
  templateUrl: './app.component.html',
  styleUrls: [ './app.component.css' ]
})
export class AppComponent  {
  name = 'Angular';
}
