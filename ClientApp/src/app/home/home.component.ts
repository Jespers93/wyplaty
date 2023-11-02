import { Component } from '@angular/core';
import { MainService } from '../services/main.service';

@Component({
  selector: 'app-home',
  templateUrl: './home.component.html',
})
export class HomeComponent {

  constructor(private serviceMain: MainService){}

  LoadExcelToDatabase(){
    this.serviceMain.LoadExcelToDatabase().subscribe();
  }
}
