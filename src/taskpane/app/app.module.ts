import { NgModule } from "@angular/core";
import { BrowserModule } from "@angular/platform-browser";
import AppComponent from "./app.component";
//import SelectionFormComponent from "./selection-form.component";
import { LocationStrategy, HashLocationStrategy } from "@angular/common";
//import { FormsModule } from '@angular/forms';
//import { HttpClientModule } from '@angular/common/http';
@NgModule({
  declarations: [AppComponent],
  imports: [BrowserModule],
  providers: [
    { provide: LocationStrategy, useClass: HashLocationStrategy },
    // Other providers suppressed
  ],
  bootstrap: [AppComponent],
})
export class AppModule { }
