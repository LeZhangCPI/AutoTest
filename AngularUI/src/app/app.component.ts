import { Component } from '@angular/core';
import { FormControl, FormGroup, Validators } from '@angular/forms';
import { HttpClient } from '@angular/common/http';
import { MessageService } from "primeng/api";
import { ChangeDetectorRef } from '@angular/core';
@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css'],
  providers: [MessageService],
})
export class AppComponent {
  isReportSelected: boolean = false;
  title = 'Computer Packages Inc. Automation Test Main Page';
pageName = '';
  entityForm = new FormGroup({
    entityStatus: new FormControl('', Validators.required),
  });

  constructor(private messageService: MessageService, private http: HttpClient,private cd: ChangeDetectorRef) { }

  onSubmit() {
    if (this.entityForm.valid) {

      const entityStatus = this.entityForm.get('entityStatus')?.value;

      //if choose "Reports", send the request to Flask backend
      if(entityStatus === 'Reports') {
        this.showSuccess("Automation Test is Running...");
        this.cd.detectChanges();

        this.http.post('http://127.0.0.1:5000/AutoTest', {entityStatus}).subscribe({
          next: (response) => this.showSuccess("DueDateListWithExcel.py script executed successfully"),
          error: (error) => this.showError("Submission failed")
        });
      } else {
        this.showError("Selected option does not trigger any action.");
      }
    } else {
      this.showError("Please fill in all required fields.");
    }
  }

  showSuccess(message: string) {
    this.messageService.add({ severity: 'success', summary: 'Success', detail: message });
  }

  showError(message: string) {
    this.messageService.add({ severity: 'error', summary: 'Error', detail: message, sticky: true });
  }
}
