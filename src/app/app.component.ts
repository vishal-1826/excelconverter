import { Component } from '@angular/core';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  tableData: any[] = [];
  tableHeaders: string[] = [];
  filteredData: any[] = [];
  filterText = '';
  sortDirection: { [key: string]: boolean } = {};
  fileType: string | null = null;
  fileName: string = 'table';

  onFileDrop(event: DragEvent) {
    event.preventDefault();
    const file = event.dataTransfer?.files?.[0];
    if (file) {
      this.setFileType(file);
      this.readExcel(file);
    }
  }

  onFileSelect(event: any) {
    const file = event.target.files[0];
    if (file) {
      this.setFileType(file);
      this.readExcel(file);
    }
  }

  setFileType(file: File) {
    const ext = file.name.split('.').pop()?.toLowerCase();
    if (ext === 'xlsx' || ext === 'xls') {
      this.fileType = 'excel';
    } else if (ext === 'csv') {
      this.fileType = 'csv';
    } else if (ext === 'html' || ext === 'htm') {
      this.fileType = 'html';
    } else if (ext === 'json') {
      this.fileType = 'json';
    } else if (ext === 'pdf') {
      this.fileType = 'pdf';
    } else {
      this.fileType = null;
    }
  }

  readExcel(file: File) {
    const reader = new FileReader();
    reader.onload = (e: any) => {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const worksheet = workbook.Sheets[workbook.SheetNames[0]];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      this.tableHeaders = jsonData[0] as string[];
      this.tableData = jsonData.slice(1);
      this.applyFilter();
    };
    reader.readAsArrayBuffer(file);
  }

  applyFilter() {
    const lower = this.filterText.toLowerCase();
    this.filteredData = this.tableData.filter(row =>
      row.some((cell: any) => String(cell).toLowerCase().includes(lower))
    );
  }
 sortBy(column: string, index: number) {
    const dir = this.sortDirection[column] = !this.sortDirection[column];
    this.filteredData.sort((a, b) => {
      const valA = a[index] ?? '';
      const valB = b[index] ?? '';
      return dir ? valA.localeCompare(valB) : valB.localeCompare(valA);
    });
  }
    downloadAs(format: string) {
    const aoa = [this.tableHeaders, ...this.tableData];
    if (format === 'csv' || format === 'xlsx') {
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
      XLSX.writeFile(wb, `${this.fileName || 'table'}.${format}`);
    } else if (format === 'html') {
      let html = '<table><thead><tr>' + this.tableHeaders.map(h => `<th>${h}</th>`).join('') + '</tr></thead><tbody>';
      this.tableData.forEach(row => {
        html += '<tr>' + row.map((cell: any) => `<td>${cell}</td>`).join('') + '</tr>';
      });
      html += '</tbody></table>';
      const blob = new Blob([html], { type: 'text/html' });
      saveAs(blob, `${this.fileName || 'table'}.html`);
    } else if (format === 'json') {
      const jsonData = this.tableData.map(row =>
        Object.fromEntries(this.tableHeaders.map((key, i) => [key, row[i]]))
      );
      const blob = new Blob([JSON.stringify(jsonData, null, 2)], { type: 'application/json' });
      saveAs(blob, `${this.fileName || 'table'}.json`);
    } else if (format === 'pdf') {
      const doc = new jsPDF();
      autoTable(doc, {
        head: [this.tableHeaders],
        body: this.tableData,
        showHead: 'firstPage',
      });
      doc.save(`${this.fileName || 'table'}.pdf`);
    }
  }
}