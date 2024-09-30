import { Component } from '@angular/core';
import { saveAs } from 'file-saver';
import * as XLSX from 'xlsx';
import { CommonModule } from '@angular/common';

@Component({
  selector: 'app-main',
  templateUrl: './main.component.html',
  standalone: true,
  imports: [CommonModule],
  styleUrl: './main.component.css'
})
export class MainComponent {

  data: any[][] = [];

  onFileChange(evt: any) {
    const target: DataTransfer = <DataTransfer>(evt.target);
    if (target.files.length !== 1) {
      throw new Error('Solo se permite subir un archivo a la vez');
    }

    const reader: FileReader = new FileReader();
    reader.onload = (e: any) => {
      const bstr: string = e.target.result;
      const wb: XLSX.WorkBook = XLSX.read(bstr, { type: 'binary' });

      // Asumimos que queremos la primera hoja
      const wsname: string = wb.SheetNames[2];
      const ws: XLSX.WorkSheet = wb.Sheets[wsname];

      // Convertir a formato JSON
      this.data = XLSX.utils.sheet_to_json(ws, { header: 1 });
      console.log(this.data);
    };
    reader.readAsBinaryString(target.files[0]);
  }

  updateCell(rowIndex: number, colIndex: number, event: any) {
    const newValue = event.target.innerText;
    // rowIndex +1 porque la primera fila es el encabezado
    this.data[rowIndex + 1][colIndex] = newValue;
    console.log(`Actualizado: Fila ${rowIndex + 1}, Columna ${colIndex} a ${newValue}`);
  }

  exportToExcel() {
    /* Crear una nueva hoja de trabajo */
    const ws: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet(this.data);

    /* Crear un nuevo libro de trabajo y añadir la hoja */
    const wb: XLSX.WorkBook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

    /* Generar el archivo Excel en formato binario */
    const wbout: string = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });

    /* Guardar el archivo usando FileSaver */
    const blob: Blob = new Blob([wbout], { type: 'application/octet-stream' });
    saveAs(blob, 'archivo-modificado.xlsx');
  }

}
