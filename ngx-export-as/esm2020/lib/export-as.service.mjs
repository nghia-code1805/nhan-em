import { Injectable } from '@angular/core';
import { Observable } from 'rxjs';
import html2canvas from 'html2canvas';
import * as XLSX from 'xlsx';
import HTMLtoDOCX from 'html-to-docx';
import html2pdf from 'html2pdf.js';
import * as i0 from "@angular/core";
window['html2canvas'] = html2canvas;
export class ExportAsService {
    constructor() { }
    /**
     * Main base64 get method, it will return the file as base64 string
     * @param config your config
     */
    get(config) {
        // structure method name dynamically by type
        const func = 'get' + config.type.toUpperCase();
        // if type supported execute and return
        if (this[func]) {
            return this[func](config);
        }
        // throw error for unsupported formats
        return new Observable((observer) => { observer.error('Export type is not supported.'); });
    }
    /**
     * Save exported file in old javascript way
     * @param config your custom config
     * @param fileName Name of the file to be saved as
     */
    save(config, fileName) {
        // set download
        config.download = true;
        // get file name with type
        config.fileName = fileName + '.' + config.type;
        return this.get(config);
    }
    /**
     * Converts content string to blob object
     * @param content string to be converted
     */
    contentToBlob(content) {
        return new Observable((observer) => {
            // get content string and extract mime type
            const arr = content.split(','), mime = arr[0].match(/:(.*?);/)[1], bstr = atob(arr[1]);
            let n = bstr.length;
            const u8arr = new Uint8Array(n);
            while (n--) {
                u8arr[n] = bstr.charCodeAt(n);
            }
            observer.next(new Blob([u8arr], { type: mime }));
            observer.complete();
        });
    }
    /**
     * Removes base64 file type from a string like "data:text/csv;base64,"
     * @param fileContent the base64 string to remove the type from
     */
    removeFileTypeFromBase64(fileContent) {
        const re = /^data:[^]*;base64,/g;
        const newContent = re[Symbol.replace](fileContent, '');
        return newContent;
    }
    /**
     * Structure the base64 file content with the file type string
     * @param fileContent file content
     * @param fileMime file mime type "text/csv"
     */
    addFileTypeToBase64(fileContent, fileMime) {
        return `data:${fileMime};base64,${fileContent}`;
    }
    /**
     * create downloadable file from dataURL
     * @param fileName downloadable file name
     * @param dataURL file content as dataURL
     */
    downloadFromDataURL(fileName, dataURL) {
        // create blob
        this.contentToBlob(dataURL).subscribe(blob => {
            // download the blob
            this.downloadFromBlob(blob, fileName);
        });
    }
    /**
     * Downloads the blob object as a file
     * @param blob file object as blob
     * @param fileName downloadable file name
     */
    downloadFromBlob(blob, fileName) {
        // get object url
        const url = window.URL.createObjectURL(blob);
        // check for microsoft internet explorer
        if (window.navigator && window.navigator['msSaveOrOpenBlob']) {
            // use IE download or open if the user using IE
            window.navigator['msSaveOrOpenBlob'](blob, fileName);
        }
        else {
            this.saveFile(fileName, url);
        }
    }
    saveFile(fileName, url) {
        // if not using IE then create link element
        const element = document.createElement('a');
        // set download attr with file name
        element.setAttribute('download', fileName);
        // set the element as hidden
        element.style.display = 'none';
        // append the body
        document.body.appendChild(element);
        // set href attr
        element.href = url;
        // click on it to start downloading
        element.click();
        // remove the link from the dom
        document.body.removeChild(element);
    }
    getPDF(config) {
        return new Observable((observer) => {
            if (!config.options) {
                config.options = {};
            }
            config.options.filename = config.fileName;
            const element = document.getElementById(config.elementIdOrContent);
            const pdf = html2pdf().set(config.options).from(element ? element : config.elementIdOrContent);
            const download = config.download;
            const pdfCallbackFn = config.options.pdfCallbackFn;
            if (download) {
                if (pdfCallbackFn) {
                    this.applyPdfCallbackFn(pdf, pdfCallbackFn).save();
                }
                else {
                    pdf.save();
                }
                observer.next();
                observer.complete();
            }
            else {
                if (pdfCallbackFn) {
                    this.applyPdfCallbackFn(pdf, pdfCallbackFn).outputPdf('datauristring').then(data => {
                        observer.next(data);
                        observer.complete();
                    });
                }
                else {
                    pdf.outputPdf('datauristring').then(data => {
                        observer.next(data);
                        observer.complete();
                    });
                }
            }
        });
    }
    applyPdfCallbackFn(pdf, pdfCallbackFn) {
        return pdf.toPdf().get('pdf').then((pdfRef) => {
            pdfCallbackFn(pdfRef);
        });
    }
    getPNG(config) {
        return new Observable((observer) => {
            const element = document.getElementById(config.elementIdOrContent);
            html2canvas(element, config.options).then((canvas) => {
                const imgData = canvas.toDataURL('image/PNG');
                if (config.type === 'png' && config.download) {
                    this.downloadFromDataURL(config.fileName, imgData);
                    observer.next();
                }
                else {
                    observer.next(imgData);
                }
                observer.complete();
            }, err => {
                observer.error(err);
            });
        });
    }
    getCSV(config) {
        return new Observable((observer) => {
            const element = document.getElementById(config.elementIdOrContent);
            const csv = [];
            const rows = element.querySelectorAll('table tr');
            for (let index = 0; index < rows.length; index++) {
                const rowElement = rows[index];
                const row = [];
                const cols = rowElement.querySelectorAll('td, th');
                for (let colIndex = 0; colIndex < cols.length; colIndex++) {
                    const col = cols[colIndex];
                    row.push('"' + col.innerText + '"');
                }
                csv.push(row.join(','));
            }
            const csvContent = 'data:text/csv;base64,' + this.btoa(csv.join('\n'));
            if (config.download) {
                this.downloadFromDataURL(config.fileName, csvContent);
                observer.next();
            }
            else {
                observer.next(csvContent);
            }
            observer.complete();
        });
    }
    getTXT(config) {
        const nameFrags = config.fileName.split('.');
        config.fileName = `${nameFrags[0]}.txt`;
        return this.getCSV(config);
    }
    getXLS(config) {
        return new Observable((observer) => {
            const element = document.getElementById(config.elementIdOrContent);
            const ws3 = XLSX.utils.table_to_sheet(element, config.options);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws3, config.fileName);
            const out = XLSX.write(wb, { type: 'base64' });
            const xlsContent = 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,' + out;
            if (config.download) {
                this.downloadFromDataURL(config.fileName, xlsContent);
                observer.next();
            }
            else {
                observer.next(xlsContent);
            }
            observer.complete();
        });
    }
    getXLSX(config) {
        return this.getXLS(config);
    }
    getDOCX(config) {
        return new Observable((observer) => {
            const contentDocument = document.getElementById(config.elementIdOrContent).outerHTML;
            const content = '<!DOCTYPE html>' + contentDocument;
            HTMLtoDOCX(content, null, config.options).then(converted => {
                if (config.download) {
                    this.downloadFromBlob(converted, config.fileName);
                    observer.next();
                    observer.complete();
                }
                else {
                    const reader = new FileReader();
                    reader.onloadend = () => {
                        const base64data = reader.result;
                        observer.next(base64data);
                        observer.complete();
                    };
                    reader.readAsDataURL(converted);
                }
            });
        });
    }
    getDOC(config) {
        return this.getDOCX(config);
    }
    getJSON(config) {
        return new Observable((observer) => {
            const data = []; // first row needs to be headers
            const headers = [];
            const table = document.getElementById(config.elementIdOrContent);
            for (let index = 0; index < table.rows[0].cells.length; index++) {
                headers[index] = table.rows[0].cells[index].innerHTML.toLowerCase().replace(/ /gi, '');
            }
            // go through cells
            for (let i = 1; i < table.rows.length; i++) {
                const tableRow = table.rows[i];
                const rowData = {};
                for (let j = 0; j < tableRow.cells.length; j++) {
                    rowData[headers[j]] = tableRow.cells[j].innerHTML;
                }
                data.push(rowData);
            }
            const jsonString = JSON.stringify(data);
            const jsonBase64 = this.btoa(jsonString);
            const dataStr = 'data:text/json;base64,' + jsonBase64;
            if (config.download) {
                this.downloadFromDataURL(config.fileName, dataStr);
                observer.next();
            }
            else {
                observer.next(data);
            }
            observer.complete();
        });
    }
    getXML(config) {
        return new Observable((observer) => {
            let xml = '<?xml version="1.0" encoding="UTF-8"?><Root><Classes>';
            const tritem = document.getElementById(config.elementIdOrContent).getElementsByTagName('tr');
            for (let i = 0; i < tritem.length; i++) {
                const celldata = tritem[i];
                if (celldata.cells.length > 0) {
                    xml += '<Class name="' + celldata.cells[0].textContent + '">\n';
                    for (let m = 1; m < celldata.cells.length; ++m) {
                        xml += '\t<data>' + celldata.cells[m].textContent + '</data>\n';
                    }
                    xml += '</Class>\n';
                }
            }
            xml += '</Classes></Root>';
            const base64 = 'data:text/xml;base64,' + this.btoa(xml);
            if (config.download) {
                this.downloadFromDataURL(config.fileName, base64);
                observer.next();
            }
            else {
                observer.next(base64);
            }
            observer.complete();
        });
    }
    btoa(content) {
        return btoa(unescape(encodeURIComponent(content)));
    }
}
ExportAsService.ɵfac = i0.ɵɵngDeclareFactory({ minVersion: "12.0.0", version: "13.0.2", ngImport: i0, type: ExportAsService, deps: [], target: i0.ɵɵFactoryTarget.Injectable });
ExportAsService.ɵprov = i0.ɵɵngDeclareInjectable({ minVersion: "12.0.0", version: "13.0.2", ngImport: i0, type: ExportAsService });
i0.ɵɵngDeclareClassMetadata({ minVersion: "12.0.0", version: "13.0.2", ngImport: i0, type: ExportAsService, decorators: [{
            type: Injectable
        }], ctorParameters: function () { return []; } });
//# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJmaWxlIjoiZXhwb3J0LWFzLnNlcnZpY2UuanMiLCJzb3VyY2VSb290IjoiIiwic291cmNlcyI6WyIuLi8uLi8uLi8uLi9wcm9qZWN0cy9uZ3gtZXhwb3J0LWFzL3NyYy9saWIvZXhwb3J0LWFzLnNlcnZpY2UudHMiXSwibmFtZXMiOltdLCJtYXBwaW5ncyI6IkFBQUEsT0FBTyxFQUFFLFVBQVUsRUFBRSxNQUFNLGVBQWUsQ0FBQztBQUMzQyxPQUFPLEVBQUUsVUFBVSxFQUFFLE1BQU0sTUFBTSxDQUFDO0FBSWxDLE9BQU8sV0FBVyxNQUFNLGFBQWEsQ0FBQztBQUN0QyxPQUFPLEtBQUssSUFBSSxNQUFNLE1BQU0sQ0FBQztBQUM3QixPQUFPLFVBQVUsTUFBTSxjQUFjLENBQUM7QUFDdEMsT0FBTyxRQUFRLE1BQU0sYUFBYSxDQUFDOztBQUNuQyxNQUFNLENBQUMsYUFBYSxDQUFDLEdBQUcsV0FBVyxDQUFDO0FBR3BDLE1BQU0sT0FBTyxlQUFlO0lBRTFCLGdCQUFnQixDQUFDO0lBRWpCOzs7T0FHRztJQUNILEdBQUcsQ0FBQyxNQUFzQjtRQUN4Qiw0Q0FBNEM7UUFDNUMsTUFBTSxJQUFJLEdBQUcsS0FBSyxHQUFHLE1BQU0sQ0FBQyxJQUFJLENBQUMsV0FBVyxFQUFFLENBQUM7UUFDL0MsdUNBQXVDO1FBQ3ZDLElBQUksSUFBSSxDQUFDLElBQUksQ0FBQyxFQUFFO1lBQ2QsT0FBTyxJQUFJLENBQUMsSUFBSSxDQUFDLENBQUMsTUFBTSxDQUFDLENBQUM7U0FDM0I7UUFFRCxzQ0FBc0M7UUFDdEMsT0FBTyxJQUFJLFVBQVUsQ0FBQyxDQUFDLFFBQVEsRUFBRSxFQUFFLEdBQUcsUUFBUSxDQUFDLEtBQUssQ0FBQywrQkFBK0IsQ0FBQyxDQUFDLENBQUMsQ0FBQyxDQUFDLENBQUM7SUFDNUYsQ0FBQztJQUVEOzs7O09BSUc7SUFDSCxJQUFJLENBQUMsTUFBc0IsRUFBRSxRQUFnQjtRQUMzQyxlQUFlO1FBQ2YsTUFBTSxDQUFDLFFBQVEsR0FBRyxJQUFJLENBQUM7UUFDdkIsMEJBQTBCO1FBQzFCLE1BQU0sQ0FBQyxRQUFRLEdBQUcsUUFBUSxHQUFHLEdBQUcsR0FBRyxNQUFNLENBQUMsSUFBSSxDQUFDO1FBQy9DLE9BQU8sSUFBSSxDQUFDLEdBQUcsQ0FBQyxNQUFNLENBQUMsQ0FBQztJQUMxQixDQUFDO0lBRUQ7OztPQUdHO0lBQ0gsYUFBYSxDQUFDLE9BQWU7UUFDM0IsT0FBTyxJQUFJLFVBQVUsQ0FBQyxDQUFDLFFBQVEsRUFBRSxFQUFFO1lBQ2pDLDJDQUEyQztZQUMzQyxNQUFNLEdBQUcsR0FBRyxPQUFPLENBQUMsS0FBSyxDQUFDLEdBQUcsQ0FBQyxFQUFFLElBQUksR0FBRyxHQUFHLENBQUMsQ0FBQyxDQUFDLENBQUMsS0FBSyxDQUFDLFNBQVMsQ0FBQyxDQUFDLENBQUMsQ0FBQyxFQUMvRCxJQUFJLEdBQUcsSUFBSSxDQUFDLEdBQUcsQ0FBQyxDQUFDLENBQUMsQ0FBQyxDQUFDO1lBQ3RCLElBQUksQ0FBQyxHQUFHLElBQUksQ0FBQyxNQUFNLENBQUM7WUFDcEIsTUFBTSxLQUFLLEdBQUcsSUFBSSxVQUFVLENBQUMsQ0FBQyxDQUFDLENBQUM7WUFDaEMsT0FBTyxDQUFDLEVBQUUsRUFBRTtnQkFDVixLQUFLLENBQUMsQ0FBQyxDQUFDLEdBQUcsSUFBSSxDQUFDLFVBQVUsQ0FBQyxDQUFDLENBQUMsQ0FBQzthQUMvQjtZQUNELFFBQVEsQ0FBQyxJQUFJLENBQUMsSUFBSSxJQUFJLENBQUMsQ0FBQyxLQUFLLENBQUMsRUFBRSxFQUFFLElBQUksRUFBRSxJQUFJLEVBQUUsQ0FBQyxDQUFDLENBQUM7WUFDakQsUUFBUSxDQUFDLFFBQVEsRUFBRSxDQUFDO1FBQ3RCLENBQUMsQ0FBQyxDQUFDO0lBQ0wsQ0FBQztJQUVEOzs7T0FHRztJQUNILHdCQUF3QixDQUFDLFdBQW1CO1FBQzFDLE1BQU0sRUFBRSxHQUFHLHFCQUFxQixDQUFDO1FBQ2pDLE1BQU0sVUFBVSxHQUFXLEVBQUUsQ0FBQyxNQUFNLENBQUMsT0FBTyxDQUFDLENBQUMsV0FBVyxFQUFFLEVBQUUsQ0FBQyxDQUFDO1FBQy9ELE9BQU8sVUFBVSxDQUFDO0lBQ3BCLENBQUM7SUFFRDs7OztPQUlHO0lBQ0gsbUJBQW1CLENBQUMsV0FBbUIsRUFBRSxRQUFnQjtRQUN2RCxPQUFPLFFBQVEsUUFBUSxXQUFXLFdBQVcsRUFBRSxDQUFDO0lBQ2xELENBQUM7SUFFRDs7OztPQUlHO0lBQ0gsbUJBQW1CLENBQUMsUUFBZ0IsRUFBRSxPQUFlO1FBQ25ELGNBQWM7UUFDZCxJQUFJLENBQUMsYUFBYSxDQUFDLE9BQU8sQ0FBQyxDQUFDLFNBQVMsQ0FBQyxJQUFJLENBQUMsRUFBRTtZQUMzQyxvQkFBb0I7WUFDcEIsSUFBSSxDQUFDLGdCQUFnQixDQUFDLElBQUksRUFBRSxRQUFRLENBQUMsQ0FBQztRQUN4QyxDQUFDLENBQUMsQ0FBQztJQUNMLENBQUM7SUFFRDs7OztPQUlHO0lBQ0gsZ0JBQWdCLENBQUMsSUFBVSxFQUFFLFFBQWdCO1FBQzNDLGlCQUFpQjtRQUNqQixNQUFNLEdBQUcsR0FBRyxNQUFNLENBQUMsR0FBRyxDQUFDLGVBQWUsQ0FBQyxJQUFJLENBQUMsQ0FBQztRQUM3Qyx3Q0FBd0M7UUFDeEMsSUFBSSxNQUFNLENBQUMsU0FBUyxJQUFJLE1BQU0sQ0FBQyxTQUFTLENBQUMsa0JBQWtCLENBQUMsRUFBRTtZQUM1RCwrQ0FBK0M7WUFDL0MsTUFBTSxDQUFDLFNBQVMsQ0FBQyxrQkFBa0IsQ0FBQyxDQUFDLElBQUksRUFBRSxRQUFRLENBQUMsQ0FBQztTQUN0RDthQUFNO1lBQ0wsSUFBSSxDQUFDLFFBQVEsQ0FBQyxRQUFRLEVBQUUsR0FBRyxDQUFDLENBQUM7U0FDOUI7SUFDSCxDQUFDO0lBRU8sUUFBUSxDQUFDLFFBQWdCLEVBQUUsR0FBVztRQUM1QywyQ0FBMkM7UUFDM0MsTUFBTSxPQUFPLEdBQUcsUUFBUSxDQUFDLGFBQWEsQ0FBQyxHQUFHLENBQUMsQ0FBQztRQUM1QyxtQ0FBbUM7UUFDbkMsT0FBTyxDQUFDLFlBQVksQ0FBQyxVQUFVLEVBQUUsUUFBUSxDQUFDLENBQUM7UUFDM0MsNEJBQTRCO1FBQzVCLE9BQU8sQ0FBQyxLQUFLLENBQUMsT0FBTyxHQUFHLE1BQU0sQ0FBQztRQUMvQixrQkFBa0I7UUFDbEIsUUFBUSxDQUFDLElBQUksQ0FBQyxXQUFXLENBQUMsT0FBTyxDQUFDLENBQUM7UUFDbkMsZ0JBQWdCO1FBQ2hCLE9BQU8sQ0FBQyxJQUFJLEdBQUcsR0FBRyxDQUFDO1FBQ25CLG1DQUFtQztRQUNuQyxPQUFPLENBQUMsS0FBSyxFQUFFLENBQUM7UUFDaEIsK0JBQStCO1FBQy9CLFFBQVEsQ0FBQyxJQUFJLENBQUMsV0FBVyxDQUFDLE9BQU8sQ0FBQyxDQUFDO0lBQ3JDLENBQUM7SUFFTyxNQUFNLENBQUMsTUFBc0I7UUFDbkMsT0FBTyxJQUFJLFVBQVUsQ0FBQyxDQUFDLFFBQVEsRUFBRSxFQUFFO1lBQ2pDLElBQUksQ0FBQyxNQUFNLENBQUMsT0FBTyxFQUFFO2dCQUNuQixNQUFNLENBQUMsT0FBTyxHQUFHLEVBQUUsQ0FBQzthQUNyQjtZQUNELE1BQU0sQ0FBQyxPQUFPLENBQUMsUUFBUSxHQUFHLE1BQU0sQ0FBQyxRQUFRLENBQUM7WUFDMUMsTUFBTSxPQUFPLEdBQWdCLFFBQVEsQ0FBQyxjQUFjLENBQUMsTUFBTSxDQUFDLGtCQUFrQixDQUFDLENBQUM7WUFDaEYsTUFBTSxHQUFHLEdBQUcsUUFBUSxFQUFFLENBQUMsR0FBRyxDQUFDLE1BQU0sQ0FBQyxPQUFPLENBQUMsQ0FBQyxJQUFJLENBQUMsT0FBTyxDQUFDLENBQUMsQ0FBQyxPQUFPLENBQUMsQ0FBQyxDQUFDLE1BQU0sQ0FBQyxrQkFBa0IsQ0FBQyxDQUFDO1lBRS9GLE1BQU0sUUFBUSxHQUFHLE1BQU0sQ0FBQyxRQUFRLENBQUM7WUFDakMsTUFBTSxhQUFhLEdBQUcsTUFBTSxDQUFDLE9BQU8sQ0FBQyxhQUFhLENBQUM7WUFDbkQsSUFBSSxRQUFRLEVBQUU7Z0JBQ1osSUFBSSxhQUFhLEVBQUU7b0JBQ2pCLElBQUksQ0FBQyxrQkFBa0IsQ0FBQyxHQUFHLEVBQUUsYUFBYSxDQUFDLENBQUMsSUFBSSxFQUFFLENBQUM7aUJBQ3BEO3FCQUFNO29CQUNMLEdBQUcsQ0FBQyxJQUFJLEVBQUUsQ0FBQztpQkFDWjtnQkFDRCxRQUFRLENBQUMsSUFBSSxFQUFFLENBQUM7Z0JBQ2hCLFFBQVEsQ0FBQyxRQUFRLEVBQUUsQ0FBQzthQUNyQjtpQkFBTTtnQkFDTCxJQUFJLGFBQWEsRUFBRTtvQkFDakIsSUFBSSxDQUFDLGtCQUFrQixDQUFDLEdBQUcsRUFBRSxhQUFhLENBQUMsQ0FBQyxTQUFTLENBQUMsZUFBZSxDQUFDLENBQUMsSUFBSSxDQUFDLElBQUksQ0FBQyxFQUFFO3dCQUNqRixRQUFRLENBQUMsSUFBSSxDQUFDLElBQUksQ0FBQyxDQUFDO3dCQUNwQixRQUFRLENBQUMsUUFBUSxFQUFFLENBQUM7b0JBQ3RCLENBQUMsQ0FBQyxDQUFDO2lCQUNKO3FCQUFNO29CQUNMLEdBQUcsQ0FBQyxTQUFTLENBQUMsZUFBZSxDQUFDLENBQUMsSUFBSSxDQUFDLElBQUksQ0FBQyxFQUFFO3dCQUN6QyxRQUFRLENBQUMsSUFBSSxDQUFDLElBQUksQ0FBQyxDQUFDO3dCQUNwQixRQUFRLENBQUMsUUFBUSxFQUFFLENBQUM7b0JBQ3RCLENBQUMsQ0FBQyxDQUFDO2lCQUNKO2FBQ0Y7UUFDSCxDQUFDLENBQUMsQ0FBQztJQUNMLENBQUM7SUFFTyxrQkFBa0IsQ0FBQyxHQUFHLEVBQUUsYUFBYTtRQUMzQyxPQUFPLEdBQUcsQ0FBQyxLQUFLLEVBQUUsQ0FBQyxHQUFHLENBQUMsS0FBSyxDQUFDLENBQUMsSUFBSSxDQUFDLENBQUMsTUFBTSxFQUFFLEVBQUU7WUFDNUMsYUFBYSxDQUFDLE1BQU0sQ0FBQyxDQUFDO1FBQ3hCLENBQUMsQ0FBQyxDQUFDO0lBQ0wsQ0FBQztJQUVPLE1BQU0sQ0FBQyxNQUFzQjtRQUNuQyxPQUFPLElBQUksVUFBVSxDQUFDLENBQUMsUUFBUSxFQUFFLEVBQUU7WUFDakMsTUFBTSxPQUFPLEdBQWdCLFFBQVEsQ0FBQyxjQUFjLENBQUMsTUFBTSxDQUFDLGtCQUFrQixDQUFDLENBQUM7WUFDaEYsV0FBVyxDQUFDLE9BQU8sRUFBRSxNQUFNLENBQUMsT0FBTyxDQUFDLENBQUMsSUFBSSxDQUFDLENBQUMsTUFBTSxFQUFFLEVBQUU7Z0JBQ25ELE1BQU0sT0FBTyxHQUFHLE1BQU0sQ0FBQyxTQUFTLENBQUMsV0FBVyxDQUFDLENBQUM7Z0JBQzlDLElBQUksTUFBTSxDQUFDLElBQUksS0FBSyxLQUFLLElBQUksTUFBTSxDQUFDLFFBQVEsRUFBRTtvQkFDNUMsSUFBSSxDQUFDLG1CQUFtQixDQUFDLE1BQU0sQ0FBQyxRQUFRLEVBQUUsT0FBTyxDQUFDLENBQUM7b0JBQ25ELFFBQVEsQ0FBQyxJQUFJLEVBQUUsQ0FBQztpQkFDakI7cUJBQU07b0JBQ0wsUUFBUSxDQUFDLElBQUksQ0FBQyxPQUFPLENBQUMsQ0FBQztpQkFDeEI7Z0JBQ0QsUUFBUSxDQUFDLFFBQVEsRUFBRSxDQUFDO1lBQ3RCLENBQUMsRUFBRSxHQUFHLENBQUMsRUFBRTtnQkFDUCxRQUFRLENBQUMsS0FBSyxDQUFDLEdBQUcsQ0FBQyxDQUFDO1lBQ3RCLENBQUMsQ0FBQyxDQUFDO1FBQ0wsQ0FBQyxDQUFDLENBQUM7SUFDTCxDQUFDO0lBRU8sTUFBTSxDQUFDLE1BQXNCO1FBQ25DLE9BQU8sSUFBSSxVQUFVLENBQUMsQ0FBQyxRQUFRLEVBQUUsRUFBRTtZQUNqQyxNQUFNLE9BQU8sR0FBZ0IsUUFBUSxDQUFDLGNBQWMsQ0FBQyxNQUFNLENBQUMsa0JBQWtCLENBQUMsQ0FBQztZQUNoRixNQUFNLEdBQUcsR0FBRyxFQUFFLENBQUM7WUFDZixNQUFNLElBQUksR0FBUSxPQUFPLENBQUMsZ0JBQWdCLENBQUMsVUFBVSxDQUFDLENBQUM7WUFDdkQsS0FBSyxJQUFJLEtBQUssR0FBRyxDQUFDLEVBQUUsS0FBSyxHQUFHLElBQUksQ0FBQyxNQUFNLEVBQUUsS0FBSyxFQUFFLEVBQUU7Z0JBQ2hELE1BQU0sVUFBVSxHQUFHLElBQUksQ0FBQyxLQUFLLENBQUMsQ0FBQztnQkFDL0IsTUFBTSxHQUFHLEdBQUcsRUFBRSxDQUFDO2dCQUNmLE1BQU0sSUFBSSxHQUFHLFVBQVUsQ0FBQyxnQkFBZ0IsQ0FBQyxRQUFRLENBQUMsQ0FBQztnQkFDbkQsS0FBSyxJQUFJLFFBQVEsR0FBRyxDQUFDLEVBQUUsUUFBUSxHQUFHLElBQUksQ0FBQyxNQUFNLEVBQUUsUUFBUSxFQUFFLEVBQUU7b0JBQ3pELE1BQU0sR0FBRyxHQUFHLElBQUksQ0FBQyxRQUFRLENBQUMsQ0FBQztvQkFDM0IsR0FBRyxDQUFDLElBQUksQ0FBQyxHQUFHLEdBQUMsR0FBRyxDQUFDLFNBQVMsR0FBQyxHQUFHLENBQUMsQ0FBQztpQkFDakM7Z0JBQ0QsR0FBRyxDQUFDLElBQUksQ0FBQyxHQUFHLENBQUMsSUFBSSxDQUFDLEdBQUcsQ0FBQyxDQUFDLENBQUM7YUFDekI7WUFDRCxNQUFNLFVBQVUsR0FBRyx1QkFBdUIsR0FBRyxJQUFJLENBQUMsSUFBSSxDQUFDLEdBQUcsQ0FBQyxJQUFJLENBQUMsSUFBSSxDQUFDLENBQUMsQ0FBQztZQUN2RSxJQUFJLE1BQU0sQ0FBQyxRQUFRLEVBQUU7Z0JBQ25CLElBQUksQ0FBQyxtQkFBbUIsQ0FBQyxNQUFNLENBQUMsUUFBUSxFQUFFLFVBQVUsQ0FBQyxDQUFDO2dCQUN0RCxRQUFRLENBQUMsSUFBSSxFQUFFLENBQUM7YUFDakI7aUJBQU07Z0JBQ0wsUUFBUSxDQUFDLElBQUksQ0FBQyxVQUFVLENBQUMsQ0FBQzthQUMzQjtZQUNELFFBQVEsQ0FBQyxRQUFRLEVBQUUsQ0FBQztRQUN0QixDQUFDLENBQUMsQ0FBQztJQUNMLENBQUM7SUFFTyxNQUFNLENBQUMsTUFBc0I7UUFDbkMsTUFBTSxTQUFTLEdBQUcsTUFBTSxDQUFDLFFBQVEsQ0FBQyxLQUFLLENBQUMsR0FBRyxDQUFDLENBQUM7UUFDN0MsTUFBTSxDQUFDLFFBQVEsR0FBRyxHQUFHLFNBQVMsQ0FBQyxDQUFDLENBQUMsTUFBTSxDQUFDO1FBQ3hDLE9BQU8sSUFBSSxDQUFDLE1BQU0sQ0FBQyxNQUFNLENBQUMsQ0FBQztJQUM3QixDQUFDO0lBRU8sTUFBTSxDQUFDLE1BQXNCO1FBQ25DLE9BQU8sSUFBSSxVQUFVLENBQUMsQ0FBQyxRQUFRLEVBQUUsRUFBRTtZQUVqQyxNQUFNLE9BQU8sR0FBZ0IsUUFBUSxDQUFDLGNBQWMsQ0FBQyxNQUFNLENBQUMsa0JBQWtCLENBQUMsQ0FBQztZQUNoRixNQUFNLEdBQUcsR0FBRyxJQUFJLENBQUMsS0FBSyxDQUFDLGNBQWMsQ0FBQyxPQUFPLEVBQUUsTUFBTSxDQUFDLE9BQU8sQ0FBQyxDQUFDO1lBQy9ELE1BQU0sRUFBRSxHQUFHLElBQUksQ0FBQyxLQUFLLENBQUMsUUFBUSxFQUFFLENBQUM7WUFDakMsSUFBSSxDQUFDLEtBQUssQ0FBQyxpQkFBaUIsQ0FBQyxFQUFFLEVBQUUsR0FBRyxFQUFFLE1BQU0sQ0FBQyxRQUFRLENBQUMsQ0FBQztZQUN2RCxNQUFNLEdBQUcsR0FBRyxJQUFJLENBQUMsS0FBSyxDQUFDLEVBQUUsRUFBRSxFQUFFLElBQUksRUFBRSxRQUFRLEVBQUUsQ0FBQyxDQUFDO1lBQy9DLE1BQU0sVUFBVSxHQUFHLGdGQUFnRixHQUFHLEdBQUcsQ0FBQztZQUMxRyxJQUFJLE1BQU0sQ0FBQyxRQUFRLEVBQUU7Z0JBQ25CLElBQUksQ0FBQyxtQkFBbUIsQ0FBQyxNQUFNLENBQUMsUUFBUSxFQUFFLFVBQVUsQ0FBQyxDQUFDO2dCQUN0RCxRQUFRLENBQUMsSUFBSSxFQUFFLENBQUM7YUFDakI7aUJBQU07Z0JBQ0wsUUFBUSxDQUFDLElBQUksQ0FBQyxVQUFVLENBQUMsQ0FBQzthQUMzQjtZQUNELFFBQVEsQ0FBQyxRQUFRLEVBQUUsQ0FBQztRQUN0QixDQUFDLENBQUMsQ0FBQztJQUNMLENBQUM7SUFFTyxPQUFPLENBQUMsTUFBc0I7UUFDcEMsT0FBTyxJQUFJLENBQUMsTUFBTSxDQUFDLE1BQU0sQ0FBQyxDQUFDO0lBQzdCLENBQUM7SUFFTyxPQUFPLENBQUMsTUFBc0I7UUFDcEMsT0FBTyxJQUFJLFVBQVUsQ0FBQyxDQUFDLFFBQVEsRUFBRSxFQUFFO1lBQ2pDLE1BQU0sZUFBZSxHQUFXLFFBQVEsQ0FBQyxjQUFjLENBQUMsTUFBTSxDQUFDLGtCQUFrQixDQUFDLENBQUMsU0FBUyxDQUFDO1lBQzdGLE1BQU0sT0FBTyxHQUFHLGlCQUFpQixHQUFHLGVBQWUsQ0FBQztZQUNwRCxVQUFVLENBQUMsT0FBTyxFQUFFLElBQUksRUFBRSxNQUFNLENBQUMsT0FBTyxDQUFDLENBQUMsSUFBSSxDQUFDLFNBQVMsQ0FBQyxFQUFFO2dCQUN6RCxJQUFJLE1BQU0sQ0FBQyxRQUFRLEVBQUU7b0JBQ25CLElBQUksQ0FBQyxnQkFBZ0IsQ0FBQyxTQUFTLEVBQUUsTUFBTSxDQUFDLFFBQVEsQ0FBQyxDQUFDO29CQUNsRCxRQUFRLENBQUMsSUFBSSxFQUFFLENBQUM7b0JBQ2hCLFFBQVEsQ0FBQyxRQUFRLEVBQUUsQ0FBQztpQkFDckI7cUJBQU07b0JBQ0wsTUFBTSxNQUFNLEdBQUcsSUFBSSxVQUFVLEVBQUUsQ0FBQztvQkFDaEMsTUFBTSxDQUFDLFNBQVMsR0FBRyxHQUFHLEVBQUU7d0JBQ3RCLE1BQU0sVUFBVSxHQUFHLE1BQU0sQ0FBQyxNQUFnQixDQUFDO3dCQUMzQyxRQUFRLENBQUMsSUFBSSxDQUFDLFVBQVUsQ0FBQyxDQUFDO3dCQUMxQixRQUFRLENBQUMsUUFBUSxFQUFFLENBQUM7b0JBQ3RCLENBQUMsQ0FBQztvQkFDRixNQUFNLENBQUMsYUFBYSxDQUFDLFNBQVMsQ0FBQyxDQUFDO2lCQUNqQztZQUNILENBQUMsQ0FBQyxDQUFDO1FBQ0wsQ0FBQyxDQUFDLENBQUM7SUFDTCxDQUFDO0lBRU8sTUFBTSxDQUFDLE1BQXNCO1FBQ25DLE9BQU8sSUFBSSxDQUFDLE9BQU8sQ0FBQyxNQUFNLENBQUMsQ0FBQztJQUM5QixDQUFDO0lBRU8sT0FBTyxDQUFDLE1BQXNCO1FBQ3BDLE9BQU8sSUFBSSxVQUFVLENBQUMsQ0FBQyxRQUFRLEVBQUUsRUFBRTtZQUNqQyxNQUFNLElBQUksR0FBRyxFQUFFLENBQUMsQ0FBQyxnQ0FBZ0M7WUFDakQsTUFBTSxPQUFPLEdBQUcsRUFBRSxDQUFDO1lBQ25CLE1BQU0sS0FBSyxHQUFxQixRQUFRLENBQUMsY0FBYyxDQUFDLE1BQU0sQ0FBQyxrQkFBa0IsQ0FBQyxDQUFDO1lBQ25GLEtBQUssSUFBSSxLQUFLLEdBQUcsQ0FBQyxFQUFFLEtBQUssR0FBRyxLQUFLLENBQUMsSUFBSSxDQUFDLENBQUMsQ0FBQyxDQUFDLEtBQUssQ0FBQyxNQUFNLEVBQUUsS0FBSyxFQUFFLEVBQUU7Z0JBQy9ELE9BQU8sQ0FBQyxLQUFLLENBQUMsR0FBRyxLQUFLLENBQUMsSUFBSSxDQUFDLENBQUMsQ0FBQyxDQUFDLEtBQUssQ0FBQyxLQUFLLENBQUMsQ0FBQyxTQUFTLENBQUMsV0FBVyxFQUFFLENBQUMsT0FBTyxDQUFDLEtBQUssRUFBRSxFQUFFLENBQUMsQ0FBQzthQUN4RjtZQUNELG1CQUFtQjtZQUNuQixLQUFLLElBQUksQ0FBQyxHQUFHLENBQUMsRUFBRSxDQUFDLEdBQUcsS0FBSyxDQUFDLElBQUksQ0FBQyxNQUFNLEVBQUUsQ0FBQyxFQUFFLEVBQUU7Z0JBQzFDLE1BQU0sUUFBUSxHQUFHLEtBQUssQ0FBQyxJQUFJLENBQUMsQ0FBQyxDQUFDLENBQUM7Z0JBQUMsTUFBTSxPQUFPLEdBQUcsRUFBRSxDQUFDO2dCQUNuRCxLQUFLLElBQUksQ0FBQyxHQUFHLENBQUMsRUFBRSxDQUFDLEdBQUcsUUFBUSxDQUFDLEtBQUssQ0FBQyxNQUFNLEVBQUUsQ0FBQyxFQUFFLEVBQUU7b0JBQzlDLE9BQU8sQ0FBQyxPQUFPLENBQUMsQ0FBQyxDQUFDLENBQUMsR0FBRyxRQUFRLENBQUMsS0FBSyxDQUFDLENBQUMsQ0FBQyxDQUFDLFNBQVMsQ0FBQztpQkFDbkQ7Z0JBQ0QsSUFBSSxDQUFDLElBQUksQ0FBQyxPQUFPLENBQUMsQ0FBQzthQUNwQjtZQUNELE1BQU0sVUFBVSxHQUFHLElBQUksQ0FBQyxTQUFTLENBQUMsSUFBSSxDQUFDLENBQUM7WUFDeEMsTUFBTSxVQUFVLEdBQUcsSUFBSSxDQUFDLElBQUksQ0FBQyxVQUFVLENBQUMsQ0FBQztZQUN6QyxNQUFNLE9BQU8sR0FBRyx3QkFBd0IsR0FBRyxVQUFVLENBQUM7WUFDdEQsSUFBSSxNQUFNLENBQUMsUUFBUSxFQUFFO2dCQUNuQixJQUFJLENBQUMsbUJBQW1CLENBQUMsTUFBTSxDQUFDLFFBQVEsRUFBRSxPQUFPLENBQUMsQ0FBQztnQkFDbkQsUUFBUSxDQUFDLElBQUksRUFBRSxDQUFDO2FBQ2pCO2lCQUFNO2dCQUNMLFFBQVEsQ0FBQyxJQUFJLENBQUMsSUFBSSxDQUFDLENBQUM7YUFDckI7WUFDRCxRQUFRLENBQUMsUUFBUSxFQUFFLENBQUM7UUFDdEIsQ0FBQyxDQUFDLENBQUM7SUFDTCxDQUFDO0lBRU8sTUFBTSxDQUFDLE1BQXNCO1FBQ25DLE9BQU8sSUFBSSxVQUFVLENBQUMsQ0FBQyxRQUFRLEVBQUUsRUFBRTtZQUNqQyxJQUFJLEdBQUcsR0FBRyx1REFBdUQsQ0FBQztZQUNsRSxNQUFNLE1BQU0sR0FBRyxRQUFRLENBQUMsY0FBYyxDQUFDLE1BQU0sQ0FBQyxrQkFBa0IsQ0FBQyxDQUFDLG9CQUFvQixDQUFDLElBQUksQ0FBQyxDQUFDO1lBQzdGLEtBQUssSUFBSSxDQUFDLEdBQUcsQ0FBQyxFQUFFLENBQUMsR0FBRyxNQUFNLENBQUMsTUFBTSxFQUFFLENBQUMsRUFBRSxFQUFFO2dCQUN0QyxNQUFNLFFBQVEsR0FBRyxNQUFNLENBQUMsQ0FBQyxDQUFDLENBQUM7Z0JBQzNCLElBQUksUUFBUSxDQUFDLEtBQUssQ0FBQyxNQUFNLEdBQUcsQ0FBQyxFQUFFO29CQUM3QixHQUFHLElBQUksZUFBZSxHQUFHLFFBQVEsQ0FBQyxLQUFLLENBQUMsQ0FBQyxDQUFDLENBQUMsV0FBVyxHQUFHLE1BQU0sQ0FBQztvQkFDaEUsS0FBSyxJQUFJLENBQUMsR0FBRyxDQUFDLEVBQUUsQ0FBQyxHQUFHLFFBQVEsQ0FBQyxLQUFLLENBQUMsTUFBTSxFQUFFLEVBQUUsQ0FBQyxFQUFFO3dCQUM5QyxHQUFHLElBQUksVUFBVSxHQUFHLFFBQVEsQ0FBQyxLQUFLLENBQUMsQ0FBQyxDQUFDLENBQUMsV0FBVyxHQUFHLFdBQVcsQ0FBQztxQkFDakU7b0JBQ0QsR0FBRyxJQUFJLFlBQVksQ0FBQztpQkFDckI7YUFDRjtZQUNELEdBQUcsSUFBSSxtQkFBbUIsQ0FBQztZQUMzQixNQUFNLE1BQU0sR0FBRyx1QkFBdUIsR0FBRyxJQUFJLENBQUMsSUFBSSxDQUFDLEdBQUcsQ0FBQyxDQUFDO1lBQ3hELElBQUksTUFBTSxDQUFDLFFBQVEsRUFBRTtnQkFDbkIsSUFBSSxDQUFDLG1CQUFtQixDQUFDLE1BQU0sQ0FBQyxRQUFRLEVBQUUsTUFBTSxDQUFDLENBQUM7Z0JBQ2xELFFBQVEsQ0FBQyxJQUFJLEVBQUUsQ0FBQzthQUNqQjtpQkFBTTtnQkFDTCxRQUFRLENBQUMsSUFBSSxDQUFDLE1BQU0sQ0FBQyxDQUFDO2FBQ3ZCO1lBQ0QsUUFBUSxDQUFDLFFBQVEsRUFBRSxDQUFDO1FBQ3RCLENBQUMsQ0FBQyxDQUFDO0lBQ0wsQ0FBQztJQUVPLElBQUksQ0FBQyxPQUFlO1FBQzFCLE9BQU8sSUFBSSxDQUFDLFFBQVEsQ0FBQyxrQkFBa0IsQ0FBQyxPQUFPLENBQUMsQ0FBQyxDQUFDLENBQUM7SUFDckQsQ0FBQzs7NEdBM1RVLGVBQWU7Z0hBQWYsZUFBZTsyRkFBZixlQUFlO2tCQUQzQixVQUFVIiwic291cmNlc0NvbnRlbnQiOlsiaW1wb3J0IHsgSW5qZWN0YWJsZSB9IGZyb20gJ0Bhbmd1bGFyL2NvcmUnO1xyXG5pbXBvcnQgeyBPYnNlcnZhYmxlIH0gZnJvbSAncnhqcyc7XHJcblxyXG5pbXBvcnQgeyBFeHBvcnRBc0NvbmZpZyB9IGZyb20gJy4vZXhwb3J0LWFzLWNvbmZpZy5tb2RlbCc7XHJcblxyXG5pbXBvcnQgaHRtbDJjYW52YXMgZnJvbSAnaHRtbDJjYW52YXMnO1xyXG5pbXBvcnQgKiBhcyBYTFNYIGZyb20gJ3hsc3gnO1xyXG5pbXBvcnQgSFRNTHRvRE9DWCBmcm9tICdodG1sLXRvLWRvY3gnO1xyXG5pbXBvcnQgaHRtbDJwZGYgZnJvbSAnaHRtbDJwZGYuanMnO1xyXG53aW5kb3dbJ2h0bWwyY2FudmFzJ10gPSBodG1sMmNhbnZhcztcclxuXHJcbkBJbmplY3RhYmxlKClcclxuZXhwb3J0IGNsYXNzIEV4cG9ydEFzU2VydmljZSB7XHJcblxyXG4gIGNvbnN0cnVjdG9yKCkgeyB9XHJcblxyXG4gIC8qKlxyXG4gICAqIE1haW4gYmFzZTY0IGdldCBtZXRob2QsIGl0IHdpbGwgcmV0dXJuIHRoZSBmaWxlIGFzIGJhc2U2NCBzdHJpbmdcclxuICAgKiBAcGFyYW0gY29uZmlnIHlvdXIgY29uZmlnXHJcbiAgICovXHJcbiAgZ2V0KGNvbmZpZzogRXhwb3J0QXNDb25maWcpOiBPYnNlcnZhYmxlPHN0cmluZyB8IG51bGw+IHtcclxuICAgIC8vIHN0cnVjdHVyZSBtZXRob2QgbmFtZSBkeW5hbWljYWxseSBieSB0eXBlXHJcbiAgICBjb25zdCBmdW5jID0gJ2dldCcgKyBjb25maWcudHlwZS50b1VwcGVyQ2FzZSgpO1xyXG4gICAgLy8gaWYgdHlwZSBzdXBwb3J0ZWQgZXhlY3V0ZSBhbmQgcmV0dXJuXHJcbiAgICBpZiAodGhpc1tmdW5jXSkge1xyXG4gICAgICByZXR1cm4gdGhpc1tmdW5jXShjb25maWcpO1xyXG4gICAgfVxyXG5cclxuICAgIC8vIHRocm93IGVycm9yIGZvciB1bnN1cHBvcnRlZCBmb3JtYXRzXHJcbiAgICByZXR1cm4gbmV3IE9ic2VydmFibGUoKG9ic2VydmVyKSA9PiB7IG9ic2VydmVyLmVycm9yKCdFeHBvcnQgdHlwZSBpcyBub3Qgc3VwcG9ydGVkLicpOyB9KTtcclxuICB9XHJcblxyXG4gIC8qKlxyXG4gICAqIFNhdmUgZXhwb3J0ZWQgZmlsZSBpbiBvbGQgamF2YXNjcmlwdCB3YXlcclxuICAgKiBAcGFyYW0gY29uZmlnIHlvdXIgY3VzdG9tIGNvbmZpZ1xyXG4gICAqIEBwYXJhbSBmaWxlTmFtZSBOYW1lIG9mIHRoZSBmaWxlIHRvIGJlIHNhdmVkIGFzXHJcbiAgICovXHJcbiAgc2F2ZShjb25maWc6IEV4cG9ydEFzQ29uZmlnLCBmaWxlTmFtZTogc3RyaW5nKTogT2JzZXJ2YWJsZTxzdHJpbmcgfCBudWxsPiB7XHJcbiAgICAvLyBzZXQgZG93bmxvYWRcclxuICAgIGNvbmZpZy5kb3dubG9hZCA9IHRydWU7XHJcbiAgICAvLyBnZXQgZmlsZSBuYW1lIHdpdGggdHlwZVxyXG4gICAgY29uZmlnLmZpbGVOYW1lID0gZmlsZU5hbWUgKyAnLicgKyBjb25maWcudHlwZTtcclxuICAgIHJldHVybiB0aGlzLmdldChjb25maWcpO1xyXG4gIH1cclxuXHJcbiAgLyoqXHJcbiAgICogQ29udmVydHMgY29udGVudCBzdHJpbmcgdG8gYmxvYiBvYmplY3RcclxuICAgKiBAcGFyYW0gY29udGVudCBzdHJpbmcgdG8gYmUgY29udmVydGVkXHJcbiAgICovXHJcbiAgY29udGVudFRvQmxvYihjb250ZW50OiBzdHJpbmcpOiBPYnNlcnZhYmxlPEJsb2I+IHtcclxuICAgIHJldHVybiBuZXcgT2JzZXJ2YWJsZSgob2JzZXJ2ZXIpID0+IHtcclxuICAgICAgLy8gZ2V0IGNvbnRlbnQgc3RyaW5nIGFuZCBleHRyYWN0IG1pbWUgdHlwZVxyXG4gICAgICBjb25zdCBhcnIgPSBjb250ZW50LnNwbGl0KCcsJyksIG1pbWUgPSBhcnJbMF0ubWF0Y2goLzooLio/KTsvKVsxXSxcclxuICAgICAgICBic3RyID0gYXRvYihhcnJbMV0pO1xyXG4gICAgICBsZXQgbiA9IGJzdHIubGVuZ3RoO1xyXG4gICAgICBjb25zdCB1OGFyciA9IG5ldyBVaW50OEFycmF5KG4pO1xyXG4gICAgICB3aGlsZSAobi0tKSB7XHJcbiAgICAgICAgdThhcnJbbl0gPSBic3RyLmNoYXJDb2RlQXQobik7XHJcbiAgICAgIH1cclxuICAgICAgb2JzZXJ2ZXIubmV4dChuZXcgQmxvYihbdThhcnJdLCB7IHR5cGU6IG1pbWUgfSkpO1xyXG4gICAgICBvYnNlcnZlci5jb21wbGV0ZSgpO1xyXG4gICAgfSk7XHJcbiAgfVxyXG5cclxuICAvKipcclxuICAgKiBSZW1vdmVzIGJhc2U2NCBmaWxlIHR5cGUgZnJvbSBhIHN0cmluZyBsaWtlIFwiZGF0YTp0ZXh0L2NzdjtiYXNlNjQsXCJcclxuICAgKiBAcGFyYW0gZmlsZUNvbnRlbnQgdGhlIGJhc2U2NCBzdHJpbmcgdG8gcmVtb3ZlIHRoZSB0eXBlIGZyb21cclxuICAgKi9cclxuICByZW1vdmVGaWxlVHlwZUZyb21CYXNlNjQoZmlsZUNvbnRlbnQ6IHN0cmluZyk6IHN0cmluZyB7XHJcbiAgICBjb25zdCByZSA9IC9eZGF0YTpbXl0qO2Jhc2U2NCwvZztcclxuICAgIGNvbnN0IG5ld0NvbnRlbnQ6IHN0cmluZyA9IHJlW1N5bWJvbC5yZXBsYWNlXShmaWxlQ29udGVudCwgJycpO1xyXG4gICAgcmV0dXJuIG5ld0NvbnRlbnQ7XHJcbiAgfVxyXG5cclxuICAvKipcclxuICAgKiBTdHJ1Y3R1cmUgdGhlIGJhc2U2NCBmaWxlIGNvbnRlbnQgd2l0aCB0aGUgZmlsZSB0eXBlIHN0cmluZ1xyXG4gICAqIEBwYXJhbSBmaWxlQ29udGVudCBmaWxlIGNvbnRlbnRcclxuICAgKiBAcGFyYW0gZmlsZU1pbWUgZmlsZSBtaW1lIHR5cGUgXCJ0ZXh0L2NzdlwiXHJcbiAgICovXHJcbiAgYWRkRmlsZVR5cGVUb0Jhc2U2NChmaWxlQ29udGVudDogc3RyaW5nLCBmaWxlTWltZTogc3RyaW5nKTogc3RyaW5nIHtcclxuICAgIHJldHVybiBgZGF0YToke2ZpbGVNaW1lfTtiYXNlNjQsJHtmaWxlQ29udGVudH1gO1xyXG4gIH1cclxuXHJcbiAgLyoqXHJcbiAgICogY3JlYXRlIGRvd25sb2FkYWJsZSBmaWxlIGZyb20gZGF0YVVSTFxyXG4gICAqIEBwYXJhbSBmaWxlTmFtZSBkb3dubG9hZGFibGUgZmlsZSBuYW1lXHJcbiAgICogQHBhcmFtIGRhdGFVUkwgZmlsZSBjb250ZW50IGFzIGRhdGFVUkxcclxuICAgKi9cclxuICBkb3dubG9hZEZyb21EYXRhVVJMKGZpbGVOYW1lOiBzdHJpbmcsIGRhdGFVUkw6IHN0cmluZyk6IHZvaWQge1xyXG4gICAgLy8gY3JlYXRlIGJsb2JcclxuICAgIHRoaXMuY29udGVudFRvQmxvYihkYXRhVVJMKS5zdWJzY3JpYmUoYmxvYiA9PiB7XHJcbiAgICAgIC8vIGRvd25sb2FkIHRoZSBibG9iXHJcbiAgICAgIHRoaXMuZG93bmxvYWRGcm9tQmxvYihibG9iLCBmaWxlTmFtZSk7XHJcbiAgICB9KTtcclxuICB9XHJcblxyXG4gIC8qKlxyXG4gICAqIERvd25sb2FkcyB0aGUgYmxvYiBvYmplY3QgYXMgYSBmaWxlXHJcbiAgICogQHBhcmFtIGJsb2IgZmlsZSBvYmplY3QgYXMgYmxvYlxyXG4gICAqIEBwYXJhbSBmaWxlTmFtZSBkb3dubG9hZGFibGUgZmlsZSBuYW1lXHJcbiAgICovXHJcbiAgZG93bmxvYWRGcm9tQmxvYihibG9iOiBCbG9iLCBmaWxlTmFtZTogc3RyaW5nKSB7XHJcbiAgICAvLyBnZXQgb2JqZWN0IHVybFxyXG4gICAgY29uc3QgdXJsID0gd2luZG93LlVSTC5jcmVhdGVPYmplY3RVUkwoYmxvYik7XHJcbiAgICAvLyBjaGVjayBmb3IgbWljcm9zb2Z0IGludGVybmV0IGV4cGxvcmVyXHJcbiAgICBpZiAod2luZG93Lm5hdmlnYXRvciAmJiB3aW5kb3cubmF2aWdhdG9yWydtc1NhdmVPck9wZW5CbG9iJ10pIHtcclxuICAgICAgLy8gdXNlIElFIGRvd25sb2FkIG9yIG9wZW4gaWYgdGhlIHVzZXIgdXNpbmcgSUVcclxuICAgICAgd2luZG93Lm5hdmlnYXRvclsnbXNTYXZlT3JPcGVuQmxvYiddKGJsb2IsIGZpbGVOYW1lKTtcclxuICAgIH0gZWxzZSB7XHJcbiAgICAgIHRoaXMuc2F2ZUZpbGUoZmlsZU5hbWUsIHVybCk7XHJcbiAgICB9XHJcbiAgfVxyXG5cclxuICBwcml2YXRlIHNhdmVGaWxlKGZpbGVOYW1lOiBzdHJpbmcsIHVybDogc3RyaW5nKSB7XHJcbiAgICAvLyBpZiBub3QgdXNpbmcgSUUgdGhlbiBjcmVhdGUgbGluayBlbGVtZW50XHJcbiAgICBjb25zdCBlbGVtZW50ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYScpO1xyXG4gICAgLy8gc2V0IGRvd25sb2FkIGF0dHIgd2l0aCBmaWxlIG5hbWVcclxuICAgIGVsZW1lbnQuc2V0QXR0cmlidXRlKCdkb3dubG9hZCcsIGZpbGVOYW1lKTtcclxuICAgIC8vIHNldCB0aGUgZWxlbWVudCBhcyBoaWRkZW5cclxuICAgIGVsZW1lbnQuc3R5bGUuZGlzcGxheSA9ICdub25lJztcclxuICAgIC8vIGFwcGVuZCB0aGUgYm9keVxyXG4gICAgZG9jdW1lbnQuYm9keS5hcHBlbmRDaGlsZChlbGVtZW50KTtcclxuICAgIC8vIHNldCBocmVmIGF0dHJcclxuICAgIGVsZW1lbnQuaHJlZiA9IHVybDtcclxuICAgIC8vIGNsaWNrIG9uIGl0IHRvIHN0YXJ0IGRvd25sb2FkaW5nXHJcbiAgICBlbGVtZW50LmNsaWNrKCk7XHJcbiAgICAvLyByZW1vdmUgdGhlIGxpbmsgZnJvbSB0aGUgZG9tXHJcbiAgICBkb2N1bWVudC5ib2R5LnJlbW92ZUNoaWxkKGVsZW1lbnQpO1xyXG4gIH1cclxuXHJcbiAgcHJpdmF0ZSBnZXRQREYoY29uZmlnOiBFeHBvcnRBc0NvbmZpZyk6IE9ic2VydmFibGU8c3RyaW5nIHwgbnVsbD4ge1xyXG4gICAgcmV0dXJuIG5ldyBPYnNlcnZhYmxlKChvYnNlcnZlcikgPT4ge1xyXG4gICAgICBpZiAoIWNvbmZpZy5vcHRpb25zKSB7XHJcbiAgICAgICAgY29uZmlnLm9wdGlvbnMgPSB7fTtcclxuICAgICAgfVxyXG4gICAgICBjb25maWcub3B0aW9ucy5maWxlbmFtZSA9IGNvbmZpZy5maWxlTmFtZTtcclxuICAgICAgY29uc3QgZWxlbWVudDogSFRNTEVsZW1lbnQgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZChjb25maWcuZWxlbWVudElkT3JDb250ZW50KTtcclxuICAgICAgY29uc3QgcGRmID0gaHRtbDJwZGYoKS5zZXQoY29uZmlnLm9wdGlvbnMpLmZyb20oZWxlbWVudCA/IGVsZW1lbnQgOiBjb25maWcuZWxlbWVudElkT3JDb250ZW50KTtcclxuXHJcbiAgICAgIGNvbnN0IGRvd25sb2FkID0gY29uZmlnLmRvd25sb2FkO1xyXG4gICAgICBjb25zdCBwZGZDYWxsYmFja0ZuID0gY29uZmlnLm9wdGlvbnMucGRmQ2FsbGJhY2tGbjtcclxuICAgICAgaWYgKGRvd25sb2FkKSB7XHJcbiAgICAgICAgaWYgKHBkZkNhbGxiYWNrRm4pIHtcclxuICAgICAgICAgIHRoaXMuYXBwbHlQZGZDYWxsYmFja0ZuKHBkZiwgcGRmQ2FsbGJhY2tGbikuc2F2ZSgpO1xyXG4gICAgICAgIH0gZWxzZSB7XHJcbiAgICAgICAgICBwZGYuc2F2ZSgpO1xyXG4gICAgICAgIH1cclxuICAgICAgICBvYnNlcnZlci5uZXh0KCk7XHJcbiAgICAgICAgb2JzZXJ2ZXIuY29tcGxldGUoKTtcclxuICAgICAgfSBlbHNlIHtcclxuICAgICAgICBpZiAocGRmQ2FsbGJhY2tGbikge1xyXG4gICAgICAgICAgdGhpcy5hcHBseVBkZkNhbGxiYWNrRm4ocGRmLCBwZGZDYWxsYmFja0ZuKS5vdXRwdXRQZGYoJ2RhdGF1cmlzdHJpbmcnKS50aGVuKGRhdGEgPT4ge1xyXG4gICAgICAgICAgICBvYnNlcnZlci5uZXh0KGRhdGEpO1xyXG4gICAgICAgICAgICBvYnNlcnZlci5jb21wbGV0ZSgpO1xyXG4gICAgICAgICAgfSk7XHJcbiAgICAgICAgfSBlbHNlIHtcclxuICAgICAgICAgIHBkZi5vdXRwdXRQZGYoJ2RhdGF1cmlzdHJpbmcnKS50aGVuKGRhdGEgPT4ge1xyXG4gICAgICAgICAgICBvYnNlcnZlci5uZXh0KGRhdGEpO1xyXG4gICAgICAgICAgICBvYnNlcnZlci5jb21wbGV0ZSgpO1xyXG4gICAgICAgICAgfSk7XHJcbiAgICAgICAgfVxyXG4gICAgICB9XHJcbiAgICB9KTtcclxuICB9XHJcblxyXG4gIHByaXZhdGUgYXBwbHlQZGZDYWxsYmFja0ZuKHBkZiwgcGRmQ2FsbGJhY2tGbikge1xyXG4gICAgcmV0dXJuIHBkZi50b1BkZigpLmdldCgncGRmJykudGhlbigocGRmUmVmKSA9PiB7XHJcbiAgICAgIHBkZkNhbGxiYWNrRm4ocGRmUmVmKTtcclxuICAgIH0pO1xyXG4gIH1cclxuXHJcbiAgcHJpdmF0ZSBnZXRQTkcoY29uZmlnOiBFeHBvcnRBc0NvbmZpZyk6IE9ic2VydmFibGU8c3RyaW5nIHwgbnVsbD4ge1xyXG4gICAgcmV0dXJuIG5ldyBPYnNlcnZhYmxlKChvYnNlcnZlcikgPT4ge1xyXG4gICAgICBjb25zdCBlbGVtZW50OiBIVE1MRWxlbWVudCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKGNvbmZpZy5lbGVtZW50SWRPckNvbnRlbnQpO1xyXG4gICAgICBodG1sMmNhbnZhcyhlbGVtZW50LCBjb25maWcub3B0aW9ucykudGhlbigoY2FudmFzKSA9PiB7XHJcbiAgICAgICAgY29uc3QgaW1nRGF0YSA9IGNhbnZhcy50b0RhdGFVUkwoJ2ltYWdlL1BORycpO1xyXG4gICAgICAgIGlmIChjb25maWcudHlwZSA9PT0gJ3BuZycgJiYgY29uZmlnLmRvd25sb2FkKSB7XHJcbiAgICAgICAgICB0aGlzLmRvd25sb2FkRnJvbURhdGFVUkwoY29uZmlnLmZpbGVOYW1lLCBpbWdEYXRhKTtcclxuICAgICAgICAgIG9ic2VydmVyLm5leHQoKTtcclxuICAgICAgICB9IGVsc2Uge1xyXG4gICAgICAgICAgb2JzZXJ2ZXIubmV4dChpbWdEYXRhKTtcclxuICAgICAgICB9XHJcbiAgICAgICAgb2JzZXJ2ZXIuY29tcGxldGUoKTtcclxuICAgICAgfSwgZXJyID0+IHtcclxuICAgICAgICBvYnNlcnZlci5lcnJvcihlcnIpO1xyXG4gICAgICB9KTtcclxuICAgIH0pO1xyXG4gIH1cclxuXHJcbiAgcHJpdmF0ZSBnZXRDU1YoY29uZmlnOiBFeHBvcnRBc0NvbmZpZyk6IE9ic2VydmFibGU8c3RyaW5nIHwgbnVsbD4ge1xyXG4gICAgcmV0dXJuIG5ldyBPYnNlcnZhYmxlKChvYnNlcnZlcikgPT4ge1xyXG4gICAgICBjb25zdCBlbGVtZW50OiBIVE1MRWxlbWVudCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKGNvbmZpZy5lbGVtZW50SWRPckNvbnRlbnQpO1xyXG4gICAgICBjb25zdCBjc3YgPSBbXTtcclxuICAgICAgY29uc3Qgcm93czogYW55ID0gZWxlbWVudC5xdWVyeVNlbGVjdG9yQWxsKCd0YWJsZSB0cicpO1xyXG4gICAgICBmb3IgKGxldCBpbmRleCA9IDA7IGluZGV4IDwgcm93cy5sZW5ndGg7IGluZGV4KyspIHtcclxuICAgICAgICBjb25zdCByb3dFbGVtZW50ID0gcm93c1tpbmRleF07XHJcbiAgICAgICAgY29uc3Qgcm93ID0gW107XHJcbiAgICAgICAgY29uc3QgY29scyA9IHJvd0VsZW1lbnQucXVlcnlTZWxlY3RvckFsbCgndGQsIHRoJyk7XHJcbiAgICAgICAgZm9yIChsZXQgY29sSW5kZXggPSAwOyBjb2xJbmRleCA8IGNvbHMubGVuZ3RoOyBjb2xJbmRleCsrKSB7XHJcbiAgICAgICAgICBjb25zdCBjb2wgPSBjb2xzW2NvbEluZGV4XTtcclxuICAgICAgICAgIHJvdy5wdXNoKCdcIicrY29sLmlubmVyVGV4dCsnXCInKTtcclxuICAgICAgICB9XHJcbiAgICAgICAgY3N2LnB1c2gocm93LmpvaW4oJywnKSk7XHJcbiAgICAgIH1cclxuICAgICAgY29uc3QgY3N2Q29udGVudCA9ICdkYXRhOnRleHQvY3N2O2Jhc2U2NCwnICsgdGhpcy5idG9hKGNzdi5qb2luKCdcXG4nKSk7XHJcbiAgICAgIGlmIChjb25maWcuZG93bmxvYWQpIHtcclxuICAgICAgICB0aGlzLmRvd25sb2FkRnJvbURhdGFVUkwoY29uZmlnLmZpbGVOYW1lLCBjc3ZDb250ZW50KTtcclxuICAgICAgICBvYnNlcnZlci5uZXh0KCk7XHJcbiAgICAgIH0gZWxzZSB7XHJcbiAgICAgICAgb2JzZXJ2ZXIubmV4dChjc3ZDb250ZW50KTtcclxuICAgICAgfVxyXG4gICAgICBvYnNlcnZlci5jb21wbGV0ZSgpO1xyXG4gICAgfSk7XHJcbiAgfVxyXG5cclxuICBwcml2YXRlIGdldFRYVChjb25maWc6IEV4cG9ydEFzQ29uZmlnKTogT2JzZXJ2YWJsZTxzdHJpbmcgfCBudWxsPiB7XHJcbiAgICBjb25zdCBuYW1lRnJhZ3MgPSBjb25maWcuZmlsZU5hbWUuc3BsaXQoJy4nKTtcclxuICAgIGNvbmZpZy5maWxlTmFtZSA9IGAke25hbWVGcmFnc1swXX0udHh0YDtcclxuICAgIHJldHVybiB0aGlzLmdldENTVihjb25maWcpO1xyXG4gIH1cclxuXHJcbiAgcHJpdmF0ZSBnZXRYTFMoY29uZmlnOiBFeHBvcnRBc0NvbmZpZyk6IE9ic2VydmFibGU8c3RyaW5nIHwgbnVsbD4ge1xyXG4gICAgcmV0dXJuIG5ldyBPYnNlcnZhYmxlKChvYnNlcnZlcikgPT4ge1xyXG5cclxuICAgICAgY29uc3QgZWxlbWVudDogSFRNTEVsZW1lbnQgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZChjb25maWcuZWxlbWVudElkT3JDb250ZW50KTtcclxuICAgICAgY29uc3Qgd3MzID0gWExTWC51dGlscy50YWJsZV90b19zaGVldChlbGVtZW50LCBjb25maWcub3B0aW9ucyk7XHJcbiAgICAgIGNvbnN0IHdiID0gWExTWC51dGlscy5ib29rX25ldygpO1xyXG4gICAgICBYTFNYLnV0aWxzLmJvb2tfYXBwZW5kX3NoZWV0KHdiLCB3czMsIGNvbmZpZy5maWxlTmFtZSk7XHJcbiAgICAgIGNvbnN0IG91dCA9IFhMU1gud3JpdGUod2IsIHsgdHlwZTogJ2Jhc2U2NCcgfSk7XHJcbiAgICAgIGNvbnN0IHhsc0NvbnRlbnQgPSAnZGF0YTphcHBsaWNhdGlvbi92bmQub3BlbnhtbGZvcm1hdHMtb2ZmaWNlZG9jdW1lbnQuc3ByZWFkc2hlZXRtbC5zaGVldDtiYXNlNjQsJyArIG91dDtcclxuICAgICAgaWYgKGNvbmZpZy5kb3dubG9hZCkge1xyXG4gICAgICAgIHRoaXMuZG93bmxvYWRGcm9tRGF0YVVSTChjb25maWcuZmlsZU5hbWUsIHhsc0NvbnRlbnQpO1xyXG4gICAgICAgIG9ic2VydmVyLm5leHQoKTtcclxuICAgICAgfSBlbHNlIHtcclxuICAgICAgICBvYnNlcnZlci5uZXh0KHhsc0NvbnRlbnQpO1xyXG4gICAgICB9XHJcbiAgICAgIG9ic2VydmVyLmNvbXBsZXRlKCk7XHJcbiAgICB9KTtcclxuICB9XHJcblxyXG4gIHByaXZhdGUgZ2V0WExTWChjb25maWc6IEV4cG9ydEFzQ29uZmlnKTogT2JzZXJ2YWJsZTxzdHJpbmcgfCBudWxsPiB7XHJcbiAgICByZXR1cm4gdGhpcy5nZXRYTFMoY29uZmlnKTtcclxuICB9XHJcblxyXG4gIHByaXZhdGUgZ2V0RE9DWChjb25maWc6IEV4cG9ydEFzQ29uZmlnKTogT2JzZXJ2YWJsZTxzdHJpbmcgfCBudWxsPiB7XHJcbiAgICByZXR1cm4gbmV3IE9ic2VydmFibGUoKG9ic2VydmVyKSA9PiB7XHJcbiAgICAgIGNvbnN0IGNvbnRlbnREb2N1bWVudDogc3RyaW5nID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoY29uZmlnLmVsZW1lbnRJZE9yQ29udGVudCkub3V0ZXJIVE1MO1xyXG4gICAgICBjb25zdCBjb250ZW50ID0gJzwhRE9DVFlQRSBodG1sPicgKyBjb250ZW50RG9jdW1lbnQ7XHJcbiAgICAgIEhUTUx0b0RPQ1goY29udGVudCwgbnVsbCwgY29uZmlnLm9wdGlvbnMpLnRoZW4oY29udmVydGVkID0+IHtcclxuICAgICAgICBpZiAoY29uZmlnLmRvd25sb2FkKSB7XHJcbiAgICAgICAgICB0aGlzLmRvd25sb2FkRnJvbUJsb2IoY29udmVydGVkLCBjb25maWcuZmlsZU5hbWUpO1xyXG4gICAgICAgICAgb2JzZXJ2ZXIubmV4dCgpO1xyXG4gICAgICAgICAgb2JzZXJ2ZXIuY29tcGxldGUoKTtcclxuICAgICAgICB9IGVsc2Uge1xyXG4gICAgICAgICAgY29uc3QgcmVhZGVyID0gbmV3IEZpbGVSZWFkZXIoKTtcclxuICAgICAgICAgIHJlYWRlci5vbmxvYWRlbmQgPSAoKSA9PiB7XHJcbiAgICAgICAgICAgIGNvbnN0IGJhc2U2NGRhdGEgPSByZWFkZXIucmVzdWx0IGFzIHN0cmluZztcclxuICAgICAgICAgICAgb2JzZXJ2ZXIubmV4dChiYXNlNjRkYXRhKTtcclxuICAgICAgICAgICAgb2JzZXJ2ZXIuY29tcGxldGUoKTtcclxuICAgICAgICAgIH07XHJcbiAgICAgICAgICByZWFkZXIucmVhZEFzRGF0YVVSTChjb252ZXJ0ZWQpO1xyXG4gICAgICAgIH1cclxuICAgICAgfSk7XHJcbiAgICB9KTtcclxuICB9XHJcblxyXG4gIHByaXZhdGUgZ2V0RE9DKGNvbmZpZzogRXhwb3J0QXNDb25maWcpOiBPYnNlcnZhYmxlPHN0cmluZyB8IG51bGw+IHtcclxuICAgIHJldHVybiB0aGlzLmdldERPQ1goY29uZmlnKTtcclxuICB9XHJcblxyXG4gIHByaXZhdGUgZ2V0SlNPTihjb25maWc6IEV4cG9ydEFzQ29uZmlnKTogT2JzZXJ2YWJsZTxhbnlbXSB8IG51bGw+IHtcclxuICAgIHJldHVybiBuZXcgT2JzZXJ2YWJsZSgob2JzZXJ2ZXIpID0+IHtcclxuICAgICAgY29uc3QgZGF0YSA9IFtdOyAvLyBmaXJzdCByb3cgbmVlZHMgdG8gYmUgaGVhZGVyc1xyXG4gICAgICBjb25zdCBoZWFkZXJzID0gW107XHJcbiAgICAgIGNvbnN0IHRhYmxlID0gPEhUTUxUYWJsZUVsZW1lbnQ+ZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoY29uZmlnLmVsZW1lbnRJZE9yQ29udGVudCk7XHJcbiAgICAgIGZvciAobGV0IGluZGV4ID0gMDsgaW5kZXggPCB0YWJsZS5yb3dzWzBdLmNlbGxzLmxlbmd0aDsgaW5kZXgrKykge1xyXG4gICAgICAgIGhlYWRlcnNbaW5kZXhdID0gdGFibGUucm93c1swXS5jZWxsc1tpbmRleF0uaW5uZXJIVE1MLnRvTG93ZXJDYXNlKCkucmVwbGFjZSgvIC9naSwgJycpO1xyXG4gICAgICB9XHJcbiAgICAgIC8vIGdvIHRocm91Z2ggY2VsbHNcclxuICAgICAgZm9yIChsZXQgaSA9IDE7IGkgPCB0YWJsZS5yb3dzLmxlbmd0aDsgaSsrKSB7XHJcbiAgICAgICAgY29uc3QgdGFibGVSb3cgPSB0YWJsZS5yb3dzW2ldOyBjb25zdCByb3dEYXRhID0ge307XHJcbiAgICAgICAgZm9yIChsZXQgaiA9IDA7IGogPCB0YWJsZVJvdy5jZWxscy5sZW5ndGg7IGorKykge1xyXG4gICAgICAgICAgcm93RGF0YVtoZWFkZXJzW2pdXSA9IHRhYmxlUm93LmNlbGxzW2pdLmlubmVySFRNTDtcclxuICAgICAgICB9XHJcbiAgICAgICAgZGF0YS5wdXNoKHJvd0RhdGEpO1xyXG4gICAgICB9XHJcbiAgICAgIGNvbnN0IGpzb25TdHJpbmcgPSBKU09OLnN0cmluZ2lmeShkYXRhKTtcclxuICAgICAgY29uc3QganNvbkJhc2U2NCA9IHRoaXMuYnRvYShqc29uU3RyaW5nKTtcclxuICAgICAgY29uc3QgZGF0YVN0ciA9ICdkYXRhOnRleHQvanNvbjtiYXNlNjQsJyArIGpzb25CYXNlNjQ7XHJcbiAgICAgIGlmIChjb25maWcuZG93bmxvYWQpIHtcclxuICAgICAgICB0aGlzLmRvd25sb2FkRnJvbURhdGFVUkwoY29uZmlnLmZpbGVOYW1lLCBkYXRhU3RyKTtcclxuICAgICAgICBvYnNlcnZlci5uZXh0KCk7XHJcbiAgICAgIH0gZWxzZSB7XHJcbiAgICAgICAgb2JzZXJ2ZXIubmV4dChkYXRhKTtcclxuICAgICAgfVxyXG4gICAgICBvYnNlcnZlci5jb21wbGV0ZSgpO1xyXG4gICAgfSk7XHJcbiAgfVxyXG5cclxuICBwcml2YXRlIGdldFhNTChjb25maWc6IEV4cG9ydEFzQ29uZmlnKTogT2JzZXJ2YWJsZTxzdHJpbmcgfCBudWxsPiB7XHJcbiAgICByZXR1cm4gbmV3IE9ic2VydmFibGUoKG9ic2VydmVyKSA9PiB7XHJcbiAgICAgIGxldCB4bWwgPSAnPD94bWwgdmVyc2lvbj1cIjEuMFwiIGVuY29kaW5nPVwiVVRGLThcIj8+PFJvb3Q+PENsYXNzZXM+JztcclxuICAgICAgY29uc3QgdHJpdGVtID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoY29uZmlnLmVsZW1lbnRJZE9yQ29udGVudCkuZ2V0RWxlbWVudHNCeVRhZ05hbWUoJ3RyJyk7XHJcbiAgICAgIGZvciAobGV0IGkgPSAwOyBpIDwgdHJpdGVtLmxlbmd0aDsgaSsrKSB7XHJcbiAgICAgICAgY29uc3QgY2VsbGRhdGEgPSB0cml0ZW1baV07XHJcbiAgICAgICAgaWYgKGNlbGxkYXRhLmNlbGxzLmxlbmd0aCA+IDApIHtcclxuICAgICAgICAgIHhtbCArPSAnPENsYXNzIG5hbWU9XCInICsgY2VsbGRhdGEuY2VsbHNbMF0udGV4dENvbnRlbnQgKyAnXCI+XFxuJztcclxuICAgICAgICAgIGZvciAobGV0IG0gPSAxOyBtIDwgY2VsbGRhdGEuY2VsbHMubGVuZ3RoOyArK20pIHtcclxuICAgICAgICAgICAgeG1sICs9ICdcXHQ8ZGF0YT4nICsgY2VsbGRhdGEuY2VsbHNbbV0udGV4dENvbnRlbnQgKyAnPC9kYXRhPlxcbic7XHJcbiAgICAgICAgICB9XHJcbiAgICAgICAgICB4bWwgKz0gJzwvQ2xhc3M+XFxuJztcclxuICAgICAgICB9XHJcbiAgICAgIH1cclxuICAgICAgeG1sICs9ICc8L0NsYXNzZXM+PC9Sb290Pic7XHJcbiAgICAgIGNvbnN0IGJhc2U2NCA9ICdkYXRhOnRleHQveG1sO2Jhc2U2NCwnICsgdGhpcy5idG9hKHhtbCk7XHJcbiAgICAgIGlmIChjb25maWcuZG93bmxvYWQpIHtcclxuICAgICAgICB0aGlzLmRvd25sb2FkRnJvbURhdGFVUkwoY29uZmlnLmZpbGVOYW1lLCBiYXNlNjQpO1xyXG4gICAgICAgIG9ic2VydmVyLm5leHQoKTtcclxuICAgICAgfSBlbHNlIHtcclxuICAgICAgICBvYnNlcnZlci5uZXh0KGJhc2U2NCk7XHJcbiAgICAgIH1cclxuICAgICAgb2JzZXJ2ZXIuY29tcGxldGUoKTtcclxuICAgIH0pO1xyXG4gIH1cclxuXHJcbiAgcHJpdmF0ZSBidG9hKGNvbnRlbnQ6IHN0cmluZykge1xyXG4gICAgcmV0dXJuIGJ0b2EodW5lc2NhcGUoZW5jb2RlVVJJQ29tcG9uZW50KGNvbnRlbnQpKSk7XHJcbiAgfVxyXG5cclxufVxyXG4iXX0=