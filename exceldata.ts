import { App, Editor, MarkdownView, Modal, Notice, Plugin, PluginSettingTab, Setting, TFile, Vault } from 'obsidian';
import * as exp from 'constants';
import * as XLSX from 'xlsx';
import * as ExcelJS from 'exceljs'; 

// 判断是否为 Excel 文件的链接格式
export function isExcelLink(link: string): boolean {
    return !!link.match(/\[\[([^#]+\.xls.?)(?:#([^!|]+)?)?(?:!([^|]+))?\]\]/);
}


// 解析 Excel 引用格式
export function parseExcelReference(link: string):  {bookname: string, sheetname: string, range: string} | null {
    //let match = decodeURIComponent(link).match(/\[\[([^#]+)(?:#([^!|]+))?(?:!([^|]+))?\]\]/);
    let match = decodeURIComponent(link).match(/\[\[([^#]+)(?:#([^!|]+)?)?(?:!([^|]+))?\]\]/);
    console.log('oblink：\n' + link);
    if (match) {
        let fileName = match[1];
        let sheetName = match[2] || '';
        let range = (match[3] || '').toUpperCase();   
        return {'bookname': fileName, 'sheetname': sheetName, 'range': range};
    }
    return null;
}

// 获取 Excel 数据
export async function getExcelData(oblink: string): Promise<Record<string, any>[]> {

    const fileinfo = parseExcelReference(oblink);
    console.log(fileinfo);

    let files = this.app.vault.getFiles();
    console.log(files);

    let file = files.find((f: TFile) => f.name === fileinfo?.bookname);
    console.log(file);
    
    const relativePath: string = file ? file.path : '';
    const absolutePath: string = file ? this.app.vault.adapter.getFullPath(file.path) : '';
    console.log('相对路径：' + relativePath);
    console.log('绝对路径：' + absolutePath);

    let workbook;
    // let workbook = new ExcelJS.Workbook();
    let bookdata: ArrayBuffer = await this.app.vault.readBinary(file);
    console.log(bookdata);
    try {
        console.log(333);
        workbook = await XLSX.read(bookdata, { type: 'array' });
        // await workbook.xlsx.readFile(absolutePath);
        console.log(444);
    } catch (error) {
        console.error('Error reading the file:', error);
        return [];
    }
    console.log(555);

    // 如果 sheetName 为空则默认第一个工作表
    console.log('读取 Excel 工作表：' + fileinfo?.sheetname)
	let sheet: XLSX.WorkSheet = workbook.Sheets[fileinfo?.sheetname || workbook.SheetNames[0]];
    // 如果找不到 sheetName 则报错
    if (!sheet) {
        console.error('Error reading the Sheet:', fileinfo?.sheetname);
        return [];
    }
    console.log('成功读取工作表：');
    console.log(sheet);

    // 如果没有传入 range 参数，则默认获取该工作表的所有 usedrange
    console.log('读取 Excel 数据区域：' + fileinfo?.range)
    // range 必须为大写字母
	let data: Record<string, any>[] = XLSX.utils.sheet_to_json(sheet, { range: fileinfo?.range || sheet['!ref'] });
    console.log('成功读取数据：');
    console.log(data);
	return data;
}

// 把 Record<string, any> 转换为 csv 文本的函数
export function convertToCSV(data: Record<string, any>[]): string {
    const header = Object.keys(data[0]).join(',');
    const rows = data.map(obj => Object.values(obj).join(',')).join('\n');
    return `${header}\n${rows}`;
}