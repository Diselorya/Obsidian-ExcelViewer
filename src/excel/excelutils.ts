import * as XLSX from 'xlsx';
import path from 'path';
import * as C from "../constants";


export async function newExcelFile(
    contextMenuFolderPath: string | null
) {
    // 创建一个新的工作簿
    let wb = XLSX.utils.book_new();

    // 创建一个新的工作表
    let ws = XLSX.utils.aoa_to_sheet([[]]);

    // 将工作表添加到工作簿中
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

    // 定义文件路径
    let filePath = path.join(contextMenuFolderPath || '', C.DEFAULT_EXCEL_NAME + '.' + C.EXCEL_EXTENSION);

    // 保存工作簿到一个新的文件
    XLSX.writeFile(wb, filePath);

    return filePath;
}