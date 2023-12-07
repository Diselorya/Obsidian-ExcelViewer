import { App, Editor, MarkdownView, Modal, Notice, Plugin, PluginSettingTab, Setting } from 'obsidian';
import * as exceldata from './exceldata';
import { text } from 'stream/consumers';

// 显示当前行内容
export function getCurrentLineContent(): string {
    const editor = this.app.workspace.getActiveViewOfType(MarkdownView)?.editor;
    if (editor) {
        const cursor = editor.getCursor();
        if (cursor) {
            const line = cursor.line;
            const lineContent = editor.getLine(line);
            console.log("当前行内容：\n" + lineContent);
            return lineContent;
        }
    }
	return '';
}

// 显示测试文本
export async function getText(): Promise<string> {
    const text = getCurrentLineContent();
    let content = exceldata.parseExcelReference(text)
    console.log("解析引用:" + content);
    if (content) {
        // return content.join('\n');
        let data: Record<string, any>[] = await exceldata.getExcelData(text);
        let csv: string = exceldata.convertToCSV(data);
        console.log("csv：" + csv);
        return csv;
    } else {
        return getCurrentLineContent();
    }
}