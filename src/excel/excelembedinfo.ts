import { App, Vault, TFile } from 'obsidian';
import * as C from "../constants";


export default class ExcelEmbedInfo {
    file: TFile | undefined;
    isExcel: boolean = false;
    name: string | undefined;
    sheetName: string | undefined;
    size: {
        width: number;
        height: number;		
    };

    // Range 属性必须是大写，否则 XLSX 不认
    private _range: string | undefined = undefined;
    set range(value: string | undefined) {
        this._range = value?.toUpperCase();
    }
    get range(): string | undefined {
        return this._range;
    }

    constructor(
		file?: TFile,
		sheetName: string | undefined = undefined,
		range: string | undefined = undefined,
		width: number = 0,
		height: number = 0
	) {
        this.name = file?.name;
        this.file = file;
        this.sheetName = sheetName;
        this.range = range;
        this.size = {
            width: width,
            height: height
        };
        this.isExcel = C.EXCEL_EXTENSIONS.some(ext => file?.name.endsWith(ext));
    }

    isExcelLink(oblink: string): boolean {
        return oblink.match(C.EMBED_LINK_EXCEL_REGEX) != null;    
    }

	loadFromObLink(oblink: string, sizeStr?: string | null, app?: App): boolean {
		const matches = oblink.match(C.EMBED_LINK_EXCEL_REGEX);
		if (!matches) return false;

		let [, bookname, sheetName, range, size] = matches;
		console.log("ExcelEmbedInfo.loadFromObLink", matches);

        this.isExcel = this.isExcelLink(oblink);

        // 文件名和工作表名真有可能以空格开头或结尾
		this.name = bookname;
		this.sheetName = sheetName;

        if (app)
        {
            this.file = app.vault.getAbstractFileByPath(this.name) as TFile ?? undefined;
        }

        if (range)
        {
            range = range.trim();
            this.range = (
                    (range.match(C.EXCEL_RANGE_A1_REGEX) != null)
                    || (range.match(C.EXCEL_RANGE_R1C1_REGEX) != null)
                ) ? range : undefined;
        }
            
        // Obsidian 的链接中，“|” 后面的数据会被放在 alt 属性里，而不是 src
        if (sizeStr) size = sizeStr.trim();
        const s = size?.trim().match(C.EMBED_SIZE_REGEX);
		if (s) {
            this.size = {
                width: parseInt(s[1]),
                height: parseInt(s[2])
			}
		} else {
			this.size = {
				width: 0,
				height: 0
			};
		}
        return true;
	}

	loadFormLinkEl(linkEl: HTMLElement, app?: App): boolean {
        console.log("ExcelEmbedInfo.loadFormLinkEl", linkEl);

		const src = linkEl.getAttribute("src");
        console.log("ExcelEmbedInfo.loadFormLinkEl.src", src);
		if (!src) return false;

        // 尺寸可能会被提取到 alt 属性中，也可能是一些其他的信息
		const alt = linkEl.getAttribute("alt");
        const size = alt?.match(C.EMBED_SIZE_REGEX_STRICT)?.[0];
        console.log("ExcelEmbedInfo.loadFormLinkEl.alt", size);

		return this.loadFromObLink(src, size, app);
	}
}