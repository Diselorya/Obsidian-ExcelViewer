import { App, MarkdownView, WorkspaceLeaf, TFile } from 'obsidian';
import * as XLSX from 'xlsx';
import * as C from "../constants";
import { hasLoadedEmbeddedExcel, } from "../render/embedded/embed-utils";
import xlinfo from "./excelembedinfo";
import * as xlstyle from "./excelembedstyle";
import { table } from 'console';

export async function excelEmbedding(app: App, leaf: WorkspaceLeaf, linkEl: HTMLElement) {
    console.log("excelEmbedding.linkEl", linkEl);

	// 判断是否已嵌入了 Excel 表格，如果已经嵌入了，就不再处理，否则会因为异步导致重复嵌入
	if (hasLoadedEmbeddedExcel(linkEl)) return;

	// 获取链接中的 Excel 嵌入信息
	let info = new xlinfo();
	info.loadFormLinkEl(linkEl);
	if (!info.isExcel) return;
	
	// 创建一个容器，用于放置 Excel 表格
	let containerEl = linkEl.createDiv({ cls: C.EXCEL_VIEW_EL_NAME,	}); 	
	console.log("excelEmbedding.linkEl", linkEl);

    // 从链接中解析出 Excel 文件的 TFile 对象
	if (!info.file) {
		const sourcePath = (leaf.view as MarkdownView).file?.path ?? "";
		console.log("sourcePath", sourcePath);		
		info.file = app.metadataCache.getFirstLinkpathDest(info.name ?? '', sourcePath) ?? undefined;
		console.log("tfile", info.file);
	}

    // 读取 Excel 文件，如果在 Obsidian 库中，则通过 Obsidian 库读取，否则通过 XLSX 库直接读取链接地址
    let wb: XLSX.WorkBook;
    if (info.file) {
        wb = XLSX.read(await app.vault.readBinary(info.file), {type: 'array'});
    } else {
        wb = XLSX.readFile(info.name ?? '');
    }
    console.log("workbook", wb);

    // 创建一个表格显示 Excel 数据，并放入容器中
    let tabEl = createTableFromWorkbooks(wb, info);
	containerEl.appendChild(tabEl);  

	// 设置容器的样式
	xlstyle.setContainerStyle(containerEl, info);

    console.log("renderContainerEl", containerEl);

    return;
}


function createTableFromWorkbooks(wb: XLSX.WorkBook, info?: xlinfo): HTMLDivElement {    
    // 创建一个分页切换容器
    let tabContainer = document.createElement('div');
	tabContainer.addClass(C.WORKBOOK_EL_NAME);

    // 创建一个页签标题容器
    let tabTitleContainer = document.createElement('div');
    tabTitleContainer.classList.add('worksheet-names');

	let maxWidth = 0;
	let maxHeight = 0;

    // 遍历工作簿中的每个工作表
    for (let sheetName in wb.Sheets) {
        // 获取工作表
        let sheet = wb.Sheets[sheetName];

        // 创建一个表格
        let table = createTableFromWorksheet(sheet, info);

        // 创建一个分页
        let tab = document.createElement('div');
        tab.style.display = 'none'; // 默认隐藏所有分页
		tab.setAttribute("name", sheetName);
		
		maxWidth = Math.max(maxWidth, table.scrollWidth);
		maxHeight = Math.max(maxHeight, table.scrollHeight);

        // 将表格添加到分页中
        tab.appendChild(table);

        // 将分页添加到分页切换容器中
        tabContainer.appendChild(tab);

        // 创建一个页签标题
        let tabTitle = document.createElement('button');
        tabTitle.textContent = sheetName;

        // 当点击页签标题时，切换到对应的页签
        tabTitle.addEventListener('click', (event) => {   		
			console.log('点击页签标题事件：', tabTitle);	

			event.stopPropagation();
			event.preventDefault();

			switchToWorksheet(tabContainer, tabTitle.textContent ?? '');

            // 显示当前页签
            tab.style.display = 'block';
        });

        // 将页签标题添加到页签标题容器中
        tabTitleContainer.appendChild(tabTitle);
    }
	console.log('工作簿尺寸：', maxWidth, maxHeight);
	
	tabContainer.style.width = '100%';
	tabContainer.style.height = '100%';

    // 默认显示第一个分页
    if (tabContainer.firstChild) {
        switchToWorksheet(tabContainer, (tabContainer.firstChild as HTMLElement).getAttribute('name') ?? '');
    }

    // 将页签标题容器添加到分页切换容器的顶部
	xlstyle.setTabTitleStyle(tabTitleContainer);
	tabContainer.insertBefore(tabTitleContainer, tabContainer.firstChild);

	// 如果指定了工作表名称，则切换至该工作表并隐藏切换页签
	console.log('工作表名称：', info?.sheetName);
	if (info?.sheetName) {
		switchToWorksheet(tabContainer, info.sheetName);
		tabTitleContainer.style.display = 'none';
	}

    return tabContainer;
}

// 切换到指定的工作表
function switchToWorksheet(tabContainer: HTMLDivElement, sheetName: string) {
	console.log('切换到工作表.tabContainer：', tabContainer);
	// 遍历所有分页
	for (let child of Array.from(tabContainer.children)) {
		console.log('切换到工作表.child：', child);
		// 判断是否为工作表分页
		if (!child.classList.contains('worksheet-names')) {
			console.log('判断当前工作表：', child.getAttribute('name'), sheetName, child.getAttribute('name') === sheetName);
			// 如果当前分页是要切换的工作表
			if (child.getAttribute('name') === sheetName) {
				// 切换到该分页
				(child as HTMLElement).style.display = 'block';
			} else {
				// 隐藏其他分页
				(child as HTMLElement).style.display = 'none';
			}
		} else {
			for (let btn of Array.from(child.children)) {
				// 如果当前分页是要切换的工作表
				if (btn.textContent === sheetName) {
					// 高亮
					(btn as HTMLElement).style.backgroundColor = 'var(--tag-color)';
					(btn as HTMLElement).style.color = 'var(--text-on-accent)';
				} else {
					// 取消高亮
					(btn as HTMLElement).style.backgroundColor = '';
					(btn as HTMLElement).style.color = '';
				}
			}
		}
	}
}

// 创建一个表格，装入 XLSX.WORKSHEET 类型的数据，并返回该表格
function createTableFromWorksheet(worksheet: XLSX.WorkSheet, info?: xlinfo): HTMLTableElement {
    // 创建一个新的表格元素
    let tableEl = document.createElement('table');
	tableEl.addClass(C.WORKSHEET_EL_NAME);

	// 这个事件监听器会在捕获阶段被触发，这样它可以在其他事件监听器之前阻止事件的传播和默认行为，屏蔽嵌入链接的点击事件
	tableEl.addEventListener('click', (event) => {
		event.stopPropagation();
		event.preventDefault();
	}, true);

    // 获取工作表的数据
	let data: string[][] = XLSX.utils.sheet_to_json(
		worksheet, {header: 1, range: info?.range || worksheet['!ref'], defval: '', });
	console.log('工作表数据（含区域）', data);

    // 遍历每一行数据
    for (let row of data) {
        // 创建一个新的行元素
        let rowEl = document.createElement('tr');

        // 遍历每一列数据
        for (let cell of row) {
            // 创建一个新的单元格元素
            let cellEl = document.createElement('td');

            // 设置单元格的文本内容
            cellEl.textContent = cell;

            // 将单元格添加到行中
            rowEl.appendChild(cellEl);
        }

        // 将行添加到表格中
        tableEl.appendChild(rowEl);
    }

    // 设置表格的样式
    const rowHeight = 25;
    const colWidth = 100;
    // const rowHeight = parseInt(plugin.settings.rowHeight); // 假设每行的高度为20px
    // const colWidth = parseInt(plugin.settings.colWidth); // 假设每列的宽度为60px
	console.log(rowHeight, colWidth);


	// 设置表格样式
    xlstyle.setTableStyle(tableEl, info);

    // 返回创建的表格
    return tableEl;
}
