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
    let tabEl = createTableFromWorkbooks(wb, info, linkEl);
	containerEl.appendChild(tabEl);  
	// 设置 table 的 name 为工作簿名称
	tabEl.setAttribute("name", info.name ?? '');

	// 设置容器的样式
	xlstyle.setContainerStyle(containerEl, info);

	// 如果指定了工作表名称，则切换至该工作表并隐藏切换页签
	const tabTitleContainer = tabEl.querySelector(`.worksheet-names`) as HTMLDivElement;
	const tabContainer = tabEl.querySelector(`.${C.WORKBOOK_EL_NAME}`) as HTMLDivElement;
	const firstSheetName = wb.SheetNames[0];

	console.log('最后切换一次工作表名称：', info?.sheetName);
	if (info?.sheetName && tabContainer && firstSheetName) {
		tabTitleContainer.style.display = 'none';
		requestAnimationFrame(() => {
			if (typeof info.sheetName === 'string') {
				switchToWorksheet(tabContainer, info.sheetName, linkEl, info);
			}
		});
	} else {
		requestAnimationFrame(() => {
			switchToWorksheet(tabContainer, firstSheetName, linkEl, info);
		});
	}

    console.log("renderContainerEl", containerEl);

    return;
}


function createTableFromWorkbooks(wb: XLSX.WorkBook, info?: xlinfo, linkEl?: HTMLElement): HTMLDivElement {    
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
		(tabTitle as HTMLElement).style.marginTop = '8px'; // 设置上外边距
		(tabTitle as HTMLElement).style.marginBottom = '8px'; // 设置下外边距

        // 当点击页签标题时，切换到对应的页签
        tabTitle.addEventListener('click', (event) => {   		
			console.log('点击页签标题事件：', tabTitle);	

			// 阻止原来的事件传播和默认行为
			event.stopPropagation();
			event.preventDefault();

			switchToWorksheet(tabContainer, tabTitle.textContent ?? '', linkEl, info);
        });

        // 将页签标题添加到页签标题容器中
        tabTitleContainer.appendChild(tabTitle);
    }
	console.log('工作簿尺寸：', maxWidth, maxHeight);
	
	tabContainer.style.width = '100%';
	tabContainer.style.height = '100%';

    // 将页签标题容器添加到分页切换容器的顶部
	xlstyle.setTabTitleStyle(tabTitleContainer);
	tabContainer.insertBefore(tabTitleContainer, tabContainer.firstChild);

    // 默认显示第一个分页
	const firstSheetName = wb.SheetNames[0];
    if (tabContainer.firstChild) {
        switchToWorksheet(tabContainer, firstSheetName, linkEl, info);
    }

    return tabContainer;
}

// 获取当前激活的标签页
function getActiveTabEl(tabContainer: HTMLElement): HTMLElement | null {
	for (let child of Array.from(tabContainer.children)) {
	  if ((child as HTMLElement).style.display !== 'none') {
		return child as HTMLElement;
	  }
	}
	return null;
}

// 获取当前激活的标签页的名称
function getActiveTabName(tabContainer: HTMLElement): string | null {
	const activeTabEl = getActiveTabEl(tabContainer);
	return activeTabEl?.getAttribute('name') ?? null;
}

// 获取当前激活的标签页的实际宽度
function getActiveTabWidth(tabContainer: HTMLElement): number | null {
	const activeTabEl = getActiveTabEl(tabContainer);
	return activeTabEl?.scrollWidth ?? null;
}

// 切换到指定的工作表
function switchToWorksheet(tabContainer: HTMLDivElement, sheetName: string, linkEl?: HTMLElement, info?: xlinfo) {
	console.log('切换到工作表.tabContainer：', tabContainer);

	if (!tabContainer?.children) return;

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

				console.log('判断是否根据工作表宽度调整容器宽度：', linkEl, info, (linkEl && (!info?.size.width)));

				// 设置链接元素的宽度为链接元素当前宽度和工作表实际宽度的最小值
				if (linkEl && (!info?.size.width)) {	
					// 获取 child 下 class==C.WORKSHEET_EL_NAME 的元素
					const ws = (Array.from(child.children).find(c => c.classList.contains(C.WORKSHEET_EL_NAME)) as HTMLTableElement);			
					console.log('Worksheet 页面的宽度.switch.before：', linkEl.offsetWidth);
					console.log('Worksheet 表格的实际像素宽度.switch.before：', ws.offsetWidth);		
					requestAnimationFrame(() => {
						console.log('Worksheet 页面的宽度.switch.after：', linkEl.offsetWidth);
						console.log('Worksheet 表格的实际像素宽度.switch.after：', ws.offsetWidth);
						console.log('Worksheet 表格的最适合宽度.display：', linkEl.style.display, ws.style.display);
						xlstyle.setLinkElWidthByTable(ws, linkEl);
					});
				}
			} else {
				// 隐藏其他分页
				(child as HTMLElement).style.display = 'none';
			}
		} else {
			// 处理页签标题按钮
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

	// if (linkEl) {
	// 	// 设置链接元素居中显示
	// 	console.log('设置链接元素居中显示：', linkEl);
	// 	linkEl.style.marginLeft = 'auto';
	// 	linkEl.style.marginRight = 'auto';
	// }
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


// 获取上级元素中最近处的 ".internal-embed" 元素，没用，取不到上级元素
function getInternalEmbedEl(el: HTMLElement): HTMLElement | null {
	console.log('获取上级元素中最近处的 ".internal-embed" 元素.el：', el);
	let parent = el.parentElement;
	console.log('获取上级元素中最近处的 ".internal-embed" 元素.parent：', parent);
	while (parent) {
		console.log('获取上级元素中最近处的 ".internal-embed" 元素.parent：', parent);
		if (parent.classList.contains('internal-embed')) {
			return parent as HTMLElement;
		}
		parent = parent.parentElement;
	}
	return null;
}