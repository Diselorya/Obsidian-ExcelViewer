import xlinfo from "./excelembedinfo";
import * as C from "../constants";

// 设置表格样式
export function setTableStyle(
    tableEl: HTMLTableElement,  // 要设置样式的表格
	info?: xlinfo,
    maxCellWidth: string = '',     // 每列的最大宽度
    mimCellWidth: string = '50px',     // 每列的最小宽度  
) {
    // 添加新的样式
    tableEl.style.cssText = `
        width: 100%;
        border-collapse: collapse;
        border: 1px solid var(--background-modifier-border-focus);
        padding: 8px;
        text-align: left;
        color: var(--text-normal);
    `;

    // 设置行的样式
    const rows = tableEl.querySelectorAll('tr');
    rows.forEach((row, index) => {
        if (index % 2 === 0) {
            row.style.backgroundColor = 'var(--color-base-25)';
        }
        let color: string;
        row.addEventListener('mouseover', () => {
            color = row.style.backgroundColor;
            row.style.backgroundColor = 'var(--color-base-30)';
        });
        row.addEventListener('mouseout', () => {
            row.style.backgroundColor = index % 2 === 0 ? color : '';
        });
    });

    // 设置单元格的样式
    const cells = tableEl.querySelectorAll('th, td');
    cells.forEach(cell => {
        (cell as HTMLElement).style.border = '1px solid var(--background-modifier-border-focus)';
        (cell as HTMLElement).style.padding = '8px';
        (cell as HTMLElement).style.textAlign = 'left';
    });


	// 找到表格首行
	const firstRow = tableEl.querySelector('tr');
	// 如果找到了表格首行，就设置其背景色
	if (firstRow) {
		(firstRow as HTMLElement).style.backgroundColor = 'var(--table-header-background)';
        (firstRow as HTMLElement).style.color = 'var(--table-header-color)';
	}

	// 优化点：冻结首行首列
	// 优化点：工具栏：可选对齐方式等
	// 优化点：图片提取
	// 优化点：点击放大全屏查看表格并自动调整尺寸
    // 优化点：改了链接的尺寸，但是渲染没有刷新，要切换一下页面才行
    // 优化点：应用/取消应用 Excel 的单元格填充色、字体、颜色

    // 设置表格中每个单元格的最小、最大宽度
    if (mimCellWidth) tableEl.style.cssText += 'td { min-width: ' + mimCellWidth + 'px; }';
    if (maxCellWidth) tableEl.style.cssText += 'td { max-width: ' + maxCellWidth + 'px; }';

    // 设置表格的宽度为最适合的宽度，而不是填满容器
    tableEl.style.width = 'max-content';

	console.log('工作表样式：', tableEl.style.height, '|', tableEl.style.width);
}


// 设置容器样式
export function setContainerStyle(
    containerEl: HTMLDivElement,  // 要设置样式的表格
	info?: xlinfo,
    width?: string,     // 宽度  
    height?: string,     // 高度
) {
	console.log('设定的容器尺寸：', info?.size, width, height);

    containerEl.style.width = 'max-content';
    const maxFitWidth = containerEl.offsetWidth;
    containerEl.style.width = 'min-content';
    const minFitWidth = containerEl.offsetWidth;
    console.log('容器的最适合宽度：', maxFitWidth, '~', minFitWidth);

    containerEl.style.height = 'max-content';
    const maxFitHeight = containerEl.offsetHeight;
    containerEl.style.height = 'min-content';
    const minFitHeight = containerEl.offsetHeight;
    console.log('容器的最适合高度：', maxFitHeight, '~', minFitHeight);

	if (info?.size)
	{
        if (width) containerEl.style.width = width;
        else {
            const w = info.size.width;
            if (w === 0) containerEl.style.width = '100%';
            else containerEl.style.width = `${w}px`;
        }

        if (height) containerEl.style.height = height;
        else {
            const h = info.size.height;
            if (h === 0) containerEl.style.height = '100%';
            else containerEl.style.height = `${h}px`;
        }
	}

    // 限制宽高不得超过容器的最适合宽度
    if (containerEl.offsetWidth > maxFitWidth) containerEl.style.width = `${maxFitWidth}px`;
    // if (containerEl.offsetWidth < minFitWidth) containerEl.style.width = `${minFitWidth}px`;
    
    // 滚动条必须设置在容器的父节点，否则会出现滚动条和渲染的宽度不一致的问题
}

// 设置工作表切换页签的样式
export function setTabTitleStyle(tabTitle: HTMLDivElement) {
	tabTitle.style.textAlign = 'left';
    const buttons = tabTitle.querySelectorAll('button');
    buttons.forEach(button => {
        // button.style.clipPath = 'polygon(10% 0%, 90% 0%, 100% 100%, 0% 100%)';
        button.style.borderRadius = '5px';
    });
}

// 通过表格元素设置嵌入链接区域的宽度
export function setLinkElWidthByTable(tabEl: HTMLTableElement, linkEl: HTMLElement) {
	// 重置 linkEl 的宽度为 auto，否则会导致一次比一次小
    console.log('Worksheet 表格的最适合宽度：', tabEl);
	linkEl.style.width = 'auto'
    tabEl.style.width = 'max-content';
    console.log('Worksheet 页面的宽度.before：', linkEl.offsetWidth);
    console.log('Worksheet 表格的实际像素宽度.before：', tabEl.offsetWidth);

	// requestAnimationFrame 是一个浏览器提供的 API
	// 它接受一个回调函数作为参数，并在下一次浏览器重绘之前调用这个回调函数。
	// 这意味着当你的回调函数被调用时，所有的 DOM 更新（包括元素的渲染）都已经完成。
	// 如果直接写逻辑，会导致获取到的宽度不正确，因为此时 DOM 还没有渲染完成
	// requestAnimationFrame(() => {
		const pageWidth = linkEl.offsetWidth;
		const wsWidth = tabEl.offsetWidth + 16;
        console.log('Worksheet 页面的宽度.after', pageWidth);
		console.log('Worksheet 表格的实际像素宽度.after', wsWidth);
        console.log('Worksheet 表格的最适合宽度.display：', linkEl.style.display, tabEl.style.display);
		setLinkElWidth(linkEl, wsWidth, pageWidth);
	// });

	// requestAnimationFrame(() => {
		// 如果调整宽度后 linkEl 的宽度仍然大于表格实际宽度，则 linkEl 的类子元素均居中显示
		if (linkEl.offsetWidth > tabEl.offsetWidth) {
			const wss = linkEl.querySelectorAll(`.${C.WORKSHEET_EL_NAME}`);
			console.log('设置子元素居中显示：', wss);
			for (let e of Array.from(wss)) {
				console.log('设置子元素居中显示：', e);
				(e as HTMLElement).style.marginLeft = 'auto';
				(e as HTMLElement).style.marginRight = 'auto';
			}

			const els = linkEl.querySelectorAll(`.worksheet-names`);
			console.log('设置子元素居中显示：', els);
			for (let e of Array.from(els)) {
				console.log('设置子元素居中显示：', e);
				(e as HTMLElement).style.display = 'flex';
				(e as HTMLElement).style.justifyContent = 'center';
				(e as HTMLElement).style.alignItems = 'center';
			}
		} else {
			const wss = linkEl.querySelectorAll(`.${C.WORKSHEET_EL_NAME}`);
			console.log('设置子元素居中显示：', wss);
			for (let e of Array.from(wss)) {
				console.log('设置子元素居中显示：', e);
				(e as HTMLElement).style.marginLeft = '';
				(e as HTMLElement).style.marginRight = '';
			}

			const els = linkEl.querySelectorAll(`.worksheet-names`);
			console.log('设置子元素居中显示：', els);
			for (let e of Array.from(els)) {
				console.log('设置子元素居中显示：', e);
				(e as HTMLElement).style.display = '';
				(e as HTMLElement).style.justifyContent = '';
				(e as HTMLElement).style.alignItems = '';
			}
		}
	// });
}

// 设置嵌入链接区域的宽度
export function setLinkElWidth(linkEl: HTMLElement, width: number, pageWidth: number) {
	// 设置链接元素的宽度为链接元素当前宽度和工作表实际宽度的最小值
	let w = Math.min(linkEl.offsetWidth, width);
	// 连接元素宽度不能小于 300px 或者 30% 的页面宽度
	w = Math.max(w, Math.max(pageWidth * 0.2, 200));
	// 设置链接元素的宽度
	// linkEl.style.width = w + 'px';

	console.log('设置链接元素的宽度：', w);
}