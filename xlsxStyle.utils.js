//字符串转字符流
function s2ab(s) {
	var buf = new ArrayBuffer(s.length);
	var view = new Uint8Array(buf);
	for (var i = 0; i != s.length; ++i)
		view[i] = s.charCodeAt(i) & 0xFF;
	return buf;
}

//初始化
function init(workBook, sheetName, cell) {
	if (!workBook.Sheets[sheetName][cell].s) {
		workBook.Sheets[sheetName][cell].s = {};
	}
}

function init1(workBook, sheetName, cell, attr) {
	init(workBook, sheetName, cell);
	if (!workBook.Sheets[sheetName][cell].s[attr]) {
		workBook.Sheets[sheetName][cell].s[attr] = {};
	}
}

function init2(workBook, sheetName, cell, attr1, attr2) {
	init(workBook, sheetName, cell);
	init1(workBook, sheetName, cell, attr1);
	if (!workBook.Sheets[sheetName][cell].s[attr1][attr2]) {
		workBook.Sheets[sheetName][cell].s[attr1][attr2] = {};
	}
}

//单元格合并 startCell=A1 endCell=B5
function mergeCells(workBook, sheetName, startCell, endCell) {
	var sc = startCell.substr(0, 1).charCodeAt(0) - 65;
	var sr = startCell.substr(1);
	sr = parseInt(sr) - 1;
	var ec = endCell.substr(0, 1).charCodeAt(0) - 65;
	var er = endCell.substr(1)
	er = parseInt(er) - 1;

	var merges = [{
		s: { //s start 始单元格
			c: sc, //cols 开始列
			r: sr //rows 开始行
		},
		e: { //e end  末单元格
			c: ec, //cols 结束列
			r: er //rows 结束行
		}
	}];
	if (!workBook.Sheets[sheetName]["!merges"]) {
		workBook.Sheets[sheetName]["!merges"] = merges;
	} else {
		workBook.Sheets[sheetName]["!merges"] = workBook.Sheets[sheetName]["!merges"].concat(merges);
	}

	return workBook;

}

//merges=[{s: {c: 0, r: 0},e: {c: 3, r: 0}}]
function mergeCellsByObj(workBook, sheetName, merges) {

	if (workBook.Sheets[sheetName]["!merges"]) {
		workBook.Sheets[sheetName]["!merges"] = workBook.Sheets[sheetName]["!merges"].concat(merges);
	} else {
		workBook.Sheets[sheetName]["!merges"] = merges;
	}

	return workBook;

}


//设置每列列宽,单位px cols= [{wpx: 45}, {wpx: 165}, {wpx: 45}, {wpx: 45}]
function setColWidth(workBook, sheetName, cols) {
	workBook.Sheets[sheetName]["!cols"] = cols;
	return workBook;
}


//workBook.Sheets[sheetName][cell].s

//一次性设置多样式 styles etc： 
/*
{
  "font": {
    "sz": 14,
    "bold": true,
    "color": {
      "rgb": "FFFFAA00"
    }
  },
  "fill": {
    "bgColor": {
      "indexed": 64
    },
    "fgColor": {
      "rgb": "FFFFFF00"
    }
  }
}
*/
function setCellStyle(workBook, sheetName, cell, styles) {
	workBook.Sheets[sheetName][cell].s = styles;
	return workBook;
}

/*Fill*/

//填充颜色 fill={"bgColor": {"indexed": 64},"fgColor": {"rgb": "FFFFFF00"}}
function setFillStyles(workBook, sheetName, cell, styles) {
	init(workBook, sheetName, cell);
	workBook.Sheets[sheetName][cell].s.fill = styles;
	return workBook;
}

function setFillStylesAll(workBook, sheetName,styles) {
	for (cell in workBook.Sheets[sheetName]) {
		if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
			setFillStyles(workBook, sheetName, cell, styles);
		}
	}
}

//patternType="solid" or "none"”
function setFillPatternType(workBook, sheetName, cell, patternType) {
	init1(workBook, sheetName, cell, "fill");
	workBook.Sheets[sheetName][cell].s.fill.patternType = patternType;
	return workBook;
}

function setFillPatternTypeAll(workBook, sheetName,patternType) {
	for (cell in workBook.Sheets[sheetName]) {
		if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
			setFillPatternType(workBook, sheetName, cell, patternType);
		}
	}
}

//前景颜色(单元格颜色) rgb 
//COLOR_SPEC属性值:{ auto: 1}指定自动值,{ rgb: "FFFFAA00" }指定16进制的ARGB,{ theme: "1", tint: "-0.25"}指定主题颜色和色调的整数索引（默认值为0）,{ indexed: 64} 默认值 fill.bgColor
function setFillFgColor(workBook, sheetName, cell, COLOR_SPEC) {
	init1(workBook, sheetName, cell, "fill");
	workBook.Sheets[sheetName][cell].s.fill.fgColor = COLOR_SPEC;
	return workBook;
}

function setFillFgColorAll(workBook, sheetName,COLOR_SPEC) {
	for (cell in workBook.Sheets[sheetName]) {
		if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
			setFillFgColor(workBook, sheetName, cell, COLOR_SPEC);
		}
	}
}

//使用RGB值设置颜色
function setFillFgColorRGB(workBook, sheetName, cell, rgb) {
	init2(workBook, sheetName, cell, "fill", "fgColor");
	workBook.Sheets[sheetName][cell].s.fill.fgColor.rgb = rgb;
	return workBook;
}

function setFillFgColorRGBAll(workBook, sheetName,rgb) {
	for (cell in workBook.Sheets[sheetName]) {
		if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
			setFillFgColorRGB(workBook, sheetName, cell, rgb);
		}
	}
}

//单元格背景颜色（貌似没用）
function setFillBgColor(workBook, sheetName, cell, COLOR_SPEC) {
	init1(workBook, sheetName, cell, "fill");
	workBook.Sheets[sheetName][cell].s.fill.bgColor = COLOR_SPEC;
	return workBook;
}

function setFillBgColorAll(workBook, sheetName,COLOR_SPEC) {
	for (cell in workBook.Sheets[sheetName]) {
		if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
			setFillBgColor(workBook, sheetName, cell, COLOR_SPEC);
		}
	}
}

//单元格背景颜色
function setFillBgColorRGB(workBook, sheetName, cell, rgb) {
	init2(workBook, sheetName, cell, "fill", "bgColor");
	workBook.Sheets[sheetName][cell].s.fill.bgColor.rgb = rgb;
	return workBook;
}

function setFillBgColorRGBAll(workBook, sheetName,rgb) {
	for (cell in workBook.Sheets[sheetName]) {
		if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
			setFillBgColorRGB(workBook, sheetName, cell, rgb);
		}
	}
}

/*Font*/

//字体风格，可一次性在styles中设置所有font风格
function setFontStyles(workBook, sheetName, cell, styles) {
	init(workBook, sheetName, cell);
	workBook.Sheets[sheetName][cell].s.font = styles;
	return workBook;
}

function setFontStylesAll(workBook, sheetName,styles) {
	for (cell in workBook.Sheets[sheetName]) {
		if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
			setFontStyles(workBook, sheetName, cell, styles);
		}
	}
}

//字体 type="Calibri" 
function setFontType(workBook, sheetName, cell, type) {
	init1(workBook, sheetName, cell, "font");
	workBook.Sheets[sheetName][cell].s.font.name = type;
	return workBook;
}

function setFontTypeAll(workBook, sheetName,type) {
	for (cell in workBook.Sheets[sheetName]) {
		if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
			setFontType(workBook, sheetName, cell, type);
		}
	}
}

//字体大小
function setFontSize(workBook, sheetName, cell, size) {
	init1(workBook, sheetName, cell, "font");
	workBook.Sheets[sheetName][cell].s.font.sz = size;
	return workBook;
}

function setFontSizeAll(workBook, sheetName,size) {
	for (cell in workBook.Sheets[sheetName]) {
		if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
			setFontSize(workBook, sheetName, cell, size);
		}
	}
}

//字体颜色 COLOR_SPEC
function setFontColor(workBook, sheetName, cell, COLOR_SPEC) {
	init1(workBook, sheetName, cell, "font");
	workBook.Sheets[sheetName][cell].s.font.color = COLOR_SPEC;
	//workBook.Sheets[sheetName][cell].s.font.color.rgb = rgb;
	return workBook;
}

function setFontColorAll(workBook, sheetName,COLOR_SPEC) {
	for (cell in workBook.Sheets[sheetName]) {
		if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
			setFontColor(workBook, sheetName, cell, COLOR_SPEC);
		}
	}
}

//字体颜色RGB
function setFontColorRGB(workBook, sheetName, cell, rgb) {
	init2(workBook, sheetName, cell, "font", "color");
	workBook.Sheets[sheetName][cell].s.font.color.rgb = rgb;
	return workBook;
}

function setFontColorRGBAll(workBook, sheetName,rgb) {
	for (cell in workBook.Sheets[sheetName]) {
		if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
			setFontColorRGB(workBook, sheetName, cell, rgb);
		}
	}
}

//是否粗体 boolean isBold
function setFontBold(workBook, sheetName, cell, isBold) {
	init1(workBook, sheetName, cell, "font");
	workBook.Sheets[sheetName][cell].s.font.bold = isBold;
	return workBook;
}

function setFontBoldAll(workBook, sheetName,isBold) {
	for (cell in workBook.Sheets[sheetName]) {
		if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
			setFontBold(workBook, sheetName, cell, isBold);
		}
	}
}

//是否下划线 boolean isUnderline
function setFontUnderline(workBook, sheetName, cell, isUnderline) {
	init1(workBook, sheetName, cell, "font");
	workBook.Sheets[sheetName][cell].s.font.underline = isUnderline;
	return workBook;
}

function setFontUnderlineAll(workBook, sheetName,isUnderline) {
	for (cell in workBook.Sheets[sheetName]) {
		if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
			setFontUnderline(workBook, sheetName, cell, isUnderline);
		}
	}
}

//是否斜体 boolean isItalic
function setFontItalic(workBook, sheetName, cell, isItalic) {
	init1(workBook, sheetName, cell, "font");
	workBook.Sheets[sheetName][cell].s.font.italic = isItalic;
	return workBook;
}

function setFontItalicAll(workBook, sheetName,isItalic) {
	for (cell in workBook.Sheets[sheetName]) {
		if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
			setFontItalic(workBook, sheetName, cell, isItalic);
		}
	}
}

//是否删除线 boolean isStrike
function setFontStrike(workBook, sheetName, cell, isStrike) {
	init1(workBook, sheetName, cell, "font");
	workBook.Sheets[sheetName][cell].s.font.strike = isStrike;
	return workBook;
}

function setFontStrikeAll(workBook, sheetName,isStrike) {
	for (cell in workBook.Sheets[sheetName]) {
		if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
			setFontStrike(workBook, sheetName, cell, isStrike);
		}
	}
}

//是否outline boolean isOutline
function setFontOutline(workBook, sheetName, cell, isOutline) {
	init1(workBook, sheetName, cell, "font");
	workBook.Sheets[sheetName][cell].s.font.outline = isOutline;
	return workBook;
}

function setFontOutlineAll(workBook, sheetName,isOutline) {
	for (cell in workBook.Sheets[sheetName]) {
		if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
			setFontOutline(workBook, sheetName, cell, isOutline);
		}
	}
}

//是否阴影 boolean isShadow
function setFontShadow(workBook, sheetName, cell, isShadow) {
	init1(workBook, sheetName, cell, "font");
	workBook.Sheets[sheetName][cell].s.font.shadow = isShadow;
	return workBook;
}

function setFontShadowAll(workBook, sheetName,isShadow) {
	for (cell in workBook.Sheets[sheetName]) {
		if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
			setFontShadow(workBook, sheetName, cell, isShadow);
		}
	}
}

//是否vertAlign boolean isVertAlign
function setFontVertAlign(workBook, sheetName, cell, isVertAlign) {
	init1(workBook, sheetName, cell, "font");
	workBook.Sheets[sheetName][cell].s.font.vertAlign = isVertAlign;
	return workBook;
}

function setFontVertAlignAll(workBook, sheetName,isVertAlign) {
	for (cell in workBook.Sheets[sheetName]) {
		if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
			setFontVertAlign(workBook, sheetName, cell, isVertAlign);
		}
	}
}


/*numFmt*/

function setNumFmt(workBook, sheetName, cell, numFmt) {
	init(workBook, sheetName, cell);
	workBook.Sheets[sheetName][cell].s.numFmt = numFmt;
	return workBook;
}

function setNumFmtAll(workBook, sheetName,numFmt) {
	for (cell in workBook.Sheets[sheetName]) {
		if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
			setNumFmt(workBook, sheetName, cell, numFmt);
		}
	}
}

/*Alignment*/

//文本对齐 alignment={vertical:top,horizontal:top,}
function setAlignmentStyles(workBook, sheetName, cell, styles) {
	init(workBook, sheetName, cell);
	workBook.Sheets[sheetName][cell].s.alignment = styles;
	return workBook;
}

function setAlignmentStylesAll(workBook, sheetName,styles) {
	for (cell in workBook.Sheets[sheetName]) {
		if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
			setAlignmentStyles(workBook, sheetName, cell, styles);
		}
	}
}

//文本垂直对齐 vertical	="bottom" or "center" or "top"
function setAlignmentVertical(workBook, sheetName, cell, vertical) {
	init1(workBook, sheetName, cell, "alignment");
	workBook.Sheets[sheetName][cell].s.alignment.vertical = vertical;
	return workBook;
}

function setAlignmentVerticalAll(workBook, sheetName,vertical) {
	for (cell in workBook.Sheets[sheetName]) {
		if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
			setAlignmentVertical(workBook, sheetName, cell, vertical);
		}
	}
}

//文本水平对齐 "bottom" or "center" or "top"
function setAlignmentHorizontal(workBook, sheetName, cell, horizontal) {
	init1(workBook, sheetName, cell, "alignment");
	workBook.Sheets[sheetName][cell].s.alignment.horizontal = horizontal;
	return workBook;
}

function setAlignmentHorizontalAll(workBook, sheetName,horizontal) {
	for (cell in workBook.Sheets[sheetName]) {
		if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
			setAlignmentHorizontal(workBook, sheetName, cell, horizontal);
		}
	}
}

//自动换行
function setAlignmentWrapText(workBook, sheetName, cell, isWrapText) {
	init1(workBook, sheetName, cell, "alignment");
	workBook.Sheets[sheetName][cell].s.alignment.isWrapText = isWrapText;
	return workBook;
}

function setAlignmentWrapTextAll(workBook, sheetName,isWrapText) {
	for (cell in workBook.Sheets[sheetName]) {
		if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
			setAlignmentWrapText(workBook, sheetName, cell, isWrapText);
		}
	}
}

function setAlignmentReadingOrder(workBook, sheetName, cell, readingOrder) {
	init1(workBook, sheetName, cell, "alignment");
	workBook.Sheets[sheetName][cell].s.alignment.readingOrder = readingOrder;
	return workBook;
}

function setAlignmentReadingOrderAll(workBook, sheetName,readingOrder) {
	for (cell in workBook.Sheets[sheetName]) {
		if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
			setAlignmentReadingOrder(workBook, sheetName, cell, readingOrder);
		}
	}
}

//文本旋转角度 0-180，255 is special, aligned vertically
function setAlignmentTextRotation(workBook, sheetName, cell, textRotation) {
	init1(workBook, sheetName, cell, "alignment");
	workBook.Sheets[sheetName][cell].s.alignment.textRotation = textRotation;
	return workBook;
}

function setAlignmentTextRotationAll(workBook, sheetName,textRotation) {
	for (cell in workBook.Sheets[sheetName]) {
		if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
			setAlignmentTextRotation(workBook, sheetName, cell, textRotation);
		}
	}
}

/*Border*/

//单元格四周边框默认样式
const borderAll = {
	top: {
		style: 'thin'
	},
	bottom: {
		style: 'thin'
	},
	left: {
		style: 'thin'
	},
	right: {
		style: 'thin'
	}
};

//边框默样式：细线黑色
const defaultBorderStyle = {
	style: 'thin'
};

//边框 styles={top:{ style:"thin",color:"FFFFAA00"},bottom:{},...}
function setBorderStyles(workBook, sheetName, cell, styles) {
	init(workBook, sheetName, cell);
	workBook.Sheets[sheetName][cell].s.border = styles;
	return workBook;
}

function setBorderStylesAll(workBook, sheetName,styles) {
	for (cell in workBook.Sheets[sheetName]) {
		if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
			setBorderStyles(workBook, sheetName, cell, styles);
		}
	}
}

//设置单元格上下左右边框
function setBorderDefault(workBook, sheetName, cell) {
	init(workBook, sheetName, cell);
	workBook.Sheets[sheetName][cell].s.border = borderAll;
	return workBook;
}

//设置所有单元默认格边框
function setBorderDefaultAll(workBook, sheetName) {
	for (cell in workBook.Sheets[sheetName]) {
		if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
			setBorderDefault(workBook, sheetName, cell);
		}
	}
}

//上边框
function setBorderTop(workBook, sheetName, cell, top) {
	init(workBook, sheetName, cell, "border");
	workBook.Sheets[sheetName][cell].s.border.top = top;
	return workBook;
}

function setBorderTopAll(workBook, sheetName,top) {
	for (cell in workBook.Sheets[sheetName]) {
		if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
			setBorderTop(workBook, sheetName, cell, top);
		}
	}
}

//上边框默样式
function setBorderTopDefault(workBook, sheetName, cell) {
	init1(workBook, sheetName, cell, "border");
	workBook.Sheets[sheetName][cell].s.border.top = defaultBorderStyle;
	return workBook;
}

function setBorderTopDefaultAll(workBook, sheetName) {
	for (cell in workBook.Sheets[sheetName]) {
		if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
			setBorderTopDefault(workBook, sheetName, cell);
		}
	}
}

//下边框
function setBorderBottom(workBook, sheetName, cell, bottom) {
	init1(workBook, sheetName, cell, "border");
	workBook.Sheets[sheetName][cell].s.border.bottom = bottom;
	return workBook;
}

function setBorderBottomAll(workBook, sheetName,bottom) {
	for (cell in workBook.Sheets[sheetName]) {
		if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
			setBorderBottom(workBook, sheetName, cell, bottom);
		}
	}
}


//下边框默样式
function setBorderBottomDefault(workBook, sheetName, cell) {
	init1(workBook, sheetName, cell, "border");
	workBook.Sheets[sheetName][cell].s.border.bottom = defaultBorderStyle;
	return workBook;
}

function setBorderBottomDefaultAll(workBook, sheetName) {
	for (cell in workBook.Sheets[sheetName]) {
		if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
			setBorderBottomDefault(workBook, sheetName, cell);
		}
	}
}

//左边框
function setBorderLeft(workBook, sheetName, cell, left) {
	init1(workBook, sheetName, cell, "border");
	workBook.Sheets[sheetName][cell].s.border.left = left;
	return workBook;
}

function setBorderLeftAll(workBook, sheetName,left) {
	for (cell in workBook.Sheets[sheetName]) {
		if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
			setBorderLeft(workBook, sheetName, cell, left);
		}
	}
}

function setBorderLeftDefault(workBook, sheetName, cell) {
	init1(workBook, sheetName, cell, "border");
	workBook.Sheets[sheetName][cell].s.border.left = defaultBorderStyle;
	return workBook;
}

function setBorderLeftDefaultAll(workBook, sheetName) {
	for (cell in workBook.Sheets[sheetName]) {
		if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
			setBorderLeftDefault(workBook, sheetName, cell);
		}
	}
}

//右边框
function setBorderRight(workBook, sheetName, cell, right) {
	init1(workBook, sheetName, cell, "border");
	workBook.Sheets[sheetName][cell].s.border.right = right;
	return workBook;
}

function setBorderRightAll(workBook, sheetName,right) {
	for (cell in workBook.Sheets[sheetName]) {
		if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
			setBorderRight(workBook, sheetName, cell, right);
		}
	}
}

function setBorderRightDefault(workBook, sheetName, cell) {
	init1(workBook, sheetName, cell, "border");
	workBook.Sheets[sheetName][cell].s.border.right = defaultBorderStyle;
	return workBook;
}

function setBorderRightDefaultAll(workBook, sheetName) {
	for (cell in workBook.Sheets[sheetName]) {
		if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
			setBorderRightDefault(workBook, sheetName, cell);
		}
	}
}

//对角线
function setBorderDiagonal(workBook, sheetName, cell, diagonal) {
	init1(workBook, sheetName, cell, "border");
	workBook.Sheets[sheetName][cell].s.border.diagonal = diagonal;
	return workBook;
}

function setBorderDiagonalAll(workBook, sheetName,diagonal) {
	for (cell in workBook.Sheets[sheetName]) {
		if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
			setBorderDiagonal(workBook, sheetName, cell, diagonal);
		}
	}
}

function setBorderDiagonalDefault(workBook, sheetName, cell) {
	init1(workBook, sheetName, cell, "border");
	workBook.Sheets[sheetName][cell].s.border.diagonal = defaultBorderStyle;
	return workBook;
}

function setBorderDiagonalDefaultAll(workBook, sheetName) {
	for (cell in workBook.Sheets[sheetName]) {
		if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
			setBorderDiagonalDefault(workBook, sheetName, cell);
		}
	}
}

function setBorderDiagonalUp(workBook, sheetName, cell, isDiagonalUp) {
	init1(workBook, sheetName, cell, "border");
	workBook.Sheets[sheetName][cell].s.border.diagonalUp = isDiagonalUp;
	return workBook;
}

function setBorderDiagonalUpAll(workBook, sheetName,isDiagonalUp) {
	for (cell in workBook.Sheets[sheetName]) {
		if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
			setBorderDiagonalUp(workBook, sheetName, cell, isDiagonalUp);
		}
	}
}

function setBorderDiagonalDown(workBook, sheetName, cell, isDiagonalDown) {
	init1(workBook, sheetName, cell, "border");
	workBook.Sheets[sheetName][cell].s.border.diagonalDown = isDiagonalDown;
	return workBook;
}

function setBorderDiagonalDownAll(workBook, sheetName,isDiagonalDown) {
	for (cell in workBook.Sheets[sheetName]) {
		if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
			setBorderDiagonalDown(workBook, sheetName, cell, isDiagonalDown);
		}
	}
}

//默认样式，多单元格设置样式

//设置所有单元格字体样式
function setFgColorStylesAll(workBook, sheetName,fontType,fontColor,fontSize) {
	for (cell in workBook.Sheets[sheetName]) {
		if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
			setFontType(workBook, sheetName, cell,fontType);
			setFontColorRGB(workBook, sheetName, cell,fontColor);
			setFontSize(workBook, sheetName, cell,fontSize);
		}
	}
}

//设置第一行标题自定义样式
function setTitleStyles(workBook, sheetName, fgColor, fontColor, alignment, isBold, fontSize, ) {
	for (cell in workBook.Sheets[sheetName]) {
		if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
			row = cell.substr(1);
			if (row == '1') {
				setFillFgColorRGB(workBook, sheetName, cell, fgColor);
				setFontColor(workBook, sheetName, cell, fontColor);
				setAlignmentHorizontal(workBook, sheetName, cell, alignment);
				setFontBold(workBook, sheetName, cell, isBold);
				setFontSize(workBook, sheetName, cell, fontSize);
			}

		}
	}
}

//设置第一行标题默认样式
function setTitleStylesDefault(workBook, sheetName) {
	for (cell in workBook.Sheets[sheetName]) {
		row = cell.substr(1);
		if (row == '1') {
			setFillFgColorRGB(workBook, sheetName, cell, 'FFFF00');
			setAlignmentHorizontal(workBook, sheetName, cell, 'center');
			setFontBold(workBook, sheetName, cell, true);
			setFontSize(workBook, sheetName, cell, '20');
		}

	}
}

//设置双数行背景色灰色，便于阅读
function setEvenRowColorGrey(workBook, sheetName) {
	for (cell in workBook.Sheets[sheetName]) {
		if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
			row = parseInt(cell.substr(1));
			if (row % 2 == 0) {
				setFillFgColorRGB(workBook, sheetName, cell, 'DCDCDC');
			}

		}
	}
}