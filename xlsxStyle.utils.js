/*
@author Ctrl
@version 20190917
@aim 对xlsx-style方法进行二次封装 方便调用以导出带样式Excel
@usage XSU.xxxFunction()
*/


var XSU;

XSU=({

	//字符串转字符流
	s2ab: function(s) {
		var buf = new ArrayBuffer(s.length);
		var view = new Uint8Array(buf);
		for (var i = 0; i != s.length; ++i)
			view[i] = s.charCodeAt(i) & 0xFF;
		return buf;
	},

	//初始化
	init: function(workBook, sheetName, cell) {
		if (!workBook.Sheets[sheetName][cell].s) {
			workBook.Sheets[sheetName][cell].s = {};
		}
	},

	init1: function(workBook, sheetName, cell, attr) {
		this.init(workBook, sheetName, cell);
		if (!workBook.Sheets[sheetName][cell].s[attr]) {
			workBook.Sheets[sheetName][cell].s[attr] = {};
		}
	},

	init2: function(workBook, sheetName, cell, attr1, attr2) {
		this.init(workBook, sheetName, cell);
		this.init1(workBook, sheetName, cell, attr1);
		if (!workBook.Sheets[sheetName][cell].s[attr1][attr2]) {
			workBook.Sheets[sheetName][cell].s[attr1][attr2] = {};
		}
	},

	//根据ref的单元格范围新建范围内所有单元格,不存在的单元格置为空值,已存在的不处理
	initAllCell: function(workBook, sheetName) {
		var ref = workBook.Sheets[sheetName]["!ref"].split(":");
		var startCell = ref[0];
		var endCell = ref[1];
		var sc = XLSX.utils.decode_cell(startCell).c;
		var sr = XLSX.utils.decode_cell(startCell).r;
		var ec = XLSX.utils.decode_cell(endCell).c;
		var er = XLSX.utils.decode_cell(endCell).r;
		var isExist;
		for (c = sc; c <= ec; c++) { //初始化所有单元格
			for (r = sr; r <= er; r++) {
				var temp = XLSX.utils.encode_cell({
					c: c,
					r: r
				});
				isExist = false;
				for (cell in workBook.Sheets[sheetName]) {
					if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
						if (temp == cell) {
							isExist = true;
							break;
						}
					}
				}
				if (!isExist) { //单元格不存在则新建单元格
					XLSX.utils.sheet_add_aoa(workBook.Sheets[sheetName], [
						['']
					], {
						origin: temp
					});
				}

			}
		}
	},

	//单元格合并 startCell=A1 endCell=B5
	mergeCells: function(workBook, sheetName, startCell, endCell) {
		/*var sc = startCell.substr(0, 1).charCodeAt(0) - 65;
		var sr = startCell.substr(1);
		sr = parseInt(sr) - 1;
		var ec = endCell.substr(0, 1).charCodeAt(0) - 65;
		var er = endCell.substr(1)
		er = parseInt(er) - 1;*/

		var sc = XLSX.utils.decode_cell(startCell).c;
		var sr = XLSX.utils.decode_cell(startCell).r;
		var ec = XLSX.utils.decode_cell(endCell).c;
		var er = XLSX.utils.decode_cell(endCell).r;

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

	},

	//merges=[{s: {c: 0, r: 0},e: {c: 3, r: 0}}]
	mergeCellsByObj: function(workBook, sheetName, merges) {
		if (workBook.Sheets[sheetName]["!merges"]) {
			workBook.Sheets[sheetName]["!merges"] = workBook.Sheets[sheetName]["!merges"].concat(merges);
		} else {
			workBook.Sheets[sheetName]["!merges"] = merges;
		}

		return workBook;

	},


	//设置每列列宽,单位px cols= [{wpx: 45}, {wpx: 165}, {wpx: 45}, {wpx: 45}]
	setColWidth: function(workBook, sheetName, cols) {
		workBook.Sheets[sheetName]["!cols"] = cols;
		return workBook;
	},

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
	setCellStyle: function(workBook, sheetName, cell, styles) {
		workBook.Sheets[sheetName][cell].s = styles;
		return workBook;
	},

	/*Fill*/

	//填充颜色 fill={"bgColor": {"indexed": 64},"fgColor": {"rgb": "FFFFFF00"}}
	setFillStyles: function(workBook, sheetName, cell, styles) {
		this.init(workBook, sheetName, cell);
		workBook.Sheets[sheetName][cell].s.fill = styles;
		return workBook;
	},

	setFillStylesAll: function(workBook, sheetName, styles) {
		for (cell in workBook.Sheets[sheetName]) {
			if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
				this.setFillStyles(workBook, sheetName, cell, styles);
			}
		}
	},

	//patternType="solid" or "none"”
	setFillPatternType: function(workBook, sheetName, cell, patternType) {
		this.init1(workBook, sheetName, cell, "fill");
		workBook.Sheets[sheetName][cell].s.fill.patternType = patternType;
		return workBook;
	},

	setFillPatternTypeAll: function(workBook, sheetName, patternType) {
		for (cell in workBook.Sheets[sheetName]) {
			if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
				this.setFillPatternType(workBook, sheetName, cell, patternType);
			}
		}
	},

	//前景颜色(单元格颜色) rgb 
	//COLOR_SPEC属性值:{ auto: 1}指定自动值,{ rgb: "FFFFAA00" }指定16进制的ARGB,{ theme: "1", tint: "-0.25"}指定主题颜色和色调的整数索引（默认值为0）,{ indexed: 64} 默认值 fill.bgColor
	setFillFgColor: function(workBook, sheetName, cell, COLOR_SPEC) {
		this.init1(workBook, sheetName, cell, "fill");
		workBook.Sheets[sheetName][cell].s.fill.fgColor = COLOR_SPEC;
		return workBook;
	},

	setFillFgColorAll: function(workBook, sheetName, COLOR_SPEC) {
		for (cell in workBook.Sheets[sheetName]) {
			if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
				this.setFillFgColor(workBook, sheetName, cell, COLOR_SPEC);
			}
		}
	},

	//使用RGB值设置颜色
	setFillFgColorRGB: function(workBook, sheetName, cell, rgb) {
		this.init2(workBook, sheetName, cell, "fill", "fgColor");
		workBook.Sheets[sheetName][cell].s.fill.fgColor.rgb = rgb;
		return workBook;
	},

	setFillFgColorRGBAll: function(workBook, sheetName, rgb) {
		for (cell in workBook.Sheets[sheetName]) {
			if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
				this.setFillFgColorRGB(workBook, sheetName, cell, rgb);
			}
		}
	},

	//单元格背景颜色（貌似没用）
	setFillBgColor: function(workBook, sheetName, cell, COLOR_SPEC) {
		this.init1(workBook, sheetName, cell, "fill");
		workBook.Sheets[sheetName][cell].s.fill.bgColor = COLOR_SPEC;
		return workBook;
	},

	setFillBgColorAll: function(workBook, sheetName, COLOR_SPEC) {
		for (cell in workBook.Sheets[sheetName]) {
			if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
				this.setFillBgColor(workBook, sheetName, cell, COLOR_SPEC);
			}
		}
	},

	//单元格背景颜色
	setFillBgColorRGB: function(workBook, sheetName, cell, rgb) {
		this.init2(workBook, sheetName, cell, "fill", "bgColor");
		workBook.Sheets[sheetName][cell].s.fill.bgColor.rgb = rgb;
		return workBook;
	},

	setFillBgColorRGBAll: function(workBook, sheetName, rgb) {
		for (cell in workBook.Sheets[sheetName]) {
			if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
				this.setFillBgColorRGB(workBook, sheetName, cell, rgb);
			}
		}
	},

	/*Font*/

	//字体风格，可一次性在styles中设置所有font风格
	setFontStyles: function(workBook, sheetName, cell, styles) {
		this.init(workBook, sheetName, cell);
		workBook.Sheets[sheetName][cell].s.font = styles;
		return workBook;
	},

	setFontStylesAll: function(workBook, sheetName, styles) {
		for (cell in workBook.Sheets[sheetName]) {
			if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
				this.setFontStyles(workBook, sheetName, cell, styles);
			}
		}
	},

	//字体 type="Calibri" 
	setFontType: function(workBook, sheetName, cell, type) {
		this.init1(workBook, sheetName, cell, "font");
		workBook.Sheets[sheetName][cell].s.font.name = type;
		return workBook;
	},

	setFontTypeAll: function(workBook, sheetName, type) {
		for (cell in workBook.Sheets[sheetName]) {
			if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
				this.setFontType(workBook, sheetName, cell, type);
			}
		}
	},

	//字体大小
	setFontSize: function(workBook, sheetName, cell, size) {
		this.init1(workBook, sheetName, cell, "font");
		workBook.Sheets[sheetName][cell].s.font.sz = size;
		return workBook;
	},

	setFontSizeAll: function(workBook, sheetName, size) {
		for (cell in workBook.Sheets[sheetName]) {
			if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
				this.setFontSize(workBook, sheetName, cell, size);
			}
		}
	},

	//字体颜色 COLOR_SPEC
	setFontColor: function(workBook, sheetName, cell, COLOR_SPEC) {
		this.init1(workBook, sheetName, cell, "font");
		workBook.Sheets[sheetName][cell].s.font.color = COLOR_SPEC;
		//workBook.Sheets[sheetName][cell].s.font.color.rgb = rgb;
		return workBook;
	},

	setFontColorAll: function(workBook, sheetName, COLOR_SPEC) {
		for (cell in workBook.Sheets[sheetName]) {
			if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
				this.setFontColor(workBook, sheetName, cell, COLOR_SPEC);
			}
		}
	},

	//字体颜色RGB
	setFontColorRGB: function(workBook, sheetName, cell, rgb) {
		this.init2(workBook, sheetName, cell, "font", "color");
		workBook.Sheets[sheetName][cell].s.font.color.rgb = rgb;
		return workBook;
	},

	setFontColorRGBAll: function(workBook, sheetName, rgb) {
		for (cell in workBook.Sheets[sheetName]) {
			if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
				this.setFontColorRGB(workBook, sheetName, cell, rgb);
			}
		}
	},

	//是否粗体 boolean isBold
	setFontBold: function(workBook, sheetName, cell, isBold) {
		this.init1(workBook, sheetName, cell, "font");
		workBook.Sheets[sheetName][cell].s.font.bold = isBold;
		return workBook;
	},

	setFontBoldAll: function(workBook, sheetName, isBold) {
		for (cell in workBook.Sheets[sheetName]) {
			if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
				this.setFontBold(workBook, sheetName, cell, isBold);
			}
		}
	},

	//设置某列为粗体
	setFontBoldOfCols: function(workBook, sheetName, isBold, col) {
		for (cell in workBook.Sheets[sheetName]) {
			if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
				if (cell.substr(0, 1) == col) {
					this.setFontBold(workBook, sheetName, cell, isBold);
				}
			}
		}
	},

	//设置某行为粗体
	setFontBoldOfRows: function(workBook, sheetName, isBold, row) {
		for (cell in workBook.Sheets[sheetName]) {
			if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
				if (cell.substr(1) == row) {
					this.setFontBold(workBook, sheetName, cell, isBold);
				}
			}
		}
	},

	//是否下划线 boolean isUnderline
	setFontUnderline: function(workBook, sheetName, cell, isUnderline) {
		this.init1(workBook, sheetName, cell, "font");
		workBook.Sheets[sheetName][cell].s.font.underline = isUnderline;
		return workBook;
	},

	setFontUnderlineAll: function(workBook, sheetName, isUnderline) {
		for (cell in workBook.Sheets[sheetName]) {
			if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
				this.setFontUnderline(workBook, sheetName, cell, isUnderline);
			}
		}
	},

	//是否斜体 boolean isItalic
	setFontItalic: function(workBook, sheetName, cell, isItalic) {
		this.init1(workBook, sheetName, cell, "font");
		workBook.Sheets[sheetName][cell].s.font.italic = isItalic;
		return workBook;
	},

	setFontItalicAll: function(workBook, sheetName, isItalic) {
		for (cell in workBook.Sheets[sheetName]) {
			if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
				this.setFontItalic(workBook, sheetName, cell, isItalic);
			}
		}
	},

	//是否删除线 boolean isStrike
	setFontStrike: function(workBook, sheetName, cell, isStrike) {
		this.init1(workBook, sheetName, cell, "font");
		workBook.Sheets[sheetName][cell].s.font.strike = isStrike;
		return workBook;
	},

	setFontStrikeAll: function(workBook, sheetName, isStrike) {
		for (cell in workBook.Sheets[sheetName]) {
			if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
				this.setFontStrike(workBook, sheetName, cell, isStrike);
			}
		}
	},

	//是否outline boolean isOutline
	setFontOutline: function(workBook, sheetName, cell, isOutline) {
		this.init1(workBook, sheetName, cell, "font");
		workBook.Sheets[sheetName][cell].s.font.outline = isOutline;
		return workBook;
	},

	setFontOutlineAll: function(workBook, sheetName, isOutline) {
		for (cell in workBook.Sheets[sheetName]) {
			if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
				this.setFontOutline(workBook, sheetName, cell, isOutline);
			}
		}
	},

	//是否阴影 boolean isShadow
	setFontShadow: function(workBook, sheetName, cell, isShadow) {
		this.init1(workBook, sheetName, cell, "font");
		workBook.Sheets[sheetName][cell].s.font.shadow = isShadow;
		return workBook;
	},

	setFontShadowAll: function(workBook, sheetName, isShadow) {
		for (cell in workBook.Sheets[sheetName]) {
			if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
				this.setFontShadow(workBook, sheetName, cell, isShadow);
			}
		}
	},

	//是否vertAlign boolean isVertAlign
	setFontVertAlign: function(workBook, sheetName, cell, isVertAlign) {
		this.init1(workBook, sheetName, cell, "font");
		workBook.Sheets[sheetName][cell].s.font.vertAlign = isVertAlign;
		return workBook;
	},

	setFontVertAlignAll: function(workBook, sheetName, isVertAlign) {
		for (cell in workBook.Sheets[sheetName]) {
			if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
				this.setFontVertAlign(workBook, sheetName, cell, isVertAlign);
			}
		}
	},

	/*numFmt*/

	setNumFmt: function(workBook, sheetName, cell, numFmt) {
		this.init(workBook, sheetName, cell);
		workBook.Sheets[sheetName][cell].s.numFmt = numFmt;
		return workBook;
	},

	setNumFmtAll: function(workBook, sheetName, numFmt) {
		for (cell in workBook.Sheets[sheetName]) {
			if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
				this.setNumFmt(workBook, sheetName, cell, numFmt);
			}
		}
	},

	/*Alignment*/

	//文本对齐 alignment={vertical:top,horizontal:top,}
	setAlignmentStyles: function(workBook, sheetName, cell, styles) {
		this.init(workBook, sheetName, cell);
		workBook.Sheets[sheetName][cell].s.alignment = styles;
		return workBook;
	},

	setAlignmentStylesAll: function(workBook, sheetName, styles) {
		for (cell in workBook.Sheets[sheetName]) {
			if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
				this.setAlignmentStyles(workBook, sheetName, cell, styles);
			}
		}
	},

	//文本垂直对齐 vertical	="bottom" or "center" or "top"
	setAlignmentVertical: function(workBook, sheetName, cell, vertical) {
		this.init1(workBook, sheetName, cell, "alignment");
		workBook.Sheets[sheetName][cell].s.alignment.vertical = vertical;
		return workBook;
	},

	setAlignmentVerticalAll: function(workBook, sheetName, vertical) {
		for (cell in workBook.Sheets[sheetName]) {
			if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
				this.setAlignmentVertical(workBook, sheetName, cell, vertical);
			}
		}
	},

	//文本水平对齐 "bottom" or "center" or "top"
	setAlignmentHorizontal: function(workBook, sheetName, cell, horizontal) {
		this.init1(workBook, sheetName, cell, "alignment");
		workBook.Sheets[sheetName][cell].s.alignment.horizontal = horizontal;
		return workBook;
	},

	setAlignmentHorizontalAll: function(workBook, sheetName, horizontal) {
		for (cell in workBook.Sheets[sheetName]) {
			if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
				this.setAlignmentHorizontal(workBook, sheetName, cell, horizontal);
			}
		}
	},

	//自动换行
	setAlignmentWrapText: function(workBook, sheetName, cell, isWrapText) {
		this.init1(workBook, sheetName, cell, "alignment");
		workBook.Sheets[sheetName][cell].s.alignment.wrapText = isWrapText;
		return workBook;
	},

	setAlignmentWrapTextAll: function(workBook, sheetName, isWrapText) {
		for (cell in workBook.Sheets[sheetName]) {
			if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
				this.setAlignmentWrapText(workBook, sheetName, cell, isWrapText);
			}
		}
	},

	setAlignmentReadingOrder: function(workBook, sheetName, cell, readingOrder) {
		this.init1(workBook, sheetName, cell, "alignment");
		workBook.Sheets[sheetName][cell].s.alignment.readingOrder = readingOrder;
		return workBook;
	},

	setAlignmentReadingOrderAll: function(workBook, sheetName, readingOrder) {
		for (cell in workBook.Sheets[sheetName]) {
			if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
				this.setAlignmentReadingOrder(workBook, sheetName, cell, readingOrder);
			}
		}
	},

	//文本旋转角度 0-180，255 is special, aligned vertically
	setAlignmentTextRotation: function(workBook, sheetName, cell, textRotation) {
		this.init1(workBook, sheetName, cell, "alignment");
		workBook.Sheets[sheetName][cell].s.alignment.textRotation = textRotation;
		return workBook;
	},

	setAlignmentTextRotationAll: function(workBook, sheetName, textRotation) {
		for (cell in workBook.Sheets[sheetName]) {
			if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
				this.setAlignmentTextRotation(workBook, sheetName, cell, textRotation);
			}
		}
	},

	/*Border*/

	//单元格四周边框默认样式
	borderAll :{
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
	},

	//边框默样式：细线黑色
	defaultBorderStyle : {
		style: 'thin'
	},

	//边框 styles={top:{ style:"thin",color:"FFFFAA00"},bottom:{},...}
	setBorderStyles: function(workBook, sheetName, cell, styles) {
		this.init(workBook, sheetName, cell);
		workBook.Sheets[sheetName][cell].s.border = styles;
		return workBook;
	},

	setBorderStylesAll: function(workBook, sheetName, styles) {
		this.initAllCell(workBook, sheetName);
		for (cell in workBook.Sheets[sheetName]) {
			if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
				this.setBorderStyles(workBook, sheetName, cell, styles);
			}
		}
	},

	//设置单元格上下左右边框默认样式
	setBorderDefault: function(workBook, sheetName, cell) {
		this.init(workBook, sheetName, cell);
		workBook.Sheets[sheetName][cell].s.border = this.borderAll;
		return workBook;
	},

	//设置所有单元默认格边框
	setBorderDefaultAll: function(workBook, sheetName) {
		this.initAllCell(workBook, sheetName);
		for (cell in workBook.Sheets[sheetName]) {
			if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
				this.setBorderDefault(workBook, sheetName, cell);
			}
		}
	},

	//上边框
	setBorderTop: function(workBook, sheetName, cell, top) {
		this.init(workBook, sheetName, cell, "border");
		workBook.Sheets[sheetName][cell].s.border.top = top;
		return workBook;
	},

	setBorderTopAll: function(workBook, sheetName, top) {
		this.initAllCell(workBook, sheetName);
		for (cell in workBook.Sheets[sheetName]) {
			if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
				this.setBorderTop(workBook, sheetName, cell, top);
			}
		}
	},

	//上边框默样式
	setBorderTopDefault: function(workBook, sheetName, cell) {
		this.init1(workBook, sheetName, cell, "border");
		workBook.Sheets[sheetName][cell].s.border.top = this.defaultBorderStyle;
		return workBook;
	},

	setBorderTopDefaultAll: function(workBook, sheetName) {
		this.initAllCell(workBook, sheetName);
		for (cell in workBook.Sheets[sheetName]) {
			if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
				this.setBorderTopDefault(workBook, sheetName, cell);
			}
		}
	},

	//下边框
	setBorderBottom: function(workBook, sheetName, cell, bottom) {
		this.init1(workBook, sheetName, cell, "border");
		workBook.Sheets[sheetName][cell].s.border.bottom = bottom;
		return workBook;
	},

	setBorderBottomAll: function(workBook, sheetName, bottom) {
		this.initAllCell(workBook, sheetName);
		for (cell in workBook.Sheets[sheetName]) {
			if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
				this.setBorderBottom(workBook, sheetName, cell, bottom);
			}
		}
	},


	//下边框默样式
	setBorderBottomDefault: function(workBook, sheetName, cell) {
		this.init1(workBook, sheetName, cell, "border");
		workBook.Sheets[sheetName][cell].s.border.bottom = this.defaultBorderStyle;
		return workBook;
	},

	setBorderBottomDefaultAll: function(workBook, sheetName) {
		this.initAllCell(workBook, sheetName);
		for (cell in workBook.Sheets[sheetName]) {
			if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
				this.setBorderBottomDefault(workBook, sheetName, cell);
			}
		}
	},

	//左边框
	setBorderLeft: function(workBook, sheetName, cell, left) {
		this.init1(workBook, sheetName, cell, "border");
		workBook.Sheets[sheetName][cell].s.border.left = left;
		return workBook;
	},

	setBorderLeftAll: function(workBook, sheetName, left) {
		this.initAllCell(workBook, sheetName);
		for (cell in workBook.Sheets[sheetName]) {
			if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
				this.setBorderLeft(workBook, sheetName, cell, left);
			}
		}
	},

	setBorderLeftDefault: function(workBook, sheetName, cell) {
		this.init1(workBook, sheetName, cell, "border");
		workBook.Sheets[sheetName][cell].s.border.left = this.defaultBorderStyle;
		return workBook;
	},

	setBorderLeftDefaultAll: function(workBook, sheetName) {
		this.initAllCell(workBook, sheetName);
		for (cell in workBook.Sheets[sheetName]) {
			if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
				this.setBorderLeftDefault(workBook, sheetName, cell);
			}
		}
	},

	//右边框
	setBorderRight: function(workBook, sheetName, cell, right) {
		this.init1(workBook, sheetName, cell, "border");
		workBook.Sheets[sheetName][cell].s.border.right = right;
		return workBook;
	},

	setBorderRightAll: function(workBook, sheetName, right) {
		this.initAllCell(workBook, sheetName);
		for (cell in workBook.Sheets[sheetName]) {
			if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
				this.setBorderRight(workBook, sheetName, cell, right);
			}
		}
	},
	setBorderRightDefault: function(workBook, sheetName, cell) {
		this.init1(workBook, sheetName, cell, "border");
		workBook.Sheets[sheetName][cell].s.border.right = this.defaultBorderStyle;
		return workBook;
	},

	setBorderRightDefaultAll: function(workBook, sheetName) {
		this.initAllCell(workBook, sheetName);
		for (cell in workBook.Sheets[sheetName]) {
			if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
				this.setBorderRightDefault(workBook, sheetName, cell);
			}
		}
	},

	//对角线
	setBorderDiagonal: function(workBook, sheetName, cell, diagonal) {
		this.init1(workBook, sheetName, cell, "border");
		workBook.Sheets[sheetName][cell].s.border.diagonal = diagonal;
		return workBook;
	},

	setBorderDiagonalAll: function(workBook, sheetName, diagonal) {
		this.initAllCell(workBook, sheetName);
		for (cell in workBook.Sheets[sheetName]) {
			if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
				this.setBorderDiagonal(workBook, sheetName, cell, diagonal);
			}
		}
	},

	setBorderDiagonalDefault: function(workBook, sheetName, cell) {
		this.init1(workBook, sheetName, cell, "border");
		workBook.Sheets[sheetName][cell].s.border.diagonal = this.defaultBorderStyle;
		return workBook;
	},

	setBorderDiagonalDefaultAll: function(workBook, sheetName) {
		this.initAllCell(workBook, sheetName);
		for (cell in workBook.Sheets[sheetName]) {
			if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
				this.setBorderDiagonalDefault(workBook, sheetName, cell);
			}
		}
	},

	setBorderDiagonalUp: function(workBook, sheetName, cell, isDiagonalUp) {
		this.init1(workBook, sheetName, cell, "border");
		workBook.Sheets[sheetName][cell].s.border.diagonalUp = isDiagonalUp;
		return workBook;
	},

	setBorderDiagonalUpAll: function(workBook, sheetName, isDiagonalUp) {
		this.initAllCell(workBook, sheetName);
		for (cell in workBook.Sheets[sheetName]) {
			if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
				this.setBorderDiagonalUp(workBook, sheetName, cell, isDiagonalUp);
			}
		}
	},

	setBorderDiagonalDown: function(workBook, sheetName, cell, isDiagonalDown) {
		this.init1(workBook, sheetName, cell, "border");
		workBook.Sheets[sheetName][cell].s.border.diagonalDown = isDiagonalDown;
		return workBook;
	},

	setBorderDiagonalDownAll: function(workBook, sheetName, isDiagonalDown) {
		this.initAllCell(workBook, sheetName);
		for (cell in workBook.Sheets[sheetName]) {
			if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
				this.setBorderDiagonalDown(workBook, sheetName, cell, isDiagonalDown);
			}
		}
	},

	//默认样式，多单元格设置样式

	//设置所有单元格字体样式
	setFgColorStylesAll: function(workBook, sheetName, fontType, fontColor, fontSize) {
		for (cell in workBook.Sheets[sheetName]) {
			if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
				this.setFontType(workBook, sheetName, cell, fontType);
				this.setFontColorRGB(workBook, sheetName, cell, fontColor);
				this.setFontSize(workBook, sheetName, cell, fontSize);
			}
		}
	},

	//设置第一行标题自定义样式
	setTitleStyles: function(workBook, sheetName, fgColor, fontColor, alignment, isBold, fontSize, ) {
		for (cell in workBook.Sheets[sheetName]) {
			if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
				row = cell.substr(1);
				if (row == '1') {
					this.setFillFgColorRGB(workBook, sheetName, cell, fgColor);
					this.setFontColor(workBook, sheetName, cell, fontColor);
					this.setAlignmentHorizontal(workBook, sheetName, cell, alignment);
					this.setFontBold(workBook, sheetName, cell, isBold);
					this.setFontSize(workBook, sheetName, cell, fontSize);
				}
			}
		}
	},

	//设置第一行标题默认样式
	setTitleStylesDefault: function(workBook, sheetName) {
		for (cell in workBook.Sheets[sheetName]) {
			row = cell.substr(1);
			if (row == '1') {
				//setFillFgColorRGB(workBook, sheetName, cell, 'FFFF00');
				this.setAlignmentHorizontal(workBook, sheetName, cell, 'center');
				this.setFontBold(workBook, sheetName, cell, true);
				this.setFontSize(workBook, sheetName, cell, '20');
			}
		}
	},

	//设置双数行背景色灰色，便于阅读
	setEvenRowColorGrey: function(workBook, sheetName) {
		for (cell in workBook.Sheets[sheetName]) {
			if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
				row = parseInt(cell.substr(1));
				if (row % 2 == 0) {
					this.setFillFgColorRGB(workBook, sheetName, cell, 'DCDCDC');
				}

			}
		}
	},

	//合并同一列中内容一样的相邻行
	mergeSameColCells: function(workBook,sheetName,col) {
		var cells=[];
		for (cell in workBook.Sheets[sheetName]) {
			if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
				if (cell.substr(0, 1) == col) {
					cells.push(cell);//获得该列单元格数组，升序
				}
			}
		}
		for(var i = 0;i<cells.length-1;){
			for(var j=i+1;j<cells.length;j++){
				//内容一样且不为空则合并
				if(workBook.Sheets[sheetName][cells[i]].v == workBook.Sheets[sheetName][cells[j]].v && workBook.Sheets[sheetName][cells[i]].v != ""){
					this.mergeCells(workBook,sheetName,cells[i],cells[j]);
					if(j==cells.length-1){
						i=j;
					}
				}
				else{	//当且仅当相邻的两个cell值相同时才合并
					i=j;
					break;
				}
			}
		}
	},

	//合并同一行中内容一样的相邻列
	mergeSameRowCells: function(workBook,sheetName,row) {
		var cells=[];
		for (cell in workBook.Sheets[sheetName]) {
			if (cell != '!cols' && cell != '!merges' && cell != '!ref') {
				if (cell.substr(1) == row) {
					cells.push(cell);//获得该列单元格数组，升序
				}
			}
		}
		for(var i = 0;i<cells.length-1;){
			for(var j=i+1;j<cells.length;j++){
				if(workBook.Sheets[sheetName][cells[i]].v == workBook.Sheets[sheetName][cells[j]].v){
					this.mergeCells(workBook,sheetName,cells[i],cells[j]);
					if(j==cells.length-1){
						i=j;
					}
				}
				else{	//当且仅当相邻的两个cell值相同时才合并
					i=j;
					break;
				}
			}
		}
	},

	//当前表格最大行数 return int
	getMaxRow: function(workBook,sheetName) {
		var length = 0;
		for (var ever in workBook.Sheets[sheetName]) {
			temp = parseInt(ever.substr(1));
			if (temp > length) {
				length = temp;
			}
		}
		return length;
	},

	//当前表格最大列数 A起步 return string
	getMaxCol: function(workBook,sheetName) {
		var length = 'A';
		for (var ever in workBook.Sheets[sheetName]) {
			temp = ever.substr(0, 1);
			if (temp > length) {
				length = temp;
			}
		}
		return length;
	},

})