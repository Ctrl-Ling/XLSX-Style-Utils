//字符串转字符流
function s2ab(s) {
    var buf = new ArrayBuffer(s.length);
    var view = new Uint8Array(buf);
    for (var i = 0; i != s.length; ++i)
        view[i] = s.charCodeAt(i) & 0xFF;
    return buf;
}

//初始化
function init(workBook,sheetName,cell){
		if(!workBook.Sheets[sheetName][cell].s){
		workBook.Sheets[sheetName][cell].s = {};
	}
}

function init1(workBook,sheetName,cell,attr){
	init(workBook,sheetName,cell);
	if(!workBook.Sheets[sheetName][cell].s[attr]){
		workBook.Sheets[sheetName][cell].s[attr]= {};
	}
}

function init2(workBook,sheetName,cell,attr1,attr2){
	init(workBook,sheetName,cell);
	init1(workBook,sheetName,cell,attr1);
	if(!workBook.Sheets[sheetName][cell].s[attr1][attr2]){
		workBook.Sheets[sheetName][cell].s[attr1][attr2]= {};
	}
}

//单元格合并 startCell=A1 endCell=B5
function mergeCells(workBook, sheetName, startCell, endCell) {
    var sc = startCell.substr(0, 1).charCodeAt(0) - 65;
    var sr = startCell.substr(1);
	sr = parseInt(sr)-1;
    var ec = endCell.substr(0, 1).charCodeAt(0) - 65;
    var er = endCell.substr(1)
	er = parseInt(er)-1;

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
	if(!workBook.Sheets[sheetName]["!merges"]){
		workBook.Sheets[sheetName]["!merges"]=merges;
	}
	else{
		workBook.Sheets[sheetName]["!merges"]=workBook.Sheets[sheetName]["!merges"].concat(merges);
	}
	
	return workBook;

}

//merges=[{s: {c: 0, r: 0},e: {c: 3, r: 0}}]
function mergeCellsByObj(workBook, sheetName, merges) {

	if (workBook.Sheets[sheetName]["!merges"]){
		workBook.Sheets[sheetName]["!merges"]=workBook.Sheets[sheetName]["!merges"].concat(merges);
	}
	else {
		workBook.Sheets[sheetName]["!merges"]=merges;
	}
	
	return workBook;

}


//设置每列列宽,单位px cols= [{wpx: 45}, {wpx: 165}, {wpx: 45}, {wpx: 45}]
function setColWidth(workBook,sheetName,cols){
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
function setCellStyle(workBook,sheetName,cell,styles){
	workBook.Sheets[sheetName][cell].s = styles;
	return workBook;
}

/*Fill*/

//填充颜色 fill={"bgColor": {"indexed": 64},"fgColor": {"rgb": "FFFFFF00"}}
function setFillStyles(workBook,sheetName,cell,styles){
	init(workBook,sheetName,cell);
	workBook.Sheets[sheetName][cell].s.fill = styles;
	return workBook;
}

//patternType="solid" or "none"”
function setFillPatternType(workBook,sheetName,cell,patternType){
	init1(workBook,sheetName,cell,"fill");
	workBook.Sheets[sheetName][cell].s.fill.patternType = patternType;
	return workBook;
}

//前景颜色(单元格颜色) rgb 
//COLOR_SPEC属性值:{ auto: 1}指定自动值,{ rgb: "FFFFAA00" }指定16进制的ARGB,{ theme: "1", tint: "-0.25"}指定主题颜色和色调的整数索引（默认值为0）,{ indexed: 64} 默认值 fill.bgColor
function setFillFgColor(workBook,sheetName,cell,COLOR_SPEC){
	init1(workBook,sheetName,cell,"fill");
	workBook.Sheets[sheetName][cell].s.fill.fgColor = COLOR_SPEC;
	return workBook;
}

function setFillFgColorRGB(workBook,sheetName,cell,rgb){
	init2(workBook,sheetName,cell,"fill","fgColor");
	workBook.Sheets[sheetName][cell].s.fill.fgColor.rgb = rgb;
	return workBook;
}

//单元格背景颜色（貌似没用）
function setFillBgColor(workBook,sheetName,cell,COLOR_SPEC){
	init1(workBook,sheetName,cell,"fill");
	workBook.Sheets[sheetName][cell].s.fill.bgColor = COLOR_SPEC;
	return workBook;
}

//单元格背景颜色
function setFillBgColorRGB(workBook,sheetName,cell,rgb){
	init2(workBook,sheetName,cell,"fill","bgColor");
	workBook.Sheets[sheetName][cell].s.fill.bgColor.rgb = rgb;
	return workBook;
}


/*Font*/

//字体风格，可一次性在styles中设置所有font风格
function setFontStyles(workBook,sheetName,cell,styles){
	init(workBook,sheetName,cell);
	workBook.Sheets[sheetName][cell].s.font = styles;
	return workBook;
}

//字体 type="Calibri" 
function setFontType(workBook,sheetName,cell,type){
	init1(workBook,sheetName,cell,"font");
	workBook.Sheets[sheetName][cell].s.font.name = type;
	return workBook;
}

//字体大小
function setFontSize(workBook,sheetName,cell,size){
	init1(workBook,sheetName,cell,"font");
	workBook.Sheets[sheetName][cell].s.font.sz = size;
	return workBook;
}

//字体颜色 COLOR_SPEC
function setFontColor(workBook,sheetName,cell,COLOR_SPEC){
	init1(workBook,sheetName,cell,"font");
	workBook.Sheets[sheetName][cell].s.font.color = COLOR_SPEC;
	//workBook.Sheets[sheetName][cell].s.font.color.rgb = rgb;
	return workBook;
}

//字体颜色RGB
function setFontColorRGB(workBook,sheetName,cell,rgb){
	init2(workBook,sheetName,cell,"font","color");
	workBook.Sheets[sheetName][cell].s.font.color.rgb = rgb;
	return workBook;
}

//是否粗体 boolean isBold
function setFontBold(workBook,sheetName,cell,isBold){
	init1(workBook,sheetName,cell,"font");
	workBook.Sheets[sheetName][cell].s.font.bold = isBold;
	return workBook;
}
//是否下划线 boolean isUnderline
function setFontUnderline(workBook,sheetName,cell,isUnderline){
	init1(workBook,sheetName,cell,"font");
	workBook.Sheets[sheetName][cell].s.font.underline = isUnderline;
	return workBook;
}
//是否斜体 boolean isItalic
function setFontItalic(workBook,sheetName,cell,isItalic){
	init1(workBook,sheetName,cell,"font");
	workBook.Sheets[sheetName][cell].s.font.italic = isItalic;
	return workBook;
}
//是否删除线 boolean isStrike
function setFontStrike(workBook,sheetName,cell,isStrike){
	init1(workBook,sheetName,cell,"font");
	workBook.Sheets[sheetName][cell].s.font.strike = isStrike;
	return workBook;
}
//是否outline boolean isOutline
function setFontOutline(workBook,sheetName,cell,isOutline){
	init1(workBook,sheetName,cell,"font");
	workBook.Sheets[sheetName][cell].s.font.outline = isOutline;
	return workBook;
}
//是否阴影 boolean isShadow
function setFontShadow(workBook,sheetName,cell,isShadow){
	init1(workBook,sheetName,cell,"font");
	workBook.Sheets[sheetName][cell].s.font.shadow = isShadow;
	return workBook;
}
//是否vertAlign boolean isVertAlign
function setFontVertAlign(workBook,sheetName,cell,isVertAlign){
	init1(workBook,sheetName,cell,"font");
	workBook.Sheets[sheetName][cell].s.font.vertAlign = isVertAlign;
	return workBook;
}


/*numFmt*/

function setNumFmt(workBook,sheetName,cell,numFmt){
	init(workBook,sheetName,cell);
	workBook.Sheets[sheetName][cell].s.numFmt = numFmt;
	return workBook;
}


/*Alignment*/

//文本对齐 alignment={vertical:top,horizontal:top,}
function setAlignmentStyles(workBook,sheetName,cell,styles){
	init(workBook,sheetName,cell);
	workBook.Sheets[sheetName][cell].s.alignment = styles;
	return workBook;
}

//文本垂直对齐 vertical	="bottom" or "center" or "top"
function setAlignmentVertical(workBook,sheetName,cell,vertical){
	init1(workBook,sheetName,cell,"alignment");
	workBook.Sheets[sheetName][cell].s.alignment.vertical = vertical;
	return workBook;
}

//文本水平对齐 "bottom" or "center" or "top"
function setAlignmentHorizontal(workBook,sheetName,cell,horizontal){
	init1(workBook,sheetName,cell,"alignment");
	workBook.Sheets[sheetName][cell].s.alignment.horizontal = horizontal;
	return workBook;
}

function setAlignmentWrapText(workBook,sheetName,cell,isWrapText){
	init1(workBook,sheetName,cell,"alignment");
	workBook.Sheets[sheetName][cell].s.alignment.isWrapText = isWrapText;
	return workBook;
}

function setAlignmentReadingOrder(workBook,sheetName,cell,readingOrder){
	init1(workBook,sheetName,cell,"alignment");
	workBook.Sheets[sheetName][cell].s.alignment.readingOrder = readingOrder;
	return workBook;
}

//文本旋转角度 0-180，255 is special, aligned vertically
function setAlignmentTextRotation(workBook,sheetName,cell,textRotation){
	init1(workBook,sheetName,cell,"alignment");
	workBook.Sheets[sheetName][cell].s.alignment.textRotation = textRotation;
	return workBook;
}



/*Border*/

//边框 styles={top:{ style:"thin",color:"FFFFAA00"},bottom:{},...}
function setBorderStyles(workBook,sheetName,cell,styles){
	init(workBook,sheetName,cell);
	workBook.Sheets[sheetName][cell].s.border = styles;
	return workBook;
}

//单元格四周边框默认样式
const borderAll = {
  top: {style: 'thin'},
  bottom: {style: 'thin'},
  left: {style: 'thin'},
  right: {style: 'thin'}
};

//边框默样式：细线黑色
const defaultBorderStyle ={style: 'thin'};

function setBorderDefault(workBook,sheetName,cell){
	//workBook.Sheets[sheetName][cell].s= {};
	/*if(!workBook.Sheets[sheetName][cell].s.alignment){
	workBook.Sheets[sheetName][cell].s.alignment = {};
	}
	if(!workBook.Sheets[sheetName][cell].s.alignment.horizontal){
	workBook.Sheets[sheetName][cell].s.alignment.horizontal = {};
	}*/
	//init1(workBook,sheetName,cell,"alignment");
	
	//init1(workBook,sheetName,cell,"border");
	//workBook.Sheets[sheetName][cell].s.alignment.horizontal = "center";
	init(workBook,sheetName,cell);
	workBook.Sheets[sheetName][cell].s.border= borderAll;
	return workBook;
}

//上边框
function setBorderTop(workBook,sheetName,cell,top){
	init(workBook,sheetName,cell,"border");
	workBook.Sheets[sheetName][cell].s.border.top = top;
	return workBook;
}

//上边框默样式
function setBorderTopDefault(workBook,sheetName,cell){
	init1(workBook,sheetName,cell,"border");
	workBook.Sheets[sheetName][cell].s.border.top = defaultBorderStyle;
	return workBook;
}

//下边框
function setBorderBottom(workBook,sheetName,cell,bottom){
	init1(workBook,sheetName,cell,"border");
	workBook.Sheets[sheetName][cell].s.border.bottom = bottom;
	return workBook;
}

//下边框默样式
function setBorderBottomDefault(workBook,sheetName,cell){
	init1(workBook,sheetName,cell,"border");
	workBook.Sheets[sheetName][cell].s.border.bottom = defaultBorderStyle;
	return workBook;
}

function setBorderLeft(workBook,sheetName,cell,left){
	init1(workBook,sheetName,cell,"border");
	workBook.Sheets[sheetName][cell].s.border.left = left;
	return workBook;
}

function setBorderLeftDefault(workBook,sheetName,cell){
	init1(workBook,sheetName,cell,"border");
	workBook.Sheets[sheetName][cell].s.border.left = defaultBorderStyle;
	return workBook;
}

function setBorderRight(workBook,sheetName,cell,right){
	init1(workBook,sheetName,cell,"border");
	workBook.Sheets[sheetName][cell].s.border.right = right;
	return workBook;
}

function setBorderRightDefault(workBook,sheetName,cell){
	init1(workBook,sheetName,cell,"border");
	workBook.Sheets[sheetName][cell].s.border.right = defaultBorderStyle;
	return workBook;
}

//对角线
function setBorderDiagonal(workBook,sheetName,cell,diagonal){
	init1(workBook,sheetName,cell,"border");
	workBook.Sheets[sheetName][cell].s.border.diagonal = diagonal;
	return workBook;
}

function setBorderDiagonalDefault(workBook,sheetName,cell){
	init1(workBook,sheetName,cell,"border");
	workBook.Sheets[sheetName][cell].s.border.diagonal = defaultBorderStyle;
	return workBook;
}

function setBorderDiagonalUp(workBook,sheetName,cell,isDiagonalUp){
	init1(workBook,sheetName,cell,"border");
	workBook.Sheets[sheetName][cell].s.border.diagonalUp = isDiagonalUp;
	return workBook;
}

function setBorderDiagonalDown(workBook,sheetName,cell,isDiagonalDown){
	init1(workBook,sheetName,cell,"border");
	workBook.Sheets[sheetName][cell].s.border.diagonalDown = isDiagonalDown;
	return workBook;
}










