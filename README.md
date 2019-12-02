# XLSX-Style-Utils(XSU)
## 基于SheetJS以及XLSX-Style的纯前端带样式导出表格为Excel工具包

## 背景
SheetJS（又名js-xlsx，npm库名称为xlsx，node库也叫node-xlsx，以下简称**JX**），免费版不支持样式调整。

（顺便吐槽下这些名字乱的不行。。实际上又是同一个东西= =

JX官方说明文档：https://github.com/SheetJS/js-xlsx

XLSX-Style（npm库命名为xlsx-style，以下简称**XS**）基于JX二次开发，使其支持样式调整，但其开发停留在2017年，所基于的JX版本老旧，缺失许多方法。因而诞生了这个项目。

XS官方说明文档：https://github.com/protobi/js-xlsx

XLSX-Style-Utils：本项目 其本体为xlsxStyle.utils.js 以下简称**XSU**


## 文件描述：

FileSaver.js 导出保存excel用到的js

test.html 基于JX官方开发demo修改的测试用例https://sheetjs.com/demos/table.html ，包含utils中的方法的测试用例

xlsx.core.min.js JX最新版核心文件，建议在将网页表格导成workbook时使用其方法

xlsxStyle.core.min.js XS最新版核心文件，因为其原本命名与JX一样，避免冲突改名成xlsxStyle

xlsxStyle.utils.js XSU本项目核心文件，基于XS的方法二次封装，更好的控制导出excel的样式。以下简称XSU

## what did I do？

由于JX和XS所暴露出来的方法调用变量名一样（都是XLSX），同时引用时必然会覆盖掉另一个，故我将XS所暴露的变量名修改为xlsxStyle。调用XS方法时请使用此变量名。调用JX方法时使用XLSX。具体原因参考:https://blog.csdn.net/tian_i/article/details/84327329

对XS的样式调整进行二次封装在utils工具包中，部分测试用例参考：

例子1：
  ```javascript
  	//test
	var wb = wb1;
	var sheetName = wb.SheetNames[0];
	utilsTest(wb);
	//使用xlsxStyle.utils（XSU）对Workbook进行样式自定义
	function utilsTest(wb){
		XSU.mergeCells(wb,sheetName,"A1","B1");
		XSU.mergeCellsByObj(wb,sheetName,[{s: {c: 0, r: 2},e: {c: 0, r: 3}}]);
		//setColWidth(wb,sheetName,[{wpx: 45}, {wpx: 165}, {wpx: 45}, {wpx: 45}]);
		
		XSU.setFillFgColorRGB(wb,sheetName,"B4","FFB6C1");
		//setFillBgColorRGB(wb,sheetName,"B4","FFB6C1");
		
		XSU.setFontSize(wb,sheetName,"B4",60);
		XSU.setFontColorRGB(wb,sheetName,"B4","00BFFF");
		XSU.setFontBold(wb,sheetName,"B4",true);
		XSU.setFontUnderline(wb,sheetName,"B4",true);
		XSU.setFontItalic(wb,sheetName,"B4",true);
		XSU.setFontStrike(wb,sheetName,"B4",true);
		XSU.setFontShadow(wb,sheetName,"B4",true);
		XSU.setFontVertAlign(wb,sheetName,"B4",true);
		
		XSU.setAlignmentVertical(wb,sheetName,"B4","top");
		XSU.setAlignmentHorizontal(wb,sheetName,"B4","center");
		
		XSU.setBorderTopDefault(wb,sheetName,"B4");
		XSU.setBorderRightDefault(wb,sheetName,"D3");
		XSU.setBorderDefault(wb,sheetName,"C4");
		
		console.log(wb);

		XSU.setBorderDefaultAll(wb,sheetName);
		XSU.setTitleStylesDefault(wb,sheetName);
		XSU.setEvenRowColorGrey(wb,sheetName);
	}

	//转换成二进制 使用xlsx-style（XS）进行转换才能得到带样式Excel
	var wbout = xlsxStyle.write(wb,wopts);
	//保存，使用FileSaver.js
	return saveAs(new Blob([XSU.s2ab(wbout)],{type:""}), "test.xlsx");
  ```


![例子1效果图](https://github.com/Ctrl-Ling/XLSX-Style-Utils/blob/master/demo.png)

例子2：
```javascript
    var wb = wb1;
    var sheet = wb.SheetNames[0];
    //自定义对应表格样式
    setWorkbookStyle: function(wb,sheet){
        var cols = XSU.getMaxCol(wb,sheet);//当前最大列数
        var rows = XSU.getMaxRow(wb,sheet);//当前最大行数
        //wb样式处理，调用xlsxStyle.utils方法

        //------------------通用表格样式----------------------------
        XSU.mergeCells(wb,sheet,"A1",cols+'1'); //合并title单元格
        XSU.setFontTypeAll(wb,sheet,'仿宋');//字体：仿宋
        XSU.setAlignmentHorizontalAll(wb,sheet,'center');//垂直居中
        XSU.setAlignmentVerticalAll(wb,sheet,'center');//水平居中
        XSU.setAlignmentWrapTextAll(wb,sheet,true);//自动换行
        XSU.setFontBoldOfCols(wb,sheet,true,'A');//设置第一列加粗
        XSU.setFontBoldOfRows(wb,sheet,true,'2');//设置第二行标题行加粗
        XSU.setBorderDefaultAll(wb,sheet);//设置所有单元格默认边框

        //-------------------------个性化----------------------------
        //列宽设置 1wch为1英文字符宽度
        var width = [{wch: 25}, {wch: 15}, {wch: 15}, {wch: 15}];
        XSU.setColWidth(wb,sheet,width);

        XSU.setTitleStylesDefault(wb,sheet);//设置A1单元格title默认样式 必须最后设置 否则可能会被其他覆盖
    }
```
utils持续更新中。只干了一些微小的工作🐸测试用例较少，建议查看utils源码
  
  ## 使用
  
  在html头部引入4个JS即可
  
  1.使用**JX**自带的方法将网页表格导出成**不带样式**的workbook（此处应该啃食一下官方文档以及下方参考文章），使用XLSX.table_to_book等方法.
  
  2.对workbook使用**XSU**方法设置样式，得到**带样式**的workbook
  
   其中，setXXX()为设置某一单元格样式的方法
  
  setXXXAll()为设置所有单元格样式的方法
  
  3.对带样式的workbook使用**XS**的方法xlsxStyle.write()处理workbook再用saveAs()保存成excel，具体参考test.html
  
  
  
  ## 建议参考文章：
  
  https://segmentfault.com/a/1190000018077543?utm_source=tag-newest
  
  https://www.cnblogs.com/liuxianan/p/js-excel.html
  
  https://www.jianshu.com/p/877631e7e411
  
  https://www.jianshu.com/p/74d405940305
  
  https://www.jianshu.com/p/869375439fee
  
  https://blog.csdn.net/tian_i/article/details/84327329
  

  
