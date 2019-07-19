# XLSX-Style-Utils
## 基于SheetJS以及XLSX-Style的纯前端带样式导出表格到Excel的工具包

## 背景
SheetJS（又名js-xlsx，npm库名称为xlsx，node库也叫node-xlsx，以下简称JX），免费版不支持样式调整。

（顺便吐槽下这些名字乱的不行。。实际上又是同一个东西= =

JX官方说明文档：https://github.com/SheetJS/js-xlsx

XLSX-Style（npm库命名为xlsx-style，以下简称XS）基于JX二次开发，使其支持样式调整，但其开发停留在2017年，所基于的JX版本老旧，缺失许多方法。因而诞生了这个项目。

XS官方说明文档：https://github.com/protobi/js-xlsx

XLSX-Style-Utils：本项目 其本体为xlsxStyle.utils.js 以下简称utils


## 文件描述：

FileSaver.js 导出保存excel用到的js

test.html 基于JX官方开发demo修改的测试用例https://sheetjs.com/demos/table.html ，包含utils中的方法的测试用例

xlsx.core.min.js JX最新版核心文件，建议在将网页表格导成workbook时使用其方法

xlsxStyle.core.min.js XS最新版核心文件，因为其原本命名与JX一样，避免冲突改名成xlsxStyle

xlsxStyle.utils.js 本项目核心文件，基于XS的方法二次封装，更好的控制导出excel的样式。以下简称utils

## what did I do？

由于JX和XS所暴露出来的方法调用变量名一样（都是XLSX），同时引用时必然会覆盖掉另一个，故我将XS所暴露的变量名修改为xlsxStyle。调用XS方法时请使用此变量名。调用JX方法时使用XLSX。具体原因参考:https://blog.csdn.net/tian_i/article/details/84327329

对XS的样式调整进行二次封装在utils工具包中，部分测试用例参考：
  ```javascript
  function doit(type, fn, dl) {
	var elt = document.getElementById('data-table');
	//从table转换成workbook
	var wb1 = XLSX.utils.table_to_book(elt, {sheet:"Sheet JS"});
  	//导出格式设置
	var wopts = { bookType:'xlsx', bookSST:false, type:'binary' };
	//test
	var wb = wb1;
	var sheetName = wb.SheetNames[0];
	utilsTest(wb);
	function utilsTest(wb){
		mergeCells(wb,sheetName,"A1","B1");
		mergeCellsByObj(wb,sheetName,[{s: {c: 0, r: 2},e: {c: 0, r: 3}}]);
		setColWidth(wb,sheetName,[{wpx: 45}, {wpx: 165}, {wpx: 45}, {wpx: 45}]);
		
		setFillFgColorRGB(wb,sheetName,"B4","FFB6C1");
		
		setFontSize(wb,sheetName,"B4",60);
		setFontColorRGB(wb,sheetName,"B4","00BFFF");
		setFontBold(wb,sheetName,"B4",true);
		setFontUnderline(wb,sheetName,"B4",true);
		setFontItalic(wb,sheetName,"B4",true);
		setFontStrike(wb,sheetName,"B4",true);
		setFontShadow(wb,sheetName,"B4",true);
		setFontVertAlign(wb,sheetName,"B4",true);
		
		setAlignmentVertical(wb,sheetName,"B4","top");
		setAlignmentHorizontal(wb,sheetName,"B4","center");
		
		setBorderTopDefault(wb,sheetName,"B4");
		setBorderRightDefault(wb,sheetName,"D3");
		setBorderDefault(wb,sheetName,"C4");
	}
	//转换成二进制
	var wbout = xlsxStyle.write(wb,wopts);
	//保存，使用FileSaver.js
	return saveAs(new Blob([s2ab(wbout)],{type:""}), "test.xlsx");
}	
  ```
  utils持续更新中。只干了一些微小的工作🐸测试用例较少，建议查看utils源码
  
  ## 使用
  
  使用JX自带的方法将网页表格导出成不带样式的workbook（此处应该啃食一下官方文档以及下方参考文章），使用XLSX.table_to_book等方法.
  
  对workbook使用utils方法设置样式，得到带样式的workbook
  
  ### ！重要！
  
  使用xlsxStyle.write()处理workbook再用saveAs()保存成excel，具体参考test.html
  
  
  
  
  
  ## 建议参考文章：
  
  https://segmentfault.com/a/1190000018077543?utm_source=tag-newest
  
  https://www.cnblogs.com/liuxianan/p/js-excel.html
  
  https://www.jianshu.com/p/877631e7e411
  
  https://www.jianshu.com/p/74d405940305
  
  https://www.jianshu.com/p/869375439fee
  
  https://blog.csdn.net/tian_i/article/details/84327329
  

  
