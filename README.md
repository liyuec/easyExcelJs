# easyExcelJs

![easyExcelJs](https://raw.githubusercontent.com/liyuec/pictures/main/easyExceljs/exceljslogo.png)

简单的操作生成漂亮的EXCEL，快速上手。提供漂亮模板直接使用

### 若存在疑问和支持，请邮件2362259057#qq.com('#'替换为'@')

<p align="left">
    <img src="https://www.oscs1024.com/platform/badge/liyuec/easyExcelJs.svg" />
    <img src="https://img.shields.io/badge/size-6.56kb-blue" />
    <img src="https://img.shields.io/badge/license-MIT-orange" />
    <img src="https://img.shields.io/badge/converage-50%25-red" />
    <img src="https://img.shields.io/badge/version-1.0.0-lightgrey" />
</p>


## npm install

组件依附 [exceljs](https://github.com/exceljs/exceljs) 和 [file-saver](https://github.com/eligrey/FileSaver.js) 进行封装，需要install相关依赖；
<p style="color:red;font-size:16px;">
    特此感谢
</p>


```shell
npm install easyexceljs -S
npm install exceljs -S
npm install file-saver -S
```

## 快速开始 生成一个Excel  以vue项目里使用为例   
```javascript
import {createExcelByOneSheet,ExcelStyleTemplate,getCellPosLetter} from "easyexceljs"

//用例数据  github上已提供   uri:https://github.com/liyuec/easyExcelJs/tree/main/expmale/testData
import headArray from "./expmaleDate/headarray";
import bodyarray from "./expmaleDate/bodyarray";

methods:{
    createExcelExpmale(){
        //new的时候需要传入基本的options，不传会默认变为'sheet1' 和 'excel'
        const excelOptions = {
            excelFileName: "XX公司年度报表",
            sheetName:'本季度报表1'
        };
        //创建一个实例
        const _createExcelByOneSheet = new createExcelByOneSheet(excelOptions);
        //定义excel head部分的格式
        const _head = [];


        
        //选择整个excel的样式  目前一共3个
         const red = ExcelStyleTemplate.red;

        //需要对应到列   其中header,key,width必传,  key 对应 bodyArray的key ， header为显示内容，width为每列宽度
        headArray.forEach((i) => {
            _head.push({
            header: i.title,
            key: i.field,
            width: 25,
            });
        });

      //赋值excel head数据（第1行）
      _createExcelByOneSheet.sheetColumnsData = [..._head];
      //赋值excel body数据， 按照head[] 数据结构进行匹配
      _createExcelByOneSheet.sheetRowsData = [...bodyarray];

      //传入样式
      _createExcelByOneSheet.rowStyleOptions.push(red.rowStyle);

      //下载得到excel
      _createExcelByOneSheet.saveAsExcel();
    }
}


```

## 1.0.0提供的3个可立即使用的模板
![模板展示](https://raw.githubusercontent.com/liyuec/pictures/main/easyExceljs/ExcelStyleTemplate_first.png)
```javascript
//引入提供模板Style
import {ExcelStyleTemplate} from "easyexceljs"

//一共三个模板
const red = ExcelStyleTemplate.red;
const blue = ExcelStyleTemplate.blue;
const green = ExcelStyleTemplate.green;
```


### red模板最终样式
![red模板样式](https://raw.githubusercontent.com/liyuec/pictures/main/easyExceljs/red.png)

### blue模板最终样式
![blue模板样式](https://raw.githubusercontent.com/liyuec/pictures/main/easyExceljs/blue.png)

### green模板最终样式
![green模板样式](https://raw.githubusercontent.com/liyuec/pictures/main/easyExceljs/green.png)

## 接下来准备
-   整行样式设置（字体，字号，颜色，背景色）
-   列：图片引入
-   列：链接引入
-   非常容易和轻松的对单独列边框设置（强调部分数据效果）
-   非常容易和轻松的对单独列样式（强调非边框外的效果）
-   提供更多 可直接使用的模板
-   提供 单excel里多个sheet 的构造
-   编写测试用例，增加覆盖率