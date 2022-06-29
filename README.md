# easyExcelJs
简单的操作生成漂亮的EXCEL，快速上手。提供漂亮模板直接使用



## npm install

```shell
npm install easyexceljs -S
npm install exceljs -S
npm install file-saver -S
```

## 快速开始  以vue项目里使用为例
```js
import {createExcelByOneSheet,ExcelStyleTemplate,getCellPosLetter} from "easyexceljs"

//用例数据  github上已提供
import headArray from "./expmaleDate/headarray";
import bodyarray from "./expmaleDate/bodyarray";

methods:{
    createExcelExpmale(){
        //创建一个实例
        const _createExcelByOneSheet = new createExcelByOneSheet(excelOptions);
    }
}


```