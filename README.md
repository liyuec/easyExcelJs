# easyExcelJs

![easyExcelJs](https://raw.githubusercontent.com/liyuec/pictures/main/easyExceljs/exceljslogo.png)

简单的操作生成漂亮的EXCEL，快速上手。提供漂亮模板直接使用

### 若存在疑问和支持，请邮件2362259057#qq.com('#'替换为'@') 或则  isSues

<p align="left">
    <img src="https://www.oscs1024.com/platform/badge/liyuec/easyExcelJs.svg" />
    <img src="https://img.shields.io/badge/size-6.56kb-blue" />
    <img src="https://img.shields.io/badge/license-MIT-orange" />
    <img src="https://img.shields.io/badge/converage-50%25-red" />
    <img src="https://img.shields.io/badge/version-1.0.0-lightgrey" />,
</p>

# 目录
<ul>
  <li><a href="#npm-install">npm install</a></li>
  <li><a href="#快速开始生成一个excel-以vue项目里使用为例">快速开始</a></li>
  <li><a href="#可见基本模板100提供的3个可立即使用的模板">可见基本模板</a></li>
  <li><a href="#可见基本模板100提供的3个可立即使用的模板">提供的模板对象</a></li>
  <li>
    <a href="#接口">其他使用</a>
    <ul>
      <li><a href="#通过行数和列数获取Excel坐标">通过行数和列数获取Excel坐标</a></li>
      <li><a href="#通过Where条件设置Cell样式">通过Where条件设置Cell样式</a></li>
      <li><a href="#通过指定行列设置cell样式">通过指定行·列，设置Cell样式</a></li>
      <li><a href="#通过指定行列设置Cell的注解">通过指定行·列，设置Cell的注解</a></li>
      <li><a href="#通过指定行列设置返回原始Cell用户可根据原始Cell进行callBack">通过指定行·列设置，返回原始Cell，用户可根据原始Cell进行callBack</a></li>
    </ul>
  </li>
  <li><a href="#保存EXCEL">保存EXCEL</a></li>
  <li><a href="#继续开发计划">继续开发计划</a></li>
</ul>


## npm install[⬆](#目录)<!-- Link generated with jump2header -->

组件依附 [exceljs](https://github.com/exceljs/exceljs) 和 [file-saver](https://github.com/eligrey/FileSaver.js) 进行封装，需要install相关依赖；
<p style="color:red;font-size:16px;">
    特此感谢
</p>

```shell
npm install easyexceljs -S
npm install exceljs -S
npm install file-saver -S
```

## 快速开始(生成一个Excel 以vue项目里使用为例)[⬆](#目录)<!-- Link generated with jump2header -->
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

        /*需要对应到列   
            其中header,key,width必传,  
            key 对应 bodyArray的key ，body中每行需对应的key
            header为第1行head显示的可见内容，比如 keyFiled1 对应的 title:“字段名称”，
            width为每列宽度
        
        */
        headArray.forEach((i) => {
            _head.push({
            header: i.title,
            key: i.field,
            width: 25
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

## 可见基本模板(1.0.0提供的3个可立即使用的模板)[⬆](#目录)<!-- Link generated with jump2header -->

![模板展示](https://raw.githubusercontent.com/liyuec/pictures/main/easyExceljs/ExcelStyleTemplate_first.png)
```javascript
//引入提供模板Style
import {ExcelStyleTemplate} from "easyexceljs"

//一共三个模板
const red = ExcelStyleTemplate.red;
const blue = ExcelStyleTemplate.blue;
const green = ExcelStyleTemplate.green;
```

#### red模板最终样式
![red模板样式](https://raw.githubusercontent.com/liyuec/pictures/main/easyExceljs/red.png)

#### blue模板最终样式
![blue模板样式](https://raw.githubusercontent.com/liyuec/pictures/main/easyExceljs/blue.png)

#### green模板最终样式
![green模板样式](https://raw.githubusercontent.com/liyuec/pictures/main/easyExceljs/green.png)




## 提供的模板对象[⬆](#目录)<!-- Link generated with jump2header -->
#### 默认提供如下样式的模板对象，可根据需求自行修改颜色， 建议不要更改属性，方便兼容
```javascript
  import {ExcelStyleTemplate,getExcelCellStyle} from "easyexceljs"

  //默认结构，未考虑版本兼容，可以理解为baseDTO
  CellStyleDTO(){
    let obj = new Object(
        {
            cellIndex:1,
            rowIndex:1,
            cellName:'',
            BorderColor:'',
            BorderStyle:'',
            font:{
                name:'',
                size:'',
                bold:'',
                color:''
            }
        }
    )

    return obj;
  }
  
  //获取基本的cell样式，可自行根据需求更改颜色，字体大小，字体等
  getExcelCellStyle = function(colorTemplate){
    var cellStyle = new CellStyleDTO();
    switch(colorTemplate){
        case "red":
            cellStyle.BorderColor = 'ffff0000'
            cellStyle.BorderStyle = 'thin'
            cellStyle.font = {
                name:'Malgun Gothic Semilight',
                size:11,
                bold:true,
                color:'ffff0000'
            }
        break;
        case "blue":
            cellStyle.BorderColor = 'ff5faee3'
            cellStyle.BorderStyle = 'thin'
            cellStyle.font = {
                name:'Malgun Gothic Semilight',
                size:11,
                bold:true,
                color:'ff5faee3'
            }
        break;
        case "green":
            cellStyle.BorderColor = 'ff48c9b0'
            cellStyle.BorderStyle = 'thin'
            cellStyle.font = {
                name:'Malgun Gothic Semilight',
                size:11,
                bold:true,
                color:'ff48c9b0'
            }
        break;
        default:
            cellStyle.BorderColor = ''
            cellStyle.BorderStyle = ''
            cellStyle.font = {
                name:'宋体',
                size:11,
                bold:false,
                color:'ff000000'
            }
        break;
    }

    return cellStyle;
}

  //默认模板样式
  ExcelStyleTemplate = {
    'red': {
        rowStyle:{
            rowNum:1,
            rowBgColor: 'FFFF0000',
            font:{
                name:'Arial',
                size:12,
                bold:true,
                color:'ffffffff'
            }
        },
        cellStyle:{
            cellName:'',
            BorderColor: 'FFFF0000',
            BorderStyle:'thin',
            font:{
                name:'Arial',
                size:11,
                bold:true,
                color:'ff707b7c'
            }
        }
    },
    'blue':{
        rowStyle:{
            rowNum:1,
            rowBgColor: 'ff5faee3',
            font:{
                name:'Arial',
                size:12,
                bold:true,
                color:'ffffffff'
            }
        },
        cellStyle:{
            cellName:'',
            BorderColor: 'ff48c9b0',
            BorderStyle:'thin',
            font:{
                name:'Arial',
                size:11,
                bold:true,
                color:'ff707b7c'
            }
        }
    },
    'green':{
        rowStyle:{
            rowNum:1,
            rowBgColor: 'ff48c9b0',
            font:{
                name:'Arial',
                size:12,
                bold:true,
                color:'ffffffff'
            }
        },
        cellStyle:{
            cellName:'',
            BorderColor: 'ffff0000',
            BorderStyle:'thin',
            font:{
                name:'Arial',
                size:11,
                bold:true,
                color:'ff707b7c'
            }
        }
    }
}

```

## 通过行数和列数获取Excel坐标[⬆](#目录)<!-- Link generated with jump2header -->
```javascript
  import {getCellPosLetter} from "easyexceljs"

  /*
    比如传入 getCellPosLetter(1,1) 得到 A1  
    传入 getCellPosLetter(27,1) 得到 AA1
  */
  getCellPosLetter(cellIndex,rowIndex)

```

## 通过Where条件设置Cell样式[⬆](#目录)<!-- Link generated with jump2header -->
#### 方法名称  setCellStyleByWhere
#### where数据结构
```javascript

  /*
    where条件必传，结构见下段代码
    cellStyle 可不指定，不指定将默认使用 默认样式

    可以链式调用
  */
  setCellStyleByWhere(where,cellStyle)
  /*
    必传
    若实体结构不正确，或则不包含对应的值，将忽略本次where条件
  */
   where:{
        valueKey: '传入头部的key的值'
        whereType: < | > | == | != | === | !== | indexOf | unIndexOf
        whereValue: number | string
    }
```

| 属性名            | 描述 |
| ---------------- | ----------- |
| >          | 找到 大于 whereValue |
| <         | 找到 小于 whereValue |
| ==        | 找到 等于 whereValue的字段，并隐式类型转换 |
| ===       | 找到 等于 whereValue的字段，并且进行类型判断 |
| !==       | 找到 不等于 whereValue的字段，并且进行类型判断 |
| indexOf   | 找到 包含 whereValue的字段，可以理解为左右模糊查询 |
| unIndexOf   | 找到 不包含 whereValue的字段  |

#### 参考使用代码

```javascript

import {createExcelByOneSheet,getExcelCellStyle} from "easyexceljs"
   
    const excelOptions = {
          excelFileName: "XX公司年度报表",
          sheetName:'本季度报表1'
    };
      //创建一个实例
    const _createExcelByOneSheet = new createExcelByOneSheet(excelOptions);
    /*
      设置header , body
      此处代码略，参照  “快速开始”
    */

    let whereSelectByUserName = {
      valueKey:'userName',
      whereType:'indexOf',
      whereValue:'李三'
    },
    whereSelectByUserId = {
      valueKey:'userId',
      whereType:'>',
      whereValue:10000
    },
    whereSelectByNickName = {
      valueKey:'NickName',
      whereType:'indexOf',
      whereValue:'用户名'
    },
    cellStyle = getExcelCellStyle('red';

    _createExcelByOneSheet
    .setCellStyleByWhere(whereSelectByUserName,cellStyle)
    .setCellStyleByWhere(whereSelectByUserId,cellStyle)
    .setCellStyleByWhere(whereSelectByUserId)

```

## 通过指定行·列设置，Cell样式[⬆](#目录)<!-- Link generated with jump2header -->

#### setCellStyleByRowCellIndex(rowCellIndex,cellStyle)  
####  rowCellIndex数据结构 = [[rowIndex,cellIndex],[rowIndex,cellIndex]]
```javascript

import {createExcelByOneSheet,getExcelCellStyle} from "easyexceljs"
   
    const excelOptions = {
          excelFileName: "XX公司年度报表",
          sheetName:'本季度报表1'
    };
      //创建一个实例
    const _createExcelByOneSheet = new createExcelByOneSheet(excelOptions);
    /*
      设置header , body
      此处代码略，参照  “快速开始”
    */
 
    let cellStyle = getExcelCellStyle('blue');

    _createExcelByOneSheet
    .setCellStyleByRowCellIndex([[rowIndex,cellIndex],[rowIndex,cellIndex]],cellStyle)
    .setCellStyleByRowCellIndex([[rowIndex,cellIndex],[rowIndex,cellIndex]])

```

## 通过指定行·列设置，Cell的注解[⬆](#目录)<!-- Link generated with jump2header -->
#### customSetValueByIndex(rowCellIndex,callBack,repairLength = 0)  
####  rowCellIndex : [[rowIndex,cellIndex],[rowIndex,cellIndex]]
####  callBack : function()
```javascript

import {createExcelByOneSheet,getExcelCellStyle} from "easyexceljs"
   
    const excelOptions = {
          excelFileName: "XX公司年度报表",
          sheetName:'本季度报表1'
    };
      //创建一个实例
    const _createExcelByOneSheet = new createExcelByOneSheet(excelOptions);
    /*
      设置header , body
      此处代码略，参照  “快速开始”
    */
 
    let cellStyle = getExcelCellStyle('blue');

    _createExcelByOneSheet
    .setCellNoteTextByRowCellIndex([[rowIndex,cellIndex],[rowIndex,cellIndex]],['注册可以是任何内容，若存在特殊表情，需要依赖系统自身解析',''])
    .setCellNoteTextByRowCellIndex([[rowIndex,cellIndex],[rowIndex,cellIndex]],['',''])

```


## 通过指定行·列设置，返回原始Cell，用户可根据原始Cell进行callBack[⬆](#目录)<!-- Link generated with jump2header -->
#### customSetValueByIndex(rowCellIndex,noteTexts,repairLength)  
####  rowCellIndex数据结构 = [[rowIndex,cellIndex],[rowIndex,cellIndex]] 或则 '*'
####  callBack : function(cell){}    其中cell是得到每列的原始数据，可进行操作
####  repairLength : int32   冗余行数  需要补充的row length 遍历，若head只有一行，则为1


```javascript

  import {createExcelByOneSheet,getExcelCellStyle} from "easyexceljs"
   
    const excelOptions = {
          excelFileName: "XX公司年度报表",
          sheetName:'本季度报表1'
    };
      //创建一个实例
    const _createExcelByOneSheet = new createExcelByOneSheet(excelOptions);
    /*
      设置header , body
      此处代码略，参照  “快速开始”
    */

    _createExcelByOneSheet
    .customSetValueByIndex('*',function(cell){
                //可打印cell 查看需要包含的属性
                console.log(cell)
                let {value} = cell;

                if(typeof value == 'string'){
                    if(value.indexOf('%') > -1){
                      let _temp = value.replace('%','');
                      if(!isNaN(parseFloat(_temp,10))){
                          value = value.replace('%','');
                          cell.value = value / 100;
                          cell.numFmt = '0.00%';
                      }
                    }else if(!isNaN(parseFloat(value,10))){
                      cell.value = parseFloat(value,10);
                    }
                }
      },2)
      .customSetValueByIndex([[1,1],[1,2]],function(cell){
          //做和业务相关的事
      })

```



## 保存EXCEL[⬆](#目录)<!-- Link generated with jump2header -->
#### saveAsExcel  其中保存完毕后，所有的设置将会清空

```javascript

  import {createExcelByOneSheet,getExcelCellStyle} from "easyexceljs"
   
    const excelOptions = {
          excelFileName: "XX公司年度报表",
          sheetName:'本季度报表1'
    };
      //创建一个实例
    const _createExcelByOneSheet = new createExcelByOneSheet(excelOptions);
    /*
      设置header , body
      此处代码略，参照  “快速开始”
    */

   _createExcelByOneSheet.saveAsExcel()

```


## 继续开发计划[⬆](#目录)<!-- Link generated with jump2header -->
-   方便的表头合并
-   方便的列合并
-   生成在线预览Excel
-   typeScipt重写
-   编写测试用例，增加覆盖率