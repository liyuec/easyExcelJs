const ExcelJS = require('exceljs');
import {saveAs} from "file-saver";
import {getCellPosLetter,conWar,conErr,conLog,_setCellStyleByWhere,_setRowStyle,_isBasicType,_getWorkBook,getType,isObject} from '../help/function';
import {ALERT_MESSAGE} from '../help/message';
import {baseModel} from './excelDto';
import {getExcelCellStyle} from '../template/index';

/*
    只创建一个sheet的excel，以后版本提供多个sheet
    必传参数：
    sheetName: sheet的名称，别整什么很多字进去，你见过sheet 的名称上百字？;

    可选参数：没有做类型判断，就不要整那些幺蛾子东西了好吗，比如创建时间，你非要传个new Object，
                少这些验证就是少体积，即便少1Kb，也算1Kb，我们彼此相信好吗。
    creator：       创建者
    lastModifiedBy：更新时间
    created：       创建时间
    modified      : 修改时间


    rowStyleOptions:所需设置的行数样式，结构[{},{}]
                    其中
                    {
                            rowNum：需要高亮的行数index，从1开始
                            rowBgColor: 背景色
                            font:{
                                name:'',
                                size:14,
                                bold:true || false
                                color:''
                            }
                    }
    cellStyleOptions:所需设置的 列  的样式 ，结构[{},{}]
                    其中
                    {
                        cellIndex: 第几列
                        rowIndex:  第几行
                        cellName:  A1 C2 或则 AA1 等
                        BoderColor: 边框颜色
                        font:{
                            name:'',
                            size:14,
                            bold:true || false
                            color:''
                        }
                    }

*/
function createExcelByOneSheet(options){
    if(!(this instanceof createExcelByOneSheet)){
        conErr(ALERT_MESSAGE.MUST_NEW);
        return;
    }

    if(!options){
        conErr(ALERT_MESSAGE.MUST_ARGUMENTS);
        return;
    }

    baseModel.call(this,options);
    this.sheetName = options.sheetName || 'sheet1';
    this.sheetColumnsData = [];
    this.sheetRowsData = [];
    this.excelFileName = options.excelFileName || 'excel';
    /*
    {
        rowNum：需要高亮的行数index，从1开始
        rowBgColor: 背景色
        font:{
            name:'',
            size:14,
            bold:true || false
            color:''
        }
    }
    */
    this.rowStyleOptions = [];
    this.cellStyleOptions = [];
    //保存的cell设置fn
    this.setCellByWhere = [];
}

/*
    获取 column 头部的数据结构
*/
createExcelByOneSheet.prototype.getColumnBaseStructure = function(){
    return {
        header:'column Name',
        key:'column Key',
        width:20
    }
}

/*
    浏览器端得到excel
*/
createExcelByOneSheet.prototype.saveAsExcel = function(){
    _isBasicType(this);

    return new Promise((reject,resolve)=>{
        const workbook = _getWorkBook(this);
        //const worksheet = workbook.addWorksheet('oh no ,please', {properties:{tabColor:{argb:'FFC0000'}}});
        const worksheet = workbook.addWorksheet(this.sheetName);
        if(this.sheetColumnsData.length > 0){
            worksheet.columns = [
                ...this.sheetColumnsData
            ]
        }

        this.sheetRowsData.forEach(i=>{
            worksheet.addRow({...i});
        })
        
        _setRowStyle(worksheet,this.rowStyleOptions);

        //设置列样式    如果有
        if(this.setCellByWhere.length > 0){
            _setCellStyleByWhere.call(this,worksheet)
        }
        //_setCellStyle(worksheet,this.cellStyleOptions);
        
        this.excelFileName = this.excelFileName.lastIndexOf('.xlsx') > -1 ? this.excelFileName : this.excelFileName + '.xlsx';
    
        workbook.xlsx.writeBuffer().then((data => {
            const blob = new Blob([data], {type: ''});
            saveAs(blob, this.excelFileName);
        }))
    })
}


/*
    根据条件设置 哪些列需要 样式   不包含第1行(head的样式)

    若不传入cellStyle 则进行默认的样式填充，（黑边框）
    where:{
        valueKey: key Value
        whereType: < | > | == | != | === | !== | indexOf | unIndexOf
        whereValue: number | string
    }
*/
createExcelByOneSheet.prototype.setCellStyleByWhere = function(where,cellStyle = undefined){
    //如果没有设置任何样式  则默认样式
    if(!cellStyle){
        cellStyle = new getExcelCellStyle()
    }

    let _where = where;

    //判断类型  且只拿自己本身的属性
    if(isObject(_where)){
         this.setCellByWhere.push({
            where:_where,
            cellStyle:cellStyle
         })
    }else{
        conErr(ALERT_MESSAGE.OBJECT_TYPE)
        return;
    }

    return this;
}

/*
    
*/
createExcelByOneSheet.prototype.setCellStyleByRowCellIndex = function(rowCellIndex,cellStyle = undefined){
    //如果没有设置任何样式  则默认样式
    if(!cellStyle){
        cellStyle = new getExcelCellStyle()
    }



}



export default createExcelByOneSheet;