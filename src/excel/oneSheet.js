const ExcelJS = require('exceljs');
import {saveAs} from "file-saver";
import {getCellPosLetter,conWar,conErr,conLog,_setCellStyleByWhere,_setCellByRowCellIndex,clearExcelOptions,
    _setRowStyle,_isBasicType,_getWorkBook,getType,isObject,_setCellNotes,_setCurrentValue} from '../help/function';
import {ALERT_MESSAGE} from '../help/message';
import {baseModel} from './excelDto';
import {getExcelCellStyle,getExcelCellNoteDTO} from '../template/index';

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
    //保存的cell样式设置    (setCellStyleByWhere 方法进入)
    this.setCellByWhere = [];
    //保存 cell样式设置   (setCellStyleByRowCellIndex 方法进入)
    this.setCellByRowCellIndex = [];
    //保存 cell 注解 (setCellNoteByRowCellIndex 方法进入)
    this.setCellNotesIndex = [];
    //保存 用户自定义callback 修改 cellName的值
    this.setCellByCustomIndex = [];
    this.setCellByCustomCallback = [];
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

        //根据 rowIndex,cellIndex 设置列样式
        if(this.setCellByRowCellIndex.length > 0){
            _setCellByRowCellIndex.call(this,worksheet)
        }
        //_setCellStyle(worksheet,this.cellStyleOptions);

        //根据 rowIndex,cellIndex 设置列 注解
        if(this.setCellNotesIndex.length > 0 ){
            _setCellNotes.call(this,worksheet);
        }

        //根据 rowIndex,cellIndex 和 用户自定义callBack 对 每列值进行修改
        if(this.setCellByCustomIndex.length > 0){
            _setCurrentValue.call(this,worksheet);
        }
        
        
        this.excelFileName = this.excelFileName.lastIndexOf('.xlsx') > -1 ? this.excelFileName : this.excelFileName + '.xlsx';
        workbook.xlsx.writeBuffer().then((data => {
            const blob = new Blob([data], {type: ''});
            saveAs(blob, this.excelFileName);
            clearExcelOptions.call(this)
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
    通过cell和row的索引设置列样式   rowCellIndex数据结构 = [[rowIndex,cellIndex],[rowIndex,cellIndex]]
*/
createExcelByOneSheet.prototype.setCellStyleByRowCellIndex = function(rowCellIndex,cellStyle = undefined){
    //如果没有设置任何样式  则默认样式
    if(!cellStyle){
        cellStyle = new getExcelCellStyle();
    }

    if(rowCellIndex.constructor === Array){
        rowCellIndex.forEach(i=>{
            this.setCellByRowCellIndex.push({
                ROW_CELL_INDEX:i,
                cellStyle:cellStyle
            })
        })
    }else{
        conErr(ALERT_MESSAGE.ROWCELL_INDEX_TYPE)
        return;
    }

    return this;

}


/*
    通过cell和row的索引设置列 注解  
    rowCellIndex数据结构 = [[rowIndex,cellIndex],[rowIndex,cellIndex]]
    noteTexts : [string,string] 
*/
createExcelByOneSheet.prototype.setCellNoteTextByRowCellIndex = function(rowCellIndex,noteTexts){
    if(rowCellIndex.constructor === Array || noteDatas.constructor === Array){
        rowCellIndex.forEach((i,index)=>{
            let _getExcelCellNoteDTO = new getExcelCellNoteDTO();
            _getExcelCellNoteDTO.text = noteTexts[index]
            this.setCellNotesIndex.push({
                ROW_CELL_INDEX:i,
                NOTE_DTO:_getExcelCellNoteDTO
            })
        })
       
    }else{
        conErr(ALERT_MESSAGE.ROWCELL_INDEX_TYPE)
        conErr(ALERT_MESSAGE.NOTES_INDEX_TYPE)
        return;
    }

    return this;
}

/*
    根据行数,列数   找到value   根据用户自定义function   对value进行修改
    rowCellIndex数据结构 = [[rowIndex,cellIndex],[rowIndex,cellIndex]]
    callBack  === function  自定义函数进行处理      callback可操作this范围的值
    不要使用箭头函数，因为箭头函数无法call到当前this

*/
createExcelByOneSheet.prototype.customSetValueByIndex = function(rowCellIndex,callBack){
    if(getType(callBack) !== 'Function'){
        conErr(ALERT_MESSAGE.MUST_FUNCTION);
        return;
    }

    rowCellIndex.forEach((i,index) => {
        this.setCellByCustomIndex.push(i)
    })

    this.setCellByCustomCallback.push(callBack);
    return this;

}

export default createExcelByOneSheet;