const ExcelJS = require('exceljs');
import {saveAs} from "file-saver";
import {getCellPosLetter,conWar,conErr,conLog,_setCellStyleByWhere,_setCellByRowCellIndex,clearExcelOptions,
    _setRowStyle,_isBasicType,_getWorkBook,getType,isObject,_setCellNotes,_setCurrentValue,_mergeCells,_alignmentCells,_setRichText,_setRowsHeight} from '../help/function';
import {ALERT_MESSAGE} from '../help/message';
import {baseModel,customSetNodeList} from './excelDto';
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
    this.customList = new customSetNodeList();
    //可能需要补充的rowLength
    this.repairLength = 0;
    //合并单元格的list
    this.mergeCellsList = [];
    //居中，缩进的list
    this.alignmentList = [];
    //行高集合
    this.rowsHeightList = [];
    //富文本
    this.richTextList = [];
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

        //合并单元格
        if(this.mergeCellsList.length > 0){
            _mergeCells.call(this,worksheet);
        }
         
        //缩进,居中
        if(this.alignmentList.length > 0){
            _alignmentCells.call(this,worksheet);
        }

        //富文本
        if(this.richTextList.length > 0){
            _setRichText.call(this,worksheet);
        }

        //行高设置
        if(this.rowsHeightList.length > 0){
            _setRowsHeight.call(this,worksheet);
        }

        //根据 rowIndex,cellIndex 和 用户自定义callBack 对 每列值进行修改
        if(this.customList.sizes > 0){
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

    一般用来整体设置样式
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
    柯里化
    合并单元格
    按照 'A4:B5'传入 
    暂不支持   按开始行，开始列，结束行，结束列合并（相当于 K10:M12）  worksheet.mergeCells(10,11,12,13);
*/
createExcelByOneSheet.prototype.mergeCells = function(cellNames1){
    if(cellNames1 === void 0){
        return this;
    }
    let _super = this;
    this.mergeCellsList.push(cellNames1)
    let merFunc = function(cellNames2){
        if(cellNames2 === void 0){
            return _super;
        }else{
            if(getType(cellNames2) === 'String'){
                _super.mergeCellsList.push(cellNames2)
            }

            return merFunc;
        }
    }
    return merFunc;
}

/*
    柯里化
    单元格居中，缩进
    {
        cellName:'A1',
        alignment:{}
    }
*/
createExcelByOneSheet.prototype.alignmentCells = function(cellAligenObj){
    if(cellAligenObj === void 0){
        return this;
    }
    this.alignmentList.push(cellAligenObj);
    let _super = this;
    let alignmentFunc = function(cellAligenObj2){
        if(cellAligenObj2 === void 0){
            return _super;
        }else{
            if(getType(cellAligenObj2) === 'Object'){
                _super.alignmentList.push(cellAligenObj2)
            }
            return alignmentFunc;
        }
    }
    return alignmentFunc;
}


/*
    柯里化
    设置富文本
    {
        cellName:'A1',
        richText:[]
    }
*/
createExcelByOneSheet.prototype.RichTextCells = function(richTextObj){
    if(richTextObj === void 0){
        return this;
    }
    this.richTextList.push(richTextObj)
    let _super = this;
    let richTextFunc = function(richTextObj2){
        if(richTextObj2 === void 0){
            return _super;
        }else{
            if(getType(richTextObj2) === 'Object'){
                _super.richTextList.push(richTextObj2)
            }

            return richTextFunc;
        }
    }
    return richTextFunc;
}

/*
    柯里化
     设置行高
    {
        rowIndex:1,
        height:
    }
*/
createExcelByOneSheet.prototype.rowsHeight = function(rowObj){
  
    if(rowObj === void 0){
        return this;
    }
    this.rowsHeightList.push(rowObj)
    let _super = this;
   
    let rowsHeightFunc = function(rowObj2){
        if(rowObj2 === void 0){
            return _super;
        }else{
            if(getType(rowObj2) === 'Object'){
                _super.rowsHeightList.push(rowObj2)
            }

            return rowsHeightFunc;
        }
    }
    return rowsHeightFunc;
}



/*
    根据行数,列数   找到value   根据用户自定义function   对value进行修改
    rowCellIndex数据结构 = [[rowIndex,cellIndex],[rowIndex,cellIndex]]      
    rowCellIndex:   '*'  则表示  除了头部的全部行和列
    callBack  === function  自定义函数进行处理      callback可操作this范围的值
    不要使用箭头函数，因为箭头函数无法call到当前this


    repairLength:需要补充的row length 遍历
*/
createExcelByOneSheet.prototype.customSetValueByIndex = function(rowCellIndex,callBack,repairLength = 0){
    if(getType(callBack) !== 'Function'){
        conErr(ALERT_MESSAGE.MUST_FUNCTION);
        return;
    }

    let nodeObj = {
        rowCellIndex:[],
        callBack:undefined
    }
    // 42 === *
    if(rowCellIndex.constructor !== Array && rowCellIndex.charCodeAt(0) === 42){
        nodeObj.rowCellIndex = rowCellIndex
    }else if(rowCellIndex.constructor === Array){
        rowCellIndex.forEach((i,index) => {
            nodeObj.rowCellIndex.push(i)
        })
    }
 
    nodeObj.callBack = callBack;
    this.repairLength = repairLength;
    this.customList.addNew(nodeObj)
    return this;
}

export default createExcelByOneSheet;