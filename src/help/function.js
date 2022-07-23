const ExcelJS = require('exceljs');

//import {saveAs} from "file-saver";

/*
    根据 第几列  得到 列数的A-Z   比如 1 得到 A    27 得到AA
    第几行 

    最终返回  A1  OR  AA1
*/
function getCellPosLetter(cellIndex,rowIndex){
    const max = Math.pow(2,31);
    if(cellIndex >= max || rowIndex >= max){
        conWar(ALERT_MESSAGE.MAX_INTEGER)
        return;
    }

	if(cellIndex <= 0){
		return 'A1';
	}
	let result = '',
	_ACharCode = 'A'.charCodeAt();
	while(cellIndex > 0){
		cellIndex--;
		result = String.fromCharCode(cellIndex % 26 + _ACharCode) + result;
		cellIndex = ((cellIndex - cellIndex % 26) / 26);
	}
	
	return result + rowIndex
}



const _toString = Object.prototype.toString;

function _getConsole(type){
    switch(type){
        case 'log':
            return function(msg){
                console.log(msg)
            }
        break;
        case 'warn':
            return function(msg){
                console.warn(msg)
            }
        break;
        case 'error':
            return function(msg){
                console.error(msg)
            }
        break;
    }
}

const conWar = _getConsole('warn')
const conErr = _getConsole('error')
const conLog = _getConsole('log')





function _setCellStyle(worksheet,cellStyleOptions){
    if(!cellStyleOptions.length > 0 ){
        return;
    }

    cellStyleOptions.forEach(i => {
        let _cellName = '';
        
        _cellName = i.cellName ? i.cellName : getCellPosLetter(i.cellIndex || 1,i.rowIndex || 1);
        if(i.BorderColor){
            worksheet.getCell(_cellName).border = {
               /*  top: {style:'thick', color: {argb:i.BorderColor}},
                left: {style:'thick', color: {argb:i.BorderColor}},
                bottom: {style:'thick', color: {argb:i.BorderColor}},
                right: {style:'thick', color: {argb:i.BorderColor}} */
                top: {style:i.BorderStyle ? i.BorderStyle : '', color: {argb:i.BorderColor}},
                left: {style:i.BorderStyle ? i.BorderStyle : '', color: {argb:i.BorderColor}},
                bottom: {style:i.BorderStyle ? i.BorderStyle : '', color: {argb:i.BorderColor}},
                right: {style:i.BorderStyle ? i.BorderStyle : '', color: {argb:i.BorderColor}}
            };
        }
        if(i.font){
            worksheet.getCell(_cellName).font = {
                name: i.font.name || '',
                color: { argb: i.font.color || '' },
                family: 2,
                size: i.font.size || 11,
                bold: typeof (i.font.bold) == 'boolean' ? i.font.bold : ''
                //bold:true
            };
        }
    })
}

/*
    单独设置 excel cell的注解  ，后期扩展富文本
*/
function _setCellNote(worksheet,rowIndex,cellIndex,noteDTO,poolText = true){
    let _cellName = getCellPosLetter(cellIndex|| 1,rowIndex || 1),
    _note = {
        protection:{
            locked: noteDTO.protection.locked,
            lockText: noteDTO.protection.lockText
        },
        editAs:noteDTO.editAs
    }
    if(poolText){
        _note.texts = [
            {'font': {'bold': true, 'size': 11, 'color': {'argb': 'ff696969','theme': 1}, 'name': '宋体', 'family': 2, 'scheme': 'minor'}, 'text': noteDTO.text}
        ]
    }else{
        //富文本text
    }

    worksheet.getCell(_cellName).note = _note;
}
/*
    根据rowindex,cellindex   设置列 注解
    rowCellIndex数据结构 = [[rowIndex,cellIndex],[rowIndex,cellIndex]]
    noteDatas : [NOTE_DTO,NOTE_DTO]
    note:{} 
*/
function _setCellNotes(worksheet){
    this.setCellNotesIndex.forEach(i=>{
        let [rowIndex,cellIndex] = i.ROW_CELL_INDEX,
        noteDTO = i.NOTE_DTO

        _setCellNote(worksheet,rowIndex,cellIndex,noteDTO);
    })
}

//通过 rowIndex ,cellIndex 设置样式
function _setCellByRowCellIndex(worksheet){
    this.setCellByRowCellIndex.forEach(i => {
        // i 就是 ROW_CELL_INDEX 和 cellStyle == [rowIndex,cellIndex];
        let [rowIndex,cellIndex] = i.ROW_CELL_INDEX,
        _cellStyle =  i.cellStyle;

        _cellStyle.cellIndex = cellIndex || 1;
        _cellStyle.rowIndex = rowIndex || 1;
        _setCellStyle(worksheet,[_cellStyle])

    })
}

//通过where条件设置 cell样式
function _setCellStyleByWhere(worksheet){
    this.setCellByWhere.forEach(i=>{
        let {where,cellStyle} = i;
        let {valueKey,whereType,whereValue} = where;

        let cellIndex = 1,rowIndexArr = [];

        //找列数
        for(let i =0;i<this.sheetColumnsData.length;i++){
            //console.log('setCell:',this.sheetColumnsData[i],valueKey)
            if(this.sheetColumnsData[i]['key'] == valueKey){
                cellIndex += i;
                break;
            }
        }

        //找行数
        this.sheetRowsData.forEach((rowI,rowIndex) =>{
            let _v = rowI[valueKey],
            _rowIndex = 0;
            switch(whereType.trim().toLowerCase()){
                case "<":
                    _rowIndex = _v < whereValue ? rowIndex + 1 : 0
                break;
                case ">":
                    _rowIndex = _v > whereValue ? rowIndex + 1 : 0
                break;
                case "==":
                    _rowIndex = _v == whereValue ? rowIndex + 1 : 0
                break;
                case "!=":
                    _rowIndex = _v != whereValue ? rowIndex + 1 : 0
                break;
                case "===":
                    _rowIndex = _v === whereValue ? rowIndex + 1 : 0
                break;
                case "!==":
                    _rowIndex = _v !== whereValue ? rowIndex + 1 : 0
                break;
                case "indexof":
                    _rowIndex = _v.indexof(whereValue) > -1 ? rowIndex + 1 : 0
                break;
                case "unindexof":
                    _rowIndex = _v.indexof(whereValue) === -1 ? rowIndex + 1 : 0
                break;
            }
            if(_rowIndex > 0){
                rowIndexArr.push(_rowIndex);
            }
        })

        //当前where条件下的 setCell 样式
        //console.log('rowIndexArr:',rowIndexArr,cellIndex)
        rowIndexArr.forEach(rowIndex =>{
            cellStyle.cellIndex = cellIndex;
            cellStyle.rowIndex = rowIndex;
            _setCellStyle(worksheet,[cellStyle])
        })
        
    })
}

function _setRowStyle(worksheet,rowStyleOptions){
    if(!rowStyleOptions.length > 0){
        return;
    }
    rowStyleOptions.forEach(i=>{
        const row = worksheet.getRow(i.rowNum);
        row.height = i.height || 21.5;
        row.fill = {
            type: 'gradient',
            gradient: 'angle',
            degree: 0,
            stops: [
                {position:0, color:{argb:i.rowBgColor}},
                {position:0.5, color:{argb:i.rowBgColor}},
                {position:1, color:{argb:i.rowBgColor}}
            ],
        };
        row.font = { 
            name: i.font.name, size: i.font.size, bold: i.font.bold ,color:{argb:i.font.color}
        }

        row.commit();
    })
}

//根据callback 修改当前 cellName的值
function _setCurrentValue(worksheet){
    let callback = this.setCellByCustomCallback[0];

    this.setCellByCustomIndex.forEach(i=>{
        let [cellIndex,rowIndex] = i,
        _cellName = getCellPosLetter(cellIndex|| 1,rowIndex || 1);

        let cell = worksheet.getCell(_cellName)

        callback.call(this,cell)
    }) 
}


function _isBasicType(wr){
    if(getType(wr.sheetColumnsData) !== 'Array'){
        conErr(ALERT_MESSAGE.MUST_COLUMN_TYPE);
        return;
    }
    if(getType(wr.sheetRowsData) !== 'Array'){
        conErr(ALERT_MESSAGE.MUST_ROW_TYPE);
        return;
    }
    if(getType(wr.excelFileName) !== 'String'){
        conErr(ALERT_MESSAGE.MUST_FILENAME);
        return;
    }
}

function _getWorkBook(wr){
    const workbook = new ExcelJS.Workbook();
    workbook.creator = wr.creator
    workbook.lastModifiedBy = wr.lastModifiedBy
    workbook.created =  wr.created
    workbook.modified = wr.modified
    workbook.lastPrinted =  wr.lastPrinted
    workbook.properties.date1904 = true;
    workbook.calcProperties.fullCalcOnLoad = true;
    return workbook
}

function getType(target){
    return _toString.call(target).slice(8,-1);
}

function isObject(target){
    return target !== null && _toString.call(target).slice(8,-1) === 'Object'
}

function clearExcelOptions(){
    this.sheetColumnsData = [];
    this.sheetRowsData = [];
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


export {
    getCellPosLetter,conWar,conErr,conLog,
    _setCellStyle,_setRowStyle,_setCellStyleByWhere,_setCellByRowCellIndex,_setCellNotes,_setCurrentValue,
    _isBasicType,_getWorkBook,getType,clearExcelOptions,
    isObject
}