const ExcelJS = require('exceljs');
import {saveAs} from "file-saver";

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
        if(i.BoderColor){
            worksheet.getCell(_cellName).border = {
                top: {style:'thick', color: {argb:i.BoderColor}},
                left: {style:'thick', color: {argb:i.BoderColor}},
                bottom: {style:'thick', color: {argb:i.BoderColor}},
                right: {style:'thick', color: {argb:i.BoderColor}}
            };
        }
        if(i.font){
            worksheet.getCell(_cellName).font = {
                name: i.font.name || 'Arial Black',
                color: { argb: i.font.color || 'FF00FF00' },
                family: 2,
                size: i.font.size || 18
            };
        }
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



export {
    getCellPosLetter,conWar,conErr,conLog,
    _setCellStyle,_setRowStyle,_isBasicType,_getWorkBook,getType
}