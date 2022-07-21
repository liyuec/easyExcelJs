import {CellStyleDTO} from '../excel/excelDto'

/*
    argba 在线颜色 转换 https://sunpma.com/other/rgb/
*/
const ExcelStyleTemplate = {
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
            BoderColor: 'FFFF0000',
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
            BoderColor: 'ff48c9b0',
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
            BoderColor: 'ffff0000',
            font:{
                name:'Arial',
                size:11,
                bold:true,
                color:'ff707b7c'
            }
        }
    }
}

const getExcelCellStyle = function(colorTemplate){
    var cellStyle = new CellStyleDTO();
    switch(colorTemplate){
        case "red":
            cellStyle.BoderColor = 'ffff0000'
            cellStyle.font = {
                name:'Malgun Gothic Semilight',
                size:11,
                bold:true,
                color:'ffff0000'
            }
        break;
        case "blue":
            cellStyle.BoderColor = 'ff5faee3'
            cellStyle.font = {
                name:'Malgun Gothic Semilight',
                size:11,
                bold:true,
                color:'ff5faee3'
            }
        break;
        case "green":
            cellStyle.BoderColor = 'ff48c9b0'
            cellStyle.font = {
                name:'Malgun Gothic Semilight',
                size:11,
                bold:true,
                color:'ff48c9b0'
            }
        break;
        default:
            cellStyle.BoderColor = ''
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


export {
    ExcelStyleTemplate,getExcelCellStyle
}