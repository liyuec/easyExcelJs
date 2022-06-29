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

export {
    ExcelStyleTemplate
}