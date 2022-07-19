function baseModel(options){
    this.creator = options.creator || '默认用户';
    this.lastModifiedBy = options.lastModifiedBy || '';
    this.created = options.created || new Date();
    this.modified = options.modified || new Date();
    this.lastPrinted = options.lastPrinted || new Date();
    this.version = '1.1.0';
}


/*
    列样式的DTO
*/
function CellStyleDTO(){
    let obj = new Object(
        {
            cellIndex:1,
            rowIndex:1,
            cellName:'',
            BoderColor:'',
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



export {
    baseModel,CellStyleDTO
}