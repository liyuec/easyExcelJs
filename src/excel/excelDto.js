function baseModel(options){
    this.creator = options.creator || '默认用户';
    this.lastModifiedBy = options.lastModifiedBy || '';
    this.created = options.created || new Date();
    this.modified = options.modified || new Date();
    this.lastPrinted = options.lastPrinted || new Date();
    this.version = '1.1.6';
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

/*
    注解实体
*/
function CellNoteDTO(){
    let obj = new Object(
        {
            text:'',
            protection: {
                locked: true,
                lockText: false
            },
            //twoCells  oneCells    absolute
            editAs: 'absolute'
        }
    )

    return obj;
}


/*
    目前Babel6  不支持 #    难得换7
    customSetValueByIndex   的链表结构NODE
*/
class customNode{
  /*   next;
    value; */
    constructor(nodeObj){
        this.next = undefined;
        this.value = nodeObj
    }
    
}
/*
    目前Babel6  不支持 #        难得换7
    customSetValueByIndex   的链表结构 NodeList
*/
class customSetNodeList{
   /*  #node = null
    #size = 0
    #tail = null
    #head = null */
    constructor(){
        this.node = null
        this.size = 0
        this.tail = null
        this.head = null
        this.clear()
    }

    get sizes(){
        return this.size;
    }

    addNew(nodeObj){
        this.node = new customNode(nodeObj)
        if(this.head){
            this.tail.next = this.node;
            this.tail = this.node;
        }else{
            this.tail = this.node;
            this.head = this.node;
        }
        this.size++;
    }

    get getHead(){
        return this.head;
    }

    clear(){
        this.tail = null;
        this.head = null;
        this.size = 0;
    }
}

export {
    baseModel,CellStyleDTO,CellNoteDTO,customSetNodeList
}