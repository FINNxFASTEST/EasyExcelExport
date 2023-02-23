
function ExportExcel( column , data , file_name ){
    try{
        if(data && column){
            var woorkbook = XLSX.utils.book_new();
            var ws_data = [];
            var ws; 
            let tmp_colum = column.map((r)=>{return r.name})
            var column_name =[tmp_colum];
            var excelData = {name:"sheet1",row:column_name};

            data.map((list,index)=>{
                    let tmp_data = [];
                    column.map((tmp_column_name, c_index)=>{
                        let output_text = '';
                        if(!tmp_column_name.type){
                            tmp_column_name.type = 'text'  // default type
                        } 
                        let path_split = tmp_column_name.data_index.split('.');

                        let oject_nest_value = list

                        // ------ this for nested oject key.  ex : key =  price_vat.$numberDecimal 
                        for ( const path  of path_split ){
                            if(path in oject_nest_value){
                                oject_nest_value = oject_nest_value[path]
                            }
                        }
                        if(typeof oject_nest_value == 'array' || typeof oject_nest_value == 'object') throw 'you should process data before export'
                        switch (tmp_column_name.type){
                            case 'date':
                                output_text = new Date(oject_nest_value);
                                break;
                            case 'text':
                                output_text = oject_nest_value;
                                break;
                            case 'money':
                                output_text = new Intl.NumberFormat().format(oject_nest_value)
                                break;
                        }
                        tmp_data.push(output_text)
                    })
                    excelData.row.push(tmp_data);
            })

            woorkbook.SheetNames.push("sheet1");             // <-----  create sheet name
            ws = XLSX.utils.aoa_to_sheet(excelData.row);
            woorkbook.Sheets["sheet1"] = ws;                 
            var wopts = { bookType:'xlsx', bookSST:false, type:'array' };
            var wbout = XLSX.write(woorkbook,wopts);
            ManualDownload(new Blob([wbout],{type:"application/octet-stream"}), file_name)
        }
        
    }catch(err){
        console.error(err);
        return false;
    }
}
function ManualDownload(file , file_name){
    const a = document.createElement('a')
    url = window.URL.createObjectURL(file);
    a.href = url
    a.download = `${file_name}.xlsx` ;
    a.click();
    a.remove();
}

function ExampleCall(){
    var timeCode = moment().format('LL')
    let file_name = `Excample Name${timeCode}`
    const column = [
        {
            name       : 'เวลา',
            data_index : 'created_at',
            type       : 'date',
        },
        {
            name       : 'id',
            data_index : 'id',
            type       : 'text',
        },
        {
            name       : 'invoice',
            data_index : 'invoice',
            type       : 'text',
        },
        {
            name       : 'Name',
            data_index : 'patient',
            type       : 'text',
        },
        {
            name       : 'Phone',
            data_index : 'patient_contact',
            type       : 'text',
        },
        {
            name       : 'list',
            data_index : 'list_name',
            type       : 'text',
        },
        {
            name       : 'price',
            data_index : 'price_total',
            type       : 'money',
        },
        {
            name       : 'overdue',
            data_index : 'price_balance',
            type       : 'money',
        }
    ]
    let excelData = [
        {
             'created_at'       :   "2023-02-23T04:32:44.239Z"
            ,'id'               :   "016500002"
            ,'invoice'          :   "660064835"
            ,'patient'          :   "EXAMPLE"
            ,'list_name'        :   "EXAMPLE LIST"
            ,'patient_contact'  :   '094444444'
            ,'price_total'      :   '300'
            ,'price_balance'    :   '100'
        },
        {
            'created_at'       :   "2023-01-23T04:32:44.239Z"
           ,'id'               :   "016500302"
           ,'invoice'          :   "660023835"
           ,'patient'          :   "ตัวอย่าง จ้า"
           ,'list_name'        :   "ขนม"
           ,'patient_contact'  :   '035545334'
           ,'price_total'      :   '300'
           ,'price_balance'    :   '100'
        }

    ]

   ExportExcel(column, excelData , file_name);


}
