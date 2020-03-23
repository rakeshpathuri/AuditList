


 const xlsxFile = require('xlsx');
 const folder_list = ['./sold/','./purchase/']; 
 const fs = require('fs');
 const ndcConvert = require('./Util/ndcConvert');

var proerty_map = new Map()
.set('Amazing_ALPINE RX REPORT',['NDC','Quantity','no-dash','Alphine_RX'])
.set('Amazing_Kinray_OTC Report',['Universal NDC','Qty','no-dash','ndc','Kinray_OTC'])
.set('Amazing_Kinray_RX Report',['Universal NDC','Qty','no-dash','Kinray_RX']);
 var distinct_list = new Map();

 folder_list.forEach((e,index) => {      
   fs.readdir(e, (err, files) => {
      files.forEach(file => {       
        var wb = xlsxFile.readFile(e+file); 
        var ws = wb.Sheets[wb.SheetNames[0]];        
        var data = xlsxFile.utils.sheet_to_json(ws);         
        if(index ===0){
         
         var  a = data.map(r => {
            delete r.DATEF;
            delete r.INSCODE;
            delete r.RXNO;
            delete r.BRAND;
            delete r.INSNAME;
            delete r.BINNO;
            delete r.PCN;
            delete r.GROUPNO;
           
            if(!distinct_list.has(r.NDC)){
                r.AllDISP = r.QUANT
                distinct_list.set(r.NDC,r)                
            }else{              
               let c = distinct_list.get(r.NDC);
               c.QUANT = c.QUANT+r.QUANT;
               c.AllDISP  = c.AllDISP+ r.QUANT;
            }
            return r;
         });         
         
         distinct_list.forEach((v,k) =>{
            v.AllDISP = (v.AllDISP/v.PACKAGESIZE);
         });
         //  var d = [...distinct_list.values()].map(e=>{e.AllDISP = (e.AllDISP/e.PACKAGESIZE); return e;});
   
          
         } else {           
            /* const prop_name = file.split('.').slice(0, -1).join('.');           
            let properties = proerty_map.get(prop_name);                         
            data.map(r => { 
                   r[properties[0]] = ndcConvert.converttoformat(r[properties[0]]);
                   // console.log(r[properties[0]]);
                   let subRecord = distinct_list.get(r[properties[0]]);   
                   if(distinct_list.has( r[properties[0]])) {                                      
                     subRecord[properties[3]] = r[properties[1]];                        
                   }
               });   */            
        }
        
      });
    });
 });
 console.log(distinct_list);
 var newWb = xlsxFile.utils.book_new();          
 var newWs = xlsxFile.utils.json_to_sheet([...distinct_list.values()]);
 var wscols = [
   {wch:18},
   {wch:27},
   {wch:17},
   {wch:15},
   {wch:15}
];

          newWs['!cols'] = wscols;
          xlsxFile.utils.book_append_sheet(newWb,newWs,"New Data");         
          xlsxFile.writeFile(newWb,"Finla_result.xlsx");

          

          
 


