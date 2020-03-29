const xlsxFile = require('xlsx');
const fs = require('fs');
const ndcConvert = require('./Util/ndcConvert');

const { of,from  } = require('rxjs'); 
const { map, filter ,concatMap} = require('rxjs/operators');

const folder_list = ['./sold/','./purchase/']; 
const distinct_list = new Map();
var proerty_map = new Map()
.set('Amazing_ALPINE RX REPORT',['NDC','Quantity','no-dash','Alphine_RX',4])
.set('Amazing_Kinray_OTC Report',['Universal NDC','Qty','no-dash','Kinray_OTC',8])
.set('Amazing_Kinray_RX Report',['Universal NDC','Qty','no-dash','Kinray_RX',8])
.set('SARATOGA RX LLC',['NDC','Quantity Shipped','dash','SARATOGA_RX',0]);
let folder_count = 0;
const file_names = [];
const added_prop_list =[];

var walkSync = function(dir) {    
    files = fs.readdirSync(dir);
    files = [...files].map(e=> {file_names.push(e); return dir.concat(e);});  
    return from(files);   
 }

function readFile(filename ){
   return new Promise(function(resolve, reject) {
      const prop_name = proerty_map.get(file_names[folder_count].split('.').slice(0, -1).join('.'));
      var wb = xlsxFile.readFile(filename); 
      var ws = wb.Sheets[wb.SheetNames[0]];     
     
      if(folder_count >0 && prop_name[4] != 0){           
        wb.SheetNames.push("Test Sheet");      
        const at = xlsxFile.utils.sheet_to_json(ws, {header:1});        
        const w = xlsxFile.utils.aoa_to_sheet(at.splice(prop_name[4]));  
        wb.Sheets["Test Sheet"] = w; 
        var s = wb.Sheets[wb.SheetNames[1]];     
        resolve(s);          
      }     
      resolve(ws);
   });   
}

async function convertoJson(filename){
       let a = await(readFile(filename));
       var data = xlsxFile.utils.sheet_to_json(a);     
       return [...data];
 }

 function refineData(data) {   
   
    if(folder_count === 0){ 
    data.map((r,index) => {
      delete r.DATEF;
      delete r.INSCODE;
      delete r.RXNO;
      delete r.BRAND;
      delete r.INSNAME;
      delete r.BINNO;
      delete r.PCN;
      delete r.GROUPNO;      
      for(let i= 1;i<file_names.length;i++){      
         const prop_name = proerty_map.get(file_names[i].split('.').slice(0, -1).join('.')); 
         r[prop_name[3]] = 0;
         if(index == 0){
         added_prop_list.push([prop_name[3]]);
         }         
      }     
      if(!distinct_list.has(r.NDC)){
          r.AllDISP = r.QUANT
          distinct_list.set(r.NDC,r)                
      }else{              
         let c = distinct_list.get(r.NDC);
         c.QUANT = c.QUANT+r.QUANT;
         c.AllDISP  = c.AllDISP+ r.QUANT;
      }     
    });         
   
    distinct_list.forEach((v,k) =>{
      v.AllDISP = (v.AllDISP/v.PACKAGESIZE);
    });
   
  } else {      
   const prop_name = file_names[folder_count].split('.').slice(0, -1).join('.');           
   let properties = proerty_map.get(prop_name);         
   data.map(r => { 
      if(properties[2] == 'no-dash'){
         r[properties[0]] = ndcConvert.converttoformat(r[properties[0]]);  
      }                     
          if(distinct_list.has(r[properties[0]])) {   
            let subRecord = distinct_list.get(r[properties[0]]);                                             
            subRecord[properties[3]] = r[properties[1]] ;                        
          }
      });       
  }
  folder_count++;  
 }

 function generateXl(){   
   distinct_list.forEach((v,k) =>{     
      const keySize = Object.keys(v).length;
      const  keys = Object.keys( v );
      let totalpurchased = 0;
      added_prop_list.forEach(e=>{                     
           totalpurchased = totalpurchased +  v[e];
      });    
      v.totalpurchased = totalpurchased;      
    });

    distinct_list.forEach((v,k,index) =>{        
    v.distace = (parseFloat(v.totalpurchased) - parseFloat(v.AllDISP));
    });
    
    var sort_list = [...distinct_list.values()].sort(vsort);
   const newWb = xlsxFile.utils.book_new();          
   const  newWs = xlsxFile.utils.json_to_sheet(sort_list);
   const wscols = [
     {wch:18},
     {wch:27},
     {wch:17},
     {wch:15},
     {wch:15}
  ];
   newWs['!cols'] = wscols;
   xlsxFile.utils.book_append_sheet(newWb,newWs,"New Data");         
   xlsxFile.writeFile(newWb,"Finla_result.xlsx");
 }


 function vsort(a,b){
   let comparison = 0;
   const bandA = a.distace;
   const bandB = b.distace;
   if (bandA > bandB) {
     comparison = 1;
   } else if (bandA < bandB) {
     comparison = -1;
   }
   return comparison;
 }

 from(folder_list).pipe(concatMap(e=> walkSync(e))).pipe(concatMap(e=> convertoJson(e))).pipe(map(e=> refineData(e))).subscribe(r=>generateXl());