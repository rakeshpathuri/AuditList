const xlsxFile = require('xlsx');
const fs = require('fs');
const Excel = require('exceljs');
const ndcConvert = require('./Util/ndcConvert');
let proerty_map = require('./Util/RuleBook');
const { from  } = require('rxjs'); 
const { map, concatMap,last} = require('rxjs/operators');

const folder_list = ['./sold/','./purchase/']; 
const distinct_list = new Map();
proerty_map =proerty_map.proerty_map;
let folder_count = 0;
const file_names = [];
const added_prop_list =[];
const binData = new Map();
const binFilterData = new Map();

const walkSync = function(dir) {    
    files = fs.readdirSync(dir);
    files = [...files].map(e=> {file_names.push(e); return dir.concat(e);});  
    return from(files);   
 }

 const binRead = function(dir) {    
  files = fs.readdirSync(dir);
  files = [...files].map(e=> {return dir.concat(e);});  
  return from(files);   
}

function readFile(filename ){
   return new Promise(function(resolve, reject) {
      const prop_name = proerty_map.get(file_names[folder_count].split('.').slice(0, -1).join('.'));
      let wb = xlsxFile.readFile(filename); 
      let ws = wb.Sheets[wb.SheetNames[0]];     
     
      if(folder_count >0 && prop_name[4] != 0){           
        wb.SheetNames.push("Test Sheet");      
        const at = xlsxFile.utils.sheet_to_json(ws, {header:1});        
        const w = xlsxFile.utils.aoa_to_sheet(at.splice(prop_name[4]));  
        wb.Sheets["Test Sheet"] = w; 
        let s = wb.Sheets[wb.SheetNames[1]];     
        resolve(s);          
      }     
      resolve(ws);
   });   
}

async function convertoJson(filename){
       let a = await(readFile(filename));
       let data = xlsxFile.utils.sheet_to_json(a);  
       return [...data];
 }

 function readDataFromBin(filename){
  let wb = xlsxFile.readFile(filename); 
  let ws = wb.Sheets[wb.SheetNames[0]]; 
  let data = xlsxFile.utils.sheet_to_json(ws);  
  const file = filename.split('.').slice(0, -1).join('.').split("/").pop();
  const d = [...data].map(e=> e.BinNumber); 
  binData.set(file,d); 
  binFilterData.set(file,new Array());
  return [...data]
}

 function refineData(data) {   
   
    if(folder_count === 0){ 
    data.map(r=> {return {NDC:r.NDC,DRGNAME:r.DRGNAME,DRUGSTRONG:r.DRUGSTRONG,PACKAGESIZE:r.PACKAGESIZE,QUANT:r.QUANT,BINNO:r.BINNO};}).map((r,index) => {
          
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
      if(v.PACKAGESIZE >0){
      v.AllDISP = (v.AllDISP/v.PACKAGESIZE);
      } else{
        v.AllDISP = 0;
      }
    });
   
  } else {      
    
   const prop_name = file_names[folder_count].split('.').slice(0, -1).join('.');           
   let properties = proerty_map.get(prop_name);        
   let testObject =new Map();
   
   data = data.map(res =>{   
    if(testObject.has(res[properties[0]])) {      
        let testdata = testObject.get(res[properties[0]]);         
        testdata[properties[1]] =  testdata[properties[1]]+res[properties[1]];         
    }else{
      testObject.set(res[properties[0]],res);
    }
   });

   [...testObject.values()].map(r => { 
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

 function generateXl(e){   
    
   distinct_list.forEach((v,k) =>{         
      let totalpurchased = 0;
      added_prop_list.forEach(e=>{                     
           totalpurchased = totalpurchased +  v[e];
      });    
      v.totalpurchased = totalpurchased;      
    });

    distinct_list.forEach((v,k,index) =>{        
    v.distace = (parseFloat(v.totalpurchased) - parseFloat(v.AllDISP));
    });
    
   let sort_list = [...distinct_list.values()].sort(vsort); 
 
   for (let r of sort_list) {  
    for (let [k, v] of binData) {      
       if(v.indexOf(r.BINNO) > 0){                
        let rr = binFilterData.get(k);
        rr.push(r); 
            
       }
    }
   }  
   const workbook = new Excel.Workbook();
   let worksheet = workbook.addWorksheet('AllDisp');
   let a = sort_list[0];
   a = Object.keys(a); 

 
   worksheet.columns = a.map(r=>{  return {header:r,key:r};});
   worksheet.addRows(sort_list.map(r=> Object.values(r)));

   binFilterData.forEach((v,k)=>{       
    let worksheet = workbook.addWorksheet(k);
    worksheet.columns = a.map(r=>{  return {header:r,key:r};});
    worksheet.addRows(v.map(r=> Object.values(r)));   
    
   worksheet.getRow(1).eachCell((cell) => {
    cell.font = { bold: true };
    cell.font = {color: {argb: "004e47cc"}};
    
  });

  a.map((e,i)=>worksheet.getColumn(i+1).width =20);  
  worksheet.autoFilter = 'A1:'+String.fromCharCode(65+a.length-1)+'1';
  worksheet.views = [{ state: 'frozen',  ySplit: 1, activeCell: 'B2' },];
   });

   worksheet.getRow(1).eachCell((cell) => {
    cell.font = { bold: true };
    cell.font = {color: {argb: "004e47cc"}};
    
  });

  a.map((e,i)=>worksheet.getColumn(i+1).width =20);  
  worksheet.autoFilter = 'A1:'+String.fromCharCode(65+a.length-1)+'1';
  worksheet.views = [{ state: 'frozen',  ySplit: 1, activeCell: 'B2' },];
  
const fileName = 'FinalResult_'+new Date().getTime();

// save workbook to disk
workbook
  .xlsx
  .writeFile(fileName+'.xlsx')
  .then(() => {
    console.log("saved");
  })
  .catch((err) => {
    console.log("err", err);
  });
  
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

 from(folder_list).pipe(concatMap(e=> walkSync(e))).pipe(concatMap(e=> convertoJson(e))).pipe(map(e=> refineData(e))).pipe(last()).subscribe(e=> generateXl(e));  
 from(['./Bin Number/']).pipe(concatMap(e=> binRead(e))).pipe(concatMap(e=> readDataFromBin(e))).subscribe();
