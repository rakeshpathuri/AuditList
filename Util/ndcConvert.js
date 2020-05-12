module.exports = {
    converttoformat:function(NDC){
     
      if(NDC){          
        let string =NDC.toString().trim();
        if(string.length === 11){          
            return string.substring(0,5)+'-'+string.substring(5,9)+'-'+string.substring(9,11);
          }else{
            return NDC;
          }
    }  
    return  NDC;
  }
}

   