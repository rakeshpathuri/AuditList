module.exports = {
    converttoformat:function(NDC){
      if(NDC){          
        let string =NDC.toString();
        if(string.length === 11){
            return NDC.substring(0,5)+'-'+NDC.substring(5,9)+'-'+NDC.substring(9,11);
          }else{
            return NDC;
          }
    }  
    return  NDC;
  }
}

   