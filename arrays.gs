function myFunction() {
  var newarray = [
                    ["one", 1],
                    ["two", 2]
                  ];
  newarray.push(["New",3]);
  
  //Logger.log(newarray)
  
  for(i=0; i<=10; i++){
    
    newarray.push([i]);
    //Logger.log(newarray)
  }
  
  total = 2700;
  
  if(total>900){
      numOfPallet = (total -(total%900))/900;
      //Logger.log(numOfPallet);
    }
  total = (total + numOfPallet*15 );
  //Logger.log(total)
  
  var settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  
  settingsSheet.getRange(10, 1).setNote("My NOTE");
  
  settingsSheet.getRange(10, 1).clear();
  settingsSheet.clearContents();
  
}
