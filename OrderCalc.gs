//GLOBAL VAR
var orderSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('modified_tab_to_past');
var settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
var resultsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Results');
 

function mainRun(){
  resultsSheet.clearContents();
  orderArr = getOrdersList();  
  totalWeightList = countKGandValume(orderArr); //get Array with total  weight
  transferDataToResSheet(orderArr, totalWeightList);  
  
}


function getOrdersList() {
 var lr = orderSheet.getLastRow();
 var lastColumn;
  
 for( var k = 2; k <= lr ; k++) {
     var m = 3; //start from column C
     while( orderSheet.getRange(k,m).isBlank() == false) { // equal while( !orderSheet.getRange(k,m).isBlank())
         m = m +1;
     }
      lastColumn = m-1;
  } 
  var orderArray =[];
  //Logger.log("last row = "+lr);
  //Logger.log("last col = "+lastColumn);
  for (var i=1; i <= lr; i++){
    if(i==1){
      var firstRow = [];
      for(var j=1; j<=lastColumn; j++){
        temp = orderSheet.getRange(i, j).getValue();// 1st row
          firstRow.push(temp);
      }      
      //Logger.log(firstRow); 
      orderArray.push(firstRow);
      
    }else{
      var nextRow = [];
       for(var j=1; j<=lastColumn; j++){
          temp = orderSheet.getRange(i, j).getValue();
          nextRow.push(temp);
      }
      //Logger.log(nextRow);
      orderArray.push(nextRow);
    }
  }
  //Logger.log(orderArray[1][2]);
  return orderArray;
}

//============================transfer data to results sheet=======================
function transferDataToResSheet(orderArray, totalWeightList){
  
  for (var i= 0; i<=orderArray.length-1; i++ ){
    if(i==0){
      //-----------------------------print header
      //print constant header      
      destination = orderArray[0][1];
      resultsSheet.getRange(i+1,2).setValue(destination);
      resultsSheet.getRange(i+1,3).setValue("Total Weight (KG)");
      resultsSheet.getRange(i+1,4).setValue("Carrier");
      resultsSheet.getRange(i+1,6).setValue("Offer");
      //print dinamic header
      rowLength = orderArray[0].length;
      for(j=0;j<=rowLength-1;j++){
        if(j>=2){
          resultsSheet.getRange(i+1,j+4).setValue(orderArray[i][j]);
          //Logger.log(destination);
        }
      }
    }else{
     //-----------------------print data in table 
      resultsSheet.getRange(i+1,1).setValue(orderArray[i][0]);//print company names
      resultsSheet.getRange(i+1,2).setValue(orderArray[i][1]);// print destination code
      //++++++++++++++print best carrier and offer
      destination = orderArray[i][1];
      weight = totalWeightList[i-1];
//      Logger.log("Dest: " + destination);
//      Logger.log("Weight:" + weight);      
      bestPriceAndOfferName = bestCarrier(destination, weight );
      resultsSheet.getRange(i+1,4).setValue(bestPriceAndOfferName[2]);//print Carrier
      resultsSheet.getRange(i+1,5).setValue(bestPriceAndOfferName[1]);//print Offer
      //Logger.log(bestPriceAndOfferName)
      
      
      
      rowLength = orderArray[0].length;
      for(j=0;j<=rowLength-1;j++){        
        if(j>=2){//-------------------------------print all velues in table from F2 to end
          resultsSheet.getRange(i+1,j+4).setValue(orderArray[i][j]);
          //Logger.log(destination);
        }
      }
    }
  }
}

//=============================count kg and volume=================================
function countKGandValume(orderArray){
  var resultsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Results');
  var numberOfBoxes = 0;
  var boxes33, boxes75, boxes30 = 0;
  var total, totalKG33, totalKG75, totalKG30 =0; 
  var totalArr = [];
  //get settings values
  var weightOfBox33 = settingsSheet.getRange(1, 2).getValue();
  var weightOfBox75 = settingsSheet.getRange(2, 2).getValue();
  var weightOfBox30 = settingsSheet.getRange(3, 2).getValue();
  var pallet = settingsSheet.getRange(4, 2).getValue();
  var maxWeightOfPallet = settingsSheet.getRange(5, 2).getValue();
  
 //loop in company names
  for (var i= 1; i<=orderArray.length-1; i++ ){
    // get total boxes of each
    boxes33 = getSummaryForVolume(numberOfBoxes, orderArray, i, 33);
    boxes75 = getSummaryForVolume(numberOfBoxes, orderArray, i, 75);
    boxes30 = getSummaryForVolume(numberOfBoxes, orderArray, i, 30);
    //----------formula to count total in KG        
    totalKG33 = (boxes33*weightOfBox33);
    totalKG75 = (boxes75*weightOfBox75);
    totalKG30 = (boxes30*weightOfBox30);
    total =  (totalKG33 + totalKG75 + totalKG30);
    var numOfPallet = 0;
    if(total>maxWeightOfPallet){
      numOfPallet = (total -(total%maxWeightOfPallet))/maxWeightOfPallet;      
    }
    total = (total + numOfPallet * pallet );//include pallets
    //carrier = bestCarrier(total, destination);
    totalArr.push(total);    
    
    resultsSheet.getRange(i+1,orderArray[0].length+5).setValue(boxes33);
    resultsSheet.getRange(i+1,orderArray[0].length+6).setValue(boxes75);
    resultsSheet.getRange(i+1,orderArray[0].length+7).setValue(boxes30);
    resultsSheet.getRange(i+1, 3).setValue(total);
      
//    Logger.log("boxes33: "+boxes33);
//    Logger.log("boxes75: "+boxes75);
//    Logger.log("boxes30: "+boxes30);
//    Logger.log("TOTAL: "+total);
  }  
  //Print Header total by volume
  resultsSheet.getRange(1,orderArray[0].length+5).setValue("Total boxes33");
  resultsSheet.getRange(1,orderArray[0].length+6).setValue("Total boxes75");
  resultsSheet.getRange(1,orderArray[0].length+7).setValue("Total boxes30");
  
  Logger.log("TOTAL-ARR: "+totalArr);
  return totalArr;  
}

// =========================Get summary for each volume====================
function getSummaryForVolume(numberOfBoxes, orderArray, i, volume){
  var rowLength = orderArray[0].length;
  
  //looking gor matches in header  
  for( j=0;j<=rowLength-1;j++){                //go through the row
    if(j>=2){                                 //numbers of boxes start from array index 2
      
      var a = orderArray[0][j].indexOf(volume);
      if(a>-1){
        // found column with 33cl and SUM value of boxes
        numberOfBoxes = numberOfBoxes + orderArray[i][j]; //orderSheet.getRange(2, j+4).getValue();
        //Logger.log( "Match 33cl: "+numberOfBoxes);
        //Logger.log( orderArray[0][j]);
      }
    }
  }
  return numberOfBoxes;
}

//======================================FOUND BEST CARRIER============================
var sss = SpreadsheetApp.openById('1vgH82c3Sbro8L4C36hfcjgEQy8MTWDytFaKUVe9sp1Q');
//var ss = sss.getUrl();// return URL

//-----------------------------get all sheets names in tabNamesArray
var tabNamesArray = new Array();
var sheets = sss.getSheets();
for (var i=0; i<sheets.length; i++){
  tabNamesArray.push([sheets[i].getName()]);
}

//==========================================bestCarrier=================
function bestCarrier(destination, weight ){
  //var destination = 20;
  //var weight = 312;
  var offer,carrier = "";
  var rowNum, columnNum, deliveryPrice; 
  var deliveryPrice;
  //------------------------------loop for all sheets with carriers
  
  for(index = 0; index<tabNamesArray.length; index++){//-------------------------------tabNamesArray.length
    //Logger.log(tabNamesArray[index]);
    var sheet = sss.getSheetByName(tabNamesArray[index]);
    var last_column = sheet.getLastColumn();
    var last_row = sheet.getLastRow();
    
    // -----------------------------------found destination row
    rowNum = getRowNumber(last_row, sheet, destination);  
    //-------------------------------------found column with weigth criteria
    columnNum = getColumnNumber(last_column, sheet, weight);
    
    if(rowNum&&columnNum){//chack if rowNum and columnNum is availible
      deliveryPriceTemp = sheet.getRange(rowNum, columnNum).getValue();      
      // --------------- chose lower price
      Logger.log("deliveryPriceTemp: "+deliveryPriceTemp);
      if(index==0 || !deliveryPrice){
        deliveryPrice = deliveryPriceTemp;
        offer = sheet.getRange(3, columnNum).getValue();
        carrier = sheet.getRange(1, 2).getValue();
      }else if(deliveryPrice>deliveryPriceTemp){
        deliveryPrice = deliveryPriceTemp;
        offer = sheet.getRange(3, columnNum).getValue();
        carrier = sheet.getRange(1, 2).getValue();
      }
    }
  }
  //fill data in array for Results sheet
  priceANDoffer = [
                    [deliveryPrice],
                    [offer],
                    [carrier]
                  ];
   Logger.log("deliveryPrice: "+priceANDoffer[0]);
   Logger.log("offer: "+priceANDoffer[1]);
   Logger.log("carrier: "+priceANDoffer[2]);
   //Browser.msgBox(priceANDoffer[0]);
  return priceANDoffer;
}

//=============found destination row========================================================!!!!!!!!!!!!
function getRowNumber(last_row, sheet, destination){
    var rowNum
    var destinationRow = sheet.getRange(7, 1, last_row-6, 1).getValues();
    for (j=0; j< destinationRow.length; j++){
      if(destinationRow[j] == destination){
      rowNum = sheet.getRange(j+7, 1).getRow();
      //Logger.log("rowNum: "+rowNum);
      }
    }
  //Logger.log("rowNum: "+rowNum);
  return rowNum;
}
//=============found column with weigth criteria========
function getColumnNumber(last_column, sheet, weight){
  var columnNum;
  var arrUpDown = sheet.getRange(4, 4, 2, last_column-3).getValues();  
  for(i=0; i<arrUpDown[0].length; i++) {
      if( (weight > arrUpDown[0][i]) && (weight <= arrUpDown[1][i])){
        columnNum = i+4// sheet.getRange(3, i).getColumn();
        //Logger.log("columnNum: "+columnNum);
      }
  }
  //Logger.log("columnNum: "+columnNum);
  return columnNum;
}
























