function addTech(){

  let sheetR = SpreadsheetApp.getActive().getSheetByName('FirstFile');
  let sheetW = SpreadsheetApp.getActive().getSheetByName('SecondFile')


  var DonnesDuTableauLine1 = SpreadsheetApp.getActive().getSheetByName('FirstFile').getRange('A1:C1')
  
  DonnesDuTableauLine1.getCell(1,1).setValue('Non')
  DonnesDuTableauLine1.getCell(1,2).setValue('Matricule')
  var cell = sheetR.getRange("C1")
  cell.setNumberFormat("@")


  let techList = []
  let datasheetR = sheetR.getDataRange().getValues()

  
  for (var i in datasheetR){
    techList.push([datasheetR[i][1],datasheetR[i][2]])
  }

  var jour = DonnesDuTableauLine1.getCell(1,3).getValue();

  let matriculeList = []
  let chargeList = []
  let datasheetW = sheetW.getDataRange().getValues()

  for (var i in datasheetW){
    matriculeList.push([datasheetW[i][2]])
    chargeList.push([datasheetW[i][3]])
  }
 
  let range = sheetW.getRange("A1:AI100")
  let col = sheetW.getLastColumn() + 1 
  let listChargeTache = []; 

  for (var i in matriculeList){
    var matricule = matriculeList[i][0]
    var charge = chargeList[i][0] 
    for (var j in techList){
        var matriculeTech = techList[j][0]
        if (matricule == matriculeTech){
          listChargeTache.push(parseInt(charge)) 
          var row = parseInt(i);
          range.getCell(row+1, col).setValue(techList[j][1])
          range.getCell(row+1, col+1).setValue(techList[j][1])
          if (charge > techList[j][1]){
            range.getCell(row+1, col).setFontColor('red')
            range.getCell(row+1, col+1).setFontColor('red')
          }
          else{
            range.getCell(row+1, col).setFontColor('#1c4587')
            range.getCell(row+1, col+1).setFontColor('#1c4587')
          }}}
  }
  
  var sumChargeTacheprv = 0
  for (var i in listChargeTache){
    var nbChargeTache = listChargeTache[i]
    if(!(isNaN(nbChargeTache))){
      var sumChargeTacheprv = sumChargeTacheprv + nbChargeTache
    }
  }

  sum = sheetW.getRange(69,col).getValues()

  range.getCell(1, col).setValue('Charge engag??e / Charge jour\nNb de dossier\nplannifi??\n'+ jour).setFontColor('black')
  range.getCell(1, col+1).setValue(sumChargeTacheprv + ' / ' + sum +'\n\nNb de t??che\n'+ jour).setFontColor('black')
}
