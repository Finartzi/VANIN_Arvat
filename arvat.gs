const startrow = 3

function get_sheets_() {
  const ss = SpreadsheetApp.getActive()
  const sh = ss.getSheets()
  let response = []
  
  sh.forEach(sheety)
  function sheety(item)
  {
    response.push([item.getName(), item.getSheetId()])
  }
  Logger.log(response)
  return response
}

function get_last_row_(sheetName="Arvottavat esineet"){
  const ss = SpreadsheetApp.getActive()
  const sh = ss.getSheetByName(sheetName)
  let lastRow = sh.getLastRow()
  //Logger.log(lastRow)
  return lastRow
}

// Nuolinäppöin sorttaus
function SORT_ITEMS(sheetName="Arvottavat esineet", sortCol = 1){
  const ss = SpreadsheetApp.getActive()
  const sh = ss.getSheetByName(sheetName)
  sh.sort(sortCol, true)
}

// Luo uusi lista -- nuolinäppäin
function CREATE_LIST(){
  const ss = SpreadsheetApp.getActive()
  const sh = ss.getSheetByName("Arvat")
  sh.clearContents()
  let winnings = randomize_item_list_()
  sh.appendRow(["Vantaan Invalidit VANIN ry",""])
  sh.appendRow(["Muokkaa tätä titteliä", ""])
  
  for (let i = 0; i < winnings.length; i++){
    //Logger.log(winnings[i])
    sh.appendRow(winnings[i])
    //.setRowHeight(i,30)
  }


}

function format_output(){
  const ss = SpreadsheetApp.getActive()
  const sh = ss.getSheetByName("Arvat")
  let range = sh.getRange("A3:D101")
  sh.setActiveRange(range)
  sh.getActiveRange().applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY)
}


function create_item_list_(){
  const ss = SpreadsheetApp.getActive()
  const sh = ss.getSheetByName("Arvottavat esineet")
  let items = sh.getSheetValues(startrow, 1, get_last_row_() + 1 - startrow, 3)
  let items_list = []
  items.forEach(pieces)
  function pieces(piece){
    if (piece[2]){      
      for (let i = 0; i < piece[1]; i++){
        items_list.push(piece[0])
        //Logger.log(piece[0])
      }
    }
  }
  return items_list
}

function randomize_item_list_(item_list=create_item_list_()){
  let randz = []
  item_list.splice(item_list.indexOf("Päävoitto"), 1)
  shuffleArray_(item_list)  
  for (let i = 1; i < item_list.length; i++){
    randz.push([i + ".", item_list[i]])
  }
  randz.push([100 + ".", "Päävoitto"])
  return randz
}

/**
 * Randomize array element order in-place.
 * Using Durstenfeld shuffle algorithm.
 */
function shuffleArray_(array) {
    for (var i = array.length - 1; i > 0; i--) {
        var j = Math.floor(Math.random() * (i + 1));
        var temp = array[i];
        array[i] = array[j];
        array[j] = temp;
    }
}
