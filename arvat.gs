const startrow = 3

function get_sheets_() {
  const ss = SpreadsheetApp.getActive()
  const sh = ss.getSheets()
  let response = []
  sh.forEach(sheety)
  function sheety(item) {
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
  let winnings = separate_given_numbers_(apply_notes_())
  sh.appendRow(["Vantaan Invalidit VANIN ry",""])
  sh.appendRow(["Muokkaa tätä titteliä", ""])
  for (let i = 0; i < winnings.length; i++){
    //Logger.log(winnings[i])
    sh.appendRow(winnings[i])
  }
}

function format_output(){
  const ss = SpreadsheetApp.getActive()
  const sh = ss.getSheetByName("Arvat")
  let range = sh.getRange("A3:D101")
  sh.setActiveRange(range)
  sh.getActiveRange().applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY)
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

// spreadsheet.getCurrentCell().offset(-1, 1).setNote('13, 14');
function read_notes_(){
  const ss = SpreadsheetApp.getActive()
  const sh = ss.getSheetByName("Arvottavat esineet")
  let items = sh.getSheetValues(startrow, 1, get_last_row_() + 1 - startrow, 3)
  let notes = sh.getRange(startrow, 1, get_last_row_() + 1 - startrow, 3).getNotes()

  for (let i = 0; i < items.length; i++){
    notes[i].pop()
    notes[i].reverse()
    notes[i].pop()
    items[i].push(notes[i])
  }
  return items 
}

function apply_notes_(){
  let items = read_notes_()
  let items_list = []

  items.forEach(pieces)
  function pieces(piece){
    let tmp = []
    if (piece[2]){
      tmp = piece[3][0].split(",")
      let rounds = piece[1] - tmp.length
      for (let i = 0; i < rounds; i++){
        tmp.push("")
      }
      for (let i = 0; i < piece[1]; i++){
        items_list.push([piece[0], tmp[i]])
      }
    }
  }
  //Logger.log(items_list)
  return items_list
}

function separate_given_numbers_(itemslist=apply_notes_()){
  const ss = SpreadsheetApp.getActive()
  const sh = ss.getSheetByName("Arvottavat esineet")
  let given_numbers = []
  let non_given_numbers = []
  let reserved_numbers = []
  let free_numbers = []
  let randz = []

  itemslist.forEach(pieces)
  function pieces(piece){
    let tmp = Number(piece[1])
    if(tmp == 0){
      non_given_numbers
  .push(piece[0])
    } else {
      given_numbers.push([tmp, piece[0]])
      reserved_numbers.push(tmp)
    }
  }
  shuffleArray_(non_given_numbers) 
  for (let i = 1; i < 100; i++){ 
    if(!reserved_numbers.includes(i)){
      free_numbers.push(i)
    }
  }
  for (let i = 0; i < free_numbers.length; i++){ 
    randz.push([free_numbers[i], non_given_numbers[i]])
  }
  for (let i = 0; i < given_numbers.length; i++){ 
    randz.push(given_numbers[i])
  }
  randz.sort(compareNumbers)
  function compareNumbers(num1, num2) {
    if(num1[0] > num2[0])
      return 1;
    if(num1[0] < num2[0])
      return -1;
    return 0;
  }
  //Logger.log(randz)
  return randz
}
