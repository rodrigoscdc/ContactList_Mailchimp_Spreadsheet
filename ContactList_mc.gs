/*****************************************************************
* 
* Mailchimp API - Get Contacts from mailchimplist to Google Sheets
* by rodrigoscdc 2017
* 
*/

// Add a menu at your spreadsheet
function onOpen() {
  var menu = SpreadsheetApp.getUi().createMenu('MailChimp');
  menu.addItem('Grab list', 'mailchimpContactsList').addToUi();
}

// Replace to your mailchimp API-KEY
var API_KEY = 'YOUR API-KEY MAILCHIMP';

/*
*
* Script to grab the List name and show in a pop-up dialog to user
* choose an index that will point to specific list.
* This Function will be called by mailchimpContactsList()
*
*/

function getLists(){
  var root = 'https://us13.api.mailchimp.com/3.0/';
  var endpoint = 'lists?count=30';
  
  var params = {
    'method': 'GET',
    'muteHttpExceptions': true,
    'headers': {
      'Authorization': 'apikey ' + API_KEY
    }
  };
  
  try {
    var response = UrlFetchApp.fetch(root+endpoint, params);
    var data = response.getContentText();
    var json = JSON.parse(data);
    var lists = json['lists'];
    
    var ui = SpreadsheetApp.getUi();
    var answer = false;
    //Logger.log(lists);
    while(!answer){
    var a = [];
    for (var i in lists){
        //lists.join("; );
        //Logger.log(lists[i]);
        a.push(i + ': ' + lists[i]['name']);
      }
      //answer = true;
      //Logger.log(a);
      
      var userResponse = ui.prompt("Listas MailChimp", "Digite o Número referenta a lista que se quer os dados: " + a.join(' ### '), ui.ButtonSet.YES_NO);
      if (userResponse.getSelectedButton() == ui.Button.YES){
        var listUser = parseInt(userResponse.getResponseText());
        for (var i in lists){
          if (listUser < lists.length){
            //Logger.log("verdade");
            //Logger.log(lists[listUser]);
            return lists[listUser];
          } 
        }
        
      } else {
        ui.alert("Operação Cancelada!");
        return false;
      }   
    }
  }
  catch(error){
    Logger.log(error);
  }
}

/*********************************************************
*
* Scripts that is called by user from Mailchimp Menu
*
*/

function mailchimpContactsList() {
  
  // URL and params for the Mailchimp API
  var listas = getLists();
  var listIdnt = listas['id'];
  var root = 'https://us13.api.mailchimp.com/3.0/';
  var endpoint = 'lists/' + listIdnt + '/members?count=' + listas['stats']['member_count'];
  //Logger.log(endpoint);
  // parameters for url fetch
  var params = {
    'method': 'GET',
    'muteHttpExceptions': true,
    'headers': {
      'Authorization': 'apikey ' + API_KEY
    }
  };
  
  try {
    // call the Mailchimp API
    var response = UrlFetchApp.fetch(root+endpoint, params);
    var data = response.getContentText();
    var json = JSON.parse(data);
    
    // get just campaign data
    var members = json['members'];
    //Logger.log(members);
    // blank array to hold the campaign data for Sheet
    var campaignData = [];

    // Use headers to data in sheet, replace for your choice
    campaignData.push(['id', 'E-mail', 'Nome', 'Telefone', 'Cidade', 'Origem E-mail']);
    
  
    // Add the campaign data to the array
    for (var i = 0; i < members.length; i++) {
      
      campaignData.push([
        i,
        members[i]['email_address'],
        // use replace for the id of merge fields that is in your list
        members[i]['merge_fields']["FNAME"],
        members[i]['merge_fields']["PHONE"],
        members[i]['merge_fields']["CITY"],
        members[i]['email_client']
      ]);
    }
    
    // Log the campaignData array
    //Logger.log(campaignData);
    
    // select the campaign output sheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet;

    // test if sheet name already exist if not, it will be created
    if (!ss.getSheetByName(listas['name'])){
      ss.insertSheet(listas['name']);
    }
    sheet = ss.getSheetByName(listas['name']);
    sheet.setTabColor('green');
    
    // calculate the number of rows and columns needed
    var numRows = campaignData.length;
    var numCols = campaignData[0].length;
    
    sheet.getRange(1, 1, 1).setValue('Relação de Contatos na Lista ' + listas['name']);
    //sheet.getRange(3, 1, 1, numCols).setValues();
    //Logger.log
    // output the numbers to the sheet
    sheet.getRange(4,1,numRows,numCols).setValues(campaignData);
    
    sheet.getRange(1, 1).setFontSize(16);
    
    sheet.getRange(1, 1, 1, numCols).merge();
    sheet.getRange(1, 1, 1, numCols).setHorizontalAlignment('center');
    
    
  }
  catch(error) {
    // deal with any errors
    Logger.log(error);
  }
  
}