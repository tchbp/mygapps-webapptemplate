//WebApp ฟอร์มมอบตัว On line
var idnwsplogo = "1xUdhjmLjVwD742QImsa3snzXUDA3J6ny";
var wsID = "1jUbSOfabc_ToiDM_Hzc9ZJdNfH5SR-keVrU7ZfkHXls"; 
var hovID = "1OJN23m64xIrAbOlPRg8ZpOkaJxfsOW12wsHjfaI7xjk";

function doGet() {
    var htmlout = HtmlService.createTemplateFromFile('index');
  
    htmlout.logo = DriveApp.getFileById(idnwsplogo).getDownloadUrl();
    htmlout.province = getprv();

    return htmlout.evaluate()
                .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function chkreg(zid){
  var rng = SpreadsheetApp.openById(hovID).getSheetByName("regstd").getDataRange().getA1Notation();
  //Logger.log(stdrng);
    var shNew = SpreadsheetApp.openById(hovID).insertSheet();
    shNew.getRange(1,1).setFormula(`=QUERY(regstd!${rng},"select B,D,E,F,J,C where C = ${zid}")`);

    let nfound = shNew.getLastRow();
    let std2frm = [];
    if (nfound > 1){
      std2frm = shNew.getRange(2,1,nfound - 1, shNew.getLastColumn()).getValues();
      std2frm = std2frm[0];
      
    }
    
    //Logger.log(std2frm);
    //std2frm = std2frm[0];
    //Logger.log(std2frm);
    SpreadsheetApp.openById(hovID).deleteSheet(shNew);
       
    //Logger.log(std2frm);
    //Logger.log(std2frm.length);
    return std2frm;
    
}
function chkHov(zid='1419902302577') {
  var shHov = SpreadsheetApp.openById(hovID).getSheetByName("handover");
  var foundHov = shHov.getRange(2,18,shHov.getLastRow()-1,1).createTextFinder(zid).findNext();
  var hovarr = [];
  if (foundHov != null){
    hovarr = shHov.getRange(foundHov.getRow(),1,1,shHov.getLastColumn()).getValues();
  }
  if (hovarr.length > 0){
    hovarr = hovarr[0];
    var hovdata = {};
    //hovarr.forEach(function(v,i){
    //  hovdata["hov"+i] = v;
    //});
    hovdata["hov[]"] = hovarr;

    var JSONstring = JSON.stringify(hovdata);
    //Logger.log(JSONstring);
    //var hovJSON = ContentService.createTextOutput(JSONstring);
    //hovJSON.setMimeType(ContentService.MimeType.JSON);
    //Logger.log(hovJSON);
    let hovobj = JSON.parse(JSONstring);
    //Logger.log(hovobj);
    //Logger.log(hovobj["hov[]"].length);
    return hovobj ; 
  } else {
    return null;
  }
   
}

function getprv(){
    var prvdata = "Province";
    var ws = SpreadsheetApp.openById(wsID);
    var shData = ws.getSheetByName(prvdata);
  
    var prvth = shData.getRange(2,2,shData.getLastRow()-1,1).getValues();
    var stropt ="";
    for (var i = 0; i < prvth.length;i++){
      stropt += "<option value='" + prvth[i] + "' >" + prvth[i] + "</option>\n";
    }
    var template = HtmlService.createTemplate(stropt);
    return template.evaluate().getContent();
  }
  
  function getdist(prv="อุดรธานี"){
    var prvdata = "TambonDatabase";
 
    var rng = SpreadsheetApp.openById(wsID).getSheetByName(prvdata).getDataRange().getA1Notation();
  //Logger.log(stdrng);
    var shNew = SpreadsheetApp.openById(wsID).insertSheet();
    shNew.getRange(1,1).setFormula(`=QUERY(TambonDatabase!${rng},"select L,count(L) where O = '${prv}' group by L ")`);
    var stdArr = shNew.getRange(2,1,shNew.getLastRow() - 1, 1).getValues();
    //Logger.log(stdArr);
    SpreadsheetApp.openById(wsID).deleteSheet(shNew);
    var dist2opt = [];
    for (const d of stdArr) {
        dist2opt.push(...d);
    }
    //Logger.log(dist2opt);
    return dist2opt;
  
  }
  function gettumbol(dist="กุมภวาปี"){
    var prvdata = "TambonDatabase";
 
    var rng = SpreadsheetApp.openById(wsID).getSheetByName(prvdata).getDataRange().getA1Notation();
  //Logger.log(stdrng);
    var shNew = SpreadsheetApp.openById(wsID).insertSheet();
    shNew.getRange(1,1).setFormula(`=QUERY(TambonDatabase!${rng},"select D,count(D) where L = '${dist}' group by D ")`);
    var stdArr = shNew.getRange(2,1,shNew.getLastRow() - 1, 1).getValues();
    //Logger.log(stdArr);
    SpreadsheetApp.openById(wsID).deleteSheet(shNew);
    var tb2opt = [];
    for (const d of stdArr) {
        tb2opt.push(...d);
    }
    //Logger.log(tb2opt);
    return tb2opt;
  
  }
  function onSendFrm(obj2get){
    //Logger.log(obj2get);
    let hovarr = [];
    hovarr.push(obj2get.stdid);
    hovarr.push(obj2get.pref);
    hovarr.push(obj2get.fname);
    hovarr.push(obj2get.sname);
    hovarr.push(obj2get.classr);
    for (i = 0; i <= 29;i++){
      hovarr.push(obj2get["data" + i]);
    }
    //Logger.log(hovarr);
    let shHov = SpreadsheetApp.openById(hovID).getSheetByName("handover");
    shHov.getRange(shHov.getLastRow() + 1, 1, 1, shHov.getLastColumn()).setValues([hovarr]);
    return true;
  }