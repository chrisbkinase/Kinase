function main() {
  var SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1nhqr1zeMsqlZBkIwbnbMQpbBnNoPqsDblvvMFLcxwBA/edit#gid=0";
  var SHEET_NAME = 'Overall';
  var today = new Date();
  var date_str = [today.getFullYear(),(today.getMonth() + 1),today.getDate()].join("-");
  
   Logger.log(date_str);
  var qsArray = new Array();
  var spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  var qs_sheet = spreadsheet.getSheetByName(SHEET_NAME);
  var productCatRange = spreadsheet.getRangeByName("ProdCats").getValues();
Logger.log(productCatRange);
  var prodCatRange = productCatRange[0];
Logger.log(prodCatRange);
//  var campaign_iter = AdWordsApp.campaigns()
 //  .withCondition("LabelNames CONTAINS_ANY ['Power Tools']")
  
  var tot_weighted_qs = 0;
  var tot_imps = 0;
  //for loop cycles through product categories based on range in spreadsheet
  
  for (var x = 0; x<prodCatRange.length; x++){
      Logger.log(x);
      Logger.log(prodCatRange[x]);
    
  //get list of keywords in each product category with more than 10 impressions for the last 7 days
    
  var kw_iter = AdWordsApp.keywords()
    .withCondition("Status = ENABLED")
    .withCondition("CampaignName CONTAINS '" + prodCatRange[x] +"'")
    .forDateRange("LAST_7_DAYS")
    .withCondition("Impressions > 10")
    .orderBy("Impressions DESC")
    .withLimit(50000)
    .get();
 
  //temporary variables for quality score
    
  var prodcat_weighted_qs = 0;
  var prodcat_imps = 0;
  var count = 0;
  
//calculate quality score for product category 
//also keep track of impressions and weighted qs for account qs
//count
  var qsBreakdownArray = [prodCatRange[x],date_str,0,0,0,0,0,0,0,0,0,0,0];
  var bd_sheet = spreadsheet.getSheetByName('Breakdown');
  while(kw_iter.hasNext()) {
    var kw = kw_iter.next();
    var kw_stats = kw.getStatsFor("LAST_7_DAYS");
    var imps = kw_stats.getImpressions();
    var qs = kw.getQualityScore();
    qsBreakdownArray[qs + 2] += 1;
    prodcat_weighted_qs += (qs * imps);
    tot_weighted_qs += (qs * imps);
    prodcat_imps += imps;
    tot_imps += imps; 
  }
     Logger.log(qsBreakdownArray);
    
    //add date, product category and quality score breakdown to Breakdown sheet
    
    if(bd_sheet != null){
      bd_sheet.appendRow(qsBreakdownArray);
    }
    //push product category qs into array
    
  var prodcat_qs = 0;  
  var prodcat_qs = prodcat_weighted_qs / prodcat_imps;
    Logger.log(prodcat_qs);
  qsArray.push(prodcat_qs);
  Logger.log(qsArray);
  }
  
  //end of for loop
  //calculate account qs
  
  var acct_qs = tot_weighted_qs / tot_imps;
  Logger.log([date_str, acct_qs, qsArray]);
  var fullArray = new Array();
  fullArray = [date_str, acct_qs];
  
  //create array of date and all quality scores
  //append to next row of spreadsheet
  
  for(x=0; x<prodCatRange.length; x++){
    fullArray.push(qsArray[x]);
    Logger.log(fullArray);
  }
  qs_sheet.appendRow(fullArray);
}