/*
//
CREATED BY JAMES JARA www.jamesjara.com #jamesjara
Get Visits Google Analytic ( magic ) from url value CELL
help in www.jamesjara.com , Get the visits from your google analytic account from a url value 

Usage For Analytics:
1. Clicking the button Analytics By Cell URL with a cell selected.
2. Using the function on the cell =getVisits( CELL )

Usage For Adsense:
1. Clicking the button Adsense By Cell URL with a cell selected.
2. Using the function on the cell =getAdsense( CELL )


Please if you creates another function send to jamesjara@gmail.com with your name,etc.. to mantain a unique version,
example:

/** 
* autor:
* description:
function getTopKeywords(){}

@license : use as you wish but share :)
 
http://code.google.com/p/google-analytic-spreadsheet/
 
*/


/**
 * @autor jamesjara
 * @description  used to create the menu on the spreadsheet GUI
 */
function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [ {name: "Get visits of last 3 months", functionName: "getVisits2"},{name: "Get Earning of last 3 months", functionName: "getAdsense2"}];
  ss.addMenu("Analytics By Cell URL", menuEntries);  
}

/* ug messy function :/ but maybe could be usefull for any programator.
function ValidateURL( URL ){
  //Get accounts
  var accounts = Analytics.Management.Accounts.list();  
  for (var i = 0; i <= accounts.getItems().length  ; ++i) {      
      if (accounts.getItems()[i]){
      var SuperAccount = accounts.getItems()[i].getId();
      var PropiedadesWeb = Analytics.Management.Webproperties.list(SuperAccount);
      //Now we get the subaccounts
      var SubAccount;
      for(var x = 0; x<= PropiedadesWeb.getItems().length; ++x ){
        if ( PropiedadesWeb.getItems()[x]){
          if ( URL == PropiedadesWeb.getItems()[x].getWebsiteUrl() ){
          Logger.log(PropiedadesWeb.getItems()[x].getWebsiteUrl());
            //return PropiedadesWeb.getItems()[x].getId();              
            var profiles = Analytics.Management.Profiles.list( '~all' , '~all' );
            if (profiles.getItems()) {
              var firstProfile = profiles.getItems()[0].getId();
              Logger.log(profiles.getItems());
              return firstProfile;              
            } else {
              throw new Error('No profiles found.');
            }
          } 
        }   
      } 
     } 
  };
  throw new Error ("Site not found in your account.");
}*/

/**
 * @autor jamesjara
 * @description  used to validate if the current url is on the current analytics profile 
 * @param a url string ,exactly as used on your analytic account.
 * @return if the url is valid will return the respective id of the current profile.
 */
function ValidateURL( URL ){  
  //Get accounts
  var accounts = Analytics.Management.Accounts.list();  
  for (var i = 0; i <= accounts.getItems().length  ; ++i) {   
    var profiles = Analytics.Management.Profiles.list( '~all' , '~all' );
    if (profiles.getItems()) {     
      for (var x = 0; x <= profiles.getItems().length  ; ++x) {
        if(profiles.getItems()[x]){
          var firstProfile = profiles.getItems()[x].getId();
          var WebsiteUrl = profiles.getItems()[x].getWebsiteUrl();
          Logger.log(WebsiteUrl);
          if ( URL == WebsiteUrl ){
            return profiles.getItems()[x].getId();
          }
        }
      }
    } else {
      throw new Error('No profiles found.');
    }
  }  
}


/**
 * @autor jamesjara
 * @description  wrapper used to call the function from the button
 */
function getAdsense2() {
     getAdsense(SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveCell().getValue());
}


/**
 * @autor jamesjara
 * @description used to get adsense 'EARNINGS', 'CLICKS'  from a custom domain
 * @param a String with the custom domain
 * @return last 3 months visits
 */
function getAdsense( DOMAIN ){
  if (DOMAIN == undefined || DOMAIN == null || DOMAIN == "")
  {
     throw new Error (" DOMAIN cant be empty.");      
  }
  var startDate = getLastNdays(72);   
  var endDate   = getLastNdays(0);// Today.
  var args = {
    'filter':['DOMAIN_NAME=@'+DOMAIN.match(/:\/\/(www[0-9]?\.)?(.[^/:]+)/)[2]],
    'metric': ['EARNINGS', 'CLICKS'],
    'dimension': ['MONTH','DOMAIN_NAME']};
  var report = AdSense.Reports.generate(startDate, endDate, args).getRows();
  Logger.log(args);
  if (report){
    //This is a hack :( because a cell cant logging into Adsense
    var result = '"'+report[0][2]+','+report[0][3]+'","'+report[1][2]+','+report[1][3]+'","'+report[2][2]+','+report[2][3]+'"';
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveCell().setFormula("=({"+result+"})");
    return;
  }  else throw new Error (" No adsense profiles found."); 
}


/**
 * @autor jamesjara
 * @description  wrapper used to call the function from the button
 */
function getVisits2() {
     getVisits(SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveCell().getValue());
}


/**
 * @autor jamesjara
 * @description used to get visits from a determinated profile
 * @param a String profile Id , https://developers.google.com/apps-script/class_analytics_v3_schema_profile
 * @return last 3 months visits
 */
function getVisits( ProfileId ) {
  ProfileId =  ValidateURL( ProfileId );
  if (ProfileId == undefined || ProfileId == null || ProfileId == "")
  {
     throw new Error (" No profiles found.");      
  }
  try {    
    var ProfileId   = 'ga:' + ProfileId;
    var startDate = getLastNdays(72);   
    var endDate   = getLastNdays(0);// Today.
    
    var optArgs = {
      'dimensions': 'ga:month',              // Comma separated list of dimensions.
      //'sort': '-ga:visits,ga:keyword',         // Sort by visits descending, then keyword.
      //'segment': 'dynamic::ga:isMobile==Yes',  // Process only mobile traffic.
      //'filters': 'ga:source==google',          // Display only google traffic.
      'start-index': '1',
      'max-results': '250'                     // Display the first 250 results.
    };
    
    // Make a request to the API.
    var results = Analytics.Data.Ga.get(
      ProfileId,                 
      startDate,             
      endDate,               
      'ga:visits', // Comma seperated list of metrics.
      optArgs);
      // https://www.googleapis.com/analytics/v3/data/ga?ids=ga%3A37121403&metrics=ga%3Avisitors%2Cga%3Avisits&start-date=2012-10-27&end-date=2012-11-10&max-results=50
      // SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveCell().setValue( results );

    if (results.getRows()) {
      //This is a hack :( because a cell cant logging into Analytics
      var result = results.getRows()[0][1]+","+results.getRows()[1][1]+","+results.getRows()[2][1];
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveCell().setFormula("=({"+result+"})");
      return;
    } else {
      throw new Error('No data found');
    }   
  } catch(error) {
    throw new Error( error.message );
  } 
};


function getLastNdays(nDaysAgo) {
  var today = new Date(); 
  var before = new Date();
  before.setDate(today.getDate() - nDaysAgo);
  return Utilities.formatDate(before, 'GMT', 'yyyy-MM-dd');
}
 
