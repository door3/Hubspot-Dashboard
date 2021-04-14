/*
HUBSPOT Api:
https://legacydocs.hubspot.com/docs/overview
 
Google Sheets Api:
https://developers.google.com/sheets/api/reference/rest
 
 
Assignment:
I was tasked with creating a database using data from hubspot, we will be focusing on company objects. 
 
First, the user must paste their Hubspot API key in our main function.
Second, the user will select a date, which will let the script know which companies to display in google sheets.
 
*Optional
Third, The user will select and adjust any properties they want included in the "get_all_company_property_names()" function. * More instructions included in the function itself on how to do this 
 
After these three steps have been completed the User can run the program.
 
Now, the program will query hubspot for all companies,and explictly state which properties we  want returned.
The, script will loop multiple times until all Companies have been returned and stored into a singular array. 
*Note you can only recieve 250 objects per call. Also, you can only make 100 calls every 10 seconds. * Check your hubspot account API restrictions
*Therfore, if you have too many companies and exceed your call limit the program will crash. If this is the case, you must build a sleep function to pause the program for 10 seconds once you are nearing 100 calls.
 
When you query hubspot for a specific company object this is an example of the JSON it returns
 
Example response:
{
  "portalId": 62515,
  "companyId": 10444744,
  "isDeleted": false,
  "properties": {
    "description": {
      "value": "A far better description than before",
      "timestamp": 1403218621658,
      "source": "API",
      "sourceId": null,
      "versions": [
        {
          "name": "description",
          "value": "A far better description than before",
          "timestamp": 1403218621658,
          "source": "API",
          "sourceVid": [
            
          ]
        }
      ]
    },
    "name": {
      "value": "A company name",
      "timestamp": 1403217668394,
      "source": "API",
      "sourceId": null,
      "versions": [
        {
          "name": "name",
          "value": "A company name",
          "timestamp": 1403217668394,
          "source": "API",
          "sourceVid": [
            
          ]
        }
      ]
    },
    "createdate": {
      "value": "1403217668394",
      "timestamp": 1403217668394,
      "source": "API",
      "sourceId": null,
      "versions": [
        {
          "name": "createdate",
          "value": "1403217668394",
          "timestamp": 1403217668394,
          "source": "API",
          "sourceVid": [
            
          ]
        }
      ]
    }
  }
}
 
 
*Note: Properties are stored in a singular array called "properties".
 
After we have all of our data in an array, we modify certain properties such as dates, we replace the timestamp returned in milliseconds with a javascript data object. 
This makes sorting easier.
The property we are sorting by in our script is "hs_lastmodifieddate".
Finally, we display our data on the spreadsheet.
 
 
 
//Please read hubspot documentation if you don't understand this section.
 
Basic Example
NameOfArray.companies.length; This returns a number which represents how many company objects are in your companies array.
 
NameOfArray.companies[x].properties.someProperty.value; where x is a number. This line would return the the value of a specific property for a selected company 
 
This notation confused me at first, which is why I provided examples.
 
 
*/
 
 
 
 
// This is our main function
function main() {
 
  // Your Hubspot Unique API Key
  // The date from which you want your data onward

  var hubspotApiKey = "";

   
  //Insert a date if you do not get any data back.The format for javascript date objects is "MM/DD/YYYY"  
  // *IF YOUR DATE IS TOO OLD, YOU WILL GET ALOT OF DATA,THIS WILL RUN VERY SLOW. We default to today. 
  var getDataAfterThisDate = new Date(); 
 
 
  //Whenever you see the word company/companies, I am refering to Hubspot company Objects
 
  var allHubspotCompanies = [];
  allHubspotCompanies = get_all_companies(hubspotApiKey);
 
 
  // We display data on our sheet via the last_modified_date of a company
  
  replace_last_modified_date_with_javascript_date_object(allHubspotCompanies);
  sort_companies_by_last_modified_date(allHubspotCompanies);
  remove_companies_older_than_provided_date(allHubspotCompanies,getDataAfterThisDate);
 
 
  //SpreadsheeApp is proprietary to Google Sheets, please see documenation listed above.*
  // This section grabs the first sheet tab, which is the tab we will be writing to
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var allSheetTabs = [];
  allSheetTabs = activeSpreadsheet.getSheets();
  var activeSheetBeingModified = allSheetTabs[0]; 
  
 
  // Display all of our data on our Google Sheet
  display_companies_on_active_sheet(activeSheetBeingModified,allHubspotCompanies);
 
 
 
}
 
 
/*
 * This function returns every company in a Hubspot Database.
 * @param {string} hubspotApiKey - A variable which contains a unique Api key for Hubspot
 * @return {object} An array of Hubspot Company Objects
 * 
*/
function get_all_companies(hubspotApiKey){
      
    hubspotUrlQuery =  construct_initial_hubspot_query(hubspotApiKey);  
    
    var allHubspotCompanies = [];
    allHubspotCompanies = get_companies_from_query(hubspotUrlQuery); 
    
    // Hubspot Url queries are limited to returning 250 objects at a time.
    // These two variables, help call remaining hubspot objects
    var pageOffset = allHubspotCompanies.offset;
    var hasMoreCompanies = allHubspotCompanies ['has-more'];
  
    while(hasMoreCompanies == true){
      
      var offsetHubspotUrlQuery = construct_offset_hubspot_query(hubspotApiKey,pageOffset);
      
      var remainingCompanies = [];
      remainingCompanies = get_companies_from_query(offsetHubspotUrlQuery); 
 
    
      hasMoreCompanies = remainingCompanies['has-more'];
      pageOffset = remainingCompanies.offset;
 
       
      // Append the recently returned company objects array to our original array
      allHubspotCompanies.companies = allHubspotCompanies.companies.concat(remainingCompanies.companies);
 
    }
 
    //return our array which now contains every company from hubspot
    return allHubspotCompanies;
 
}
 
 
 
 
/*
 * This function creates and returns a url, which is used to query Hubspot.
 * @param {string} hubspotApiKey - A variable which contains a unique API key for Hubspot
 * @return {string} A string which can be used to query Hubspot
 * 
*/ 
function construct_initial_hubspot_query(hubspotApiKey){
 
  var companyPropertyNames = [];
  companyPropertyNames = get_all_company_property_names();
 
  var hubspotUrlPlusProperties = "https://api.hubapi.com/companies/v2/companies/paged?hapikey="+ hubspotApiKey +"&includeAssociations=true&limit=250";
 
  //Add desired properties to the end of our Hubspot query
  for(i=0; i < companyPropertyNames.length; i++){
    hubspotUrlPlusProperties += ("&properties=" + companyPropertyNames[i]);
  }
 
  //Note if your Url is too long, meaning you have too many properties you will need a way to shorten your url. FireBase dynamic links is a viable option.
  return hubspotUrlPlusProperties;
 
}
 
 
 
 
/*
 * This function creates and returns a url, which is used to query Hubspot.
 * @param {string} hubspotApiKey - A variable which contains a unique API key for Hubspot
 * @param {string} pageOffset - A variable which represents the pagination offset required to not call duplicate Hubspot objects
 * @return {string} A string which can be used to query Hubspot
 * 
*/
function construct_offset_hubspot_query(hubspotApiKey,pageOffset){
 
  var companyPropertyNames = [];
  companyPropertyNames = get_all_company_property_names();
 
  var hubspotUrlPlusPropertiesAndOffset = "https://api.hubapi.com/companies/v2/companies/paged?hapikey="+ hubspotApiKey +"&includeAssociations=true&limit=250&offset=" + pageOffset ;
  
  //Add desired properties to the end of our Hubspot query
  for(i=0; i < companyPropertyNames.length; i++){
    hubspotUrlPlusPropertiesAndOffset += ("&properties=" + companyPropertyNames[i]);
  }
 
   //Note if your Url is too long, meaning you have too many properties you will need a way to shorten your url. FireBase dynamic links is a viable option.
  return hubspotUrlPlusPropertiesAndOffset;
}
 
 
 
 
/*
 * This function queries a Hubspot Url which responds with text in JSON format. 
 * @param {string} url - A variable which contains a URL which links to Hubspot
 * @return {string} A string which contains the query response as text
 * 
*/
function fetch_url_response(url){
   //UrlFetchApp is proprietary to Google Scripts and Google sheets, please see documentation
    var urlResponse = UrlFetchApp.fetch(url);
    var urlResponseAsText = urlResponse.getContentText();
    
    return urlResponseAsText;
}
 
 
 
 
/*
 * This function parses a query response into JSON, then returns this parsed response.
 * @param {string} url - A variable which contains a URL which links to Hubspot
 * @return {object} An array of Company Objects 
 * 
*/
function get_companies_from_query(url){
  
  //fetch our data as text using provided url
  var urlResponseAsText = fetch_url_response(url);
 
  //parse Json data into an array 
  var allCompanies = JSON.parse(urlResponseAsText);
 
  return allCompanies;
}
 
 
 
 
/*
 * This function assigns Javascript date objects to every Hubspot Company Object.
 * @param {object} allHubspotCompanies - An array of Hubspot Company Objects
 * @return: void
*/
function replace_last_modified_date_with_javascript_date_object(allHubspotCompanies){
  // Hubspot returns "dates" in milliseconds, therefore we use this timestamp and create a date Object. This makes it easier to sort by the date.
 
  //Assign Date object to every company lastmodifieddate property
  for(i = 0 ; i < allHubspotCompanies.companies.length; i++){
    if(allHubspotCompanies.companies[i].properties.hs_lastmodifieddate == null){
     
     // do nothing if for some reason this property does not exist
    
    }
    else{
 
      // The value of a companies last_modified_date is now a javascript data object created with the correct millisecond timestamp.
      allHubspotCompanies.companies[i].properties.hs_lastmodifieddate.value = new Date(Number(allHubspotCompanies.companies[i].properties.hs_lastmodifieddate.value));
 
    }
  }   
 
}
 
 
/*
 * This function sorts every Hubspot Company Object by Last Modifed Date.
 * @param {object} allHubspotCompanies - An array of Hubspot Company Objects
 * @return: void
*/
function sort_companies_by_last_modified_date(allHubspotCompanies){
 
  // Sort companies by Last Modified Date, New to Old
  allHubspotCompanies.companies.sort(function(a,b){
  return b.properties.hs_lastmodifieddate.value - a.properties.hs_lastmodifieddate.value;
  });
 
 
}
 
 
 
 
/*
 * This function removes Hubspot Company Objects from the passed in array,by Last Modifed Date. Dependent on the Date passed in.
 * @param {object} allHubspotCompanies - An array of Hubspot Company Objects
 * @param {object} getDataAfterThisDate - Javascript Date Object
 * @return: void
*/
function remove_companies_older_than_provided_date(allHubspotCompanies,getDataAfterThisDate){
   
  // getDataAfterThisDate is a javascript Date Object
  // hs_lastmodifieddate.value is also a javascript Date Object
 
  //Remove all objects from our array after this index
  var indexToSplice ;
 
  for(i = 0; i < allHubspotCompanies.companies.length; i++ ){
 
    if(allHubspotCompanies.companies[i].properties.hs_lastmodifieddate.value > getDataAfterThisDate){
     //do nothing
    }
    else{
      
      indexToSplice = i+1;
      break;
    }
  }
 
  // Get all elements up until the splice index
  // This Shortens our array and makes writing data to our sheet less time consuming
  allHubspotCompanies.companies.splice(indexToSplice);
 
 
}
 
 
 
 
/*
 * This function writes all of our Hubspot Company Object Data to our Google Sheet.
 * * @param {object} activeSheetBeingModified - A object which represents which google Sheet Tab we will be changing.
 * @param {object} allHubspotCompanies - An array of Hubspot Company Objects
 * @return: void
*/
function display_companies_on_active_sheet(activeSheetBeingModified,allHubspotCompanies){
  
  // activate() and clear() our google Sheet functions, see documentation*
  activeSheetBeingModified.activate();
  activeSheetBeingModified.clear();
  
  var companyPropertyNames = get_all_company_property_names();
  set_column_header_names(activeSheetBeingModified);
   
 
 
  // In this section we are looping through all the rows and columns, and displaying each company's data
  // i == row and j == column
 
  for(i = 0; i < allHubspotCompanies.companies.length; i++){
    for(j =0; j < companyPropertyNames.length; j++){
 
      // We always have to check if a property exists, this is due to custom properties. Also the code fails if we try to access a null value
      if(allHubspotCompanies.companies[i].properties[companyPropertyNames[j]]==null){
       // do nothing we have a null value
      }
      else if(allHubspotCompanies.companies[i].properties[companyPropertyNames[j]]['value'] !== null){
 
        activeSheetBeingModified.getRange(i+2,j+1).setValue(allHubspotCompanies.companies[i].properties[companyPropertyNames[j]]['value']);
       
      }
     }
   }
 
}
 
 
 
 
/*
 * This function returns an array of Company Property Names as string. The order of the array determines how data is displayed on the Google sheet.
 * @return{object} An array of Strings which represent Company Object property names
*/
function get_all_company_property_names(){
  
  // The order in which the Property Names are pushed into the array, is the order in which they will be displayed in google Sheets
  // This array is used to loop through all the properties of a company object
  // All of these are default Hubspot properties, You can add/create company properties in hubspot and append them here.
  var propertyNames = [];
 
  propertyNames.push("name");
  propertyNames.push("hs_lastmodifieddate");
  propertyNames.push("hs_target_account");
  propertyNames.push("hs_lead_status");
  propertyNames.push("total_revenue");
  propertyNames.push("zip");
  propertyNames.push("domain"); 
  propertyNames.push("numberofemployees");
  propertyNames.push("hs_num_decision_makers");
  propertyNames.push("hs_num_blockers");
  propertyNames.push("num_associated_deals");
  propertyNames.push("recent_deal_amount");
  propertyNames.push("num_contacted_notes");
  propertyNames.push("city");
 
  return propertyNames;
}
 
 
 
 
/*
 * This function sets the first row of our Google Sheet, which is our header
 * @param {object} activeSheetBeingModified - A object which represents which google Sheet Tab we will be changing.
 * @return: void
*/
function set_column_header_names(activeSheetBeingModified){
  
  var columnNames = get_all_company_property_names();
 
  for(i= 0; i < columnNames.length; i++){
    activeSheetBeingModified.getRange(1,i+1).setValue(columnNames[i]);
  }
 
  //This must be fixed later this should be dynamic
  // This section changes the text-style in the first row of our sheet
  var sheetRange = SpreadsheetApp.getActive().getRange('A1:Z1');
  var pageStyle = SpreadsheetApp.newTextStyle()
  .setFontSize(12)
  .setBold(true)
  .build();
  sheetRange.setTextStyle(pageStyle);
 
}
 
 
 
 
 
 
 
 
 
 

