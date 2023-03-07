/**
 * This function outputs the UPC values that need to be deleted from the UPC Database because 
 * they are associated with a SKU number that no longer exists.
 * 
 * @author Jarren Ralf
 */
function deleted_UPCs()
{
  const sheet = SpreadsheetApp.getActive().getSheetByName("Deleted UPCs - 23 March 2021");
  const adagioData = Utilities.parseCsv(DriveApp.getFilesByName("AllSKUsAdagio_23Mar2021.csv").next().getBlob().getDataAsString());
  adagioData.pop(); // Remove the last row (which are 'Totals')
  const header = adagioData.shift();
  const sku = header.indexOf('Number')
  const skus = adagioData.map(g => [(g[sku].substring(0, 4) + g[sku].substring(5, 9) + g[sku].substring(10)).trim()]);
  const data = Utilities.parseCsv(DriveApp.getFilesByName("BarcodeInput.csv").next().getBlob().getDataAsString());
  const filteredData = data.filter(e => skus.filter(f => e[1].toUpperCase() == f[0]).length == 0);
  sheet.getRange(1, 2, filteredData.length, filteredData[0].length).setNumberFormat('@').setValues(filteredData);
}

/**
 * This function outputs the UPC values (and their associated SKUs) that need to be remarried in the UPC Database because 
 * they do not contain the most recent description.
 * 
 * @author Jarren Ralf
 */
function remarried_UPCs()
{
  const sheet = SpreadsheetApp.getActive().getSheetByName("Remarried UPCs - 23 Mar 2021");
  const adagioData = Utilities.parseCsv(DriveApp.getFilesByName("AllSKUsAdagio_23Mar2021.csv").next().getBlob().getDataAsString());
  adagioData.pop(); // Remove the last row (which are 'Totals')
  const header = adagioData.shift();
  const sku = header.indexOf('Number')
  const description = header.indexOf('Description')
  const skus = adagioData.map(g => [(g[sku].substring(0, 4) + g[sku].substring(5, 9) + g[sku].substring(10)).trim(), g[description]]);
  const data = Utilities.parseCsv(DriveApp.getFilesByName("BarcodeInput.csv").next().getBlob().getDataAsString());
  const upcHeader = data.shift();
  const filteredData = data.filter(e => skus.filter(f => e[1].toUpperCase() == f[0] && e[2] != f[1]).length != 0);
  filteredData.unshift(upcHeader);
  sheet.getRange(1, 2, filteredData.length, filteredData[0].length).setNumberFormat('@').setValues(filteredData);
}

/**
 * This function outputs 
 * 
 * @author Jarren Ralf
 */
function itemsToDelete()
{
  const sheet = SpreadsheetApp.getActive().getSheetByName("Items to Delete");
  const adagioData = Utilities.parseCsv(DriveApp.getFilesByName("ItemToDeleteFromAdagioAndAccess.csv").next().getBlob().getDataAsString());
  adagioData.pop(); // Remove the last row (which are 'Totals')
  const header = adagioData.shift();
  const sku = header.indexOf('Number');
  const rcpt = header.indexOf('Last Rcpt Date');
  const ship = header.indexOf('Last Shipment Date');
  const qty  = header.indexOf('Qty On Hand');
  const created = header.indexOf('Created Date');

  header[sku] = 'Item #';
  header[created] = 'Created Year';

  const data = adagioData.filter(item => {
    item[sku] = (item[sku].substring(0, 4) + item[sku].substring(5, 9) + item[sku].substring(10)).trim();
    item[created] = item[created].substring(6);
    return item[qty] == 0 && item[rcpt] === ' ' && item[ship] === ' ' && item[created] < 2020;
  })

  data.unshift(header);
  sheet.getRange(1, 1, data.length, data[0].length).setNumberFormat('@').setValues(data);
}

/**
 * This function outputs 
 * 
 * @author Jarren Ralf
 */
function duplicateItems()
{
  const sheet = SpreadsheetApp.getActive().getSheetByName("Duplicates - 07 Apr 2021");
  const adagioData = Utilities.parseCsv(DriveApp.getFilesByName("AdagioDatabase_07Apr2021.csv").next().getBlob().getDataAsString());
  adagioData.pop(); // Remove the last row (which are 'Totals')
  const header = adagioData.shift();
  const sku = header.indexOf('Number');
  const descrip = header.indexOf('Description');
  header[sku] = 'Item #';
  const data = adagioData.filter(item => {
    item[sku] = (item[sku].substring(0, 4) + item[sku].substring(5, 9) + item[sku].substring(10)).trim();
    return adagioData.filter( item_ => item[descrip] === item_[descrip]).length > 1;
  })

  data.unshift(header);
  sheet.getRange(1, 1, data.length, data[0].length).setNumberFormat('@').setValues(data);
}


/**
 * This function 
 * 
 * @author Jarren Ralf
 */
function import_SKUs()
{
  const sheet = SpreadsheetApp.getActive().getSheetByName("Sheet9");
  const data = Utilities.parseCsv(DriveApp.getFilesByName("AllSKUsAdagio_23Mar2021.csv").next().getBlob().getDataAsString());
  data.pop();
  const header = data.shift();
  const sku = header.indexOf('Number')
  const skus = data.map(g => [(g[sku].substring(0, 4) + g[sku].substring(5, 9) + g[sku].substring(10)).trim()]);
  skus.unshift(['SKU'])
  sheet.getRange(1, 1, skus.length, skus[0].length).setNumberFormat('@').setValues(skus);
}

function itemsNotCreatedInAccess()
{
  const sheet = SpreadsheetApp.getActive().getSheetByName('Copy of Copy of Access (Kris & Brent)');

  // Adagio 
  const adagioData = Utilities.parseCsv(DriveApp.getFilesByName("AdagioDatabase_30Mar2021.csv").next().getBlob().getDataAsString());
  adagioData.pop();
  const adagioHeader = adagioData.shift();
  const sku = adagioHeader.indexOf('Number');

  // Access
  const accessData = Utilities.parseCsv(DriveApp.getFilesByName("AccessDatabase_31Mar2021.csv").next().getBlob().getDataAsString());
  const accessHeader = accessData.shift();
  const sku_ = accessHeader.indexOf('Item code - ACCPAC')
  
  const data = adagioData.filter(item => {
    item[sku] = (item[sku].substring(0, 4) + item[sku].substring(5, 9) + item[sku].substring(10)).trim();
    return accessData.filter(item_ => item_[sku_] == item[sku]).length == 0;
  });
  data.unshift(adagioHeader);

  sheet.getRange(1, 1, data.length, data[0].length).setNumberFormat('@').setValues(data);
}

/**
 * This function 
 * 
 * @author Jarren Ralf
 */
function adagioImportForConversion()
{
  const sheet = SpreadsheetApp.getActive().getSheetByName("Conversions");
  const csvData = Utilities.parseCsv(DriveApp.getFilesByName("conversions.csv").next().getBlob().getDataAsString());
  const header = csvData.shift();
  const numRows = csvData.length;
  const numCols = csvData[0].length;
  const numberFormats = new Array(numRows).fill(null).map(() => ['@', '@','@', '#.#', '#.#', '#.#', '#.#', '@', '@']);
  numberFormats.unshift(new Array(numCols).fill('@'));
  csvData.sort(sortByCategories).unshift(header);
  sheet.clearContents().getRange(1, 1, numRows + 1, numCols).setNumberFormats(numberFormats).setValues(csvData);
}

/**
 * This function imports the UPC database.
 * 
 * @author Jarren Ralf
 */
function import_UPCs()
{
  const sheet = SpreadsheetApp.getActive().getSheetByName("UPC Database - 23 March 2021");
  const data = Utilities.parseCsv(DriveApp.getFilesByName("BarcodeInput.csv").next().getBlob().getDataAsString());
  sheet.getRange(1, 1, data.length, data[0].length).setNumberFormat('@').trimWhitespace().setValues(data);
}

/**
 * This function imports the UPC database.
 * 
 * @author Jarren Ralf
 */
function import_inventory()
{
  const sheet = SpreadsheetApp.getActive().getSheetByName("Sheet13");
  const data = Utilities.parseCsv(DriveApp.getFilesByName("inventory.csv").next().getBlob().getDataAsString());
  sheet.getRange(1, 1, data.length, data[0].length).setNumberFormat('@').trimWhitespace().setValues(data);
}

function removeInactiveNoTS()
{
  const sheet = SpreadsheetApp.getActiveSheet()
  const range = sheet.getDataRange();
  const values = range.getValues();
  const newValues = values.map(g => {var temp = g[1].split(' - '); g[1] = temp[2] + ' - ' + temp[3]; return g});
  range.setNumberFormat('@').setValues(newValues)
}

/**
* Sorts data by the categories while ignoring capitals and pushing blanks to the bottom of the list.
*
* @param  {String[]} a : The current array value to compare
* @param  {String[]} b : The next array value to compare
* @return {String[][]} The output data.
* @author Jarren Ralf
*/
function sortByCategories(a, b)
{
  return (a[8].toLowerCase() == b[8].toLowerCase()) ? 0 : (a[8] == '') ? 1 : (b[8] == '') ? -1 : (a[8].toLowerCase() < b[8].toLowerCase()) ? -1 : 1;
}