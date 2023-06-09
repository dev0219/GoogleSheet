const { GoogleSpreadsheet } = require('google-spreadsheet');



async function SheetHandle() {
    const doc = new GoogleSpreadsheet('1FWZvqaNskUxkFLQP8_vxqL9RuFlBnWo-EgfIL_HqeF4');
 
    // useServiceAuth //
    await doc.useServiceAccountAuth({
        client_email: '',
        private_key: '',
    });

    // Load the spreadsheet
    await doc.loadInfo();

    // Get sheet data about each tab //
    const preSales_sheet = doc.sheetsByTitle['PreSales']; //  PreSals Sheet
    const leads_sheet = doc.sheetsByTitle['Leads'];//  Leads Sheet
    const intros_sheet = doc.sheetsByTitle['Intros'];//  Intros Sheet
    const members_sheet = doc.sheetsByTitle['Members'];//  Members Sheet
    const declined_sheet = doc.sheetsByTitle['Declined'];//  Declined Sheet

    // Clear Sheets //
    await leads_sheet.clear();
    await intros_sheet.clear();
    await members_sheet.clear();
    await declined_sheet.clear();

    try {
         // Get the rows from the PreSales sheet
        const preSales_rows = await preSales_sheet.getRows(); // Get the rows from the PreSales sheet
    
        await leads_sheet.resize({
          rowCount: preSales_rows.length + 1, // Add 1 for the header row
          columnCount: preSales_sheet.columnCount,
        });
        await intros_sheet.resize({
            rowCount: preSales_rows.length + 1, // Add 1 for the header row
            columnCount: preSales_sheet.columnCount,
          });
        await members_sheet.resize({
            rowCount: preSales_rows.length + 1, // Add 1 for the header row
            columnCount: preSales_sheet.columnCount,
        });
        await declined_sheet.resize({
            rowCount: preSales_rows.length + 1, // Add 1 for the header row
            columnCount: preSales_sheet.columnCount,
        });
    
        let headerRow = preSales_sheet.headerValues;
        headerRow[0] = 'PreSales ID';
        await leads_sheet.setHeaderRow(headerRow);
        await intros_sheet.setHeaderRow(headerRow);
        await members_sheet.setHeaderRow(headerRow);
        await declined_sheet.setHeaderRow(headerRow);
    
        const jsonData = preSales_rows.map(row => {
          const rowData = Object.values(row._rawData);
          return Object.fromEntries(headerRow.map((header, index) => [header, rowData[index]]));
        });
        let headerLine = {}
        headerRow.forEach((header) => {
            if (header == '') {
                headerLine['PreSales ID'] = 'PreSales ID';
            } else {
                headerLine[header] = header;
            }
           
        })

        // Declare the data vars for each tab //
        let membersJsonData = [];
        let declinedJsonData = [];
        let introsJsonData = [];
        let leadsJsonData = [];

        // Push data with filtered data //
        jsonData.forEach((row, index) => {
            row[Object.keys(row)[0]] = index+2;
           if ( row['Attended Event/Workout'] == 'Yes' || row['Attended Event/Workout'] == 'No' ) {
            introsJsonData.push(row)
           } else if ( row['Deposit Paid'] == 'Yes' ) {
            membersJsonData.push(row)
           } else if ( row['Declined Membership'] == 'yes' || (row['DNC'] == 'yes' && row['DNT'] == 'yes')) {
            declinedJsonData.push(row)
           } else {
            leadsJsonData.push(row)
           }
        });

        // Push data on google sheet //
        const leadSheet_update = await leads_sheet.addRows(leadsJsonData);
        const introsSheet_update = await intros_sheet.addRows(introsJsonData);
        const membersSheet_update = await members_sheet.addRows(membersJsonData);
        const declinedSheet_update = await declined_sheet.addRows(declinedJsonData);

        return 'Updated Successfully!';

      } catch (error) {
        
        // error Handling //

        console.error('Error fetching rows:', error);
        throw error;

      }
}
  
  module.exports = { SheetHandle };