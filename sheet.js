const { GoogleSpreadsheet } = require('google-spreadsheet');



async function SheetHandle() {
    const doc = new GoogleSpreadsheet('1Jy0BA29yyAk6DE7cnWnTRDDmwgSgt61Y5m6_oIeJ1U0');
 
    // useServiceAuth //
    try {
      await doc.useServiceAccountAuth({
        client_email: 'sales-176@fifth-audio-389203.iam.gserviceaccount.com',
        private_key: '-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQDmgLa0rjqQs4jX\nukiWs8VhXsNy399OtxY36V154WZVIdi+565aJlFtWgKKLJY+B2Dh7yQm8pY7FRTy\npLcM9ln0J7jhDWNrMXAEgV9emWYTUMAOGkOesQBKmhfef+kD9Dj+GAE6GyTWM7+f\nIKP+4OQlXcC/vznzRELO5eDHF/voa62IBdvJavqcXsl/S+8wgJ3zYLf1FUZSvdWu\nSBGF7QEDAU8KjuRShkLht/1CGW5dWiK8d33n0jV3s7vIruewzKpjUot3CF9QFqnm\n1bj+ICcKHHLvJl6/sFUJ1saa/A7qzZdg1TAMjxaIKELqnTy5+EDpmKPqw9wsM69R\nVjodhi9LAgMBAAECggEAAJZ8HIfZN5Yvwic3Q8ZSypsch0pE3kD9JMiPGxF3OHoX\nr+E2+gN6mllBamcPX4+oDlP7VQVLn1F72GRvRxKVwvFOMMur4Ob4gdQIxI/+HpUG\nNBhWQHnzv4E5OwbD0KEZjGAS5Yb43FgNCnO27K+LEjSGItkumLfj6dwYnpYpZ7My\nKPEM0pISu5lVBEnkdSWRufNZsIbJVPX826nmbYBvQIJwQlkoOHFb7pr7gBFdiiGy\nHQlwweoHnWyWA/uRYS+labaAZPiRpuTOQt2RQwLqfRjq8qOb3RUPbMjrhu7urauH\nN1Ny0og+dZ7+oa7saO/JYJCclfcKJiKXIHV3D49bAQKBgQD2sH0Pkuctrd5THYZt\nAvUWxiIXAEYkF7P0ukSLpn6ulEgExUI6y441KPslkelUzadFXXEE8DVqN65vVNg1\nF+0bcKMLRceI0I65jhDcQVTXYOoH/8oD4b33/eskHxND8Nz+L7AREeOwyJICFE4l\nKRe+R0S5H1uy92L3MgRsua+5gQKBgQDvM9SE59eumGweH6N8F9sQpJnBT5MATqpf\nfvP2tMEA062YTlzYF2O4aJtkOOuR6EeTcRFoSj4J3ZTPaSRcsAOBBWa7t0h2G8eF\n1bZ5NWzN9KY/gB9wlHBpiZrlL1W7CvuYzXwZZHXM4tXkTfQRWNJYoylpP16XhzZU\nfFugUb0WywKBgQCm49VnoNm3RIZBMptLypFmu9o6N1w8dLCxIwbWd2gC0Zw7Zfwd\nbLKjcjseOq1hStQcEFAyqBNq7YqcFQlEOhCV3MjhCm8oNvgnvl3XtHciUpVhngHc\nSG9Ng8H9qOxOrXfEmxyBH9orMjXdJEGN0FQYzXxHxVgzJjwUwgjOSX0BgQKBgGu1\nRO6LjrzZeWWfDXhlLYky9ODsud6bjW/utF/USEvdBP/d2UScU5TH1aCtWLWciA5G\nDXaOZ5z9n3I9f9gUkZ9ZFUdVYlV8cL083Ct7+QBMN0fEo2OIE44SHiMwy0Or0Fqf\nvE+awsh9I57n0wy0mBK9dXokxK0qfzZPwNpRs/k9AoGAfYQxLXDhEenG+y9ZYUyC\nZL4pDqHA/srB6JDgxZrrKFWNxHEPoApqYzjDcsw+mJmfZNSYYq3FJAWuc4svpstm\nvQfAU1JfQbFB8xFeghzi0KohLNMV2y6+Vrsxo7jN9P2WomisEtc+1o4F4DW03bUO\naTZ43DF2W5v4RAKRa7qwmeo=\n-----END PRIVATE KEY-----\n',
    });

    // Load the spreadsheet
    await doc.loadInfo();
    } catch (error) {
      console.log("-- console error")
      console.log(error)      
    }
    

    // Get sheet data about each tab //
    const preSales_sheet = doc.sheetsByTitle['PreSales']; //  PreSals Sheet
    const leads_sheet = doc.sheetsByTitle['Leads'];//  Leads Sheet
    const intros_sheet = doc.sheetsByTitle['Intros'];//  Intros Sheet
    const members_sheet = doc.sheetsByTitle['Members'];//  Members Sheet
    const declined_sheet = doc.sheetsByTitle['Declined'];//  Declined Sheet   

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
        headerRow[1] = 'PreSales ID';
        const jsonData = preSales_rows.map(row => {
          const rowData = Object.values(row._rawData);
          return Object.fromEntries(headerRow.map((header, index) => [header, rowData[index]]));
        });
        
      
        jsonData.forEach((json_row) => {
          delete json_row['Entry ID'];
        })

        // Declare the data vars for each tab //
        let membersJsonData = [];
        let declinedJsonData = [];
        let introsJsonData = [];
        let leadsJsonData = [];

        // Push data with filtered data //
        headerRow.shift()
        jsonData.forEach((row, index) => {
            if ( row['Deposit Paid'] == 'Yes' ) {
              membersJsonData.push(row)
             }
            else if ( row['Attended Event/Workout'] == 'Yes' ) {
            introsJsonData.push(row)
           }  else if ( row['Declined Membership'] == 'Yes' || (row['DNC'] == 'YES' && row['DNT'] == 'Y')) {
            declinedJsonData.push(row)
           } else {
            leadsJsonData.push(row)
           }
        });

        // Push data on google sheet //
        await leads_sheet.clear();
        await leads_sheet.setHeaderRow(headerRow);
        const leadSheet_update = await leads_sheet.addRows(leadsJsonData);
        await intros_sheet.clear();
        await intros_sheet.setHeaderRow(headerRow);
        const introsSheet_update = await intros_sheet.addRows(introsJsonData);
        await members_sheet.clear();
        await members_sheet.setHeaderRow(headerRow);
        const membersSheet_update = await members_sheet.addRows(membersJsonData);
        await declined_sheet.clear();
        await declined_sheet.setHeaderRow(headerRow);
        const declinedSheet_update = await declined_sheet.addRows(declinedJsonData);

        return 'Updated Successfully!';

      } catch (error) {
        
        // error Handling //

        console.error('Error fetching rows:', error);
        throw error;

      }
}
  
  module.exports = { SheetHandle };