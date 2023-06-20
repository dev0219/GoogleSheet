
const SheetHandle = require("./sheet").SheetHandle;
var cron = require('node-cron');

async function startCronJob() {
    cron.schedule('*/15 * * * *', async () => {
        try {
        await SheetHandle();
        console.log('Sheet data updated successfully.');
        } catch (error) {
        console.error('Error updating sheet data:', error);
        }
    });
  
}
startCronJob();
