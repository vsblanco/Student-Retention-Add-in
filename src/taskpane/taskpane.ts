/**
 * Handler for worksheet changed events
 * @param event - WorksheetChangedEventArgs containing change details
 */
async function onWorksheetChanged(event: Excel.WorksheetChangedEventArgs): Promise<void> {
  await Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getItem(event.worksheetId);
    const changedRange = worksheet.getRange(event.address);
    
    changedRange.load("values, address");
    await context.sync();
    
    console.log(`Worksheet changed: ${event.changeType}`);
    console.log(`Address: ${changedRange.address}`);
    console.log(`New values:`, changedRange.values);
    
    // Add your custom logic here
    // Example: validate data, update calculations, trigger analysis, etc.
  });
}

/**
 * Register the worksheet change event handler
 */
async function registerWorksheetChangeHandler(): Promise<void> {
  await Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getActiveWorksheet();
    worksheet.onChanged.add(onWorksheetChanged);
    await context.sync();
    
    console.log("Worksheet change handler registered");
  });
}
