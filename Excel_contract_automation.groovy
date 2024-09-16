//Excel Automation script
function main(workbook: ExcelScript.Workbook) {

    let sheet = workbook.getActiveWorksheet();




    // Get the used range of the worksheet

    let usedRange = sheet.getUsedRange();

    let usedValues = usedRange.getValues();




    // Initialize lastRow

    let lastRow: number = 0;




    // Find the last non-empty cell in column C (column index 2)

    for (let i = usedValues.length - 1; i >= 0; i--) {

        if (usedValues[i][2] !== null && usedValues[i][2] !== "") {

            lastRow = i + 1; // Adding 1 because arrays are 0-based

            break;

        }

    }




    // Check if any data was found in column C

    if (lastRow < 2) {

        console.log("No data available in column C.");

        return;

    }




    // Get the ranges for columns C, D, and E based on lastRow

    let contractStartDatesRange = sheet.getRange(`C2:C${lastRow}`);

    let contractEndDatesRange = sheet.getRange(`D2:D${lastRow}`);

    let cancellationPeriodsRange = sheet.getRange(`E2:E${lastRow}`);




    // Get the values from the ranges

    let contractStartDates: (string | number)[][] = contractStartDatesRange.getValues();

    let contractEndDates: (string | number)[][] = contractEndDatesRange.getValues();

    let cancellationPeriods: (string | number)[][] = cancellationPeriodsRange.getValues();




    console.log("Contract Start Dates:", contractStartDates);

    console.log("Contract End Dates:", contractEndDates);

    console.log("Cancellation Periods:", cancellationPeriods);




    // Iterate through the rows

    for (let i = 0; i < contractStartDates.length; i++) {

        let contractStartDate = contractStartDates[i][0];

        let contractEndDate = contractEndDates[i][0]; // Excel serial number

        let cancellationPeriod = cancellationPeriods[i][0]; // Cancellation period in month




        // Convert Excel date serial number to JavaScript date

        let contractEndDateParsed = new Date((contractEndDate - 25569) * 86400 * 1000);




        // Check for valid data (skip empty rows)

        if (contractEndDate && cancellationPeriod) {

            let cancellationPeriodParsed: number = parseInt(cancellationPeriod.toString());




            // Calculate the check date: contract end date minus cancellation period minus 1 month

            let checkDate: Date = new Date(contractEndDateParsed);

            checkDate.setMonth(checkDate.getMonth() - cancellationPeriodParsed - 1);




            // Today's date

            let today: Date = new Date();




            // If today's date is on or after the check date, update status

            if (today >= checkDate) {

                sheet.getRange(`F${i + 2}`).setValue("Handlung erforderlich");

            } else {

                sheet.getRange(`F${i + 2}`).setValue("laufend");

            }

        }

    }

}
