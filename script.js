document.getElementById('csvFileInput').addEventListener('change', handleFileSelect);

function handleFileSelect(event) {
    const fileInput = event.target;
    const file = fileInput.files[0];

    if (file) {
        const reader = new FileReader();

        reader.onload = function (e) {
            const csvData = parseCSV(e.target.result);
            filterAndCreateSheets(csvData);
        };

        reader.readAsText(file);
    }
}

function parseCSV(csvText) {
    const lines = csvText.split('\n');
    const header = lines[0].split(',');

    return lines.slice(1).map(line => {
        const cells = line.split(',');
        const row = {};

        cells.forEach((cell, index) => {
            const columnName = header[index].trim();

            // Remove double quotes and convert to numeric value for specific columns
            if (columnName === '# of Primary Splitters' || columnName === '# of Secondary Splitters') {
                const numericValue = parseInt(cell.replace(/"/g, '').trim(), 10);
                row[columnName] = isNaN(numericValue) ? '' : numericValue;
            } else {
                row[columnName] = cell.trim();
            }
        });

        return row;
    });
}

const producerName=["TALHA","jEGAN","NIK","AFIQAH","DANIA","NADIA","Azbin"];


function filterAndCreateSheets(csvData) {
    // Identify unique values in the v_plan column
    const uniquePlans = [...new Set(csvData.map(row => row.v_plan))];

    // Create a workbook
    const wb = XLSX.utils.book_new();

    // Common columns to copy to all sheets
    const commonColumns = ['ID', 'Service Group', 'Service Set' ,'Service Area', 'Tier Rating', '# of Primary Splitters', '# of Secondary Splitters', 'Equipment', 'vetro_id', 'v_created_time', 'v_last_edited_time'];

    // Additional headers for other sheets
    const additionalHeaders = {
        'Passthrough': ['Passthrough'],
        'Producer': ['Producer'],
        'Complete Date': ['Complete Date']
    };
    // Ask the user for a custom file name
    const customFileName = prompt("Enter the file name:") + " Handhole Status.xlsx";
    const fileName = customFileName ? customFileName : 'Handhole Status.xlsx';

    // Initialize an array to store the overview data
    const overviewData = [];

    // Add the "Overview" sheet
    const overviewWs = XLSX.utils.json_to_sheet([], { header: ['NO', 'SG', 'Overall', 'Completed', 'No Splitter', 'Remaining', 'Producer', 'Remark',,'PRODUCER','No.SG(completed)','No.HH(completed)'] });
    XLSX.utils.book_append_sheet(wb, overviewWs, 'Overview');

    // Counter for numbering sheets
    let sheetNumber = 1,index_=0;
    // Iterate over each unique v_plan and filter data
    uniquePlans.forEach(plan => {
        // Check if plan is not undefined or null
        if (plan !== undefined && plan !== null) {
            // Filter and copy only the specified columns
            const filteredData = csvData
                .filter(row => row.v_plan === plan)
                .map(row => {
                    // Include common columns
                    const newRow = commonColumns.reduce((obj, key) => ({ ...obj, [key]: row[key] }), {});

                    // Add additional headers for the specific sheet
                    Object.keys(additionalHeaders).forEach(header => {
                        if (row[header]) {
                            newRow[header] = row[header];
                        }
                    });

                    return newRow;
                });

            // Check if the filteredData array has at least one row with data before appending the sheet

            if (filteredData.length > 0) {
                // Sanitize the plan name to remove invalid characters
                const sanitizedPlan = plan.replace(/[\\\/?*\[\]:]/g, '_');

                producerformula_top = `=UPPER(`
                producerformula_middle = []
                for(let i = 0;i < producerName.length-1; i++){
                    producerformula_middle.push('IFNA(')
                }
                producerformula_middle = producerformula_middle.join('')
                console.log('producerformula_middle: ',producerformula_middle)

                producerformula_btm = `))`

                //loop to generate producer formula
                let FinalFormula =[],LastFormula=[]
                for(let i=0;i<producerName.length;i++){
                    producerformula = 'VLOOKUP(' + '"' + producerName[i] + '"' +",'" + sanitizedPlan + "'!M2:M150,1,FALSE)"

                    if(i==0){
                        FinalFormula.push(producerformula_top, producerformula_middle ,producerformula + ',')
                    }
                    else if( i < producerName.length -1){
                        FinalFormula.push(producerformula + ')' +  ',')
                    }
                    else{
                        FinalFormula.push(producerformula + producerformula_btm)
                    }

                }

                LastFormula = FinalFormula.join('')
                console.log('LastFormula',LastFormula)
                // Convert filtered data to sheet
                const ws = XLSX.utils.json_to_sheet(filteredData, { header: [...commonColumns, ...Object.keys(additionalHeaders)] });

                // Append the sheet to the workbook with the sanitized plan name
                XLSX.utils.book_append_sheet(wb, ws, sanitizedPlan);

                // Add data for the Overview sheet
                overviewData.push({
                    'NO': sheetNumber++,
                    'SG': `SG0${sheetNumber - 1}`,
                    'Overall': `=COUNTA('${sanitizedPlan}'!A2:A150)`,
                    'Completed': `=COUNTIF('${sanitizedPlan}'!L2:L150,"Y")`,
                    'No Splitter': `=COUNTIF('${sanitizedPlan}'!L2:L150,"N")`,
                    //'Remaining': `=C${sheetNumber}-(D${sheetNumber}+E${sheetNumber})`
                    'Remaining': `=IF(C${sheetNumber} -(D${sheetNumber}+E${sheetNumber})=0,"COMPLETED",C${sheetNumber}-(D${sheetNumber}+E${sheetNumber}))`,
                    'Producer': LastFormula,
                    'Remark':``,
                    '':``,
                    
                });
            }
        }
        index_+=1
    });
    console.log('overviewData: ',overviewData)
    // Update the "Overview" sheet with the final overview data
    //XLSX.utils.sheet_add_json(overviewWs, overviewData, { header: ['NO', 'SG', 'Overall', 'Completed', 'No Splitter', 'Remaining','Producer'] });

    // Add the "TOTAL" row at the bottom of the Overview sheet
    const totalRow = {
        'NO': '',
        'SG': 'TOTAL',
        'Overall': `=SUM(C2:C${overviewData.length + 1})`,
        'Completed': `=SUM(D2:D${overviewData.length + 1})`,
        'No Splitter': `=SUM(E2:E${overviewData.length + 1})`,
        'Remaining': `=SUM(F2:F${overviewData.length + 1})`
    };
    overviewData.push(totalRow);

    //add producer,PRODUCER,NOSG,COMPLETED
    for(i in producerName){
        let index = Number(i)
        SGCompleted_value = `=COUNTIFS($G$2:$G$21,J${index+2},$F$2:$F$21,"COMPLETED")`
        HHCompleted_value = `=SUMIF($G$2:$G$21,J${index+2},$D$2:$D$21)`

        if(index < overviewData.length){
            overviewData[index]['PRODUCER'] = producerName[index]
            overviewData[index]['No.SG(Completed)'] = SGCompleted_value
            overviewData[index]['No.HH(Completed)'] = HHCompleted_value
        }
        else{
            overviewData.push({
                'PRODUCER': producerName[index],
                'No.SG(Completed)' : SGCompleted_value,
                'No.HH(Completed)' : HHCompleted_value
            })
        }
    }

    XLSX.utils.sheet_add_json(overviewWs, overviewData);

    // Save the workbook to a file with the custom file name
    XLSX.writeFile(wb, fileName);
}
