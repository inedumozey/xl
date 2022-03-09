const { dialog } = require('electron');
const xlsx = require('xlsx');
const os = require('os');
const path = require('path');

async function excelToJson({window=null, title="Open File", multiple=true}){
    try{
        if(!multiple){
            multiSelections = ''
        }else{
            multiSelections = "multiSelections"
        }
        
        // show an open dialog
        const { filePaths, canceled } = await dialog.showOpenDialog(window, {
            title,
            defaultPath: path.join(os.homedir()),
            properties: [multiSelections]
        })

        if(filePaths && !canceled){
            //loop through the filePaths array
            for(let i=0; i<filePaths.length; i++){
                let filePath = filePaths[i];

                //read each of the file and the workbook
                const wb = xlsx.readFileSync(filePath);
                const sheets = wb.SheetNames
                
                // loop through the sheets and read them from the workbook
                for(let k=0; k<sheets.length; k++){
                    const sheet = sheets[k];
                    const ws = wb.Sheets[sheet]

                    // get data from each of the sheet
                    const data = xlsx.utils.sheet_to_json(ws);
                    return data
                }
            }
        }
    }
    catch(err){
        console.log(err);
        return;
    }
}

module.exports = excelToJson