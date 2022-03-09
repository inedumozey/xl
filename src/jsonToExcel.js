const xlsx = require('xlsx');
const path = require('path');
const os = require('os');
const fs = require('fs');
const { dialog } = require("electron");
const check = require('@mozeyinedu/check');

console.log(check);

async function jsonToExcel({
    data=[],
    save=false,
    dir=null,
    filename="file",
    window=null,
    title="Saving New Excel File"
}){
   
    try{
        if(check.isObject(data) || check.isArray(data)){
           
            const wb = xlsx.utils.book_new();
            const ws = xlsx.utils.json_to_sheet(data);
            xlsx.utils.book_append_sheet(wb, ws, "sheet1");

            if(save){
                let directory;

                if(!dir){
                    directory = path.join(os.homedir(), 'Documents')
            
                }else{
                    if(!fs.lstatSync(dir).isDirectory()){
                        throw new Error('Invalid directory')
                    }else{
                        directory = dir;
                    }
                }

                if(directory){
                    
                    //open save dialog
                    const { filePath, canceled } = await dialog.showSaveDialog(window, {
                        title,
                        defaultPath: path.join(directory, `${filename.split('.')[0]}.xlsx`),
                        filters: [
                            { name: "Excel", extensions: ["xlsx", "xls", "xlsm"] },
                            { name: "Plain Text", extensions: ["txt"] }
                        ]
                    });
                    
                    if(filePath && !canceled){
                        xlsx.writeFile(wb, filePath);
                        return "Path: " + filePath;
                    }
                }

            }else{
                if(!dir || !fs.lstatSync(dir).isDirectory()){
                    throw new Error('Invalid directory')
            
                }else{
                    //save the new excel file to the path provided
                    let filename_ = `${dir}/${filename.split('.')[0]}.xlsx`;
                    xlsx.writeFile(wb, filename_)
                    return "Path: " + filename_;
                }
            }
        }else{
            throw new Error("Invalid data")
        }
    }
    catch(err){
        console.log(err);
        return;
    }

}

module.exports = jsonToExcel