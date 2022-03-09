# @mozeyinedu/xl

## Installation

`npm i @mozeyinedu/xl`

## Description

xl is an electron package module that converts json file to excel and save in a specified directory or current directory. It also read excel file and convert to json data.

### Has two methods

...

    xl.jsonToExcel();
    // converts json to excel

...

...

    xl.excelToJson()
    // converts excel to json

...

### Converting JSON to Excel

...

    xl.jsonToExcel();

...

It receives an object with 6 properties

1. window: this is the app window, its optional,

2. data: the actual json data. default is set to an empty array. if provided, it must be an array of object or object or json data otherwise throws an error

3. save: true/false. default is false.

   - If set to true tells the app to open a save dailog and save the excel file on the user Documents by default or save to a specified directory seen below.
   - If set to false or not specified, a directory must then be provided as discussed below.

4. dir: directory where the excel file will be saved.

   - If the 'save' property as seen above is set to false or not specified. then directory must be provided, it throws an error when empty or when directory is not valid.
   - If the 'save' property is set to true, default directory is Documents. It throws an error if the directory is set to a directory that is invalid. You can use the 'os' node module to set a valid directory, require('path').join(require('os').homedir(), 'directory name e.g Desktop, Downloads, Documents etc').

5. filename: default is 'file'. No need you include extension name, since by default, the user specified extension is removed (whatever or however the extension is) and replaced with excel extension.

6. title: String that will appear at the top of the save dialog. Default is 'Saving New Excel File'

### example

...

    //opens save dialog and to Desktop

    async function exportExcel1(win){
        let filepath = await xl.jsonToExcel({
            window: win,
            data,
            save: true,
            title: 'Save',
            dir: require('path').join(require('os').homedir(), "Desktop"),
            filename: 'inventory'
        });

        if(filepath){
            console.log(filepath);
        }
    }

    //the function is called and win is passed to it as a paramter in app.on('ready, ()=>{})

...

...

    //saves in the app current directory

    async function exportExcel2(){
        let filepath = await xl.jsonToExcel({
            data,
            dir: path.join(__dirname),
            filename: 'inventory'
        });

        if(filepath){
            console.log(filepath);
        }
    }

...



### Converting Excel to JSON

...

    xl.excelTojson()

...

It receives an object with two properties of which all are optional
1. window: window: this is the app window, its optional,
2. title: String that will appear at the top of the save dialog. Default is 'Open File',
3. multiple: true/false. Default is true ans allows multiple file selection but when change to false will allow only sinlge file selection 

If non excel file is selected, JSON data will not be produced but still wont block your code, it runs asynchronously
The function returns promise with the data in array

### examples
...

    // no any property in the object parameter
    async function importExcel(){
        const data = await xl.excelToJson({});
        console.log(data);
    }

...

...

    //multiple set to false
    async function importExcel(window){
        const data = await xl.excelToJson({
            multiple: false,
            window
        });
        console.log(data);
    }

...

...

    //title set
    async function importExcel(window){
        const data = await xl.excelToJson({
            multiple: true,
            title: 'Exporting Excel file',
            window
        })
        console.log(data);
    }

...

#### Reach me;
* whatsApp/Call: +2348036000347
* [Facebook](https://www.facebook.com/mozey.inedu.3)