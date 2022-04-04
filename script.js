// this is where we begin to write code

// The function below is used to validate a excel file once it has been uploaded.

function upload() {
    var files = document.getElementById('file_upload').files;
    if (files.length==0) {
        alert("Please choose any file...");
        return;
    }

    var filename = files[0].name;
    var extension = filename.substring(filename.lastIndexOf(".")).toUpperCase();

    if (extension == '.XLS' || extension == '.XLSM') {
        excelFileToJSON(files[0]);
    }
    else {
        alert("Please select a valid excel file.");
    }
}

// Method to read an excel file and convert data to JSON, which is a step after validation

function excelFileToJSON(file) {

    //Inside excelFileToJSON(), we have read the data of the excel file by using a file reader as a binary string using readAsBinaryString() method.
   
    try {
        var reader = new FileReader();
        reader.readAsBinaryString(file);
        reader.onload = function(e) {
            var data = e.target.result;
            var workbook = XLSX.read(data, {
                type : 'binary'
            });
            var result = {};
            workbook.SheetNames.forEach(function(sheetName){
                var roa = XLSX.utils.sheet_to_row_object_arrray(workbook.Sheets[sheetName]);
                if (roa.length > 0) {
                    result[sheetName] = roa;
                }
            });
            //displaying the json result
            var resultEle=document.getElementById("json-result");
            resultEle.value=JSON.stringify(result, null, 4);
            resultEle.style.display='block';
            }
    }catch(e) {
        console.error(e);
    }
}

// The method below is to display the data in HTML table
function displayJSONtoHTMLTable(jsonData) {
    var table = document.getElementById("display_excel_data");
    if (jsonData.length > 0) {
        var htmlData='<tr><th>coins</th><th>Bin Price</th><th>landing</th><th>final output</th><th>to bank</th><th>diff</th><th>diff%</th><th>buy</th><th>sell</th><th>Column</th><th>spread</th><th>coin qty/th></tr>';
            for(var i=0;i<jsonData.length;i++){
                var row=jsonData[i];
                htmlData+='<tr><td>'+row["coins"]+'</td><td>'+row["landing"]+'</td><td>'+row["final output"]+'</td><td>'+row["to bank"]+'</td><td>'+row["diff"]+'</td><td>'+row["diff%"]+'</td><td>'+row["Buy"]+'</td><td>'+row["Sell"]
                      +'</td><td>'+row["Column"]+'</td><td>'+row["spread"]+'</td></tr>'+row["coin qty"]+'</td></tr>';
            }
            table.innerHTML=htmlData;
        }
        else{
            table.innerHTML='There is no data in Excel';
        }
}