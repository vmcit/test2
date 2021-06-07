let selectedFile;
console.log(window.XLSX);
document.getElementById('input').addEventListener("change", (event) => {
    selectedFile = event.target.files[0];
})

let datajson=[{
    "name":"jayanth",
    "data":"scd",
    "abc":"sdef"
}]


document.getElementById('button').addEventListener("click", () => {
    XLSX.utils.json_to_sheet(datajson, 'out.xlsx');
    if(selectedFile){
        if (!isExcelFile(selectedFile.name)) {
            toastr.error('File không đúng định dạng.');
            return;
        }
        let fileReader = new FileReader();
        fileReader.readAsBinaryString(selectedFile);
        fileReader.onload = (event)=>{
         let data = event.target.result;
         let workbook = XLSX.read(data,{type:"binary"});
         console.log(workbook);

             var header = ["STT","certificateId", "scoreExecute", "scoreTheory",
                        "certificateCode", "certificateName", "field", "activeDate", "exprireDate", "scanPath",
                        "status"];
                    var XL_row_object = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], {
                        header: header,
                        range: 1
                    });

                    var isValidDataStruct = true;
                    XL_row_object.forEach(function (data) {
                                Object.keys(data).forEach(function (key) {
                                    if (key == 'undefined') isValidDataStruct = false;
                                })
                                data.status = 1;
                                data.certificateId = 0;
                            });
                    if (!isValidDataStruct) {
                        toastr.error('Cấu trúc file excel không hợp lệ.');
                        return;
                    }
                
         workbook.SheetNames.forEach(sheet => {
              let rowObject = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheet]);
              rowObject.forEach(function(currentItem) {
                delete currentItem["STT"];
            });
              console.log(rowObject);
              document.getElementById("jsondata").innerHTML = JSON.stringify(rowObject,undefined,4)
         });
         reader.onerror = function (event) {
            console.error("File could not be read! Code " + event.target.error.code);
        };
        reader.readAsBinaryString(selectedFile);
        }
    }
});

// check dinh dang file
function isExcelFile(fileName) {
    var validExts = new Array(".xlsx", ".xls");
    var fileExt = fileName.substring(fileName.lastIndexOf('.'));
    if (validExts.indexOf(fileExt) < 0) return false;
    else return true;
}

// sua code lan 1 nhanh master
// sua lan 4 cuongvm
// sua lan 3 vmcit94
<<<<<<< HEAD
// sua lan 6 vmcit94
=======
// sua lan 5 cuongvm
>>>>>>> cuongvm
