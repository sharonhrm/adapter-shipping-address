let selectedFile;
console.log(window.XLSX);

//fedex sheet
document.getElementById('input').addEventListener("change", (event) => {
    selectedFile = event.target.files[0];
})

//adapter sheet
document.getElementById('inputPrice').addEventListener("change", (event) => {
    selectedFileAdapter = event.target.files[0];
})

let data=[];
let adapterSheetArr=[];
let shippingJsonArr=[];
var fedexJson = "";
var adapterJson = "";


document.getElementById('button').addEventListener("click", () => {    
    XLSX.utils.json_to_sheet(data, 'out.xlsx');

    //first we get fedex sheet
    if(selectedFile){
        document.getElementById("jsondata000").innerHTML = "---- Reading Fedex sheet";
        let fileReader = new FileReader();
        fileReader.readAsBinaryString(selectedFile);
        fileReader.onload = (event)=>{
         let data = event.target.result;
         let workbook = XLSX.read(data,{type:"binary"});
         console.log(workbook);

         var i = 0;
         workbook.SheetNames.forEach(sheet => {
            if(i == 0){
              let rowObject = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheet]);
              //console.log(rowObject);
              fedexJson = JSON.stringify(rowObject,undefined,4); 
              shippingJsonArr = JSON.parse(fedexJson);
              //console.log(fedexJson);
              //document.getElementById("jsondata").innerHTML = JSON.stringify(rowObject,undefined,4)
              i++;
            }
         });
        }
    }

    if(selectedFileAdapter){
        document.getElementById("jsondata00").innerHTML = "---- Reading Adapter sheet";  
        let fileReader = new FileReader();
        fileReader.readAsBinaryString(selectedFileAdapter);
        fileReader.onload = (event)=>{
         let data = event.target.result;
         let workbook = XLSX.read(data,{type:"binary"});
         console.log(workbook);

         var i = 0;
         workbook.SheetNames.forEach(sheet => {
            if(i == 0){
              let rowObject1 = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheet]);
              //adapterSheetArr.push(rowObject1);
              adapterJson = JSON.stringify(rowObject1,undefined,4);    
              adapterSheetArr = JSON.parse(adapterJson); 
              i++;
            }
         });
        }
    }
    
    document.getElementById("jsondata0").innerHTML = "<span style='font-size:24px;'><b>---- CLICK ON SUBMIT AGAIN</b></span>";  
    let shipmentArr=[];
    let addedNames=[];
    let found = false;

    adapterSheetArr.forEach(function(data3, index2) {  
        found = false;      
        shipmentArr[index2] = [];

        if(addedNames.includes(data3['OrderID'])){
            found = true;
            shipmentArr[index2]['OrderID'] = data3['OrderID'];
            shipmentArr[index2]['warehouseNotes'] = "Duplicate";
            shipmentArr[index2]['Qty'] = data3['Qty'];
        }else{
            shippingJsonArr.forEach(function(data1, index) { 
                if(data3['OrderID'] == data1['Name']){
                    if(data1['Shipping Name'] != "" && data1['Shipping Name'] != null && data1['Shipping Name'] != undefined){
                        console.log(data1['Shipping Name']);
                        addedNames.push(data1['Name']);                        
                        shipmentArr[index2]['OrderID'] = data1['Name'];
                        shipmentArr[index2]['warehouseNotes'] = data3['Notes'];
                        shipmentArr[index2]['Qty'] = data3['Qty'];
                        shipmentArr[index2]['email'] = data1['Email'];
                        shipmentArr[index2]['shippingName'] = data1['Shipping Name'];
                        shipmentArr[index2]['shippingStreet'] = data1['Shipping Street'];
                        //shipmentArr[index2]['shippingAddress1'] = data1['Shipping Address1'];
                        //shipmentArr[index2]['shippingAddress2'] = data1['Shipping Address2'];
                        //shipmentArr[index2]['shippingCompany'] = data1['Shipping Company'];
                        shipmentArr[index2]['shippingCity'] = data1['Shipping City'];
                        shipmentArr[index2]['shippingZip'] = data1['Shipping Zip'];
                        shipmentArr[index2]['shippingProvince'] = data1['Shipping Province'];
                        shipmentArr[index2]['shippingCountry'] = data1['Shipping Country'];
                        shipmentArr[index2]['shippingPhone'] = data1['Shipping Phone'];
                        shipmentArr[index2]['orderNotes'] = data1['Notes']; 
                        found = true;
                    }
                }             
            });
        }      

        if(found == false){
            shipmentArr[index2]['OrderID'] = data3['OrderID'];
            shipmentArr[index2]['warehouseNotes'] = "Not Found";
        }
    });

    if(shipmentArr.length > 0){
        document.getElementById("jsondata4").innerHTML = "---- Parsing things into excel";
        var priceFilename='adapterShipmentChina.xlsx';
        
        var ws = XLSX.utils.json_to_sheet(shipmentArr);
        var wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "shipment");
        XLSX.writeFile(wb,priceFilename);

        document.getElementById("jsondata5").innerHTML = "---- All done!";
    }

});