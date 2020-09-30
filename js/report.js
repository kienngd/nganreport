"use strict";
var data = null;
var bLoaded = false;
function ProcessData() {    
    // Set status text
    setMessage("Processing...");
    // Set button disable
    setButtonDisabled(true);
    // Load file data
    LoadFileData();
    // Process data
    ProcessExcel();
    /* if( bLoaded == true) {
        // Process data
        // alert("process excel data here");
        bLoaded = false;
        
        // Set status text
        setMessage("Finished");
        // Set button disable
        setButtonDisabled(false);
    } else {
        // Waiting for one second
        setTimeout(ProcessData, 1000);
    } */
}

function LoadFileData() {
    let fileUpload = document.getElementById("fileUpload");
    let regex = /^([a-zA-Z0-9\s_\\.\-:])+(.xls|.xlsx)$/;
    if (regex.test(fileUpload.value.toLowerCase())) {
        if (typeof (FileReader) != "undefined") {
            var reader = new FileReader();
            //For Browsers other than IE.
            if (reader.readAsBinaryString) {
                reader.onload = function (e) {
                    /* data = e.target.result;
                    bLoaded = true; */
                };

                reader.onloadstart = function () {
                    // alert("Load start");
                };

                reader.onprogress = function () {
                    // alert("Progress");
                };

                reader.onerror = function () {
                    alert("Error when load data");
                };

                reader.onloadend = function (e) {
                    // alert("Load end");
                    data = e.target.result;
                    bLoaded = true;
                };
                reader.readAsBinaryString(fileUpload.files[0]);
            } else {
                //For IE Browser.
                reader.onload = function (e) {
                    var tmpData = "";
                    var bytes = new Uint8Array(e.target.result);
                    for (var i = 0; i < bytes.byteLength; i++) {
                        tmpData += String.fromCharCode(bytes[i]);
                    }
                    data = tmpData;
                    bLoaded = true;
                };
                reader.readAsArrayBuffer(fileUpload.files[0]);
            }

        } else {
            alert("This browser does not support HTML5.");
        }
    } else {
        alert("Please upload a valid Excel file.");
    }
}

function ProcessExcel() {
    if( bLoaded !== true) {
        // Waiting for one second
        setTimeout(ProcessExcel, 500);
        return(false);
    }
    bLoaded = false;
    var workbook = XLSX.read(data, {
        type: 'binary'
    });
    // Read the first sheet
    var tmpSheetName = workbook.SheetNames[0];
    var productionData = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[tmpSheetName]);

    // Read the 2nd sheet
    tmpSheetName = workbook.SheetNames[1];
    var boomListData = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[tmpSheetName]);
    
    setMessage("Generating report...");
    // Generate the 3rd sheet's data: product sum            
    let productReportData = processProductionData(productionData);            
    
    // Generate the 4th sheet's data: boom list with sum
    let boomList = preProcessMaterialData(boomListData);            
    let materialReportData = processMaterialData(productReportData, boomList);

    // Create new sheet
    tmpSheetName = "Report1";
    let tmpData = [["Ngay xuat", "Ma thanh pham", "So luong"]];
    for(const [index, value] of productReportData) {
        tmpData.push([value["ngayxuat"], value["mathanhpham"], value["soluong"]]);                
    }
    let tmpSheet = XLSX.utils.aoa_to_sheet(tmpData);
    XLSX.utils.book_append_sheet(workbook, tmpSheet, tmpSheetName);

    tmpSheetName = "Report2";
    tmpData = [["Ngay xuat", "Ma thanh pham", "Ma linh kien", "So luong"]];
    for(const [index, value] of materialReportData) {
        tmpData.push([value["ngayxuat"], value["mathanhpham"], value["malinhkien"], value["soluong"]]);
    }
    tmpSheet = XLSX.utils.aoa_to_sheet(tmpData);
    XLSX.utils.book_append_sheet(workbook, tmpSheet, tmpSheetName);
    XLSX.writeFile(workbook, "report.xlsx");
    // Set status text
    setMessage("Finished");
    // Set button disable
    setButtonDisabled(false);
}

function processProductionData(productionData) {
    var retData = new Map();
    var nCount = productionData.length;
    for (var i = 0; i < nCount; i++) {
        let tmpNgayXuat = ("ngày xuất hàng" in productionData[i]) ? productionData[i]["ngày xuất hàng"] : '-';
        tmpNgayXuat = tmpNgayXuat.trim();
        let tmpMaThanhPham = ("mã thành phẩm" in productionData[i]) ? productionData[i]["mã thành phẩm"] : '-';
        tmpMaThanhPham = tmpMaThanhPham.trim();
        let tmpSoLuong = ("số lượng" in productionData[i]) ? productionData[i]["số lượng"] : "0";
        tmpSoLuong = tmpSoLuong.trim();
        let tmpKey = '' + tmpNgayXuat + '@' + tmpMaThanhPham;
        let tmpObj = new Object();
        if( retData.has(tmpKey)) {
            // key is exist, update old data
            tmpObj = retData.get(tmpKey);
            if("soluong" in tmpObj) {
                tmpObj.soluong = tmpObj.soluong + parseFloat(tmpSoLuong);
            } else {
                tmpObj.soluong = parseFloat(tmpSoLuong);
            }

        } else {
            // key is not exists so create new
            tmpObj = {
                "ngayxuat": tmpNgayXuat, 
                "mathanhpham": tmpMaThanhPham, 
                "soluong": parseFloat(tmpSoLuong),
            }
        }
        retData.set(tmpKey, tmpObj);
    }
    return(retData);
}

function preProcessMaterialData(boomList) {
    let retData = new Map();
    let nCount = boomList.length;
    let nAlert = 5;
    for (var i = 0; i < nCount; i++) {
        let tmpMaThanhPham = ("mã thành phẩm" in boomList[i]) ? boomList[i]["mã thành phẩm"] : "-";
        tmpMaThanhPham = tmpMaThanhPham.trim();
        let tmpMaLinhKien = ("mã linh kiện" in boomList[i]) ? boomList[i]["mã linh kiện"] : "-";
        tmpMaLinhKien = tmpMaLinhKien.trim();
        let tmpSoLuong = ("số lượng" in boomList[i]) ? boomList[i]["số lượng"] : "0";
        tmpSoLuong = tmpSoLuong.trim();
        tmpSoLuong = parseFloat(tmpSoLuong);
        // Check data is duplicate
        if( retData.has(tmpMaThanhPham) && retData.get(tmpMaThanhPham).has(tmpMaLinhKien) && (nAlert > 0)) {
            alert("Mã linh kiện: "+tmpMaLinhKien+" bị lặp lại trong: " + tmpMaThanhPham);
            nAlert--;
        }
        let tmpObj = new Map();
        if( !retData.has(tmpMaThanhPham)) {
            // Key is not exists, create new                    
            // tmpObj = new Map();
        } else {
            // Key is exists, update
            tmpObj = retData.get(tmpMaThanhPham);
        }                
        tmpObj.set(tmpMaLinhKien, tmpSoLuong);
        retData.set(tmpMaThanhPham, tmpObj);        
    }
    // console.log(retData);
    return(retData);
}

function processMaterialData(productReportData, boomList) {
    let retData = new Map();

    if(!(productReportData instanceof Map) || !(boomList instanceof Map)) {
        alert("Dữ liệu không hợp lệ");
        return(retData);
    }                        
    
    for(const [index, product] of productReportData ) {
        // Debug::
        // console.log("process: ", product["mathanhpham"]);
        let tmpMaThanhPham = product["mathanhpham"];
        let tmpNgayXuat = product["ngayxuat"];
        let tmpSoLuong = product["soluong"];

        if( boomList.has(tmpMaThanhPham) && (boomList.get(tmpMaThanhPham) instanceof Map)) {
            let  tmpBoomList = boomList.get(tmpMaThanhPham);
            if (tmpBoomList instanceof Map) {
                for(const [materialCode, materialQuantity] of tmpBoomList) {
                    // Debug::
                    // console.log("==== Process: ", materialCode);
                    let tmpKey = "" + tmpNgayXuat + "@" + tmpMaThanhPham + "@"+materialCode;                            
                    // console.log(material);
                    let tmpObject = new Object();
                    if(!retData.has(tmpKey)) {
                        // If key is not exists, create new                        
                        tmpObject = {
                            "ngayxuat": tmpNgayXuat, 
                            "mathanhpham": tmpMaThanhPham, 
                            "malinhkien": materialCode, 
                            "soluong": 0, 
                        }
                    } else {
                        // If key is exists, update
                        tmpObject = retData.get(tmpKey);
                    }
                    tmpObject["soluong"] = tmpObject["soluong"] + materialQuantity*tmpSoLuong
                    retData.set(tmpKey, tmpObject);
                }
            } else {
                alert("Invalid boomlist data");
            }
        }
    }
    return(retData);
}

function setMessage(message) {
    let obj = document.getElementById("divMessage")
    if(obj != null) {     
        obj.innerHTML = message;
    }
}

function setButtonDisabled(disabled) {
    if (disabled !== false) {
        disabled = true;
    }
    let obj = document.getElementById("btnProcess")
    if(obj != null) {     
        obj.disabled = disabled;
        if(disabled == true) {
            obj.style.background = "#FF0000";
        } else {
            obj.style.background = "";
        }
    }
}