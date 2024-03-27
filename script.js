const FONT_COLOR = "#FFFFFF";
const BG_COLOR = "#1B0AE1";

function onOpen() {
    // Create a custom menu in the spreadsheet
    let ui = SpreadsheetApp.getUi();
    ui.createMenu("Better Contact")
        .addItem("Find Emails", "openSidebar")
        .addToUi();
}

function openSidebar() {
    // Create and display the sidebar
    let html = HtmlService.createHtmlOutputFromFile("index")
        .setTitle("Email Finder")
        .setWidth(300);
    SpreadsheetApp.getUi().showSidebar(html);
}


function getSheetRowCount(){
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    return sheet.getLastRow()
}




function postDataToAPI(API_KEY, payload) {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    let txtFirstNameCol = getColumnIndex(payload.txtFirstName);
    let txtLastNameCol = getColumnIndex(payload.txtLastName);
    let txtCompanyNameCol = getColumnIndex(payload.txtCompanyName);
    let txtDomainCol = getColumnIndex(payload.txtCompanyDomain);
    let txtLinkedInCol = getColumnIndex(payload.txtLinkedIn);

    // Get the last row with data in any of the specified columns
    let lastRow = sheet.getLastRow();
    let startingRow = 2;
    let skipRows = parseInt(payload.txtSkipRows)

    if(skipRows >= lastRow){
        return {
            "status":"error",
            "message": "You cannot skip more rows then exist"
        }
    }

    if (payload.cbSheetHasHeaders == "NO") {
        startingRow = 1
    }

    if (skipRows > 0) {
        startingRow = skipRows + 1
    }

    if (payload.cbSheetHasHeaders == "YES" && skipRows < 1) {
        lastRow = lastRow - 1
    }

    // getRange(startRow, startColumn, numRows, numColumns)
    // Get the data in the specified columns for all rows
    let dataRange = sheet.getRange(startingRow, 1, lastRow, sheet.getLastColumn());
    let data = dataRange.getValues();

    // Convert data to a list of dictionaries
    let dataListToPost = [];
    data.map(function (row, index) {
        if (
            row[txtFirstNameCol] != '' ||
            row[txtLastNameCol] != '' ||
            row[txtCompanyNameCol] != ''
        ) {

            let dataObject = {
                "first_name": row[txtFirstNameCol],
                "last_name": row[txtLastNameCol],
                "company": row[txtCompanyNameCol],
                "company_domain": row[txtDomainCol],
                "linkedin_url": row[txtLinkedInCol],
                "custom_fields": {
                    "row_id": index
                }
            };
            dataListToPost.push(dataObject);
        }
    });

    response = makePostCall(API_KEY, { data: dataListToPost });

    let returnData = {
        "payload": payload,
        "api_key": API_KEY,
        "id": response.id
    }

    response["data"] = returnData
    return response

}

function makePostCall(API_KEY, postData) {
    const URL = "https://app.bettercontact.rocks/api/v2/async?api_key=" + API_KEY;
    // Options for the fetch request
    var options = {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(postData),
    };

    try {
        response = UrlFetchApp.fetch(URL, options);
        if (response.getResponseCode() === 200 || response.getResponseCode() === 201) {
            responseData = JSON.parse(response.getContentText());
            responseData.status = "success";
            responseData.message = "Your sheet has been submitted for enrichment.";
            return responseData
        } else {
            return {
                "status": "error",
                "message": response.getContentText()
            };
        }
    }
    catch (exception) {
        return {
            "status": "error",
            "message": exception.message.replace("(use muteHttpExceptions option to examine full response)", "")
        }
    }
}

function getColumnIndex(field) {
    return field.toUpperCase().charCodeAt(0) - 64 - 1;
}

function runGetRequest() {
    var data = {
        "api_key": "2705f1db8959a2431e54",
        "payload": {
            "txtRequestId": "",
            "txtSkipRows": "20",
            "cbSkipRows": "on",
            "cbAgree": "on",
            "cbSheetHasHeaders": "YES",
            "txtCompanyName": "C",
            "txtLastName": "B",
            "txtFirstName": "A",
            "txtLinkedIn": "E",
            "txtCompanyDomain": "D"
        },
        "id": "216d5d6964df472dd06a"
    }
    let response = checkStatusAndGetData(data)
    console.log(response)
}


function createHeaders(payload) {
    let headers = [];
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getActiveSheet();


    if (payload.cbSheetHasHeaders == "NO") {
        let txtFirstNameCol = getColumnIndex(payload.txtFirstName);
        let txtLastNameCol = getColumnIndex(payload.txtLastName);
        let txtCompanyNameCol = getColumnIndex(payload.txtCompanyName);
        let txtDomainCol = getColumnIndex(payload.txtCompanyDomain);
        let txtLinkedInCol = getColumnIndex(payload.txtLinkedIn);

        if (!isNaN(txtFirstNameCol))
            headers.splice(txtFirstNameCol, 0, "First Name")

        if (!isNaN(txtLastNameCol))
            headers.splice(txtLastNameCol, 0, "Last Name")

        if (!isNaN(txtCompanyNameCol))
            headers.splice(txtCompanyNameCol, 0, "Company Name")

        if (!isNaN(txtDomainCol))
            headers.splice(txtDomainCol, 0, "Company Domain")

        if (!isNaN(txtLinkedInCol))
            headers.splice(txtLinkedInCol, 0, "Linkedin URL")
    }


    headers.push("Email")
    headers.push("Email Provider")
    headers.push("Delivery Status")

    let rangeToFill = null
    if (payload.cbSheetHasHeaders == "NO") {
        sheet.insertRowBefore(1)
        Logger.log("A new row has been inserted on top")
        rangeToFill = sheet.getRange(1, 1, 1, headers.length)
    }
    else {
        // getRange(startRow, startColumn, numRows, numColumns)
        rangeToFill = sheet.getRange(1, sheet.getLastColumn() + 1, 1, 3)
    }



    rangeToFill.setValues([headers]);
    Logger.log("Header values are filled")
}

function checkStatusAndGetData(params) {
    let response = getDataFromBetterConnect(params)
    if (response.status == "terminated") {
        writeEmailsToSheet(response.data, params)
        delete response["data"]
    }
    return response
}

function writeEmailsToSheet(data, params) {
    let payload = params.payload;
    let sortedData = data.sort((a, b) => a.custom_fields[0].value - b.custom_fields[0].value);


    let skipRows = parseInt(payload.txtSkipRows)
    let startingRow = 2
    if (skipRows < 1) {
        createHeaders(payload)
    }
    else {
        startingRow = skipRows + 1
    }

    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let currentSheet = ss.getActiveSheet();
    let startingCol = currentSheet.getLastColumn() - 2

    let newDataSet = [];
    sortedData.map((row) => {
        row_id = row.custom_fields[0].value
        let tempData = [];
        tempData.push(row.contact_email_address);
        tempData.push(row.contact_email_address_provider);
        tempData.push(row.contact_email_address_status);
        newDataSet.push(tempData);
    });

    // getRange(startRow, startColumn, numRows, numColumns)
    let rangeToFill = currentSheet.getRange(startingRow, startingCol, newDataSet.length, 3)
    rangeToFill.setValues(newDataSet);
    Logger.log("New Sheet has been created and enriched data is filled there");
}

function getDataFromBetterConnect(params) {
    try {
        const URL = `https://app.bettercontact.rocks/api/v2/async/${params.id}?api_key=${params.api_key}`;
        response = UrlFetchApp.fetch(URL);
        if (response.getResponseCode() === 200 || response.getResponseCode() === 201) {
            responseData = JSON.parse(response.getContentText());
            if (responseData.status == "not started yet") {
                responseData.status = "pending"
                responseData.message = "Enrichment task did not start yet. Please wait"
            }
            else if (responseData.status == "in progress") {
                responseData.status = "processing"
                responseData.message = "Enrichment task is in progress"
            }
            else if (responseData.status == "terminated") {
                responseData.status = "terminated"
                responseData.message = "Enrichment task is terminated. Results are available."
            }
            else {
                responseData.status = "error"
                responseData.message = "Enrichment task is in error."
            }
            return responseData
        } else {
            return {
                "status": "error",
                "message": response.getContentText()
            };
        }
    }
    catch (exception) {

        if (exception.message.includes("Unvalid request_id")) {
            return {
                "status": "pending",
                "message": "Enrichment task did not start yet. Please wait"
            }
        }
        else {

            return {
                "status": "error",
                "message": exception.message.replace("(use muteHttpExceptions option to examine full response)", "")
            }

        }
    }
}
