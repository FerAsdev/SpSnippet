/** 
 * @file This library helps you to implement the basic interactions in sharepoint with web services such as
 *       mono and multi site CRUD, upload and delete files to/from a library and access to basic logged user info. 
 *       Also has some basic JS utilities like basic notification function and form error notifications.
 * @summary Library with basic functions to interact with sharepoint lists.
 * @author Felipe Pulido <fpulido.mendoza@gmail.com>
 * @author Fernando Aguilar <fernando.asdev@gmail.com>
 * @requires JQuery1.12.4+
 * @requires SPServices2014.02
 *
 * @copyright F2 2016 - 2017
 * @version 3.2
 * @license
 * Copyright (c) 2017 F2 Inc.
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE. */

/** 
 * A string with the url of your site (used to access the data of logged user)
 * @constant {string} */
const URL_SITE = "http://yoursiteurl.us/url/of/your/site"; // Make sure it doesn't have any trail slash

/**
 * @callback listCallback
 * @param {Object[]} listItems
 *
 * @callback successCallback
 * @callback errorCallback
 */

/* CRUD  */
/**
 * Gets the item(s) requested to a list in Sharepoint, filtered by the query.
 * @example <caption>First usage of getListItems() function (basic).</caption>
 * // returns the ID, FirstName and LastName of every Employee in the Employees list
 * getListItems("Employees", ["ID", "FirstName", "LastName"]);
 *
 * @example <caption>Second usage of getListItems() function (with filtering query).</caption>
 * // returns the ID, FirstName and LastName of the employee with ID equal to 2
 * var query = "<Query>\
                    <Where>\
                        <Eq>\
                            <FieldRef Name='ID'></FieldRef>
                            <Value Type='Number'>2</Value>
                        </Eq>\
                    </Where>\
 *              </Query>"
 * getListItems("Employees", ["ID", "FirstName", "LastName"], query)
 *
 * @example <caption>Third usage of getListItems() function (with callback).</caption>
 * // returns the ID, FirstName and LastName of every Employee in the Employees list and populates an unordered list in the DOM.
 *  function showEmployees(employees) {
        $("#body").append("<ul>");
        $.each(employees, function(i, employee) {
            $("#body").append("<li>"+employee.FirstName+" "+employee.LastName+"</li>");
        });
        $("#body").append("</ul>");
    }
 * getListItems("Employees", ["ID", "FirstName", "LastName"], "", showEmployees);
 *
 * @param {string} lName - The name of the SP list where the items will be searched.
 * @param {string[]} fields - The fields that will be taken from each item.
 * @param {string} [query=<Query></Query>] - CAML query to search only needed items.
 * @param {listCallback} [callback] - Executes a callback function if the service's response was successfull.
 * @returns {Boolean|Object|Array} A boolean if no items were found, an Object if there was only one item or an Array of Objects if more items were found.
 */
function getListItems(lName, fields, query, callback){
    query = query ? query : "<Query></Query>";
    isAsync = callback ? true : false;
    var i = 0;
    var listItems = [];
    var viewFields = "";
    fields.forEach(function(item,index) {
        viewFields += "<FieldRef Name='"+item+"'/>";
    });
    viewFields = "<ViewFields>"+viewFields+"</ViewFields>";

    $().SPServices({
        operation: "GetListItems",
        async: isAsync,
        listName: lName,
        CAMLViewFields: viewFields,
        CAMLQuery: query,                                                                             
        completefunc: function (xData, Status){
            if(Status == "success") {
                $(xData.responseXML).SPFilterNode("z:row").each(function(){
                    listItems[i] = {};
                    var row = $(this);
                    fields.forEach(function(item, index){
                        listItems[i][item] = row.attr('ows_'+item);
                    });
                    i++;
                });
                if(isAsync)
                    callback(listItems);
            }
        }
    });
    if(listItems.length > 1)
        return listItems;
    else {
        if(listItems.length == 1)
            return listItems[0];
        else
            return false;
    }
}

/*
 * An array that defines the data that will be sent with the file.
 * @typedef {*} valuesArray
 */

/**
 * Creates a new item in the given list.
 * @example <caption>How to use createNewListItem() function.</caption>
 * // returns true if the process was successfully done, or false otherwise.
 * createNewListItem("Employees", ["FirstName", "LastName"], ["John", "Doe"]);
 * 
 * @param {string} lName - The name of the SP list where the item will be created.
 * @param {string[]} fields - An array with the names of the columns that will receive a value.
 * @param {valuesArray[]} values - An array with the values that will be set in each column (the length of fields[] and values[] must be the same).
 * @returns {Boolean} A boolean indicating if the creation could or couldn't be done as expected.
 */
function createNewListItem(lName, fields, values) {
    var fieldValues = [];
    var response = false;
    fields.forEach(function(item, index) {
        var pairs = [fields[index], values[index]];
        fieldValues.push(pairs);
    });
    
    $().SPServices({
        operation: 'UpdateListItems',
        async: false,
        batchCmd: 'New',
        listName: lName,
        valuepairs: fieldValues,
        completefunc: function(xData, Status) {
            var newId = $(xData.responseXML).SPFilterNode("z:row").attr("ows_ID");
            if (Status == "success" && newId >= 0)  {
                response = true;
            } else  {   
                response = false;
            }
        }
    });
    return response;
}

/**
 * Updates an item in the given list.
 * @example <caption>Usage of updateListItem() function.</caption>
 * // returns true if the process was successfully done, or false otherwise. You must declare at least two fields / values: the ID and the field of the item which you want to update.
 * updateListItem("Employees", ["ID", "FirstName", "LastName"], [152, "John", "Doe"]); // Assuming 152 is the item's ID
 *
 * @param {string} lName - The SP list name where the selected item should exist to be modified.
 * @param {string[]} fields - A string array with the internal names of the columns in which item's values will be modified.
 * @param {valuesArray[]} values - An array with the values that will overwrite the existing values in the declared fields.
 * @returns {Boolean} A boolean indicating if the transaction has been done as expected.
 */
function updateListItem(lName, fields, values) {
    var fieldValues = "";
    fields.forEach(function(item, index) {
        fieldValues += "<Field Name='"+fields[index]+"'>"+values[index]+"</Field>";
    });
    var batch = "<Batch OnError='Continue' PreCalc='TRUE'><Method ID='1' Cmd='Update'>"+fieldValues+"</Method></Batch>";
    var response;
    $().SPServices({
        operation: "UpdateListItems",
        async: false,
        listName: lName,
        updates: batch,
        completefunc: function(xData, Status) {
            var newId = $(xData.responseXML).SPFilterNode("z:row").attr("ows_ID");
            if (Status == "success" && newId >= 0) {
                response = true;
            } else  {
                response = false;
            }
        }
    });
    return response;
}

/**
 * Deletes an item in the selected list.
 * @example <caption>Usage of deleteListItem() function.</caption>
 * // returns true if the item was successfully deleted, or false otherwise.
 * deleteListItem("Employees", 152);
 *
 * @param {string} lName - The SP list that has the item that will be deleted.
 * @param {number} id - The id of the item that sharepoint gave it when it was created.
 * @returns {Boolean} false if the process wasn't successfully done, otherwise, true.
 */
function deleteListItem(lName, id) {
    var fieldValues = "<Field Name='ID'>"+id+"</Field>";
    var batch = "<Batch OnError='Continue' PreCalc='TRUE'><Method ID='1' Cmd='Delete'>"+fieldValues+"</Method></Batch>";
    var response;
    $().SPServices({
        operation: "UpdateListItems",
        async: false,
        listName: lName,
        updates: batch,
        completefunc: function(xData, Status) {
            var newId = $(xData.responseXML).SPFilterNode("z:row").attr("ows_ID");
            if (Status == "success" && newId >= 0) {
                response = true;
            } else  {
                response = false;
            }
        }
    });
    return response;
}

/**
 * An array that defines the data that will be sent with the file.
 * @typedef {array} fieldsArray
 * @property {string} 0 - The value type.
 * @property {string} 1 - Display name of the column in the Sharepoint Library.
 * @property {string} 2 - Internal name of the column in the Sharepoint Library.
 * @property {*} 3 - The value that the item's declared column will take.
 */

/**
 * Allows the user to upload a file to a Sharepoint library.
 * @example <caption>Usage of uploadFile() function.</caption>
 * // Calls the corresponding callback function depending on the status of the transaction.
 * var urlSite = "http://example.us/sites/mainsite/subsite/site";
 * var lName = "EmployeesDocuments";
 * var idInputFile = "fileInput"; // You should have a <input type='file' id='fileInput'/> in the DOM
 * var fields = [["Text", "Description", "Description", "Description of the file that will be uploaded."],
                 ["Number", "EmployeeID", "EmployeeID", 152]];
        // Field Type | Display Name | Internal Name | Value
        // You can declare as many fields as columns in the library.
 * var filename = "JohnDoe.txt"; // I'm assuming you have file defined somewhere in your code.
 * function onSuccess() {
        alert("Wohooo file uploaded");
   }
 * function onError() {
        alert("Oops, something went wrong :("); 
   }
 * uploadFile(urlSite, lName, idInputFile, fields, file, onSuccess, onError);
 *
 * @param {string} urlSite - The URL where the library belongs to.
 * @param {string} lName - The library's name.
 * @param {string} idInputFile - The id of the file input that is being used to upload the file.
 * @param {fieldsArray[]} fields.
 * @param {string} filename - The file's name.
 * @param {file} file - The file to upload.
 * @param {successCallback} [successCallback] - A callback function to call if the file was uploaded successfully.
 * @param {errorCallback} [errorCallback] - A callback function to call if the file couldn't be uploaded.
 */
function uploadFile(urlSite, lName, idInputFile, fields, filename, file, successCallback, errorCallback) {
    var path = $("#"+idInputFile).val();
    var fieldInformation = "";
    $.each(fields, function(i, values) {
        var type = values[0];
        var displayName = values[1];
        var internalName = values[2];
        var value = values[3];
        fieldInformation += "<FieldInformation Type='"+type+"' DisplayName='"+displayName+"' InternalName='"+internalName+"' Value='"+value+"'/>";
    });
    filereader = new FileReader();
    filereader.filename = filename;
    filereader.onload = function() {
        data = filereader.result;
        n = data.indexOf(';base64,')+8;
        data = data.substring(n);
        var soapEnv =
        "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'> \
            <soap:Body>\
                <CopyIntoItems xmlns='http://schemas.microsoft.com/sharepoint/soap/'>\
                    <SourceUrl>" + path + "</SourceUrl>\
                        <DestinationUrls>\
                            <string>"+urlSite+"/"+lName+"/" + filename + "</string>\
                        </DestinationUrls>\
                        <Fields>\
                            "+fieldInformation+"\
                        </Fields>\
                    <Stream>" + data + "</Stream>\
                </CopyIntoItems>\
            </soap:Body>\
        </soap:Envelope>";

        $.ajax({
            url: urlSite + "/_vti_bin/Copy.asmx",
            beforeSend: function (xhr) { 
                xhr.setRequestHeader("SOAPAction", "http://schemas.microsoft.com/sharepoint/soap/CopyIntoItems"); },
            type: "POST",
            dataType: "xml",
            data: soapEnv,
            success: function() {
                if(successCallback)
                    successCallback();
                else
                    console.log("File uploaded successfully");
            },
            error: function(response) {
                if(errorCallback)
                    errorCallback(response);
                else
                    console.log("The file couldn't be uploaded");
            },
            contentType: "text/xml; charset=\"utf-8\""
        });
    };
    filereader.readAsDataURL(file);
}

/**
 * Delete a file from a library
 * @example <caption>How to use deleteFile() function.</caption>
 * // returns true if the transaction was done succesfully or false if it wasn't.
 * var lName = "EmployeesDocuments";
 * var docPath = "http://example.us/sites/mainsite/subsite/site/EmployeesDocuments/JohnDoe.txt";
 * deleteFile(lName, docPath, 12); // Assuming 12 is the id of the document in the library
 *
 * @param {string} lName - The name of the library that the file belongs to.
 * @param {string} filePath - The URL of the file that will be deleted.
 * @param {number} id - The item's id.
 */
function deleteFile(lName, filePath, id) {
    var batchCmd = "<Batch OnError='Continue'>\
                        <Method ID='1' Cmd='Delete'>\
                            <Field Name='ID'>" + id + "</Field>\
                            <Field Name='FileRef'>" + filePath + "</Field>\
                        </Method>\
                    </Batch>";
    var response = false;

    $().SPServices({
        operation: "UpdateListItems",
        async: false,
        listName: lName,
        updates: batchCmd,
        completefunc: function ( xData, Status ) {
            $( xData.responseXML ).SPFilterNode( 'ErrorCode' ).each( function(){
                responseError = $( this ).text();
                if ( responseError === '0x00000000' ) {
                    response = true;
                } else {
                    response = false;
                }
            });
        }
    });
    return response;
}

/* CROSS-SITE SP FUNCTIONS */
/**
 * @todo Document this function
 */
function getExternalListItems(lName, query, fields, webUrl, callback){
    isAsync = callback ? true : false;
    var i = 0;
    var listItems = [];
    var viewFields = "";
    fields.forEach(function(item,index) {
        viewFields += "<FieldRef Name='"+item+"'/>";
    });
    viewFields = "<ViewFields>"+viewFields+"</ViewFields>";

    $().SPServices({
        operation: "GetListItems",
        async: isAsync,
        webURL: webUrl,
        listName: lName,
        CAMLViewFields: viewFields,
        CAMLQuery: query,                                                                             
        completefunc: function (xData, Status){
            if(Status == "success") {
                $(xData.responseXML).SPFilterNode("z:row").each(function(){
                    listItems[i] = {};
                    var row = $(this);
                    fields.forEach(function(item, index){
                        listItems[i][item] = row.attr('ows_'+item);
                    });
                    i++;
                });
                if(isAsync)
                    callback(listItems);
            }
        }
    });
    return listItems;
}

/**
 * @todo Document this function
 */
function createExternalListItem(lName, fields, values, webUrl) {
    var fieldValues = [];
    var response = false;
    fields.forEach(function(item, index) {
        var pairs = [fields[index], values[index]];
        fieldValues.push(pairs);
    });

    $().SPServices({
        operation: 'UpdateListItems',
        async: false,
        batchCmd: 'New',
        listName: lName,
        webURL: webUrl,
        valuepairs: fieldValues,
        completefunc: function(xData, Status) {
            var newId = $(xData.responseXML).SPFilterNode("z:row").attr("ows_ID");
            if (Status == "success" && newId >= 0)  {
                response = true;
            } else  {   
                response = false;
            }
        }
    });
    return response;
}

/**
 * @todo Document this function
 */
function updateExternalListItem(lName, fields, values, webUrl) {
    var fieldValues = "";
    fields.forEach(function(item, index) {
        fieldValues += "<Field Name='"+fields[index]+"'>"+values[index]+"</Field>";
    });
    var batch = "<Batch OnError='Continue' PreCalc='TRUE'><Method ID='1' Cmd='Update'>"+fieldValues+"</Method></Batch>";
    var response;
    $().SPServices({
        operation: "UpdateListItems",
        async: false,
        webURL: webUrl,
        listName: lName,
        updates: batch,
        completefunc: function(xData, Status) {
            var newId = $(xData.responseXML).SPFilterNode("z:row").attr("ows_ID");
            if (Status == "success" && newId >= 0) {
                response = true;
            } else  {
                response = false;
            }
        }
    });
    return response;
}

/* UTILITIES  */
/**
 * @todo Document this function
 */
function getCurrentUser() {
    var idUser, name;
    var user = [];
    var soapEnv = "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'> \
                    <soap:Body> \
                        <GetCurrentUserInfo xmlns='http://schemas.microsoft.com/sharepoint/soap/directory/' /> \
                    </soap:Body> \
                </soap:Envelope>";
    $.ajax({
        url: URL_SITE+"/_vti_bin/usergroup.asmx",
        async:false,
        type: "POST",
        dataType: "xml",
        data: soapEnv,
        complete: function(xData, status){
            idUser = $(xData.responseXML).find("User").attr("ID");  
            name = $(xData.responseXML).find("User").attr("Name");  
        },
        contentType: "text/xml; charset=\"utf-8\""
    });
    user = [idUser, name];
    return user;
}

/**
 * @todo Document this function
 */
function getUrlVars() {
    var vars = {};
    var parts = window.location.href.replace(/[?&]+([^=&]+)=([^&]*)/gi, function(m,key,value) {
        vars[key] = value;
    });
    return vars;
}

/**
 * @todo Document this function
 */
function responseToArray(objArr, property, person) {
    var arrayObtained = [];
    $.each(objArr, function(i, obj) {
        if(arrayObtained.indexOf(obj[property]) == -1)
            arrayObtained.push(obj[property]);
    });
    if(person) {
        arrayObtained.sort(function(a, b) {
            if(a.split("#")[1] > b.split("#")[1]) return 1;
            if(a.split("#")[1] < b.split("#")[1]) return -1;
            return 0;
        });
    }
    return arrayObtained;
}

/* NOTIFICATIONS  */
/**
 * @todo Document this function
 */
var msgColors = {
    success: "#51b956",
    error: "#CA2A2A",
    warning: "#E68F0C"
};

/**
 * @todo Document this function
 */
function setColor(color, value) {
    msgColors[color] = value;
}

/**
 * @todo Document this function
 */
function showMessage(msg, color) {
    if($("#update-status").length == 0) {
        alert("You need to add the .message div tag in order to use notifications.");
        // Add this tag to your html <div id="update-status" class="message"></div>
        /* and this style to your css
            .message {
                display: none;
                width: 30%;
                position: fixed;
                bottom: 4%;
                right: 2%;
                padding: 2%;
                text-align: center;
                border: none;
                border-radius: 10px;
                color: #FFFFFF;
                font-weight: bold;
                background: #898989;
                z-index: 999999;
                -webkit-box-shadow: 0px 0px 35px -8px rgba(0,0,0,0.75);
                -moz-box-shadow: 0px 0px 35px -8px rgba(0,0,0,0.75);
                box-shadow: 0px 0px 35px -8px rgba(0,0,0,0.75);
            }
         */
        return false;
    }
    if(msgColors[color] == undefined) {
        alert("The color "+color+" is not defined");
        return false;
    }

    $("#update-status").text(msg);
    $("#update-status").css("background-color", msgColors[color]);
    $("#update-status").fadeIn(1000,function() {
        setTimeout(function() {
            $("#update-status").fadeOut(1000);
        },3000);
    });
}

/**
 * @todo Document this function
 */
$(function() {
    $("#errorBox").hide();
    $("select, input, textarea").data("error", "");

    $("body").on("mousemove", "select, input, textarea", function(evt) {
        if($(this).data("error") != "")
            $("#errorBox").css({ "top": evt.pageY +20, "left": evt.pageX + 20 });
    });

    $("body").on("mouseenter", "select, input, textarea", function(evt) {
        if($(this).data("error") != "")
            $("#errorBox").addClass("errorBox").text($(this).data("error")).show();
    });    

    $("body").on("mouseleave", "select, input, textarea", function(evt) {
        $("#errorBox").removeClass("errorBox").text("").hide();
    });

    $("body").on("change keyup", "select, input, textarea", function() {
        $(this).css("border","1px solid #66afe9").data("error","");
    });
    /* 
        <div id="errorBox" class="errorBox"></div>
        .errorBox {
            border: 1px solid gray;
            border-radius: 5px;
            border-top-left-radius: 0px;
            background-color: white;
            position: absolute;
            display: block;
            padding: 0.5%;
            z-index: 9999;
        }
     */
});