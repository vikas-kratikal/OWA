/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */
let header = '',
    from = null,
    to = null,
    subject = '',
    bodyHtml = '',
    bodyText = '',
    mailAttachments = [],
    mailRecieveTime = '',
    accessToken = '',
    key = null,
    mailEml64 = '',
    currentItem = null,
    userProfile = null,
    authToken = null,
    mailId = null;

/**
 * Error Display Function
 */

function showErrorMessage(message) {
    $('#analysis').hide();
    $('#error-message').html(message);
    $('#error-message').fadeIn();
}

function showPageLoader() {
    $('.page-loader').removeClass('d-none')
    $('.page-loader').addClass('d-block')
}

function hidePageLoader() {
    $('.page-loader').removeClass('d-block')
    $('.page-loader').addClass('d-none')
}

function Refresh() {
    showPageLoader()
    from = Office.context.mailbox.item.from.emailAddress;
    to = Office.context.mailbox.userProfile.emailAddress;
    subject = Office.context.mailbox.item.subject;

    getDataFromMail();
    // document.getElementById("hackRecord").onclick = showHackRecord;
    // document.getElementById("back").onclick = backButton;
    // document.getElementById("ham").onclick = markHam;
    // $('#hackRecordTable').hide();
    // $('#hacksafe').hide();
    // $('#analysis').show();
    // $('#back').hide();
}

(function loadOffice() {
    $("#loadingText").html("Initializing....");
    console.log("Initializing office....");
    // First check if the script already exists on the dom
    // by searching for an id
    let id = 'office'
    if (document.getElementById(id) === null) {
        let script = document.createElement('script')
        script.setAttribute('src', 'https://appsforoffice.microsoft.com/lib/1/hosted/office.js')
        script.setAttribute('id', id)
        document.body.appendChild(script)

        // now wait for it to load...
        script.onload = () => {
            // script has loaded, you can now use it safely
            console.log("Office loaded successfully")
            initializeOffice()
        }

        // now wait for it to load...
        script.onerror = () => {
            // script has loaded, you can now use it safely
            console.log("Office not loaded")
            showErrorMessage("Office not loaded")
        }
    } else {
        console.log("Inside else of office load....");
    }
})()

function initializeOffice() {
    Office.initialize = function () { };
    $("#loadingText").html("Office Initialized....");

    Office.onReady(function (info) {
        showPageLoader();
        from = Office.context.mailbox.item.from.emailAddress;
        to = Office.context.mailbox.userProfile.emailAddress;
        subject = Office.context.mailbox.item.subject;

        console.log('insde initialize office', { from, to });
        getDataFromMail();
        // document.getElementById("hackRecord").onclick = showHackRecord;
        // document.getElementById("refresh").onclick = Refresh;
        // document.getElementById("back").onclick = backButton;
        // document.getElementById("ham").onclick = markHam;
        // $('#analysis').hide();
        // $('#back').hide();
        hidePageLoader()
    });
}

function getDataFromMail() {
    Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
        showPageLoader();
        $("#loadingText").html("Data Loading..... ");
        const ewsId = Office.context.mailbox.item.itemId;
        authToken = result.value;
        mailId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
        console.log('mail id => ', mailId);
        hidePageLoader()
    });

    let attachments = Office.context.mailbox.item.attachments;
    console.log('attachment => ', attachments);
    attachments.forEach((el)=>{
        $('#scanMailAttachment').append(`<li><p>${el.name}<p> <button id=${el.id} class="btn btn-primary">Scan</button></li>`)
        document.getElementById(`${el.id}`).onclick = scanAttachment;
    })

    getMailHeaders(function (rawHeader) {
        header = rawHeader;
        $('#run').show();
        document.getElementById("run").onclick = run;
    });
}


function getMailHeaders(callback) {
    try {
        Office.context.mailbox.item.getAllInternetHeadersAsync(
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                    callback(asyncResult.value);
                } else {
                    $("#loadingText").html("Headers fething failed....!");
                    callback("");
                }
            });
    } catch (err) {
        console.log("Error while fetching headers", err);
        $("#loadingText").html("Headers fething failed....!");
        callback("");
    }
}

function reportSpam() {
    console.log('reportSpam called from taskpane.js');
    let data = { from, to, subject, mailId, header, authToken }
    mailId='';
    console.log('reportSpam => ', data);
    if (from && to && subject && mailId && header && authToken) {
        console.log('inside if');
        let url = "http://localhost:8080/api/outlook/reportSpam";
        $.ajax({
            headers: {
                'Accept': 'application/json',
                'Content-Type': 'application/json'
            },
            url: url,
            type: "POST",
            data: JSON.stringify(data),
        })
            .done(function (item) {
                console.log('itemmmm', item);
                var item = Office.context.mailbox.item;
                // moveToJunk(item.itemId);
                Refresh()
                hidePageLoader()
            })
            .fail(function (jqXHR, textStatus) {
                hidePageLoader()
                document.getElementById("warning").show()
                document.getElementById("app-body").hide();
                console.log('error sending', textStatus, jqXHR);
            })
    }else{
        showErrorMessage("Auth token or mailId fething failed....!")
        hidePageLoader()
    }
}

function attachmentScan(attachID) {
    console.log('attachmentScan called from taskpane.js');
    let data = { mailId, authToken }
    mailId='';
    console.log('reportSpam => ',attachID);
    if ( attachID && mailId && authToken) {
        console.log('inside if');
        let url = "http://localhost:8080/XXXXXXXXX";
        $.ajax({
            headers: {
                'Accept': 'application/json',
                'Content-Type': 'application/json'
            },
            url: url,
            type: "POST",
            data: JSON.stringify(data),
        })
            .done(function (item) {
                console.log('itemmmm', item);
                var item = Office.context.mailbox.item;
                // moveToJunk(item.itemId);
                Refresh()
                hidePageLoader()
            })
            .fail(function (jqXHR, textStatus) {
                hidePageLoader()
                document.getElementById("warning").show()
                document.getElementById("app-body").hide();
                console.log('error sending', textStatus, jqXHR);
            })
    }else{
        showErrorMessage("Attachment fetching....!")
        hidePageLoader()
    }
    hidePageLoader()
}

export async function run() {
    showPageLoader()
    Office.onReady()
        .then(function () {
            reportSpam();
        })
        .catch(function (err) {
            console.log(err);
        })
}

function scanAttachment(e) {
    showPageLoader()
    Office.onReady()
        .then(function () {
            attachmentScan(e.target.id);
        })
        .catch(function (err) {
            console.log(err);
        })
}