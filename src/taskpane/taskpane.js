/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */
let API_get_next_Message = "https://api.bextra.io/email/get_next_message/"
let API_url_check_if_new_email = "https://api.bextra.io/email/are_there_any_messages_to_send/"

Office.initialize = () => {
  console.log("Initialized")
};

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  // Get a reference to the current message
  var item = Office.context.mailbox.item;
  console.log(item)
  API_Get_Next_Message()
  // Write message property value to the task pane
  document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + item.subject;
}

function sendEmail(message_OBJ) {

  let bcc = message_OBJ[0]["bcc"] != null ? 
  `          <t:BccRecipients>` +
  `           <t:Mailbox><t:EmailAddress>` + message_OBJ[0]["bcc"] + `</t:EmailAddress></t:Mailbox>` +
  `          </t:BccRecipients>`
  : ""
  let cc = message_OBJ[0]["cc"] != null ? 
  `          <t:CcRecipients>` +
  `           <t:Mailbox><t:EmailAddress>` + message_OBJ[0]["cc"] + `</t:EmailAddress></t:Mailbox>` +
  `          </t:CcRecipients>`
  : ""
  var request = `<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:xs="http://www.w3.org/2001/XMLSchema" targetNamespace="http://schemas.microsoft.com/exchange/services/2006/messages" elementFormDefault="qualified" version="Exchange2016" id="messages">` +
        `  <soap:Header><t:RequestServerVersion Version="Exchange2010" /></soap:Header>` +
        `  <soap:Body>` +
        `    <m:CreateItem MessageDisposition="SendAndSaveCopy">` +
        `      <m:SavedItemFolderId><t:DistinguishedFolderId Id="sentitems" /></m:SavedItemFolderId>` +
        `      <m:Items>` +
        `        <t:Message>` +
        `          <t:Subject>${message_OBJ[0]["subject"]}</t:Subject>` +
        `          <t:Body BodyType="HTML"><![CDATA[${message_OBJ[0]["body"]}]]></t:Body>` +
        `          <t:ToRecipients>` +
        `            <t:Mailbox><t:EmailAddress>ermascio@gmail.com</t:EmailAddress></t:Mailbox>` +
        `          </t:ToRecipients>` +
        bcc + cc +
        `        </t:Message>` +
        `      </m:Items>` +
        `    </m:CreateItem>` +
        `  </soap:Body>` +
        `</soap:Envelope>`;
        console.log(request)

      Office.context.mailbox.makeEwsRequestAsync(request, function (asyncResult) {
        console.log(asyncResult.value)
        if (asyncResult.status == "failed") {
          console.log("Action failed with error: " + asyncResult.error.message);
        }
        else {
          console.log(`Message sent to ${message_OBJ[0]["to_email_address"]}`);
        }
      });
}

function get_last_sent_email_ID() {
  var request = '<?xml version="1.0" encoding="utf-8"?>'+
'  <soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"'+
'                 xmlns:t="https://schemas.microsoft.com/exchange/services/2006/types">'+
'    <soap:Body>'+
'      <FindItem xmlns="https://schemas.microsoft.com/exchange/services/2006/messages"'+
                 'xmlns:t="https://schemas.microsoft.com/exchange/services/2006/types"'+
                'Traversal="Shallow">'+
'        <ItemShape>'+
'          <t:BaseShape>IdOnly</t:BaseShape>'+
        '</ItemShape>'+
        '<ParentFolderIds>'+
          '<t:DistinguishedFolderId Id="deleteditems"/>'+
        '</ParentFolderIds>'+
      '</FindItem>'+
    '</soap:Body>'+
  '</soap:Envelope>'
}

async function loop_Get_Next_Message_Until_None(campaign_ids_Array) {
  async function wait_for_ending_read_emails() {
      return new Promise ((res, rej) => {
          let check_if_ended = setInterval(() => {
              if (localStorage.getItem("Reading_Email_Campaign_is_Active?") == "NO") {
                  console.log("Finished Reading Emails")
                  clearInterval(check_if_ended)
                  res()
              }
          }, 3000);
      })
  }
  localStorage.setItem("Is_Email_Campaign_Active?", "YES")

  console.log("Gmail Automation Started! With campaign ID:" + campaign_ids_Array[0])

  let number_of_Emails_Sent = 0
  for (let i = 0; i < campaign_ids_Array.length;) {
        let email_Sending_result = await API_Get_Next_Message(campaign_ids_Array[i])
        if (email_Sending_result == "SENT") {
            number_of_Emails_Sent++
            console.log("1 Email Sent, continuing with next one within same campaign id")
            //Every 100 emails sent, start reading sequence
            if (number_of_Emails_Sent % 100 == 0) {
                localStorage.setItem("Reading_Email_Campaign_is_Active?", "YES")
                console.log("Reading email message sending next")
                await wait_for_ending_read_emails()
            }
        }
        else {
            console.log(email_Sending_result)
            console.log("changing campaign id")
            i++
            if (i == campaign_ids_Array.length) {
                localStorage.setItem("Is_Email_Campaign_Active?", "NO")
                console.log("All Emails Sent!")
                chrome.notifications.create({
                    type: 'basic',
                    iconUrl: 'Images/128.png',
                    title: `Finished. Total Emails Sent: ${number_of_Emails_Sent}`,
                    message: 'Gmail Automation Finished!',
                    priority: 1
                })
                localStorage.setItem("Reading_Email_Campaign_is_Active?", "YES")
                console.log("Reading email message sending next")
                chrome.tabs.sendMessage(google_tab_Id, {
                    message: "start_listen_Automation",
                    last_message_id: localStorage.getItem("last_MessageID_read"),
                    date_limit: localStorage.getItem("reading_days_limit")
                    })
                
                await wait_for_ending_read_emails()
                }
            }
  }
}

async function API_Get_Next_Message() {
  return new Promise((res, rej) => {
      
      console.log("calling API_Get_Next_Message")
      var xhr = new XMLHttpRequest();
      let api_URL = API_get_next_Message + "27/7";
      xhr.open("GET", api_URL);
  
      xhr.setRequestHeader("Accept", "application/json");
      xhr.setRequestHeader("Content-Type", "application/json");
      console.log(api_URL)
      xhr.onreadystatechange = function () {
      if (xhr.readyState === 4) {
          console.log(xhr);
          console.log(xhr.response)
          //If error response
          if (xhr.status.toString().substring(0,1) != "2") {
            alert(`API CALL ERROR - Response: \n\n ${xhr.response}`)
            res("Error API Call")
          }
          else if (xhr.response == "") {
              res("Error API Response is Empty")
          }
          //If got a valid response
          else if (xhr.response != "") {
              let obj = JSON.parse(xhr.response);
              if (obj != ""){
                  // obj[0]["body"] = obj[0]["body"].replace("<", "&lt;")
                  // obj[0]["body"] = obj[0]["body"].replace(">", "&gt;")
                  // obj[0]["body"] = obj[0]["body"].replace("&", "&amp;")
                  console.log(obj)
                  sendEmail(obj)
              }
              //Waiting for Email sent confirmation from Content script
              // let confirmation_email_sent = setInterval(function() {
              //     if (message_ID != null) {
              //         console.log("Received Message ID")
              //         clearInterval(confirmation_email_sent)
              //         //send message id to API
              //         API_Call_Send_Message_ID(obj[0].message_id, message_ID)
              //         message_ID = null
              //         //Closing loop and check if exist a new message
              //         res("SENT")
              //     }
              // }, 1500)
          }
      }
  };
  xhr.send();    
})
}

async function API_Check_If_New_Messages_To_Send(one_or_all_Emails_in_campaign) {
  console.log("calling API_Check_If_New_Messages_To_Send")
  var xhr = new XMLHttpRequest();
  let api_URL = API_url_check_if_new_email + "ermascio@live.it"; // CHANGE IT BACK localStorage.getItem("selected_email")
  console.log(api_URL)
  xhr.open("GET", api_URL);

  xhr.setRequestHeader("Accept", "application/json");
  xhr.setRequestHeader("Content-Type", "application/json");

  xhr.onreadystatechange = function () {
  if (xhr.readyState === 4) {
      console.log(xhr.status);
      //If error response
      if (xhr.status.toString().substring(0,1) != "2") {
        alert(`API CALL ERROR - Response: \n\n ${xhr.response}`)
      }
      //If got a valid response
      else  {
          if (xhr.response == "") {
              console.log("API response is empty, No New Email Campaign to Send")
          }
          else {
              console.log(xhr)
              let obj = JSON.parse(xhr.response);
              console.log(obj)
              console.log(obj[0].campaign_id)
              let campaign_ids = []
              let promis = new Promise((res, rej) => {
                  let i = 0;
                  for (i; i < obj.length; i++) {
                      console.log("inside loop")
                      campaign_ids.push(obj[i].campaign_id)
                  }
                  res()
              }).then((res) => {
                  localStorage.setItem("Campaign_ids",campaign_ids)
                  console.log(campaign_ids)
                  if (one_or_all_Emails_in_campaign == "ALL_EMAIL") {
                      loop_Get_Next_Message_Until_None(campaign_ids)
                  }
                  else {
                      get_and_send_One_Email(campaign_ids)
                  }
              })
          }
      }
  }
};

 
  xhr.send();
}