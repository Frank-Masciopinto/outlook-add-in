/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */
let API_get_next_Message = "https://api.bextra.io/email/get_next_message/"
let API_url_check_if_new_email = "https://api.bextra.io/email/are_there_any_messages_to_send/"
let API_send_Id_URL = "https://api.bextra.io/email/update_email_send_ce/"

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
  send_email_loop()
  // Write message property value to the task pane
  document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + item.subject;
}

async function send_email_loop() {
  console.log("*** send_email_loop Started ***")
  while (true) {
    let next_email_to_send = await API_Get_Next_Message()
    if (next_email_to_send.match(/Error API Response is Empty||Error API Call/)) {
      console.log(next_email_to_send)
      console.log("Email Sending Sequence Stopped")
      break
    }
    else {
      let email_was_sent = await sendEmail(next_email_to_send)
      if (email_was_sent == true) {
        let email_sent_ID = await get_last_sent_email_ID()
        await API_Call_Send_Message_ID(email_sent_ID)
        break
      }
      else {
        console.log("Email Not Sent")
        break
      }
    }
  }
}

async function API_Call_Send_Message_ID(message_ID) {
  return new Promise((res, rej) => {
    console.log("calling API_Call_Send_Message_ID")
    var xhr = new XMLHttpRequest();
    let api_URL = API_send_Id_URL + "27/7" // + await LS.getItem("campaign_id") + "/" + await LS.getItem("user_id");
    xhr.open("POST", api_URL);
  
    xhr.setRequestHeader("Accept", "application/json");
    xhr.setRequestHeader("Content-Type", "application/json");
  
    xhr.onreadystatechange = function () {
    if (xhr.readyState === 4) {
        console.log(xhr.status);
        //If error response
        if (xhr.status.toString().substring(0,1) != "2") {
          alert(`API CALL ERROR - Response: \n\n ${xhr.response}`)
          res("ERROR")
        }
        //If got a valid response
        else  {
            if (xhr.response == "") {
                console.log("API response is empty, No New Email Campaign to Send")
                res("API response is empty")
            }
            else {
                console.log(xhr.response)
                console.log("message_ID sent")
            }
        }
    }};
    let api_message = {
        "email_message_id": message_ID
    }
    console.log(JSON.stringify(api_message))
    xhr.send(JSON.stringify(api_message));
  })
}

async function sendEmail(message_OBJ) {
return new Promise((res, rej) => {

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
        if (asyncResult.status != "succeeded") {
          console.log("Action failed with error: " + asyncResult.error.message);
          res(false)
        }
        else {
          console.log(`Message sent to ${message_OBJ[0]["to_email_address"]}`);
          res(true)
        }
      });
})
}

async function get_last_sent_email_ID() {
  return new Promise((res, rej) => {
    var request = `<?xml version="1.0" encoding="UTF-8"?>
    <soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:mes="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:typ="http://schemas.microsoft.com/exchange/services/2006/types">
    <soap:Header>
      <typ:RequestServerVersion Version="Exchange2010_SP2" />
    </soap:Header>
    <soap:Body>
      <mes:FindItem Traversal="Shallow">
         <mes:ItemShape>
            <typ:BaseShape>Default</typ:BaseShape>
         </mes:ItemShape>
         <mes:ParentFolderIds>
            <typ:DistinguishedFolderId Id="sentitems" />
         </mes:ParentFolderIds>
      </mes:FindItem>
     </soap:Body>
    </soap:Envelope>`
  
    Office.context.mailbox.makeEwsRequestAsync(request, function (asyncResult) {
      console.log(asyncResult)
      if (asyncResult.status != "succeeded") {
        console.log("Action failed with error: " + asyncResult.error.message);
        res("Failed")
      }
      else {
        console.log(`Request Successfully Sent`);
        let parser = new DOMParser();
        let xmlDoc = parser.parseFromString(asyncResult.value, "text/xml");
        let id = xmlDoc.getElementsByTagName("t:Items")[0].childNodes[0].getElementsByTagName("t:ItemId")[0].getAttribute("Id")
        console.log(`ID Retrieved Below`);
        console.log(id)
        res(id)
      }
    });
  })
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

                  console.log(obj)
                  res(obj)
              }

          }
      }
  };
  xhr.send();    
})
}
