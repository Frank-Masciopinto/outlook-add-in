/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */
let API_get_next_Message = "https://api.bextra.io/email/get_next_message/"
let API_url_check_if_new_email = "https://api.bextra.io/email/are_there_any_messages_to_send/"
let API_send_Id_URL = "https://api.bextra.io/email/update_email_send_ce/"
let api_ERROR_URL = "https://api.bextra.io/error/report"
let api_ce_login = "https://api.bextra.io/ce/login";

Office.initialize = () => {
  console.log("Initialized")
};

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("greeting").innerHTML = `Hi ${Office.context.mailbox.userProfile.displayName}, Start Sending Email with BPersonal`
    document.getElementById("run").onclick = run;
    let stored_regkey = Office.context.roamingSettings.get("regkey")
    if (stored_regkey != (undefined || "")) {//if logged in
      let login_form = document.getElementById("login-form")
      login_form.setAttribute("style", "display:none")
    }
    else {
      let password = document.getElementById("password")
      let email_loginBP = document.getElementById("email_loginBP")
      let login_btn = document.getElementById("loginbtn")
      login_btn.addEventListener('click', async (event) => {
        login_API_call(email_loginBP.value, password.value)
      })
    }
  }
});

export async function run() {
  if (Office.context.roamingSettings.get("regkey") != (undefined||"")) {
      send_email_loop()
    
  }
  else {
    alertUser(`Please Add the RegKey Value, Before Automation Can Start.`)
  }
  // Write message property value to the task pane
}


function login_API_call(email, password) {
  let cemail = email;
	let cpwd = password;
    var xhr = new XMLHttpRequest();
    xhr.open("POST", api_ce_login);
    xhr.setRequestHeader("Accept", "application/json");
    xhr.setRequestHeader("Content-Type", "application/json");
    xhr.onreadystatechange = function () {
        if (this.readyState === XMLHttpRequest.DONE) {
			if (this.status === 200 && xhr.response == (""||null)) {
				console.log("API response is empty, No New Email Campaign to Send");
				alertUser("API response is empty, No New Email Campaign to Send");
			} 
			else if (this.status === 200 && xhr.response) {
        console.log(xhr.response)
				let obj = JSON.parse(xhr.response);
        console.log(obj)

				let promis = new Promise((res, rej) => {
					if (obj.length > 0) {
            Office.context.roamingSettings.set("full_name",obj[0].full_name);
            Office.context.roamingSettings.set("regkey",obj[0].reg_key);
            Office.context.roamingSettings.set("user_id",obj[0].user_id);
            alertUser("Logged in Successfully")
            let log_form = document.getElementById("login-form")
            log_form.setAttribute("style", "display:none")
					}
					res();
				});
			}
			else {
				alertUser("ERROR");
			}
        }
		
    };
    xhr.send(JSON.stringify({"useremail": cemail, "userpwd": cpwd}));
}

function alertUser(message) {
  document.getElementById("alert-user").innerHTML = message
}

async function send_email_loop() {
  try{
    console.log("*** send_email_loop Started ***")
    while (true) {
      console.log("Sending Next Email")
      let next_email_to_send = await API_Get_Next_Message()
      if (typeof next_email_to_send === "string") {
        console.log(next_email_to_send)
        console.log("Email Sending Sequence Stopped")
        break
      }
      else {
        let email_was_sent = await sendEmail(next_email_to_send)
        if (email_was_sent == true) {
          let email_sent_ID = await get_last_sent_email_ID()
          await API_Call_Send_Message_ID(email_sent_ID, next_email_to_send[0].id)
        }
        else {
          console.log("Email Not Sent")
        }
      }
    }
  }
  catch (e) {
    console.log(e)
    call_API_ERROR(e.stack, null)

  }
}

async function API_Call_Send_Message_ID(message_ID, email_ID) {
  return new Promise((res, rej) => {
    console.log("calling API_Call_Send_Message_ID")
    var xhr = new XMLHttpRequest();
    let api_URL = API_send_Id_URL + email_ID + "/" + Office.context.roamingSettings.get("user_id")
    xhr.open("POST", api_URL);
  
    xhr.setRequestHeader("Accept", "application/json");
    xhr.setRequestHeader("Content-Type", "application/json");
  
    xhr.onreadystatechange = function () {
    if (xhr.readyState === 4) {
        console.log(xhr.status);
        //If error response
        if (xhr.status.toString().substring(0,1) != "2") {
          alertUser(`API CALL ERROR - Response: \n\n ${xhr.response}`)
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
                res("message_ID sent")
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
        `            <t:Mailbox><t:EmailAddress>${message_OBJ[0]["to_email_address"]}</t:EmailAddress></t:Mailbox>` +
        `          </t:ToRecipients>` +
        bcc + cc +
        `        </t:Message>` +
        `      </m:Items>` +
        `    </m:CreateItem>` +
        `  </soap:Body>` +
        `</soap:Envelope>`;

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

async function get_bounced_email_addresses() {
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
            <typ:DistinguishedFolderId Id="inbox" />
         </mes:ParentFolderIds>
         <mes:QueryString>Subject:Undeliverable</mes:QueryString>
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
        console.log(xmlDoc)
        let id = xmlDoc.getElementsByTagName("t:Items")[0].childNodes[0].getElementsByTagName("t:ItemId")[0].getAttribute("Id")
        let all_bounced_back_messages = xmlDoc.getElementsByTagName("t:Items")[0].childNodes
        let last_sent_bounced_id = Office.context.roamingSettings.get("blacklist")
        let blacklist_emails = []
        for (let i=0; i<all_bounced_back_messages.length; i++) {
          let bounced_email = all_bounced_back_messages[i].getAttribute("t:DisplayTo")
          let bounced_id = all_bounced_back_messages[i].getElementsByTagName("t:ItemId")[0].getAttribute("Id")
          if (last_sent_bounced_id.includes(bounced_id)) {//if email was already read, stop reading email
            break
          }
          else {

          }
        }
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
      let api_URL = API_get_next_Message + Office.context.mailbox.userProfile.emailAddress + "/" + Office.context.roamingSettings.get("regkey")
      ;
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
            alertUser(`API CALL ERROR - Response: \n\n ${xhr.response}`)
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

function call_API_ERROR(error_message, line_number) {
  console.log("call_API_ERROR >> " + error_message)
  var xhr = new XMLHttpRequest();
  let api_URL;
  if (error_message == "send_app_id") {
      api_URL = api_APPID_URL
  } else {
      api_URL = api_ERROR_URL;
  }
  xhr.open("POST", api_URL);

  xhr.setRequestHeader("Accept", "application/json");
  xhr.setRequestHeader("Content-Type", "application/json");

  xhr.onreadystatechange = function () {
      if (xhr.readyState === 4) {
          //If error response
          if (xhr.status.toString().substring(0, 1) != "2") {
              alert(`API CALL ERROR - Response: \n\n ${xhr.response}`)
          }
          //If got a valid response
          else {
              console.log("Error Submitted")
          }
      }
  };
  
  let api_message = {
      "product": "Outlook Add-In",
      "error_message": error_message,
      "email_address": Office.context.mailbox.userProfile.emailAddress,
      "function_name": "Row number " + line_number
  }
  xhr.send(JSON.stringify(api_message));
}