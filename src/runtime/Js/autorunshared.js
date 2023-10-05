// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// Contains code for event-based activation on Outlook on web, on Windows, and on Mac (new UI preview).

/**
 * Checks if signature exists.
 * If not, displays the information bar for user.
 * If exists, insert signature into new item (appointment or message).
 * @param {*} eventObj Office event object
 * @returns
 */
function checkSignature(eventObj) {
  console.log("Office.context", Office.context)
  
  getCurrentUser(eventObj);
}

function getCurrentUser(eventObj){
  Office.context.mailbox.getCallbackTokenAsync({
    isRest: true
  }, function (result) {
    console.log(result);  

    //const apiUrl = Office.context.mailbox.restUrl + "/v2.0/Users('" + Office.context.mailbox.userProfile.emailAddress + "')";
    //const apiUrl = Office.context.mailbox.restUrl + "/v1.0/Users('" + Office.context.mailbox.userProfile.emailAddress + "')/Contacts";
    const apiUrl = Office.context.mailbox.restUrl + "/beta/Users('" + Office.context.mailbox.userProfile.emailAddress + "')/people";

    $.ajax({  
      method: 'GET',  
      url: apiUrl,  
      headers: {  
          'Authorization': 'Bearer ' + result.value,  
          'Content-Type': 'application/json'  
      },  
    }).success(function(response) {  
        const curUser = response.value.filter(x => x.UserPrincipalName == Office.context.mailbox.userProfile.emailAddress)[0];
        console.log(curUser);
        setSignatureTemplate(curUser, eventObj);
    }).error(function(error) {
       console.log(error);
       setSignatureTemplate({
        Title: "",
        Phones: [{ Type: "Business", Number: ""}],
        OfficeLocation: ""
       }, eventObj);
    });
  });
}

function setSignatureTemplate(curUser, eventObj){
  const busNo = curUser.Phones.filter(x => x.Type == "Business")[0];
  const emailTemplate = 
  
			'<span style="font-size:14px"><b>'+ Office.context.mailbox.userProfile.displayName +'</b></span>'+
      '<br />'+
			'<span style="font-size:14px">'+ curUser.Title +'<span>'+

    '<br />'+
    '<br />'+
    '<table style="border:0;border-spacing:0;" cellspacing="0">'+
      '<tr>'+
        '<td style="padding-right: 10px;">'+
          '<img height="90" width="150" src="https://raw.githubusercontent.com/ejdomantay/lidiar-group/main/Lidiar%20Northern%20Territory%20Logo.png"></img>'+
        '</td>'+
        '<td>'+
          '<table style="border:0;border-spacing:0;" cellspacing="0">'+		
            '<tr>'+
              '<td style="background-color: red; width: 5px; height: 95px;">'+
            '</td>'+
            '</tr>'+
          '</table>'+
        '</td>'+
        '<td style="padding-left:5px;">'+
          '<table style="border:0;border-spacing:0; font-size:14px; line-height: 16px;" cellspacing="0">	'+	
            '<tr>'+
              '<td>'+
                  '<span><b>Lidiar Northern Territory Pty Ltd, part of the Lidiar Group</b></span>'+
                  '</td>'+
            '</tr>'+
            '<tr>'+
              '<td>'+
                  '<span style="color:red">m.</span>'+
                  '<span> '+ (busNo ? busNo.Number : "") +'</span>'+
                  '<span> | <span>'+
                  '<span style="color:red">e.</span>'+
                  '<span> '+ Office.context.mailbox.userProfile.emailAddress +'</span>'+
                  '</td>'+
            '</tr>'+
            '<tr>'+
              '<td>'+
                '<span style="color:red">o.</span>'+
                '<span> '+ (curUser.OfficeLocation ? curUser.OfficeLocation : "Level 3, 240 Queen Street Brisbane QLD 4000") +'</span>'+
                  
              '</td>'+
            '</tr>'+
            '<tr>'+
            '<td>'+
              '<span style="color:red">w.</span> <span>www.lidiargroup.com.au</span>'+
                  '</td>'+
              '</tr>'+
            '<tr>'+
              
            '<td>'+

              '<img style="" height="15" width="15" src="https://upload.wikimedia.org/wikipedia/commons/thumb/c/c9/Microsoft_Office_Teams_%282018%E2%80%93present%29.svg/2203px-Microsoft_Office_Teams_%282018%E2%80%93present%29.svg.png"></img>'+
                
                '	<a href="https://teams.microsoft.com/l/chat/0/0?users='+ Office.context.mailbox.userProfile.emailAddress +'">'+
          
                ' Chat with me on Teams'+
                  '	</a>'+
                '	</td>'+
              '</tr>'+	
            '</table>'+
          '</td>'+
        '</tr>'+
      '</table>'+
    '<div>'+
    '<p style="color:gray; font-size:13px;padding-top: 16px;">'+
          'Disclaimer: The information contained in this email is intended only for the use of the person(s) to whom it is addressed and may be confidential or contain privileged information. All information contained in this electronic communication is solely for the use of the individual(s) or entity to which it was addressed. If you are not the intended recipient you are hereby notified that any perusal, use, distribution, copying or disclosure is strictly prohibited. If you have received this email in error please immediately advise us by return email and delete the email without making a copy.'+
          '</p>'+
          '</div>'+
    '<div>'+
    '<p style="font-size:13px;padding-top: 10px;">'+
          'Please consider the environment before printing this email.'+
          '</p>'+
      '</div>'
 


    if(Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Appointment)
    {
      Office.context.mailbox.item.body.setAsync(
        "<br/><br/>" + emailTemplate,
        {
          coercionType: "html",
          asyncContext: eventObj,
        },
        function (asyncResult) {
    
          asyncResult.asyncContext.completed();
        }
      );
    }
    else{
    Office.context.mailbox.item.body.setSignatureAsync(
      emailTemplate,
      {
        coercionType: "html",
        asyncContext: eventObj,
      },
      function (asyncResult) {
        asyncResult.asyncContext.completed();
      }
    );
  }
}

function requestToken() {  
  console.log(Office.context); 
  Office.context.auth.getAccessTokenAsync(function(result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
        const token = result.value;
        console.log("token", token); 
    } else {
        console.log("Error obtaining token", result.error);
    }
});
}

function getUserInformation(token, mailAddress) {  
  $.ajax({  
      method: 'GET',  
      url: "https://graph.microsoft.com/v1.0/users/" + mailAddress,  
      headers: {  
          'Authorization': 'Bearer ' + token,  
          'Content-Type': 'application/json'  
      },  
  }).success(function(response) {  
      console.log(response);  
  }).error(function(error) {});  
} 
/**
 * For Outlook on Windows and on Mac only. Insert signature into appointment or message.
 * Outlook on Windows and on Mac can use setSignatureAsync method on appointments and messages.
 * @param {*} compose_type The compose type (reply, forward, newMail)
 * @param {*} user_info Information details about the user
 * @param {*} eventObj Office event object
 */
function insert_auto_signature(compose_type, user_info, eventObj) {
  let template_name = get_template_name(compose_type);
  let signature_info = get_signature_info(template_name, user_info);
  addTemplateSignature(signature_info, eventObj);
}

/**
 * 
 * @param {*} signatureDetails object containing:
 *  "signature": The signature HTML of the template,
    "logoBase64": The base64 encoded logo image,
    "logoFileName": The filename of the logo image
 * @param {*} eventObj 
 * @param {*} signatureImageBase64 
 */
function addTemplateSignature(signatureDetails, eventObj, signatureImageBase64) {
  if (is_valid_data(signatureDetails.logoBase64) === true) {
    //If a base64 image was passed we need to attach it.
    Office.context.mailbox.item.addFileAttachmentFromBase64Async(
      signatureDetails.logoBase64,
      signatureDetails.logoFileName,
      {
        isInline: true,
      },
      function (result) {
        //After image is attached, insert the signature
        Office.context.mailbox.item.body.setSignatureAsync(
          signatureDetails.signature,
          {
            coercionType: "html",
            asyncContext: eventObj,
          },
          function (asyncResult) {
            asyncResult.asyncContext.completed();
          }
        );
      }
    );
  } else {
    //Image is not embedded, or is referenced from template HTML
    Office.context.mailbox.item.body.setSignatureAsync(
      signatureDetails.signature,
      {
        coercionType: "html",
        asyncContext: eventObj,
      },
      function (asyncResult) {
        asyncResult.asyncContext.completed();
      }
    );
  }
}

/**
 * Creates information bar to display when new message or appointment is created
 */
function display_insight_infobar() {
  Office.context.mailbox.item.notificationMessages.addAsync("fd90eb33431b46f58a68720c36154b4a", {
    type: "insightMessage",
    message: "Please set your signature with the Office Add-ins sample.",
    icon: "Icon.16x16",
    actions: [
      {
        actionType: "showTaskPane",
        actionText: "Set signatures",
        commandId: get_command_id(),
        contextData: "{''}",
      },
    ],
  });
}

/**
 * Gets template name (A,B,C) mapped based on the compose type
 * @param {*} compose_type The compose type (reply, forward, newMail)
 * @returns Name of the template to use for the compose type
 */
function get_template_name(compose_type) {
  if (compose_type === "reply") return Office.context.roamingSettings.get("reply");
  if (compose_type === "forward") return Office.context.roamingSettings.get("forward");
  return Office.context.roamingSettings.get("newMail");
}

/**
 * Gets HTML signature in requested template format for given user
 * @param {\} template_name Which template format to use (A,B,C)
 * @param {*} user_info Information details about the user
 * @returns HTML signature in requested template format
 */
function get_signature_info(template_name, user_info) {
  if (template_name === "templateB") return get_template_B_info(user_info);
  if (template_name === "templateC") return get_template_C_info(user_info);
  return get_template_A_info(user_info);
}

/**
 * Gets correct command id to match to item type (appointment or message)
 * @returns The command id
 */
function get_command_id() {
  if (Office.context.mailbox.item.itemType == "appointment") {
    return "MRCS_TpBtn1";
  }
  return "MRCS_TpBtn0";
}

/**
 * Gets HTML string for template A
 * Embeds the signature logo image into the HTML string
 * @param {*} user_info Information details about the user
 * @returns Object containing:
 *  "signature": The signature HTML of template A,
    "logoBase64": The base64 encoded logo image,
    "logoFileName": The filename of the logo image
 */
function get_template_A_info(user_info) {
  const logoFileName = "sample-logo.png";
  let str = "";
  if (is_valid_data(user_info.greeting)) {
    str += user_info.greeting + "<br/>";
  }

  str += "<table>";
  str += "<tr>";
  // Embed the logo using <img src='cid:...
  str +=
    "<td style='border-right: 1px solid #000000; padding-right: 5px;'><img src='cid:" +
    logoFileName +
    "' alt='MS Logo' width='24' height='24' /></td>";
  str += "<td style='padding-left: 5px;'>";
  str += "<strong>" + user_info.name + "</strong>";
  str += is_valid_data(user_info.pronoun) ? "&nbsp;" + user_info.pronoun : "";
  str += "<br/>";
  str += is_valid_data(user_info.job) ? user_info.job + "<br/>" : "";
  str += user_info.email + "<br/>";
  str += is_valid_data(user_info.phone) ? user_info.phone + "<br/>" : "";
  str += "</td>";
  str += "</tr>";
  str += "</table>";

  // return object with signature HTML, logo image base64 string, and filename to reference it with.
  return {
    signature: str,
    logoBase64:
      "iVBORw0KGgoAAAANSUhEUgAAACIAAAAiCAYAAAA6RwvCAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsQAAA7EAZUrDhsAAAEeSURBVFhHzdhBEoIwDIVh4EoeQJd6YrceQM+kvo5hQNokLymO/4aF0/ajlBl1fL4bEp0uj3K9XQ/lGi0MEcB3UdD0uVK1EEj7TIuGeBaKYCgIswCLcUMid8mMcUEiCMk71oRYE+Etsd4UD0aFeBBSFtOEMAgpg6lCIggpitlAMggpgllBeiAkFjNDeiIkBlMgeyAkL6Z6WJdlEJJnjvF4vje/BvRALNN23tyRXzVpd22dHSZtLhjMHemB8cxRINZZyGCssbL2vCN7YLwItHo0PTEMAm3OSA8Mi0DVw5rBRBCoCkERTBSBmhDEYDII5PqlZy1iZSGQuiOSZ6JW3rEuCIpgmDFuCGImZuEUBHkWiOweDUHaQhEE+pM/aobhBZaOpYLJeeeoAAAAAElFTkSuQmCC",
    logoFileName: logoFileName,
  };
}

/**
 * Gets HTML string for template B
 * References the signature logo image from the HTML
 * @param {*} user_info Information details about the user
 * @returns Object containing:
 *  "signature": The signature HTML of template B,
    "logoBase64": null since this template references the image and does not embed it ,
    "logoFileName": null since this template references the image and does not embed it
 */
function get_template_B_info(user_info) {
  let str = "";
  if (is_valid_data(user_info.greeting)) {
    str += user_info.greeting + "<br/>";
  }

  str += "<table>";
  str += "<tr>";
  // Reference the logo using a URI to the web server <img src='https://...
  str +=
    "<td style='border-right: 1px solid #000000; padding-right: 5px;'><img src='https://officedev.github.io/Office-Add-in-samples/Samples/outlook-set-signature/assets/sample-logo.png' alt='Logo' /></td>";
  str += "<td style='padding-left: 5px;'>";
  str += "<strong>" + user_info.name + "</strong>";
  str += is_valid_data(user_info.pronoun) ? "&nbsp;" + user_info.pronoun : "";
  str += "<br/>";
  str += user_info.email + "<br/>";
  str += is_valid_data(user_info.phone) ? user_info.phone + "<br/>" : "";
  str += "</td>";
  str += "</tr>";
  str += "</table>";

  return {
    signature: str,
    logoBase64: null,
    logoFileName: null,
  };
}

/**
 * Gets HTML string for template C
 * @param {*} user_info Information details about the user
 * @returns Object containing:
 *  "signature": The signature HTML of template C,
    "logoBase64": null since there is no image,
    "logoFileName": null since there is no image
 */
function get_template_C_info(user_info) {
  let str = "";
  if (is_valid_data(user_info.greeting)) {
    str += user_info.greeting + "<br/>";
  }

  str += user_info.name;

  return {
    signature: str,
    logoBase64: null,
    logoFileName: null,
  };
}

/**
 * Validates if str parameter contains text.
 * @param {*} str String to validate
 * @returns true if string is valid; otherwise, false.
 */
function is_valid_data(str) {
  return str !== null && str !== undefined && str !== "";
}

Office.actions.associate("checkSignature", checkSignature);
