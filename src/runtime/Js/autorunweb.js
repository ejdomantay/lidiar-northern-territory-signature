// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// This file contains code only used by autorunweb.html when loaded in Outlook on the web.

Office.initialize = function (reason) {};

/**
 * For Outlook on the web, insert signature into appointment or message.
 * Outlook on the web does not support using setSignatureAsync on appointments,
 * so this method will update the body directly.
 * @param {*} compose_type The compose type (reply, forward, newMail)
 * @param {*} user_info Information details about the user
 * @param {*} eventObj Office event object
 */
function insert_auto_signature(compose_type, user_info, eventObj) {
  // const emailTemplate = 
  
	// 		'<span style="font-size:14px"><b>'+ Office.context.mailbox.userProfile.displayName +'</b></span>'+
  //     '<br />'+
	// 		'<span style="font-size:14px">%%Title%%<span>'+

  //   '<br />'+
  //   '<br />'+
  //   '<table style="border:0;border-spacing:0;" cellspacing="0">'+
  //     '<tr>'+
  //       '<td style="padding-right: 20px;">'+
  //         '<img  height="82" width="120" src="https://lidiargrouplive.com.au/wp-content/uploads/2019/08/cropped-header-logo-square-1.png"></img>'+
  //       '</td>'+
  //       '<td>'+
  //         '<table style="border:0;border-spacing:0;" cellspacing="0">'+		
  //           '<tr>'+
  //             '<td style="background-color: red; width: 5px; height: 95px;">'+
  //           '</td>'+
  //           '</tr>'+
  //         '</table>'+
  //       '</td>'+
  //       '<td style="padding-left:5px;">'+
  //         '<table style="border:0;border-spacing:0; font-size:14px; line-height: 16px;" cellspacing="0">	'+	
  //           '<tr>'+
  //             '<td>'+
  //                 '<span><b>Lidiar Group Pty Ltd</b></span>'+
  //                 '</td>'+
  //             '</tr>'+
  //           '<tr>'+
  //           '<td>'+
  //             '<span style="color:red">m.</span>'+
  //                 '<span>%%MobilePhone%%</span>'+
  //                 '<span>|<span>'+
  //                 '<span style="color:red">e.</span>'+
  //                 '<span>'+ Office.context.mailbox.userProfile.emailAddress +'</span>'+
  //                 '</td>'+
  //             '</tr>'+
  //           '<tr>'+
  //           '<td>'+
  //             '<span style="color:red">o.</span>'+
  //                 '<span>%%StreetAddress%%, %%City%% %%StateOrProvince%% %%PostalCode%%</span>'+
                  
  //                 '</td>'+
  //             '</tr>'+
  //           '<tr>'+
  //           '<td>'+
  //             '<span style="color:red">w.</span> <span>www.lidiargroup.com.au</span>'+
  //                 '</td>'+
  //             '</tr>'+
  //           '<tr>'+
              
  //           '<td>'+

  //             '<img style="" height="15" width="15" src="https://upload.wikimedia.org/wikipedia/commons/thumb/c/c9/Microsoft_Office_Teams_%282018%E2%80%93present%29.svg/2203px-Microsoft_Office_Teams_%282018%E2%80%93present%29.svg.png"></img>'+
                
  //               '	<a href="https://teams.microsoft.com/l/chat/0/0?users='+ Office.context.mailbox.userProfile.emailAddress +'">'+
          
  //               'Chat with me on Teams'+
  //                 '	</a>'+
  //               '	</td>'+
  //             '</tr>'+	
  //           '</table>'+
  //         '</td>'+
  //       '</tr>'+
  //     '</table>'+
  //   '<div>'+
  //   '<p style="color:gray; font-size:13px;padding-top: 16px;">'+
  //         'Disclaimer: The information contained in this email is intended only for the use of the person(s) to whom it is addressed and may be confidential or contain privileged information. All information contained in this electronic communication is solely for the use of the individual(s) or entity to which it was addressed. If you are not the intended recipient you are hereby notified that any perusal, use, distribution, copying or disclosure is strictly prohibited. If you have received this email in error please immediately advise us by return email and delete the email without making a copy.'+
  //         '</p>'+
  //         '</div>'+
  //   '<div>'+
  //   '<p style="font-size:13px;padding-top: 10px;">'+
  //         'Please consider the environment before printing this email.'+
  //         '</p>'+
  //     '</div>'

  //   if(Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Appointment)
  //   {
  //     Office.context.mailbox.item.body.setAsync(
  //       "<br/><br/>" + emailTemplate,
  //       {
  //         coercionType: "html",
  //         asyncContext: eventObj,
  //       },
  //       function (asyncResult) {
    
  //         asyncResult.asyncContext.completed();
  //       }
  //     );
  //   }
  //   else{
  //     Office.context.mailbox.item.body.setSignatureAsync(
  //       emailTemplate,
  //       {
  //         coercionType: "html",
  //         asyncContext: eventObj,
  //       },
  //       function (asyncResult) {
  //         asyncResult.asyncContext.completed();
  //       }
  //     );
  // }
}

/**
 * For Outlook on the web, set signature for current appointment
* @param {*} signatureDetails object containing:
 *  "signature": The signature HTML of the template,
    "logoBase64": The base64 encoded logo image,
    "logoFileName": The filename of the logo image
 * @param {*} eventObj Office event object
 */
function set_body(signatureDetails, eventObj) {

  if (is_valid_data(signatureDetails.logoBase64) === true) {
    //If a base64 image was passed we need to attach it.
    Office.context.mailbox.item.addFileAttachmentFromBase64Async(
      signatureDetails.logoBase64,
      signatureDetails.logoFileName,
      {
        isInline: true,
      },
      function (result) { 
        Office.context.mailbox.item.body.setAsync(
        "<br/><br/>" + signatureDetails.signature,
        {
          coercionType: "html",
          asyncContext: eventObj,
        },
        function (asyncResult) {

          asyncResult.asyncContext.completed();
        }
      );
    });
  } else {
    Office.context.mailbox.item.body.setAsync(
      "<br/><br/>" + signatureDetails.signature,
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
