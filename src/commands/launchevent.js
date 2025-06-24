//Simple console log for debugging
console.log("Debug: launchevent.js file loaded");


//The onready() function will make the code wait till the office resources are all ready
//without onready() the code starts executing before office is brought in properly
Office.onReady(() => {
    console.log("Debug: Office.onReady called");
    console.log("Debug: Registering onMessageSendHandler");
    Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
    console.log("Debug: Handler registered successfully");
}).catch((error) => {
    console.error("Debug: Office.js failed:", error);
});


//This is the main function which executes when you click on the send mail button
//it is an event based activated function
function onMessageSendHandler(event) 
{


//We are defining a variable item which stores the mail information
//defining several variables to store the information to
  console.log("DEBUG 1a");
  const item = Office.context.mailbox.item;
  let to = "";
  let from = "";
  let subject = "";
  let cc = "";
  let bcc = "";
  let body = "";
  let attachment = "";


//The getAsync() helps to access the mailitem and extract the info we want
//Here we are accessing the 'to' information and storing it into variable to
    item.to.getAsync(function (toResult) {
      to = toResult.value;


//Here we are accessing the 'from' information and storing it into variable from
      item.from.getAsync(function (fromResult) {
        from = fromResult.value;


//Here we are accessing the 'subject' information and storing it into variable subject
        item.subject.getAsync(function (subjectResult) {
          subject = subjectResult.value;


//Here we are accessing the 'cc' information and storing it into variable cc
          item.cc.getAsync(function (ccResult) {
            cc = ccResult.value;


//Here we are accessing the 'bcc' information and storing it into variable bcc
            item.bcc.getAsync(function (bccResult) {
              bcc = bccResult.value;


//Here we are accessing the 'attachments' information and storing it into array variable attachments as there can be multiple attachments
              item.getAttachmentsAsync(function (attachmentResult) {
                const attachments = attachmentResult.value || [];


//First condition is checked immediately, the file name. if it does contain a blocked name, then the event is not allowed to pass and a error message pop up
                if (hasBlockedAttachmentNames(attachments)) {
                  event.completed({ 
                    allowEvent: false,
                    errorMessage: "Blocked attachment detected.",
                    errorMessageMarkdown: "One or more of the attachments have an invalid name"
                  });
                  return;
                }


//The mail_mode is a key and it has a value, by default the value is null but if the user had pressed any of the private,public,protected buttons. they would have stored some value
//mail_mode is stored in the browser local storage
                const selectedMode = localStorage.getItem("mail_mode");


//If the user does not select any mode, an error message pops up telling the user to select a mail mode
                if (!selectedMode || selectedMode.trim() === "") {
                  event.completed({
                    allowEvent: false,
                    errorMessage: "Mail mode not selected.",
                    errorMessageMarkdown: "Please select a mail mode before sending the email."
                  });
                  alert("Please select a mail mode (e.g., private/public) before sending the email.");
                  return;
                }


//after making sure the user has selected a mail mode, we take that information and send it to the python backend as a json
                const fd = new FormData();
                fd.append("mode", JSON.stringify(selectedMode));
                fetch("http://127.0.0.1:5000/receive_sizetoken", {
                  method: "POST",
                  body: fd
                })
                .then(response => response.json())
                .then(data => {
                  console.log("Token status:", data.status);


//The python backend after receiving the mail_mode checks which mode it is and sends the front end a token
//1 for private, 2 for protected and 3 for public
//it checks the attachment size only for the private and protected modes
                  if (data.status !== 3) {
                    if (hasBlockedAttachmentSize(attachments)) {
                      event.completed({ 
                        allowEvent: false,
                        errorMessage: "File size limit exceeded.",
                        errorMessageMarkdown: "One or more of the attachments exceed the maximum size limit of 5MB"
                      });
                      return;
                    }
                  }
                  appendMessageAndSend();
                })
                .catch(error => {
                  console.error("Failed to get token status:", error);
                  appendMessageAndSend();
                });


//This function is used for appending an automatic message to the body of every mail you send
//Whatever mode you have selected. It adds a line specifiying that the mail was sent in said mode
                function appendMessageAndSend() {
                  item.body.getAsync("html", { asyncContext: event }, function (bodyResult) {
                    const event = bodyResult.asyncContext;
                    let currentBody = bodyResult.value || "";
                    let appendedMessage = `<br/><br/><i>This message was sent under ${selectedMode} constraint.</i><!-- MailManagerAppended -->`;
                    if (!currentBody.includes("<!-- MailManagerAppended -->")) {
                      let newBody = currentBody;
                      item.body.setAsync(newBody, { coercionType: "html" }, function (setResult) {
                        if (setResult.status === Office.AsyncResultStatus.Succeeded) {
                          continueSend();
                        } else {
                          event.completed({ allowEvent: false, errorMessage: "Failed to append message to email." });
                        }
                      });
                    } else {
                      continueSend();
                    }


//Finally the mail information which we previously stored are all stored into a formdata as json strings
                    function continueSend() {
                      item.body.getAsync("text", { asyncContext: event }, function (bodyResultText) {
                        body = bodyResultText.value;

                        const formData = new FormData();
                        formData.append("to", JSON.stringify(to));
                        formData.append("from", JSON.stringify(from));
                        formData.append("subject", subject);
                        formData.append("cc", JSON.stringify(cc));
                        formData.append("bcc", JSON.stringify(bcc));
                        formData.append("body", body);
                        formData.append("attachment", attachment);


//pending counts how many attachments user has attached
                        let pending = attachments.length;
                        if (pending === 0) {
                          sendFormData(formData, event);
                        } else {
                          attachments.forEach(att => {
                            item.getAttachmentContentAsync(att.id, function (contentResult) {
                              if (contentResult.status === Office.AsyncResultStatus.Succeeded) {
                                const content = contentResult.value.content;
                                const fileType = contentResult.value.format;
                                const filename = att.name;


//As the attachments cannot be sent directly they are converted to a format which the frontend can send to python backend
                                if (fileType === "base64") {
                                  const byteCharacters = atob(content);
                                  const byteArrays = [];
                                  for (let offset = 0; offset < byteCharacters.length; offset += 512) {
                                    const slice = byteCharacters.slice(offset, offset + 512);
                                    const byteNumbers = new Array(slice.length);
                                    for (let i = 0; i < slice.length; i++) {
                                      byteNumbers[i] = slice.charCodeAt(i);
                                    }
                                    const byteArray = new Uint8Array(byteNumbers);
                                    byteArrays.push(byteArray);
                                  }
                                  const blob = new Blob(byteArrays, { type: "application/octet-stream" });
                                  formData.append("attachments", blob, filename);
                                }
                                pending--;
                                if (pending === 0) {
                                  sendFormData(formData, event);
                                }
                              } else {
                                console.error("Attachment fetch error:", contentResult.error);
                                pending--;
                                if (pending === 0) {
                                  sendFormData(formData, event);
                                }
                              }
                            });
                          });
                        }
                      });
                    }
                  });
                }
              });
            });
          });
        });
      });
    });
}




//The sendformdata() function is what sends the information to the backend to be saved.
//It calls the python backend running locally and sends the json file containing all the mail info.
function sendFormData(formData, event) {
  fetch("http://127.0.0.1:5000/receive_email", {
    method: "POST",
    body: formData
  })
    .then(response => response.json())
    .then(data => {


//Checks the attachment content for secret or confidential keywords,This is fully checked in the python backend
      if (data.status === "sensitive") {
        event.completed({
          allowEvent: false,
          errorMessage: "Sensitive content found.",
          errorMessageMarkdown: "This email contains confidential information in attachments."
        });
        return;
      }


//This function is used for appending an automatic message to the body of every mail you send
//Whatever mode you have selected. It adds a line specifiying that the mail was sent in said mode
//localstorage.clear() makes sure once the mail is sent, the mail_mode is reset to null
      const item = Office.context.mailbox.item;
      item.body.getAsync("html", function (bodyResult) {
        let currentBody = bodyResult.value || "";
        const selectedMode = localStorage.getItem("mail_mode") || "private";
        const appendedMessage = `<br/><br/><i>This message was sent under ${selectedMode} constraint.</i><!-- MailManagerAppended -->`;
        if (!currentBody.includes("<!-- MailManagerAppended -->")) {
          const newBody = currentBody + appendedMessage;
          item.body.setAsync(newBody, { coercionType: "html" }, function (setResult) {
            if (setResult.status === Office.AsyncResultStatus.Succeeded) {
              localStorage.clear(); 
              event.completed({ allowEvent: true });
            } else {
              event.completed({ allowEvent: false, errorMessage: "Failed to append privacy message." });
            }
          });
        } else {
          localStorage.clear();
          event.completed({ allowEvent: true });
        }
      });
    })
    .catch(error => {
      console.error("Failed to send email data:", error);
      localStorage.clear(); 
      event.completed({ allowEvent: true });
    });
}


//This is the function to check the attachments names, you can always add more names which you want blocked
function hasBlockedAttachmentNames(attachments) 
{
  const blockedNames = ["virus.exe", "malware.js", "blockedfile.txt", "virus.txt", "unidentified.txt", "malware.txt"];
  return attachments.some(att => blockedNames.includes(att.name));
}


//This is the function that checks the assignment size, currently it checks if attachment is larger than 5mb
function hasBlockedAttachmentSize(attachments) 
{
  return attachments.some(att => att.size > 5242880);
}


//The actions.associate associates the function name with the actual function
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
