// At the very top of launchevent.js
console.log("Debug: launchevent.js file loaded");

Office.onReady(() => {
    console.log("Debug: Office.onReady called");
    console.log("Debug: Registering onMessageSendHandler");
    Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
    console.log("Debug: Handler registered successfully");
}).catch((error) => {
    console.error("Debug: Office.js failed:", error);
});

function onMessageSendHandler(event) {
  console.log("DEBUG 1a");
  console.log(event);
  const item = Office.context.mailbox.item;

  let to = "";
  let from = "";
  let subject = "";
  let cc = "";
  let bcc = "";
  let body = "";
  let attachment = "";

  item.getAttachmentsAsync(function (toAttachment) {
    attachment = toAttachment.value;

    item.to.getAsync(function (toResult) {
      to = toResult.value;

      item.from.getAsync(function (fromResult) {
        from = fromResult.value;

        item.subject.getAsync(function (subjectResult) {
          subject = subjectResult.value;

          item.cc.getAsync(function (ccResult) {
            cc = ccResult.value;

            item.bcc.getAsync(function (bccResult) {
              bcc = bccResult.value;

              const selectedMode = localStorage.getItem("mail_mode") || "private";
              const fd = new FormData();
              fd.append("mode", JSON.stringify(selectedMode));
                
              item.body.getAsync("html", { asyncContext: event }, function (bodyResult) {
                const event = bodyResult.asyncContext;
                let currentBody = bodyResult.value || "";
                let appendedMessage = "<br/><br/><i>This message was sent in ${selectedMode}.</i>";
                let newBody = currentBody + appendedMessage;

                item.body.setAsync(newBody, { coercionType: "html" }, function (setResult) {
                  if (setResult.status === Office.AsyncResultStatus.Succeeded) {
                    item.body.getAsync("text", { asyncContext: event }, function (bodyResultText) {
                      body = bodyResultText.value;

                      item.getAttachmentsAsync(function (attachmentResult) {
                        const attachments = attachmentResult.value || [];

                        if (hasBlockedAttachmentNames(attachments)) {
                          event.completed({ 
                            allowEvent: false,
                            errorMessage: "Blocked attachment detected.",
                            errorMessageMarkdown: "One or more of the attachments have an invalid name"
                          });
                          return;
                        }

                        fetch("http://127.0.0.1:5000/receive_sizetoken", {
                          method: "POST"
                        })
                        .then(response => response.json())
                        .then(data => {
                          console.log("Token status:", data.status);
                          if (data.status !== 3) {
                              console.log("entered the size checking");
                              console.log(event);
                            if (hasBlockedAttachmentSize(attachments)) {
                              event.completed({ 
                                allowEvent: false,
                                errorMessage: "File size limit exceeded.",
                                errorMessageMarkdown: "One or more of the attachments exceed the maximum size limit of 5MB"
                              });
                              return;
                            }
                          }
                          
                          processEmailData();
                        })
                        .catch(error => {
                          console.error("Failed to get token status:", error);
                          processEmailData();
                        });

                        function processEmailData() {
                          const formData = new FormData();
                          formData.append("to", JSON.stringify(to));
                          formData.append("from", JSON.stringify(from));
                          formData.append("subject", subject);
                          formData.append("cc", JSON.stringify(cc));
                          formData.append("bcc", JSON.stringify(bcc));
                          formData.append("body", body);
                          formData.append("attachment", attachment);

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
                        }
                      });
                    });
                  } else {
                    event.completed({ allowEvent: false, errorMessage: "Failed to append message to email." });
                  }
                });
              });
            });
          });
        });
      });
    });
  });
}

function sendFormData(formData, event) {
  fetch("http://127.0.0.1:5000/receive_email", {
    method: "POST",
    body: formData
  })
    .then(response => response.json())
    .then(data => {
      if (data.status === "sensitive") {
        event.completed({
          allowEvent: false,
          errorMessage: "Sensitive content found.",
          errorMessageMarkdown: "This email contains confidential information in attachments."
        });
      } else {
        console.log("Email data sent successfully:", data);
        event.completed({ allowEvent: true });
      }
    })
    .catch(error => {
      console.error("Failed to send email data:", error);
      event.completed({ allowEvent: true });
    });
}

function hasBlockedAttachmentNames(attachments) {
  const blockedNames = ["virus.exe", "malware.js", "blockedfile.txt", "virus.txt", "unidentified.txt", "malware.txt"];
  return attachments.some(att => blockedNames.includes(att.name));
}

function hasBlockedAttachmentSize(attachments) {
  return attachments.some(att => att.size > 5242880);
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);



/*
console.log("Debug: launchevent.js file loaded");

Office.onReady(() => {
    console.log("Debug: Office.onReady called");
    console.log("Debug: Registering onMessageSendHandler");
    Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
    
    console.log("Debug: Handler registered successfully");
}).catch((error) => {
    console.error("Debug: Office.js failed:", error);
});

function onMessageSendHandler(event) {
  console.log("DEBUG 1a");
  console.log(event);
  const item = Office.context.mailbox.item;

  let to = "";
  let from = "";
  let subject = "";
  let cc = "";
  let bcc = "";
  let body = "";
  let attachment = "";

  item.getAttachmentsAsync(function (toAttachment) {
    attachment = toAttachment.value;

    item.to.getAsync(function (toResult) {
      to = toResult.value;

      item.from.getAsync(function (fromResult) {
        from = fromResult.value;

        item.subject.getAsync(function (subjectResult) {
          subject = subjectResult.value;

          item.cc.getAsync(function (ccResult) {
            cc = ccResult.value;

            item.bcc.getAsync(function (bccResult) {
              bcc = bccResult.value;

              item.body.getAsync("text", { asyncContext: event }, function (bodyResult) {
                const event = bodyResult.asyncContext;
                body = bodyResult.value;

                item.getAttachmentsAsync(function (attachmentResult) {
                  const attachments = attachmentResult.value || [];

                  // Check for blocked attachment names
                  if (hasBlockedAttachmentNames(attachments)) {
                    event.completed({ 
                      allowEvent: false,
                      errorMessage: "Blocked attachment detected.",
                      errorMessageMarkdown: "One or more of the attachments have an invalid name"
                    });
                    return;
                  }

                  const selectedMode = localStorage.getItem("mail_mode") || "private";
                  const fd = new FormData();
                  fd.append("mode", JSON.stringify(selectedMode));

                  fetch("http://127.0.0.1:5000/receive_sizetoken", {
                    method: "POST",
                    body: fd
                  })
                  .then(response => response.json())
                  .then(data => {
                    console.log("Token status:", data.status);
                    if (data.status !== 3) {
                        console.log("entered the size checking");
                        console.log(event);
                      if (hasBlockedAttachmentSize(attachments)) {
                        event.completed({ 
                          allowEvent: false,
                          errorMessage: "File size limit exceeded.",
                          errorMessageMarkdown: "One or more of the attachments exceed the maximum size limit of 5MB"
                        });
                        return;
                      }
                    }                    
                    processEmailData();
                  })
                  .catch(error => {
                    console.error("Failed to get token status:", error);
                    
                    processEmailData();
                  });

                  function processEmailData() {
                    const formData = new FormData();
                    formData.append("to", JSON.stringify(to));
                    formData.append("from", JSON.stringify(from));
                    formData.append("subject", subject);
                    formData.append("cc", JSON.stringify(cc));
                    formData.append("bcc", JSON.stringify(bcc));
                    formData.append("body", body);
                    formData.append("attachment", attachment);

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
                  }
                });
              });
            });
          });
        });
      });
    });
  });
}

function sendFormData(formData, event) {
  fetch("http://127.0.0.1:5000/receive_email", {
    method: "POST",
    body: formData
  })
    .then(response => response.json())
    .then(data => {
      if (data.status === "sensitive") {
        event.completed({
          allowEvent: false,
          errorMessage: "Sensitive content found.",
          errorMessageMarkdown: "This email contains confidential information in attachments."
        });
      } else {
        console.log("Email data sent successfully:", data);
        event.completed({ allowEvent: true });
      }
    })
    .catch(error => {
      console.error("Failed to send email data:", error);
      event.completed({ allowEvent: true });
    });
}

function hasBlockedAttachmentNames(attachments) {
  const blockedNames = ["virus.exe", "malware.js", "blockedfile.txt", "virus.txt", "unidentified.txt", "malware.txt"];
  return attachments.some(att => blockedNames.includes(att.name));
}

function hasBlockedAttachmentSize(attachments) {
  return attachments.some(att => att.size > 5242880);
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
*/
