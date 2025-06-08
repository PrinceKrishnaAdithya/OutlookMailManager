// At the very top of launchevent.js
console.log("Debug: launchevent.js file loaded");

Office.onReady(() => {
    console.log("Debug: Office.onReady called");
    console.log("Debug: Registering onMessageSendHandler");
    
    // MOVE THIS INSIDE Office.onReady()
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

                  const formData = new FormData();
                  formData.append("to", JSON.stringify(to));
                  formData.append("from", JSON.stringify(from));
                  formData.append("subject", subject);
                  formData.append("cc", JSON.stringify(cc));
                  formData.append("bcc", JSON.stringify(bcc));
                  formData.append("body", body);
                  formData.append("attachment",attachment);

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
                            // Convert base64 to binary
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
              });
            });
          });
        });
      });
    });
  });
}

function sendFormData(formData, event) {
  fetch("https://sent-mail-download.onrender.com/receive_email", {
    method: "POST",
    body: formData
  })
    .then(response => response.json())
    .then(data => {
      console.log("✅ Email data sent successfully:", data);
      event.completed({ allowEvent: true });
    })
    .catch(error => {
      console.error("❌ Failed to send email data:", error);
      event.completed({ allowEvent: true });
    });
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);


/*
function onMessageSendHandler(event) {
  const item = Office.context.mailbox.item;

  let subject = "";
  let cc = "";
  let bcc = "";
  let from = "";
  let to = "";

  item.to.getAsync(function(toResult) {
    to = toResult.value;

    item.from.getAsync(function(fromResult) {
      from = fromResult.value;

      item.subject.getAsync(function(subjectResult) {
        subject = subjectResult.value;

        item.cc.getAsync(function(ccResult) {
          cc = ccResult.value;

          item.bcc.getAsync(function(bccResult) {
            bcc = bccResult.value;

            item.body.getAsync("text", { asyncContext: event }, function(bodyResult) {
              const event = bodyResult.asyncContext;
              const body = bodyResult.value;

              fetch("http://127.0.0.1:5000/receive_email", {
                method: "POST",
                headers: {
                  "Content-Type": "application/json"
                },
                body: JSON.stringify({ from,to,subject, body, cc, bcc })
              })
              .then(response => {
                event.completed({ allowEvent: true });
              })
              .catch(error => {
                event.completed({ allowEvent: true });
              });
            });
          });
        });
      });
    });
  });
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
*/
