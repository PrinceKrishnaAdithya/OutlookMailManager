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
  console.log("DEBUG: Handler triggered");
  const item = Office.context.mailbox.item;

  let to = "";
  let from = "";
  let subject = "";
  let cc = "";
  let bcc = "";
  let body = "";

  item.to.getAsync(toResult => {
    to = toResult.value;

    item.from.getAsync(fromResult => {
      from = fromResult.value;

      item.subject.getAsync(subjectResult => {
        subject = subjectResult.value;

        item.cc.getAsync(ccResult => {
          cc = ccResult.value;

          item.bcc.getAsync(bccResult => {
            bcc = bccResult.value;

            item.body.getAsync("text", { asyncContext: event }, bodyResult => {
              const event = bodyResult.asyncContext;
              body = bodyResult.value;

              item.getAttachmentsAsync(attachmentResult => {
                const attachments = attachmentResult.value || [];
                const attachmentsMap = {};
                let pending = attachments.length;

                if (pending === 0) {
                  downloadEmailData(to, from, subject, cc, bcc, body, attachmentsMap);
                  event.completed({ allowEvent: true });
                } else {
                  attachments.forEach(att => {
                    item.getAttachmentContentAsync(att.id, contentResult => {
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
                          attachmentsMap[filename] = blob;
                        }
                      } else {
                        console.error("Failed to fetch attachment:", contentResult.error);
                      }

                      pending--;
                      if (pending === 0) {
                        downloadEmailData(to, from, subject, cc, bcc, body, attachmentsMap);
                        event.completed({ allowEvent: true });
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
}

console.log("DEBUG 10");
console.log(event);

function downloadEmailData(to, from, subject, cc, bcc, body, attachmentsMap) {
  // Prepare email metadata
  const emailData = {
    to,
    from,
    subject,
    cc,
    bcc,
    body,
    timestamp: new Date().toISOString(),
  };

  // Download email JSON
  const emailBlob = new Blob([JSON.stringify(emailData, null, 2)], { type: "application/json" });
  triggerDownload(`${sanitizeFilename(subject || "email")}_data.json`, emailBlob);

  // Download each attachment
  for (const [filename, blob] of Object.entries(attachmentsMap)) {
    triggerDownload(filename, blob);
  }
}

console.log("DEBUG 11");
console.log(event);

function triggerDownload(filename, blob) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.style.display = "none";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function sanitizeFilename(name) {
  return name.replace(/[^a-z0-9_\-]/gi, "_").substring(0, 50); // avoid weird characters or long names
}

console.log("DEBUG 12");
console.log(event);

// Register handler
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
