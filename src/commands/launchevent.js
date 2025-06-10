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

  console.log("log 1");
  console.log(event);
  const item = Office.context.mailbox.item;

  let to = "";
  let from = "";
  let subject = "";
  let cc = "";
  let bcc = "";
  let body = "";
  const attachmentsMap = {};

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
                let pending = attachments.length;

                if (pending === 0) {
                  console.log("log 2");
                  console.log(event);
                  // No attachments, just download the metadata
                  downloadEmailData(to, from, subject, cc, bcc, body, attachmentsMap);
                    
                  event.completed({ allowEvent: true });
                } else {
                  attachments.forEach(att => {
                    item.getAttachmentContentAsync(att.id, function (contentResult) {
                      if (contentResult.status === Office.AsyncResultStatus.Succeeded) {
                        const content = contentResult.value.content;
                        const fileType = contentResult.value.format;
                        const filename = att.name;

                        if (fileType === "base64") {
                          const mimeType = getMimeType(filename);
                          const blob = base64ToBlob(content, mimeType);
                          attachmentsMap[filename] = blob;
                          console.log("✅ Created and stored blob for:", filename);
                        }
                      } else {
                        console.error("❌ Failed to fetch attachment:", contentResult.error);
                      }

                      pending--;
                      if (pending === 0) {
                        console.log("log 3");
                        console.log(event);
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

console.log("log 4");
console.log(event);
function downloadEmailData(to, from, subject, cc, bcc, body, attachmentsMap) {
  const emailData = {
    to,
    from,
    subject,
    cc,
    bcc,
    body,
    timestamp: new Date().toISOString(),
  };

  const emailBlob = new Blob([JSON.stringify(emailData, null, 2)], { type: "application/json" });
  triggerDownload(`${sanitizeFilename(subject || "email")}_data.json`, emailBlob);

  for (const [filename, blob] of Object.entries(attachmentsMap)) {
    console.log(`📥 Downloading attachment: ${filename}`);
    triggerDownload(filename, blob);
  }
}

console.log("log 5");
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
console.log("log 6");
  console.log(event);
function sanitizeFilename(name) {
  return name.replace(/[^a-z0-9_\-]/gi, "_").substring(0, 50);
}
console.log("log 7");
  console.log(event);
function base64ToBlob(base64, mimeType = "application/octet-stream") {
  const byteCharacters = atob(base64);
  const byteArrays = [];

  for (let offset = 0; offset < byteCharacters.length; offset += 512) {
    const slice = byteCharacters.slice(offset, offset + 512);
    const byteNumbers = new Array(slice.length);
    for (let i = 0; i < slice.length; i++) {
      byteNumbers[i] = slice.charCodeAt(i);
    }
    byteArrays.push(new Uint8Array(byteNumbers));
  }

  return new Blob(byteArrays, { type: mimeType });
}
console.log("log 9");
  console.log(event);
function getMimeType(filename) {
  const ext = filename.split('.').pop().toLowerCase();
  const map = {
    pdf: "application/pdf",
    txt: "text/plain",
    jpg: "image/jpeg",
    jpeg: "image/jpeg",
    png: "image/png",
    docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    py: "text/x-python",
    json: "application/json",
    zip: "application/zip"
  };
  return map[ext] || "application/octet-stream";
}

// Register your handler
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
