function setModeProtected(event) {
  OfficeRuntime.storage.setItem("mail_mode", "protected")
    .then(() => {
      console.log("Mail mode set to protected");
      event.completed();
    })
    .catch((error) => {
      console.error("Storage error:", error);
      event.completed();
    });
}

window.setModeProtected = setModeProtected;
