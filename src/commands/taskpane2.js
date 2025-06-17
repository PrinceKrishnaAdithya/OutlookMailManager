async function setModeProtected(event) {
    await OfficeRuntime.storage.setItem("mail_mode", "protected");
    console.log("Set mode to protected using OfficeRuntime.storage");
    event.completed();
}
