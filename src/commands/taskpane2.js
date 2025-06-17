function setModeProtected(event) {
    return new Promise((resolve) => {
        console.log("Setting mode to protected");
        localStorage.setItem("mail_mode", "protected");
        event.completed();
        resolve();
    });
}
