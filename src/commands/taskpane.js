function setModePrivate(event) {
    localStorage.setItem("mail_mode", "private");
    console.log("✳️ Mode set to private");
    event.completed();
}
