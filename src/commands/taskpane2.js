function setModeProtected(event) {
    localStorage.setItem("mail_mode", "protected");
    event.completed();
}
