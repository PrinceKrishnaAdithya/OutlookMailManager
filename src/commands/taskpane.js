function setModePrivate(event) {
    localStorage.setItem("mail_mode", "private");
    event.completed();
}
