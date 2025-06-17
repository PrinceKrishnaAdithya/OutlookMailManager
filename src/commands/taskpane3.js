function setModePublic(event) {
    localStorage.setItem("mail_mode", "public");
    event.completed();
}
