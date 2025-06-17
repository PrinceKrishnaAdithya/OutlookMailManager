function setModePrivate(event) {
    console.log("sending");
    localStorage.setItem("mail_mode", "private");
    console.log("sent");
    event.completed();
}
