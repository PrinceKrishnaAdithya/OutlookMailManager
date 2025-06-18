function runTaskpane2Logic(event) {
    localStorage.setItem("mail_mode", "protected");
    console.log("Mode set to protected");
    event.completed();
}
