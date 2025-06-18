function runTaskpaneLogic(event) {
    localStorage.setItem("mail_mode", "private");
    console.log("Mode set to private");
    event.completed();
}
