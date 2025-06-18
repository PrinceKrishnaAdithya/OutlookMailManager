function runTaskpane3Logic(event) {
    localStorage.setItem("mail_mode", "public");
    console.log("Mode set to public");
    event.completed();
}
