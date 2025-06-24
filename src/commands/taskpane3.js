//This is the function called when you press the public button
//In your browser localstorage, it saves the value "public" associated with the key "mail_mode"
//in launchevent, this localstorage is accessed and the value is read
function runTaskpane3Logic(event) 
{
    localStorage.setItem("mail_mode", "public");
    console.log("Mode set to public");
    event.completed();
}
