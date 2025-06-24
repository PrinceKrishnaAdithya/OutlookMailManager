//This is the function called when you press the protected button
//In your browser localstorage, it saves the value "protected" associated with the key "mail_mode"
//in launchevent, this localstorage is accessed and the value is read
function runTaskpane2Logic(event) 
{
    localStorage.setItem("mail_mode", "protected");
    console.log("Mode set to protected");
    event.completed();
}
