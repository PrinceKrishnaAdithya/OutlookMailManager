//This is the function called when you press the private button
//In your browser localstorage, it saves the value "private" associated with the key "mail_mode"
//in launchevent, this localstorage is accessed and the value is read
function runTaskpaneLogic(event) 
{
    localStorage.setItem("mail_mode", "private");
    console.log("Mode set to private");
    event.completed();
}
