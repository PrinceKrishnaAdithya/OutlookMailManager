function runTaskpaneLogic(event) {
    console.log("DEBUG 1a");
    let token = "1";
    const formdata = new FormData();
    formdata.append("token", JSON.stringify(token));

    fetch("http://127.0.0.1:5000/receive_token", {
        method: "POST",
        body: formdata
    })
    .then(response => response.json())
    .then(data => {
        console.log(" Response:", data);
        event.completed();
    })
    .catch(error => {
        console.error("Error:", error);
        event.completed();
    });
}
