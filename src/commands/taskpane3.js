console.log("DEBUG 1a");
let token = "3";
const formdata = new FormData();
formdata.append("token",JSON.stringify(token))
fetch("http://127.0.0.1:5000/receive_token", {
    method: "POST",
    body: formdata
})
.then(response => response.json())
.then(data => {
    console.log("✅ Response:", data);
})
.catch(error => {
    console.error("❌ Error:", error);
});
