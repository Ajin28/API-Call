
async function api() {
    console.log("stated")
    let response = await fetch("https://aybotazurefunctionapp.azurewebsites.net/api/GenerateAYBotToken?code=DjpE11lNWCsOaKN1OZIUjOHhAPb/4ze6JdCNqGS2Qcddp5VSLeESdQ==");
    let token = await response.json();
    console.log(token)
    document.getElementById("token").textContent = token.source
}