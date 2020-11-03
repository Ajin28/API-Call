
async function api() {
    console.log("stated")
    let response = await fetch("https://aybotazurefunctionapp.azurewebsites.net/api/GenerateAYBotToken?code=DjpE11lNWCsOaKN1OZIUjOHhAPb/4ze6JdCNqGS2Qcddp5VSLeESdQ==");
    let token = await response.json();
    console.log(token)
    document.getElementById("ay_bot").src = token.source
}

function toggleChatbox(slidedirection) {
    $("#ay_ChatBotBox").toggle();
}
function switchChatbox() {
    $('.setAlignment').toggleClass('ay_leftAlign').toggleClass('ay_rightAlign');
}

function init() {
    $('head').append(
        [
            $('<link/>', { 'rel': "stylesheet", 'href': "https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/css/fabric.min.css" }),
            $('<link/>', { 'rel': "stylesheet", 'href': "https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/css/fabric.components.min.css" }),
            $('<script/>', { 'src': "https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/js/fabric.min.js" }),
            //bootstrap
            $('<link/>', { 'rel': "stylesheet", 'href': "https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/css/bootstrap.min.css", 'integrity': "sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm", 'crossorigin': "anonymous" }),
            $('<script/>', { 'src': "https://code.jquery.com/jquery-3.2.1.slim.min.js", 'integrity': "sha384-KJ3o2DKtIkvYIK3UENzmM7KCkRr/rE9/Qpg6aAZGJwFDMVNA/GpGFF93hXpG5KkN", 'crossorigin': "anonymous" }),
            $('<script/>', { 'src': "https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.12.9/umd/popper.min.js", 'integrity': "sha384-ApNbgh9B+Y1QKtv3Rn7W3mgPxhU9K/ScQsAP7hUibX39j7fakFPskvXusvfa0b4Q", 'crossorigin': "anonymous" }),
            $('<script/>', { 'src': "https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/js/bootstrap.min.js", 'integrity': "sha384-JZR6Spejh4U02d8jOt6vLEHfe/JQGiRRSQQxSfFWpi1MquVdAyjUar5+76PVCmYl", 'crossorigin': "anonymous" }),
            //custom stylesheet
            $('<link/>', { 'rel': "stylesheet", 'href': "https://agreeyabotstorage.blob.core.windows.net/chat-wrapper/stylesheet.css" })
        ]
    )


    var header_div1 =
        $('<div/>', { "class": "ay_chatbotImgContainer" })
            .append(
                [
                    $('<img/>', { 'src': 'https://agreeyabotstorage.blob.core.windows.net/chat-wrapper/Bot Logo 2.png', 'class': 'ay_chatbotimg' }),
                    $('<span/>', { 'class': 'ay_botName', 'text': 'Mary' })
                ]
            )

    var btn1 =
        $('<button/>', { 'class': "helpicon", 'title': "Help" }).append(
            $('<span/>', { 'class': "ms-Icon ms-Icon--Help", 'aria-hidden': "true" })
        )
    var btn2 =
        $('<button/>', { 'class': "switch", 'onClick': "switchChatbox()", "title": "Switch Position" }).append(
            $('<span/>', { 'class': "ms-Icon ms-Icon--Switch", 'aria-hidden': "true" })
        )
    var btn3 =
        $('<button/>', { 'class': "minimize", 'onClick': "toggleChatbox()", 'title': "Minimize" }).append(
            $('<span/>', { 'class': "ms-Icon ms-Icon--ChromeMinimize pt-2", 'aria-hidden': "true" })
        )

    var header_div2 =
        $('<div/>', { 'class': 'ay_headerRightcontent' }).append(btn1, btn2, btn3)

    var header =
        $('<div/>', { "class": "ay_chatboxheaderContainer" })
            .append(header_div1, header_div2);

    var body_div2 =
        $('<div/>', { 'class': 'ay_chatboticon ay_rightAlign setAlignment', 'style': 'background-color: rgb(126, 211, 33);' }).append(
            $('<a/>', { 'href': "javascript:void(0)", 'onClick': "toggleChatbox()" }).append(
                $('<span/>', { 'class': "ms-Icon ms-Icon--Message" })
            )
        )

    var chatbox = $('<div/>', { 'class': 'ay_chatbox', }).append(
        $("<iframe/>", { 'id': 'ay_bot' })
    )

    var body_div1 =
        $('<div/>', { 'class': 'ay_chatboxMainWrapper ay_rightAlign setAlignment', 'id': 'ay_ChatBotBox', }).append(
            header,
            chatbox
        )


    $('body').append(
        $('<div/>', { 'id': "Chatbot" }).append(
            body_div1,
            body_div2
        )
    );

    api();
}