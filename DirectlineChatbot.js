//Customiztion


const custom = {
    wrapper: {
        color: "#0769AD",
        bot_name: "Mary",
        bot_avatar: "https://agreeyabotstorage.blob.core.windows.net/chat-wrapper/Bot Logo 2.png"

    },
    chatbot: {
        bot_avatar_image: "https://agreeyabotstorage.blob.core.windows.net/chat-wrapper/Bot Logo 2.png",
        user_avatar_image: 'https://upload.wikimedia.org/wikipedia/commons/thumb/a/a5/Instagram_icon.png/900px-Instagram_icon.png',
        user_avatar_intials: "Me",
        bot_avatar_initails: "Bot",
        bot_bubble_color: 'rgba(0, 0, 255, .1)',
        user_bubble_color: 'rgba(0, 255, 0, .1)',
        fontFamily: "'Roboto', sans-serif",
        fontWeight: '400'


    }

}









async function api() {
    //console.log("stated")
    //let response = await fetch("https://aybotazurefunctionapp.azurewebsites.net/api/GenerateAYBotToken?code=DjpE11lNWCsOaKN1OZIUjOHhAPb/4ze6JdCNqGS2Qcddp5VSLeESdQ==");
    //let token = await response.json();
    //console.log(token)
    //document.getElementById("ay_bot").src = token.source

    const styleSet = window.WebChat.createStyleSet({
        alignItems: "left",
        justifyContent: "left",
        bubbleBackground: custom.chatbot.bot_bubble_color,
        bubbleFromUserBackground: custom.chatbot.user_bubble_color,
        botAvatarImage: custom.chatbot.bot_avatar_image,
        botAvatarInitials: custom.chatbot.bot_avatar_initails,
        //userAvatarImage: custom.chatbot.user_avatar,
        userAvatarInitials: custom.chatbot.user_avatar_intials,
        rootHeight: '99%',
        rootWidth: '100%',
        bubbleFromUserBorderRadius: 5,
        bubbleBorderRadius: 5,


    });

    // After generated, you can modify the CSS rules
    styleSet.textContent = {
        ...styleSet.textContent,
        fontFamily: custom.chatbot.fontFamily,
        fontWeight: custom.chatbot.fontWeight
    };

    //const res = await fetch('https:<YOUR_TOKEN_SERVER/API>', { method: 'POST' });
    //const { token } = await res.json();
    const res = await fetch("https://aybotazurefunctionapp.azurewebsites.net/api/GenerateDirectLineToken?code=swhdQ830yeVmGlROB5K6rqQo/b32jLWT9JyPSdDmMsldGiLuyV0vkQ==", { method: "post" });
    const json = await res.json();
    const token = json.token.token;
    console.log(token);

    const selectedLang = $("#lang option:selected").val();
    console.log(selectedLang);





    window.WebChat.renderWebChat({
            directLine: window.WebChat.createDirectLine({ token }),
            locale: selectedLang,
            // Passing 'styleSet' when rendering Web Chat
            styleSet,
            webSpeechPonyfillFactory: window.WebChat.createBrowserWebSpeechPonyfillFactory(),

        },
        document.getElementById('ay_bot')
    );
}

function toggleChatbox(slidedirection) {
    $("#ay_ChatBotBox").toggle();
}

function switchChatbox() {
    $('.setAlignment').toggleClass('ay_leftAlign').toggleClass('ay_rightAlign');
}

function init() {
    build();
    api();
}

function build() {
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
            $('<link/>', { 'rel': "stylesheet", 'href': "https://agreeyabotstorage.blob.core.windows.net/chat-wrapper/stylesheet.css" }),
            //webchat cnd
            $('<script/>', { 'src': "https://cdn.botframework.com/botframework-webchat/latest/webchat.js" })

        ]
    )



    var header_div1 =
        $('<div/>', { "class": "ay_chatbotImgContainer" })
        .append(
            [
                $('<img/>', { 'src': custom.wrapper.bot_avatar, 'class': 'ay_chatbotimg' }),
                $('<span/>', { 'class': 'ay_botName', 'text': custom.wrapper.bot_name })
            ]
        )

    var dropdown = $('<span/>', { 'class': "ay_languageDDLContainer" }).append(
        $('<select/>', { 'id': 'lang', 'class': 'ay_languageDDL' }).append(
            [
                $('<option/>', { 'value': 'en', 'text': "English", 'class': "options" }),
                $('<option/>', { 'value': 'fr', 'text': "French", 'class': "options" }),
                $('<option/>', { 'value': 'de', 'text': "German", 'class': "options" }),

            ]
        )

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
        $('<div/>', { 'class': 'ay_headerRightcontent' }).append(dropdown, btn1, btn2, btn3)

    var header =
        $('<div/>', { "class": "ay_chatboxheaderContainer", 'style': `background-color:${custom.wrapper.color}` })
        .append(header_div1, header_div2);

    var body_div2 =
        $('<div/>', { 'class': 'ay_chatboticon ay_rightAlign setAlignment', 'style': `background-color:${custom.wrapper.color}` }).append(
            $('<a/>', { 'href': "javascript:void(0)", 'onClick': "toggleChatbox()" }).append(
                $('<span/>', { 'class': "ms-Icon ms-Icon--Message" })
            )
        )

    var chatbox = $('<div/>', { 'class': 'ay_chatbox', }).append(
        $("<div/>", { 'id': 'ay_bot' })
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


    $("#lang").change(renderWebChat);

}

function renderWebChat() {


    api();

}