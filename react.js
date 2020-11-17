async function init() {
    const res = await fetch("https://aybotazurefunctionapp.azurewebsites.net/api/GenerateDirectLineToken?code=swhdQ830yeVmGlROB5K6rqQo/b32jLWT9JyPSdDmMsldGiLuyV0vkQ==", { method: "post" });
    const json = await res.json();
    const token = json.token.token;
    console.log(token);

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
    const styleOptions = {
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


    };

    // After generated, you can modify the CSS rules
    styleOptions.textContent = {
        ...styleOptions.textContent,
        fontFamily: custom.chatbot.fontFamily,
        fontWeight: custom.chatbot.fontWeight
    };



    const {
        hooks: { usePostActivity },
        ReactWebChat
    } = window.WebChat;

    const { useCallback } = window.React;

    // const BotActivityDecorator = ({ activityID, children }) => {
    //     const postActivity = usePostActivity();

    //     const handleDownvoteButton = useCallback(() => {
    //         postActivity({
    //             type: 'messageReaction',
    //             reactionsAdded: [{ activityID, helpful: -1 }]
    //         });
    //     }, [activityID, postActivity]);

    //     const handleUpvoteButton = useCallback(() => {
    //         postActivity({
    //             type: 'messageReaction',
    //             reactionsAdded: [{ activityID, helpful: 1 }]
    //         });
    //     }, [activityID, postActivity]);

    //     return (
    //         <div className="botActivityDecorator">

    //             <div className="botActivityDecorator__buttonBar">
    //                 <div>
    //                     <button className="botActivityDecorator__button" onClick={handleUpvoteButton}>
    //                         <i class="far fa-thumbs-up"></i>
    //                     </button>
    //                 </div>
    //                 <div>
    //                     <button className="botActivityDecorator__button" onClick={handleDownvoteButton}>
    //                         <i class="far fa-thumbs-down"></i>
    //                     </button>
    //                 </div>
    //             </div>
    //         </div>
    //     );
    // };

    const store = window.WebChat.createStore({}, ({ dispatch }) => next => action => {
        return next(action);
    });

    const BotActivityDecorator = ({ activityID, children }) => {
        const [{ feedbackGiven, activateButtons, selectedButton }, setFeedbackState] =
            React.useState({ feedbackGiven: false, activateButtons: false, selectedButton: '' });


        const upButtonClass = selectedButton === 'UP' ?
            'botActivityDecorator__button selected' : 'botActivityDecorator__button';
        const downButtonClass = selectedButton === 'DOWN' ?
            'botActivityDecorator__button selected' : 'botActivityDecorator__button';


        const postActivity = usePostActivity();
        const handleDownvoteButton = useCallback(() => {
            postActivity({
                type: "DIRECT_LINE/POST_ACTIVITY",
                meta: "keyboard",
                payload: {
                    activity: {
                        name: "reactionAdded",
                        type: "messageReaction",
                        value: { activityID },
                        reactionsAdded: [{ activityID, helpful: -1 }],
                        replyToID: activityID
                    }
                }
            });
        }, [activityID, postActivity]);

        const handleUpvoteButton = useCallback(() => {
            postActivity({
                type: "DIRECT_LINE/POST_ACTIVITY",
                meta: "keyboard",
                payload: {
                    activity: {
                        name: "reactionAdded",
                        type: "messageReaction",
                        value: { activityID },
                        reactionsAdded: [{ activityID, helpful: -1 }],
                        replyToID: activityID
                    }
                }
            });
        }, [activityID, postActivity]);

        return <div className="botActivityDecorator" onMouseEnter={() =>
            setFeedbackState({ feedbackGiven, activateButtons: true, selectedButton })}
            onMouseLeave={() => setFeedbackState({ feedbackGiven, activateButtons: false, selectedButton })}>
            {children}

            {activateButtons && <div className="botActivityDecorator__buttonBar">
                <button className={upButtonClass + ' up'}
                    disabled={feedbackGiven}
                    onClick={() => {
                        setFeedbackState({ feedbackGiven: true, activateButtons, selectedButton: 'UP' });
                        handleDownvoteButton();

                    }}><span className="far fa-thumbs-up"></span></button>
                <button className={downButtonClass + ' down'}
                    disabled={feedbackGiven}

                    onClick={() => {
                        setFeedbackState({ feedbackGiven: true, activateButtons, selectedButton: 'DOWN' });
                        handleUpvoteButton();

                    }}><span className="far fa-thumbs-down"></span></button>
            </div>}
        </div>;
    }

    const activityMiddleware = () => next => (...setUpArgs) => {
        const [card] = setUpArgs;

        if (card.activity.from.role === 'bot' && !card.activity.HideFeedbackButtons) {
            return (...renderArgs) => (
                <BotActivityDecorator key={card.activity.id} activityID={card.activity.id}>
                    {next(card)(...renderArgs)}
                </BotActivityDecorator>
            );
        }

        return next(card);
    };
    window.ReactDOM.render(
        <ReactWebChat
            styleOptions={styleOptions}
            activityMiddleware={activityMiddleware}

            directLine={window.WebChat.createDirectLine({ token })} />,



        document.getElementById('webchat')
    );

    document.querySelector('#webchat > *').focus();



}

// class BotContainer extends React.Component<IBotProps, {}>{
//     constructor(props) {
//         super(props);
//     }

//     async get_token() {
//         const res = await fetch("https://aybotazurefunctionapp.azurewebsites.net/api/GenerateDirectLineToken?code=swhdQ830yeVmGlROB5K6rqQo/b32jLWT9JyPSdDmMsldGiLuyV0vkQ==", { method: "post" });
//         const json = await res.json();
//         const token = json.token.token;
//         console.log(token);
//         return token;
//     }

//     BotActivityDecorator = ({ activityID, children }) => {
//         const postActivity = usePostActivity();

//         const handleDownvoteButton = useCallback(() => {
//             postActivity({
//                 type: 'messageReaction',
//                 reactionsAdded: [{ activityID, helpful: -1 }]
//             });
//         }, [activityID, postActivity]);

//         const handleUpvoteButton = useCallback(() => {
//             postActivity({
//                 type: 'messageReaction',
//                 reactionsAdded: [{ activityID, helpful: 1 }]
//             });
//         }, [activityID, postActivity]);

//         return (
//             <div className="botActivityDecorator">
//                 <ul className="botActivityDecorator__buttonBar">
//                     <li>
//                         <button className="botActivityDecorator__button" onClick={handleUpvoteButton}>
//                             <i class="far fa-thumbs-up"></i>
//                         </button>
//                     </li>
//                     <li>
//                         <button className="botActivityDecorator__button" onClick={handleDownvoteButton}>
//                             <i class="far fa-thumbs-down"></i>
//                         </button>
//                     </li>
//                 </ul>
//                 <div className="botActivityDecorator__content">{children}</div>
//             </div>
//         );
//     };

//     activityMiddleware = () => next => card => {
//         if (card.activity.from.role === 'bot') {
//             return children => (
//                 <BotActivityDecorator key={card.activity.id} activityID={card.activity.id}>
//                     {next(card)(children)}
//                 </BotActivityDecorator>
//             );
//         }

//         return next(card);
//     };
//     async render() {
//         const custom = {
//             wrapper: {
//                 color: "#0769AD",
//                 bot_name: "Mary",
//                 bot_avatar: "https://agreeyabotstorage.blob.core.windows.net/chat-wrapper/Bot Logo 2.png"

//             },
//             chatbot: {
//                 bot_avatar_image: "https://agreeyabotstorage.blob.core.windows.net/chat-wrapper/Bot Logo 2.png",
//                 user_avatar_image: 'https://upload.wikimedia.org/wikipedia/commons/thumb/a/a5/Instagram_icon.png/900px-Instagram_icon.png',
//                 user_avatar_intials: "Me",
//                 bot_avatar_initails: "Bot",
//                 bot_bubble_color: 'rgba(0, 0, 255, .1)',
//                 user_bubble_color: 'rgba(0, 255, 0, .1)',
//                 fontFamily: "'Roboto', sans-serif",
//                 fontWeight: '400'


//             }

//         }
//         const styleOptions = {
//             alignItems: "left",
//             justifyContent: "left",
//             bubbleBackground: custom.chatbot.bot_bubble_color,
//             bubbleFromUserBackground: custom.chatbot.user_bubble_color,
//             botAvatarImage: custom.chatbot.bot_avatar_image,
//             botAvatarInitials: custom.chatbot.bot_avatar_initails,
//             //userAvatarImage: custom.chatbot.user_avatar,
//             userAvatarInitials: custom.chatbot.user_avatar_intials,
//             rootHeight: '99%',
//             rootWidth: '100%',
//             bubbleFromUserBorderRadius: 5,
//             bubbleBorderRadius: 5,


//         };

//         // After generated, you can modify the CSS rules
//         styleOptions.textContent = {
//             ...styleOptions.textContent,
//             fontFamily: custom.chatbot.fontFamily,
//             fontWeight: custom.chatbot.fontWeight
//         };

//         const token = await this.get_token();

//         return (
//             <ReactWebChat activityMiddleware={activityMiddleware}
//                 styleOptions={styleOptions}
//                 directLine={window.WebChat.createDirectLine({ token })} />

//         );

//     }


// }