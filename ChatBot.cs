using Agreeya.ChatBot.Dialogs.Main;
using Agreeya.ChatBot.Dialogs.Ticket;
using Agreeya.ChatBot.Models;
using Agreeya.ChatBot.Translation;
using Agreeya.ChatBot.Trial;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Connector;
using Microsoft.Bot.Connector.Authentication;
using Microsoft.Bot.Schema;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using TelemetryConstants = Agreeya.ChatBot.Middleware.Telemetry.TelemetryConstants;

namespace Agreeya.ChatBot
{
    public class ChatBot : ActivityHandler
    {
        private readonly BotServices _services;
        private readonly ConversationState _conversationState;
        private readonly UserState _userState;
        private readonly MicrosoftTranslator _translator;
        private readonly IBotTelemetryClient _telemetryClient;
        private DialogSet _dialogs;
        private readonly IConfiguration _configuration;
        private readonly ICosmosDbService _cosmosDbService;
        private MainResponses _responder = new MainResponses();
        private ActivityLog log;

        /// <summary>
        /// Initializes a new instance of the <see cref="EnterpriseBotLab"/> class.
        /// </summary>
        /// <param name="botServices">Bot services.</param>
        /// <param name="conversationState">Bot conversation state.</param>
        /// <param name="userState">Bot user state.</param>
        public ChatBot(BotServices botServices, ConversationState conversationState, UserState userState, IBotTelemetryClient telemetryClient, IConfiguration configuration, MicrosoftTranslator translator, ICosmosDbService cosmosDbService,
            ActivityLog log)
        {
            _conversationState = conversationState ?? throw new ArgumentNullException(nameof(conversationState));
            _userState = userState ?? throw new ArgumentNullException(nameof(userState));
            _services = botServices ?? throw new ArgumentNullException(nameof(botServices));
            _telemetryClient = telemetryClient ?? throw new ArgumentNullException(nameof(telemetryClient));
            _configuration = configuration ?? throw new ArgumentNullException(nameof(configuration));
            _translator = translator ?? throw new ArgumentNullException(nameof(translator));
            _dialogs = new DialogSet(_conversationState.CreateProperty<DialogState>(nameof(ChatBot)));
            _cosmosDbService = cosmosDbService ?? throw new ArgumentNullException(nameof(cosmosDbService));
            _dialogs.Add(new MainDialog(_services, _conversationState, _userState, _telemetryClient, _configuration, _translator, _cosmosDbService));
            _dialogs.Add(new AskForTicket(_services, _conversationState, _userState, _telemetryClient, _configuration, _translator));
            this.log = log;
        }

        /// <summary>
        /// Run every turn of the conversation. Handles orchestration of messages.
        /// </summary>
        /// <param name="turnContext">Bot Turn Context.</param>
        /// <param name="cancellationToken">Task CancellationToken.</param>
        /// <returns>A <see cref="Task"/> representing the asynchronous operation.</returns>
        public override async Task OnTurnAsync(ITurnContext turnContext, CancellationToken cancellationToken)
        {
            // Client notifying this bot took to long to respond (timed out)
            if (turnContext.Activity.Code == EndOfConversationCodes.BotTimedOut)
            {
                _services.TelemetryClient.TrackTrace($"Timeout in {turnContext.Activity.ChannelId} channel: Bot took too long to respond.");
                return;
            }

            var dc = await _dialogs.CreateContextAsync(turnContext);
            if (turnContext.Activity.ChannelId.ToLower().Contains("teams"))
            {
                MicrosoftAppCredentials.TrustServiceUrl(turnContext.Activity.ServiceUrl);
                var members = (await turnContext.TurnState.Get<IConnectorClient>().Conversations.GetConversationMembersAsync(turnContext.Activity.Conversation.Id).ConfigureAwait(false)).ToList();
                foreach (var member in members)
                {
                    MicrosoftTeamUser teamMember = JsonConvert.DeserializeObject<MicrosoftTeamUser>(member.Properties.ToString());
                    if (!string.IsNullOrEmpty(teamMember.Email))
                    {
                        turnContext.Activity.From.Id = teamMember.Email;
                    }
                }
            }
            if (turnContext.Activity.Type == ActivityTypes.Message && Constants.ProjectFeatures.EnableTrialCheck)
            {
                var conversationStateAccessors = _conversationState.CreateProperty<ConversationData>(nameof(ConversationData));
                var conversationData = await conversationStateAccessors.GetAsync(turnContext, () => new ConversationData());
                if (!conversationData.IsTrialExpired)
                {
                    string customernamekey = ClientKeyDecrypt.DecryptString(_configuration.GetSection("customerNameKey").Get<string>());
                    var clientTrialData = await _cosmosDbService.GetItemsAsync("SELECT * FROM c WHERE c.key='" + customernamekey + "'");
                    if (clientTrialData != null && clientTrialData.Count() > 0)
                    {
                        Item item = clientTrialData.ElementAt(0);
                        DateTime eDate = Convert.ToDateTime(item.EndDate);
                        if (eDate < DateTime.Today)
                        {
                            await _responder.ReplyWith(dc.Context, MainResponses.ResponseIds.TrialExpired);
                            conversationData.IsTrialExpired = true;
                            return;
                        }
                    }
                    else
                    {
                        await _responder.ReplyWith(dc.Context, MainResponses.ResponseIds.TrialExpired);
                        conversationData.IsTrialExpired = true;
                        return;
                    }
                }
                else
                {
                    await _responder.ReplyWith(dc.Context, MainResponses.ResponseIds.TrialExpired);
                    return;
                }
            }

            if (turnContext.Activity.Type == ActivityTypes.MessageReaction)
            {
                await base.OnTurnAsync(turnContext, cancellationToken);
            }
            else
            {
                if (dc.ActiveDialog != null)
                {
                    var result = await dc.ContinueDialogAsync();
                }
                else
                {
                    await dc.BeginDialogAsync(nameof(MainDialog));
                }
            }
        }

        protected override async Task OnReactionsAddedAsync(IList<MessageReaction> messageReactions, ITurnContext<IMessageReactionActivity> turnContext, CancellationToken cancellationToken)
        {
            try
            {
                // The ReplyToId property of the inbound MessageReaction Activity will correspond to a Message Activity which
                // was previously sent from this bot.
                var activity = await log.Find(turnContext.Activity.ReplyToId);
                if (activity == null)
                {
                    // If we had sent the message from the error handler we wouldn't have recorded the Activity Id and so we
                    // shouldn't expect to see it in the log.
                    await turnContext.SendActivityAsync($"Activity {turnContext.Activity.ReplyToId} not found in the log.");
                }
                else
                {
                    _telemetryClient.TrackEvent(TelemetryConstants.UserFeedbackVote, new Dictionary<string, string> {
                    { "Answer", getTextFromActivity(activity) },
                    { TelemetryConstants.UserFeedbackVote, Configuration.Messages("Feedback", "Yes") },
                    { TelemetryConstants.ChannelIdProperty, activity.ChannelId },
                    { TelemetryConstants.FromIdProperty, activity.From.Id },
                    { TelemetryConstants.ConversationNameProperty, activity.Conversation.Name },
                    { TelemetryConstants.LocaleProperty, activity.Locale },
                    { TelemetryConstants.RecipientIdProperty, activity.Recipient.Id },
                    { TelemetryConstants.RecipientNameProperty, activity.Recipient.Name }
                });
                }

                await turnContext.SendActivityWithoutFeedback(MessageFactory.Text(Configuration.Messages("Feedback", "ThumbsUp")));
            }
            catch (Exception)
            {
                await turnContext.SendActivityWithoutFeedback(MessageFactory.Text(Configuration.Messages("UnknownError")));
                throw;
            }
        }

        protected override async Task OnReactionsRemovedAsync(IList<MessageReaction> messageReactions, ITurnContext<IMessageReactionActivity> turnContext, CancellationToken cancellationToken)
        {
            try
            {
                // The ReplyToId property of the inbound MessageReaction Activity will correspond to a Message Activity which
                // was previously sent from this bot.
                var activity = await log.Find(turnContext.Activity.ReplyToId);
                if (activity == null)
                {
                    // If we had sent the message from the error handler we wouldn't have recorded the Activity Id and so we
                    // shouldn't expect to see it in the log.
                    await turnContext.SendActivityAsync($"Activity {turnContext.Activity.ReplyToId} not found in the log.");
                }
                else
                {
                    _telemetryClient.TrackEvent(TelemetryConstants.UserFeedbackVote, new Dictionary<string, string> {
                    { "Answer", getTextFromActivity(activity) },
                    { TelemetryConstants.UserFeedbackVote, Configuration.Messages("Feedback", "No") },
                    { TelemetryConstants.ChannelIdProperty, activity.ChannelId },
                    { TelemetryConstants.FromIdProperty, activity.From.Id },
                    { TelemetryConstants.ConversationNameProperty, activity.Conversation.Name },
                    { TelemetryConstants.LocaleProperty, activity.Locale },
                    { TelemetryConstants.RecipientIdProperty, activity.Recipient.Id },
                    { TelemetryConstants.RecipientNameProperty, activity.Recipient.Name }
                });
                }

                //await turnContext.SendActivityWithoutFeedback(MessageFactory.Text(Configuration.Messages("Feedback", "ThumbsDown")));

                var dc = await _dialogs.CreateContextAsync(turnContext);
                await dc.BeginDialogAsync(nameof(AskForTicket), Configuration.Messages("Feedback", "ThumbsDown"));
            }
            catch (Exception)
            {
                await turnContext.SendActivityWithoutFeedback(MessageFactory.Text(Configuration.Messages("UnknownError")));
                throw;
            }
        }

        private string getTextFromActivity(Activity activity)
        {
            List<string> textList = new List<string>();

            if (!string.IsNullOrEmpty(activity.Text)) textList.Add(activity.Text);

            activity.Attachments?.All(delegate (Attachment attachment)
            {
                if (!string.IsNullOrEmpty(attachment.Content as string)) textList.Add(attachment.Content as string);
                else if (null != (attachment.Content as HeroCard)) textList.Add((attachment.Content as HeroCard).Text);
                else if (null != (attachment.Content as ThumbnailCard)) textList.Add((attachment.Content as ThumbnailCard).Text);

                return false;
            });

            return string.Join(Environment.NewLine, textList);
        }
    }
}