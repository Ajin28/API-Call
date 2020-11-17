import * as React from "react";
import {
  SPHttpClient,
  SPHttpClientResponse,
  SPHttpClientConfiguration
} from "@microsoft/sp-http";
import "./aybot.scss";
import { IBotProps, MenuItem } from "./IBotProps";
import ChatBotDialog from "./ChatBotDialog";
import cookie from 'react-cookies';
import SPUtility from "./SPUtility";
import {
  createStore,
  createStyleSet,
  createCognitiveServicesSpeechServicesPonyfillFactory,
  createDirectLine,
  Components
} from "botframework-webchat";
import ReactWebChat from "botframework-webchat";
import * as moment from "moment";
import fetchSpeechServicesToken from "./fetchSpeechServicesToken";

import { PrimaryButton, Button } from 'office-ui-fabric-react/lib/Button';
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';

import * as SUI from 'simple-update-in';

import { AppConfigurationClient } from '@azure/app-configuration';

const store = createStore({}, ({ dispatch }) => next => action => {
  if (action.type === "DIRECT_LINE/POST_ACTIVITY") {
  } else if (action.type === "DIRECT_LINE/CONNECT_FULFILLED") {
    dispatch({
      type: "WEB_CHAT/SEND_EVENT",
      payload: {
        name: "webchat/join",
        value: { language: window.navigator.language }
      }
    });
  } else if (action.type === "DIRECT_LINE/INCOMING_ACTIVITY") {
    const { activity } = action.payload;
    if (activity.type === "message") {
    }
  }
  return next(action);
});

export default class BotContainer extends React.Component<IBotProps, {}> {
  private LanguageOptions: any;
  public directLine: any;
  public webSpeechPonyfillFactory: any;
  public state = {
    minimized: true,
    newMessage: false,
    chatDisplay: "none",
    iconDisplay: "block",
    side: "right",
    styleSet: "",
    modalIsOpen: true,
    afterOpenModal: true,
    closeModal: false,
    dialogStatus: true,
    _token: null,
    localeLang: "en-US",
    photoUrl: this.props.photoUrl,
    LastConversation: null,
    chatBotDialogProps: this.props.chatBotDialogProps,
    botConnected: false
  };
  public user = {
    id: this.props.context.pageContext.user.email,
    name: this.props.context.pageContext.user.displayName
  };
  public BotName: string = "";
  public Boticon: string = "";
  public FeedbackTimeout: string = "";
  public TokenRegion: string = "";
  public LangauageAvailable: any;
  public BotSecret: string = "";
  public ShowHideAttachment: string = "";
  public SpeechToken: any = "";
  public chatBotDialogProps: any = "";
  public BotAdminGroup: string = "BOTAdmin";
  public IsBotAdmin: boolean;
  public PickedColor: string;
  public Token: string;
  public store: any;
  public LastConversationDuration: any;
  public isFeedBackSent: boolean;
  private appConfigClient: AppConfigurationClient;
  private configConnString: string;
  private popupWelcomeMessage: string;
  private welcomeHeading: string;
  private welcomeMessage: string;
  private helpMessage: string;
  private menuItemTicket: MenuItem = { value: '', enabled: false };
  private menuItemSearch: MenuItem = { value: '', enabled: false };
  private menuItemTeams: MenuItem = { value: '', enabled: false };
  private menuItemMyActivities: MenuItem = { value: '', enabled: false };
  private menuItemSelfService: MenuItem = { value: '', enabled: false };
  private thumbsUpTooltip: string;
  private thumbsDownTooltip: string;
  private dialog: ChatBotDialog;

  constructor(props) {
    super(props);
    this.handleFetchToken = this.handleFetchToken.bind(this);
    this.handleMaximizeButtonClick = this.handleMaximizeButtonClick.bind(this);
    this.handleMinimizeButtonClick = this.handleMinimizeButtonClick.bind(this);
    this.handleSwitchButtonClick = this.handleSwitchButtonClick.bind(this);
    this.handleTextChangedEvent = this.handleTextChangedEvent.bind(this);
    this.helpButtonClick = this.helpButtonClick.bind(this);
    this.closeDialog = this.closeDialog.bind(this);
    this.openModal = this.openModal.bind(this);
    this.afterOpenModal = this.afterOpenModal.bind(this);
    this.closeModalYes = this.closeModalYes.bind(this);
    this.closeModalNo = this.closeModalNo.bind(this);
    this.BotName =
      this.props.chatBotDialogProps.BotName == undefined
        ? "Mary"
        : this.props.chatBotDialogProps.BotName;
    this.Boticon =
      this.props.chatBotDialogProps.Boticon == undefined
        ? "https://acuvate.com/wp-content/uploads/2017/04/sia-500x500-1-362x362.png"
        : this.props.chatBotDialogProps.Boticon;
    this.LangauageAvailable = this.props.chatBotDialogProps.LangauageAvailable;
    this.BotSecret = this.props.chatBotDialogProps.BotSecret;
    this.TokenRegion = this.props.chatBotDialogProps.TokenRegion;

    this.ShowHideAttachment = this.props.chatBotDialogProps.ShowHideAttachment;
    this.context = this.props.chatBotDialogProps.context;
    this.chatBotDialogProps = this.props.chatBotDialogProps;
    this.PickedColor = this.props.chatBotDialogProps.PickedColor;
    // this.FeedbackTimeout = this.props.chatBotDialogProps.FeedbackTimeout;
    // this.SpeechToken = this.props.chatBotDialogProps.SpeechToken;
    this.props.chatBotDialogProps.BotName = this.BotName;
    this.props.chatBotDialogProps.Boticon = this.Boticon;
    // this.props.chatBotDialogProps.FeedbackTimeout = this.FeedbackTimeout;
    // this.props.chatBotDialogProps.BotSecret = this.props.chatBotDialogProps.BotSecret;

    initializeIcons();
    this.getConfigValuesFromAppConfig();
  }

  private async getConfigValuesFromAppConfig() {
    try {
      this.configConnString = this.props.chatBotDialogProps.ConfigConnString;

      if (this.configConnString) {
        this.appConfigClient = new AppConfigurationClient(this.configConnString);

        // this.popupWelcomeMessage = (await this.appConfigClient.getConfigurationSetting({ key: 'Message:WebChat:PopupWelcomeMessage' })).value;
        // this.SpeechToken = (await this.appConfigClient.getConfigurationSetting({ key: 'Setting:SpeechToken' })).value;
        // this.FeedbackTimeout = (await this.appConfigClient.getConfigurationSetting({ key: 'Setting:FeedbackTimeout' })).value;
        // this.welcomeHeading = (await this.appConfigClient.getConfigurationSetting({ key: 'Message:WelcomeHeading' })).value;
        // this.welcomeMessage = (await this.appConfigClient.getConfigurationSetting({ key: 'Message:WelcomeText' })).value;
        // this.helpMessage = (await this.appConfigClient.getConfigurationSetting({ key: 'Message:HelpText' })).value;
        // this.thumbsDownTooltip = (await this.appConfigClient.getConfigurationSetting({ key: 'Message:WebChat:Tooltip:ThumbsDown' })).value;
        // this.thumbsUpTooltip = (await this.appConfigClient.getConfigurationSetting({ key: 'Message:WebChat:Tooltip:ThumbsUp' })).value;

        this.appConfigClient.getConfigurationSetting({ key: 'Message:WelcomeHeading' })
          .then(res => this.welcomeHeading = res.value);
        this.appConfigClient.getConfigurationSetting({ key: 'Message:WelcomeText' })
          .then(res => this.welcomeMessage = res.value);
        this.appConfigClient.getConfigurationSetting({ key: 'Message:HelpText' })
          .then(res => this.helpMessage = res.value);
        this.appConfigClient.getConfigurationSetting({ key: 'Message:WebChat:Tooltip:ThumbsDown' })
          .then(res => this.thumbsDownTooltip = res.value);
        this.appConfigClient.getConfigurationSetting({ key: 'Message:WebChat:Tooltip:ThumbsUp' })
          .then(res => this.thumbsUpTooltip = res.value);

        this.appConfigClient.getConfigurationSetting({ key: 'Message:WebChat:PopupWelcomeMessage' })
          .then(res => {
            this.popupWelcomeMessage = res.value;
            this.setState({});
          });
        this.appConfigClient.getConfigurationSetting({ key: 'Setting:SpeechToken' })
          .then(res => {
            this.SpeechToken = res.value;
            this.setState({});
          });
        this.appConfigClient.getConfigurationSetting({ key: 'Setting:FeedbackTimeout' })
          .then(res => {
            this.FeedbackTimeout = res.value;
            this.setState({});
          });
        this.menuItemTicket = {
          value: (await this.appConfigClient.getConfigurationSetting({ key: 'Message:MenuItems:Ticket' })).value,
          enabled: (await this.appConfigClient.getConfigurationSetting({ key: 'Setting:TicketManagement' }))
            .value.trim().toLowerCase() === '1'
        };
        this.menuItemSearch = {
          value: (await this.appConfigClient.getConfigurationSetting({ key: 'Message:MenuItems:Search' })).value,
          enabled: (await this.appConfigClient.getConfigurationSetting({ key: 'Setting:EnterpriseSearch' }))
            .value.trim().toLowerCase() === '1'
        };
        this.menuItemTeams = {
          value: (await this.appConfigClient.getConfigurationSetting({ key: 'Message:MenuItems:Teams' })).value,
          enabled: (await this.appConfigClient.getConfigurationSetting({ key: 'Setting:TeamsManagement' }))
            .value.trim().toLowerCase() === '1'
        };
        this.menuItemMyActivities = {
          value: (await this.appConfigClient.getConfigurationSetting({ key: 'Message:MenuItems:Scheduler' })).value,
          enabled: (await this.appConfigClient.getConfigurationSetting({ key: 'Setting:TaskManagement' }))
            .value.trim().toLowerCase() === '1'
        };
        this.menuItemSelfService = {
          value: (await this.appConfigClient.getConfigurationSetting({ key: 'Message:MenuItems:Self-Service' })).value,
          enabled: (await this.appConfigClient.getConfigurationSetting({ key: 'Setting:SelfService' }))
            .value.trim().toLowerCase() === '1'
        };
        this.setState({});
        // if (this.dialog) this.dialog.update();
      }

      this.checkIfUserIsBotAdmin();
    }
    catch (ex) {
      console.log(ex);
    }
  }

  private clicked() {
    const dialog = this.dialog = new ChatBotDialog(
      this.props.context,
      this.props.domElement,
      this.chatBotDialogProps,
      {
        welcomeHeading: this.welcomeHeading,
        welcomeMessage: this.welcomeMessage,
        helpMessage: this.helpMessage,
        menuItemTicket: this.menuItemTicket,
        menuItemSearch: this.menuItemSearch,
        menuItemTeams: this.menuItemTeams,
        menuItemMyActivities: this.menuItemMyActivities,
        menuItemSelfService: this.menuItemSelfService,
      }
    );
    if (this.BotName) {
      dialog.BotName = this.BotName;
    }
    if (this.LangauageAvailable) {
      dialog.LangauageAvailable = this.LangauageAvailable;
    }
    if (this.BotSecret) {
      dialog.BotSecret = this.BotSecret;
    }
    if (this.Boticon) {
      dialog.Boticon = this.Boticon;
    }
    if (this.FeedbackTimeout) {
      dialog.FeedbackTimeout = this.FeedbackTimeout;
    }
    if (this.TokenRegion) {
      dialog.TokenRegion = this.TokenRegion;
    }
    if (this.ShowHideAttachment) {
      dialog.ShowHideAttachment = this.ShowHideAttachment;
    }
    if (this.SpeechToken) {
      dialog.SpeechToken = this.SpeechToken;
    }
    if (this.PickedColor) {
      dialog.PickedColor = this.PickedColor;
    }
    dialog.show();
    dialog.onSubmit = () => {
      if (!dialog.Submitted) return;

      this.BotName = dialog.BotName;
      this.Boticon = dialog.Boticon;
      this.FeedbackTimeout = dialog.FeedbackTimeout;
      this.TokenRegion = dialog.TokenRegion;
      this.LangauageAvailable = dialog.LangauageAvailable;
      this.BotSecret = dialog.BotSecret;
      this.ShowHideAttachment = dialog.ShowHideAttachment;
      this.SpeechToken = dialog.SpeechToken;
      this.PickedColor = dialog.PickedColor;
      this.configConnString = dialog.configConnString;
      this.welcomeHeading = dialog.messages.welcomeHeading;
      this.welcomeMessage = dialog.messages.welcomeMessage;
      this.helpMessage = dialog.messages.helpMessage;
      this.menuItemTicket = dialog.messages.menuItemTicket;
      this.menuItemSearch = dialog.messages.menuItemSearch;
      this.menuItemTeams = dialog.messages.menuItemTeams;
      this.menuItemMyActivities = dialog.messages.menuItemMyActivities;
      this.menuItemSelfService = dialog.messages.menuItemSelfService;

      // save the  termSet in properties
      SPUtility.loadSPDependencies().then(() => {
        this.props.chatBotDialogProps.BotName = this.BotName;
        this.props.chatBotDialogProps.Boticon = this.Boticon;
        this.props.chatBotDialogProps.FeedbackTimeout = this.FeedbackTimeout;
        this.props.chatBotDialogProps.TokenRegion = this.TokenRegion;
        this.props.chatBotDialogProps.LangauageAvailable = this.LangauageAvailable;
        this.props.chatBotDialogProps.BotSecret = this.BotSecret;
        this.props.chatBotDialogProps.ShowHideAttachment = this.ShowHideAttachment;
        this.props.chatBotDialogProps.SpeechToken = this.SpeechToken;
        this.props.chatBotDialogProps.PickedColor = this.PickedColor;
        this.props.chatBotDialogProps.ConfigConnString = this.configConnString;
        this.saveProperties();
      });
      this.saveAppConfig().then(() => {
        this.restartBotService();
        setTimeout(() => { dialog.close(); }, 2000);
      });
    };
  }

  private async saveAppConfig() {
    try {
      if (!this.appConfigClient) {
        this.appConfigClient = new AppConfigurationClient(this.configConnString);
      }
      this.appConfigClient.setConfigurationSetting({ key: 'Message:WelcomeHeading', value: this.welcomeHeading });
      this.appConfigClient.setConfigurationSetting({ key: 'Message:WelcomeText', value: this.welcomeMessage });
      this.appConfigClient.setConfigurationSetting({ key: 'Message:HelpText', value: this.helpMessage });
      this.appConfigClient.setConfigurationSetting({ key: 'Message:MenuItems:Ticket', value: this.menuItemTicket.value });
      this.appConfigClient.setConfigurationSetting({ key: 'Message:MenuItems:Search', value: this.menuItemSearch.value });
      this.appConfigClient.setConfigurationSetting({ key: 'Message:MenuItems:Teams', value: this.menuItemTeams.value });
      this.appConfigClient.setConfigurationSetting({ key: 'Message:MenuItems:Scheduler', value: this.menuItemMyActivities.value });
      this.appConfigClient.setConfigurationSetting({ key: 'Message:MenuItems:Self-Service', value: this.menuItemSelfService.value });

      this.appConfigClient.setConfigurationSetting({ key: 'Setting:BotName', value: this.BotName });
      this.appConfigClient.setConfigurationSetting({ key: 'Setting:BotIcon', value: this.Boticon });
      await this.appConfigClient.setConfigurationSetting({ key: 'Setting:FeedbackTimeout', value: this.FeedbackTimeout });
      await this.appConfigClient.setConfigurationSetting({ key: 'Setting:TokenRegion', value: this.TokenRegion });
      await this.appConfigClient.setConfigurationSetting({ key: 'Setting:BotSecret', value: this.BotSecret });
      await this.appConfigClient.setConfigurationSetting({ key: 'Setting:SpeechToken', value: this.SpeechToken });
      await this.appConfigClient.setConfigurationSetting({ key: 'Setting:PickedColor', value: this.PickedColor });
      await this.appConfigClient.setConfigurationSetting({ key: 'Setting:LanguageAvailable', value: JSON.stringify(this.LangauageAvailable) });
      await this.appConfigClient.setConfigurationSetting({ key: 'Setting:TicketManagement', value: String(Number(this.menuItemTicket.enabled)) });
      await this.appConfigClient.setConfigurationSetting({ key: 'Setting:EnterpriseSearch', value: String(Number(this.menuItemSearch.enabled)) });
      await this.appConfigClient.setConfigurationSetting({ key: 'Setting:TeamsManagement', value: String(Number(this.menuItemTeams.enabled)) });
      await this.appConfigClient.setConfigurationSetting({ key: 'Setting:TaskManagement', value: String(Number(this.menuItemMyActivities.enabled)) });
      await this.appConfigClient.setConfigurationSetting({ key: 'Setting:SelfService', value: String(Number(this.menuItemSelfService.enabled)) });
    }
    catch (ex) {
      console.log('Exception:', ex);
    }
  }

  private saveProperties() {
    this.context = new SP.ClientContext(
      this.props.context.pageContext.web.absoluteUrl
    );
    let oWebsite = this.context.get_web();
    let collUserCustomAction = oWebsite.get_userCustomActions();
    this.context.load(collUserCustomAction);
    this.context.executeQueryAsync(
      (sender, args) => {
        let customActionEnumerator = collUserCustomAction.getEnumerator();

        while (customActionEnumerator.moveNext()) {
          let oUserCustomAction = customActionEnumerator.get_current();
          if (oUserCustomAction.get_title() === "Aybot") {
            oUserCustomAction.set_clientSideComponentProperties(
              JSON.stringify(this.props.chatBotDialogProps)
            );

            oUserCustomAction.update();
            this.context.load(oUserCustomAction);

            this.context.executeQueryAsync(
              (s, arg) => { },
              (s, arg) => {
                console.log(
                  "saveProperties: ",
                  arg.get_message() + "\n" + arg.get_stackTrace()
                );
              }
            );

            break;
          }
        }

        // update component
        let _language: any = this.props.chatBotDialogProps.LangauageAvailable;
        if (_language != null && _language.length > 0) {
          let _key = _language[0].split(":")[0];
          this.setState({ localeLang: _key });
        }
        this.setState({});
      },
      (sender, args) => {
        console.log(
          "saveProperties: ",
          args.get_message() + "\n" + args.get_stackTrace()
        );
      }
    );
  }

  private handleTextChangedEvent(event) {
    this.setState({ localeLang: event.target.value });
  }
  private async handleFetchToken() { }
  private handleMaximizeButtonClick() {
    this.setState(() => ({
      minimized: false,
      newMessage: false,
      chatDisplay: "block",
      iconDisplay: "none",
      modalIsOpen: false
    }));
  }
  private helpButtonClick() {
    this.store.dispatch({
      type: "WEB_CHAT/SEND_MESSAGE",
      payload: { text: "help" }
    });
  }
  private handleMinimizeButtonClick() {
    this.setState(() => ({
      minimized: true,
      newMessage: false,
      chatDisplay: "none",
      iconDisplay: "block"
    }));
  }

  private handleSwitchButtonClick() {
    let side = this.state.side;
    this.setState(({ }) => ({
      side: side === "left" ? "right" : "left"
    }));
  }
  private closeDialog() {
    this.setState(({ }) => ({
      dialogStatus: false
    }));
  }
  private checkIfEventNeedsToPosted = lastConversationTime => {
    let currentDate = moment(new Date());
    let timeEnd = moment(lastConversationTime);
    let diff = currentDate.diff(timeEnd);
    let diffDuration = moment.duration(diff).asMinutes();
    let feedbackTimeout: number;
    feedbackTimeout = parseInt(this.FeedbackTimeout, 10);

    if (diffDuration > feedbackTimeout) {
      console.log(
        "Bot will send the event here: " +
        diffDuration +
        " & Feedback Timeout: " +
        feedbackTimeout
      );
      this.PostEventToBot();
      clearInterval(this.LastConversationDuration);
    }
  }

  private PostEventToBot() {
    this.store.dispatch({
      type: "WEB_CHAT/SEND_EVENT",
      payload: {
        name: "displayFeedbackDialog",
        text: "feedback",
        locale: "en-US"
      }
    });
    this.isFeedBackSent = true;
  }

  private restartBotService() {
    this.store.dispatch({
      type: "WEB_CHAT/SEND_EVENT",
      payload: {
        name: "restartBotService"
      }
    });
  }

  public async componentDidMount() {
    this.store = createStore(
      {},
      ({ dispatch }) => next => action => {
        if (action.type === "DIRECT_LINE/CONNECT_FULFILLED") {
          dispatch({
            type: "WEB_CHAT/SEND_EVENT",
            payload: {
              name: "webchat/join",
              value: { language: window.navigator.language }
            }
          });
          setTimeout(() => this.setState({ botConnected: true }), 1000);
        } else if (action.type === "DIRECT_LINE/INCOMING_ACTIVITY") {
          const { activity } = action.payload;
          if (activity.type != undefined && activity.type === "message") {
            clearInterval(this.LastConversationDuration);
            let LastConversation = new Date();
            if (!this.isFeedBackSent) {
              this.LastConversationDuration = setInterval(
                () => this.checkIfEventNeedsToPosted(LastConversation),
                3000
              );
            }
          }
        }
        else if (action.type === 'DIRECT_LINE/POST_ACTIVITY') {
          action = SUI(
            action,
            ['payload', 'activity', 'channelData', 'browserName'],
            () => this.getBrowserName()
          );
        }

        return next(action);
      }
    );

    const myHeaders = new Headers();
    //cookie.remove('WantToChat');
    let WantToChat = cookie.load('WantToChat');
    if (WantToChat == "No") {
      this.setState({ modalIsOpen: false });
    }

    myHeaders.append("Authorization", "Bearer " + this.props.chatBotDialogProps.BotSecret);
    const BOtToken = await fetch("https://directline.botframework.com/v3/directline/tokens/generate", { method: "POST", headers: myHeaders }).then(r => r.json());
    if (this.props.chatBotDialogProps.BotSecret) {
      this.directLine = createDirectLine({
        token: fetchSpeechServicesToken(this.props.chatBotDialogProps.BotSecret)
      });
      this.directLine = createDirectLine({
        token: this.props.chatBotDialogProps.BotSecret,
        webSpeechPonyfillFactory: createCognitiveServicesSpeechServicesPonyfillFactory(
          {
            // TODO: [P3] Fetch token should be able to return different region
            region: this.props.chatBotDialogProps.TokenRegion,
            token: fetchSpeechServicesToken(
              this.props.chatBotDialogProps.BotSecret
            )
          }
        )
      });
    }

    const SpeechHeaders = new Headers();
    SpeechHeaders.append("Ocp-Apim-Subscription-Key", this.SpeechToken);
    const authorizationToken = await fetch("https://westus.api.cognitive.microsoft.com/sts/v1.0/issueToken", { method: "POST", headers: SpeechHeaders }).then(r => r.text());
    const region = this.props.chatBotDialogProps.TokenRegion;
    this.webSpeechPonyfillFactory = await createCognitiveServicesSpeechServicesPonyfillFactory(
      { region: region || "westus", authorizationToken: authorizationToken }
    );
  }

  private async checkIfUserIsBotAdmin() {
    await this.props.context.spHttpClient
      .get(
        this.props.context.pageContext.web.absoluteUrl +
        `/_api/web/sitegroups/getByName('` +
        this.BotAdminGroup +
        `')/Users?$filter=Email eq '` +
        this.props.context.pageContext.user.email +
        "'",
        SPHttpClient.configurations.v1
      )
      .then((responseListCustomer: SPHttpClientResponse) => {
        if (responseListCustomer.ok) {
          responseListCustomer.json().then(responseJSON => {
            if (responseJSON != null) {
              if (responseJSON.value.length > 0) {
                this.IsBotAdmin = true;
              } else {
                this.IsBotAdmin = false;
              }

              this.setState({});
            }
          });
        }
      });
  }

  public createSelectItems() {
    let _language: any = this.props.chatBotDialogProps.LangauageAvailable;
    if (_language == undefined) {
      _language = ["en-US:English"];
    }
    let items = [];
    for (let i = 0; i <= _language.length - 1; i++) {
      let _key = _language[i].split(":")[0];
      let _value = _language[i].split(":")[1];

      if (i == 0) {
        items.push(
          <option key={i + 1} value={_key}>
            {_value}
          </option>
        );
      } else {
        items.push(
          <option key={i + 1} value={_key}>
            {_value}
          </option>
        );
      }
    }
    return items;
  }
  private renderEditIcon() {
    let EditButton: any;
    if (this.IsBotAdmin) {
      EditButton = (
        <button className="switch" onClick={() => { this.clicked(); }}>
          <span className="ms-Icon ms-Icon--PageEdit" aria-hidden="true" />{" "}
        </button>
      );
    } else {
      EditButton = <div />;
    }
    return EditButton;
  }

  private GetInitials(name: string) {
    let parts = name.split(" ");
    let initials = "";
    for (var i = 0; i < parts.length; i++) {
      if (parts[i].length > 0 && parts[i] !== "") {
        if (i == 0 || i == parts.length - 1) initials += parts[i][0];
      }
    }
    return initials;
  }

  public GetToken = async () => {
    const myHeaders = new Headers();
    myHeaders.append(
      "Authorization",
      "Bearer " + this.props.chatBotDialogProps.BotSecret
    );
    const response = await fetch(
      "https://directline.botframework.com/v3/directline/tokens/generate",
      { method: "POST", headers: myHeaders }
    );
    const { BOtToken } = await response.json();
    return BOtToken;
  }

  private closeModalYes() {
    this.setState({ modalIsOpen: false });
    this.setState(() => ({
      minimized: false,
      newMessage: false,
      chatDisplay: "block",
      iconDisplay: "none"
    }));
    cookie.save('WantToChat', 'Yes');
  }
  private closeModalNo() {
    this.setState({ modalIsOpen: false });
    cookie.save('WantToChat', 'No');
  }
  private afterOpenModal() {
    // references are now sync'd and can be accessed.
    // this.subtitle.style.color = '#f00';
  }
  private openModal() {
    this.setState({ modalIsOpen: true });
  }

  private attachmentMiddleware = () => next => ({ activity, attachment, ...others }) => {
    const { activities } = this.store.getState();
    const messageActivities = activities.filter(activityItem => activityItem.type === 'message');
    const recentBotMessage = messageActivities.pop() === activity;

    switch (attachment.contentType) {
      case 'application/vnd.microsoft.card.adaptive':
        return <Components.AdaptiveCardContent
          content={attachment.content} disabled={!recentBotMessage} />;
      case 'application/vnd.microsoft.card.hero':
        return <Components.HeroCardContent
          content={attachment.content} disabled={!recentBotMessage} />;
      case 'application/vnd.microsoft.card.thumbnail':
        return <Components.ThumbnailCardContent
          content={attachment.content} disabled={!recentBotMessage}
        />;

      default:
        return next({ activity, attachment, ...others });
    }
  }

  private BotActivityDecorator = ({ activityID, children }) => {
    const [{ feedbackGiven, activateButtons, selectedButton }, setFeedbackState] =
      React.useState({ feedbackGiven: false, activateButtons: false, selectedButton: '' });

    const upButtonClass = selectedButton === 'UP' ?
      'botActivityDecorator__button selected' : 'botActivityDecorator__button';
    const downButtonClass = selectedButton === 'DOWN' ?
      'botActivityDecorator__button selected' : 'botActivityDecorator__button';

    return <div className="botActivityDecorator" onMouseEnter={() =>
      setFeedbackState({ feedbackGiven, activateButtons: true, selectedButton })}
      onMouseLeave={() => setFeedbackState({ feedbackGiven, activateButtons: false, selectedButton })}>
      {children}
      {activateButtons && 
      <div className="botActivityDecorator__buttonBar">
        <button className={upButtonClass + ' up'}
          title={this.thumbsUpTooltip}
          disabled={feedbackGiven}
          onClick={() => {
            setFeedbackState({ feedbackGiven: true, activateButtons, selectedButton: 'UP' });
            this.store.dispatch({
              type: "DIRECT_LINE/POST_ACTIVITY",
              meta: "keyboard",
              payload: {
                activity: {
                  name: "reactionAdded",
                  type: "messageReaction",
                  value: { activityID },
                  reactionsAdded: [{ activityID, helpful: 1 }],
                  replyToID: activityID
                }
              }
            });
          }}><span className="glyphicon glyphicon-thumbs-up"></span></button>
        <button className={downButtonClass + ' down'}
          disabled={feedbackGiven}
          title={this.thumbsDownTooltip}
          onClick={() => {
            setFeedbackState({ feedbackGiven: true, activateButtons, selectedButton: 'DOWN' });
            this.store.dispatch({
              type: "DIRECT_LINE/POST_ACTIVITY",
              meta: "keyboard",
              payload: {
                activity: {
                  name: "reactionRemoved",
                  type: "messageReaction",
                  value: { activityID },
                  reactionsRemoved: [{ activityID, helpful: -1 }],
                  replyToID: activityID
                }
              }
            });
          }}><span className="glyphicon glyphicon-thumbs-down"></span></button>
      </div>}
    </div>;
  }

  private activityMiddleware = () => next => card => {
    if (card.activity.from.role === 'bot' && !card.activity.HideFeedbackButtons) {
      return children => (
        <this.BotActivityDecorator activityID={card.activity.id} key={card.activity.id}>
          {next(card)(children)}
        </this.BotActivityDecorator>
      );
    } else {
      return next(card);
    }
  }

  public render() {
    const {
      state: {
        minimized,
        newMessage,
        side,
        _token,
        localeLang,
        chatDisplay,
        iconDisplay,
        photoUrl
      }
    } = this;

    const styleSet = createStyleSet({
      alignItems: "left",
      justifyContent: "left"
    });
    styleSet.avatar = {
      alignItems: "left",
      justifyContent: "left"
    };
    const styleOptions = {
      // Colors
      bubbleBackground: "rgba(240, 240, 240, 0.1)",
      bubbleFromUserBackground:
        this.props.chatBotDialogProps.PickedColor && "rgb(210, 240, 255)",
      bubbleFromUserTextColor: "black",
      bubbleBorderRadius: 10,
      bubbleFromUserBorderRadius: 10,
      suggestedActionBorderRadius: 10,
      bubbleMaxWidth: 600, // maximum width of text message

      // Avatar
      botAvatarImage: this.props.chatBotDialogProps.Boticon,
      userAvatarImage: this.props.photoUrl,
      userAvatarInitials: this.GetInitials(this.props.userName),

      // Send box
      hideSendBox: false,
      hideUploadButton: true,
      sendBoxButtonColor: "#767676",
      sendBoxButtonColorOnDisabled: "#CCC",
      sendBoxButtonColorOnFocus: "#333",
      sendBoxButtonColorOnHover: "#333",
      sendBoxButtonAlignItems: "left",
      sendBoxButtonJustifyContent: "left",
      sendBoxHeight: 40,

      // Suggested actions
      suggestedActionTextColor: "black",
      suggestedActionBorder: "solid 1px rgb(0, 99, 177)",
      suggestedActionAlignItems: "left",
      suggestedActionJustifyContent: "left",
      suggestedActionHeight: 30
    };

    const botStyle = { backgroundColor: this.props.chatBotDialogProps.PickedColor };
    return (
      <div>
        <div className="minimizable-web-chat">
          {minimized ? (
            <button className="maximize" onClick={this.handleMaximizeButtonClick} style={this.props.chatBotDialogProps.PickedColor && { backgroundColor: this.props.chatBotDialogProps.PickedColor }} >
              <span className={_token ? "ms-Icon ms-Icon--MessageFill" : "ms-Icon ms-Icon--Message"} />
              {newMessage && (<span className="ms-Icon ms-Icon--CircleShapeSolid red-dot" />)}
            </button>
          ) : (
              <div className={side === "left" ? "chat-box left" : "chat-box right"} style={{ backgroundColor: this.props.chatBotDialogProps.PickedColor }} >
                <header style={this.props.chatBotDialogProps.PickedColor && { backgroundColor: this.props.chatBotDialogProps.PickedColor }}>
                  {this.props.chatBotDialogProps.Boticon ? (
                    <div>
                      <img src={this.props.chatBotDialogProps.Boticon} width="34px" className="botImage" alt={this.props.chatBotDialogProps.BotName + ' bot'} />
                    </div>
                  ) : (<div></div>)}
                  <div className="botIcon">
                    {this.props.chatBotDialogProps.BotName}{" "}
                  </div>
                  <div className="filler" />
                  <label className='label-noshow' htmlFor='lang'>Select Language:</label>
                  <select className="languageSelector" id="lang"
                    style={{ backgroundColor: this.props.chatBotDialogProps.PickedColor }}
                    onChange={this.handleTextChangedEvent} >
                    {this.createSelectItems()} </select>
                  <button className="minimize" onClick={this.helpButtonClick}>
                    <span className="ms-Icon ms-Icon--Help" aria-hidden="true" />
                  </button>
                  {this.renderEditIcon()}
                  <button className="switch" onClick={this.handleSwitchButtonClick} >
                    <span className="ms-Icon ms-Icon--Switch" aria-hidden="true" />
                  </button>
                  <button className="minimize" onClick={this.handleMinimizeButtonClick} >
                    <span className="ms-Icon ms-Icon--ChromeMinimize" aria-hidden="true" />
                  </button>
                </header>

                {this.props.chatBotDialogProps.BotSecret ? (
                  this.SpeechToken ? (<ReactWebChat userID={this.props.email}
                    username={this.props.userName} directLine={this.directLine}
                    styleOptions={styleOptions} resize="detect"
                    locale={localeLang} store={this.store}
                    attachmentMiddleware={this.attachmentMiddleware}
                    activityMiddleware={this.activityMiddleware}
                    webSpeechPonyfillFactory={this.webSpeechPonyfillFactory} />
                  ) : (<ReactWebChat userID={this.props.email} username={this.props.userName}
                    directLine={this.directLine} styleOptions={styleOptions} resize="detect"
                    locale={localeLang} store={this.store}
                    attachmentMiddleware={this.attachmentMiddleware}
                    activityMiddleware={this.activityMiddleware}
                  />)
                ) : (<div> Please configure bot properties </div>)}

                {this.props.chatBotDialogProps.BotSecret && !this.state.botConnected &&
                  <div className='loader'><div className="lds-dual-ring"></div></div>}
              </div>
            )}
        </div>

        {this.state.modalIsOpen && <div id="popup-box-container">
          <div id="popup-body">
            <a onClick={this.closeModalNo} id="novocall-popup-close" className="close-button">×</a>
            <div className="popup-box">
              <div className="popup-header" style={botStyle} >
                <img className="widget-popup-profile-image img-circle" src={this.props.chatBotDialogProps.Boticon} alt={this.props.chatBotDialogProps.BotName + ' bot'} />
                <div className="widget-popup-profile-info">
                  <div className="popup-name">{this.props.chatBotDialogProps.BotName}</div>
                </div>
              </div>
              <div className="popup-body">
                <div className="widget-popup-welcome-message"
                  dangerouslySetInnerHTML={{ __html: this.popupWelcomeMessage }}>
                </div>
                <div className="popup-button-small">
                  <Button onClick={this.closeModalYes} style={botStyle}>Help me</Button>
                </div>
              </div>
            </div>
          </div>
        </div>}
      </div>
    );
  }

  private getBrowserName() {
    let _window: any = window;
    let _document: any = document;
    // Opera 8.0+
    let isOpera = (!!_window.opr && !!_window.opr.addons) || !!_window.opera || navigator.userAgent.indexOf(' OPR/') >= 0;

    // Firefox 1.0+
    let isFirefox = typeof _window.InstallTrigger !== 'undefined';

    // Safari 3.0+ "[object HTMLElementConstructor]" 
    let isSafari = /constructor/i.test(_window.HTMLElement) || ((p) => { return p.toString() === "[object SafariRemoteNotification]"; })(!_window['safari'] || (typeof _window.safari !== 'undefined' && _window.safari.pushNotification));

    // Internet Explorer 6-11
    let isIE = /*@cc_on!@*/false || !!_document.documentMode;

    // Edge 20+
    let isEdge = !isIE && !!_window.StyleMedia;

    // Chrome 1 - 79
    let isChrome = !!_window.chrome && (!!_window.chrome.webstore || !!_window.chrome.runtime);

    // Edge (based on chromium) detection
    let isEdgeChromium = isChrome && (navigator.userAgent.indexOf("Edg") != -1);

    // Blink engine detection
    let isBlink = (isChrome || isOpera) && !!_window.CSS;

    if (isChrome) return 'Google Chrome';
    if (isFirefox) return 'Mozilla Firefox';
    if (isIE) return 'Internet Explorer';
    if (isSafari) return 'Safari';
    if (isEdge) return 'MS Edge';
    if (isEdgeChromium) return 'MS Edge (Chromium)';
  }
}
