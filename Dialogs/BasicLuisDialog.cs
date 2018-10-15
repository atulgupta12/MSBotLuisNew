using System;
using System.Configuration;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;

using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Builder.Luis;
using Microsoft.Bot.Builder.Luis.Models;
using Microsoft.Bot.Connector;
using System.Net.Http;
using System.Collections.Generic;
using Newtonsoft.Json;
using System.Text;
using Microsoft.SharePoint.Client.Search.Query;

namespace Microsoft.Bot.Sample.LuisBot
{
    // For more information about this template visit http://aka.ms/azurebots-csharp-luis
    [Serializable]
    public class BasicLuisDialog : LuisDialog<object>
    {
        const string SPAccessTokenKey = "SPAccessToken";
        const string SPSite = "https://abcatul.sharepoint.com";
        static string qnamaker_endpointKey = "830e6887-5031-4227-9818-ff4891b44023";
		static string qnamaker_endpointDomain = "botsharepointsearch";
		static string HR_kbID = "b664476d-faa0-4709-8dc7-f5e3662bd31c";

        private static readonly Dictionary<string, string> PropertyMappings
        = new Dictionary<string, string>
    {
        { "TypeOfDocument", "kbDocType" },
        { "Software", "kbTopic" }
    };
        [Serializable]
        public class PartialMessage
        {
            public string Text { set; get; }
        }
        private PartialMessage message;

        //internal BasicLuisDialog() { }

        protected override async Task MessageReceived(IDialogContext context,
            IAwaitable<Microsoft.Bot.Connector.IMessageActivity> item)
        {
            var msg = await item;

            if (string.IsNullOrEmpty(context.UserData.Get<string>(SPAccessTokenKey)))
            {
                MicrosoftAppCredentials cred = new MicrosoftAppCredentials(
                    ConfigurationManager.AppSettings["MicrosoftAppId"],
                    ConfigurationManager.AppSettings["MicrosoftAppPassword"]);
                StateClient stateClient = new StateClient(cred);
                BotState botState = new BotState(stateClient);
                BotData botData = await botState.GetUserDataAsync(msg.ChannelId, msg.From.Id);
                context.UserData.SetValue<string>(SPAccessTokenKey, botData.GetProperty<string>(SPAccessTokenKey));
            }

            this.message = new PartialMessage { Text = msg.Text };
            await base.MessageReceived(context, item);
        }

        public QnAMakerService hrQnAService = new QnAMakerService("https://" + qnamaker_endpointDomain + ".azurewebsites.net", HR_kbID, qnamaker_endpointKey);

		
        public BasicLuisDialog() : base(new LuisService(new LuisModelAttribute(
            ConfigurationManager.AppSettings["LuisAppId"], 
            ConfigurationManager.AppSettings["LuisAPIKey"], 
            domain: ConfigurationManager.AppSettings["LuisAPIHostName"])))
        {
        }

        [LuisIntent("None")]
        public async Task NoneIntent(IDialogContext context, LuisResult result)
        {
            await this.ShowLuisResult(context, result);
        }

        // Go to https://luis.ai and create a new intent, then train/publish your luis app.
        // Finally replace "Greeting" with the name of your newly created intent in the following handler
        [LuisIntent("Greeting")]
        public async Task GreetingIntent(IDialogContext context, LuisResult result)
        {
			var qnaMakerAnswer = await hrQnAService.GetAnswer(result.Query);
			await context.PostAsync($"{qnaMakerAnswer}");
			context.Wait(MessageReceived);
            //await this.ShowLuisResult(context, result);
        }

        [LuisIntent("Cancel")]
        public async Task CancelIntent(IDialogContext context, LuisResult result)
        {
            await this.ShowLuisResult(context, result);
        }

        [LuisIntent("Help")]
        public async Task HelpIntent(IDialogContext context, LuisResult result)
        {
            await this.ShowLuisResult(context, result);
        }
		
		[LuisIntent("FindDocumentation")]
        public async Task FindIntent(IDialogContext context, LuisResult result)
        {
            var reply = context.MakeMessage();
            try
            {
                reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;
                reply.Attachments = new List<Microsoft.Bot.Connector.Attachment>();
                StringBuilder query = new StringBuilder();
                bool QueryTransformed = false;
                if (result.Entities.Count > 0)
                {
                    QueryTransformed = true;
                    foreach (var entity in result.Entities)
                    {
                        if (PropertyMappings.ContainsKey(entity.Type))
                        {
                            query.AppendFormat("{0}:'{1}' ", PropertyMappings[entity.Type], entity.Entity);
                        }
                    }
                }
                else
                {
                    //should replace all special chars
                    query.Append(this.message.Text.Replace("?", ""));
                }

                using (ClientContext ctx = new ClientContext(SPSite))
                {
                    ctx.AuthenticationMode = ClientAuthenticationMode.Anonymous;
                    ctx.ExecutingWebRequest +=
                        delegate (object oSender, WebRequestEventArgs webRequestEventArgs)
                        {
                            webRequestEventArgs.WebRequestExecutor.RequestHeaders["Authorization"] =
                                "Bearer " + context.UserData.Get<string>("SPAccessToken");
                        };
                    KeywordQuery kq = new KeywordQuery(ctx);
                    kq.QueryText = string.Concat(query.ToString(), " IsDocument:1");
                    kq.RowLimit = 5;
                    SearchExecutor se = new SearchExecutor(ctx);
                    ClientResult<ResultTableCollection> results = se.ExecuteQuery(kq);
                    ctx.ExecuteQuery();

                    if (results.Value != null && results.Value.Count > 0 && results.Value[0].RowCount > 0)
                    {
                        reply.Text += (QueryTransformed == true) ? "I found some interesting reading for you!" : "I found some potential interesting reading for you!";
                        BuildReply(results, reply);
                    }
                    else
                    {
                        if (QueryTransformed)
                        {
                            //fallback with the original message
                            kq.QueryText = string.Concat(this.message.Text.Replace("?", ""), " IsDocument:1");
                            kq.RowLimit = 3;
                            se = new SearchExecutor(ctx);
                            results = se.ExecuteQuery(kq);
                            ctx.ExecuteQuery();
                            if (results.Value != null && results.Value.Count > 0 && results.Value[0].RowCount > 0)
                            {
                                reply.Text += "I found some potential interesting reading for you!";
                                BuildReply(results, reply);
                            }
                            else
                                reply.Text += "I could not find any interesting document!";
                        }
                        else
                            reply.Text += "I could not find any interesting document!";

                    }

                }

            }
            catch (Exception ex)
            {
                reply.Text = ex.Message;
            }
            await context.PostAsync(reply);
            context.Wait(MessageReceived);
            //await this.ShowLuisResult(context, result);
        }

        void BuildReply(ClientResult<ResultTableCollection> results, IMessageActivity reply)
        {
            foreach (var row in results.Value[0].ResultRows)
            {
                List<CardAction> cardButtons = new List<CardAction>();
                List<CardImage> cardImages = new List<CardImage>();
                string ct = string.Empty;
                string icon = string.Empty;
                switch (row["FileExtension"].ToString())
                {
                    case "docx":
                        ct = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                        icon = "https://cdn2.iconfinder.com/data/icons/metro-ui-icon-set/128/Word_15.png";
                        break;
                    case "xlsx":
                        ct = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                        icon = "https://cdn2.iconfinder.com/data/icons/metro-ui-icon-set/128/Excel_15.png";
                        break;
                    case "pptx":
                        ct = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
                        icon = "https://cdn2.iconfinder.com/data/icons/metro-ui-icon-set/128/PowerPoint_15.png";
                        break;
                    case "pdf":
                        ct = "application/pdf";
                        icon = "https://cdn4.iconfinder.com/data/icons/CS5/256/ACP_PDF%202_file_document.png";
                        break;

                }
                cardButtons.Add(new CardAction
                {
                    Title = "Open",
                    Value = (row["ServerRedirectedURL"] != null) ? row["ServerRedirectedURL"].ToString() : row["Path"].ToString(),
                    Type = ActionTypes.OpenUrl
                });
                cardImages.Add(new CardImage(url: icon));
                ThumbnailCard tc = new ThumbnailCard();
                tc.Title = (row["Title"] != null) ? row["Title"].ToString() : "Untitled";
                tc.Text = (row["Description"] != null) ? row["Description"].ToString() : string.Empty;
                tc.Images = cardImages;
                tc.Buttons = cardButtons;
                reply.Attachments.Add(tc.ToAttachment());
            }
        }

        private async Task ShowLuisResult(IDialogContext context, LuisResult result) 
        {
            await context.PostAsync($"You have reached {result.Intents[0].Intent}. You said: {result.Query}");
            context.Wait(MessageReceived);
        }
    }
	
	public class Metadata
	{
		public string name { get; set; }
		public string value { get; set; }
	}

	public class Answer
	{
		public IList<string> questions { get; set; }
		public string answer { get; set; }
		public double score { get; set; }
		public int id { get; set; }
		public string source { get; set; }
		public IList<object> keywords { get; set; }
		public IList<Metadata> metadata { get; set; }
	}

	public class QnAAnswer
	{
		public IList<Answer> answers { get; set; }
	}
	
	[Serializable]
	public class QnAMakerService
	{
		private string qnaServiceHostName;
		private string knowledgeBaseId;
		private string endpointKey;

		public QnAMakerService(string hostName, string kbId, string endpointkey)
		{
			qnaServiceHostName = hostName;
			knowledgeBaseId = kbId;
			endpointKey = endpointkey;

		}
		async Task<string> Post(string uri, string body)
		{
			using (var client = new HttpClient())
			using (var request = new HttpRequestMessage())
			{
				request.Method = HttpMethod.Post;
				request.RequestUri = new Uri(uri);
				request.Content = new StringContent(body, Encoding.UTF8, "application/json");
				request.Headers.Add("Authorization", "EndpointKey " + endpointKey);

				var response = await client.SendAsync(request);
				return  await response.Content.ReadAsStringAsync();
			}
		}
		public async Task<string> GetAnswer(string question)
		{
			string uri = qnaServiceHostName + "/qnamaker/knowledgebases/" + knowledgeBaseId + "/generateAnswer";
			string questionJSON = @"{'question': '" + question + "'}";

			var response = await Post(uri, questionJSON);

			var answers = JsonConvert.DeserializeObject<QnAAnswer>(response);
			if (answers.answers.Count > 0)
			{
				return answers.answers[0].answer;
			}
			else
			{
				return "No good match found.";
			}
		}
	}
}

