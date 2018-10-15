using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Http;
using System.Security.Claims;
using System.Net;

using System.Net.Http.Headers;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System.Configuration;
using System.Threading.Tasks;
using Microsoft.Bot.Connector;
using System.Net.Http;

namespace LuisBot.Controllers
{
    [Authorize]
    public class WebChatController : ApiController
    {
        private string clientId = ConfigurationManager.AppSettings["ida:ClientId"];
        private string appKey = ConfigurationManager.AppSettings["ida:ClientSecret"];
        private string aadInstance = ConfigurationManager.AppSettings["ida:AADInstance"];
        private string tenantId = ConfigurationManager.AppSettings["ida:TenantId"];

        private string SPAccessToken = null;
        public HttpResponseMessage Get()
        {
            GetSPTokenSilent().Wait();
            var userId = ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier).Value;

            string WebChatString =
                new WebClient().DownloadString("https://webchat.botframework.com/embed/SharepointSearchLuis?s=ukaTOD836(]*fqnnDHHY58:" +
                HttpUtility.UrlEncode(userId) + "&atul.gupta@ABCAtul.onmicrosoft.com=" + HttpUtility.UrlEncode(ClaimsPrincipal.Current.Identity.Name));

            WebChatString = WebChatString.Replace("/css/botchat.css", "https://webchat.botframework.com/css/botchat.css");
            WebChatString = WebChatString.Replace("/scripts/botchat.js", "https://webchat.botframework.com/scripts/botchat.js");
            var response = new HttpResponseMessage();
            response.Content = new StringContent(WebChatString);
            response.Content.Headers.ContentType = new MediaTypeHeaderValue("text/html");
            var botCred = new MicrosoftAppCredentials(
                ConfigurationManager.AppSettings["MicrosoftAppId"],
                ConfigurationManager.AppSettings["MicrosoftAppPassword"]);
            var stateClient = new StateClient(botCred);
            BotState botState = new BotState(stateClient);
            BotData botData = new BotData(eTag: "*");
            botData.SetProperty<string>("SPAccessToken", SPAccessToken);
            stateClient.BotState.SetUserDataAsync("webchat", userId, botData).Wait();
            return response;
        }

        private Task GetSPTokenSilent()
        {
            return Task.Run(async () =>
            {
                string userObjectID = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier").Value;
                string signedInUserID = ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier).Value;
                AuthenticationContext authContext = new AuthenticationContext(aadInstance + tenantId, new ADALTokenCache(signedInUserID));
                ClientCredential cred = new ClientCredential(clientId, appKey);
                AuthenticationResult res = await authContext.AcquireTokenSilentAsync("https://abcatul.sharepoint.com", cred,
                    new UserIdentifier(userObjectID, UserIdentifierType.UniqueId));
                SPAccessToken = res.AccessToken;
            });

        }
    }
}