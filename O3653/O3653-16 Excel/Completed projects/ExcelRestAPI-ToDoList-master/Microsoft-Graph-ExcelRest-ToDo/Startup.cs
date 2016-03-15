using System;
using System.Configuration;
using System.Threading.Tasks;
using System.Globalization;
using System.IdentityModel.Tokens;
using System.Web;
using Owin;
using Microsoft.Owin;
using Microsoft.Owin.Security;
using Microsoft.Owin.Security.Cookies;
using Microsoft.Owin.Security.Notifications;
using Microsoft.Owin.Security.OpenIdConnect;
using Microsoft.IdentityModel.Protocols;
using Microsoft_Graph_ExcelRest_ToDo.TokenStorage;
using ADAL = Microsoft.IdentityModel.Clients.ActiveDirectory;

[assembly: OwinStartup(typeof(Microsoft_Graph_ExcelRest_ToDo.Startup))]

namespace Microsoft_Graph_ExcelRest_ToDo
{
    public class Startup
    {
        public static string appId = ConfigurationManager.AppSettings["ida:AppId"];
        public static string appSecret = ConfigurationManager.AppSettings["ida:AppSecret"];
        public static string aadInstance = ConfigurationManager.AppSettings["ida:AADInstance"];

        public void Configuration(IAppBuilder app)
        {
            app.SetDefaultSignInAsAuthenticationType(CookieAuthenticationDefaults.AuthenticationType);

            app.UseCookieAuthentication(new CookieAuthenticationOptions());

            app.UseOpenIdConnectAuthentication(
              new OpenIdConnectAuthenticationOptions
              {
            // The `Authority` represents the auth endpoint - https://login.microsoftonline.com/common/
            // The 'ResponseType' indicates that we want an authorization code and an ID token 
            // In a real application you could use issuer validation for additional checks, like making 
            // sure the user's organization has signed up for your app, for instance.

            ClientId = appId,
                  Authority = string.Format(CultureInfo.InvariantCulture, aadInstance, "common", ""),
                  ResponseType = "code id_token",
                  PostLogoutRedirectUri = "/",
                  TokenValidationParameters = new TokenValidationParameters
                  {
                      ValidateIssuer = false,
                  },
                  Notifications = new OpenIdConnectAuthenticationNotifications
                  {
                      AuthenticationFailed = OnAuthenticationFailed,
                      AuthorizationCodeReceived = OnAuthorizationCodeReceived
                  }
              }
            );
        }

        private Task OnAuthenticationFailed(AuthenticationFailedNotification<OpenIdConnectMessage,
          OpenIdConnectAuthenticationOptions> notification)
        {
            notification.HandleResponse();
            notification.Response.Redirect("/Error?message=" + notification.Exception.Message);
            return Task.FromResult(0);
        }

        private async Task OnAuthorizationCodeReceived(AuthorizationCodeReceivedNotification notification)
        {
            // Get the user's object id (used to name the token cache)
            string userObjId = notification.AuthenticationTicket.Identity
              .FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier").Value;

            // Create a token cache
            HttpContextBase httpContext = notification.OwinContext.Get<HttpContextBase>(typeof(HttpContextBase).FullName);
            SessionTokenCache tokenCache = new SessionTokenCache(userObjId, httpContext);

            // Exchange the auth code for a token
            ADAL.ClientCredential clientCred = new ADAL.ClientCredential(appId, appSecret);

            // Create the auth context
            ADAL.AuthenticationContext authContext = new ADAL.AuthenticationContext(
              string.Format(CultureInfo.InvariantCulture, aadInstance, "common", ""),
              false, tokenCache);

            ADAL.AuthenticationResult authResult = await authContext.AcquireTokenByAuthorizationCodeAsync(
              notification.Code, notification.Request.Uri, clientCred, "https://graph.microsoft.com");
        }
    }
}
