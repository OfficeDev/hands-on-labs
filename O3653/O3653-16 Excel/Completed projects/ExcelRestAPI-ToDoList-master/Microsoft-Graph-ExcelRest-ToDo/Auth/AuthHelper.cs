using System.Threading.Tasks;
using System.Web;
using Microsoft_Graph_ExcelRest_ToDo.TokenStorage;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Owin.Security;
using Microsoft.Owin.Security.OpenIdConnect;

namespace Microsoft_Graph_ExcelRest_ToDo.Auth
{
    public class AuthHelper
    {
        // This is the logon authority
        // i.e. https://login.microsoftonline.com/common
        public string Authority { get; set; }
        // This is the application ID obtained from registering at
        // https://apps.dev.microsoft.com
        public string AppId { get; set; }
        // This is the application secret obtained from registering at
        // https://apps.dev.microsoft.com
        public string AppSecret { get; set; }
        // This is the token cache
        public SessionTokenCache TokenCache { get; set; }

        public AuthHelper(string authority, string appId, string appSecret, SessionTokenCache tokenCache)
        {
            Authority = authority;
            AppId = appId;
            AppSecret = appSecret;
            TokenCache = tokenCache;
        }

        public async Task<string> GetUserAccessToken(string redirectUri)
        {
            AuthenticationContext authContext = new AuthenticationContext(Authority, false, TokenCache);

            ClientCredential credential = new ClientCredential(AppId, AppSecret);

            try
            {
                AuthenticationResult authResult = await authContext.AcquireTokenSilentAsync("https://graph.microsoft.com", credential,
                  new UserIdentifier(TokenCache.UserObjectId, UserIdentifierType.UniqueId));
                return authResult.AccessToken;
            }
            catch (AdalSilentTokenAcquisitionException)
            {
                HttpContext.Current.Request.GetOwinContext().Authentication.Challenge(
                  new AuthenticationProperties() { RedirectUri = redirectUri },
                  OpenIdConnectAuthenticationDefaults.AuthenticationType);

                return null;
            }
        }
    }
}