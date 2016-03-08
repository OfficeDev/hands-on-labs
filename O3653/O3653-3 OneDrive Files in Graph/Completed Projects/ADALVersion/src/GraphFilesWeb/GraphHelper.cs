

namespace GraphFilesWeb
{
    public static class GraphHelper
    {
        public static string GetGraphAccessToken()
        {
            //var AzureAdGraphResourceURL = "https://graph.microsoft.com/";

            //var Authority = config["Authentication:AzureAd:AADInstance"] + "Common";

            //var signInUserId = ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier).Value;
            //var userObjectId = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier").Value;
            //var clientCredential = new ClientCredential(config["Authentication:AzureAd:ClientId"]);
            //var userIdentifier = new UserIdentifier(userObjectId, UserIdentifierType.UniqueId);

            //// create auth context
            //AuthenticationContext authContext = new AuthenticationContext(Authority, new ADALTokenCache(signInUserId));
            //var result = await authContext.AcquireTokenSilentAsync(AzureAdGraphResourceURL, clientCredential, userIdentifier);

            //return result.AccessToken;

            return "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6Ik1uQ19WWmNBVGZNNXBPWWlKSE1iYTlnb0VLWSIsImtpZCI6Ik1uQ19WWmNBVGZNNXBPWWlKSE1iYTlnb0VLWSJ9.eyJhdWQiOiJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20iLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC85ZDg0MGQxYS0zOGJkLTRlYzItOWVmMi0xMTZmNmJiNmVlMGEvIiwiaWF0IjoxNDU2OTg4ODAzLCJuYmYiOjE0NTY5ODg4MDMsImV4cCI6MTQ1Njk5MjcwMywiYWNyIjoiMSIsImFtciI6WyJwd2QiXSwiYXBwaWQiOiIyMWNjOWIyYS0yZmVhLTQwNjktYmQ1MC0xODk2ZGY3MTY3ZGUiLCJhcHBpZGFjciI6IjEiLCJmYW1pbHlfbmFtZSI6IkdyZWdnIiwiZ2l2ZW5fbmFtZSI6IlJ5YW4iLCJpcGFkZHIiOiIyMy45OS4zLjEwMCIsIm5hbWUiOiJSeWFuIEdyZWdnIiwib2lkIjoiZWZlZTFiNzctZmIzYi00ZjY1LTk5ZDYtMjc0YzExOTE0ZDEyIiwicHVpZCI6IjEwMDMzRkZGOEM0MUQwNkEiLCJzY3AiOiJNeUZpbGVzLldyaXRlIFVzZXIuUmVhZCBVc2VyLlJlYWQuQWxsIiwic3ViIjoiVjcyY21ZSngzdy04MTFEOHZISmRMWUZDa0NFOFJDZktxZmh5Qk92dVZOQSIsInRpZCI6IjlkODQwZDFhLTM4YmQtNGVjMi05ZWYyLTExNmY2YmI2ZWUwYSIsInVuaXF1ZV9uYW1lIjoicmdyZWdnQHNlYXR0bGVhcHB3b3Jrcy5vbm1pY3Jvc29mdC5jb20iLCJ1cG4iOiJyZ3JlZ2dAc2VhdHRsZWFwcHdvcmtzLm9ubWljcm9zb2Z0LmNvbSIsInZlciI6IjEuMCJ9.MuL7pBLoWouFftc53P_nHeJHfI9CxOVPc5k96wPWMzQ8HQAcydswIDbQdEF_uK1krgRUbockk3Vzzar0I_TZK4u78dSHB2Yz9lNDlax2XPoUa0ugO1tb1qpnVL_p29Es4KqtHKvW6Svz6pMdTuStgdpVuXlgt0rZF2_67ogFwHfpdT4qgkVPAA-atLMWhz4XWfxQFHiULEyBBD9K8ULA94-_3Ln5HMx0gq5GlscEiZbYoAzMAUgaUADXoBYPhg-EyivL4VeXEZSv0lRjKMQewG45QDNH2y4-b6mqUW7Q-QCAGpB0E_Tqn_4M70ocmgIWcEoqBqSFBQyA953GnhUZEg";
        }


    }
}
