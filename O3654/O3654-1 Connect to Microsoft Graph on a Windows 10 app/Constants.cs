// ------------------------------------------------------------------------------
//  Copyright (c) 2016 Microsoft Corporation
// 
//  Permission is hereby granted, free of charge, to any person obtaining a copy
//  of this software and associated documentation files (the "Software"), to deal
//  in the Software without restriction, including without limitation the rights
//  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
//  copies of the Software, and to permit persons to whom the Software is
//  furnished to do so, subject to the following conditions:
// 
//  The above copyright notice and this permission notice shall be included in
//  all copies or substantial portions of the Software.
// 
//  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
//  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
//  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
//  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
//  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
//  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
//  THE SOFTWARE.
// ------------------------------------------------------------------------------

namespace Microsoft.Graph.Authentication
{
    public static class Constants
    {
        public static class Authentication
        {
            public const string AccessTokenKeyName = "access_token";

            public const string AccessTokenTypeKeyName = "token_type";

            public const string AuthenticationCancelled = "authentication_cancelled";

            public const string AuthorizationCodeGrantType = "authorization_code";

            public const string AuthorizationServiceKey = "authorization_service";

            public const string ClientIdKeyName = "client_id";

            public const string ClientSecretKeyName = "client_secret";

            public const string CodeKeyName = "code";

            public const string ErrorDescriptionKeyName = "error_description";

            public const string ErrorKeyName = "error";

            public const string ExpiresInKeyName = "expires_in";

            public const string GrantTypeKeyName = "grant_type";

            public const string RedirectUriKeyName = "redirect_uri";

            public const string RefreshTokenKeyName = "refresh_token";

            public const string ResponseTypeKeyName = "response_type";

            public const string ScopeKeyName = "scope";

            public const string TokenResponseTypeValueName = "token";

            public const string TokenServiceKey = "token_service";

            public const string TokenTypeKeyName = "token_type";

            public const string UserIdKeyName = "user_id";

            internal const string ActiveDirectoryAuthenticationServiceUrl = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize";

            internal const string ActiveDirectorySignOutUrl = "https://login.microsoftonline.com/common/oauth2/v2.0/logout";

            internal const string ActiveDirectoryTokenServiceUrl = "https://login.microsoftonline.com/common/oauth2/v2.0/token";
        }

        public static class Headers
        {
            public const string Bearer = "Bearer";

            public const string SdkVersionHeaderName = "X-ClientService-ClientTag";

            public const string FormUrlEncodedContentType = "application/x-www-form-urlencoded";

            public const string SdkVersionHeaderValue = "SDK-Version=CSharp-v{0}";

            public const string ThrowSiteHeaderName = "X-ThrowSite";
        }

        public static class Serialization
        {
            public const string ODataType = "@odata.type";
        }

        public static class Url
        {
            public const string Drive = "drive";

            public const string Root = "root";

            public const string AppRoot = "approot";
        }
    }
}
