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
    using System;
    using System.Collections.Generic;
    using System.Net;
    using System.Net.Http;
    using System.Net.Http.Headers;
    using System.Text;
    using System.Threading.Tasks;
    using Windows.Security.Authentication.Web;

    public class OAuth2AuthenticationProvider : IAuthenticationProvider
    {
        private const string authenticationServiceUrl = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize";
        private const string tokenServiceUrl = "https://login.microsoftonline.com/common/oauth2/v2.0/token";

        private string clientId;
        private string[] scopes;
        private string returnUrl;

        private string accessToken;
        private string refreshToken;

        private DateTimeOffset expiration;

        private IHttpProvider httpProvider;

        /// <summary>
        /// Creates a new instance of <see cref="OAuth2AuthenticationProvider"/> for authentication in Windows Store applications.
        /// </summary>
        public OAuth2AuthenticationProvider(
            string clientId,
            string returnUrl,
            string[] scopes,
            IHttpProvider httpProvider = null)
        {
            if (string.IsNullOrEmpty(clientId))
            {
                throw new ServiceException(
                    new Error
                    {
                        Code = GraphErrorCode.InvalidRequest.ToString(),
                        Message = "Client ID is required to authenticate using OAuth2AuthenticationProvider.",
                    });
            }

            if (string.IsNullOrEmpty(returnUrl))
            {
                throw new ServiceException(
                    new Error
                    {
                        Code = GraphErrorCode.InvalidRequest.ToString(),
                        Message = "Return URL is required to authenticate using OAuth2AuthenticationProvider.",
                    });
            }

            if (scopes == null || scopes.Length == 0)
            {
                throw new ServiceException(
                    new Error
                    {
                        Code = GraphErrorCode.InvalidRequest.ToString(),
                        Message = "No scopes have been requested for authentication.",
                    });
            }

            this.clientId = clientId;
            this.returnUrl = returnUrl;
            this.scopes = scopes;
            this.httpProvider = httpProvider ?? new HttpProvider();
        }

        /// <summary>
        /// Presents authentication UI to the user.
        /// </summary>
        /// <returns>The task to await.</returns>
        public async Task AuthenticateAsync()
        {
            if (string.IsNullOrEmpty(this.accessToken) || this.expiration <= DateTimeOffset.Now.UtcDateTime.AddMinutes(5))
            {
                await this.GetAuthenticationResultAsync();

                if (string.IsNullOrEmpty(accessToken))
                {
                    throw new ServiceException(
                        new Error
                        {
                            Code = GraphErrorCode.AuthenticationFailure.ToString(),
                            Message = "Failed to retrieve a valid authentication token for the user."
                        });
                }
            }
        }

        /// <summary>
        /// Adds the current access token to the request headers. This method will silently refresh the access
        /// token, if needed and a refresh token is present, but will not prompt authentication UI to the user.
        /// Throws an exception if the access token has expired and cannot be refreshed.
        /// </summary>
        /// <param name="request">The <see cref="HttpRequestMessage"/> to authenticate.</param>
        /// <returns>The task to await.</returns>
        public async Task AuthenticateRequestAsync(HttpRequestMessage request)
        {
            if (!string.IsNullOrEmpty(this.accessToken) && !(this.expiration <= DateTimeOffset.Now.UtcDateTime.AddMinutes(5)))
            {
                request.Headers.Authorization = new AuthenticationHeaderValue("bearer", this.accessToken);
            }
            else
            {
                this.accessToken = null;

                await this.RefreshAccessTokenAsync();

                if (!string.IsNullOrEmpty(this.accessToken))
                {
                    request.Headers.Authorization = new AuthenticationHeaderValue("bearer", this.accessToken);
                }
                else
                {
                    throw new ServiceException(
                        new Error
                        {
                            Code = "authenticationRequired",
                            Message = "Please call AuthenticateAsync to prompt the user for authentication.",
                        });
                }
            }
        }

        private async Task GetAuthenticationResultAsync()
        {
            // Log the user in if we haven't already pulled their credentials from the cache.
            var code = await this.GetAuthorizationCodeAsync(returnUrl);

            if (!string.IsNullOrEmpty(code))
            {
                await this.SendTokenRequestAsync(this.GetCodeRedemptionRequestBody(code));
            }
        }

        private async Task<string> GetAuthorizationCodeAsync(string returnUrl = null)
        {
            var requestUri = new Uri(this.GetAuthorizationCodeRequestUrl(returnUrl));

            var result = await WebAuthenticationBroker.AuthenticateAsync(WebAuthenticationOptions.None, requestUri, new Uri(this.returnUrl));

            IDictionary<string, string> authenticationResponseValues = null;
            if (result != null && !string.IsNullOrEmpty(result.ResponseData))
            {
                authenticationResponseValues = UrlHelper.GetQueryOptions(new Uri(result.ResponseData));

                this.ThrowIfError(authenticationResponseValues);
            }
            else if (result != null && result.ResponseStatus == WebAuthenticationStatus.UserCancel)
            {
                throw new ServiceException(new Error { Code = GraphErrorCode.AuthenticationCancelled.ToString() });
            }
            
            string code;
            if (authenticationResponseValues != null && authenticationResponseValues.TryGetValue("code", out code))
            {
                return code;
            }

            return null;
        }

        /// <summary>
        /// Refresh the current access token, if possible.
        /// </summary>
        /// <returns>The task to await.</returns>
        private Task RefreshAccessTokenAsync()
        {
            if (!string.IsNullOrEmpty(this.refreshToken))
            {
                return this.SendTokenRequestAsync(this.GetRefreshTokenRequestBody(refreshToken));
            }

            return Task.FromResult(0);
        }

        private async Task SendTokenRequestAsync(string requestBodyString)
        {
            var httpRequestMessage = new HttpRequestMessage(HttpMethod.Post, OAuth2AuthenticationProvider.tokenServiceUrl);

            httpRequestMessage.Content = new StringContent(requestBodyString, Encoding.UTF8, Constants.Headers.FormUrlEncodedContentType);

            using (var authResponse = await this.httpProvider.SendAsync(httpRequestMessage))
            using (var responseStream = await authResponse.Content.ReadAsStreamAsync())
            {
                var responseValues =
                    this.httpProvider.Serializer.DeserializeObject<IDictionary<string, string>>(
                        responseStream);

                if (responseValues != null)
                {
                    this.ThrowIfError(responseValues);

                    this.refreshToken = responseValues[Constants.Authentication.RefreshTokenKeyName];
                    this.accessToken = responseValues[Constants.Authentication.AccessTokenKeyName];
                    this.expiration = DateTimeOffset.UtcNow.Add(new TimeSpan(0, 0, int.Parse(responseValues[Constants.Authentication.ExpiresInKeyName])));
                }
                else
                {
                    throw new ServiceException(
                        new Error
                        {
                            Code = GraphErrorCode.AuthenticationFailure.ToString(),
                            Message = "Authentication failed. No response values returned from token authentication flow."
                        });
                }
            }
        }

        /// <summary>
        /// Gets the request URL for OAuth authentication using the code flow.
        /// </summary>
        /// <param name="returnUrl">The return URL for the request. Defaults to the service info value.</param>
        /// <returns>The OAuth request URL.</returns>
        private string GetAuthorizationCodeRequestUrl(string returnUrl = null)
        {
            var requestUriStringBuilder = new StringBuilder();
            requestUriStringBuilder.Append(OAuth2AuthenticationProvider.authenticationServiceUrl);
            requestUriStringBuilder.AppendFormat("?{0}={1}", Constants.Authentication.RedirectUriKeyName, returnUrl);
            requestUriStringBuilder.AppendFormat("&{0}={1}", Constants.Authentication.ClientIdKeyName, this.clientId);
            requestUriStringBuilder.AppendFormat("&{0}={1}", Constants.Authentication.ResponseTypeKeyName, Constants.Authentication.CodeKeyName);
            requestUriStringBuilder.AppendFormat("&{0}={1}", Constants.Authentication.ScopeKeyName, WebUtility.UrlEncode(string.Join(" ", this.scopes)));

            return requestUriStringBuilder.ToString();
        }

        /// <summary>
        /// Gets the request body for redeeming an authorization code for an access token.
        /// </summary>
        /// <param name="code">The authorization code to redeem.</param>
        /// <param name="returnUrl">The return URL for the request. Defaults to the service info value.</param>
        /// <returns>The request body for the code redemption call.</returns>
        private string GetCodeRedemptionRequestBody(string code)
        {
            var requestBodyStringBuilder = new StringBuilder();
            requestBodyStringBuilder.AppendFormat("{0}={1}", Constants.Authentication.RedirectUriKeyName, this.returnUrl);
            requestBodyStringBuilder.AppendFormat("&{0}={1}", Constants.Authentication.ClientIdKeyName, this.clientId);
            requestBodyStringBuilder.AppendFormat("&{0}={1}", Constants.Authentication.CodeKeyName, code);
            requestBodyStringBuilder.AppendFormat("&{0}={1}", Constants.Authentication.GrantTypeKeyName, Constants.Authentication.AuthorizationCodeGrantType);
            requestBodyStringBuilder.AppendFormat("&{0}={1}", Constants.Authentication.ScopeKeyName, WebUtility.UrlEncode(string.Join(" ", this.scopes)));

            return requestBodyStringBuilder.ToString();
        }

        /// <summary>
        /// Gets the request body for redeeming a refresh token for an access token.
        /// </summary>
        /// <param name="refreshToken">The refresh token to redeem.</param>
        /// <returns>The request body for the redemption call.</returns>
        private string GetRefreshTokenRequestBody(string refreshToken)
        {
            var requestBodyStringBuilder = new StringBuilder();
            requestBodyStringBuilder.AppendFormat("{0}={1}", Constants.Authentication.RedirectUriKeyName, this.returnUrl);
            requestBodyStringBuilder.AppendFormat("&{0}={1}", Constants.Authentication.ClientIdKeyName, this.clientId);
            requestBodyStringBuilder.AppendFormat("&{0}={1}", Constants.Authentication.RefreshTokenKeyName, refreshToken);
            requestBodyStringBuilder.AppendFormat("&{0}={1}", Constants.Authentication.GrantTypeKeyName, Constants.Authentication.RefreshTokenKeyName);
            requestBodyStringBuilder.AppendFormat("&{0}={1}", Constants.Authentication.ScopeKeyName, WebUtility.UrlEncode(string.Join(" ", this.scopes)));

            return requestBodyStringBuilder.ToString();
        }

        private void ThrowIfError(IDictionary<string, string> responseValues)
        {
            if (responseValues != null)
            {
                string error = null;
                string errorDescription = null;

                if (responseValues.TryGetValue(Constants.Authentication.ErrorDescriptionKeyName, out errorDescription) ||
                    responseValues.TryGetValue(Constants.Authentication.ErrorKeyName, out error))
                {
                    this.ParseAuthenticationError(error, errorDescription);
                }
            }
        }

        private void ParseAuthenticationError(string error, string errorDescription)
        {
            throw new ServiceException(
                new Error
                {
                    Code = GraphErrorCode.AuthenticationFailure.ToString(),
                    Message = errorDescription ?? error
                });
        }
    }
}

