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
    using System.Net.Http;
    using System.Net.Http.Headers;
    using System.Text;
    using System.Threading.Tasks;
    using Windows.Security.Authentication.Web;

    public class OAuth2AuthenticationProvider : IAuthenticationProvider
    {
        private AppConfig appConfig;
        private IHttpProvider httpProvider;
        private IOAuthRequestStringBuilder oAuthRequestStringBuilder;
        private IWebAuthenticationUi webAuthenticationUi;

        /// <summary>
        /// Creates a new instance of <see cref="OAuth2AuthenticationProvider"/> for authentication in Windows Store applications.
        /// </summary>
        /// <param name="appConfig">The configuration details for authenticating the application.</param>
        /// <param name="httpProvider">The HTTP provider for sending HTTP requests.</param>
        public OAuth2AuthenticationProvider(AppConfig appConfig, IHttpProvider httpProvider = null)
        {
            if (appConfig == null)
            {
                throw new ServiceException(
                    new Error
                    {
                        Code = GraphErrorCode.InvalidRequest.ToString(),
                        Message = "AppConfig is required to authenticate using OAuth2AuthenticationProvider.",
                    });
            }

            if (string.IsNullOrEmpty(appConfig.ClientId))
            {
                throw new ServiceException(
                    new Error
                    {
                        Code = GraphErrorCode.InvalidRequest.ToString(),
                        Message = "Client ID is required to authenticate using OAuth2AuthenticationProvider.",
                    });
            }

            if (string.IsNullOrEmpty(appConfig.AuthenticationServiceUrl))
            {
                throw new ServiceException(
                    new Error
                    {
                        Code = GraphErrorCode.InvalidRequest.ToString(),
                        Message = "Authentication service URL is required to authenticate using OAuth2AuthenticationProvider.",
                    });
            }

            if (appConfig.Scopes == null)
            {
                throw new ServiceException(
                    new Error
                    {
                        Code = GraphErrorCode.InvalidRequest.ToString(),
                        Message = "No scopes have been requested for authentication.",
                    });
            }

            this.appConfig = appConfig;
            this.httpProvider = httpProvider ?? new HttpProvider();
        }

        /// <summary>
        /// Gets the current authenticated account session.
        /// </summary>
        public AccountSession AuthenticatedSession { get; internal set; }

        internal IOAuthRequestStringBuilder OAuthRequestStringBuilder
        {
            get
            {
                if (this.oAuthRequestStringBuilder == null)
                {
                    this.oAuthRequestStringBuilder = new OAuthRequestStringBuilder(this.appConfig);
                }

                return this.oAuthRequestStringBuilder;
            }

            set
            {
                this.oAuthRequestStringBuilder = value;
            }
        }

        internal IWebAuthenticationUi WebAuthenticationUi
        {
            get
            {
                if (this.webAuthenticationUi == null)
                {
                    this.webAuthenticationUi = new WebAuthenticationBrokerWebAuthenticationUi();
                }

                return this.webAuthenticationUi;
            }

            set
            {
                this.webAuthenticationUi = value;
            }
        }

        /// <summary>
        /// Presents authentication UI to the user.
        /// </summary>
        /// <returns>The task to await.</returns>
        public async Task AuthenticateAsync()
        {
            var authResult = await this.ProcessCachedAccountSessionAsync();

            if (authResult != null)
            {
                return;
            }

            authResult = await this.GetAuthenticationResultAsync();

            if (authResult == null || string.IsNullOrEmpty(authResult.AccessToken))
            {
                throw new ServiceException(
                    new Error
                    {
                        Code = GraphErrorCode.AuthenticationFailure.ToString(),
                        Message = "Failed to retrieve a valid authentication token for the user."
                    });
            }

            this.AuthenticatedSession = authResult;
        }

        /// <summary>
        /// Refreshes the current account session using the refresh token stored in AuthenticatedSession (if available).
        /// </summary>
        /// <returns>The task to await.</returns>
        public Task RedeemRefreshTokenAsync()
        {
            return this.RedeemRefreshTokenAsync(null);
        }

        /// <summary>
        /// Redeems the provided refresh token for an access token. If no refresh token is provided the refresh is
        /// performed using the refresh token stored in AuthenticatedSession (if available).
        /// </summary>
        /// <param name="refreshToken">The refresh token to redeem for an access token.</param>
        /// <returns>The task to await.</returns>
        public async Task RedeemRefreshTokenAsync(string refreshToken)
        {
            if (string.IsNullOrEmpty(refreshToken))
            {
                if (this.AuthenticatedSession == null || string.IsNullOrEmpty(this.AuthenticatedSession.RefreshToken))
                {
                    throw new ServiceException(
                        new Error
                        {
                            Code = GraphErrorCode.InvalidRequest.ToString(),
                            Message = "Refresh token is required to refresh authentication."
                        });
                }

                this.AuthenticatedSession = await this.RefreshAccessTokenAsync(this.AuthenticatedSession.RefreshToken);
            }
            else
            {
                this.AuthenticatedSession = await this.RefreshAccessTokenAsync(refreshToken);
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
            var authResult = await this.ProcessCachedAccountSessionAsync();

            if (authResult == null)
            {
                throw new ServiceException(
                    new Error
                    {
                        Code = GraphErrorCode.InvalidRequest.ToString(),
                        Message = "The current authentication token has expired and a new one cannot be silently retrieved. Please re-authenticate the user.",
                    });
            }

            if (!string.IsNullOrEmpty(authResult.AccessToken))
            {
                var tokenTypeString = string.IsNullOrEmpty(authResult.AccessTokenType)
                    ? Constants.Headers.Bearer
                    : authResult.AccessTokenType;
                request.Headers.Authorization = new AuthenticationHeaderValue(tokenTypeString, authResult.AccessToken);
            }
        }

        internal Task<AccountSession> GetAuthenticationResultAsync()
        {
            return this.GetAccountSessionAsync();
        }

        internal async Task<AccountSession> GetAccountSessionAsync()
        {
            var returnUrl = string.IsNullOrEmpty(this.appConfig.ReturnUrl)
                ? WebAuthenticationBroker.GetCurrentApplicationCallbackUri().ToString()
                : this.appConfig.ReturnUrl;

            // Log the user in if we haven't already pulled their credentials from the cache.
            var code = await this.GetAuthorizationCodeAsync(returnUrl);

            if (!string.IsNullOrEmpty(code))
            {
                var authResult = await this.SendTokenRequestAsync(this.OAuthRequestStringBuilder.GetCodeRedemptionRequestBody(code, returnUrl));

                return authResult;
            }

            return null;
        }

        internal async Task<string> GetAuthorizationCodeAsync(string returnUrl = null)
        {
            if (this.WebAuthenticationUi != null)
            {
                returnUrl = returnUrl ?? this.appConfig.ReturnUrl;

                var requestUri = new Uri(this.OAuthRequestStringBuilder.GetAuthorizationCodeRequestUrl(returnUrl));

                var authenticationResponseValues = await this.WebAuthenticationUi.AuthenticateAsync(
                    requestUri,
                    new Uri(returnUrl));
                OAuthErrorHandler.ThrowIfError(authenticationResponseValues);

                string code;
                if (authenticationResponseValues != null && authenticationResponseValues.TryGetValue("code", out code))
                {
                    return code;
                }
            }

            return null;
        }

        internal async Task<AccountSession> ProcessCachedAccountSessionAsync()
        {
            if (this.AuthenticatedSession != null)
            {
                // If we have cached credentials and they're not expiring, return them.
                if (!string.IsNullOrEmpty(this.AuthenticatedSession.AccessToken) && !this.AuthenticatedSession.IsExpiring())
                {
                    return this.AuthenticatedSession;
                }

                // If we don't have an access token or it's expiring, see if we can refresh the access token.
                if (!string.IsNullOrEmpty(this.AuthenticatedSession.RefreshToken))
                {
                    this.AuthenticatedSession = await this.RefreshAccessTokenAsync(this.AuthenticatedSession.RefreshToken);

                    if (this.AuthenticatedSession != null && !string.IsNullOrEmpty(this.AuthenticatedSession.AccessToken))
                    {
                        return this.AuthenticatedSession;
                    }
                }

                this.AuthenticatedSession = null;
            }

            return null;
        }

        internal Task<AccountSession> RefreshAccessTokenAsync(string refreshToken)
        {
            return this.SendTokenRequestAsync(this.OAuthRequestStringBuilder.GetRefreshTokenRequestBody(refreshToken));
        }

        internal async Task<AccountSession> SendTokenRequestAsync(string requestBodyString)
        {
            var httpRequestMessage = new HttpRequestMessage(HttpMethod.Post, this.appConfig.TokenServiceUrl);

            httpRequestMessage.Content = new StringContent(requestBodyString, Encoding.UTF8, Constants.Headers.FormUrlEncodedContentType);

            using (var authResponse = await this.httpProvider.SendAsync(httpRequestMessage))
            using (var responseStream = await authResponse.Content.ReadAsStreamAsync())
            {
                var responseValues =
                    this.httpProvider.Serializer.DeserializeObject<IDictionary<string, string>>(
                        responseStream);

                if (responseValues != null)
                {
                    OAuthErrorHandler.ThrowIfError(responseValues);
                    return new AccountSession(responseValues, this.appConfig.ClientId);
                }

                throw new ServiceException(
                    new Error
                    {
                        Code = GraphErrorCode.AuthenticationFailure.ToString(),
                        Message = "Authentication failed. No response values returned from token authentication flow."
                    });
            }
        }
    }
}
