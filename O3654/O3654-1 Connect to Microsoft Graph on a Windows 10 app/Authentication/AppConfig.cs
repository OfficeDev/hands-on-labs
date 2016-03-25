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
    public class AppConfig
    {
        public AppConfig()
        {
            this.AuthenticationServiceUrl = Constants.Authentication.ActiveDirectoryAuthenticationServiceUrl;
            this.TokenServiceUrl = Constants.Authentication.ActiveDirectoryTokenServiceUrl;
        }

        /// <summary>
        /// Gets or sets the client ID for authentication.
        /// </summary>
        public string ClientId { get; set; }

        /// <summary>
        /// Gets or sets the authentication service URL for authentication.
        /// Defaults to "https://login.microsoftonline.com/common/oauth2/v2.0/authorize".
        /// </summary>
        public string AuthenticationServiceUrl { get; set; }

        /// <summary>
        /// Gets or sets the application return URL for authentication.
        /// </summary>
        public string ReturnUrl { get; set; }

        /// <summary>
        /// Gets or sets the requested scopes for authentication.
        /// </summary>
        public string[] Scopes { get; set; }

        /// <summary>
        /// Gets or sets the base URL for the authentication token service.
        /// Defaults to "https://login.microsoftonline.com/common/oauth2/v2.0/token".
        /// </summary>
        public string TokenServiceUrl { get; set; }
    }
}
