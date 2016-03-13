




using System.Web.Http;

namespace WebApplication1
{
    public static class WebHookConfig
    {
        public static void Register(HttpConfiguration config)
        {

			config.InitializeReceiveGitHubWebHooks();

        }
    }
}
