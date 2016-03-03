using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Graph;

namespace GraphFilesWeb
{
    public class FileRepository
    {
        private readonly string graphAccessToken;
        private readonly GraphServiceClient graphClient;

        public FileRepository(string accessToken)
        {
            graphAccessToken = GraphHelper.GetGraphAccessToken();
            graphClient = new GraphServiceClient(new DelegateAuthenticationProvider(async (System.Net.Http.HttpRequestMessage request) =>
            {
                request.Headers.TryAddWithoutValidation("Authorization", "Bearer " + graphAccessToken);
            }));
        }

        public async Task<IChildrenCollectionPage> GetMyFilesAsync(int pageSize)
        {
            // build the request for items in the root folder
            var request = graphClient.Me.Drive.Root.Children.Request().Top(pageSize);

            // get there results of the request
            var results = await request.GetAsync();
            return results;
        }

        public async Task<bool> DeleteItemAsync(string id, string etag)
        {
            // create request to delete the item
            var request = graphClient.Me.Drive.Items[id].Request();

            // Execute the delete action on this request
            try
            {
                await request.DeleteAsync();
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public async Task<DriveItem> UploadFileAsync(System.IO.Stream filestream, string filename)
        {
            // Create a request to upload the file using simple PUT action
            var request = graphClient.Me.Drive.Root.ItemWithPath(filename).Content.Request();

            // Submit the request with the contents of the filestream
            var result = await request.PutAsync<DriveItem>(filestream);
            return result;
        }


    }
}
