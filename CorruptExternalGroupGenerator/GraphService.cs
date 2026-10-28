using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.External.Connections.Item;
using Microsoft.Graph.Models.ExternalConnectors;
using Microsoft.Graph.Models.ODataErrors;

namespace CorruptExternalGroupGenerator
{
    internal class GraphService(ClientSecretCredential Credentials, string ExternalConnectionId)
    {
        private GraphServiceClient Client() => new(Credentials);
        private ExternalConnectionItemRequestBuilder ExternalConnection(GraphServiceClient client)
            => client.External.Connections[ExternalConnectionId];

        public async Task<ExternalConnection?> GetConnectionAsync(CancellationToken cancellationToken)
        {
            ExternalConnection? connection = null;
            using var client = Client();
            try
            {
                connection = await ExternalConnection(client).GetAsync(null, cancellationToken);
            }
            catch (ODataError odex) when (odex.ResponseStatusCode == 404)
            {
                // doesn't exist
            }
            return connection;
        }

        public async Task<Schema?> SetSchemaAsync(Schema schema, CancellationToken cancellationToken)
        {
            using var client = Client();
            return await ExternalConnection(client).Schema.PatchAsync(schema, cancellationToken: cancellationToken);
        }
        public async Task<Schema?> GetSchemaAsync(CancellationToken cancellationToken)
        {
            using var client = Client();
            return await ExternalConnection(client).Schema.GetAsync(cancellationToken: cancellationToken);
        }

        public async Task<bool> CreateConnectionAsync(CancellationToken cancellationToken)
        {
            using var client = Client();

            await client.External.Connections.PostAsync(new ExternalConnection
            {
                Id = ExternalConnectionId,
                Name = "Testing External Groups",
                Description = "Testing External Groups"
            }, cancellationToken: cancellationToken);

            return true;
        }

        public async Task<ExternalGroup?> GetExternalGroupAsync(string groupId, CancellationToken cancellationToken)
        {
            ExternalGroup? group = null;
            using var client = Client();
            try
            {
                group = await ExternalConnection(client).Groups[groupId].GetAsync(cancellationToken: cancellationToken);
            }
            catch(ODataError odex) when (odex.ResponseStatusCode == 404)
            {
                // not found
            }
            return group;
        }
        public async Task<ExternalGroup?> CreateExternalGroupAsync(string groupId, CancellationToken cancellationToken)
        {
            using var client = Client();

            return await ExternalConnection(client).Groups.PostAsync(
                body: new ExternalGroup
                {
                    Id = groupId,
                    Description = groupId,
                    DisplayName = groupId
                },
                cancellationToken: cancellationToken);
        }
        public async Task DeleteExternalGroupAsync(string groupId, CancellationToken cancellationToken)
        {
            using var client = Client();

            try
            {
                await ExternalConnection(client).Groups[groupId].DeleteAsync(cancellationToken: cancellationToken);
            }
            catch (ODataError odex) when (odex.ResponseStatusCode == 404)
            {
                // not found
            }
        }
    }
}
