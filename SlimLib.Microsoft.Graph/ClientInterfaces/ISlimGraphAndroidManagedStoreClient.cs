using SlimLib.Auth.Azure;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Threading;

namespace SlimLib.Microsoft.Graph
{
    public interface ISlimGraphAndroidManagedStoreClient
    {
        GraphOperation<JsonDocument?> GetEnterpriseSettingsAsync(IAzureTenant tenant, ScalarRequestOptions? options = default, CancellationToken cancellationToken = default);
        GraphOperation<JsonDocument?> CreateGooglePlayWebTokenAsync(IAzureTenant tenant, JsonObject data, InvokeRequestOptions? options = default, CancellationToken cancellationToken = default);
        GraphOperation AddAppsAsync(IAzureTenant tenant, JsonObject data, InvokeRequestOptions? options = default, CancellationToken cancellationToken = default);
        GraphOperation SyncAppsAsync(IAzureTenant tenant, InvokeRequestOptions? options = default, CancellationToken cancellationToken = default);
    }
}
