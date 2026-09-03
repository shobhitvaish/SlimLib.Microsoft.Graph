using SlimLib.Auth.Azure;
using System.Net.Http;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Threading;

namespace SlimLib.Microsoft.Graph
{
    partial class SlimGraphClientImpl
    {
        GraphOperation<JsonDocument?> ISlimGraphAndroidManagedStoreClient.GetEnterpriseSettingsAsync(IAzureTenant tenant, ScalarRequestOptions? options, CancellationToken cancellationToken)
        {
            var link = ODataLinkBuilder.BuildLink(options, "deviceManagement/androidManagedStoreAccountEnterpriseSettings");

            return new(this, tenant, HttpMethod.Get, link, options, default, static doc => doc);
        }

        GraphOperation<JsonDocument?> ISlimGraphAndroidManagedStoreClient.CreateGooglePlayWebTokenAsync(IAzureTenant tenant, JsonObject data, InvokeRequestOptions? options, CancellationToken cancellationToken)
        {
            var link = ODataLinkBuilder.BuildLink(options, "deviceManagement/androidManagedStoreAccountEnterpriseSettings/createGooglePlayWebToken");

            return new(this, tenant, HttpMethod.Post, link, options, JsonSerializer.SerializeToUtf8Bytes(data), static doc => doc);
        }

        GraphOperation ISlimGraphAndroidManagedStoreClient.AddAppsAsync(IAzureTenant tenant, JsonObject data, InvokeRequestOptions? options, CancellationToken cancellationToken)
        {
            var link = ODataLinkBuilder.BuildLink(options, "deviceManagement/androidManagedStoreAccountEnterpriseSettings/addApps");

            return new(this, tenant, HttpMethod.Post, link, options, JsonSerializer.SerializeToUtf8Bytes(data));
        }

        GraphOperation ISlimGraphAndroidManagedStoreClient.SyncAppsAsync(IAzureTenant tenant, InvokeRequestOptions? options, CancellationToken cancellationToken)
        {
            var link = ODataLinkBuilder.BuildLink(options, "deviceManagement/androidManagedStoreAccountEnterpriseSettings/syncApps");

            return new(this, tenant, HttpMethod.Post, link, options, new byte[] { 0x7B, 0x7D });
        }
    }
}
