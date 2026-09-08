using SlimLib.Auth.Azure;
using System;
using System.Net.Http;
using System.Runtime.CompilerServices;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Threading;
using System.Threading.Tasks;

namespace SlimLib.Microsoft.Graph
{
    partial class SlimGraphClientImpl
    {
        GraphOperation<JsonDocument?> ISlimGraphMobileAppsClient.GetMobileAppAsync(IAzureTenant tenant, Guid appID, ScalarRequestOptions? options, CancellationToken cancellationToken)
        {
            var link = ODataLinkBuilder.BuildLink(options, $"deviceAppManagement/mobileApps/{appID}");

            return new(this, tenant, HttpMethod.Get, link, options, default, static doc => doc);
        }

        GraphOperation<JsonDocument?> ISlimGraphMobileAppsClient.CreateMobileAppAsync(IAzureTenant tenant, JsonObject data, InvokeRequestOptions? options, CancellationToken cancellationToken)
        {
            var link = ODataLinkBuilder.BuildLink(options, "deviceAppManagement/mobileApps");

            return new(this, tenant, HttpMethod.Post, link, options, JsonSerializer.SerializeToUtf8Bytes(data), static doc => doc);
        }

        GraphOperation<JsonDocument?> ISlimGraphMobileAppsClient.UpdateMobileAppAsync(IAzureTenant tenant, Guid appID, JsonObject data, InvokeRequestOptions? options, CancellationToken cancellationToken)
        {
            var link = ODataLinkBuilder.BuildLink(options, $"deviceAppManagement/mobileApps/{appID}");

            return new(this, tenant, HttpMethod.Patch, link, options, JsonSerializer.SerializeToUtf8Bytes(data), static doc => doc);
        }

        GraphOperation ISlimGraphMobileAppsClient.DeleteMobileAppAsync(IAzureTenant tenant, Guid appID, InvokeRequestOptions? options, CancellationToken cancellationToken)
        {
            var link = ODataLinkBuilder.BuildLink(options, $"deviceAppManagement/mobileApps/{appID}");

            return new(this, tenant, HttpMethod.Delete, link, options, default);
        }


        GraphArrayOperation<JsonDocument> ISlimGraphMobileAppsClient.GetMobileAppsAsync(IAzureTenant tenant, ListRequestOptions? options, CancellationToken cancellationToken)
        {
            var nextLink = ODataLinkBuilder.BuildLink(options, "deviceAppManagement/mobileApps");

            return new(this, tenant, HttpMethod.Get, nextLink, options, default, static doc => doc);
        }

        GraphArrayOperation<JsonDocument> ISlimGraphMobileAppsClient.GetMobileAppAssignmentsAsync(IAzureTenant tenant, Guid appID, ListRequestOptions? options, CancellationToken cancellationToken)
        {
            var nextLink = ODataLinkBuilder.BuildLink(options, $"deviceAppManagement/mobileApps/{appID}/assignments");

            return new(this, tenant, HttpMethod.Get, nextLink, options, default, static doc => doc);
        }

        GraphOperation<JsonDocument?> ISlimGraphMobileAppsClient.GetMobileAppAssignmentAsync(IAzureTenant tenant, Guid appID, string assignmentID, ScalarRequestOptions? options, CancellationToken cancellationToken)
        {
            var link = ODataLinkBuilder.BuildLink(options, $"deviceAppManagement/mobileApps/{appID}/assignments/{assignmentID}");

            return new(this, tenant, HttpMethod.Get, link, options, default, static doc => doc);
        }

        GraphOperation<JsonDocument?> ISlimGraphMobileAppsClient.CreateMobileAppAssignmentAsync(IAzureTenant tenant, Guid appID, JsonObject data, InvokeRequestOptions? options, CancellationToken cancellationToken)
        {
            var link = ODataLinkBuilder.BuildLink(options, $"deviceAppManagement/mobileApps/{appID}/assignments");

            return new(this, tenant, HttpMethod.Post, link, options, JsonSerializer.SerializeToUtf8Bytes(data), static doc => doc);
        }

        GraphOperation<JsonDocument?> ISlimGraphMobileAppsClient.UpdateMobileAppAssignmentAsync(IAzureTenant tenant, Guid appID, string assignmentID, JsonObject data, InvokeRequestOptions? options, CancellationToken cancellationToken)
        {
            var link = ODataLinkBuilder.BuildLink(options, $"deviceAppManagement/mobileApps/{appID}/assignments/{assignmentID}");

            return new(this, tenant, HttpMethod.Patch, link, options, JsonSerializer.SerializeToUtf8Bytes(data), static doc => doc);
        }

        GraphOperation ISlimGraphMobileAppsClient.DeleteMobileAppAssignmentAsync(IAzureTenant tenant, Guid appID, string assignmentID, InvokeRequestOptions? options, CancellationToken cancellationToken)
        {
            var link = ODataLinkBuilder.BuildLink(options, $"deviceAppManagement/mobileApps/{appID}/assignments/{assignmentID}");

            return new(this, tenant, HttpMethod.Delete, link, options, default);
        }


        GraphOperation<JsonDocument?> ISlimGraphMobileAppsClient.GetMobileAppContentAsync(IAzureTenant tenant, Guid appID, string type, string mobileAppContentID, ScalarRequestOptions? options, CancellationToken cancellationToken)
        {
            var link = ODataLinkBuilder.BuildLink(options, $"deviceAppManagement/mobileApps/{appID}/{type}/contentVersions/{mobileAppContentID}");

            return new(this, tenant, HttpMethod.Get, link, options, default, static doc => doc);
        }

        async Task<string> ISlimGraphMobileAppsClient.CreateMobileAppContentAsync(IAzureTenant tenant, Guid appID, string type, InvokeRequestOptions? options, CancellationToken cancellationToken)
        {
            var link = ODataLinkBuilder.BuildLink(options, $"deviceAppManagement/mobileApps/{appID}/{type}/contentVersions");

            using var response = await PostAsync(tenant, link, new byte[] { 0x7B, 0x7D }, options, cancellationToken).ConfigureAwait(false);
            return response?.RootElement.GetProperty("id").GetString() ?? "";
        }

        async Task ISlimGraphMobileAppsClient.CommitMobileAppContentAsync(IAzureTenant tenant, Guid appID, string type, string mobileAppContentID, InvokeRequestOptions? options, CancellationToken cancellationToken)
        {
            var data = new JsonObject
            {
                ["@odata.type"] = "#" + type,
                ["committedContentVersion"] = mobileAppContentID,
            };

            await ((ISlimGraphMobileAppsClient)this).UpdateMobileAppAsync(tenant, appID, data, cancellationToken: cancellationToken);
        }


        GraphOperation<JsonDocument?> ISlimGraphMobileAppsClient.GetMobileAppContentFilesAsync(IAzureTenant tenant, Guid appID, string type, string mobileAppContentID, Guid mobileAppContentFileID, ScalarRequestOptions? options, CancellationToken cancellationToken)
        {
            var link = ODataLinkBuilder.BuildLink(options, $"deviceAppManagement/mobileApps/{appID}/{type}/contentVersions/{mobileAppContentID}/files/{mobileAppContentFileID}");

            return new(this, tenant, HttpMethod.Get, link, options, default, static doc => doc);
        }

        async Task<JsonDocument?> ISlimGraphMobileAppsClient.CreateMobileAppContentFilesAsync(IAzureTenant tenant, Guid appID, string type, string mobileAppContentID, JsonObject data, InvokeRequestOptions? options, CancellationToken cancellationToken)
        {
            var link = ODataLinkBuilder.BuildLink(options, $"deviceAppManagement/mobileApps/{appID}/{type}/contentVersions/{mobileAppContentID}/files");

            return await PostAsync(tenant, link, JsonSerializer.SerializeToUtf8Bytes(data), options, cancellationToken).ConfigureAwait(false);
        }

        async Task ISlimGraphMobileAppsClient.CommitMobileAppContentFilesAsync(IAzureTenant tenant, Guid appID, string type, string mobileAppContentID, Guid mobileAppContentFileID, JsonObject data, InvokeRequestOptions? options, CancellationToken cancellationToken)
        {
            var link = ODataLinkBuilder.BuildLink(options, $"deviceAppManagement/mobileApps/{appID}/{type}/contentVersions/{mobileAppContentID}/files/{mobileAppContentFileID}/commit");

            using var doc = await PostAsync(tenant, link, JsonSerializer.SerializeToUtf8Bytes(data), options, cancellationToken).ConfigureAwait(false);
        }

        GraphOperation ISlimGraphMobileAppsClient.AssignMobileAppAsync(IAzureTenant tenant, Guid appID, JsonObject data, InvokeRequestOptions? options, CancellationToken cancellationToken)
        {
            var link = ODataLinkBuilder.BuildLink(options, $"deviceAppManagement/mobileApps/{appID}/assign");

            return new(this, tenant, HttpMethod.Post, link, options, JsonSerializer.SerializeToUtf8Bytes(data));
        }

        GraphOperation<JsonDocument?> ISlimGraphMobileAppsClient.GetMobileAppCategoryAsync(IAzureTenant tenant, Guid appCategoryID, ScalarRequestOptions? options, CancellationToken cancellationToken)
        {
            var link = ODataLinkBuilder.BuildLink(options, $"deviceAppManagement/mobileAppCategories/{appCategoryID}");
            return new(this, tenant, HttpMethod.Get, link, options, default, static doc => doc);
        }

        GraphOperation<JsonDocument?> ISlimGraphMobileAppsClient.CreateMobileAppCategoryAsync(IAzureTenant tenant, JsonObject data, InvokeRequestOptions? options, CancellationToken cancellationToken)
        {
            var link = ODataLinkBuilder.BuildLink(options, "deviceAppManagement/mobileAppCategories");
            return new(this, tenant, HttpMethod.Post, link, options, JsonSerializer.SerializeToUtf8Bytes(data), static doc => doc);
        }

        GraphOperation<JsonDocument?> ISlimGraphMobileAppsClient.UpdateMobileAppCategoryAsync(IAzureTenant tenant, Guid appCategoryID, JsonObject data, InvokeRequestOptions? options, CancellationToken cancellationToken)
        {
            var link = ODataLinkBuilder.BuildLink(options, $"deviceAppManagement/mobileAppCategories/{appCategoryID}");
            return new(this, tenant, HttpMethod.Patch, link, options, JsonSerializer.SerializeToUtf8Bytes(data), static doc => doc);
        }

        GraphOperation ISlimGraphMobileAppsClient.DeleteMobileAppCategoryAsync(IAzureTenant tenant, Guid appCategoryID, InvokeRequestOptions? options, CancellationToken cancellationToken)
        {
            var link = ODataLinkBuilder.BuildLink(options, $"deviceAppManagement/mobileAppCategories/{appCategoryID}");
            return new(this, tenant, HttpMethod.Delete, link, options, default);
        }

        GraphArrayOperation<JsonDocument> ISlimGraphMobileAppsClient.GetMobileAppCategoriesAsync(IAzureTenant tenant, ListRequestOptions? options, CancellationToken cancellationToken)
        {
            var nextLink = ODataLinkBuilder.BuildLink(options, "deviceAppManagement/mobileAppCategories");
            return new(this, tenant, HttpMethod.Get, nextLink, options, default, static doc => doc);
        }

        GraphOperation ISlimGraphMobileAppsClient.AssignCategoryToMobileAppAsync(IAzureTenant tenant, Guid appID, JsonObject data, InvokeRequestOptions? options, CancellationToken cancellationToken)
        {
            var link = ODataLinkBuilder.BuildLink(options, $"deviceAppManagement/mobileApps/{appID}/categories/$ref");
            return new(this, tenant, HttpMethod.Post, link, options, JsonSerializer.SerializeToUtf8Bytes(data));
        }

        GraphOperation ISlimGraphMobileAppsClient.RemoveCategoryFromMobileAppAsync(IAzureTenant tenant, Guid appID, Guid appCategoryID, InvokeRequestOptions? options, CancellationToken cancellationToken)
        {
            var link = ODataLinkBuilder.BuildLink(options, $"deviceAppManagement/mobileApps/{appID}/categories/{appCategoryID}/$ref");
            return new(this, tenant, HttpMethod.Delete, link, options, default);
        }
    }
}