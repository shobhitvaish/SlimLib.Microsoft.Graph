using SlimLib.Auth.Azure;
using System;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Threading;
using System.Threading.Tasks;

namespace SlimLib.Microsoft.Graph
{
    public interface ISlimGraphMobileAppsClient
    {
        GraphOperation<JsonDocument?> GetMobileAppAsync(IAzureTenant tenant, Guid appID, ScalarRequestOptions? options = default, CancellationToken cancellationToken = default);
        GraphOperation<JsonDocument?> CreateMobileAppAsync(IAzureTenant tenant, JsonObject data, InvokeRequestOptions? options = default, CancellationToken cancellationToken = default);
        GraphOperation<JsonDocument?> UpdateMobileAppAsync(IAzureTenant tenant, Guid appID, JsonObject data, InvokeRequestOptions? options = default, CancellationToken cancellationToken = default);
        GraphOperation DeleteMobileAppAsync(IAzureTenant tenant, Guid appID, InvokeRequestOptions? options = default, CancellationToken cancellationToken = default);

        GraphArrayOperation<JsonDocument> GetMobileAppsAsync(IAzureTenant tenant, ListRequestOptions? options = default, CancellationToken cancellationToken = default);
        GraphArrayOperation<JsonDocument> GetMobileAppAssignmentsAsync(IAzureTenant tenant, Guid appID, ListRequestOptions? options = default, CancellationToken cancellationToken = default);

        GraphOperation<JsonDocument?> GetMobileAppAssignmentAsync(IAzureTenant tenant, Guid appID, string assignmentID, ScalarRequestOptions? options = default, CancellationToken cancellationToken = default);
        GraphOperation<JsonDocument?> CreateMobileAppAssignmentAsync(IAzureTenant tenant, Guid appID, JsonObject data, InvokeRequestOptions? options = default, CancellationToken cancellationToken = default);
        GraphOperation<JsonDocument?> UpdateMobileAppAssignmentAsync(IAzureTenant tenant, Guid appID, string assignmentID, JsonObject data, InvokeRequestOptions? options = default, CancellationToken cancellationToken = default);
        GraphOperation DeleteMobileAppAssignmentAsync(IAzureTenant tenant, Guid appID, string assignmentID, InvokeRequestOptions? options = default, CancellationToken cancellationToken = default);

        GraphOperation<JsonDocument?> GetMobileAppContentAsync(IAzureTenant tenant, Guid appID, string type, string mobileAppContentID, ScalarRequestOptions? options = default, CancellationToken cancellationToken = default);
        Task<string> CreateMobileAppContentAsync(IAzureTenant tenant, Guid appID, string type, InvokeRequestOptions? options = default, CancellationToken cancellationToken = default);
        Task CommitMobileAppContentAsync(IAzureTenant tenant, Guid appID, string type, string mobileAppContentID, InvokeRequestOptions? options = default, CancellationToken cancellationToken = default);

        GraphOperation<JsonDocument?> GetMobileAppContentFilesAsync(IAzureTenant tenant, Guid appID, string type, string mobileAppContentID, Guid mobileAppContentFileID, ScalarRequestOptions? options = default, CancellationToken cancellationToken = default);
        Task<JsonDocument?> CreateMobileAppContentFilesAsync(IAzureTenant tenant, Guid appID, string type, string mobileAppContentID, JsonObject data, InvokeRequestOptions? options = default, CancellationToken cancellationToken = default);
        Task CommitMobileAppContentFilesAsync(IAzureTenant tenant, Guid appID, string type, string mobileAppContentID, Guid mobileAppContentFileID, JsonObject data, InvokeRequestOptions? options = default, CancellationToken cancellationToken = default);

        GraphOperation AssignMobileAppAsync(IAzureTenant tenant, Guid appID, JsonObject data, InvokeRequestOptions? options = default, CancellationToken cancellationToken = default);

        GraphOperation<JsonDocument?> GetMobileAppCategoryAsync(IAzureTenant tenant, Guid appCategoryID, ScalarRequestOptions? options = default, CancellationToken cancellationToken = default);
        GraphOperation<JsonDocument?> CreateMobileAppCategoryAsync(IAzureTenant tenant, JsonObject data, InvokeRequestOptions? options = default, CancellationToken cancellationToken = default);
        GraphOperation<JsonDocument?> UpdateMobileAppCategoryAsync(IAzureTenant tenant, Guid appCategoryID, JsonObject data, InvokeRequestOptions? options = default, CancellationToken cancellationToken = default);
        GraphOperation DeleteMobileAppCategoryAsync(IAzureTenant tenant, Guid appCategoryID, InvokeRequestOptions? options = default, CancellationToken cancellationToken = default);
        GraphArrayOperation<JsonDocument> GetMobileAppCategoriesAsync(IAzureTenant tenant, ListRequestOptions? options = default, CancellationToken cancellationToken = default);

        GraphOperation AssignCategoryToMobileAppAsync(IAzureTenant tenant, Guid appID, JsonObject data, InvokeRequestOptions? options = default, CancellationToken cancellationToken = default);
        GraphOperation RemoveCategoryFromMobileAppAsync(IAzureTenant tenant, Guid appID, Guid appCategoryID, InvokeRequestOptions? options = default, CancellationToken cancellationToken = default);
    }
}