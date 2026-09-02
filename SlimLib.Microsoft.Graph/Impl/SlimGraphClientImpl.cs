using Microsoft.Extensions.Logging;
using SlimLib.Auth.Azure;
using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace SlimLib.Microsoft.Graph
{
    internal sealed partial class SlimGraphClientImpl : SlimODataClientBase, ISlimGraphAdministrativeUnitsClient, ISlimGraphAndroidManagedStoreClient, ISlimGraphApplicationsClient, ISlimGraphAuditEventsClient, ISlimGraphAuditLogsClient, ISlimGraphDeviceManagementReportsClient, ISlimGraphOrganizationsClient, ISlimGraphOrgContactsClient, ISlimGraphDevicesClient, ISlimGraphDirectoryRolesClient, ISlimGraphDetectedAppsClient, ISlimGraphMobileAppsClient, ISlimGraphManagedDevicesClient, ISlimGraphGroupsClient, ISlimGraphSubscribedSkusClient, ISlimGraphServicePrincipalsClient, ISlimGraphPrivilegedAccessClient, ISlimGraphUsersClient, ISlimGraphDeviceLocalCredentialsClient, ISlimGraphPartnerBillingReportsClient, ISlimGraphTenantRelationshipsClient, ISlimGraphBitLockerClient, ISlimGraphWindowsDeviceUpdatesClient
    {
        private readonly ILogger<SlimGraphClient> logger;

        public SlimGraphClientImpl(IAuthenticationProvider authenticationProvider, HttpClient httpClient, ILogger<SlimGraphClient> logger)
            : base(authenticationProvider, httpClient, logger)
        {
            this.logger = logger;
        }

        protected override string Scope => SlimGraphConstants.ScopeDefault;

#pragma warning disable CS0618 // SlimGraphException is intentionally thrown for backward compatibility

        protected override SlimApiException CreateApiError(HttpStatusCode statusCode, IEnumerable<KeyValuePair<string, IEnumerable<string>>> headers, string errorCode, string errorMessage)
            => new SlimGraphException(statusCode, headers, errorCode, errorMessage);

#pragma warning restore CS0618

        private async Task<Results.Delta.DeltaResult<JsonElement>> GetDeltaAsync(IAzureTenant tenant, string nextLink, DeltaRequestOptions? options, CancellationToken cancellationToken)
        {
            var result = new List<JsonElement>();

            string? nLink = nextLink;
            string? dLink = default;

            do
            {
                using var doc = await GetAsync(tenant, nLink, options, cancellationToken).ConfigureAwait(false);

                if (doc is not null)
                {
                    if (doc.RootElement.TryGetProperty("value", out var items) && items.ValueKind == JsonValueKind.Array)
                    {
                        foreach (var item in items.EnumerateArray())
                        {
                            if (cancellationToken.IsCancellationRequested)
                                break;

                            result.Add(item);
                        }
                    }

                    if (cancellationToken.IsCancellationRequested)
                        break;

                    HandleNextLink(doc.RootElement, ref nLink);
                    HandleDeltaLink(doc.RootElement, ref dLink);
                }

            } while (nLink != null);

            return new Results.Delta.DeltaResult<JsonElement>(result, dLink);
        }

        private async Task<SlimGraphPicture?> GetPictureAsync(IAzureTenant tenant, string requestUri, CancellationToken cancellationToken)
        {
            using var response = await SendInternalAsync(tenant, HttpMethod.Get, requestUri, null, null, cancellationToken);

            if (response.StatusCode == HttpStatusCode.NoContent)
            {
                logger.LogInformation("Got no content for HTTP request to {requestUri}.", requestUri);
                return null;
            }

            var buffer = await response.Content.ReadAsByteArrayAsync().ConfigureAwait(false);

            return new SlimGraphPicture(buffer, response.Content.Headers.ContentType);
        }

        private static void HandleDeltaLink(JsonElement root, ref string? deltaLink)
        {
            if (root.TryGetProperty("@odata.deltaLink", out var el))
            {
                deltaLink = el.GetString();
            }
            else
            {
                deltaLink = null;
            }
        }
    }
}