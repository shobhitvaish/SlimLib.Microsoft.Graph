using Microsoft.Extensions.Logging;
using SlimLib.Auth.Azure;
using System.Collections.Generic;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;

namespace SlimLib.Microsoft.Graph
{
    public class SlimGraphClient : ISlimGraphClient
    {
        private readonly SlimGraphClientImpl impl;

        public SlimGraphClient(IAuthenticationProvider authenticationProvider, HttpClient httpClient, ILogger<SlimGraphClient> logger)
        {
            impl = new SlimGraphClientImpl(authenticationProvider, httpClient, logger);
        }

        public ISlimGraphAdministrativeUnitsClient AdministrativeUnits => impl;
        public ISlimGraphAndroidManagedStoreClient AndroidManagedStore => impl;
        public ISlimGraphApplicationsClient Applications => impl;
        public ISlimGraphAuditEventsClient AuditEvents => impl;
        public ISlimGraphAuditLogsClient AuditLogs => impl;
        public ISlimGraphDeviceManagementReportsClient DeviceManagementReports => impl;
        public ISlimGraphDirectoryObjectsClient DirectoryObjects => impl;
        public ISlimGraphOrganizationsClient Organization => impl;
        public ISlimGraphOrgContactsClient OrgContacts => impl;
        public ISlimGraphDevicesClient Devices => impl;
        public ISlimGraphDirectoryRolesClient DirectoryRoles => impl;
        public ISlimGraphDetectedAppsClient DetectedApps => impl;
        public ISlimGraphMobileAppsClient MobileApps => impl;
        public ISlimGraphManagedDevicesClient ManagedDevices => impl;
        public ISlimGraphGroupsClient Groups => impl;
        public ISlimGraphServicePrincipalsClient ServicePrincipals => impl;
        public ISlimGraphSubscribedSkusClient SubscribedSkus => impl;
        public ISlimGraphPrivilegedAccessClient PrivilegedAccess => impl;
        public ISlimGraphUsersClient Users => impl;
        public ISlimGraphWindowsDeviceUpdatesClient WindowsDeviceUpdates => impl;
        public ISlimGraphDeviceLocalCredentialsClient DeviceLocalCredentials => impl;
        public ISlimGraphPartnerBillingReportsClient PartnerBillingReports => impl;
        public ISlimGraphTenantRelationshipsClient TenantRelationships => impl;
        public ISlimGraphBitLockerClient BitLocker => impl;

        public Task BatchRequestAsync(IAzureTenant tenant, IList<GraphOperation> operations, CancellationToken cancellationToken = default) => impl.BatchRequestAsync(tenant, operations, cancellationToken);
    }
}