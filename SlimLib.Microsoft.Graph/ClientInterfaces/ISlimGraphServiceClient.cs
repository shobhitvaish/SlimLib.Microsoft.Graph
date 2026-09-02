using SlimLib.Auth.Azure;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace SlimLib.Microsoft.Graph
{
    public interface ISlimGraphClient
    {
        ISlimGraphAdministrativeUnitsClient AdministrativeUnits { get; }
        ISlimGraphAndroidManagedStoreClient AndroidManagedStore { get; }
        ISlimGraphApplicationsClient Applications { get; }
        ISlimGraphAuditEventsClient AuditEvents { get; }
        ISlimGraphAuditLogsClient AuditLogs { get; }
        ISlimGraphDirectoryObjectsClient DirectoryObjects { get; }
        ISlimGraphOrganizationsClient Organization { get; }
        ISlimGraphOrgContactsClient OrgContacts { get; }
        ISlimGraphDevicesClient Devices { get; }
        ISlimGraphDirectoryRolesClient DirectoryRoles { get; }
        ISlimGraphDetectedAppsClient DetectedApps { get; }
        ISlimGraphMobileAppsClient MobileApps { get; }
        ISlimGraphManagedDevicesClient ManagedDevices { get; }
        ISlimGraphGroupsClient Groups { get; }
        ISlimGraphDeviceManagementReportsClient DeviceManagementReports { get; }
        ISlimGraphServicePrincipalsClient ServicePrincipals { get; }
        ISlimGraphSubscribedSkusClient SubscribedSkus { get; }
        ISlimGraphPrivilegedAccessClient PrivilegedAccess { get; }
        ISlimGraphUsersClient Users { get; }
        ISlimGraphWindowsDeviceUpdatesClient WindowsDeviceUpdates { get; }
        ISlimGraphDeviceLocalCredentialsClient DeviceLocalCredentials { get; }
        ISlimGraphPartnerBillingReportsClient PartnerBillingReports { get; }
        ISlimGraphTenantRelationshipsClient TenantRelationships { get; }
        ISlimGraphBitLockerClient BitLocker { get; }

        Task BatchRequestAsync(IAzureTenant tenant, IList<GraphOperation> operations, CancellationToken cancellationToken = default);
    }
}