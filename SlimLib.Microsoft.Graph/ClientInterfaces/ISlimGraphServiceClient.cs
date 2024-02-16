﻿namespace SlimLib.Microsoft.Graph
{
    public interface ISlimGraphClient
    {
        ISlimGraphAdministrativeUnitsClient AdministrativeUnits { get; }
        ISlimGraphApplicationsClient Applications { get; }
        ISlimGraphAuditEventsClient AuditEvents { get; }
        ISlimGraphAuditLogsClient AuditLogs { get; }
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
        ISlimGraphDeviceLocalCredentialsClient DeviceLocalCredentials { get; }
    }
}