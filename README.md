# GCX Sensitivity Settings

## Summary

Apply permissions based on the sensitivity level of the site:
- Apply the appropriate sensitivity label
- Update the site collection administrator group
- Give full control access the to the support group
- Give read-only access to the read-only group(s)
- Remove the temporary owner that was used to perform actions that required elevated access
  
## Prerequisites

The following user accounts (as reflected in the app settings) are required:

| Account             | Membership requirements                                  |
| ------------------- | -------------------------------------------------------- |
| user_name           | |

## Version 

![dotnet 6](https://img.shields.io/badge/net6.0-blue.svg)

## API permission

MSGraph

| API / Permissions name    | Type        | Admin consent | Justification                       |
| ------------------------- | ----------- | ------------- | ----------------------------------- |
| Group.ReadWrite.All       | Delegated   | Yes           | Apply Label code                    | 

Sharepoint

| API / Permissions name    | Type      | Admin consent | Justification                          |
| ------------------------- | --------- | ------------- | -------------------------------------- |
| AllSites.FullControl      | Delegated | Yes           | Add users to the Site Admin Collection |

## App setting

| Name                     | Description                                                                   					          |
| ------------------------ | ------------------------------------------------------------------------------------------------ |
| AzureWebJobsStorage      | Connection string for the storage acoount                                     					          |
| clientId                 | The application (client) ID of the app registration                           					          |
| keyVaultUrl              | Address for the key vault                                                     					          |
| ownerId				           | Id of  the service account to add as temporary owner in order to authorize delegated permissions |
| proBLabelId              | Id of the "Protected B sensitivity" label                                                        |
| readOnlyGroup            | Login name of groups that are provided read-only access                                          |
| sca_login_name           | Login name of the group for site collection administrator access for unclassified sites          |
| sca_prob_login_name      | Login name of the group for site collection administrator access for protected sites             |
| secretName               | Secret name used to authorize the function app                                					          |
| secretNamePassword       | The secret name for the delegated user (user_name) password                                      |
| sharePointUrl			       | The base url under which new sites will be created 	    						                            |
| support_group_login_name | Login name of the support group for elevated access                                              |
| tenantId                 | Id of the Azure tenant that hosts the function app                            					          |
| tenantName			         | Name of the tenant that hosts the function app                                					          |
| unclassifiedLabelId      | Id of the "Unclassified" sensitivity label                                                       |
| user_name                | Delegated authentication user for applying sensitivity label and removing owner (ownerId)        |

## Version history

Version|Date|Comments
-------|----|--------
1.0|2023-10-10|Initial release

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**
