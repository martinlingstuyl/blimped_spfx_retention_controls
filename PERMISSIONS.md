# Retention Controls Permissions Configuration

## Overview

By default, users with SharePoint edit permissions on the site can clear retention labels and toggle record status. However, for the retention controls extension, you can configure additional permissions to restrict these actions to specific Entra ID groups or users.
The Retention Controls extension has a simple permissions system that performs two permission checks:

1. **SharePoint Site Edit Permissions**: Checks if the user has native SharePoint edit permissions on the site
2. **Custom Permissions**: Checks against configured Entra ID groups or specific users

Both checks must pass for a user to have edit capabilities.

If no permissions are configured, the default behavior applies. 

## Caching Strategy

- Permissions are cached in browser `sessionStorage` using `RetentionControls_<siteUrl>` as the key
- Cache is automatically cleared when the browser session ends
- Since permissions are checked at the site level, they apply to all libraries within the site, regardless of broken permission inheritance. 

> If you need a fresh permissions check after making changes, you can simply start a new browser tab or clear the sessionStorage through the browsers dev tools.

## Configuration

To configure permissions, update your SharePoint Framework extension properties in the app catalog or tenant store. The permissions configuration should be added to the `RetentionControlsCommandSet` properties.

### Example Configuration

```json
{
  "permissions": {
    "entries": [
      {
        "groupId": "12345678-1234-1234-1234-123456789012"
      },
      {
        "userName": "admin@contoso.com"
      }
    ]
  }
}
```

## Microsoft Graph permissions in SharePoint Framework

Because the extension calls Microsoft Graph for retrieving user group memberships, the SharePoint Entra ID principal should have `User.Read` permissions granted if you want to use permission configurations. Visit the API access page in your SharePoint admin center to verify if the permissions have been granted.