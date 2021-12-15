# Graph API Email, Presence, MailboxSettings, Calendar client

[![.NET Core Desktop](https://github.com/damienbod/EmailCalandarsClient/actions/workflows/dotnet-desktop.yml/badge.svg)](https://github.com/damienbod/EmailCalandarsClient/actions/workflows/dotnet-desktop.yml)

## Blogs

[Send Emails using Microsoft Graph API and a desktop client](https://damienbod.com/2021/08/09/send-emails-using-microsoft-graph-api-and-a-desktop-client/)

## History

2021-12-14 Added more graph clients

2021-12-13 Updated to .NET 6, added calendar events, UserMailbox settings client

## Requirements

To send emails using Microsoft Graph API, you need to have an office license for the Azure Active Directory user which sends the email.

You can sign-in here to check this:

https://www.office.com

## Setup Azure App registration

The API permissions require the Graph delegated permissions.

![azureAppRegistration_emailCalendar_01](https://github.com/damienbod/EmailCalandarsClient/blob/main/images/azureAppRegistration_emailCalendar_01.png)


### Email client

The Azure App registration requires the Graph API delegated **Mail.Send** and the **Mail.ReadWrite** scopes.

```xml
  <appSettings>
    <add key="AADInstance" value="https://login.microsoftonline.com/{0}/v2.0"/>
    <add key="Tenant" value="--your-tenant--"/>
    <add key="ClientId" value="--your-client-id--"/>
    <add key="Scope" value="User.read Mail.Send Mail.ReadWrite"/>
  </appSettings>
```

### Presence client

The Azure App registration requires the Graph API delegated **User.Read.All**  scope.

```xml
  <appSettings>
    <add key="AADInstance" value="https://login.microsoftonline.com/{0}/v2.0"/>
    <add key="Tenant" value="--your-tenant--"/>
    <add key="ClientId" value="--your-client-id--"/>
    <add key="Scope" value="User.read User.Read.All"/>
  </appSettings>
```

### User Mailbox settings client

The Azure App registration requires the Graph API delegated **User.Read.All** **MailboxSettings.Read** scopes.

```xml
  <appSettings>
    <add key="AADInstance" value="https://login.microsoftonline.com/{0}/v2.0"/>
    <add key="Tenant" value="--your-tenant--"/>
    <add key="ClientId" value="--your-client-id--"/>
    <add key="Scope" value="User.read User.Read.All MailboxSettings.Read"/>
  </appSettings>
```

### Calendar client

The Azure App registration requires the Graph API delegated **User.Read.All** **Calendars.Read** Calendars.ReadWrite** scopes.

```xml
  <appSettings>
    <add key="AADInstance" value="https://login.microsoftonline.com/{0}/v2.0"/>
    <add key="Tenant" value="--your-tenant--"/>
    <add key="ClientId" value="--your-client-id--"/>
    <add key="Scope" value="User.read User.Read.All Calendars.Read Calendars.ReadWrite"/>
  </appSettings>
```

## Nuget packages

```xml
<ItemGroup>
  <PackageReference Include="Microsoft.Identity.Client" Version="4.39.0" />
  <PackageReference Include="Microsoft.Graph" Version="4.11.0" />
  <PackageReference Include="Newtonsoft.Json" Version="13.0.1" />
</ItemGroup>
```

## C# to Visual Basic

I converted the project from C# to Visual Basic using the telerik tool:

https://converter.telerik.com/

## Links

https://docs.microsoft.com/en-us/graph/outlook-send-mail-from-other-user

https://stackoverflow.com/questions/43795846/graph-api-daemon-app-with-user-consent

https://winsmarts.com/managed-identity-as-a-daemon-accessing-microsoft-graph-8d1bf87582b1

https://cmatskas.com/create-a-net-core-deamon-app-that-calls-msgraph-with-a-certificate/

https://docs.microsoft.com/en-us/graph/api/user-sendmail?view=graph-rest-1.0&tabs=http

https://docs.microsoft.com/en-us/answers/questions/43724/sending-emails-from-daemon-app-using-graph-api-on.html

https://stackoverflow.com/questions/56110910/sending-email-with-microsoft-graph-api-work-account

https://docs.microsoft.com/en-us/graph/sdks/choose-authentication-providers?tabs=CS#InteractiveProvider

https://docs.microsoft.com/en-us/graph/api/cloudcommunications-getpresencesbyuserid?view=graph-rest-1.0&tabs=csharp

https://docs.microsoft.com/en-us/graph/api/user-get-mailboxsettings?view=graph-rest-1.0&tabs=csharp

https://docs.microsoft.com/en-us/graph/api/user-list-events?view=graph-rest-1.0&tabs=csharp

## More information

For more information see MSAL.NET's conceptual documentation:
- [Quickstart: Register an application with the Microsoft identity platform](https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app)
- [Quickstart: Configure a client application to access web APIs](https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-configure-app-access-web-apis)
- [Recommended pattern to acquire a token in public client applications](https://github.com/AzureAD/microsoft-authentication-library-for-dotnet/wiki/AcquireTokenSilentAsync-using-a-cached-token#recommended-call-pattern-in-public-client-applications)
- [Acquiring tokens interactively in public client applications](https://github.com/AzureAD/microsoft-authentication-library-for-dotnet/wiki/Acquiring-tokens-interactively) 
- [Customizing Token cache serialization](https://github.com/AzureAD/microsoft-authentication-library-for-dotnet/wiki/token-cache-serialization)
