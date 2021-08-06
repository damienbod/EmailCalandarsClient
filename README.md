# Graph API Email client

## Setup Azure App registration


## Nuget packages

```xml
<ItemGroup>
	<PackageReference Include="Microsoft.Identity.Client" Version="4.35.1" />
	<PackageReference Include="Microsoft.Identity.Web.MicrosoftGraphBeta" Version="1.15.2" />
	<PackageReference Include="Newtonsoft.Json" Version="13.0.1" />
</ItemGroup>
```

## Links

https://docs.microsoft.com/en-us/graph/outlook-send-mail-from-other-user

https://stackoverflow.com/questions/43795846/graph-api-daemon-app-with-user-consent

https://winsmarts.com/managed-identity-as-a-daemon-accessing-microsoft-graph-8d1bf87582b1

https://cmatskas.com/create-a-net-core-deamon-app-that-calls-msgraph-with-a-certificate/

https://docs.microsoft.com/en-us/graph/api/user-sendmail?view=graph-rest-1.0&tabs=http

https://docs.microsoft.com/en-us/answers/questions/43724/sending-emails-from-daemon-app-using-graph-api-on.html

https://stackoverflow.com/questions/56110910/sending-email-with-microsoft-graph-api-work-account

https://docs.microsoft.com/en-us/graph/sdks/choose-authentication-providers?tabs=CS#InteractiveProvider

https://converter.telerik.com/


## More information

For more information see MSAL.NET's conceptual documentation:
- [Quickstart: Register an application with the Microsoft identity platform](https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app)
- [Quickstart: Configure a client application to access web APIs](https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-configure-app-access-web-apis)
- [Recommended pattern to acquire a token in public client applications](https://github.com/AzureAD/microsoft-authentication-library-for-dotnet/wiki/AcquireTokenSilentAsync-using-a-cached-token#recommended-call-pattern-in-public-client-applications)
- [Acquiring tokens interactively in public client applications](https://github.com/AzureAD/microsoft-authentication-library-for-dotnet/wiki/Acquiring-tokens-interactively) 
- [Customizing Token cache serialization](https://github.com/AzureAD/microsoft-authentication-library-for-dotnet/wiki/token-cache-serialization)
