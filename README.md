# Outlook Bridge for DeepSeek Harness

Outlook mail and calendar inside DeepSeek Harness. The plugin includes folder and unread views, inline replies, bulk mark as read, a monthly calendar, and AI-assisted mail summaries and drafts.

## Install and connect

Install the bundle from a supported Harness plugin source, then open **Outlook** in the sidebar. The panel can show demo data before an account is connected.

For your own Microsoft account, register an application in Microsoft Entra and choose the account types you intend to support. Enter its application ID and tenant in Outlook Bridge settings. Enable public client flows for device-code sign-in. If you use the authorization-code flow, register the native-client redirect URI shown by the plugin. Follow your organization's Microsoft consent requirements.

The plugin requests delegated Microsoft Graph permissions for the user profile, reading and updating mail, sending mail, managing calendar events, and supported To Do accounts, along with offline access. See [Microsoft identity platform documentation](https://learn.microsoft.com/en-us/entra/identity-platform/).

## Where data goes

Microsoft Graph requests run on the local Harness host. Tokens and connection settings are stored in the host's Outlook Bridge data directory. The UI requests mail and calendar information through the local host.

If you use an AI action, the selected email or typed instruction is sent to the model configured in Harness. Suggested replies open as editable drafts. You send them from the compose or reply controls.

Weather requests use [Open-Meteo](https://open-meteo.com/). When you use device location, approximate coordinates are used to request weather and to look up the city or district with [BigDataCloud](https://www.bigdatacloud.com/geocoding-apis/free-reverse-geocode-to-city-api). You can also choose a city manually.

## License

MIT. Outlook and DeepSeek Harness are independent products. This is a community plugin.
