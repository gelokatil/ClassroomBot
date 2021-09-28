# ClassroomBot for Kubernetes Setup
This is a setup guide for this project.

# Requirements

- Azure subscription on same tenant as O365
- **Dev deploy only** :
  - Ngrok pro licence to allow TCP + HTTP tunnelling.
  - SSL certificate for NGrok URL, as per [this guide](https://github.com/microsoftgraph/microsoft-graph-comms-samples/blob/master/Samples/V1.0Samples/AksSamples/teams-recording-bot/docs/setup/certificate.md#%23generate-ssl-certificate).
- **Production deploy** :
  - Public bot domain (root-level) + DNS control for domain.
- Node JS LTS 14 to build Teams manifest.
- Docker for Windows to build bot image.
- Source code: [https://github.com/sambetts/ClassroomBot](https://github.com/sambetts/ClassroomBot)
- Bot permissions:
  - AccessMedia.All
  - JoinGroupCall.All
  - JoinGroupCallAsGuest.All

# Prerequisite Information

Most of these values we&#39;ll get after creating the resources below.

- Bot service DNS name - $botDomain.
- Azure container registry name/URL - $acrName (for &quot;contosoacr&quot;).
- Azure App Service to host Teams App; the DNS hostname - $teamsAppDNS.
- Application Insights instrumentation key - $appInsightsKey
- Bot App ID &amp; secret – we&#39;ll use the same app registration for the Teams App too.
  - $applicationId, $applicationSecret
- An Azure AD user object ID for which the bot will impersonate when editing meetings.
  - [https://docs.microsoft.com/en-us/graph/cloud-communication-online-meeting-application-access-policy](https://docs.microsoft.com/en-us/graph/cloud-communication-online-meeting-application-access-policy)
  - Bots can&#39;t edit/create online meetings as themselves; it must be done via a user impersonation with that user having rights to also edit meetings. For that user, we need the Azure AD object ID - $botUserId.

# Setup Steps

Build Docker image of bot.

1. Create an Azure container registry to push/pull bot image to.
2. With Docker in &quot;Windows container&quot; mode, build a bot image.
3. Tag image for your container registry and push. Take note of version tag (e.g &quot;classroombotregistry.azurecr.io/classroombot:1.0.5&quot; – this is your $containerTag).

Create Azure Resources

Create: Bot Service, Teams App Service, Application Insights

1. Create Azure bot. Add channel to Teams, with calling enabled with endpoint: https://$botDomain/api/calling
2. Take note of app ID &amp; secret – the secret of which is stored in an associated key vault.
3. Create app service for Teams App, with Node 14 LTS runtime stack.
  1. Recommended: Linux app service, on Free/Basic tier.
  2. Take note of URL hostname; this is your $teamsAppDNS. It can be the standard free Azure websites DNS.

Setup Teams App SSO &amp; allow Bot to access online meetings on behalf of a user

1. Edit the application registration to allow SSO.
2. For your bot user ID impersonation ID, $botUserId, run this to allow the bot to remove members from meetings - [https://docs.microsoft.com/en-us/graph/cloud-communication-online-meeting-application-access-policy](https://docs.microsoft.com/en-us/graph/cloud-communication-online-meeting-application-access-policy)

Create AKS resource

1. Create public IP address (standard SKU) for bot domain &amp; create/update DNS A-record. Resource-group can be the same as AKS resource.
2. Run &quot;setup.ps1&quot; to create AKS + bot architecture, with parameters:

- $azureLocation – example: &quot;westeurope&quot;
- $resourceGroupName – example: &quot;ClassroomBotProd&quot;
- $publicIpName – example: &quot;AksIpStandard&quot;
- $botDomain – example: &quot;classroombot.teamsplatform.app&quot;
- $acrName – example: &quot;classroombotregistry&quot;
- $AKSClusterName– example: &quot;ClassroomCluster&quot;
- $applicationId – example: &quot;151d9460-b018-4904-8f81-14203ac3cb4f&quot;
- $applicationSecret – example: &quot;9p96lolQJSD~\*\*\*\*\*\*\*\*\*\*\*\*&quot; (example truncated)
- $botName – example: &quot;ClassroomBotProd&quot;
- $containerTag– example: &quot;latest&quot;

Publish Teams App into App Service

- In &quot;ClassroomBot/TeamsApp/classroombot-teamsapp/.env&quot;, edit:
  - PUBLIC\_HOSTNAME - $teamsAppDNS
  - BOT\_HOSTNAME - $botDomain
  - TAB\_APP\_ID – your app ID - $applicationId
  - TAB\_APP\_URI – your app secret - $applicationSecret
  - MICROSOFT\_APP\_ID – your app ID (repeated) - $applicationId
  - MICROSOFT\_APP\_PASSWORD – your app secret (repeated) - $applicationSecret
  - APPLICATION\_ID - **generate a new GUID**.
  - PACKAGE\_NAME – generate your own package name.
- Publish classroombot-teamsapp website in App Service with VSCode and [this extension](https://marketplace.visualstudio.com/items?itemName=ms-azuretools.vscode-azureappservice).

Build Teams App Manifest

- Inside &quot;classroombot-teamsapp&quot; folder, run &quot;gulp manifest&quot;.
- Open &quot;classroombot-teamsapp/package&quot; and you&#39;ll find {PACKAGE\_NAME}.zip
  - This needs to be installed into Teams either via App Studio, or Teams Administration deployment.

Allow App Account to Impersonate User

[https://docs.microsoft.com/en-us/graph/cloud-communication-online-meeting-application-access-policy](https://docs.microsoft.com/en-us/graph/cloud-communication-online-meeting-application-access-policy)

The application will impersonate a user to update meetings using this method, but this requires setup.

- Connect-MicrosoftTeams
- New-CsApplicationAccessPolicy -Identity Meeting-Update-Policy -AppIds &quot;$applicationId&quot; -Description &quot;Policy to allow meetings to be updated by a bot&quot;
- Grant-CsApplicationAccessPolicy -PolicyName Meeting-Update-Policy -Identity &quot;$userId&quot;

Install the PS module with &quot;Install-Module -Name MicrosoftTeams&quot;

