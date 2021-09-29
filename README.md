# ClassroomBot Setup Guide

ClassroomBot is a &quot;Real-time Media Platform&quot; Teams bot that is designed to control students in a Teams meeting, so the organizer/teacher doesn&#39;t need to. It&#39;s a fairly basic for solution now &amp; just ejects meeting members that don&#39;t have their webcam activated but can also record the meeting audio streams too.

It&#39;s designed to run either locally via NGrok or in Azure Kubernetes Service so it can scale.

# Requirements

1. Azure subscription on same Azure tenant as Office 365/Teams
2. **Dev deploy only** :
    - Ngrok with pro licence (auth key needed to allow TCP + HTTP tunnelling).
    - SSL certificate for NGrok URL (see below).
    - Visual Studio 2019
3. **Production deploy** :
    - Public bot domain (root-level) + DNS control for domain.
4. Node JS LTS 14 to build Teams manifest.
5. Docker for Windows to build bot image.
6. Source code: [https://github.com/sambetts/ClassroomBot](https://github.com/sambetts/ClassroomBot)
7. Permissions in Azure AD application:
    - Application permissions (for bot):
        - AccessMedia.All
        - JoinGroupCall.All
        - JoinGroupCallAsGuest.All
        - OnlineMeetings.ReadWrite.All
    - Delegated permissions for Teams App (requested dynamically):
        - OnlineMeetings.ReadWrite - to create meetings.
        - ChannelMessage.Send - to publish meetings in a channel
        - GroupMember.Read.All - to read members of a group, for mentions
        - Directory.Read.All - to resolve group memebers to users
        - email 
        - openid 
        - profile 
        - offline_access

# Required Configuration Information

Most of these values we&#39;ll get after creating the resources below.

1. **Dev only** :
    - NGrok domains, TCP address, and auth token for pro license - $ngrokAuthToken.
    - SSL certificate thumbprint - $certThumbPrint.
2. Bot service DNS name - $botDomain.
    - **Production only:** this is your own domain.
    - **Dev** : this is your reserved NGrok domain
3. **Production only:**
    - Azure container registry name/URL - $acrName (for &quot;contosoacr&quot;).
    - Azure App Service to host Teams App; the DNS hostname - $teamsAppDNS.
    - Application Insights instrumentation key - $appInsightsKey
4. Azure AD: tenant ID, Bot App ID &amp; secret – we&#39;ll use the same app registration for the Teams App too.
    - $azureAdTenantId, $applicationId, $applicationSecret
5. Azure Bot Service name – $botName
6. An Azure AD user object ID for which the bot will impersonate when editing meetings.
    - [https://docs.microsoft.com/en-us/graph/cloud-communication-online-meeting-application-access-policy](https://docs.microsoft.com/en-us/graph/cloud-communication-online-meeting-application-access-policy)
    - Bots can&#39;t edit/create online meetings as themselves; it must be done via a user impersonation with that user having rights to also edit meetings. For that user, we need the Azure AD object ID - $botUserId.

# Setup Steps

These steps differ depending on whether you plan on running the bot in AKS/K8 or directly on from Visual Studio for developing the solution.

In these steps both the bot &amp; the Teams App share an Azure AD application registration, created normally by the Azure Bot Service.

## Prepare Local Files

Some files aren&#39;t tracked in git, so need creating locally from the templates.

- Copy &quot;deploy\cluster-issuer - template.yaml&quot; to &quot;deploy\cluster-issuer.yaml&quot;
  - Edit &quot;cluster-issuer.yaml&quot; and replace &quot;$YOUR\_EMAIL\_HERE&quot; with your own email.
  - This is used for LetsEncrypt and needs to be a proper email address; not a free one (Gmail, Outlook, etc)
- Copy &quot;TeamsApp\classroombot-teamsapp\template.env&quot; to &quot;TeamsApp\classroombot-teamsapp\\.env&quot;
- Copy &quot;BotService\Bot.Console\template.env&quot; to &quot;BotService\Bot.Console\\.env&quot;

## Dev Only: Setup NGrok configuration

For developer machines you&#39;ll want to run the bot directly from Visual Studio 2019 instead of in a container. For this to happen, we need inbound tunnelling to the right places.

1. In [https://dashboard.ngrok.com/](https://dashboard.ngrok.com/), reserve a TCP address &amp; x2 domains, all based in the US region.
    - Reserved TCP address for Skype Media endpoint – take note of address ($streamingAddressFull) &amp; port of the TCP addres ($streamingAddressPort).
    - Domain for bot service - $botDomain.
    - Domain for Teams app - $teamsAppDNS.

1. Configure &quot;%userprofile%\\.ngrok2\ngrok.yml&quot; with those domains &amp; TCP address like so:
    - authtoken: $ngrokAuthToken
    - tunnels:
    - classroombot:
        - addr: &quot;https://localhost:9441&quot;
        - proto: http
        - subdomain: classroombot
        - host-header: localhost
    - classroombotapp:
        - addr: &quot;http://localhost:3007&quot;
        - proto: http
        - subdomain: classroombotapp
        - host-header: localhost
    - media:
        - addr: 8445
        - proto: tcp
        - remote\_addr: &quot;1.tcp.ngrok.io:26065&quot;

Note: indentation is done with tab only as is standard yaml formatting. Also, the subdomains in this config file are what you&#39;ve registered in your NGrok account. The &quot;remote\_addr&quot; is the reserved TCP address given (they can&#39;t be specified, only given).

NGrok can be started with all tunnels with this command:

- ngrok start --all

The NGrok output should look something like this:

- Region United States
- tcp://1.tcp.ngrok.io:26065 -> localhost:8445
- http://classroombot.ngrok.io -> https://localhost:9441
- https://classroombot.ngrok.io -> https://localhost:9441
- http://classroombotapp.ngrok.io -> http://localhost:3007
- https://classroombotapp.ngrok.io -> http://localhost:3007

## Dev Only: Generate SSL for Bot Media TCP

As this bot receives audio/video streams it must expose a TCP endpoint with SSL in addition to the normal HTTP endpoints. For dev we must request these certificates manually; in production there is an AKS service we deploy to do it automatically.

1. Generate an SSL certificate for your NGrok addresses as per [this guide](https://github.com/microsoftgraph/microsoft-graph-comms-samples/blob/master/Samples/V1.0Samples/AksSamples/teams-recording-bot/docs/setup/certificate.md#%23generate-ssl-certificate).
    - In short, you need to use [certbot](https://certbot.eff.org/lets-encrypt/windows-other.html) to generate SSL certificates via LetsEncrypt.
    - Open port 80 of your bot domain with a specific ngrok command:
        - ngrok http 80 -subdomain $botDomain
    - Now run certbot to validate you own the domain &amp; download the certificates.
        - certbot certonly --standalone
    - This will create a temporary webserver that LetsEncrypt will read to validate ownership of the domain $botDomain – your NGrok tunnel domain. Once validated, certificates for that domain are downloaded in PEM format.
2. Once the PEM files have been created by certbot, you need to convert them to PFX format with [Open SSL](https://slproweb.com/products/Win32OpenSSL.html), in the directory the PEM files were created:
    - penssl pkcs12 -export -out mycert.pfx -inkey privkey1.pem -in cert1.pem -certfile chain1.pem
3. You&#39;ll need to create a password for the PFX file.
4. Now install the certificate into the machine&#39;s certificate store via MMC.
5. Take note of the certificate thumbprint – $certThumbPrint

## Create Azure Resources

Create: Bot Service, and production only: Teams App Service, Application Insights

1. Create Azure bot. Add channel to Teams, with calling enabled with endpoint: https://$botDomain/api/calling
2. Take note of app ID &amp; secret – the secret of which is stored in an associated key vault.
3. Production only: Create app service for Teams App, with Node 14 LTS runtime stack.

    - **Recommended** : Linux app service, on Free/Basic tier.
    - Take note of URL hostname; this is your $teamsAppDNS. It can be the standard free Azure websites DNS.
    - Dev only: your Teams app will be hosted by &quot;gulp&quot; and NGrok (which is your $teamsAppDNS). No need to create anything in Azure for it.
## Production only: build Docker image of bot

1. Create an Azure container registry to push/pull bot image to.
2. With Docker in &quot;Windows container&quot; mode, build a bot image from the root directory.
    - docker build -f ./build/Dockerfile . -t [TAG]
        - [TAG] is the FQDN of you container registry + image name, e.g. &quot;classroombotregistry.azurecr.io/classroombot:1.0.5&quot;
3. Push image to container registry with &quot;docker push&quot;. Take note of version tag (e.g &quot;classroombotregistry.azurecr.io/classroombot:1.0.5&quot; – this number/value is your $containerTag).
    - You may need to authenticate to your ACR first with &quot;az acr login --name $acrName&quot;

## Setup Teams App SSO
SSO is needed so users don&#39;t have to login again when on the Teams app tab.
1. Edit the application registration to allow SSO for the Teams App - [https://docs.microsoft.com/en-us/microsoftteams/platform/tabs/how-to/authentication/auth-aad-sso#develop-an-sso-microsoft-teams-tab](https://docs.microsoft.com/en-us/microsoftteams/platform/tabs/how-to/authentication/auth-aad-sso#develop-an-sso-microsoft-teams-tab).

## Production only: Create AKS resource via PowerShell

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

## Edit Teams App Environment Variables

1. In &quot;ClassroomBot/TeamsApp/classroombot-teamsapp/.env&quot;, edit:
    - PUBLIC\_HOSTNAME - $teamsAppDNS
    - BOT\_HOSTNAME - $botDomain
    - TAB\_APP\_ID – your app ID - $applicationId
    - TAB\_APP\_URI – your app secret - $applicationSecret
    - MICROSOFT\_APP\_ID – your app ID (repeated) - $applicationId
    - MICROSOFT\_APP\_PASSWORD – your app secret (repeated) - $applicationSecret
    - APPLICATION\_ID - **generate a new GUID**.
    - PACKAGE\_NAME – generate your own package name.

## Production only: Publish Teams App into App Service
Run the application from a dedicated website using the app-service created above.
1. Publish classroombot-teamsapp website in App Service with VSCode and [this extension](https://marketplace.visualstudio.com/items?itemName=ms-azuretools.vscode-azureappservice).

## Dev Only: Run Teams App with Gulp
If you&#39;re going to be editing the Teams app, do so with gulp.
1. In &quot;ClassroomBot/TeamsApp/classroombot-teamsapp&quot; run commands:

    - npm install
    - gulp serve --debug
This will run the application locally in debug-mode. NGrok should make it accessible via the $teamsAppDNS tunnel.
## Build &amp; Deploy Teams App Manifest

1. Inside &quot;classroombot-teamsapp&quot; folder, run &quot;gulp manifest&quot;.
2. Open &quot;classroombot-teamsapp/package&quot; and you&#39;ll find {PACKAGE\_NAME}.zip
    - This needs to be installed into Teams either via App Studio, or Teams Administration deployment.

## Allow App Account to Impersonate User
[https://docs.microsoft.com/en-us/graph/cloud-communication-online-meeting-application-access-policy](https://docs.microsoft.com/en-us/graph/cloud-communication-online-meeting-application-access-policy)
1. The application will impersonate a user to update meetings using this method, but this requires setup.

    - Connect-MicrosoftTeams
    - New-CsApplicationAccessPolicy -Identity Meeting-Update-Policy -AppIds &quot;$applicationId&quot; -Description &quot;Policy to allow meetings to be updated by a bot&quot;
    - Grant-CsApplicationAccessPolicy -PolicyName Meeting-Update-Policy -Identity &quot;$userId&quot;
    - Set-CsApplicationMeetingConfiguration -AllowRemoveParticipantAppIds @{Add="$applicationId"}
    
Install the PS module with &quot;Install-Module -Name MicrosoftTeams&quot;
## Dev Only: Run Solution from Visual Studio
If you&#39;re deploying to AKS, this won&#39;t be necessary.
1. Open &quot;ClassroomBot\BotService\Bot.Console\.env&quot; and update the following values.

    - AzureSettings\_\_BotName - $botName
    - AzureSettings\_\_AadAppId - $applicationId
    - AzureSettings\_\_AadTenantId - $azureAdTenantId
    - AzureSettings\_\_AadAppSecret - $applicationSecret
    - AzureSettings\_\_ServiceDnsName - $botDomain
    - AzureSettings\_\_CertificateThumbprint - $certThumbPrint
    - AzureSettings\_\_InstancePublicPort - $streamingAddressPort

1. Run Visual Studio as administrator and start debugging &quot;Bot.Console&quot;.

# Running Solution
Once the bot service is running and the Teams app is deployed, you need to install the teams app if you&#39;re not side-loading with App Studio. Search for &quot;classroombot&quot; in the Teams app store and you should find it there. Open it and you&#39;ll see the single tab with the &quot;create a meeting&quot; page if the URLs are all correct &amp; working. 

The 1st time you run the app you&#39;ll have to grant access to the graph permissions if not done so proactively. The app will open a new window with a login &amp; consent flow, but will fail to redirect back. This is fine; we just want the consent for now.Refresh the app tab after granting consent and the list of teams your user is joined to should show.Click &quot;start meeting&quot; against one, enter some meeting details, and click &quot;start new class&quot;. This will create a new meeting for the group, and post in the default channel that a new meeting has started with a mention for anyone in that group.

# Updating Solution
For the Teams App, just republish any changes directly to the App Service. There are more elegant ways of doing this, but this works & is easy at least.

For the bot:
- Build a new image, with a new tag (previous tag was 1.0.6). From the solution root:
    - docker build -f ./build/Dockerfile . -t classroombotregistry.azurecr.io/classroombot:1.0.7
- Login & push image:
    - az acr login --name classroombotregistry
    - docker push classroombotregistry.azurecr.io/classroombot:1.0.7
- Update helm deployment (check params 1st):
    - helm upgrade classroombot ./deploy/classroombot --namespace classroombot --set host=classroombot.teamsplatform.app --set public.ip=20.73.235.135 --set image.domain="classroombotregistry.azurecr.io" --set image.tag=1.0.7 --set scale.replicaCount=3
- Verify pods are running after a few minutes:
    - kubectl get pods -n classroombot

**Important:** you have to pass in host, IP etc, so make sure you get the right values!
