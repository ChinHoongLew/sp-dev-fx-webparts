# react-quick-links-fluent

## Summary

Short summary on functionality and used technologies.

[picture of the solution in action, if possible]

## Used SharePoint Framework Version

![SPFx 1.21.1](https://img.shields.io/badge/version-1.21.1-green.svg)
![Node.js v22 ](https://img.shields.io/badge/Node.js-v20-green.svg) 

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Prerequisites
pre-create the list with the format below
| Column Name| Type                                               |
| -----------| ------------------------------------------------------- |
| Title          | Single Line of Text  |
| ICON  | Single Line of Text  | refer to https://uifabricicons.azurewebsites.net/ |
| LINK  | Single Line of Text  | the link of the quick link |
| POSITION    | Number  | position of the quick link |
| TARGET    | Choice (_blank,_self) | open in new window? | 
| GROUP     | Choice (Group1,Group2) | Group by |
| COLOR | Single Line of Text| color of the icon and font |
| BGCOLOR | Single Line of Text | background color of the quick link |
 
## Property Panel Configuration
|Name|Description|
|List Name| the name of the list|
|Group by?| filter query|
|Margin| the margin in between tiles|
|Padding| padding of the tiles|
|Max Width|maximum width of the tiles|
|Min Height|minimum height of the tiles|
|Grid Width|width of the whole grid|

## Powershell command to create the list
 <# provide the Site URL #>
$SiteURL = "https://yourtenantsite.sharepoint.com/sites/yoursiteurl"

<# Create "Quicklink" Lists #>
$ListTitle = "QuickLink Settings"
New-PnPList -Title $ListTitle -Template GenericList 


Add-PnPField -List $ListTitle -DisplayName "ICON" -InternalName "ICON" -Type Text -AddToDefaultView
Add-PnPField -List $ListTitle -DisplayName "LINK" -InternalName "LINK" -Type Text -AddToDefaultView
Add-PnPField -List $ListTitle -DisplayName "POSITION" -InternalName "POSITION" -Type Number -AddToDefaultView
Add-PnPField -List $ListTitle -DisplayName "TARGET" -InternalName "TARGET" -Type Choice -Group "TARGET" -AddToDefaultView -Choices "_blank","_self"
Add-PnPField -List $ListTitle -DisplayName "GROUP" -InternalName "GROUP" -Type Choice -Group "GROUP" -AddToDefaultView -Choices "MAIN", "GROUP1"
Add-PnPField -List $ListTitle -DisplayName "COLOR" -InternalName "COLOR" -Type Text -AddToDefaultView
Add-PnPField -List $ListTitle -DisplayName "BGCOLOR" -InternalName "BGCOLOR" -Type Text -AddToDefaultView

## Solution

| Solution    | Author(s)                                               |
| ----------- | ------------------------------------------------------- |
| React-Quick-Links-Fluent | ChinHoong Lew (https://www.linkedin.com/in/lewchinhoong/) |

## Version history

| Version | Date             | Comments        |
| ------- | ---------------- | --------------- |
| 1.0     | July 8, 2025 | Initial release |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - **npm install**
  - **gulp serve**

> Include any additional steps as needed.

## Features

Description of the extension that expands upon high-level summary above.

This extension illustrates the following concepts:

- @pnp/sp
- fluent UI - icon - https://uifabricicons.azurewebsites.net/

> Notice that better pictures and documentation will increase the sample usage and the value you are providing for others. Thanks for your submissions advance.

> Share your web part with others through Microsoft 365 Patterns and Practices program to get visibility and exposure. More details on the community, open-source projects and other activities from http://aka.ms/m365pnp.

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft teams](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development

<img src="https://m365-visitor-stats.azurewebsites.net/sp-dev-fx-webparts/samples/react-quick-links-fluent" />