# Group People

## Summary

Display the members of a target SharePoint group. An alternative to display dynamically people without edit the page.

![](assets/LsOnline-SPFx-GroupPeople.gif)

### Used SharePoint Framework Version

![SPFx 1.11.0](https://img.shields.io/badge/SPFx-1.11.0-success.svg)

## Applies to

* [SharePoint Framework Developer][1]
* [Office 365 developer tenant][2]

## Prerequisites

No prerequisites

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

To build manually the package, please make sure you have the prerequisites like illustrated to the [Set up your SharePoint Framework development environment][3] article and follow the next steps:

* Clone this repository
* in the command line run:
  * `npm i`
  * `gulp bundle --ship`
  * `gulp package-solution --ship`
* deploy the package to your **Tenant App Catalog** or **Site Collection App Catalog**
* add the web part to a page

## Features

This SharePoint Framework Web Part allow to:

- retrieving the SharePoint Groups from the current site
- retrieving SharePoint users profiles properties
- passing Web Part properties to React components
- building dynamic web part properties (SharePoint groups)
- managing the displayed title
- choose to keep web part visible or not if nothing to show
- manage personas size (regular, large or extra large)
- manage user profiles properties to show

### Next steps

- Get members of AD security groups - recursively (Graph API)
- Get members of Office 365 group (Graph API)
- Get user profile from Microsoft 365 (Graph API)

[1]: https://docs.microsoft.com/sharepoint/dev/spfx/sharepoint-framework-overview
[2]: https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-developer-tenant
[3]: https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-development-environment