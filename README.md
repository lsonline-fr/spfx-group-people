# Group People

## Summary

Display the members of a target SharePoint group. An alternative to display dynamically people without edit the page.

![](assets/LsOnline-SPFx-GroupPeople.gif)

### Used SharePoint Framework Version

![SPFx 1.8.1](https://img.shields.io/badge/SPFx-1.8.1-success.svg)

## Applies to

* [SharePoint Framework][1]
* [Office 365 tenant][2]

## Prerequisites
 
 * React 
 * PnP-JS-Core
 * React UI Fabric
 * SPFx Controls PlaceHolder

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Prerequisites

- SharePoint Online tenant with Office Graph content enabled

## Minimal Path to Awesome

To build manually the package, please make sure you have the prerequisites like illustrated to the [Set up your SharePoint Framework development environment][3] article and follow the next steps:

* clone this repo
* in the command line run:
  * `npm i`
  * `gulp bundle --ship`
  * `gulp package-solution --ship`
* deploy the package to your **Tenant App Catalog** or **Site Collection App Catalog**
* add the web part to a page

## Features

This SharePoint Framework Web Part allow to:

- retrieving the SharePoint Groups from the current web
- retrieving users profiles properties
- passing Web Part properties to React components
- building dynamic web part properties
- managing the displayed title

[1]: https://dev.office.com/sharepoint
[2]: https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant
[3]: https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-development-environment