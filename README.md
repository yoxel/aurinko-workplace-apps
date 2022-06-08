# Workplace app examples

Aurinko provides a backend API for single-page apps with a focus on workplace integrations and unified APIs. Aurinko is built to support workplace apps/addins/addons like Office 365 addins, MS Teams apps, Zoom/WebEx in-meeting apps. It will allow your webapp to authenticate a user using Office 365, Gmail, Salesforce,... oauth flow, and will manage a user session based on that id. 

The benefits of using Aurinko as the backend are the following:

1. Pre-built backend API, openid-based user management is provided, supports Outlook's idToken.
2. Unified OAuth flows, and access token management. OAuth tokens never reside on the client-side.
3. Unified error handling. 

## Virtualized/unified API

ISVs will appreciate the fact that they can build apps for one API to work with many data providers:

1. [Unified mailbox API](https://docs.aurinko.io/article/8-what-is-aurinko) to access email/calendar/contacts data across multiple providers (Office 365, Gmail, MS Exchange, IMAP).

2. Virtualized/unified CRM API to define unified data models specifically for your app needs. Currently supporting Salesforce, Sugar CRM, HubSpot, ...


## Outlook addin

Here is how to set up Aurinko.io as a backend API for your Outlook addin: [Create your first Outlook addin](https://docs.aurinko.io/article/36-create-your-first-outlook-addin)!

The addin manifest file that Aurinko generates for you is just a quickstart example that you can [install in Outlook](https://docs.aurinko.io/article/37-office-365-installing-outlook-addin) right away. The manifest activates the addin's READ mode using outlook/read.html code. 

Check out this article [Outlook addins](https://docs.aurinko.io/article/30-outlook-addins) to learn more.

## MS Teams app

Here is how to set up Aurinko.io as a backend API for your Teams app (tab): [Create your first Teams app](https://docs.aurinko.io/article/40-create-your-first-ms-teams-app)!

The app package that Aurinko generates for you is just a quickstart example that you can install in [MS Teams](https://docs.aurinko.io/article/38-installing-ms-teams-app) right away. The package activates one Teams tab (bot is not activated).

Check out this article [MS Teams apps](https://docs.aurinko.io/article/31-microsoft-teams-apps) to learn more.
