---
page_type: sample
description: This sample demonstrates how to use the Microsoft Graph JavaScript SDK to send email in Excel from Office Add-ins.
products:
- ms-graph
- microsoft-graph-email-api
languages:
- javascript
---

# Microsoft Excel Mail Merge Sample Office Add-in

[![Node.js build](https://github.com/microsoftgraph/msgraph-training-office-addin/actions/workflows/node.js.yml/badge.svg)](https://github.com/microsoftgraph/msgraph-training-office-addin/actions/workflows/node.js.yml) ![License.](https://img.shields.io/badge/license-MIT-green.svg)

This sample demonstrates how to use the Microsoft Graph JavaScript SDK to send emails in Excel from Office Add-ins.

# Required Steps & How to Run

## Prerequisites

To run the completed project in this folder, you need the following:

- [Node.js](https://nodejs.org) installed on your development machine. (**Note:** This tutorial was written with Node version 16.14.0. The steps in this guide may work with other versions, but that has not been tested.)
- Either a personal Microsoft account with a mailbox on Outlook.com, or a Microsoft work or school account.

If you don't have a Microsoft account, there are a couple of options to get a free account:

- You can [sign up for a new personal Microsoft account](https://signup.live.com/signup?wa=wsignin1.0&rpsnv=12&ct=1454618383&rver=6.4.6456.0&wp=MBI_SSL_SHARED&wreply=https://mail.live.com/default.aspx&id=64855&cbcxt=mai&bk=1454618383&uiflavor=web&uaid=b213a65b4fdc484382b6622b3ecaa547&mkt=E-US&lc=1033&lic=1).
- You can [sign up for the Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program) to get a free Microsoft 365 subscription.

## Register a web application with the Azure Active Directory admin center

1. Open a browser and navigate to the [Microsoft Entra admin center](https://aad.portal.azure.com). Login using a **personal account** (aka: Microsoft Account) or **Work or School Account**.

1. Select **Identity** in the left-hand navigation, then select **App registrations** under **Applications**.

1. Select **New registration**. On the **App registrations** page, set the values as follows.

    - Set **Name** to `Office Add-in Graph Tutorial`.
    - Set **Supported account types** to **Accounts in any organizational directory and personal Microsoft accounts**.
    - Under **Redirect URI**, set the first drop-down to `Single-page application (SPA)` and set the value to `https://localhost:3000/consent.html`.

1. Select **Register**. On the **Office Add-in Graph Tutorial** page, copy the value of the **Application (client) ID** and save it, you will need it in the next step.

## Configure the sample

1. Edit the `consent.js` file and make the following changes.
    - Replace `YOUR_APP_ID_HERE` with the **Application Id** you got from the App Registration Portal.
1. In your command-line interface (CLI), navigate to this directory and run the following command to install requirements.

    ```
    npm install
    ```

## Run the sample

Run the following command in your CLI to start the application.
```
npm run build
npm start
```

## Expected result

A webpack server will be hosted on https://localhost:3000/, as the CLI shows:

![](https://github.com/SiruiSun-MSFT/Mail-Management-Add-in-for-Excel/blob/main/assets/webpack.png)

An Excel desktop application will be auto-launched and the Mail Merge Addin will be auto-run on the right taskpane area.

![](https://github.com/SiruiSun-MSFT/Mail-Management-Add-in-for-Excel/blob/main/assets/taskpane.png)

Please follow the steps below:

1. Create Sample Data, including valid email address (required) and other information.

2. Verify Template and Data, the To Line must contain the column name of the email address.

3. Send Email, which will pop up a dialog to get the consent of Microsoft Graph. After sign-in, the email will be send out.

![](https://github.com/SiruiSun-MSFT/Mail-Management-Add-in-for-Excel/blob/main/assets/mail.png)

## Questions and comments

We'd love to get your voice about the Microsoft Excel Mail Merge Sample Office Add-in. You can send your feedback to us in the [Survey](https://forms.office.com/Pages/ResponsePage.aspx?id=v4j5cvGGr0GRqy180BHbR8GFRbAYEV9Hmqgjcbr7lOdUNVAxQklNRkxCWEtMMFRFN0xXUFhYVlc5Ni4u).

## Code of conduct

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**