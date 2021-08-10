# Google Visitor Manager

## Overview

###### This web app provides the following functionalities:

-   A google workspace user is able to share files or folders with any email holders for a period of time not longer than a year.

-   When sharing multiple files and folders, a copy, that contains the multiple files and folders selected, will be created in your shared folder, and this copy will be shared

-   The shared file permission will be revoked automatically when the specified date comes

-   Every day at `1:00 AM`, the rows of the expired files inside your `Spreadsheet` will be deleted, and the shared file will be deleted from your shared folder.

-   Non-google users can be access to the shared file(s)

## [Requirements](https://developers.google.com/apps-script/guides/typescript)

-   Node, npm and yarn:
    > npm install --global yarn
-   [clasp](https://developers.google.com/apps-script/guides/clasp):
    > npm install @google/clasp
-   Type definitions for Apps Script:
    > npm i -S @types/google-apps-script
-   Visual Studio Code (for TypeScript IDE autocompletion)

-   A folder to store the a copy of the file(s) you want to share (`sharedFolderId`)
-   A spreadsheet to keep track of your shared files (`sharedSpreadId`) with headers in the first row
    -   Original | File ID | Copy File ID | Name | Shared Emails | Extra Message | Expiration Date ISO | Shared Date | Version

## Installation

1. Clone this repo and install the dependencies listed in `package.json`

    ```sh
    git clone [REPO_URL]

    cd google-visitor-manager

    yarn install
    ```

2. [Create a new google cloud platform project](https://cloud.google.com/resource-manager/docs/creating-managing-projects) and store your `Project ID`

3. [Create a new API key](https://cloud.google.com/docs/authentication/api-keys) with the following parameters and store the generated key (sample `API_KEY` 🔑 : AIzaSyA8ywe**\*\*\*\***\*\*\*\***\*\*\*\***wjXTPrBY)

    - Application restrictions -> HTTP referrers(websites)
    - Website restrictions
        - \*google.com/\*
        - \*googleusercontent.com/\*
    - API restrictions -> Don't restrict key

4. Create a new file `1_propertyService.ts`(Google apps script run files alphabetically, so it will run first this file then go to `2_Code.ts`) and add the following lines of code using your own ids and key (This file contains sensitive information 🔑, DON'T EXPOSE YOUR SENSITIVE INFORMATION 🚨)

    ```ts
    var scriptProperties = PropertiesService.getScriptProperties()
    scriptProperties.setProperties({
        sharedSpreadId: '********************************************',
        sharedFolderId: '*********************************',
        developerKey: 'YOUR_DEVELOPER_KEY',
    })
    ```

5. Enable APIs & services inside Google Cloud Platform project:

    - [Setting up OAuth 2.0](https://support.google.com/cloud/answer/6158849?hl=en) with the following setting
        - Internal Application
    - [Enable API Libraries](https://cloud.google.com/endpoints/docs/openapi/enable-api)
        - Google Picker API
        - Google Drive API
        - Apps Script API

6. Login to clasp and create a new GAS project
    ```sh
    yarn clasp login
    yarn clasp create --title [scriptTitle] --type webapp
    ```
7. Specify the id of the Google Cloud Platform project

    - Run `clasp open` or open your GAS page
    - Click `Project Settings > Change project`
    - Specify the project number `************`

8. Add drive service to your GAS project by adding the following lines to the file `appsscript.json`

    ```json
    "timeZone": "Asia/Tokyo",
    "dependencies": {
        "enabledAdvancedServices": [
        {
            "userSymbol": "Drive",
            "version": "v2",
            "serviceId": "drive"
        }
        ]
    },
    "webapp": {
        "executeAs": "USER_ACCESSING",
        "access": "DOMAIN"
    }
    ```

9. Push your changes
    ```sh
    yarn clasp push
    ```
10. [Create a new deployment](https://developers.google.com/apps-script/concepts/deployments) as a `Web app`
    - Web app -> Execute as : User accessing the web app
    - Web app -> Who has access : Anyone within `YOUR_WORKSPACE_NAME`
11. Finally, [Add a time-driven trigger](https://developers.google.com/apps-script/guides/triggers/installable)
    - Choose a function to run: `removeExpiredFiles`
    - Chose which deployment should run: `Head`
    - Select event source: `Time-driven`
    - Select type of time based trigger: `Day timer`
    - Select time of day: `1am to 2am`
