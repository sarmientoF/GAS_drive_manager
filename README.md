# Google Visitor Manager

## Overview

###### This web app provides the following functionalities:

-   A google workspace user is able to share files or folders with any email holders for a period of time not longer than a year.

-   When sharing multiple files, a zip file, that contains the multiple files selected, will be created and shared

-   The shared file permission will be revoked automatically when the specified date comes

-   Non-google users can be access to the shared file(s)

- The first time of use, it will create a new Spreadsheet in your google drive root with the name `google-visitor-manger`. Also, every time you share a file(s), it will keep track of your shared file(s) by adding a new row to the `google-visitor-manger` file

## [Requirements](https://developers.google.com/apps-script/guides/typescript)

-   Node, npm and yarn(optional):
    > npm install --global yarn
-   [clasp](https://developers.google.com/apps-script/guides/clasp):
    > npm install @google/clasp
-   Type definitions for Apps Script:
    > npm i -S @types/google-apps-script
-   Visual Studio Code (for TypeScript IDE autocompletion)

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
        - \*google.com
        - \*googleusercontent.com
    - API restrictions -> Don't restrict key

4. Open the file `form.html` and update this file with your `API_KEY` 🔑
    ```js
    var DEVELOPER_KEY = 'YOUR_API_KEY'
    ```
5. Login to clasp and create a newGAS project, select webapp
    ```sh
    yarn clasp login
    yarn clasp create [scriptTitle]
    ```
6. Enable APIs & services, it'll ask for your `Project ID`
    ```sh
    yarn clasp apis list
    yarn clasp apis enable picker
    yarn clasp apis enable drive
    yarn clasp apis enable script
    ```
7. Add drive service to your GAS project by adding the following lines to the file `appsscript.json`

    ```json
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

8. Push your changes
    ```sh
    yarn clasp push
    ```
9. Finally, go to [your GAS project](https://script.google.com/home) and create a new deploy as `Web app`
    - Web app -> Execute as : User accessing the web app
    - Web app -> Who has access : Anyone within `YOUR_WORKSPACE_NAME`

## 🚀🚀 Now you are ready to go!
