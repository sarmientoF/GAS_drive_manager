///////////////// Import Variables ////////////
var scriptProperties = PropertiesService.getScriptProperties()

const sharedSpreadId = scriptProperties.getProperty('sharedSpreadId')
const sharedFolderId = scriptProperties.getProperty('sharedFolderId')
const developerKey = scriptProperties.getProperty('developerKey')

const driveFolder = DriveApp.getFolderById(sharedFolderId)
const spreadsheet = SpreadsheetApp.openById(sharedSpreadId)

const fileIdRow = 0
const copyIdRow = 1
const emailRow = 3
const expirationRow = 5
const versionRow = 7

///////////////// Trigger ////////////////////////////

function removeExpiredFiles() {
    var todayTime = new Date().getTime()
    var range = spreadsheet.getDataRange()
    var sheet = range.getValues()
    var values = sheet.slice(1, sheet.length)
    var expiredRows = []
    var count = 0
    values.forEach(function (row, index) {
        var copyId = row[copyIdRow]
        var expiration = row[expirationRow]
        var expirationDate = new Date(expiration)
        // If expired Remove Copy file
        if (expirationDate.getTime() < todayTime) {
            // Remove File
            let expiredFile = DriveApp.getFileById(copyId)
            expiredFile.setTrashed(true)
            // Add expired row
            expiredRows.push(index)
        }
    })
    expiredRows.forEach(function (index) {
        spreadsheet.deleteRow(index + 2 - count)
        count += 1
    })
}

/////////////////// Spreadsheet ///////////////////////
function getVersions(myFileId: string) {
    var range = spreadsheet.getDataRange()

    var sheet = range.getValues()

    var values = sheet.slice(1, sheet.length)

    var version = 0

    values.forEach((row) => {
        let fileId = row[fileIdRow]
        let fileVersion = row[versionRow]
        if (fileId == myFileId) {
            version = fileVersion + 1
        }
    })
    return version
}
function getViewers(fileId: string) {
    const file = DriveApp.getFileById(fileId)
    const viewers = file.getViewers()
    return viewers.map((viewer) => {
        return viewer.getEmail()
    })
}

function modifySpread(contents: any[]) {
    spreadsheet.appendRow(contents)
}

////////////////// Share File ////////////////////
const customMessage = (
    extraMessage: string,
    expiration: string,
    fileName: string,
    version: number = 0
) => {
    var verString = version == 0 ? '' : `Ver${version}_`
    return `${extraMessage}
    ファイル名: ${verString}${fileName}
    このファイルは${expiration}には削除しますので、必要に応じてそれまでにダウンロードいただくようお願いします。`
}

function uploadFromDrive(e: any) {
    try {
        // Read Form inputs
        var arr = e['array[]'] as string[] | string
        var ids = e.fileId as string
        var expiration = e.expiration as string
        var message = e.message as string
        // Get Files ids and emails
        var emails = arr.toString().split(',')
        var fileIds = ids.split(',')

        if (fileIds.length == 1) {
            var fileId = fileIds[0]
            try {
                // is file
                var file = DriveApp.getFileById(fileId)
                var version = getVersions(fileId)
                var versionName = version > 0 ? `Version${version}_` : ''
                var newName = `${versionName}${file.getName()}`

                var copy = file.makeCopy(newName, driveFolder)

                shareFileToUsers(
                    emails,
                    fileId,
                    copy.getId(),
                    newName,
                    expiration,
                    version,
                    message
                )
            } catch (error) {
                // is folder
                var {
                    id: folderId,
                    copyId: copyId,
                    name: folderName,
                    version: version_,
                } = copySingleFolder(fileIds[0])
                shareFolderToUsers(
                    emails,
                    folderId,
                    copyId,
                    folderName,
                    expiration,
                    message,
                    version_
                )
            }
        } else {
            // Find folder where files are
            var { id: folderId_, name: folderName_ } =
                copyMultipleFiles(fileIds)
            shareFolderToUsers(
                emails,
                folderId_,
                folderId_,
                folderName_,
                expiration,
                message,
                0
            )
        }
        return true
    } catch (error) {
        throw new Error('Whoops!')
    }
}

function getFiles(
    rootFolder: GoogleAppsScript.Drive.Folder,
    destFolder: GoogleAppsScript.Drive.Folder
) {
    var files = rootFolder.getFiles()
    while (files.hasNext()) {
        var file = files.next()
        file.makeCopy(destFolder)
    }
    var folders = rootFolder.getFolders()
    while (folders.hasNext()) {
        var folder = folders.next()
        var copyFolder = destFolder.createFolder(folder.getName())
        getFiles(folder, copyFolder)
    }
}
function copySingleFolder(id: string) {
    var folder = DriveApp.getFolderById(id)

    var version = getVersions(id)
    var versionName = version > 0 ? `Version${version}_` : ''
    var newName = `${versionName}${folder.getName()}`

    var outFolder = driveFolder.createFolder(newName)
    getFiles(folder, outFolder)

    return {
        id: id,
        copyId: outFolder.getId(),
        name: outFolder.getName(),
        version: version,
    }
}

function copyMultipleFiles(ids: string[]) {
    var names: Array<string> = []
    const tempName = Math.random().toString(36).substr(2, 9)
    const outFolder = driveFolder.createFolder(tempName)
    ids.forEach((id) => {
        try {
            let file = DriveApp.getFileById(id)
            file.makeCopy(file.getName(), outFolder)
            names.push(file.getName())
        } catch (err) {
            var folder = DriveApp.getFolderById(id)
            var copyFolder = outFolder.createFolder(folder.getName())
            getFiles(folder, copyFolder)
            names.push(folder.getName())
        }
    })
    outFolder.setName(names.join('__'))

    return { id: outFolder.getId(), name: outFolder.getName() }
}

function shareFileToUsers(
    emails: string[],
    fileId: string,
    copyId: string,
    fileName: string,
    expiration: string,
    version: number,
    extraMessage: string
) {
    var date = new Date(expiration)
    date.setHours(date.getHours() + 1)

    emails.forEach((email, i) => {
        let body = customMessage(extraMessage, expiration, fileName, version)
        shareFileToUser(email, copyId, body, date)
    })

    modifySpread([
        fileId,
        copyId,
        fileName,
        emails.toString(),
        extraMessage,
        date.toISOString(),
        new Date().toTimeString(),
        version,
    ])
}

function shareFolderToUsers(
    emails: string[],
    fileId: string,
    copyId: string,
    fileName: string,
    expiration: string,
    extraMessage: string,
    version: number
) {
    var date = new Date(expiration)
    date.setHours(date.getHours() + 1)

    emails.forEach((email) => {
        let body = customMessage(extraMessage, expiration, fileName)
        shareFileToUser(email, copyId, body, date)
    })

    modifySpread([
        fileId,
        copyId,
        fileName,
        emails.toString(),
        extraMessage,
        date.toISOString(),
        new Date().toTimeString(),
        version,
    ])
}

function shareFileToUser(
    email: string,
    fileId: string,
    body: string,
    expiration: Date
) {
    var permission = Drive.Permissions.insert(
        {
            value: email,
            type: 'user',
            role: 'reader',
            withLink: false,
        },
        fileId,
        {
            sendNotificationEmails: true,
            emailMessage: body,
            supportsAllDrives: true,
        }
    )
    Drive.Permissions.patch(
        {
            expirationDate: expiration.toISOString(),
        },
        fileId,
        permission.id
    )
}

/////////// Google Drive Picker /////////////////////////
function onOpen() {
    SpreadsheetApp.getUi()
        .createMenu('Picker')
        .addItem('Start', 'showPicker')
        .addToUi()
}

function showPicker() {
    var html = HtmlService.createHtmlOutputFromFile('form.html')
        .setWidth(600)
        .setHeight(425)
        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    SpreadsheetApp.getUi().showModalDialog(html, 'Select a file')
}

function getOAuthToken() {
    DriveApp.getRootFolder()
    return [ScriptApp.getOAuthToken(), developerKey]
}

function getDeveloperKey() {
    return developerKey
}
function getSpreadsheetURL() {
    return spreadsheet.getUrl()
}
///////////////// DO GET ///////////////////////////////////
function doGet() {
    return HtmlService.createTemplateFromFile('form.html')
        .evaluate()
        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .setTitle('Google Visitor Manager')
    // .setSandboxMode(HtmlService.SandboxMode.NATIVE)
    // .setXFrameOptionsMode(HtmlService.SandboxMode.NATIVE)
    // .setXFrameOptionsMode()
}
