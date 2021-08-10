const spreadName = 'google-visitor-manager'
const spreadId = findFile()
console.log('🚀🚀', spreadId)

function createSpreadSheet() {
    var folderId = DriveApp.getRootFolder().getId()
    var resource: GoogleAppsScript.Drive.Schema.File = {
        title: spreadName,
        //@ts-ignore
        mimeType: MimeType.GOOGLE_SHEETS,
        parents: [{ id: folderId }],
    }
    var fileJson = Drive.Files.insert(resource)
    var fileId = fileJson.id
    modifySpread(
        ['FileID', 'Name', 'UserEmails', 'Expiration', 'isTrashed'],
        fileId
    )
    console.log('🚀 Created Spread', fileId)
    return fileId
}
function findFile() {
    var root = DriveApp.getRootFolder()
    var files = root.getFilesByName(spreadName)
    if (files.hasNext()) {
        return files.next().getId()
    }
    return createSpreadSheet()
}

function getVersions(myFileId: string, myEmails: string[]) {
    var file = SpreadsheetApp.openById(spreadId)
    var range = file.getDataRange()

    var sheet = range.getValues()

    var values = sheet.slice(1, sheet.length)

    var versions = myEmails.map((_) => 0)
    values.forEach((row) => {
        let fileId = row[0]
        let emails = (row[2] as string).split(',')
        if (fileId == myFileId) {
            myEmails.forEach((myEmail, i) => {
                if (emails.indexOf(myEmail) > -1) versions[i] += 1
            })
        }
    })

    console.log('🚀', versions)
    return versions
}
function getViewers(fileId: string) {
    const file = DriveApp.getFileById(fileId)
    const viewers = file.getViewers()
    return viewers.map((viewer) => {
        return viewer.getEmail()
    })
}

function onOpen() {
    SpreadsheetApp.getUi()
        .createMenu('Picker')
        .addItem('Start', 'showPicker')
        .addToUi()
}

function showPicker() {
    var html = HtmlService.createHtmlOutputFromFile('dialog.html')
        .setWidth(600)
        .setHeight(425)
        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    SpreadsheetApp.getUi().showModalDialog(html, 'Select a file')
}

function getOAuthToken() {
    DriveApp.getRootFolder()
    return ScriptApp.getOAuthToken()
}

function doGet() {
    return HtmlService.createTemplateFromFile('form.html')
        .evaluate()
        .setTitle('Google GAS application')
        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
}

function modifySpread(contents: string[], id: string = spreadId) {
    console.log('🚀 modify Spread', contents)

    var file = SpreadsheetApp.openById(id)
    file.appendRow(contents)
}

const customMessage = (
    extraMessage: string,
    expiration: string,
    fileName: string,
    version: number = 0
) => {
    var verString = version == 0 ? '' : `_Ver${version}`
    return `ProCube作業の作業報告書を送ります。
    Filename: ${fileName}${verString}。
    ${extraMessage}。
    このファイルは${expiration}には削除しますので、必要に応じてそれまでにダウンロードいただくようお願いします。`
}

function uploadFromDrive(e: any) {
    // Read Form inputs
    var arr = e['array[]'] as string[] | string
    var ids = e.fileId as string
    var expiration = e.expiration as string
    var message = e.message as string
    // var expirationOr = new Date(e.expiration)
    // var expiration = new Date(expirationOr.getTime() - 9 * 60 * 60 * 1000)
    // Get Files ids and emails
    var emails = arr.toString().split(',')
    var fileIds = ids.split(',')

    if (fileIds.length == 1) {
        var fileId = fileIds[0]
        var fileName = DriveApp.getFileById(fileId).getName()
        var versions = getVersions(fileId, emails)
        shareFileToUsers(
            emails,
            fileId,
            fileName,
            expiration,
            versions,
            message
        )
    } else {
        // Find folder where files are
        var { id: zipId, name: zipName } = zipFiles(fileIds)
        shareZipToUsers(emails, zipId, zipName, expiration, message)
    }
}

function zipFiles(fileIds: string[]) {
    var firstFile = DriveApp.getFileById(fileIds[0])
    var folder = firstFile.getParents().next()
    var names: Array<string> = []
    var file = folder.createFile(
        Utilities.zip(
            fileIds.map((id, i) => {
                var file_ = DriveApp.getFileById(id)
                var name = file_.getName()
                names.push(name)
                return file_.getBlob().setName(name)
            }),
            names.join('<>') + '.zip'
        )
    )
    console.log('🚀 Zip File Created', file.getId())
    return { id: file.getId(), name: file.getName() }
}

function shareFileToUsers(
    emails: string[],
    fileId: string,
    fileName: string,
    expiration: string,
    versions: number[],
    extraMessage: string
) {
    var date = new Date(expiration)

    emails.forEach((email, i) => {
        let body = customMessage(
            extraMessage,
            expiration,
            fileName,
            versions[i]
        )
        shareFileToUser(email, fileId, body, date)
    })

    modifySpread([
        fileId,
        fileName,
        emails.toString(),
        date.toISOString(),
        'false',
    ])
}

function shareZipToUsers(
    emails: string[],
    fileId: string,
    fileName: string,
    expiration: string,
    extraMessage: string
) {
    var date = new Date(expiration)
    emails.forEach((email) => {
        let body = customMessage(extraMessage, expiration, fileName)
        shareFileToUser(email, fileId, body, date)
    })

    modifySpread([
        fileId,
        fileName,
        emails.toString(),
        date.toISOString(),
        'false',
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
