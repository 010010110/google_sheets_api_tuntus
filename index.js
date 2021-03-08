const fs = require('fs');
const readline = require('readline');
const { google } = require('googleapis');

const SHEET_ID = '1l-HAlyjpheUF7jLNevsDPQcTb2PBLJZpX05D8W46SU4';

// If modifying these scopes, delete token.json.
const SCOPES = ['https://www.googleapis.com/auth/spreadsheets'];
// The file token.json stores the user's access and refresh tokens, and is
// created automatically when the authorization flow completes for the first
// time.
const TOKEN_PATH = 'token.json';

// Load client secrets from a local file.
fs.readFile('credentials.json', (err, content) => {
    if (err) return console.log('Error loading client secret file:', err);
    // Authorize a client with credentials, then call the Google Sheets API.
    authorize(JSON.parse(content), updateGrades, listAll);
});

/**
 * Create an OAuth2 client with the given credentials, and then execute the
 * given callback function.
 * @param {Object} credentials The authorization client credentials.
 * @param {function} callback The callback to call with the authorized client.
 */
function authorize(credentials, callback) {
    const { client_secret, client_id, redirect_uris } = credentials.installed;
    const oAuth2Client = new google.auth.OAuth2(
        client_id, client_secret, redirect_uris[0]);

    // Check if we have previously stored a token.
    fs.readFile(TOKEN_PATH, (err, token) => {
        if (err) return getNewToken(oAuth2Client, callback);
        oAuth2Client.setCredentials(JSON.parse(token));
        callback(oAuth2Client);
    });
}

/**
 * Get and store new token after prompting for user authorization, and then
 * execute the given callback with the authorized OAuth2 client.
 * @param {google.auth.OAuth2} oAuth2Client The OAuth2 client to get token for.
 * @param {getEventsCallback} callback The callback for the authorized client.
 */
function getNewToken(oAuth2Client, callback) {
    const authUrl = oAuth2Client.generateAuthUrl({
        access_type: 'offline',
        scope: SCOPES,
    });
    console.log('Authorize this app by visiting this url:', authUrl);
    const rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout,
    });
    rl.question('\nEnter the code from that page here: ', (code) => {
        rl.close();
        oAuth2Client.getToken(code, (err, token) => {
            if (err) return console.error('Error while trying to retrieve access token', err);
            oAuth2Client.setCredentials(token);
            // Store the token to disk for later program executions
            fs.writeFile(TOKEN_PATH, JSON.stringify(token), (err) => {
                if (err) return console.error(err);
                console.log('Token stored to', TOKEN_PATH);
            });
            callback(oAuth2Client);
        });
    });
}

/**
 * @see https://docs.google.com/spreadsheets/d/1l-HAlyjpheUF7jLNevsDPQcTb2PBLJZpX05D8W46SU4/edit
 * @param {google.auth.OAuth2} auth The authenticated Google OAuth client.
 */

function updateGrades(auth) {
    const sheets = google.sheets({ version: 'v4', auth });
    sheets.spreadsheets.values.get({
        spreadsheetId: SHEET_ID,
        range: 'A4:F',
    }, (err, res) => {
        if (err) return console.log('The API returned an error: ' + err);
        const rows = res.data.values;

        if (rows.length) {
            let valuesCell = [];

             // through the table and set new informations
            rows.map((row) => {
                let values = [];

                if (row[2] > 60 * 0.25) {
                    values.push('Reprovado por Falta');
                    values.push(0);
                    valuesCell.push(values);
                    return;
                }
                // set into average value
                const average = (parseInt(row[3]) + parseInt(row[4]) + parseInt(row[5])) / 3;

                if (average < 50) {
                    values.push('Reprovado por Nota');
                    values.push(0);
                    valuesCell.push(values);
                    return;
                } else if (average > 70) {
                    values.push('Aprovado');
                    values.push(0);
                    valuesCell.push(values);
                    return;
                } else {
                    values.push('Exame Final');
                    let naf = Math.ceil((2 * 50) - average, 2);
                    values.push(naf);
                    valuesCell.push(values);
                }
            });

            // create const resource with the information set into valuesCell
            const resource = { values: valuesCell };

            // Update configs
            const updateOptions = {
                spreadsheetId: SHEET_ID,
                range: '!G4:H',
                valueInputOption: 'USER_ENTERED',
                resource
            };

            // Do Update and call function to make table of all spreadsheet
            sheets.spreadsheets.values.update(
                updateOptions,
                (err, result) => {
                    if (err) {
                        console.log(err);
                    } else {
                        listAll(auth);
                    }
                });
        } else {
            console.log('No data found.');
        }
    });

}

//List All spredsheet after update
function listAll(auth) {
    const sheets = google.sheets({ version: 'v4', auth });
    sheets.spreadsheets.values.get({
        spreadsheetId: SHEET_ID,
        range: 'engenharia_de_software!A4:H',
    }, (err, res) => {
        if (err) return console.log('The API returned an error: ' + err);
        const rows = res.data.values;
        if (rows.length) {
            // Print table of the entire spreadsheet
            console.table(rows);
        } else {
            console.log('No data found.');
        }
    });
}
