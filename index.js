/**
 * The purpose of this application/viber bot is to create a plain text file with the song lyrics sent from Viber to the bot. 
 * This would eliminate copying and emailing the songs, then copying them into ProPresenter. This way the bot will create a plain text
 * file and upload it to Google Drive. Since we have Google Drive synced as local folder on the ProPresenter computer, then it will be
 * just a matter of click on File > Import > File... on ProPresenter and import the plain text file created by the Bot.
 * 
 * USAGE:
 * Set the Bot call back URL by doing HTTP call to the /set-webhook path with the body as follows for bot hosted on render :
 * {
 *  "url":"https://gbec-lyrics-creator.onrender.com"
 * } 
 * 
 * Set the Bot call back URL by doing HTTP call to the /set-webhook path with the body as follows for bot running locally :
 * 
 * {
 *  "url":"http://localhost:9000"
 * } 
 * 
 *
 * Create the following environment variables:
 *    BOT_TOKEN = <Token of the Bot retrieved from Viber Bot Admin Console>
 * 
 * If the code is running on a web service (Heroku, Render, etc), create environment variables on those services
 * 
 * Once environment variables are assigned, start the server locally by running "node index.js". 
 * Use Postman or a simliar applications to issue an HTTP request on http://localhost:9000 with body as the lyrics
 */

const fs = require('fs');
const readline = require('readline');
const { google } = require('googleapis');
const util = require('util')
const path = require('path');
const {authenticate} = require('@google-cloud/local-auth');

const express = require('express');
const app = express();
const bodyParser = require("body-parser");
const { file } = require('googleapis/build/src/apis/file');
const PORT = process.env.PORT || 9000;

/**************************** GOOGLE DRIVE STUFF *******************************/

// If modifying these scopes, delete token.json.
const SCOPES = ['https://www.googleapis.com/auth/drive'];
// The file token.json stores the user's access and refresh tokens, and is
// created automatically when the authorization flow completes for the first
// time.
const TOKEN_PATH = 'token.json';
const CREDENTIALS_PATH = path.join(process.cwd(), 'credentials.json');

app.use(bodyParser.text());
app.use(bodyParser.json());

var message = '';

const ViberBot = require('viber-bot').Bot,
  BotEvents = require('viber-bot').Events,
  TextMessage = require('viber-bot').Message.Text;
  UrlMessage = require('viber-bot').Message.Url;

//Initialize the bot with the token and other matadata
const bot = new ViberBot({
  authToken: process.env.BOT_TOKEN,
  name: "GBEC",
  avatar: "http://ieecdulles.com/wp-content/uploads/2017/09/IEEC_Dulles_logo.png"
});

app.use('/viber/webhook', bot.middleware());
/*
  This is mainly for local testing. Call this endpoint with the 
  '/create' path to process message sent in the post body
*/

app.get('/',(req,res) => {
  return res.send("Hi There!")
});

app.post('/set-webhook', (req, res) => {
   //We're registering the Viber bot with the webhook
  console.log(req.body);
  bot.setWebhook(
    `${req.body.url}/viber/webhook`
  ).then(response => {
    console.log(response);
    res.send(response);
  })
  .catch(error => {
    console.log("Cannot set webhook on following server. Is it running?");
    res.send(error);
  })
  
})

app.post('/create', function (request, response) {
  console.log(request.body);
  message = request.body;
  authorize().then(createTextFile).then(() => {
    response.end("yes");
  }).catch(console.error);
  
});

//We're getting the Viber Bot Token from the environment variable. 
if (!process.env.BOT_TOKEN) {
  console.log('Could not find bot account token key');
  return;
} 

//When a user subscribes to the bot, send this message
bot.on(BotEvents.SUBSCRIBED, response => {
  response.send(new TextMessage(`Hi there ${response.userProfile.name}. I am the ${bot.name} bot! Feel free to send me song lyrics and I'll upload it for you`));
});

//When we recieve a message, we grab it and call the function that creates the slides using the Google APIs
bot.on(BotEvents.MESSAGE_RECEIVED, (textMessage, response) => {
  console.log(textMessage);
  message = textMessage.text;
  authorize().then(createTextFile).then(() => {
    response.send(new TextMessage(`Thanks for your message ${response.userProfile.name}. If this is a song lyrics, I will try my best to create a text file and upload it to Google drive right away! Please notify the GBEC Media team once you sent the song lyrics. God Bless You!`))
  }).catch(console.error);
})


/**
 * Reads previously authorized credentials from the save file.
 *
 * @return {Promise<OAuth2Client|null>}
 */
async function loadSavedCredentialsIfExist() {
  try {
    const content = await fs.readFileSync(TOKEN_PATH);
    const credentials = JSON.parse(content);
    return google.auth.fromJSON(credentials);
  } catch (err) {
    return null;
  }
}


/**
 * Serializes credentials to a file comptible with GoogleAUth.fromJSON.
 *
 * @param {OAuth2Client} client
 * @return {Promise<void>}
 */
async function saveCredentials(client) {
  const content = await fs.readFileSync(CREDENTIALS_PATH);
  const keys = JSON.parse(content);
  const key = keys.installed || keys.web;
  const payload = JSON.stringify({
    type: 'authorized_user',
    client_id: key.client_id,
    client_secret: key.client_secret,
    refresh_token: client.credentials.refresh_token,
  });
  await fs.writeFileSync(TOKEN_PATH, payload);
}


/**
 * Load or request or authorization to call APIs.
 *
 */
async function authorize() {
  let client = await loadSavedCredentialsIfExist();
  if (client) {
    return client;
  }
  client = await authenticate({
    scopes: SCOPES,
    keyfilePath: CREDENTIALS_PATH,
  });
  if (client.credentials) {
    await saveCredentials(client);
  }
  return client;
}


async function createTextFile(authClient) {
  const drive = google.drive({version: 'v3', auth: authClient});
  //Get the current date to pass to the function that gets us next sunday's date
  var date = Date();
  //Call the function that gets us next Sunday's date
  var nextSunday = nextWeekdayDate(date, 7);
  console.log(`File will be named: ${nextSunday}`)
  //MIME type for downloding and uploading files to Drive

  console.log(message);

  fs.writeFile(`./${nextSunday}.txt`, message, err => {
    if(err) {
      console.log(err);
    }
    console.log("FILE WRITTEN TO DISK SUCCESSFULLY");
  })

  const requestBody = {
    name: `${nextSunday}.txt`,
    parents: ['1hNdAzLjcSDN8-h1P7NFtHmqusXH1KpMY'],
    fields: 'id',
  };

  const media = {
    mimeType: 'text/plain',
    body: fs.createReadStream(`${nextSunday}.txt`),
  };

  try {
    const file = await drive.files.create({
      requestBody,
      media: media,
    })
    console.log('FILE UPLOADED TO DRIVE WITH ID: ', file.data.id);
    fs.unlinkSync(`${nextSunday}.txt`);
    return file.data.id;
  }
  catch(err) {
    throw err;
  }
}

function nextWeekdayDate(date, day_in_week) {
  const months = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"];
  var ret = new Date(date || new Date());
  ret.setDate(ret.getDate() + (day_in_week - 1 - ret.getDay() + 7) % 7 + 1);
  let formatted_date = ret.getDate() + "-" + months[ret.getMonth()] + "-" + ret.getFullYear()
  return formatted_date;
}


var server = app.listen(PORT, () => {
  console.log("The lyrics creator app is listening on %s", PORT)
})