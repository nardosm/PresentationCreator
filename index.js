const fs = require('fs');
const readline = require('readline');
const { google } = require('googleapis');
const util = require('util')


const express = require('express');
const app = express();
const bodyParser = require("body-parser");
const PORT = process.env.PORT || 8000;

app.use(bodyParser.text());
app.use('/viber/webhook', bot.middleware());



const ViberBot = require('viber-bot').Bot,
  BotEvents = require('viber-bot').Events,
  TextMessage = require('viber-bot').Message.Text;


/*
  This is mainly for local testing. Call this endpoint with the 
  '/create' path to process message sent in the post body
*/

app.post('/create', function (request, response) {
  createPresentation(request.body)
  response.end("yes");
});

//We're getting the Viber Bot Token from the environment variable. 
if (!process.env.BOT_TOKEN) {
  console.log('Could not find bot account token key');
  return;
}

//We're also getting the Expose URL from the environment variable to register is later to the bot
if (!process.env.EXPOSE_URL) {
  console.log("Could not find exposing url");
  return;
}

//Initialize the bot with the token and other matadata
const bot = new ViberBot({
  authToken: process.env.BOT_TOKEN,
  name: "IEEC Dulles",
  avatar: "http://ieecdulles.com/wp-content/uploads/2017/09/IEEC_Dulles_logo.png"
});

//When a user subscribes to the bot, send this message
bot.on(BotEvents.SUBSCRIBED, response => {
  response.send(new TextMessage(`Hi there ${response.userProfile.name}. I am the ${bot.name} bot! Feel free to send me song lyrics and I'll convert them to a PowerPoint for you`));
});

//When we recieve a message, we grab it and call the function that creates the slides using the Google APIs
bot.on(BotEvents.MESSAGE_RECEIVED, (message, response) => {
  //console.log(`${message.text} from ${response.userProfile.name}`);
  createPresentation(message.text);
  response.send(new TextMessage(`Thanks for your message ${response.userProfile.name}. If this is a song lyrics, I will try my best to prepare the PowerPoint right away! Have a blessed day!`))
})



  /**
   * This function gets the message from viber and does a few things:
   *    - Gets authentication token that allows us to interact with the Google Slides and Drive API
   *    - Splits up message, 
   *    - Duplicates a template presentation
   *    - Modify the template with the song lyrics
   *    - Download the modified presentaiton as a PPT
   *    - Upload the PPT file back to Google Drive
   * 
   * @param {message} String The OAuth2 client to get token for.
   */

function createPresentation(message) {

  //Split up the message into separate arrays by detecting multiple carriage returns through the Regex
  var textArray = message.split(/\n{2,}/);

  console.log(textArray);

  // If modifying these scopes, delete token.json.
  const SCOPES = ['https://www.googleapis.com/auth/presentations', 'https://www.googleapis.com/auth/drive'];
  // The file token.json stores the user's access and refresh tokens, and is
  // created automatically when the authorization flow completes for the first
  // time.
  const TOKEN_PATH = 'token.json';

  // Load client secrets from a local file.
  fs.readFile('credentials.json', (err, content) => {
    if (err) return console.log('Error loading client secret file:', err);
    // Authorize a client with credentials, then call the Google Slides API.
    authorize(JSON.parse(content), createSlide);
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
    rl.question('Enter the code from that page here: ', (code) => {
      rl.close();
      oAuth2Client.getToken(code, (err, token) => {
        if (err) return console.error('Error retrieving access token', err);
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




  function createSlide(auth) {

    const slides = google.slides({ version: 'v1', auth });
    const drive = google.drive({ version: 'v3', auth });
    //Get the current date to pass to the function that gets us next sunday's date
    var date = Date();
    //Call the function that gets us next Sunday's date
    var nextSunday = nextWeekdayDate(date, 7);
    console.log(`File will be named: ${nextSunday}`)
    //MIME type for downloding and uploading files to Drive
    var mime = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

    //Name the file with the date for next sunday
    var body = { name: nextSunday };

    // Duplicate the template file by providing the file ID
    drive.files.copy({
      'fileId': '1ZBFXk9Mn7NwzuuquhEkyvUxWc9kSwtlAvR_tWXpFMa0',
      'resource': body
    }, (err, res) => {


      let requests = []
      //Reverse the array of song lyrics because somehow the API writes the slides in reverse order
      extArray = textArray.reverse();
      for (let index = 0; index < textArray.length; index++) {

        if (err) return console.log('The API returned an error: ' + err);

        /*
          We're pushing the request to an array because it's a best practice 
          not to call the batchupdate API in a loop
        */
        var originalSlideID = "gd5b15f0a3_5_26";
        requests.push(
          {
            duplicateObject: {
              objectId: originalSlideID,
              objectIds: {
                'gd5b15f0a3_5_28': "copiedText_" + index
              }
            },
          },
          {
            deleteText: {
              objectId: "copiedText_" + index,
              textRange: {
                type: 'ALL'
              }
            },
          },
          {
            insertText: {
              objectId: "copiedText_" + index,
              insertionIndex: 0,
              text: textArray[index]
            }
          }

        )
      }
      // Call the batchUpdate API to duplicate slides, delete text, and insert text
      slides.presentations.batchUpdate({
        presentationId: res.data.id,
        resource: {
          requests
        }
      }, (err, res3) => {
        if (err) return console.log('There is an error modifying the slide: ' + err);
      });

      /*
        After we successfully create the slides with the modifications, 
        delete the tempalate slide
      */
      requests = [{
        deleteObject: {
          objectId: originalSlideID
        }
      }]
      slides.presentations.batchUpdate({
        presentationId: res.data.id,
        resource: {
          requests
        }
      }, (err, res3) => {
        if (err) return console.log('There is an error modifying the slide: ' + err);
        console.log(`OUTPUT OF MODIFIED PRESENTATION ${util.inspect(res3.data, false, null, true)}`);

        //We're going to be exporting the file locally to convert it to powerpoint
        var dest = fs.createWriteStream('/tmp/' + nextSunday + '.pptx');

        drive.files.export({
          fileId: res3.data.presentationId,
          mimeType: mime
        }, {
          responseType: 'stream'
        }, function (err, response) {
          if (err) return done(err);

          response.data.on('error', err => {
            done(err);
          }).on('end', () => {
            console.log("Successfully exported the PowerPoint file!!!");

            //Uplod the file back to Drive as a PPT
            var fileMetadata = {
              'name': nextSunday + '.pptx'
            };
            var media = {
              mimeType: mime,
              body: fs.createReadStream('/tmp/' + nextSunday + '.pptx')
            };
            drive.files.create({
              resource: fileMetadata,
              media: media
            }, function (err, file) {
              if (err) {
                // Handle error
                console.error(err);
              } else {
                console.log("Successfully uploaded the PowerPoint file!!!");
              }
            });
          })
            .pipe(dest);
        });

      });

    });
  }


  function nextWeekdayDate(date, day_in_week) {
    const months = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"];
    var ret = new Date(date || new Date());
    ret.setDate(ret.getDate() + (day_in_week - 1 - ret.getDay() + 7) % 7 + 1);
    let formatted_date = ret.getDate() + "-" + months[ret.getMonth()] + "-" + ret.getFullYear()
    return formatted_date;
  }
}


var server = app.listen(PORT, () => {

    //We're registering the Viber bot with the webhook
    bot.setWebhook(
      `${process.env.EXPOSE_URL}/viber/webhook`
    ).catch(error => {
      console.log("Cannot set webhook on following server. Is it running?");
    })
    
  console.log("The PresentationCreator app is listening on %s", PORT)
})