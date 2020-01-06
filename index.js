const fs = require('fs');
const readline = require('readline');
const { google } = require('googleapis');
const util = require('util')

const express = require('express');
const app = express();
const bodyParser = require("body-parser");
const port = 8000;


//app.use(bodyParser.urlencoded({extended: false}));
app.use(bodyParser.text());

app.post('/createSlides', (request, response) => {

  var textArray = request.body.split(/\n{2,}/);

  console.log(textArray);
  response.end("yes");
  //var replacementText = request.body.replacementText


  

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
    console.log(`The date for next Sunday is ${nextSunday}`)

    //Name the file with the date for next sunday
    var body = { name: nextSunday };



    /*
      here is where the magic happens:
        - Get next Sunday's date
        - Supply the file name using that date 
        - Copy the template file naming it with Sunday's date
        - Get the full text from the request endpoint (Viber message in the future)
        - Then modify the copy with the messages
        - Write file back to drive
        - Export file as PowerPoint file
        - Future scope 
          - Generate a word doc from the Viber message and print it
    
    */




    




    // Copy the file
    drive.files.copy({
      'fileId': '1ZBFXk9Mn7NwzuuquhEkyvUxWc9kSwtlAvR_tWXpFMa0',
      'resource': body
    }, (err, res) => {


      textArray.forEach((replacementText, index) => {
        
      


      var originalSlideID = "gd5b15f0a3_5_26";
      let requests = [{
        duplicateObject: {
          objectId: originalSlideID,
          objectIds: {
            'gd5b15f0a3_5_28': "copiedText_00" + index
          }
        },

      }]
      if (err) return console.log("ERROR OCCURED");
      console.log(res.data.id);
      //Get the file ID of the copied file and modify the slides

      slides.presentations.batchUpdate({
        presentationId: res.data.id,
        resource: {
          requests
        }
      }, (err, res2) => {
        if (err) return console.log('The API returned an error: ' + err);
        console.log(util.inspect(res2.data, false, null, true))

        let requests = [
          {
            deleteText: {
              objectId: "copiedText_00" + index,
              textRange: {
                type: 'ALL'
              }
            },
          },
          {
            insertText: {
              objectId: "copiedText_00" + index,
              insertionIndex: 0,
              text: replacementText
            }
          }
        ]


        slides.presentations.batchUpdate({
          presentationId: res.data.id,
          resource: {

            requests


          }
        }, (err, res3) => {
          if (err) return console.log('There is an error modifying the slide: ' + err);
          //response.json(util.inspect(res3.data, false, null, true));
        });
      })



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

  

});


var server = app.listen(port, function () {

  
  console.log("Example app listening on %s",  port)
})