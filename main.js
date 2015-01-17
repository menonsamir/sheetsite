var express = require('express');
var app = express();
var bodyParser = require('body-parser');
var exphbs = require('express-handlebars');

var parser = require('xml2json');

var google = require('googleapis');
var OAuth2 = google.auth.OAuth2;

var CLIENT_ID = '660626344688-oocqgepv1rr6hdj1gk3a3ciguaqchpjk.apps.googleusercontent.com';
var CLIENT_SECRET = 'aATAwBhzGav7Jn0NlwM_3L7j';
var REDIRECT_URI = 'http://localhost:3000/authorized';

var oauth2Client = new OAuth2(CLIENT_ID, CLIENT_SECRET, REDIRECT_URI);

var scopes = [
  'https://spreadsheets.google.com/feeds'
];

var authTokens = {};

app.use(bodyParser.urlencoded({ extended: false }));
app.engine('handlebars', exphbs({}));
app.set('view engine', 'handlebars');

app.get('/auth', function (req, res) {
  var url = oauth2Client.generateAuthUrl({
    access_type: 'offline', // 'online' (default) or 'offline' (gets refresh_token)
    scope: scopes // If you only need one scope you can pass it as string
  });
  res.redirect(url);
})

app.get('/authorized', function (req, res) {
  console.log(req.query.code);
  oauth2Client.getToken(req.query.code, function(err, tokens) {
    console.log("0.5");
    // Now tokens contains an access_token and an optional refresh_token. Save them.
    authTokens = tokens; 
    if (err)
      console.error(err);
    if(!err) {
      oauth2Client.setCredentials(tokens);
      google.options({ auth: oauth2Client });
      console.log("1");
      res.redirect('/view');
    }
  });
})

app.get('/home', function (req, res) {
  console.log("2");
  var Spreadsheet = require('edit-google-spreadsheet');
  Spreadsheet.load({
    debug: true,
    spreadsheetName: 'Test',
    worksheetName: 'Summer',
 /*   // Choose from 1 of the 3 authentication methods:
    //    1. Username and Password
    username: 'my-name@google.email.com',
    password: 'my-5uper-t0p-secret-password',
    // OR 2. OAuth
    oauth : {
      email: 'my-name@google.email.com',
      keyFile: 'my-private-key.pem'
    },
    // OR 3. Token*/
    accessToken : {
      type: 'Bearer',
      token: authTokens.access_token
    }
  }, function sheetReady(err, spreadsheet) {
    console.log("3");
    if (err)
      console.error(err);
    //use speadsheet!
    spreadsheet.add({ 3: { 5: "hello!" } });

        spreadsheet.send(function(err) {
          if(err) throw err;
          res.send("Updated Cell at row 3, column 5 to 'hello!'");
        });
  });
})

app.get('/view', function (req, res) {

  /*var options = {
    url: "https://spreadsheets.google.com/feeds/spreadsheets/private/full",
    headers: {
      'Authorization: Bearer': authTokens.access_token
    }
  };

  request(options, function(err, resp, body) {
    res.send(body);
  })*/
  var Spreadsheet = require('edit-google-spreadsheet');
  Spreadsheet.load({
    debug: true,
    spreadsheetName: 'Hours (Responses)',
    worksheetName: 'Summer',
    accessToken : {
      type: 'Bearer',
      token: authTokens.access_token
    },
/*    spreadsheetId: "1w5x4fFDc6PGCG0eNNuQCovb2m-0tiqDh7Epb0pIHJIg",
    worksheetId: "od6"*/
  }, function sheetReady(err, spreadsheet) {
    if (err)
      console.error(err);
    //use speadsheet

    spreadsheet.receive({getValues: true}, function(err, values, info) {
      if(err) throw err;
      spreadsheet.receive({getValues: false}, function(err, formulas, info) {
        if(err) throw err;
        process(spreadsheet, info, formulas, values, res)
      });
      // Found rows: { '3': { '5': 'hello!' } }
    });
  });
});

function process(spreadsheet, info, formulas, values, res) {
  var fullData = {};
  for (var row in formulas) {
    fullData[row] = {};
    for (var column in formulas[row]) {
      fullData[row][column] = {
        formula: formulas[row][column],
        value: values[row][column]
      };
      if (formulas[row][column] === values[row][column])
        delete fullData[row][column].formula;
    }
  }

  recognizeForm(spreadsheet, info, fullData, res);
}

function recognizeForm(spreadsheet, info, fullData, res) {
  var title = fullData["1"]["1"].value;
  var inputs = [];
  var outputs = [];

  var inpLocations = {};
  var outLocations = {};
  for (var r = 2; r <= info.lastRow; r += 1) {
    var row = fullData[r.toString()];
    var field = row["1"].value;
    function test(a) {
      if (a.hasOwnProperty("2")) {
        return a["2"].hasOwnProperty("formula");
      } else {
        return false;
      }
    }
    if (test(row)) {
        // Must be a calculated value (output)
        outputs.push({
          field: field
        });
        outLocations[field] = {row: r.toString(), column: "2"};
    } else {
      // Must be an input
      inputs.push({
        field: field,
        type: "text"
      });
      inpLocations[field] = {row: r.toString(), column: "2"};
    }
  }

  app.post("/"+encodeURIComponent(title), function (req, res) {
    var batch = {};
    for (var field in req.body) {
      var value = req.body[field];
      var loc = inpLocations[field];
      batch[loc.row] = {};
      batch[loc.row][loc.column] = value;
    }

    spreadsheet.add(batch);

    spreadsheet.send(function(err) {
      if(err) throw err;
      // getValues true so that we get the calculated values, not the formula
      spreadsheet.receive({getValues: true}, function(err, rows, info) {
        if(err) throw err;
        var outputs = {};
        for (var field in outLocations) {
          var loc = outLocations[field];
          var value = rows[loc.row][loc.column];

          outputs[field] = value;
        }
        console.log(outputs);
        res.send(JSON.stringify(outputs));
      });
    });
  });

  res.render("home", {title: encodeURIComponent(title), inputs: inputs})

}

app.listen(3000)
