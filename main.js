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
});

app.get('/view/:spreadsheetName/:worksheetName', function (req, res) {

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

  var spreadsheetName = req.params.spreadsheetName || 'Hours (Responses)'
  var worksheetName = req.params.worksheetName || "Summer";

  Spreadsheet.load({
    debug: true,
    spreadsheetName: spreadsheetName,
    worksheetName: worksheetName,
    accessToken : {
      type: 'Bearer',
      token: authTokens.access_token
    }
  }, function sheetReady(err, spreadsheet) {
    if (err) throw (err);
    spreadsheet.receive({getValues: true}, function(err, values, info) {
      if(err) throw err;
      spreadsheet.receive({getValues: false}, function(err, formulas, info) {
        if(err) throw err;

        var info = {
          spreadsheetName: spreadsheetName,
          worksheetName: worksheetName,
          lastRow: info.lastRow
        }

        var fullData = getFullData(spreadsheet, info, formulas, values, res);

        createForm(spreadsheet, info, fullData, res);
      });
    });
  });
});

function getFullData(spreadsheet, info, formulas, values, res) {
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

  return fullData;
}

function createForm(spreadsheet, info, fullData, res) {
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
        // This row is a calculated value (output)
        outputs.push({
          field: field
        });
        outLocations[field] = {row: r.toString(), column: "2"};
    } else {
        // This row is an input
      inputs.push({
        field: field,
        type: "text"
      });
      inpLocations[field] = {row: r.toString(), column: "2"};
    }
  }

  function escapeRegExp(str) {
    return str.replace(/[\-\[\]\/\{\}\(\)\*\+\?\.\\\^\$\|]/g, "\\$&");
  }

  var url = new RegExp(escapeRegExp("/view/"+encodeURIComponent(info.spreadsheetName)+"/"+encodeURIComponent(info.worksheetName)));
  console.log(url);

  app.post(url, function (req, res) {
    console.log("Recieved");
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

  res.render("home", { inputs: inputs})

}

app.listen(3000)
