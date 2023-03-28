// This Google Sheets script will post to a slack channel when a user submits data to a Google Forms Spreadsheet
// View the README for installation instructions. Don't forget to add the required slack information below.

// Source: https://github.com/markfguerra/google-forms-to-slack

/////////////////////////
// Begin customization //
/////////////////////////

// Alter this to match the incoming webhook url provided by Slack
var slackIncomingWebhookUrl = 'https://hooks.slack.com/services/YOUR-URL-HERE';

// Include # for public channels, omit it for private channels
var postChannel = "YOUR-CHANNEL-HERE";

var postIcon = ":mailbox_with_mail:";
var postUser = "Form Response";
var postColor = "#0000DD";

var messageFallback = "The attachment must be viewed as plain text.";
var messagePretext = "A user submitted a response to the form.";

///////////////////////
// End customization //
///////////////////////

// In the Script Editor, run initialize() at least once to make your code execute on form submit
function initialize() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i in triggers) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  ScriptApp.newTrigger("submitValuesToSlack")
    .forSpreadsheet(FormApp.getActiveForm().getDestinationId())
    .onFormSubmit()
    .create();

  // ScriptApp.newTrigger('copySpreadSheet').forForm(FormApp.getActiveForm()).onFormSubmit().create();
  ScriptApp.newTrigger("submitValuesToSlackBlocks")
    .forForm(FormApp.getActiveForm())
    .onFormSubmit()
    .create();
}


function getGridItemLabeledValue(item, response) {
  var idx = response.findIndex(i => i != null);

  var gridItem = item.asGridItem()

  // edge, sleepsuite, etc
  var rows = gridItem.getRows();
  // monthly, annual
  // var cols = gridItem.getColumns()

  return [rows[idx], response[idx],];
}

function submitValuesToSlackBlocks(e) {
  // MARK: START DEBUGGING
  // var formDebug = FormApp.openById('{formId}');
  // var formResponses = formDebug.getResponses();
  // e = { response: formResponses[formResponses.length - 1] };
  // MARK: END DEBUGGING

  var form = FormApp.getActiveForm();
  var response = e.response;

  // Create the message to send to Slack
  var slackMessage = {
    "blocks": [
      {
        "type": "section",
        "text": {
          "type": "mrkdwn",
          "text": "*New form submission*"
        }
      }
    ]
  };

  // Loop through all the form items and add them to the message
  var items = response.getItemResponses();
  for (var i = 0; i < items.length; i++) {
    var item = items[i].getItem();
    var value = items[i].getResponse();

    if (item.getType() == FormApp.ItemType.GRID) {
      // TODO: multiple grid items
      var formItem = form.getItems(FormApp.ItemType.GRID)[0];
      var labeledValue = getGridItemLabeledValue(formItem, value);
      value = labeledValue.join(" ");
    }

    var block = {
      "type": "section",
      "fields": [
        {
          "type": "mrkdwn",
          "text": "*" + item.getTitle() + ":*"
        },
        {
          "type": "mrkdwn",
          "text": value || " "
        }
      ]
    };

    slackMessage.blocks.push(block);
  }

  slackMessage.blocks.push({
			"type": "context",
			"elements": [
				{
					"type": "plain_text",
					"text": "Source: Google Form",
					"emoji": true
				}
			]
		})

  var options = {
    'method': 'post',
    'payload': JSON.stringify(slackMessage)
  };

  var response = UrlFetchApp.fetch(slackIncomingWebhookUrl, options);
}

// Running the code in initialize() will cause this function to be triggered this on every Form Submit
function submitValuesToSlack(e) {
  // Test code. uncomment to debug in Google Script editor
  // if (typeof e === "undefined") {
  //   e = {namedValues: {"Question1": ["answer1"], "Question2" : ["answer2"]}};
  //   messagePretext = "Debugging our Sheets to Slack integration";
  // }

  var attachments = constructAttachments(e.values);
  // var blocks = constructBlocks(e.values);

  var payload = {
    "channel": postChannel,
    "username": postUser,
    "icon_emoji": postIcon,
    "link_names": 1,
    "attachments": attachments,
  };

  var options = {
    'method': 'post',
    'payload': JSON.stringify(payload)
  };

  var response = UrlFetchApp.fetch(slackIncomingWebhookUrl, options);
}

// Creates Slack message attachments which contain the data from the Google Form
// submission, which is passed in as a parameter
// https://api.slack.com/docs/message-attachments
var constructAttachments = function (values) {
  var fields = makeFields(values);

  var attachments = [{
    "fallback": messageFallback,
    "pretext": messagePretext,
    "mrkdwn_in": ["pretext"],
    "color": postColor,
    "fields": fields
  }]

  return attachments;
}

// Creates an array of Slack fields containing the questions and answers
var makeFields = function (values) {
  var fields = [];

  var columnNames = getColumnNames();

  for (var i = 0; i < columnNames.length; i++) {
    var colName = columnNames[i];
    var val = values[i];
    fields.push(makeField(colName, val));
  }

  return fields;
}

// Creates a Slack field for your message
// https://api.slack.com/docs/message-attachments#fields
var makeField = function (question, answer) {
  var field = {
    "title": question,
    "value": answer,
    "short": false
  };
  return field;
}

// Extracts the column names from the first row of the spreadsheet
var getColumnNames = function () {
  var sheet = SpreadsheetApp.getActiveSheet();

  // Get the header row using A1 notation
  var headerRow = sheet.getRange("1:1");

  // Extract the values from it
  var headerRowValues = headerRow.getValues()[0];

  return headerRowValues;
}
