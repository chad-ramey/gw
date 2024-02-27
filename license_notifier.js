/*
GAS
Work in Progress : February 27, 2024
chad.ramey@onepeloton.com 
Forked from Google Workspace License Notifier by Goldy Arora
https://www.goldyarora.com/blog/license-notifier
*/

// Custom Menu 
// adds a custom menu to your Google Sheets UI, providing easy access to the getLicenses and buildReport functions.
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('ColSys')
    .addItem('Get Licenses', 'getLicenses')
    .addItem('Send Report', 'buildReport')
    .addToUi();
}

// Mapping of SKU IDs to SKU Names
const skuIdToNameMapping = {
  "1010020020": "Google Workspace Enterprise Plus",
  "1010020026": "Google Workspace Enterprise Standard",
  "1010340001": "Google Workspace Enterprise Plus - Archived User",
  "1010340004": "Google Workspace Enterprise Standard - Archived User"
  // ... Add all other necessary SKU mappings here
};

// Fetch license data, write to sheet and check for thresholds
function getLicenses() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("License Notifier");
  const domainName = Session.getActiveUser().getEmail().replace(/.*@/, "");
  let fileArray = [["User Email", "Assigned License"]]; // Header for the data

  const productIds = [
    "Google-Apps",
    "101031",
    "Google-Drive-storage",
    "Google-Vault",
    "101001",
    "101005",
    "101033",
    "101034"
  ];

  productIds.forEach(product => {
    let pageToken;
    do {
      let optionalArgs = { maxResults: 100, pageToken: pageToken };
      try {
        const page = AdminLicenseManager.LicenseAssignments.listForProduct(product, domainName, optionalArgs);
        page.items.forEach(prod => {
          const skuName = skuIdToNameMapping[prod.skuId] || prod.skuId; // Map SKU ID to name, or keep the ID if not found in mapping
          fileArray.push([prod.userId, skuName]);
        });
        pageToken = page.nextPageToken;
      } catch (error) {
        Logger.log("Error fetching license data for product " + product + ": " + error);
        return;
      }
    } while (pageToken);
  });

  // Clear the existing content and set the new values to avoid duplicates
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) { // Check if there's content to clear
    sheet.getRange(2, 1, lastRow - 1, 2).clearContent(); // Clear existing content from row 2 onwards
  }
  // Write the new data to the sheet starting from the second row
  sheet.getRange(2, 1, fileArray.length - 1, 2).setValues(fileArray.slice(1)); // Exclude the header when setting values

    checkAndNotifyForLicenses(); // Check for license thresholds and notify if needed
}

// Define thresholds for each license type
const licenseThresholds = {
    "Google Workspace Enterprise Plus": 20,
    "Google Workspace Enterprise Standard": 20,
    "Google Workspace Enterprise Plus - Archived User": 0,
    "Google Workspace Enterprise Standard - Archived User": 0
    // Add more license types and their thresholds as needed
  };
  
/// Check licenses against their thresholds and notify via Slack if needed
function checkAndNotifyForLicenses() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("License Notifier");
  const licenseDataRange = sheet.getRange("E2:F" + sheet.getLastRow());
  const licenseData = licenseDataRange.getValues();
  let alertsToSend = [];

  licenseData.forEach(row => {
    const licenseType = row[0];
    const licensesLeft = row[1];
    const threshold = licenseThresholds[licenseType];

    if (threshold !== undefined && licensesLeft <= threshold) {
      alertsToSend.push(`${licenseType} has only ${licensesLeft} licenses left, which is at or below the threshold of ${threshold}.`);
    }
  });

  if (alertsToSend.length > 0) {
    sendAlertToSlack(alertsToSend.join("\n"));
  }
}

// Function to send alerts to Slack
function sendAlertToSlack(message) {
  const payload = {
    "blocks": [
      {
        "type": "section",
        "text": {
          "type": "mrkdwn",
          "text": message
        }
      }
    ]
  };
  sendAlert(payload);
}

// Send alert to multiple Slack channels
function sendAlert(payload) {
  const webhooks = [
   //   "https://hooks.slack.com/services/", // channel-name
  ];

  var options = {
    "method": "post", 
    "contentType": "application/json", 
    "muteHttpExceptions": true, 
    "payload": JSON.stringify(payload) 
  };

  webhooks.forEach(webhook => {
    try {
      UrlFetchApp.fetch(webhook, options);
    } catch(e) {
      Logger.log("Error sending alert to Slack: " + e);
    }
  });
}
  
  // Build and send report to Slack channels
  function buildReport() {
    const ss = SpreadsheetApp.getActive();
    let data = ss.getSheetByName('License Notifier').getRange("E1:F5").getValues();
    let payload = buildAlert(data);
    sendAlert(payload);
  }
  
  // Construct the alert payload for Slack
  function buildAlert(data) {
    let totalLicense = data[0][1];
    let licenseBreakdown = data.slice(1).map(function(row) {
      return row[0] + ": " + row[1];
    }).join("\n");
  
    let payload = {
      "blocks": [
        {
          "type": "section",
          "text": {
            "type": "mrkdwn",
            "text": ":ice_cube::robot_face::google: *Available GW Licenses* :google::robot_face::ice_cube:"
          }
        },
        {
          "type": "divider"
        },
        {
          "type": "section",
          "text": {
            "type": "mrkdwn",
            "text": licenseBreakdown
          }
        }
      ]
    };
    return payload;
  }
  
  // Send alert to multiple Slack channels
  function sendAlert(payload) {
    const webhooks = [
   //   "https://hooks.slack.com/services/", // channel-name

    ];
  
    var options = {
      "method": "post", 
      "contentType": "application/json", 
      "muteHttpExceptions": true, 
      "payload": JSON.stringify(payload) 
    };
  
    webhooks.forEach(webhook => {
      try {
        UrlFetchApp.fetch(webhook, options);
      } catch(e) {
        Logger.log("Error sending alert to Slack: " + e);
      }
    });
  }
