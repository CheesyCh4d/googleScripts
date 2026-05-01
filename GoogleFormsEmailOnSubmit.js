function onFormSubmit(e) {
  // Debug: Check what we're receiving
  console.log("Event object: " + JSON.stringify(e));
  
  // Get the form response values - check different possible locations
  var values;
  if (e.values) {
    values = e.values;
    console.log("Using e.values");
  } else if (e.response) {
    values = e.response.getItemResponses().map(function(itemResponse) {
      return itemResponse.getResponse();
    });
    values.unshift(new Date()); // Add timestamp at beginning
    console.log("Using e.response method");
  } else {
    console.error("Cannot find form response data in event object");
    return;
  }
  
  // Debug: Log all values to see the structure
  console.log("Total values received: " + values.length);
  console.log("All values: " + JSON.stringify(values));
  
  if (values.length <= 3) {
    console.error("Not enough form responses received. Expected at least 4 values.");
    return;
  }
  
  // Define which column contains the school selection (adjust this number!)
  // Column numbers start at 0: first question = 0, second question = 1, etc.
  var schoolColumnIndex = 4; // 5th question = index 4
  var schoolSelected = values[schoolColumnIndex];
  
  console.log("School selected (index " + schoolColumnIndex + "): " + schoolSelected);
  
  // Define your email mapping - Ketchikan School District schools
  var emailMap = {
    "Ketchikan High School": {
      emails: ["KHSRegistrar@k21schools.org"],
      schoolName: "Ketchikan High School"
    },
    "Revilla High School": {
      emails: ["RHSRegistrar@k21schools.org"],
      schoolName: "Revilla High School"
    },
    "Fast Track": {
      emails: ["FTSRegistrar@k21schools.org"],
      schoolName: "Fast Track"
    },
    "Schoenbar Middle School": {
      emails: ["SMSRegistrar@k21schools.org"],
      schoolName: "Schoenbar Middle School"
    },
    "Ketchikan Charter School (Elementary and Secondary)": {
      emails: ["KCSRegistrar@k21schools.org"],
      schoolName: "Ketchikan Charter School (Elementary and Secondary)"
    },
    "Tongass School of Arts & Sciences": {
      emails: ["TSASRegistrar@k21schools.org"],
      schoolName: "Tongass School of Arts & Sciences"
    },
    "Point Higgins Elementary School": {
      emails: ["PHERegistrar@k21schools.org"],
      schoolName: "Point Higgins Elementary School"
    },
    "Houghtaling Elementary School": {
      emails: ["HTERegistrar@k21schools.org"],
      schoolName: "Houghtaling Elementary School"
    },
    "Fawn Mountain Elementary School": {
      emails: ["FMERegistrar@k21schools.org"],
      schoolName: "Fawn Mountain Elementary School"
    }
  };
  
  // Check if we have a match for the selected school
  if (emailMap[schoolSelected]) {
    var schoolInfo = emailMap[schoolSelected];
    
    // Build the email content
    var subject = "Updated student information - " + schoolInfo.schoolName;
    var body = buildEmailBody(values, schoolInfo.schoolName);
    
    // Send the email
    try {
      MailApp.sendEmail({
        to: schoolInfo.emails.join(','),
        subject: subject,
        body: body,
        htmlBody: buildHtmlEmailBody(values, schoolInfo.schoolName)
      });
      
      console.log("Email sent successfully to " + schoolInfo.emails.join(', '));
      
    } catch (error) {
      console.error("Error sending email: " + error.toString());
    }
  } else {
    console.log("No email mapping found for school: " + schoolSelected);
  }
}

function buildEmailBody(values, schoolName) {
  // Build a simple notification email with link to responses
  var body = "Hello,\n\n";
  body += "Updated student information has been submitted for " + schoolName + ".\n\n";
  body += "Please review the response in the Google Sheet:\n";
  body += getResponseSheetUrl() + "\n\n";
  body += "Submitted at: " + values[0] + "\n\n";
  body += "Thank you,\n";
  body += "KGBSD\n";
  body += "Written by Chad Jacks";
  
  return body;
}

function buildHtmlEmailBody(values, schoolName) {
  // Build a simple HTML notification email
  var sheetUrl = getResponseSheetUrl();
  var html = "<p>Hello,</p>";
  html += "<p>Updated student information has been submitted for <strong>" + schoolName + "</strong>.</p>";
  html += "<p>Please review the response in the Google Sheet:<br>";
  html += "<a href='" + sheetUrl + "'>View Registration Responses</a></p>";
  html += "<p><em>Submitted at: " + values[0] + "</em></p>";
  html += "<p>Thank you,<br>";
  html += "KGBSD<br>";
  html += "<br>";
  html += "Written by Chad Jacks</p>";
  
  return html;
}

function getResponseSheetUrl() {
  // Get the URL of the Google Sheet that stores form responses
  var form = FormApp.getActiveForm();
  var responseSheetId = form.getDestinationId();
  
  if (responseSheetId) {
    return "https://docs.google.com/spreadsheets/d/" + responseSheetId;
  } else {
    // If no response sheet is set up, return the form's response summary
    return form.getEditUrl().replace('/edit', '/responses');
  }
}