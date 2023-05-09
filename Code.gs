// the recipient variable will store the current email it needs to fill out while it is executing the program
// the email sent variable will confirm that the email has been sent
const RECIPIENT_COL  = "Recipient";
const EMAIL_SENT_COL = "Email Status"; 


// when the spreadsheet is opened up, another dropdown beside 'Help" will be created titled Deploy grades from which this script can be run from
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Deploy Grades')
      .addItem('Send Emails', 'sendEmails')
      .addToUi();
}

function sendEmails(subjectLine, sheet=SpreadsheetApp.getActiveSheet()) {
  // option to skip browser prompt if you want to use this code in other projects
  if (!subjectLine){
    subjectLine = Browser.inputBox("Mail Merge", 
                                      "Type or copy/paste the subject line of the Gmail " +
                                      "draft message you would like to mail merge with:",
                                      Browser.Buttons.OK_CANCEL);
                                      
    if (subjectLine === "cancel" || subjectLine == ""){ 
    // If no subject line, finishes up
    return;
    }
  }



  // this variable will get the gmail template that has been prewritten
  const emailTemplate = getGmailTemplateFromDrafts_(subjectLine);

  // Gets the data from the passed sheet
  const dataRange = sheet.getDataRange();

  // fetches displayed values for each row in the range
  const data = dataRange.getDisplayValues();

  // Takes into consideration that row 1 contains the appropriate headings
  const heads = data.shift(); 

  // Retrives the index of the email status column
  const emailSentColIdx = heads.indexOf(EMAIL_SENT_COL);


  const obj = data.map(r => (heads.reduce((o, k, i) => (o[k] = r[i] || '', o), {}))); // magic I stole from stack, basically turns array into an object array

  // Creates an array to record sent emails
  const out = [];



  // Loops through all the rows of data
    obj.forEach(function(row, rowIdx){
      // Only sends emails if email_sent cell is blank and not hidden by a filter
      if (row[EMAIL_SENT_COL] == ''){
        
        try {
          const msgObj = fillInTemplateFromObject_(emailTemplate.message, row);

          GmailApp.sendEmail(row[RECIPIENT_COL], msgObj.subject, msgObj.text, {
            htmlBody: msgObj.html,

            attachments: emailTemplate.attachments,
            inlineImages: emailTemplate.inlineImages
          });

          // Edits cell to record email sent date
          out.push([new Date()]);

        } catch(e) {
          // modify cell to record error
          out.push([e.message]);
        }
      } else {
        out.push([row[EMAIL_SENT_COL]]);
      }
    });

  // Updates the sheet with new data
    sheet.getRange(2, emailSentColIdx+1, out.length).setValues(out).setBackground("lightgreen");



    /**
     * Get a Gmail draft message by matching the subject line.
     * @param {string} subject_line to search for draft message
     * @return {object} containing the subject, plain and html message body and attachments
    */
    function getGmailTemplateFromDrafts_(subject_line){
      try {
        // get drafts
        const drafts = GmailApp.getDrafts();
        // filter the drafts that match subject line
        const draft = drafts.filter(subjectFilter_(subject_line))[0];
        // get the message object
        const msg = draft.getMessage();

        // Handles inline images and attachments so they can be included in the merge
        // Gets all attachments and inline image attachments
        const allInlineImages = draft.getMessage().getAttachments({includeInlineImages: true,includeAttachments:false});
        const attachments = draft.getMessage().getAttachments({includeInlineImages: false});
        const htmlBody = msg.getBody(); 

        // Creates an inline image object with the image name as key 
        // (can't rely on image index as array based on insert order)
        const img_obj = allInlineImages.reduce((obj, i) => (obj[i.getName()] = i, obj) ,{});

        //Regexp searches for all img string positions with cid
        const imgexp = RegExp('<img.*?src="cid:(.*?)".*?alt="(.*?)"[^\>]+>', 'g');
        const matches = [...htmlBody.matchAll(imgexp)];

        //Initiates the allInlineImages object
        const inlineImagesObj = {};
        // built an inlineImagesObj from inline image matches
        matches.forEach(match => inlineImagesObj[match[1]] = img_obj[match[2]]);

        return {message: {subject: subject_line, text: msg.getPlainBody(), html:htmlBody}, 
                attachments: attachments, inlineImages: inlineImagesObj };
      } catch(e) {
        throw new Error("Oops - can't find Gmail draft");
      }

      //search for the subject filer via it's header name to input into the gmail message {{name}}
      function subjectFilter_(subject_line){
        return function(element) {
          if (element.getMessage().getSubject() === subject_line) {
            return element;
          }
        }
      }
    }
    
  // All this code has been borrowed from stack and google documentation.
    //fills up the template string with the data object
    function fillInTemplateFromObject_(template, data) {
      // We have two templates one for plain text and the html body
      // Stringifing the object means we can do a global replace
      let template_string = JSON.stringify(template);

      // Token replacement
      template_string = template_string.replace(/{{[^{}]+}}/g, key => {
        return escapeData_(data[key.replace(/[{}]+/g, "")] || "");
      });
      return  JSON.parse(template_string);
    }

  function escapeData_(str) {
      return str
        .replace(/[\\]/g, '\\\\')
        .replace(/[\"]/g, '\\\"')
        .replace(/[\/]/g, '\\/')
        .replace(/[\b]/g, '\\b')
        .replace(/[\f]/g, '\\f')
        .replace(/[\n]/g, '\\n')
        .replace(/[\r]/g, '\\r')
        .replace(/[\t]/g, '\\t');
    };

}






