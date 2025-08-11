function sendReminderEmails() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(2, 1, lastRow - 1, 16).getValues(); // Assuming data starts from row 2 and includes column P.

  // Create an object to group emails by CC address
  var groupedEmails = {};

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var email = row[9]; // Email address from column J

    // Fetch the email address from column P based on the email address in column J
    var ccEmailColumnJ = row[9]; // Email address from column J
    var ccEmailColumnP = row[15]; // Use a function to perform the lookup

    var ccEmail = ccEmailColumnJ + ", " + ccEmailColumnP; // Add the email addresses from columns J and P to CC
    var toEmail = row[14]; // Email address from column O for TO
    var daysRemaining = row[10]; // Value in column K
    var managerName = row[13]; // Value in column N (نام مدیر دپارتمان)
    var teamManagerName = row[8]; // Value in column I (نام مدیر مستقیم )
    var teamName = row[11]; // Value in column L

    if (daysRemaining !== "" && daysRemaining < 51) {
      // Create or update the grouped email message
      groupedEmails[ccEmail] = groupedEmails[ccEmail] || {
        subject: "یادآور تمدید قرارداد",
        message: `${teamManagerName} عزیز سلام و وقت بخیر<br>
                  <br>این ایمیل به صورت اتوماتیک برای یادآوری تمدید قرارداد تعدادی از افراد دپارتمان‌ شما ارسال شده است<br><br>
                  <table style="border-collapse: collapse; width: 100%;">
                    <tr style="background-color: #000; color: #fff;">
                      <th style="border: 1px solid black; padding: 8px; text-align: center;">ردیف</th>
                      <th style="border: 1px solid black; padding: 8px; text-align: center;">نام و نام خانوادگی</th>
                      <th style="border: 1px solid black; padding: 8px; text-align: center;">کد پرسنلی</th>
                      <th style="border: 1px solid black; padding: 8px; text-align: center;">تاربخ اتمام قرارداد</th>
                      <th style="border: 1px solid black; padding: 8px; text-align: center;">نام تیم</th>
                      <th style="border: 1px solid black; padding: 8px; text-align: center;">مدت زمان تمدید</th>
                    </tr>`,
        cc: ccEmail, // Use the combined email addresses for CC
        to: toEmail, // Use the email from column O for TO
        rowIndex: 1 // Initialize the row index for this email
      };

      // Check if daysRemaining is negative and change cell color and text color
      var cellColor = daysRemaining < 0 ? "#FF0043" : "#FFFFFF"; // Red color for negative daysRemaining
      var textColor = daysRemaining < 0 ? "#FFFFFF" : "#000000"; // White text for negative daysRemaining

      // Add the data for this row to the table with cell color and text color
      groupedEmails[ccEmail].message += `<tr>
                                            <td style="border: 1px solid black; padding: 8px; text-align: center;">${groupedEmails[ccEmail].rowIndex}</td>
                                            <td style="border: 1px solid black; padding: 8px; text-align: center;">${row[2]}</td>
                                            <td style="border: 1px solid black; padding: 8px; text-align: center;">${row[1]}</td>
                                            <td style="border: 1px solid black; padding: 8px; text-align: center; background-color: ${cellColor}; color: ${textColor};">${row[10]}</td>
                                            <td style="border: 1px solid black; padding: 8px; text-align: center;">${teamName}</td>
                                            <td style="border: 1px solid black; padding: 8px; text-align: center;">---------</td>
                                          </tr>`;
      
      // Increment the row index for the next row in this email
      groupedEmails[ccEmail].rowIndex++;
    }
  }

  // Close the table and send the rest of the email
  for (var ccEmail in groupedEmails) {
    groupedEmails[ccEmail].message += `</table><br>
                                       اعداد منفی در لیست، نشان‌دهنده‌ی تعداد روزهایی است که از اتمام قرارداد افراد نامبرده سپری شده است<br><br>
                                       لطفاً بعد از انجام هماهنگی‌های لازم با مدیران میانی خود، مدت زمان تمدید قرارداد افراد فوق را از بین یکی از حالت‌های «یک ماهه، سه ماهه و شش ماهه» انتخاب و نتیجه را از طریق ریپلای به «همه‌ی» مخاطبان این ایمیل، اعلام نمایید<br><br>
                                       لطفا در صورت استعفا  یا قطع همکاری ، حتما تیکت خروج برای هر پرسنل ثبت شود.<br><br>
                                       با تشکر و احترام<br>`;

    // Send the email with cc as ccEmail and "me.moradi@AZKI.Com"
    MailApp.sendEmail({
      to: groupedEmails[ccEmail].to,
      cc: ccEmail + ", me.moradi@AZKI.Com",
      subject: groupedEmails[ccEmail].subject,
      body: groupedEmails[ccEmail].message,
      htmlBody: groupedEmails[ccEmail].message // Use htmlBody to send HTML-formatted email
    });
  }
}

// Function to perform VLOOKUP
function getEmailFromColumnP(sheet, emailJ) {
  var range = sheet.getRange(1, 10, sheet.getLastRow(), 2); // Assuming email addresses are in columns J and P
  var values = range.getValues();

  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === emailJ) {
      return values[i][1]; // Return the corresponding value from column P
    }
  }

  return ""; // Return an empty string if no match is found
}


