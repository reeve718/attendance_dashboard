# Streamlit app for attendance visualization

## Google App Script Setup


1. Lanuch Google Apps Script<br>
<img src="images/1.jpg" class="w-100" />

2. Find the doc ID on the link of the goolg spreadsheet. <br> <img src="images/8.jpg" class="w-100" />
Replace YOUR_DOC_ID and paste below code to the GS file and .
<img src="images/2.jpg" class="w-100" />

```bash
function doGet() {
    var spreadsheetId = "YOUR_DOC_ID";
    var ss = SpreadsheetApp.openById(spreadsheetId);
    var sheet = ss.getSheets()[0];
    var data = sheet.getDataRange().getValues();

      var response = {
        data: data
    };

    return ContentService
        .createTextOutput(JSON.stringify(response))
        .setMimeType(ContentService.MimeType.JSON);
}
```


3. Click below <br>
<img src="images/3.jpg" class="w-100" />

4. Select type <br>
<img src="images/4.jpg" class="w-100" />

5. Select type <br>
<img src="images/5.jpg" class="w-100" />

6. Grant permissions <br>
<img src="images/6.jpg" class="w-100" />

7. Copy the webpage link <br>
<img src="images/7.jpg" class="w-100" />