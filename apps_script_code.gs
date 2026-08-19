function doPost(e) {
  try {
    Logger.log("=== Processing Registration ===");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const memberSheet = ss.getSheetByName("Church Member Registrations");
    const childrenSheet = ss.getSheetByName("Children");
    
    if (!memberSheet || !childrenSheet) {
      Logger.log("ERROR: Sheets not found");
      return ContentService.createTextOutput('Error: Sheets not found').setMimeType(ContentService.MimeType.TEXT);
    }
    
    Logger.log("✓ Sheets found");
    
    const data = JSON.parse(e.postData.contents);
    Logger.log("✓ Data parsed: " + data.member.fornamn + " " + data.member.efternamn);
    
    // DEBUG: Log the raw data to see what we're getting
    Logger.log("DEBUG: Full data = " + JSON.stringify(data));
    Logger.log("DEBUG: Children array = " + JSON.stringify(data.children));
    Logger.log("DEBUG: Children length = " + (data.children ? data.children.length : 0));
    
    // Main member data
    const memberRow = [
      data.date,
      data.member.fornamn,
      data.member.efternamn,
      data.member.dopnamn,
      data.member.personnummer,
      data.member.adress,
      data.member.stad,
      data.member.telefon,
      data.member.epost,
      data.spouse ? data.spouse.fornamn : '',
      data.spouse ? data.spouse.efternamn : '',
      data.spouse ? data.spouse.epost : '',
      data.children ? data.children.length : 0,
      data.fee.manadsavgift || '',
      data.fee.regavgift || '',
      data.fee.ovrigt || '',
      data.signature
    ];
    
    try {
      memberSheet.appendRow(memberRow);
      Logger.log("✓ Member row added");
    } catch (err) {
      Logger.log("ERROR adding member: " + err.toString());
      throw err;
    }
    
    // Add children to separate sheet
    const parentName = data.member.fornamn + " " + data.member.efternamn;
    Logger.log("Parent: " + parentName);
    
    if (data.children && Array.isArray(data.children) && data.children.length > 0) {
      Logger.log("Processing " + data.children.length + " children");
      
      for (let i = 0; i < data.children.length; i++) {
        try {
          const child = data.children[i];
          Logger.log("Child " + (i + 1) + " raw data: " + JSON.stringify(child));
          
          // Skip empty children entries
          if (!child.namn || child.namn.trim() === '') {
            Logger.log("Child " + (i + 1) + " skipped (empty)");
            continue;
          }
          
          const childRow = [
            parentName,
            child.namn || '',
            child.dopnamn || '',
            child.personnummer || '',
            child.kon || '',
            data.date
          ];
          
          childrenSheet.appendRow(childRow);
          Logger.log("✓ Child " + (i + 1) + " added: " + child.namn);
        } catch (childErr) {
          Logger.log("ERROR adding child " + (i + 1) + ": " + childErr.toString());
        }
      }
    } else {
      Logger.log("No children to process or children array is empty/not array");
      if (data.children) {
        Logger.log("Children type: " + typeof data.children);
        Logger.log("Is array: " + Array.isArray(data.children));
      }
    }
    
    Logger.log("=== Processing Complete ===");
    return ContentService.createTextOutput('Success').setMimeType(ContentService.MimeType.TEXT);
    
  } catch (error) {
    Logger.log("FATAL ERROR: " + error.toString());
    Logger.log("Stack: " + error.stack);
    return ContentService.createTextOutput('Error: ' + error.toString()).setMimeType(ContentService.MimeType.TEXT);
  }
}

// Test function
function testRegistration() {
  try {
    Logger.log("=== Testing with Mock Data ===");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const memberSheet = ss.getSheetByName("Church Member Registrations");
    const childrenSheet = ss.getSheetByName("Children");
    
    if (!memberSheet) {
      Logger.log("❌ ERROR: 'Church Member Registrations' sheet not found!");
      return;
    }
    if (!childrenSheet) {
      Logger.log("❌ ERROR: 'Children' sheet not found!");
      return;
    }
    
    Logger.log("✓ Both sheets found");
    
    const mockData = {
      date: new Date().toISOString().slice(0, 10),
      forsamling: "Kidist Selassie – Stockholm",
      medlemsnr: "",
      member: {
        fornamn: "Test",
        efternamn: "Parent",
        dopnamn: "Testson",
        personnummer: "900101-1234",
        adress: "Test Street 123",
        stad: "Stockholm",
        telefon: "0701234567",
        epost: "test@example.com"
      },
      spouse: null,
      children: [
        { namn: "Child One", dopnamn: "Christen", personnummer: "200101-1111", kon: "Pojke" },
        { namn: "Child Two", dopnamn: "Christina", personnummer: "200202-2222", kon: "Flicka" },
        { namn: "Child Three", dopnamn: "Christopher", personnummer: "200303-3333", kon: "Pojke" }
      ],
      fee: {
        manadsavgift: "500",
        regavgift: "100",
        ovrigt: ""
      },
      signature: "Test Parent"
    };
    
    const memberRow = [
      mockData.date,
      mockData.member.fornamn,
      mockData.member.efternamn,
      mockData.member.dopnamn,
      mockData.member.personnummer,
      mockData.member.adress,
      mockData.member.stad,
      mockData.member.telefon,
      mockData.member.epost,
      "",
      "",
      "",
      mockData.children.length,
      mockData.fee.manadsavgift,
      mockData.fee.regavgift,
      mockData.fee.ovrigt,
      mockData.signature
    ];
    
    memberSheet.appendRow(memberRow);
    Logger.log("✓ Test member added");
    
    const parentName = mockData.member.fornamn + " " + mockData.member.efternamn;
    
    mockData.children.forEach((child, index) => {
      const childRow = [
        parentName,
        child.namn,
        child.dopnamn,
        child.personnummer,
        child.kon,
        mockData.date
      ];
      childrenSheet.appendRow(childRow);
      Logger.log("✓ Child " + (index + 1) + " added: " + child.namn);
    });
    
    Logger.log("=== Test Complete! Check both sheets ===");
    
  } catch (error) {
    Logger.log("❌ TEST ERROR: " + error.toString());
  }
}
