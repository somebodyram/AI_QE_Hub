namespace exceldata;

entity TestResults {
  ID              : UUID;
  key fileName    : String(500);   // Added 'key' here
  key sheetName   : String(100);   // Added 'key' here
  key rowNumber   : Integer;       // Added 'key' here
  scenario        : String(500);
  country         : String(50);
  actualResult    : String(2000);
  status          : String(20);
  createdAt       : Timestamp;
}
