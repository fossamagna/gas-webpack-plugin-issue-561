// import {
//   SheetMeta,
//   NamedRange,
//   SheetsAndNamedRanges,
//   SpreadsheetPropertiesGAS,
// } from './sheetsTypes';

const METADATA_INSTALLED_ON = 'installedOn';
const METADATA_ORIGINAL_SPREADSHEET_ID = 'originalSpreadsheetId';

export const setActiveSheet = (sheetName: string) => {
  SpreadsheetApp.getActive().getSheetByName(sheetName).activate();
};

export const getSheets = () => {
  const sheets = SpreadsheetApp.getActive().getSheets();
  return sheets.map((s) => {
    return {
      name: s.getName(),
      hidden: s.isSheetHidden(),
    };
  });
};

export const setActiveSheetId = (sheetId: number) => {
  const sheets = SpreadsheetApp.getActive().getSheets();
  for (let i = 0; i < sheets.length; i += 1) {
    if (sheets[i].getSheetId() === sheetId) {
      sheets[i].activate();
      return;
    }
  }
};

export const getSpreadsheetValues = (range: string) => {
  return SpreadsheetApp.getActive().getRange(range).getDisplayValues();
};

export const getSpreadsheetNamedRanges = () => {
  const namedRanges = SpreadsheetApp.getActive().getNamedRanges();
  return namedRanges.map((nr) => {
    return {
      name: nr.getName(),
    };
  });
};

/** Lets you get sheets AND named ranges in one call */
export const getSpreadsheetSheetsAndNamedRanges = () => {
  return {
    namedRanges: getSpreadsheetNamedRanges(),
    sheets: getSheets(),
  };
};

const findMetadataByKey = (key: string) => {
  const metadatas = SpreadsheetApp.getActive().getDeveloperMetadata();

  for (let i = 0; i < metadatas.length; i += 1) {
    const m = metadatas[i];
    if (m.getKey() === key) {
      return m;
    }
  }

  return null;
};

const getOriginalSpreadsheetId = () => {
  const metadata = findMetadataByKey(METADATA_ORIGINAL_SPREADSHEET_ID);

  return metadata ? metadata.getValue() : null;
};

const setOriginalSpreadsheetToSelf = () => {
  SpreadsheetApp.getActive().addDeveloperMetadata(
    METADATA_ORIGINAL_SPREADSHEET_ID,
    SpreadsheetApp.getActive().getId()
  );
};

export const markSpreadsheetInstalledOn = () => {
  // When we've been installed, we need to set the originalSpreadsheetId to us, and record the install
  setOriginalSpreadsheetToSelf();

  const date = new Date();
  SpreadsheetApp.getActive().addDeveloperMetadata(
    METADATA_INSTALLED_ON,
    date.valueOf().toString()
  );
};

const getSpreadsheetInstalledOnTimestamp = () => {
  const metadata = findMetadataByKey(METADATA_INSTALLED_ON);
  if (!metadata) {
    return null;
  }

  const timestampStr = metadata.getValue();
  return parseFloat(timestampStr);
};

export const getSpreadsheetProperties = () => {
  const id = SpreadsheetApp.getActive().getId();
  const timezone = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const originalSpreadsheetId = getOriginalSpreadsheetId();
  const installedOnTimestamp = getSpreadsheetInstalledOnTimestamp();

  return {
    id,
    timezone,
    originalSpreadsheetId,
    installedOnTimestamp,
  };
};

/**
 * Used by API Pipeline to show request errors within sheets
 */
export const API_ERROR = (message: string) => {
  throw new Error(message);
};
