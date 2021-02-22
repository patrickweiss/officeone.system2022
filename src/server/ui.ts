import { copyFolder, DriveConnector, getNextVersion } from "./oo22lib/driveConnector";
import { currentOOversion, ooTables, ooVersions } from "./oo22lib/systemEnums";

export const onOpen = () => {
  const menu = SpreadsheetApp.getUi()
    .createMenu('OfficeOne.System') // edit me!
    .addItem('Create new version', 'newVersion')

  menu.addToUi();
};

export const newVersion = () => {
  //Read System Properties
  const office2022systemDC = new DriveConnector(
    SpreadsheetApp.getActive().getId(),
    ooTables.systemMasterConfiguration,
    ooVersions.oo55
  )

  const versionFolder = DriveApp.getFileById(SpreadsheetApp.getActive().getId()).getParents().next();
  const systemFolder = versionFolder.getParents().next();

  copyFolder(versionFolder.getId(),systemFolder.getId(),currentOOversion,getNextVersion())
 
  SpreadsheetApp.getUi().alert("Subversion 0003");
};

