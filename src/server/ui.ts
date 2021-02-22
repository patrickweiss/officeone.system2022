import { copyFolder, DriveConnector, getNextVersion } from "./oo22lib/driveConnector";
import { currentOOversion, ooTables, ooVersions, systemMasterProperty } from "./oo22lib/systemEnums";

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

  //copy own system template folder
  copyFolder(versionFolder.getId(),systemFolder.getId(),currentOOversion,getNextVersion());

  //copy subsystem folder
  const oo2021systemFolderId = office2022systemDC.getMasterProperty(systemMasterProperty.officeOne2021_TemplateFolderId);
 if (oo2021systemFolderId){
   const oo2021versionFolder = DriveApp.getFolderById(oo2021systemFolderId).getFoldersByName(currentOOversion).next();
  const oo2021systemFolder = DriveApp.getFolderById(oo2021systemFolderId);
  copyFolder(oo2021versionFolder.getId(),oo2021systemFolder.getId(),currentOOversion,getNextVersion());
 }

 
  SpreadsheetApp.getUi().alert("Subversion 0005");
};

