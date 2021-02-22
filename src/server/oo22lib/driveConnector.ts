import { currentOOversion, office, ooFolders, ooTables, ooVersions, systemMasterId, systemObject } from "./systemEnums";



export class DriveConnector {
    private hostFileId: string;
    private hostTable: ooTables;
    private version: ooVersions;
    public officeFolder: GoogleAppsScript.Drive.Folder;

    private spreadsheetCache: Object = {};
    private tableDataCache: Object = {};
    private ooConfigurationCache: Object = {};
    constructor(hostFileId: string, hostTable: ooTables, version: ooVersions) {
        this.hostFileId = hostFileId;
        this.hostTable = hostTable;
        this.version = version;
        if (!this.systemInstalled)this.installSystem();
    }
    public systemInstalled(): boolean {
        return true;
    }
    public installSystem() {
    }

    public getSheetName(table: ooTables): string { return this.getProperyFromTable(ooTables.systemMasterConfiguration, table + "_TableSheet"); }


    public setSystemObject(systemObject: systemObject, object: Object): void {
        const systemDataTable = this.tableDataCache[ooTables.systemConfiguration] as unknown as any[][]
        const propertyRow = systemDataTable.filter(row => row[0] === systemObject)[0]
        propertyRow[1] = JSON.stringify(object);
        this.getSpreadsheet(ooTables.systemConfiguration)
            .getSheetByName(this.getSheetName(ooTables.systemConfiguration))
            .getDataRange()
            .setValues(systemDataTable);
        SpreadsheetApp.flush()
    }
    public setOfficeProperty(officeProperty: office, value: string): void {
        const officeDataTable = this.tableDataCache[ooTables.officeConfiguration] as unknown as any[][]
        const propertyRow = officeDataTable.filter(row => row[0] === officeProperty)[0]
        propertyRow[1] = value;
        this.getSpreadsheet(ooTables.officeConfiguration)
            .getSheetByName(this.getSheetName(ooTables.officeConfiguration))
            .getDataRange()
            .setValues(officeDataTable);
        SpreadsheetApp.flush()
    }

    private getSpreadsheet(table: ooTables): GoogleAppsScript.Spreadsheet.Spreadsheet {
        const spreadsheet = this.spreadsheetCache[this.getFileName(table)] as unknown as GoogleAppsScript.Spreadsheet.Spreadsheet;
        if (!spreadsheet) {
            throw new Error("implement office spreadsheet caching for " + this.getFileName(table));
        }
        return spreadsheet;
    }

    public getMasterProperty(name: string) { return this.getProperyFromTable(ooTables.systemMasterConfiguration, name); }

    private getValuesCache(table: ooTables) {
        let valuesCache = this.ooConfigurationCache[table];
        if (!valuesCache) {
            console.log("Fill Configuration Cache:" + table);
            const data = this.getTableData(table);
            valuesCache = new ValuesCache(data);
            this.ooConfigurationCache[table] = valuesCache;
        }
        return valuesCache;
    }
    private getProperyFromTable(table: ooTables, propertyName: string): string {
        const property = this.getValuesCache(table).getValueByName(propertyName);
        if (!property) {
            console.log(this.getTableData(table));
            throw new Error("Variable missing in Configuration:" + table + " " + propertyName);
        }
        return property;
    }
    public getTableData(table: ooTables): any[][] {
        let tableData = this.tableDataCache[table] as unknown as any[][];
        console.log("getTableData:" + table);
        if (!tableData && table === ooTables.systemMasterConfiguration) {
            tableData = SpreadsheetApp.openById(systemMasterId).getSheetByName("Configuration").getDataRange().getValues();
            this.tableDataCache[table] = tableData;
            return tableData
        }
        if (!tableData && table === ooTables.systemConfiguration) {
            const sheetName = this.getSheetName(ooTables.systemConfiguration)
            const spreadsheet = this.getSpreadsheet(ooTables.systemConfiguration)
            tableData = spreadsheet
                .getSheetByName(sheetName)
                .getDataRange().getValues();
            this.tableDataCache[table] = tableData;
            return tableData
        }
        if (!tableData) {
            throw new Error("no implementation for " + table);
        }
        return tableData;
    }
    public saveTableData(table: ooTables, data: any[][]) {
        this.tableDataCache[table] = data;
        const spreadsheet = this.getSpreadsheet(table);
        const sheetName = this.getSheetName(table);
        spreadsheet.getSheetByName(sheetName).getDataRange().setValues(data);
        SpreadsheetApp.flush();
    }


    private getFileName(table: ooTables): string {
        const tableFile = this.getMasterProperty(table + "_TableFile");
        const table_FileName = this.getMasterProperty(tableFile + "Name");
        return table_FileName + " - Version:" + this.version;
    }
    private getMasterId(table: ooTables): string {
        const tableFile = this.getMasterProperty(table + "_TableFile");
        const table_FileId = this.getMasterProperty(tableFile + "Id");
        console.log(table + " " + tableFile + " " + table_FileId);
        return table_FileId;
    }
    public getFolderNameWithVersion(folder: ooFolders) {
        return folder + " " + this.version;
    }
}



class ValuesCache {
    dataArray: any[][];
    dataHash = {};
    constructor(data: any[][]) {
        if (!data) throw new Error("no data for Values Cache");
        this.dataArray = data;
        for (let row of this.dataArray) {
            this.dataHash[row[0]] = row[1];
        }
    }
    public getValueByName(name: string) {
        return this.dataHash[name];
    }
}

export function getNextVersion():ooVersions {
    let oooNextVersion = (parseInt(currentOOversion, 10) + 1).toString();
    let nix = "";
    for (let nullen = 0; nullen < 4 - oooNextVersion.length; nullen++) {
        nix += "0";
    }
    oooNextVersion = nix + oooNextVersion;
    return oooNextVersion as ooVersions;
}

export function copyFolder(folderToCopyId:string,parentFolderId:string,oldVersion:ooVersions,newVersion:ooVersions){
    const parentFolder = DriveApp.getFolderById(parentFolderId);
    const folderToCopy = DriveApp.getFolderById(folderToCopyId);
    //create new Folder
    const folderCopy = parentFolder.createFolder(getNewName(folderToCopy.getName(),oldVersion,newVersion));
 
    //copy all files from the folder
    const fileIterator = folderToCopy.getFiles()
    while (fileIterator.hasNext()){
        const fileToCopy =fileIterator.next();
        fileToCopy.makeCopy(getNewName(fileToCopy.getName(),oldVersion,newVersion),folderCopy);
    }

    //copy all folders from the folder
    const folderIterator = folderToCopy.getFolders();
    while (folderIterator.hasNext()){
        const folderToCopy = folderIterator.next();
        copyFolder(folderToCopy.getId(),folderCopy.getId(),oldVersion,newVersion);
    }
}

function getNewName(oldName:string,oldVersion:ooVersions,newVersion:ooVersions):string{
    let folderToCopyName = oldName;
    //rename folder if it ends with version number
    if (oldVersion===folderToCopyName.substr(folderToCopyName.length - 4)){
        folderToCopyName=folderToCopyName.substr(0,folderToCopyName.length-4)+newVersion;
    }
    return folderToCopyName
}
