import * as publicUiFunctions from './ui'


interface Imenu {
    onOpen: ()=> void;
    newVersion: ()=> void;
}
declare let global:Imenu;
// Expose public functions by attaching to `global`
global.onOpen = publicUiFunctions.onOpen;
global.newVersion = publicUiFunctions.newVersion;
