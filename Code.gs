

// Ici les ID des fichiers et des dossiers utilisés


var ID_DOSSIER_ARCHIVES = '18JrjJfbPOehB0dERIuwtxun9mGTJ9puy',
    ID_FICHIER_BUDGET = '1bXTUZoMkuji9jzpGzvvRW9PWcut7NittbLM8zaUHNUY',
    ID_FORM = '10Sqxy6c4QwGkIm9fkGC4S7On6VvAFL6gBoJR-guCCNs';






//ici les mois et le nom générique du fichier. Utile pour modifier la langue.

var ARRAY_MOIS = ['Janvier','Février','Mars','Avril','Mai','Juin','Juillet','Août','Septembre','Octobre','Novembre','Décembre'],
    CODE_FICHIER = 'BudJeff';




// La fonction à laquelle est attachée un trigger pour être lancée tous les mois et lancer les différentes étapes à réaliser
function nouveauMois(){
  
  
  var date = new Date();
  if (date.getMonth() == 0){
    var nouveauMois = ARRAY_MOIS[0],
        ancienMois = ARRAY_MOIS[11],
        nouvelleAnnee = date.getFullYear().toString(),
        ancienneAnnee = (date.getFullYear() -1).toString()
        ;
    
  }
  
  else {
    var nouveauMois = ARRAY_MOIS[date.getMonth()],
        ancienMois = ARRAY_MOIS[date.getMonth() -1],
        nouvelleAnnee = date.getFullYear().toString(),
        ancienneAnnee = date.getFullYear().toString()
        ;
    
    
  
  }
  
  copierFile(ancienMois, ancienneAnnee);
  renameFichiersOrigine(nouveauMois, nouvelleAnnee);
  viderLesFormulairesLies(ID_FICHIER_BUDGET);





}




// Fonction pour renomer les fichiers d'origine pour les réutiliser le mois suivant
function renameFichiersOrigine(mois, annee){
  
var nomFichierMois = CODE_FICHIER+' '+mois+" "+annee;

DriveApp.getFileById(ID_FICHIER_BUDGET).setName(nomFichierMois);
DriveApp.getFileById(ID_FORM).setName(nomFichierMois);



}


// Fonction pour créer les répertoires d'archivage s'ils n'existent pas encore

function creerDossierMois(mois,annee){
  var nomFolderAnnee = CODE_FICHIER+' '+annee;
  var nomFolderMois = CODE_FICHIER+' '+mois+" "+annee;

  var iteratorFolderAnnee = DriveApp.getFolderById(ID_DOSSIER_ARCHIVES).getFoldersByName(nomFolderAnnee);  
  if (iteratorFolderAnnee.hasNext()){

 
    var folderAnnee=iteratorFolderAnnee.next();
    var iteratorFolderMois = folderAnnee.getFoldersByName(nomFolderMois);  
    if (iteratorFolderMois.hasNext()){
    
      //dans ce cas là on a déjà un dossier au nom du mois...
      return iteratorFolderMois.next();
    
    
    }
    else{
      // dans le cas contraire on crée le dossier
      return folderAnnee.createFolder(nomFolderMois);
    
    }
  
  }

  
  else{
    
    var folderMois = DriveApp.getFolderById(ID_DOSSIER_ARCHIVES).createFolder(nomFolderAnnee).createFolder(nomFolderMois);  
  
  }

}



function copierFile(mois,annee){
  
  
  var nomFileMois =  CODE_FICHIER+' '+mois+" "+annee;
  var dossierMois = creerDossierMois(mois,annee);
  var idFichierCopie = DriveApp.getFileById(ID_FICHIER_BUDGET).makeCopy(nomFileMois,dossierMois).getId();
  var sheetsFichierCopie = SpreadsheetApp.openById(idFichierCopie).getSheets();
  
  for (i=0; i < sheetsFichierCopie.length ; i++){
    var lienFormCopie = sheetsFichierCopie[i].getFormUrl();
    if (lienFormCopie != null){
      var formCopie = DriveApp.getFileById(FormApp.openByUrl(lienFormCopie).getId()).setName(nomFileMois);
      var parentsFormCopieIterator = formCopie.getParents();
      
      
      while (parentsFormCopieIterator.hasNext()){
        var parentFormCopie = parentsFormCopieIterator.next();
        parentFormCopie.removeFile(formCopie);
        dossierMois.addFile(formCopie);
        
        
           
           
           }
    
    
    }
    
  }


}





function viderLesFormulairesLies(idSpreadsheet){
  var listeDesSheets = SpreadsheetApp.openById(idSpreadsheet).getSheets();
  for (k = 0; k < listeDesSheets.length; k ++){
  
    var lienVersForm = listeDesSheets[k].getFormUrl();
    if (lienVersForm != null){
    
      FormApp.openByUrl(lienVersForm).deleteAllResponses();
      listeDesSheets[k].clear();
    
      
    }
    
  
  }
  

}
