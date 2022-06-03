function myFunction() {
  // get access to the desired drive folder
  let id = 'drive-folder-id';
  let mainFolder = DriveApp.getFolderById(id);
  console.log(mainFolder.getName());

  //create an array containing the names of the folders already in this folder
  let folderIter = mainFolder.getFolders();
  let exisFolderList = []
  while(folderIter.hasNext()) {
    let name = folderIter.next().getName()
    // console.log(name);
    exisFolderList.push(name)
  }
  console.log(exisFolderList)

  //function takes in two date objects and returns the corresponding folder name as a string
  function formatNameFromDates(startDate, endDate) {
    return startDate.toDateString().slice(4) + ' to ' + endDate.toDateString().slice(4);
  }

  //manually enter the important dates
  let startDate = new Date(2021, 7, 9); //the first day of the scheel week (Mon)
  let endDate = new Date(2021, 7, 13); //the last day of the school week (Fri)
  let finalDate = new Date(2022, 5, 3); //the final Friday of the school year

  //create folder names for all the weeks until the end of the school year and add them to an array
  let newNameList = [];
  let count = 1;
  while (finalDate >= endDate) {
    newNameList.push(`(${count})-${formatNameFromDates(startDate, endDate)}`);
    startDate.setDate(startDate.getDate() + 7);
    endDate.setDate(endDate.getDate() + 7);
    count++;
  }
  console.log(newNameList);

  //if the new folder name is not already in the folder, add a folder with such a name
  for (let i = 0; i < newNameList.length; i++) {
    let name = newNameList[i];
    if (!exisFolderList.includes(name)) {
      mainFolder.createFolder(name);
    }
  }
  
}