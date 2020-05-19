/**
 * Retrieve the first folder named folderName, from the parentFolder
 * If it doesn't exist, create it.
 * @param {DriveApp.Folder} parentFolder 
 * @param {string} folderName 
 * @return {Folder}
 */
function getOrCreateFolderByName(parentFolder, folderName) {
  var iterator = parentFolder.getFoldersByName(folderName);
  if (iterator.hasNext()) {
    return iterator.next()
  } else {
    return parentFolder.createFolder(folderName)
  }
}


/**
 * Move Drive file to a destination folder and remove it from all other folders.
 * @param {file} file 
 * @param {folder} destinationFolder 
 */
function moveFile(file, destinationFolder) {
  // Get previous parent folders.
  var oldParents = file.getParents();
  // Add file to destination folder.
  destinationFolder.addFile(file);
  // Remove previous parents.
  while (oldParents.hasNext()) {
    var oldParent = oldParents.next();
    // In case the destination folder was already a parent, don't remove it.
    if (oldParent.getId() != destinationFolder.getId()) {
      oldParent.removeFile(file);
    }
  }
}


/**
 * Retrieve the first file named fileName, from the parentFolder
 * If it doesn't exist, return null.
 *
 * @param {Folder} parentFolder 
 * @param {string} fileName 
 * @return {File} The File, or null
 */
function getFileByName(parentFolder, fileName) {
  var iterator = parentFolder.getFoldersByName(fileName);
  if (iterator.hasNext()) {
    return iterator.next()
  } else {
    return null;
  }
}