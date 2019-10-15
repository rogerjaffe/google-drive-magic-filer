/**
 * Google Drive Magic Filer V1.0.0 Roger Jaffe
 *
 * TL;DR -- If you just want to get started, see the quickstart guide below
 * 
 * I am a teacher and my students do their work on Google Drive and share their documents with me.
 * I needed a quicker way to organize their work in folders in My Drive instead of dragging and 
 * dropping the files one by one to their appropriate filing folder for later grading. I have several 
 * classes that have many assignments so I'm looking at a dozen or more filing folders at any given 
 * time. 
 *
 * Googla App Scripts to the rescue!!  Students share their files (doc, sheets, drawings, etc) and
 * they put a special code in the title.  I take the shared files from the "Shared with me" 
 * folder and drag them all to the special "drop folder."  The script reads the code and 
 * moves the file or folder into the folder associated with that code.  Finally the file or folder 
 * is removed from the drop folder. 
 *
 * For example, a student assignment could be titled "JaffeRoger-10ComputerSecurityRules-P025"  
 * The "P025" part is the special code that tells this script in which folder the file belongs.
 *
 * This works recursively for folders as well.  If the folder's name contains a code, the entire
 * folder is placed in the appropriate folder.  If it doesn't have a code, then the contents of
 * folder are searched for files or folders that have a code.
 * 
 * The spreadsheet URL shown in the SHEET_URL constant below points to the spreadsheet containing the 
 * special codes and the path to the folder in which the documents should be placed. Open the sheet
 * at the URL in the code to see a sample.  The DROP_FOLDER_URL points to the G-Drive folder that 
 * acts as the "drop folder".
 *
 * QUICKSTART GUIDE
 * SETUP (one time only)
 * 1. Open the spreadsheet with the URL listed below in SHEET_URL
 * 2. Make a copy of the sheet and erase the rows from row 3 to the bottom.  Keep 
 *    the top two rows to use as a template
 * 3. Copy the URL of the new sheet into the SHEET_URL variable between the quotes
 * 4. In your Google Drive, create a folder called 1DropBox (the '1' puts it at the top of the list)
 * 5. Copy the URL of the 1DropBox folder and paste it in the DROP_FOLDER_URL
 *    variable between the quotes.
 *
 * HOW TO USE
 * 1. In the SHEET_URL spreadsheet enter the following information for each folder or
 *    file you need to organize and move:
 *    
 *    Col A: Class (not required)
 *    Col B: Assignment name (not required)
 *    Col C: The code that students will include in the title of the folder or file
 *    Col D: The folder where the folder or file should be placed.  Use file path notation
 *           starting with a slash and the root folder
 *
 * 2. Make sure that students add the CODE to the title of the Google files and folders
 *    they share with you.
 * 3. When you're ready to organize student work, drag the student files / folders 
 *    from the Shared with me folder in your Google drive to the 1DropBox folder
 * 4. Select magicFiler in the drop down right under the "Resources" menu
 * 5. Click the Run arrow (underneath the "Publish" menu)
 *    Note: The first time you run the program you will be asked to give permission.
 *          Follow the prompts to give your permission so the program can access
 *          your Google document files and folders.
 *
 *    You can add more files / folders to the spreadsheet list at any time and
 *    you can run the program at any time.
 *
 * Let me know if you find this useful!  rogerjaffe@gmail.com
 *
 */

/**
 * Function to bootstrap the application
 */
function magicFiler() {
  const SHEET_URL = 'https://docs.google.com/spreadsheets/d/1EdWFjYptRKArAVp9KUmiF16-KjMsF25ueCPmXcSjziI/edit#gid=0';
  const DROP_FOLDER_URL = 'https://drive.google.com/drive/u/1/folders/1otcgbVua7NHQ56fw-q3bFTvbzh-lmQvU';

  var codes = parseCodeData(getCodeData(SHEET_URL));
  var srcRootUri = DROP_FOLDER_URL.split('/');
  var srcRoot = DriveApp.getFolderById(srcRootUri[7]);
  
  process(srcRoot, codes);
}

/**
 * Process the files and folders in the drop folder
 * 1. Process the files first
 * 2. Process the folders with codes second
 * 3. Go into the remaining folders and process recursively
 */
function process(srcFolder, codes) {
  // Deal with the files
  var files = srcFolder.getFiles();
  processItems(files, srcFolder, codes, false);
  // Then deal with the folders that are coded
  var folders = srcFolder.getFolders();
  processItems(folders, srcFolder, codes, true);
  // Finally, recurse into the folders that are left and 
  // process the same way.
  var innerFoldersIt = srcFolder.getFolders();
  while (innerFoldersIt.hasNext()) {
    var innerFolder = innerFoldersIt.next();
    process(innerFolder, codes);
  }
}

/**
 * For each item in the items iterator, look to see if it
 * has a code.  If so, then move it to the proper
 * destination folder
 */
function processItems(items, srcFolder, codes, isFolder) {
  while (items.hasNext()) {
    var item = items.next();
    var name = item.getName();
    for (var i=0; i<codes.length; i++) {
      if (name.toLowerCase().indexOf(codes[i].code.toLowerCase()) >= 0) {
        var destFolder = getDriveFolder(codes[i].dest);
        if (isFolder) {
          destFolder.addFolder(item);
          srcFolder.removeFolder(item);
        } else {
          destFolder.addFile(item);
          srcFolder.removeFile(item);
        }
      }
    }
  }
}  

/**
 * Retrieve the code data from the spreadsheet
 */
function getCodeData(sheetUrl) {
  var spreadsheet = SpreadsheetApp.openByUrl(sheetUrl);
  var codeSheet = spreadsheet.getSheetByName('Codes');
  var codeRange = codeSheet.getDataRange();
  var codeData = codeRange.getDisplayValues();
  return codeData;
}

/**
 * Parse the information from the spreadsheet.  Note
 * that only the code, and the destination path are
 * used here.  The rest of the information in the spreadsheet
 * is for convenience.
 */
function parseCodeData(data) {
  // Get the codes from the spreadsheet
  var codeObj = data.map(function(row, idx) {
    if (idx === 0) {
      return null
    } else {
      return {
        code: row[2],
        dest: row[3]
      }
    }
  })
  codeObj = codeObj.filter(function(row) {
    return row !== null
  });
  return codeObj;
}

/**
 * Get a reference to a folder when provided a UNIX-like path
 *
 * Thank you to Amit Agarwal for this function. You can find it at
 * https://ctrlq.org/code/19925-google-drive-folder-path
 */
function getDriveFolder(path) {
  var name, folder, search, fullpath;
  // Remove extra slashes and trim the path
  fullpath = path.replace(/^\/*|\/*$/g, '').replace(/^\s*|\s*$/g, '').split("/");
  // Always start with the main Drive folder
  folder = DriveApp.getRootFolder();
  for (var subfolder in fullpath) {
    name = fullpath[subfolder];
    search = folder.getFoldersByName(name);
    // If folder does not exit, create it in the current level
    folder = search.hasNext() ? search.next() : folder.createFolder(name);
  }
  return folder;
}
