//Script name: Save Document and Extract PDFs
//Author: William Dowling
//Creation Date: 10/15/15
/*
	Saves active document to user's desktop in dated folder with order number as filename.
	Saves each artboard as a PDF with artboard name as filename.
*/

//updated 27 May, 2016
//added a confirmation function to verify correct collars

//updated 15 November, 2016
//added cross platform support

//updated 23 December, 2016
//added condition to ignore any artboard named bag tag

#target Illustrator
function container()
{

	var valid = true;
	var scriptName = "save_export";

	function getUtilities ()
	{
		var utilNames = [ "Utilities_Container" ]; //array of util names
		var utilFiles = []; //array of util files
		//check for dev mode
		var devUtilitiesPreferenceFile = File( "~/Documents/script_preferences/dev_utilities.txt" );
		function readDevPref ( dp ) { dp.open( "r" ); var contents = dp.read() || ""; dp.close(); return contents; }
		if ( devUtilitiesPreferenceFile.exists && readDevPref( devUtilitiesPreferenceFile ).match( /true/i ) )
		{
			$.writeln( "///////\n////////\nUsing dev utilities\n///////\n////////" );
			var devUtilPath = "~/Desktop/automation/utilities/";
			utilFiles =[ devUtilPath + "Utilities_Container.js", devUtilPath + "Batch_Framework.js" ];
			return utilFiles;
		}

		var dataResourcePath = customizationPath + "Library/Scripts/Script_Resources/Data/";
		
		for(var u=0;u<utilNames.length;u++)
		{
			var utilFile = new File(dataResourcePath + utilNames[u] + ".jsxbin");
			if(utilFile.exists)
			{
				utilFiles.push(utilFile);	
			}
			
		}

		if(!utilFiles.length)
		{
			alert("Could not find utilities. Please ensure you're connected to the appropriate Customization drive.");
			return [];
		}

		
		return utilFiles;

	}
	var utilities = getUtilities();

	for ( var u = 0, len = utilities.length; u < len && valid; u++ )
	{
		eval( "#include \"" + utilities[ u ] + "\"" );
	}

	if ( !valid || !utilities.length) return;

	logDest.push(getLogDest());



	var docRef = app.activeDocument;
	var layers = docRef.layers;
	var aB = docRef.artboards;
	var date = getDate();
	log.l("date = " + date);

	var destString = "";
	var dest = findDest();
	log.l("dest = " + dest);

	//prompt user to confirm they have checked
	//for USA Collars

	function confirmCollars()
	{
		var bool = confirm("STOP!!!\nHave you checked the collars?\nIf you're not 100% positive the collars say 'Made in DR' hit cancel and check!");
		return bool;
	}

	var collarsChecked = confirmCollars();
	if (!collarsChecked)
		return;

	function findDest()
	{



		// if($.os.match('Windows')){
		// 	//PC
		// 	var user = $.getenv("USERNAME");
		// 	var dest = new Folder("C:\\Users\\" + user + "\\Desktop\\Today's Orders " + date);
		// 	return dest.fsName;

		// } else {
		// 	// MAC
		// 	var user = $.getenv("USER");
		// 	var dest = new Folder("/Volumes/Macintosh HD/Users/" + user + "/Desktop" + "/Today's Orders " + date);
		// 	// var dest = new Folder("~/Desktop/Today's Orders " + date);
		// 	return dest.fsName;
		// }
		destString = desktopPath + "Today's Orders " + date;
		return Folder(destString);
	}

	function getDate()
	{
		var today = new Date();
		var dd = today.getDate();
		var mm = today.getMonth() + 1; //January is 0!
		var yyyy = today.getYear();
		var yy = yyyy - 100;

		if (dd < 10)
		{
			dd = '0' + dd
		}
		if (mm < 10)
		{
			mm = '0' + mm
		}
		return mm + '.' + dd + '.' + yy;
	}



	// dest = "/Volumes/Macintosh HD" + dest + " " + date;
	if (!dest.exists)
	{
		log.l("dest does not exist. creating new dest folder at ")
		var newFolder = new Folder(destString);
		newFolder.create();
	}
	if (docRef.name.substring(0, 2) == "Un")
	{
		log.l("document is untitled. prompting user for order number.");
		var oN = prompt("Enter Order Number", "1234567");
		log.l("user entered " + oN);
		var fileName = oN;
	}
	else
	{
		var fileName = docRef.name;
	}

	log.l("setting fileName to " + fileName);

	var saveFile = new File(destString + "/" + fileName);
	log.l("saveFile.name = " + destString + "/" + fileName);

	if (saveFile.exists)
	{
		log.l("saveFile already exists. prompting for overwrite.");
		if (!confirm("This File already exists.. Do you want to overwrite?", false, "Overwrite file?"))
		{
			log.l("::::user chose not to overwrite. exiting script.");
			return;
		}
		log.l("user chose to overwrite file.");
	}
	docRef.saveAs(saveFile);
	if (fileName.indexOf(".ai") == -1)
	{
		var pdfdest = destString + "/" + fileName + "_PDFs";
	}
	else
	{
		var pdfdest = destString + "/" + fileName.substring(0, fileName.indexOf(".ai")) + "_PDFs";
	}
	var pdfFolder = new Folder(pdfdest);
	log.l("pdfFolder = " + pdfFolder);
	if (!pdfFolder.exists)
	{
		log.l("creating pdf folder");
		pdfFolder.create();
		log.l("after creating folder: pdfFolder.exists = " + pdfFolder.exists);
	}
	else if (pdfFolder.exists)
	{
		var removeFiles = pdfFolder.getFiles();
		for (var a = 0; a < removeFiles.length; a++)
		{
			removeFiles[a].remove();
		}
		pdfFolder.remove();
		pdfFolder = new Folder(pdfdest);
		pdfFolder.create();


	}

	//loop artboards to save individual PDFs

	var pdfSaveOpts = new PDFSaveOptions();
	pdfSaveOpts.preserveEditability = false;
	pdfSaveOpts.viewAfterSaving = false;
	pdfSaveOpts.compressArt = true;
	pdfSaveOpts.optimization = true;
	//     pdfSaveOpts.preset = "[Smallest File Size]";
	var seq = 1;

	for (var a = 0; a < aB.length; a++)
	{
		if (aB[a].name.toLowerCase().indexOf("bag tag") > -1)
		{
			continue;
		}
		var range = (a + 1).toString();
		pdfSaveOpts.artboardRange = range;
		var pdfFileName = pdfdest + "/" + aB[a].name;
		var ext = ".pdf";
		var thisPDF = new File(pdfFileName + ext);
		log.l("attempting to save: " + pdfFileName);
		if (thisPDF.exists)
		{
			log.l(aB[a].name + " is a duplicate of an existing pdf. Adding the sequence number: " + seq);
			thisPDF = new File(pdfFileName + " " + seq + ext);
			seq++;
		}
		docRef.saveAs(thisPDF, pdfSaveOpts);
	}

	printLog();
}
container();