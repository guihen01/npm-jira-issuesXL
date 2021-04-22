const fetch = require("node-fetch");
var prompt = require('prompt-sync')();
var fs = require('fs');
const Excel = require('exceljs')

// declaration des variables 
// --------------------------------------------------------------------------
var nbissues;                 // number of issues in the Jira project 
var KeyField = [];            // Issue key 
var IdField = [];             // Issue Id
var TypeField = [];           // Issue type
var StatusField = [];         // Issue status
var ResolutionStatus = [];    //Resolution 
var AssigneeField = [];       //Assignee username
var AssigneeDisplay = [];     //Assignee full name(displayName)
var ReporterField = [];       //Reporter username
var ReporterDisplay = [];     //Reporter full name
var CreatorField = [];        //Creator username
var CreatorDisplay = [];      //Creator full name(displayName)

// authentication data and questions 
//-----------------------------------------------------------------------------
server = ' ';
var server = prompt('enter url of server ( example : http://localhost:8080 ');
//console.log(server);

projectname = ' ';
var projectname = prompt('enter project name: ');

const url = server + "/rest/api/2/search?jql=project=" + projectname;
console.log('-------------------------------------------------------------------')
console.log('url choosed to pick list of Jira issues in project :' + url);
console.log('-------------------------------------------------------------------')

username = ' ';
var username = prompt('enter username: ');
//console.log(username);

var password = ' ';
var password = prompt('enter password : ');
//console.log(password);

//------------------------------------------------------------------------------------
// main function
//------------------------------------------------------------------------------------
async function getData(url) {
	try {
		var auth = 'Basic ' + Buffer.from(username + ':' + password).toString('base64');
		var headers = { 'Authorization': auth };
		const response = await fetch(url, { method: 'GET', headers: headers }).then((res) => { return res.json() })
			.then((json) => {
				//console.log(json);
				data = JSON.stringify(json, null, 2);
				fs.writeFileSync('List-Issues.json', data);
				extract();
				WriteXL();
			});
	}
	catch (error) {
		console.log(error);
	}
};

//----------------------------------------------------------------------------------
// execution of the function
//----------------------------------------------------------------------------------
getData(url);



//---------------------------------------------------------------------------------
// Data Fields extraction from json buffer :
//---------------------------------------------------------------------------------
function extract() {
	var data = fs.readFileSync("List-Issues.json", "UTF-8");
	//console.log(data);  

	var Obj = JSON.parse(data);

	// get number of issues 
	//--------------------------------------------
	nbissues = Obj.total;
	console.log('-------------------------------');
	console.log("Number of Issues: " + Obj.total);


	key1 = Obj.issues[0].key;
	//console.log("key 1 : " + key1);

	// get all issues key 
	//----------------------------------------
	//var KeyField = [];
	for (let i = 0; i < nbissues; i++) {
		KeyField[i] = Obj.issues[i].key;
	}

	//Test affichage des issues key
	//---------------------------------------
	//for  (let i=0; i< nbissues; i++)
	//{ console.log("key" + i + " : " + KeyField[i]); }

	// get all issues id 
	//-----------------------------------------
	//var IdField = [];
	for (let i = 0; i < nbissues; i++) {
		IdField[i] = Obj.issues[i].id;
	}

	//Test affichage des Id key
	//---------------------------------------
	//for  (let i=0; i< nbissues; i++)
	//{ console.log("id" + i + " : " + IdField[i]); }	

	// get all issues type 
	//-----------------------------------------
	//var TypeField = [];
	for (let i = 0; i < nbissues; i++) {
		TypeField[i] = Obj.issues[i].fields.issuetype.name;
	}

	//Test affichage des issues type
	//---------------------------------------
	//for  (let i=0; i< nbissues; i++)
	//{ console.log("issuetype" + i + " : " + TypeField[i]); }	

	// get all issues creator username (The username creator of the issue)
	//-----------------------------------------------------------	
	//var CreatorField = [];
	for (let i = 0; i < nbissues; i++) {
		CreatorField[i] = Obj.issues[i].fields.creator.name;
	}

	//Test affichage des creators
	//---------------------------------------
	//for  (let i=0; i< nbissues; i++)
	//{ console.log("issue creator" + i + " : " + CreatorField[i]); }

	// get all issues creator Displayname ( The display full name of the creator of the issue)
	//--------------------------------------------------------------------------------
	//var CreatorDisplay = [];
	for (let i = 0; i < nbissues; i++) {
		CreatorDisplay[i] = Obj.issues[i].fields.creator.displayName;
	}

	//Test affichage des creators displayName
	//---------------------------------------
	// for  (let i=0; i< nbissues; i++)
	//{ console.log("issue creator Displayname" + i + " : " + CreatorDisplay[i]); }

	// get all issues Reporters ( The reporter username of the issues)
	//--------------------------------------------------------------------------------
	//var ReporterField = [];
	for (let i = 0; i < nbissues; i++) {
		ReporterField[i] = Obj.issues[i].fields.reporter.name;
	}

	//Test affichage des issues reporters
	//---------------------------------------
	//for  (let i=0; i< nbissues; i++)
	//{ console.log("issues reporters " + i + " : " + ReporterField[i]); }

	// get all issues Reporters displayname (The reporter displayname of the issue)
	//--------------------------------------------------------------------------------
	//var ReporterDisplay = [];
	for (let i = 0; i < nbissues; i++) {
		ReporterDisplay[i] = Obj.issues[i].fields.reporter.displayName;
	}

	//Test affichage des issues reporters displayName
	//-----------------------------------------------
	// for  (let i=0; i< nbissues; i++)
	//{ console.log("issues reporters displayName " + i + " : " + ReporterDisplay[i]); }

	// get all issues assignee username (The assignee username of the issue)
	//--------------------------------------------------------------------------------
	//var AssigneeField = [];
	for (let i = 0; i < nbissues; i++) {
		if (Obj.issues[i].fields.assignee != null) {
			AssigneeField[i] = Obj.issues[i].fields.assignee.name;
			if (AssigneeField[i] == null) {
				AssigneeField[i] = "UnAssigned";
			}
		}
		else {
			AssigneeField[i] = "UnAssigned";
		}
	}

	//Test affichage des assignee username
	//-----------------------------------------------
	// for  (let i=0; i< nbissues; i++)
	//{ console.log("issue assignee username " + i + " : " + AssigneeField[i]); }

	// get all issues assignee displayname (The assignee displayname of the issue)
	//--------------------------------------------------------------------------------
	//var AssigneeDisplay = [];
	for (let i = 0; i < nbissues; i++) {
		if (Obj.issues[i].fields.assignee != null) {
			AssigneeDisplay[i] = Obj.issues[i].fields.assignee.displayName;
			if (AssigneeDisplay[i] == null) {
				AssigneeDisplay[i] = "UnAssigned";
			}
		}
		else {
			AssigneeDisplay[i] = "UnAssigned";
		}
	}

	//Test affichage des assignee displayName
	//-----------------------------------------------
	//for  (let i=0; i< nbissues; i++)
	//{console.log("issue assignee displayName " + i + " : " + AssigneeDisplay[i]); }

	// get all issues status resolution 
	//-----------------------------------------
	//var ResolutionStatus = [];
	for (let i = 0; i < nbissues; i++) {
		ResolutionStatus[i] = Obj.issues[i].fields.resolution;
		if (ResolutionStatus[i] == null) {
			ResolutionStatus[i] = "Unresolved";
		}
		else {
			ResolutionStatus[i] = "Resolved";
		}
	}

	//Test affichage des status resolution
	//-----------------------------------------------
	//for  (let i=0; i< nbissues; i++)
	//{ console.log("issue resolution status " + i + " : " + ResolutionStatus[i]); }

	// get all issues status
	//-----------------------------------------
	//var StatusField = [];
	for (let i = 0; i < nbissues; i++) {
		StatusField[i] = Obj.issues[i].fields.status.name;
		if (StatusField[i] == null) {
			StatusField[i] = "no status";
		}
	}

	//Test affichage des status 
	//-----------------------------------------------
	//for  (let i=0; i< nbissues; i++)
	//{ console.log("issue status " + i + " : " + StatusField[i]); }

}

//------------------------------------------------------------------------
// function for writing to Excel file
//-------------------------------------------------------------------------
async function WriteXL() {
	let workbook = new Excel.Workbook()

	let worksheet = workbook.addWorksheet('List-Issues')

	// Add a first row by sparse Array (assign to columns A, to & )
	//----------------------------------------------------------------
	const rowValues = [];
	rowValues[1] = "Project name";
	rowValues[2] = "Issue key";
	rowValues[3] = "Issue Id";
	rowValues[4] = "Issue type";
	rowValues[5] = "Issue status";
	rowValues[6] = "Resolution";
	rowValues[7] = "Assignee username";
	rowValues[8] = "Assignee full name";
	rowValues[9] = "Reporter username";
	rowValues[10] = "Reporter full name";
	rowValues[11] = "Creator username";
	rowValues[12] = "Creator full name";
	worksheet.addRow(rowValues);

	//la premiere ligne est en bold :
	worksheet.getRow(1).font = { bold: true }

	// write nbissues row in the Excel file
	//--------------------------------------------------------------
	for (let i = 0; i < nbissues; i++) {
		rowValues[1] = projectname;
		rowValues[2] = KeyField[i];    // Issue Key
		rowValues[3] = IdField[i];     // Issue Id
		rowValues[4] = TypeField[i];    // Issue type
		rowValues[5] = StatusField[i];  // Issue status
		rowValues[6] = ResolutionStatus[i];   //Resolution 
		rowValues[7] = AssigneeField[i];      //Assignee username
		rowValues[8] = AssigneeDisplay[i];    //Assignee full name(displayName)
		rowValues[9] = ReporterField[i];      //Reporter username
		rowValues[10] = ReporterDisplay[i];   //Reporter full name
		rowValues[11] = CreatorField[i];      //Creator username
		rowValues[12] = CreatorDisplay[i];     //Creator full name(displayName)
		worksheet.addRow(rowValues);
	}

	await workbook.xlsx.writeFile('List-Issues.xlsx')
	console.log('----------------------------------');
	console.log("File : List-Issues.xlsx is written");
	console.log('----------------------------------');
}




