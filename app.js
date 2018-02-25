const Excel = require('exceljs');
const program = require("commander");
const fs = require('fs');

const resultDir = "./script_result/";
const notSupportDir = "./not_support_api/";
const excelDir = "./excel_files/";

program.on('--help', function () {
let msg = `
  Examples:

    node app.js -p example -v 1.0.2.28 -s 3.18

  Folders:

    -- ./script_result/2.0_result 3.0_result\t[should contain "2.0" and "3.0"]
    -- ./not_support_api/not_support_list\t[projectName-specVersion.txt]
    -- ./excel_files/excel_file\t\t\t[excel file created by this tool]`
	console.log(msg);
})

program.option('-p, --project <project>', 'project name. eg: example')
	.option('-v, --version <version>', 'project version. eg: 1.0.2.28')
	.option('-s, --spec <spec>', 'meet SOAP Spec version. eg: 3.18')
	.parse(process.argv);

if (!program.project || !program.version || !program.spec) {
	program.outputHelp();
	process.exit(1);
} else {
	program.version = (program.version.slice(0, 1).toLowerCase() == "v" ? "" : "V") + program.version;
	program.spec = (program.spec.slice(0, 1).toLowerCase() == "v" ? "" : "V") + program.spec;
}

let files = fs.readdirSync(resultDir);
if (files.length != 2) {
	console.log("Just put the results of 2.0 and 3.0 to ./script_result folder.\n");
	process.exit(1);
}

let notSupportFiles = fs.readdirSync(notSupportDir);
let notSupportFile;
let notSupportData = "";
let pattern = new RegExp(program.project + "-v?" + program.spec.replace(/[^\.\d]/g, "").replace(".", "\\.") + "\\.txt", "gi");
for (let j in notSupportFiles) {
	pattern.compile(pattern);
	if (pattern.test(notSupportFiles[j])) {
		notSupportFile = notSupportFiles[j];
		break;
	}
}
if (!notSupportFile) {
	console.log("Can't find file for not support api. You may miss it.\n");
} else {
	try {
		notSupportData = fs.readFileSync(notSupportDir + notSupportFile, "utf-8");
		if (notSupportData.trim() == "")
			throw new Error();
		notSupportData.replace(/\r\n/g, "\n");
		var eachNotSupport = notSupportData.split(/\n/g);
		eachNotSupport = eachNotSupport.filter(item => !(item.trim() === ""))
	} catch (err) {
		notSupportData = "";
		console.log(err);
	}
}

let excel = new Excel.Workbook();
let worksheet2 = excel.addWorksheet("SOAP 2.0");
let worksheet3 = excel.addWorksheet("SOAP 3.0");

for (let i in files) {
	if (files[i].indexOf("2.0") != -1)
		var sheet = worksheet2;
	else if (files[i].indexOf("3.0") != -1)
		var sheet = worksheet3;
	else {
		console.log("You may put the wrong files in ./script_result folder.\n");
		process.exit(1);
	}
	sheet.getCell("A1").value = "Firmware Version: " + program.version;
	sheet.getCell("A2").value = "Meet SOAP Spec: " + program.spec;
	sheet.getCell("A3").value = "Test Results: (PASS/FAIL N/A:Not support)";
	sheet.getCell("A4").value = "Test Items";
	sheet.getCell("B4").value = "Results";
	sheet.getCell("C4").value = "Comments";
	sheet.mergeCells("A1:C1");
	sheet.mergeCells("A2:C2");
	sheet.mergeCells("A3:C3");
	sheet.getColumn(1).width = 80;
	sheet.getColumn(2).width = 25;
	sheet.getColumn(3).width = 50;

	try {
		var data = fs.readFileSync(resultDir + files[i], "utf-8");
	} catch (err) {
		console.log(err);
		process.exit(1);
	}

	line = data.split("\n");
	for (let i in line) {
		line[i] = line[i].replace(/\[[^\[\]]*\]/g, "");
		line[i] = line[i].replace(/ResponseTime=.*$/g, "");
		line[i] = line[i].replace(/\s:\s/g, "");
		line[i] = line[i].trim();
		if (line[i] == "")
			continue;
		var each = line[i].split(/\s+/);
		each[0] = (each[0] == undefined ? "" : each[0]);
		each[1] = (each[0] == undefined ? "" : each[1]);
		sheet.addRow([each[0], each[1]]);
	}

	sheet.eachRow(function (Row, rowNum) {
		Row.eachCell(function (Cell, cellNum) {
			if (rowNum < 5) {
				Cell.alignment = {
					vertical: 'middle',
					horizontal: 'center',
					wrapText: true
				};
				Cell.font = {
					size: 14,
					bold: true
				};
			} else if (cellNum > 1) {
				if (Row.getCell(2) == "FAIL" && notSupportData != "") {
					var isNotSupport = false;
					for (let k in eachNotSupport) {
						let action = eachNotSupport[k].replace(/^([^\s\[]*)[\s|\[].*$/gm, "$1").trim();
						let reason = eachNotSupport[k].match(/\[([^\[\]]*)\]/)[0].replace("]", "").replace("[", "");
						if (action == Row.getCell(1).value) {
							Row.getCell(2).value = "N/A";
							Row.getCell(3).value = reason;
							Row.getCell(1).font = {
								color: {
									argb: 'FF1E90FF'
								},
								italic: true
							};
							Row.getCell(2).font = {
								color: {
									argb: 'FF1E90FF'
								},
								italic: true
							};
							Row.getCell(3).font = {
								color: {
									argb: 'FF1E90FF'
								},
								italic: true
							};
							isNotSupport = true;
						}
					}
					if (!isNotSupport) {
						Row.getCell(1).font = {
							color: {
								argb: 'FFFF4500'
							},
							italic: true
						};
						Row.getCell(2).font = {
							color: {
								argb: 'FFFF4500'
							},
							italic: true
						};
					}
				}
				Cell.alignment = {
					vertical: 'middle',
					horizontal: 'center'
				};
			}
		})
	})
}

try {
	fs.accessSync(excelDir);
}
catch(err) {
	fs.mkdirSync(excelDir);
}

var fileName = excelDir + program.project + "-" + program.version + "-" + "SOAP-" + program.spec + "-Result.xlsx";
excel.xlsx.writeFile(fileName).then(function () {
	console.log("Done to generate excel file: " + fileName);
}).catch(function () {
	console.log("Failed to generate excel file.");
})