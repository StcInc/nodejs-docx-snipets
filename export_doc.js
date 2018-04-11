const docx = require('docx');
var fs = require('fs');
var parse = require('csv-parse');


function parseArgs(args, schema, strict) {
    var res = {};
    for (var i = 2; i < args.length; ++i) {
        // console.log(args[i]);
        var parts = args[i].split("=");
        paramName = parts[0].replace("--", "");
        // console.log(paramName);
        if (schema.includes(paramName)) {
            res[paramName] = parts[1];
        }
    }
    if (strict) {
        for (var i = 0; i < schema.length; ++i) {
            if (typeof res[schema[i]] == 'undefined') {
                return null;
            }
        }
    }
    return res;
}

function readCSV (inputFile, includesHeader, callback) {
    var csvData=[];
    fs.createReadStream(inputFile)
    .pipe(parse({delimiter: ','}))
    .on('data', function(csvrow) {
        csvData.push(csvrow);
    })
    .on('end',function() {
        callback(csvData);
    });
}

function addTable(doc, tableData) {
    if (tableData.length > 0) {
        const table = doc.createTable(tableData.length, tableData[0].length);
        for (var i = 0; i < tableData.length; ++i) {
            for (var j = 0; j < Math.min(tableData[i].length,tableData[0].length); ++j) {
                table.getCell(i, j).addContent(new docx.Paragraph(tableData[i][j]));
            }
        }
    }
}

function createDoc(args) {
    var doc = new docx.Document();

    doc.addParagraph(new docx.Paragraph("Parameters"));
    addTable(doc, args.paramData);

    doc.addParagraph(new docx.Paragraph("Values"));
    addTable(doc, args.csvData);

    doc.addParagraph(new docx.Paragraph("Graphics"));
    const image = doc.createImage(args.image);

    console.log("Outputting");
    var exporter = new docx.LocalPacker(doc);
    exporter.pack(args.docName);
}

function main() {
    var args = parseArgs(process.argv,
        [
            "table-data",
            "image",
            "output",
            "params"
        ], true);
    if (args) {
        readCSV(args["params"], true, function (paramData){
            readCSV(args["table-data"], true, function (csvData) {
                createDoc({
                    paramData: paramData,
                    csvData: csvData,
                    image: args.image,
                    docName: args.output
                });
            });
        });
    }
    else {
        console.error("Not all args supplied!");
    }
}


main();
