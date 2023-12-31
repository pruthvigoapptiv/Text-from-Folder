const fs = require("fs");
const pdfparse = require("pdf-parse");
const xlsx = require("xlsx");
var arr = [];
var pageNo = 0;
// Function to extract text from a PDF file
function jsonToExcel(jsonData, excelFileName, fileName) {
  // Create a worksheet
  const ws = xlsx.utils.aoa_to_sheet(jsonData);

  // Create a workbook and add the worksheet to it
  const wb = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(wb, ws, "Sheet 1");

  // Write the workbook to an Excel file
  xlsx.writeFile(wb, excelFileName);

  console.log(`Writing data for file '${fileName}'`);
}
function extractTextFromPDF(pdfPath) {
  pdfparse(pdfPath).then(function (data) {
    var i = data.numpages;
    var len = pdfPath.length;
    var pdf = pdfPath;
    var temp12 = "";
    for (let j = 9; j < len; j++) {
      temp12 += pdf[j];
      if (pdf[j] == ".") {
        break;
      }
    }
temp12+="pdf"

    var title, docNo, Dated, Remarks, GrandTotal;
    arr.push([`  `]);
    arr.push(["FileName","title", "docNo", "Dated", "Remarks", "GrandTotal", "PageNo"]);
    // console.log(data.text);

    for (let i = 0; i < data.text.length - 20; i++) {
      // conditon for Remarks using the logic that is it followed by email and ends at char ']'
      if (data.text.substr(i, 7) == "Message") {
        i = i + 9;
        Remarks = data.text.substr(i, 15);
        // console.log(Remarks);
      }

      if (data.text[i] == "]") {
        var k = i;
        for (k = i; k >= 10; k--) {
          if((data.text[k]=='m' && data.text[k-3]=='.')|| (data.text[k]=='M' && data.text[k-3]=='.')|| (data.text[k]=='N' && data.text[k-2]=='.') || (data.text[k]=='n' && data.text[k-2]=='.'))
           {
            break;
          }
        }
        k = k + 2;
        var temp = "";
        for (k = k; k < k + 60; k++) {
          temp += data.text[k];
          if (data.text[k] == "]") {
            break;
          }
        }
        title = temp;
        //  console.log(title);
      }
      //For Dated just check whether the string Dated is present or not
      if (data.text.substr(i, 5) == "Dated") {
        i = i + 7;
        Dated = data.text.substr(i, 18);
        // console.log(Dated);
      }

      // Similar logic is for Doc No.
      if (data.text.substr(i, 7) == "Doc No.") {
        i = i + 7;
        docNo = data.text.substr(i + 1, 14);

        // Syntax for exporting everything to excel

        //   console.log(title);
        // console.log(docNo);
        //   console.log(Remarks);
        //   console.log(GrandTotal);
      }
      // For Grand Total it is after the string Amount in words
      if (data.text.substr(i, 15) == "Amount in words") {
        var temp = "";
        i = i + 18;
        while (data.text[i] != ".") {
          temp += data.text[i];
          i++;
        }
        temp += data.text[i + 1];
        temp += data.text[i + 2];
        //store decimals

        GrandTotal = temp;
        pageNo++;
        arr.push([temp12,title, docNo, Dated, Remarks, GrandTotal, pageNo]);
        //  console.log(arr);
      }
    }
    arr.push([`  `]);

    pageNo = 0;
    const excelFileName = `output.xlsx`;
    jsonToExcel(arr, excelFileName, temp12);
  });
}
// Function to iterate over each PDF file in a folder
async function processPDFsInFolder(folderPath) {
  // Get a list of files in the folder
  const files = fs.readdirSync(folderPath);

  // Iterate over each file
  for (const file of files) {
    // Check if the file has a .pdf extension
    if (file.toLowerCase().endsWith(".pdf")) {
      const pdfPath = `${folderPath}/${file}`;

      // Await the result of extractTextFromPDF
      const values = extractTextFromPDF(pdfPath);
      // Add the values to the worksheet
    }
  }
}

// Replace 'path/to/your/pdf/folder' with the path to your folder containing PDFs
const pdfFolder = "./Folder";
processPDFsInFolder(pdfFolder);
