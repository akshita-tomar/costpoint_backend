const express = require('express');
const cors = require('cors');
const fileUpload = require('express-fileupload');
const app = express();
const XLSX = require("xlsx");
const ExcelJS = require("exceljs");
const fs = require('fs');
const stream = require('stream');
const workbookStream = new stream.PassThrough();


app.use(cors());
app.use(fileUpload());
app.use(express.json({ limit: '200mb' }));
app.use(express.urlencoded({ limit: '200mb', extended: true }));
app.get('/', (req, res) => {
  res.send("working");
});

app.get('/test',(req,res)=>{
 try {
  return res.json("costpoint backend working")
 }catch(error){
  return res.status(500).json({message:"Internal server error",data:error.message})
 }})


//api to get data from the file and send back 
app.post('/get_file1data', (req, res) => {
  if (!req.files || !req.files.excelFile) {
    return res.status(400).send('No file uploaded.');
  }
  const excelFile = req.files.excelFile;
  excelFile.mv('uploads/' + excelFile.name, (err) => {
    if (err) {
      console.error(err);
      return res.status(500).send('File upload failed.');
    }
    try {
      const fileData = excelFile.data;
      const workbook = XLSX.read(fileData, { type: 'buffer' });
      const sheetNames = workbook.SheetNames; // Array of sheet names
      const sheetData = [];

      for (const sheetName of sheetNames) {
        const worksheet = workbook.Sheets[sheetName];
        const jsondata = XLSX.utils.sheet_to_json(worksheet, {
          header: 1,
          defval: '',
          raw: false,
          dateNF: 'mm/dd/yyyy',
        });
        const rows = jsondata.filter((row) => row.join('').trim() !== '');
        const processedRows = rows.map((row) =>
          row.map((value) => {
            // Check if the value is a numeric string
            if (/^-?\d*\.?\d+$/.test(value)) {
              return Number(value);
            }
            return value;
          })
        );
        sheetData.push({ sheetNames, rows: processedRows });
      }
      //delete the uploaded file 
      const filePath = 'uploads/' + excelFile.name;
      fs.unlink(filePath, (err) => {
        if (err) {
          // console.error('Error deleting the file:', err);
        }
         else {
          // console.log('File deleted successfully.');
        }
      });
      return res.json({ sheetData, sheetNames });
    } catch (error) {
      console.error(error)
      return res.status(500).send('Error processing the file.');
    }
  });
});



//api to get data and perfrom comparision
app.post("/compare_data", async (req, res) => {
  let exact_matchedfile1 = [];
  let exact_matchedfile2 = [];
  let non_exactmatchfile1 = [];
  let non_excatmatchedfile2 = [];
  let indexfile1 = [];
  let indexfile2 = [];
  let indexfile1matchwith=[];
  let indexfile2matchwith=[];


  const columndata = req.body.columndata;
  const columndatafile2 = req.body.columndatafile2;
  const matchtypes = req.body.matchtypes;
  const radioInput = req.body.radioInput;
  const expectedvaluefile1 = req.body.expectedvaluefile1;
  const expectedvaluefile2 = req.body.expectedvaluefile2;
  

  function compareArraysAndStoreMatches(columndata, columndatafile2) {
  
    for (let i = 0; i < columndata.length; i++) {
      const record = columndata[i];
      const indexInFile2 = columndatafile2.indexOf(record);
      
      if (indexInFile2 !== -1 && record !== null) {
        exact_matchedfile1.push(record);
        indexfile1.push(i); // Add the index from columndata to indexfile1
        if(i===0){
          indexfile1matchwith.push("*"+columndatafile2[0]+"("+matchtypes+")")
        }else{
          indexfile1matchwith.push(indexInFile2+1);
        }
         // Add the index from columndatafile2 to indexfile2
      } else {
        non_exactmatchfile1.push(i)
        if(i===0){
          indexfile1matchwith.push("*"+columndatafile2[0]+"("+matchtypes+")");
        }else{
          indexfile1matchwith.push("no match");
        }
      }}
    

    for (let i = 0; i < columndatafile2.length; i++) {
      const record = columndatafile2[i];
      const indexInFile2 = columndata.indexOf(record);
    
      if (indexInFile2 !== -1 && record !== null) {
        exact_matchedfile2.push(record);
        indexfile2.push(i); // Add the index from columndata to indexfile1
        if(i===0){
          indexfile2matchwith.push("*"+columndata[0]+"("+matchtypes+")");
        }else{
          indexfile2matchwith.push(indexInFile2+1); // Add the index from columndatafile2 to indexfile2
        }
       
      } else {
        non_excatmatchedfile2.push(i)
        if(i===0){
          indexfile2matchwith.push("*"+columndata[0]+"("+matchtypes+")");
        }else{
          indexfile2matchwith.push("no match");
        }
      }
    }
    return {
      exact_matchedfile1,
      exact_matchedfile2, 
      non_exactmatchfile1, 
      non_excatmatchedfile2,
      indexfile1,
      indexfile2,
      indexfile1matchwith,
      indexfile2matchwith
    };
    
  }
  function compareArraysAndStoreMatchesnumeric(columndata, columndatafile2) {

    for (let i = 0; i < columndata.length; i++) {
      let record = columndata[i];
  
      if (typeof record === "number" && !isNaN(record) && columndatafile2.includes(record) && record !== null) {
          exact_matchedfile1.push(record);
          indexfile1.push(i);
  
          const indexInFile2 = columndatafile2.indexOf(record);
          if(i===0){
            indexfile1matchwith.push("*"+columndatafile2[0]+"("+matchtypes+")");
          }else{
            indexfile1matchwith.push(indexInFile2+1);
          }  
      } else {
        non_exactmatchfile1.push(i)
        if(i===0){
          indexfile1matchwith.push("*"+columndatafile2[0]+"("+matchtypes+")");
        }else{
          indexfile1matchwith.push("no match");
        }
      }
  }
  

  for (let i = 0; i < columndatafile2.length; i++) {
    let record = columndatafile2[i];

    if (typeof record === "number" && !isNaN(record) && columndata.includes(record) && record !== null) {
        exact_matchedfile2.push(record);
        indexfile2.push(i);

        const indexInFile2 = columndata.indexOf(record);
        if(i===0){
          indexfile2matchwith.push("*"+columndata[0]+"("+matchtypes+")");
        }else{
          indexfile2matchwith.push(indexInFile2+1);
        }
       
    } else {
      non_excatmatchedfile2.push(i)
     if(i===0){
      indexfile2matchwith.push("*"+columndata[0]+"("+matchtypes+")");
     }else{
      indexfile2matchwith.push("no match");
     }
     
    }
}

    return {
      exact_matchedfile1,
      exact_matchedfile2,
      non_exactmatchfile1,
      non_excatmatchedfile2,
      indexfile1,
      indexfile2,
      indexfile1matchwith,
      indexfile2matchwith
    };
  }
  function compareArraysAndStoreMatchestring(columndata, columndatafile2) {

    for (let i = 0; i < columndata.length; i++) {
      let record = columndata[i];
  
      if (typeof record === "string" && columndatafile2.includes(record) && record !== null) {
          exact_matchedfile1.push(record);
          indexfile1.push(i);
          const indexInFile2 = columndatafile2.indexOf(record);
          if(i===0){
            indexfile1matchwith.push("*"+columndatafile2[0]+"("+matchtypes+")");
          }else{
            indexfile1matchwith.push(indexInFile2+1);
          }
      } else {
        non_exactmatchfile1.push(i)
        if(i===0){
          indexfile1matchwith.push("*"+columndatafile2[0]+"("+matchtypes+")");
        }else{
          indexfile1matchwith.push("no match");
        }
      }
  }
  

  for (let i = 0; i < columndatafile2.length; i++) {
    let record = columndatafile2[i];

    if (typeof record === "string" && columndata.includes(record) && record !== null) {
        exact_matchedfile2.push(record);
        indexfile2.push(i);
        
        const indexInFile2 = columndata.indexOf(record);
        if(i===0){
          indexfile2matchwith.push("*"+columndata[0]+"("+matchtypes+")");
        }else{
          indexfile2matchwith.push(indexInFile2+1);
        }
        
    } else {
      non_excatmatchedfile2.push(i)
      if(i===0){
        indexfile2matchwith.push("*"+columndata[0]+"("+matchtypes+")");
      }
        indexfile2matchwith.push("no match");
    }
}

    return {
      exact_matchedfile1,
      exact_matchedfile2,
      non_exactmatchfile1,
      non_excatmatchedfile2,
      indexfile1,
      indexfile2,
      indexfile1matchwith,
      indexfile2matchwith
    };
  }
  // sequential match function ______________________________
  function findMatches(sourceArray, targetArray, exactMatchArray, nonExactMatchArray, indexArray,indexmatchwith, expectedvalue) {

    function calculatePercent(value, percentage) {
      if (typeof value === "string") {
        return value.substring(0, Math.floor(value.length * (percentage / 100)));
      } else if (typeof value === "number" || !isNaN(parseFloat(value))) {
        const stringValue = value.toString();
        const numDigits = Math.ceil(stringValue.length * (percentage / 100));
        return parseInt(stringValue.substring(0, numDigits));
      }
      return null;
    }

    function calculatePercentageArray(inputArray, percentage) {
      return inputArray.map(value => calculatePercent(value, percentage));
    }

    let resultArray = calculatePercentageArray(sourceArray, expectedvalue);
    const targetSet = new Set(targetArray);
   
    for (let i = 0; i < resultArray.length; i++) {
        const searchValue = resultArray[i];
        let found = false;
       
        if (targetSet.has(searchValue)) {
            exactMatchArray.push(searchValue);
            indexArray.push(i);
            found = true;
            const targetIndex = targetArray.indexOf(searchValue);
            
      if (i === 0) {
        indexmatchwith.push("*"+targetArray[0]+"("+radioInput+")");

      }else{
        indexmatchwith.push((targetArray.indexOf(searchValue))+1); 
      }
            
        } else if (typeof searchValue === "string") {
         
            for (let j = 0; j < targetArray.length; j++) {
                const element = String(targetArray[j]); // Convert element to a string
               
                if (element.indexOf(searchValue) !== -1 ) {
                    // exactMatchArray.push(searchValue);
                    indexArray.push(i);
                    const targetIndex = targetArray.indexOf(searchValue);
                  
          if (i === 0) {
            indexmatchwith.push("*"+targetArray[0]+"("+radioInput+")");
          }else{
            indexmatchwith.push(j+1)
          }
                   
                    found = true;
                    break;
                }
            }
        } else if (typeof searchValue === "number") {
            const searchStr = searchValue.toString();
            for (let j = 0; j < targetArray.length; j++) {
              
                const element = String(targetArray[j]); // Convert element to a string
                
                if (element.indexOf(searchStr) !== -1 ) {
                   
                    // exactMatchArray.push(searchValue);
                    indexArray.push(i);
                    const targetIndex = targetArray.indexOf(searchValue);
          if (i === 0) {
            indexmatchwith.push("*"+targetArray[0]+"("+radioInput+")");
          }else{
            indexmatchwith.push(j+1)
          }
                    found = true;
                    break;
                }
            }
        }

        if (!found) {
          nonExactMatchArray.push(i);
          const targetIndex = targetArray.indexOf(searchValue);
          if (i === 0) {
            indexmatchwith.push("*"+targetArray[0]+"("+radioInput+")");
          }else{
            indexmatchwith.push("no match"); 
          }
            
        }
    }
    
    const valuesToAdd = indexArray.map(index => sourceArray[index]);
    exactMatchArray.push(...valuesToAdd);
 }
   
  function processMatches(radioInput, columndata1, columndata2) {
    findMatches(columndata1, columndata2, exact_matchedfile1, non_exactmatchfile1, indexfile1,indexfile1matchwith, expectedvaluefile1);
    findMatches(columndata2, columndata1, exact_matchedfile2, non_excatmatchedfile2, indexfile2,indexfile2matchwith, expectedvaluefile2);
  }

  //sectional match function_____________________
  function findMatchesectional(sourceArray, targetArray, exactMatchArray, nonExactMatchArray, indexArray,indexmatchwith, expectedvalue) {
    function calculatePercent(value, percentage) {
      if (typeof value === "string") {
        return value.substring(0, Math.floor(value.length * (percentage / 100)));
      } else if (typeof value === "number" || !isNaN(parseFloat(value))) {
        const stringValue = value.toString();
        const numDigits = Math.ceil(stringValue.length * (percentage / 100));
        return parseInt(stringValue.substring(0, numDigits));
      }
      return null;
    }

    function calculatePercentageArray(inputArray, percentage) {
      return inputArray.map(value => calculatePercent(value, percentage));
    }

    let resultArray = calculatePercentageArray(sourceArray, expectedvalue);
    const targetSet = new Set(targetArray);

    for (let i = 0; i < resultArray.length; i++) {
        const searchValue = resultArray[i];
        let found = false;

        if (targetSet.has(searchValue)) {
            exactMatchArray.push(searchValue);
            indexArray.push(i);
            found = true;
            if (i === 0) {
              indexmatchwith.push("*"+targetArray[0]+"("+radioInput+")");
            }else{
              indexmatchwith.push((targetArray.indexOf(searchValue))+1); 
            }
           
        } else if (typeof searchValue === "string") {
            for (let j = 0; j < targetArray.length; j++) {
                const element = String(targetArray[j]); // Convert element to a string

                if (element.indexOf(searchValue) !== -1 || isJumbledMatch(searchValue, element)) {
                    // exactMatchArray.push(searchValue);
                    indexArray.push(i);
                    if (i === 0) {
                      indexmatchwith.push("*"+targetArray[0]+"("+radioInput+")");
                    }else{
                      indexmatchwith.push(j+1)
                    }
                    
                    found = true;
                    break;
                }
            }
        } else if (typeof searchValue === "number") {
            const searchStr = searchValue.toString();
            for (let j = 0; j < targetArray.length; j++) {
                const element = String(targetArray[j]); // Convert element to a string
                if (element.indexOf(searchStr) !== -1 || isJumbledMatch(searchStr, element)) {
                    // exactMatchArray.push(searchValue);
                    indexArray.push(i);
                    if (i === 0) {
                      indexmatchwith.push("*"+targetArray[0]+"("+radioInput+")");
                    }else{
                      indexmatchwith.push(j+1)
                    }
                    found = true;
                    break;
                }
            }
        }

        if (!found) {
          nonExactMatchArray.push(i);
          const targetIndex = targetArray.indexOf(searchValue);
          if (i === 0) {
            indexmatchwith.push("*"+targetArray[0]+"("+radioInput+")");
          }else{
            indexmatchwith.push("no match"); 
          }
            
        }
    }
    const valuesToAdd = indexArray.map(index => sourceArray[index]);
    exactMatchArray.push(...valuesToAdd);

    return {
        exactMatchArray,
        nonExactMatchArray,
        indexArray
    };
}

function isJumbledMatch(searchValue, targetValue) {
  for (let i = 0; i <= targetValue.length - searchValue.length; i++) {
      const substring = targetValue.substring(i, i + searchValue.length);
      if (areStringsJumbledMatch(searchValue, substring)) {
          return true;
      }
  }
  return false;
}

  function areStringsJumbledMatch(str1, str2) {
    const chars1 = str1.split("").sort().join("");
    const chars2 = str2.split("").sort().join("");
    return chars1 === chars2;
  }

  function processMatchesectional(columndata1, columndata2) {
    findMatchesectional(columndata1, columndata2, exact_matchedfile1, non_exactmatchfile1, indexfile1,indexfile1matchwith, expectedvaluefile1);
    findMatchesectional(columndata2, columndata1, exact_matchedfile2, non_excatmatchedfile2, indexfile2,indexfile2matchwith, expectedvaluefile2);
  }
  if (matchtypes === "exact match") {
    compareArraysAndStoreMatches(columndata, columndatafile2)
  }
  else if (matchtypes === "exact number match") {
    compareArraysAndStoreMatchesnumeric(columndata, columndatafile2)
  }
  else if (matchtypes === "exact text match") {
    compareArraysAndStoreMatchestring(columndata, columndatafile2)
  }
  else if (radioInput === "sequential match") {
    processMatches(radioInput, columndata, columndatafile2);
  }
  else if (radioInput === "sectional match") {
    processMatchesectional(columndata, columndatafile2)
  }

  res.json({
    exact_matchedfile1,
    exact_matchedfile2,
    non_exactmatchfile1,
    non_excatmatchedfile2,
    indexfile1,
    indexfile2,
    indexfile1matchwith,
    indexfile2matchwith
  });
})



app.post("/download_file", async (req, res) => {
  let excelData = req.body.excelData;
  let exceldatafile2 = req.body.exceldatafile2;
  let indexfile1matchwith=req.body.file1matchindexwith;
  let indexfile2matchwith=req.body.file2matchindexwith;
  let notmatchedfile1=req.body.notmatchedfile1;
  let notmatchedfile2=req.body.notmatchedfile2
  let checkbox = req.body.checkbox;

  let inputfile1 = excelData;
  let inputfile2 = exceldatafile2;
   
 

  //handle matched fields for file 1
let result = [];
let currentRow = [];

for (let i = 0; i < indexfile1matchwith.length; i++) {
  const cellValue = indexfile1matchwith[i];

  if (typeof cellValue === 'string' && cellValue.startsWith('*')) {
    if (currentRow.length > 0) {
      result.push(currentRow);
      currentRow = [];
    }
    currentRow.push(cellValue);
  } else {
    currentRow.push(cellValue);
  }
}

if (currentRow.length > 0) {
  result.push(currentRow);
}

let finalResult = [];
for (let i = 0; i < result[0].length; i++) {
  finalResult.push(result.map(row => row[i]));
}



//handle matched fields for file 2
let result2 = [];
let currentRow2 = [];

for (let i = 0; i < indexfile2matchwith.length; i++) {
  const cellValue = indexfile2matchwith[i];

  if (typeof cellValue === 'string' && cellValue.startsWith('*')) {
    if (currentRow2.length > 0) {
      result2.push(currentRow2);
      currentRow2 = [];
    }
    currentRow2.push(cellValue);
  } else {
    currentRow2.push(cellValue);
  }
}

if (currentRow2.length > 0) {
  result2.push(currentRow2);
}

let finalResult2 = [];
for (let i = 0; i < result2[0].length; i++) {
  finalResult2.push(result2.map(row => row[i]));
}




const outputfile1combinedData = [];

const maxLength = Math.max(excelData.length, finalResult.length);

for (let i = 0; i < maxLength; i++) {
    const row = {};

    if (i < excelData.length) {
        const data1 = excelData[i];
        Object.assign(row, data1);
    }
   
     
      row["Matched"] = "file2-->"; // Add "MATCHED" for other rows
  
    if (i < finalResult.length) {
        const data2 = finalResult[i];
        Object.keys(data2).forEach((key) => {
            row[key + " "] = data2[key];
        });
    }
    outputfile1combinedData.push(row);
}

let firstObject = outputfile1combinedData[0];
firstObject['all criteria'] = 'all criteria';

for (let i = 1; i < outputfile1combinedData.length; i++) {
  let obj = outputfile1combinedData[i];
  let hasNoMatch = false;

  for (let [key, value] of Object.entries(obj)) {
    if (typeof value === 'string' && value.toLowerCase() === 'no match') {
      hasNoMatch = true;
      break;
    }
  }

  obj['Last Key'] = hasNoMatch ? 'NO MATCH' : 'MATCH';
}


const outputfile2combinedData = [];

const maxLength2 = Math.max(exceldatafile2.length, finalResult2.length);

for (let i = 0; i < maxLength2; i++) {
    const row = {};

    if (i < exceldatafile2.length) {
        const data1 = exceldatafile2[i];
        Object.assign(row, data1);
    }

      row["Matched"] = "file 1-->"; // Add "MATCHED" for other rows


    if (i < finalResult2.length) {
        const data2 = finalResult2[i];
        Object.keys(data2).forEach((key) => {
            row[key + " "] = data2[key];
        } );
    }
    outputfile2combinedData.push(row);
}

let firstObjectfil2 = outputfile2combinedData[0];
firstObjectfil2['all criteria'] = 'all criteria';

for (let i = 1; i < outputfile2combinedData.length; i++) {
  let obj = outputfile2combinedData[i];
  let hasNoMatch = false;

  for (let [key, value] of Object.entries(obj)) {
    if (typeof value === 'string' && value.toLowerCase() === 'no match') {
      hasNoMatch = true;
      break;
    }
  }

  obj['Last Key'] = hasNoMatch ? 'NO MATCH' : 'MATCH';
}

//creating workbooks_______________________
  let workbook = new ExcelJS.Workbook();
   const redColor = { argb: "FF2400" }; // Red color
   const greenColor = { argb: "4CBB17" }; // Green color
   const greyColor = { argb :"808080"}

// worksheet 1_____________________
if(checkbox===false){
  let worksheet1 = workbook.addWorksheet("Input File1");
  inputfile1.forEach((data, index) => {
    const row = worksheet1.addRow(data);
    // Exclude the first row from color filling
    if (index !== 0) {
      const fillColor = !notmatchedfile1.includes(index) ? greenColor : redColor;
      row.eachCell({ includeEmpty: true }, (cell) => {
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: fillColor,
        };
      });
    }else{
      row.eachCell({ includeEmpty: true }, (cell) => {
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: greyColor,
        };
      });
    }
  });
  
}else{
   // Add sheet for matchedData1
   let worksheet1 = workbook.addWorksheet("Input File1");
   inputfile1.forEach((data, index) => {
     const row = worksheet1.addRow(data);
     const fillColor = !notmatchedfile1.includes(index) ? greenColor : redColor
     row.eachCell({ includeEmpty: true }, (cell) => {
       cell.fill = {
         type: "pattern",
         pattern: "solid",
         fgColor: fillColor,
       };
     });
    })
  }


  // worksheet 2 ______________________
  if(checkbox===false){
    let worksheet2 = workbook.addWorksheet("Input File2");
    inputfile2.forEach((data, index) => {
      const row = worksheet2.addRow(data);  
      if (index !== 0) {
      const fillColor = !notmatchedfile2.includes(index) ? greenColor : redColor
      row.eachCell({ includeEmpty: true }, (cell) => {
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: fillColor,
        };
      });
    }else{
      row.eachCell({ includeEmpty: true }, (cell) => {
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: greyColor,
        };})
    }
    });
    
  }else{
   // Add sheet for matchedData2
   let worksheet2 = workbook.addWorksheet("Input File2");
  inputfile2.forEach((data, index) => {
    const row = worksheet2.addRow(data);
    const fillColor = !notmatchedfile2.includes(index) ? greenColor : redColor
    row.eachCell({ includeEmpty: true }, (cell) => {
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: fillColor,
      };
    });
  });
}
//  worksheet 3 _______________________
const worksheet3 = workbook.addWorksheet("Output File1");
const batchSize = 1000; // Adjust the batch size as needed

const totalRows = outputfile1combinedData.length;

for (let start = 0; start < totalRows; start += batchSize) {
  const end = Math.min(start + batchSize, totalRows);
  const batchData = outputfile1combinedData.slice(start, end);

  const rows = batchData.map(data => Object.values(data));
  worksheet3.addRows(rows);
}

// worksheet 4____________________
const worksheet4 = workbook.addWorksheet("Output File2");


const totalRows2 = outputfile2combinedData.length;

for (let start = 0; start < totalRows2; start += batchSize) {
  const end = Math.min(start + batchSize, totalRows2);
  const batchData = outputfile2combinedData.slice(start, end);

  const rows = batchData.map(data => Object.values(data));
  worksheet4.addRows(rows);
}


 // streaming process to send the data in chunks so that it can handle to load of large files 
  const stream = new require('stream').PassThrough();

  // Write the workbook to the stream
  workbook.xlsx.write(stream).then(() => {
    // Set headers
    res.setHeader('Content-Disposition', 'attachment; filename="data.xlsx"');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

    // Pipe the stream to the response
    stream.pipe(res);
  });

})





app.listen(8000, () => {
  console.log('Server is running on port 8000');
});
