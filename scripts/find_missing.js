// usage:
// put result of find missing browser console script to files.json
// node find_missing.js
// see result in missing.json

const fs = require('fs');

function checkMissingFiles(fileNames) {
  const missingFiles = [];
  for (let i = 0; i < fileNames.length; i++) {
    const fileName = fileNames[i];
    if (!fs.existsSync("docs/"+fileName)) {
      missingFiles.push(fileName);
    }
  }
  return missingFiles;
}

function saveMissingFiles(missingFiles, outputFile) {
  const data = JSON.stringify(missingFiles);
  fs.writeFileSync(outputFile, data);
  console.log(`Missing files saved to ${outputFile}`);
}

function runProgram(inputFile, outputFile) {
  // Read the input file
  const fileData = fs.readFileSync(inputFile, 'utf8');
  const jsonData = JSON.parse(fileData);
  const fileNames = jsonData;

  // Check for missing files
  const missingFiles = checkMissingFiles(fileNames);

  // Save missing files to output file
  saveMissingFiles(missingFiles, outputFile);
}

// Provide the input and output file paths
const inputFile = 'files.json';
const outputFile = 'missing.json';

// Run the program
runProgram(inputFile, outputFile);