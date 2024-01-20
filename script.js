const fs = require('fs');
const xlsx = require('xlsx');

// Function to parse the xlsx data and return an array of objects
function parseXLSX(filePath) {
  const workbook = xlsx.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const data = xlsx.utils.sheet_to_json(sheet);

  return data;
}

// Function to analyze and print employee information based on the specified criteria
function analyzeEmployeeData(employeeData) {
  const output = [];
  const processedEmployees = new Set();

  for (let i = 0; i < employeeData.length; i++) {
    const employee = employeeData[i];

    // Check if the employee has already been processed for any criterion
    if (processedEmployees.has(employee['Employee Name'])) {
      continue;
    }

    const currentDate = new Date(employee['Time']);
    const sevenDaysAgo = new Date(currentDate);
    sevenDaysAgo.setDate(currentDate.getDate() - 7);

    // Criteria: Worked for 7 consecutive days
    const consecutiveDays = employeeData.filter((e) => {
      const date = new Date(e['Time']);
      return (
        e['Employee Name'] === employee['Employee Name'] &&
        date >= sevenDaysAgo &&
        date <= currentDate
      );
    });

    if (consecutiveDays.length >= 7) {
      const startDate = consecutiveDays[0]['Time'];
      const endDate = consecutiveDays[consecutiveDays.length - 1]['Time'];

      const result = `Employee ${employee['Employee Name']} worked for 7 consecutive days from ${startDate} to ${endDate}.`;
      output.push(result);
      processedEmployees.add(employee['Employee Name']);
      console.log(result);
    }

    // Criteria: Less than 10 hours but greater than 1 hour between shifts
    const nextShift = employeeData.find((e) => {
      const date = new Date(e['Time']);
      return (
        e['Employee Name'] === employee['Employee Name'] &&
        date >= currentDate &&
        date <= currentDate && // Only check for the same day in this case
        (date - currentDate) / (1000 * 60 * 60) >= 1 &&
        (date - currentDate) / (1000 * 60 * 60) < 10
      );
    });

    if (nextShift) {
      const currentEndTime = new Date(employee['Time Out']);
      const nextStartTime = new Date(nextShift['Time']);

      const hoursBetweenShifts = (nextStartTime - currentEndTime) / (1000 * 60 * 60);

      if (hoursBetweenShifts > 1 && hoursBetweenShifts < 10) {
        const result = `Employee ${employee['Employee Name']} has less than 10 hours between shifts.`;
        output.push(result);
        processedEmployees.add(employee['Employee Name']);
        console.log(result);
      }
    }

    // Criteria: Worked for more than 14 hours in a single shift
    const currentStartTime = new Date(employee['Time']);
    const currentEndTime = new Date(employee['Time Out']);

    const hoursWorked = (currentEndTime - currentStartTime) / (1000 * 60 * 60);

    if (hoursWorked > 14) {
      const result = `Employee ${employee['Employee Name']} worked for more than 14 hours in a single shift.`;
      output.push(result);
      processedEmployees.add(employee['Employee Name']);
      console.log(result);
    }
  }

  return output;
}

// Main function to read the xlsx file, perform analysis, and write output to 'output.txt'
function main() {
  // Assuming the file is named 'assignment_timecard.csv' and is in the 'source' folder
  const filePath = 'Assignment_Timecard.csv';

  try {
    const employeeData = parseXLSX(filePath);
    const output = analyzeEmployeeData(employeeData);

    // Separate results with lines
    const separatedOutput = output.join('\n') + '\n' + '-'.repeat(50) + '\n';
    
    // Write output to 'output.txt'
    fs.writeFileSync('output.txt', separatedOutput);
    console.log('Output written to output.txt.');
  } catch (error) {
    console.error(`Error reading or processing the file: ${error.message}`);
  }
}

// Execute the main function
main();
