const fs = require('fs');
const xlsx = require('xlsx');


function parseXLSX(filePath) {
  const workbook = xlsx.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const data = xlsx.utils.sheet_to_json(sheet);

  return data;
}

function analyzeEmployeeData(employeeData) {
  const output = {
    condition1: [],
    condition2: [],
    condition3: [],
    processedEmployees: new Set(),
  };

  for (let i = 0; i < employeeData.length; i++) {
    const employee = employeeData[i];

    // Check if the employee has already been processed for any criterion
    if (output.processedEmployees.has(employee['Employee Name'])) {
      continue;
    }

    const currentDate = new Date(employee['Time']);
    const sevenDaysAgo = new Date(currentDate);
    sevenDaysAgo.setDate(currentDate.getDate() - 7);

    // Criteria 1: Worked for 7 consecutive days
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

      const result = `Condition 1: Employee ${employee['Employee Name']} worked for 7 consecutive days from ${startDate} to ${endDate}.`;
      output.condition1.push(result);
      output.processedEmployees.add(employee['Employee Name']);
      console.log(result);
    }

    // Criteria 2: Less than 10 hours but greater than 1 hour between shifts
    const nextShift = employeeData.find((e) => {
      const date = new Date(e['Time']);
      return (
        e['Employee Name'] === employee['Employee Name'] &&
        date > currentDate &&
        (date - currentDate) / (1000 * 60 * 60) > 1 &&
        (date - currentDate) / (1000 * 60 * 60) < 10
      );
    });

    if (nextShift) {
      const currentEndTime = new Date(employee['Time Out']);
      const nextStartTime = new Date(nextShift['Time']);

      const hoursBetweenShifts = (nextStartTime - currentEndTime) / (1000 * 60 * 60);

      if (hoursBetweenShifts > 1 && hoursBetweenShifts < 10) {
        const result = `Condition 2: Employee ${employee['Employee Name']} has less than 10 hours between shifts.`;
        output.condition2.push(result);
        output.processedEmployees.add(employee['Employee Name']);
        console.log(result);
      }
    }

    // Criteria 3: Worked for more than 14 hours in a single shift
    const shiftDuration = (new Date(employee['Time Out']) - currentDate) / (1000 * 60 * 60);

    if (shiftDuration > 14) {
      const result = `Condition 3: Employee ${employee['Employee Name']} worked for more than 14 hours in a single shift.`;
      output.condition3.push(result);
      output.processedEmployees.add(employee['Employee Name']);
      console.log(result);
    }
  }

  return output;
}

function main() {
  const filePath = 'Assignment_Timecard.csv';

  try {
    const employeeData = parseXLSX(filePath);
    const output = analyzeEmployeeData(employeeData);

    // Write output to 'output.txt'
    let outputText = '';

    if (output.condition1.length > 0) {
      outputText += output.condition1.join('\n') + '\n' + '-'.repeat(50) + '\n';
    }

    if (output.condition2.length > 0) {
      outputText += output.condition2.join('\n') + '\n' + '-'.repeat(50) + '\n';
    }

    if (output.condition3.length > 0) {
      outputText += output.condition3.join('\n');
    } else {
      outputText += 'Condition 3: No one worked for more than 14 hours in a single shift.';
    }

    fs.writeFileSync('output.txt', outputText);
    console.log('Output written to output.txt. Check Out!!');
  } catch (error) {
    console.error(`Error reading or processing the file: ${error.message}`);
  }
}

main();
