import React, { useState, useEffect } from "react";
import * as XLSX from "xlsx";
import "./ActuarialCalculator.css";

const ActuarialCalculator = () => {
  const [employeeData, setEmployeeData] = useState([]);
  const [mortalityTable, setMortalityTable] = useState([]);
  const [error, setError] = useState(null);

  const [selectedEmployee, setSelectedEmployee] = useState(null);
const [calculationSteps, setCalculationSteps] = useState([]);

  const discountRateCurve = [
    0.0181, 0.0199, 0.0211, 0.0221, 0.023, 0.0239, 0.0246, 0.0253, 0.026, 0.0267,
    0.0274, 0.028, 0.0286, 0.0292, 0.0299, 0.0305, 0.0311, 0.0317, 0.0323, 0.0329,
    0.0335, 0.0341, 0.0348, 0.0354, 0.036, 0.0366, 0.0372, 0.0378, 0.0384, 0.0391,
    0.0397, 0.0403, 0.0409, 0.0415, 0.0421, 0.0427, 0.0434, 0.044, 0.0446, 0.0452,
    0.0458, 0.0464, 0.047, 0.0476, 0.0483, 0.0489, 0.0495,
  ];

  const resignationRates = [
    { range: [18, 29], dismissal: 0.07, resignation: 0.2 },
    { range: [30, 39], dismissal: 0.05, resignation: 0.13 },
    { range: [40, 49], dismissal: 0.04, resignation: 0.10 },
    { range: [50, 59], dismissal: 0.03, resignation: 0.07 },
    { range: [60, 67], dismissal: 0.02, resignation: 0.03 },
  ];

  const fetchDataFromAssets = async () => {
    try {
      const response = await fetch('/assets/data7.xlsx');
      if (!response.ok) throw new Error('Failed to fetch the Excel file');
      const arrayBuffer = await response.arrayBuffer();
      const data = new Uint8Array(arrayBuffer);
      const workbook = XLSX.read(data, { type: "array" });
      const employeeSheet = workbook.Sheets["data"]
        ? XLSX.utils.sheet_to_json(workbook.Sheets["data"])
        : [];

      const cleanedEmployeeSheet = employeeSheet.map(row =>
        Object.fromEntries(Object.entries(row).map(([key, value]) => [key.trim(), value]))
      );

      setEmployeeData(cleanedEmployeeSheet);
      setError(null);
    } catch (error) {
      console.error("Error loading the Excel file:", error);
      setError(error.message);
    }
  };

  const loadMortalityTable = async () => {
    try {
      const response = await fetch("/assets/mortality_table.xlsx");
      const data = await response.arrayBuffer();
      const workbook = XLSX.read(data, { type: "array" });
      const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);

      const mortalityData = jsonData
        .filter(row => !isNaN(row["age"]) && !isNaN(row["q(x)"]))
        .map(row => ({
          age: Number(row["age"]),
          qx: parseFloat(row["q(x)"]),
        }));

      setMortalityTable(mortalityData);
    } catch (error) {
      console.error("Error loading mortality data:", error);
    }
  };
  const discountRate = 0.04;

function getDiscountFactor(years) {
  return Math.pow(1 + discountRate, years);
}

function clause14Adjustment(t, yearsBeforeClause14, clause14Percentage) {
  if (t < yearsBeforeClause14) return 1;
  return (100 - clause14Percentage) / 100;
}

  const calculateEmployeeLiability = (employee, index) => {
    const salaryGrowth = index % 2 === 1 ? 0.02 : 0.04;
    const salaryGrowthFrequency = 2;
    const nextSalaryRaiseDate = new Date("2025-06-30");
  
    const startDate = parseDate2(employee["תאריך תחילת עבודה"]);
    const Date14 = parseDate2(employee["תאריך  קבלת סעיף 14"]);
    const clause14Percentage = parseFloat(employee["אחוז סעיף 14"] || 0);
    const currentDate = new Date('2024-12-31');
  
    const leaveDate = parseDate(employee["תאריך עזיבה"]);
    const salary = parseFloat(employee["שכר"].toString().replace(/,/g, ""));
    const currentAge = calculateAge(employee["תאריך לידה"]);
    const employeeID = 13;
    const gender = employee["מין"]?.trim();
    const retirementAge = gender === "נקבה" ? 64 : 67;
    const yearsUntilRetirement = Math.max(0, retirementAge - currentAge);
  
    let yearsBeforeClause14 = 0;
    if (startDate instanceof Date && !isNaN(startDate.getTime()) &&
        Date14 instanceof Date && !isNaN(Date14.getTime())) {
      yearsBeforeClause14 = Date14.getFullYear() - startDate.getFullYear();
      const startMonthDay = (startDate.getMonth() + 1) * 100 + startDate.getDate();
      const date14MonthDay = (Date14.getMonth() + 1) * 100 + Date14.getDate();
      if (date14MonthDay < startMonthDay) {
        yearsBeforeClause14 -= 1;
      }
      yearsBeforeClause14 = Math.max(0, yearsBeforeClause14);
    } else {
      yearsBeforeClause14 = 0;
    }
  
    if (!salary || isNaN(currentAge)) return "Invalid data";
  
    let liability = 0;
  
    for (let t = 0; t < yearsUntilRetirement; t++) {
      if (leaveDate && currentDate > leaveDate) {
        if (employeeID === index) console.log("DEBUG -> Left Company Before Calculation, Year:", t);
        return (0).toFixed(2);
      }
  
      let raiseYears = 0;
      if (currentDate < nextSalaryRaiseDate) {
        raiseYears = Math.floor((t - 1) / salaryGrowthFrequency);
      } else {
        raiseYears = Math.floor(t / salaryGrowthFrequency);
      }
  
      const projSalary = salary * Math.pow(1 + salaryGrowth, raiseYears);
      const futureAge = currentAge + t;
  
      // Survival probability for year t+1
      const survivalProb = calculateSurvivalProbability(currentAge, t + 1);
  
      // Mortality and resignation rates for futureAge + 1 (as per your table)
      const mortality = getMortalityRate(futureAge + 1);
      const resignation = getResignationRate(futureAge);
      const combinedRate = mortality + resignation;
  
      // Discount factor for year t+1
      const discount = getDiscountFactor(t + 1);
  
      // Years of service at time t (capped at yearsUntilRetirement)
      const yearsOfService = Math.min(t + (currentDate.getFullYear() - startDate.getFullYear()), yearsUntilRetirement);
  
      // Severance benefit: 1 month salary per year of service
      const benefit = projSalary * (yearsOfService / 12);
  
      // Clause 14 adjustment
      const clauseAdj = clause14Adjustment(t, yearsBeforeClause14, clause14Percentage);
  
      // Present value calculation
      const presentValue = (benefit * survivalProb * combinedRate * clauseAdj) / discount;
  
      if (employeeID === index) {
        console.log(`DEBUG -> Year: ${t}`, {
          raiseYears,
          projSalary,
          futureAge,
          survivalProb,
          mortality,
          resignation,
          combinedRate,
          benefit,
          clauseAdj,
          discount,
          presentValue
        });
      }
  
      liability += presentValue;
    }
  
    // Add asset payments if any
    ["תשלום מהנכס", "שווי נכס", "הפקדות", "השלמה בצ'ק"].forEach(key => {
      if (employee[key]) {
        const val = parseFloat(employee[key]);
        liability += val;
        if (employeeID === index) console.log(`DEBUG -> Added ${key}:`, val);
      }
    });
  
    if (employeeID === index) console.log("DEBUG -> Final Liability:", liability.toFixed(2));
  
    return liability.toFixed(2);
  };
  


  
  const getMortalityRate = (age) => {
  
    const data = mortalityTable.find((item) => item.age === age);
    return data ? data.qx : 0;
  };

  const getResignationRate = (age) => {
    for (const group of resignationRates) {
      const [min, max] = group.range;
      if (age >= min && age <= max) return group.resignation;
    }
    return 0.1; // Default fallback
  };

  const isEven = (n) => n % 2 === 0;

  const calculateSurvivalProbability = (age, years) => {
    let probability = 1;
    for (let i = 0; i < years; i++) {
      const mortalityRate = getMortalityRate(age + i);
      probability *= 1 - mortalityRate;
    }
    return probability;
  };



const parseDate = (dateInput) => {
  if (!dateInput) return null;
  if (!isNaN(dateInput)) {
      // If it's a number, treat as Excel serial
      return formatExcelSerialDate(dateInput);
  } else if (typeof dateInput === 'string') {
      // If it's a string like "9.1.2014"
      const parts = dateInput.split('.');
      if (parts.length === 3) {
          const [day, month, year] = parts.map(Number);
          return new Date(year, month - 1, day); // JS months are 0-based
      }
  }
  return null;
};


function formatExcelSerialDate(serialDate) {
  if (!serialDate || isNaN(serialDate)) return "";

  // Excel serial date starts on Jan 1, 1900, but Excel wrongly considers 1900 a leap year, so we subtract 2 days
  const excelEpoch = new Date(1900, 0, 1);
  const date = new Date(excelEpoch.getTime() + (serialDate - 2) * 86400000);

  return date.toLocaleDateString("he-IL"); // Or use another locale if needed
}


const parseDate2 = (dateInput) => {
  if (!dateInput) {
    return null;
  }
  if (typeof dateInput === "string") {
    const parts = dateInput.split(".");
    if (parts.length === 3) {
      const [day, month, year] = parts.map(Number);
      if (isNaN(day) || isNaN(month) || isNaN(year)) {
        return null;
      }
      const date = new Date(year, month - 1, day);
      if (isNaN(date.getTime())) {
        return null;
      }
      return date;
    }
  
    return null;
  }
  if (!isNaN(dateInput)) {

    return formatExcelSerialDate2(dateInput);
  }

  return null;
};

function formatExcelSerialDate2(serialDate) {
  if (!serialDate || isNaN(serialDate)) {

    return null;
  }
  const excelEpoch = new Date(1900, 0, 1);
  const date = new Date(excelEpoch.getTime() + (serialDate - 2) * 86400000);
  if (isNaN(date.getTime())) {

    return null;
  }
  return date; // Return Date object, not string
}

const exportEmployeeResultsToExcel = () => {
  const results = employeeData.map((employee, index) => {
    const employeeID = index + 1;
    const liability = calculateEmployeeLiability(employee, employeeID);
    
    return {
      EmployeeID: employeeID,
      Name: employee["שם"],
      LastName: employee["שם משפחה"],
      Age: calculateAge(employee["תאריך לידה"]),
      Salary: employee["שכר"],
      Gender: employee["מין"],
      StartJob: formatExcelSerialDate(employee["תאריך תחילת עבודה"]),
      DateOfGetting14: formatExcelSerialDate(employee["תאריך  קבלת סעיף 14"]),
      PercentOf14: employee["אחוז סעיף 14"],
      ValueOfProperty: employee["שווי נכס"],
      Deposits: employee["הפקדות"],
      DateOfLeave: formatExcelSerialDate(employee["תאריך עזיבה"]),
      PayFromProperty: employee["תשלום מהנכס"],
      Check: employee["השלמה בצ'ק"],
      ReasonOfLeaving: employee["סיבת עזיבה"],
      Liability: liability,
    };
  });

  const worksheet = XLSX.utils.json_to_sheet(results);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Employee Results");
  XLSX.writeFile(workbook, "employee_results.xlsx");
};


  useEffect(() => {
    loadMortalityTable();
    fetchDataFromAssets();
  }, []);


  
  function calculateAge(serialDate) {
    // Excel's serial date 1 is January 1, 1900.
    const excelEpoch = new Date(1900, 0, 1); // January 1, 1900
    
    // Convert serial date to JavaScript Date object
    const birthDate = new Date(excelEpoch.getTime() + (serialDate - 2) * 86400000); // Subtract 2 because Excel's serial date starts at 1 for Jan 1, 1900
    
    const today = new Date();
    let age = today.getFullYear() - birthDate.getFullYear();
    const m = today.getMonth() - birthDate.getMonth();
  
    // Adjust age if the birth date hasn't occurred yet this year
    if (m < 0 || (m === 0 && today.getDate() < birthDate.getDate())) {
      age--;
    }
  
    return age;
  }
  

  
  

  return (
    <div className="actuarial-app">
      <section className="results">
        <h2>📊 Employee Results</h2>
        <table>
          <thead>
            <tr>
              <th>Employee ID</th>
              <th>Name</th>
              <th>Last Name</th>
              <th>Age</th>
              <th>Salary</th>
              <th>Gender</th>
              <th>Start Job</th>
              <th>Date of getting 14</th>
              <th>Precent of 14</th>
              <th>value of propety</th>
              <th>Deposits</th>
              <th>Date of leave</th>
              <th>Pay from propety</th>
              <th>Check</th>
              <th>reason of leaving</th>
              <th>Liability</th> 
            </tr>
          </thead>
          <tbody>
  {employeeData.map((employee, index) => {
     const employeeID = index + 1;
     const liability = calculateEmployeeLiability(employee, employeeID);
    return (
      <tr key={employeeID} onClick={() => {
        setSelectedEmployee(employee);
        setCalculationSteps(getCalculationBreakdown(employee));
      }}>
        <td>{employeeID}</td>
        <td>{employee["שם"]}</td>
        <td>{employee["שם משפחה"]}</td>
        <td>{calculateAge(employee["תאריך לידה"])}</td>
        <td>{employee["שכר"]}</td>
        <td>{employee["מין"]}</td>
        <td>{formatExcelSerialDate(employee["תאריך תחילת עבודה"])}</td>
        <td>{formatExcelSerialDate(employee["תאריך  קבלת סעיף 14"])}</td>
        <td>{employee["אחוז סעיף 14"]}</td>
        <td>{employee["שווי נכס"]}</td>
        <td>{employee["הפקדות"]}</td>
        <td>{formatExcelSerialDate(employee["תאריך עזיבה"])}</td>
        <td>{employee["תשלום מהנכס"]}</td>
        <td>{employee["השלמה בצ'ק"]}</td>
        <td>{employee["סיבת עזיבה"]}</td>
        <td>{liability}</td>
      </tr>
    );
  })}
</tbody>

        </table>

        {selectedEmployee && (
  <section className="employee-breakdown">
    <h3>📋 Calculation Breakdown for {selectedEmployee["שם"]} {selectedEmployee["שם משפחה"]}</h3>
    <table>
      <thead>
        <tr>
          <th>Year</th>
          <th>Projected Salary</th>
          <th>Survival Probability</th>
          <th>Mortality</th>
          <th>Resignation</th>
          <th>Discount Factor</th>
          <th>Benefit</th>
          <th>Present Value</th>
        </tr>
      </thead>
      <tbody>
        {calculationSteps.map((step, idx) => (
          <tr key={idx}>
            <td>{step.year}</td>
            <td>{step.projectedSalary}</td>
            <td>{step.survivalProb}</td>
            <td>{step.mortality}</td>
            <td>{step.resignation}</td>
            <td>{step.discount}</td>
            <td>{step.benefit}</td>
            <td>{step.presentValue}</td>
          </tr>
        ))}
      </tbody>
    </table>
    <button onClick={() => setSelectedEmployee(null)}>❌ Close</button>
  </section>
)}

        <button onClick={exportEmployeeResultsToExcel}>📤 Export All Employee Results to Excel</button>
      </section>
    </div>
  );
};

export default ActuarialCalculator;
