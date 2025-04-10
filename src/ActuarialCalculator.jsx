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


  const getCalculationBreakdown = (employee, index) => {
    const steps = [];
    const yearsUntilRetirement = 30;
    const salaryGrowth = index % 2 === 1 ? 0.02 : 0.04;
    const salaryGrowthFrequency = 2;
    const nextSalaryRaiseDate = new Date("2025-06-30");
  
    const startDate = new Date(employee["תאריך תחילת עבודה"]);
    const birthDate = new Date(employee["תאריך לידה"]);
    const currentDate = new Date();
    const salary = parseFloat(employee["שכר"]);
    const currentAge = currentDate.getFullYear() - birthDate.getFullYear();
    const employeeID = parseInt(employee["EmployeeID"]);
  
    let liability = 0;
  
    for (let t = 0; t < yearsUntilRetirement; t++) {
      let raiseYears = currentDate < nextSalaryRaiseDate
        ? Math.floor((t - 1) / salaryGrowthFrequency)
        : Math.floor(t / salaryGrowthFrequency);
  
      const projectedSalary = salary * Math.pow(1 + salaryGrowth, raiseYears);
      const futureAge = currentAge + t;
      const survivalProb = calculateSurvivalProbability(currentAge, t + 1);
      const mortality = getMortalityRate(futureAge + 1);
      const resignation = getResignationRate(futureAge);
      const discount = Math.pow(1 + (discountRateCurve[t] || 0.05), t);
  
      const paysSeveranceOnResign = isEven(employeeID);
      const resignationRate = paysSeveranceOnResign ? resignation : 0;
  
      const combinedRate = mortality + resignationRate;
      const benefit = projectedSalary * yearsUntilRetirement * (1 - 0.05);
      const presentValue = (benefit * survivalProb * combinedRate) / discount;
  
      steps.push({
        year: t + 1,
        projectedSalary: projectedSalary.toFixed(2),
        survivalProb: survivalProb.toFixed(5),
        mortality,
        resignation: resignationRate,
        discount: discount.toFixed(5),
        benefit: benefit.toFixed(2),
        presentValue: presentValue.toFixed(2),
      });
  
      liability += presentValue;
    }
  
    return steps;
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
  const calculateEmployeeLiability = (employee, index) => {
    const yearsUntilRetirement = 30;
    const salaryGrowth = index % 2 === 1 ? 0.02 : 0.04;
    const salaryGrowthFrequency = 2;
    const nextSalaryRaiseDate = new Date("2025-06-30");
  
    const startDate = new Date(employee["תאריך תחילת עבודה"]);
    const birthDate = new Date(employee["תאריך לידה"]);
    const clause14Percentage = parseFloat(employee["אחוז סעיף 14"] || 0); // New
    const currentDate = new Date();
    const leaveDate = employee["תאריך עזיבה"] ? new Date(employee["תאריך עזיבה"]) : null;
    const salary = parseFloat(employee["שכר"].toString().replace(/,/g, ""));
    const currentAge = currentDate.getFullYear() - birthDate.getFullYear();
    const yearsOfService = (currentDate - startDate) / (1000 * 3600 * 24 * 365);
    const employeeID = parseInt(employee["EmployeeID"]);
  
    if (!salary || isNaN(currentAge)) return "Invalid data";
  
    // Clause 14 fully applied — employer owes nothing
    if (clause14Percentage === 100) return (0).toFixed(2);
  
    let liability = 0;
  
    for (let t = 0; t < yearsUntilRetirement; t++) {
      if (leaveDate && currentDate > leaveDate) {
        break;
      }
  
      let raiseYears = 0;
      if (currentDate < nextSalaryRaiseDate) {
        raiseYears = Math.floor((t - 1) / salaryGrowthFrequency);
      } else {
        raiseYears = Math.floor(t / salaryGrowthFrequency);
      }
  
      const projectedSalary = salary * Math.pow(1 + salaryGrowth, raiseYears);
      const futureAge = currentAge + t;
      const survivalProb = calculateSurvivalProbability(currentAge, t + 1);
      const mortality = getMortalityRate(futureAge + 1);
      const resignation = getResignationRate(futureAge);
      const discount = Math.pow(1 + (discountRateCurve[t] || 0.05), t);
  
      const paysSeveranceOnResign = isEven(employeeID);
      const resignationRate = paysSeveranceOnResign ? resignation : 0;
      const combinedRate = mortality + resignationRate;
  
      const benefit = projectedSalary * yearsUntilRetirement;
  
      // 💡 Adjust liability by remaining Clause 14 % (e.g., if only 72% is covered, employer still owes 28%)
      const clause14Adjustment = 1 - (clause14Percentage / 100);
  
      let presentValue = (benefit * survivalProb * combinedRate * clause14Adjustment) / discount;
  
      if (employee["תשלום מהנכס"]) {
        const assetPayment = parseFloat(employee["תשלום מהנכס"]);
        presentValue += assetPayment / discount;
      }
  
      if (employee["השלמה בצ'ק"]) {
        const checkCompletion = parseFloat(employee["השלמה בצ'ק"]);
        presentValue += checkCompletion / discount;
      }
  
      liability += presentValue;
    }
  
    return liability.toFixed(2);
  };
  

  const exportEmployeeResultsToExcel = () => {
    const results = employeeData.map((employee) => {
      const liability = calculateEmployeeLiability(employee);
      return {
        EmployeeID: employee["EmployeeID"],
        Name: employee["שם"],
        LastName: employee["שם משפחה"],
        Age: employee["גיל"],
        Salary: employee["שכר"],
        Liability: liability,
      };
    });

    

    const worksheet = XLSX.utils.json_to_sheet(results);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Employee Liabilities");
    XLSX.writeFile(workbook, "employee_liabilities.xlsx");
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
  

  function formatExcelSerialDate(serialDate) {
    if (!serialDate || isNaN(serialDate)) return "";
  
    // Excel serial date starts on Jan 1, 1900, but Excel wrongly considers 1900 a leap year, so we subtract 2 days
    const excelEpoch = new Date(1900, 0, 1);
    const date = new Date(excelEpoch.getTime() + (serialDate - 2) * 86400000);
  
    return date.toLocaleDateString("he-IL"); // Or use another locale if needed
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
    const liability = calculateEmployeeLiability(employee,index );
    return (
      <tr key={index} onClick={() => {
        setSelectedEmployee(employee);
        setCalculationSteps(getCalculationBreakdown(employee));
      }}>
        <td>{index + 1}</td>
        <td>{employee["שם"]}</td>
        <td>{employee["שם משפחה"]}</td>
        <td>{calculateAge(employee["תאריך לידה"])}</td>
        <td>{employee["שכר"]}</td>
        <td>{employee["מין"]}</td>
        <td>{formatExcelSerialDate(employee["תאריך תחילת עבודה"])}</td>
        <td>{formatExcelSerialDate(employee["תאריך קבלת סעיף 14"])}</td>
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
