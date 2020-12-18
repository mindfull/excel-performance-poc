import React, { ChangeEvent, useCallback, useState } from "react";
import logo from "./excel.svg";
import spinner from "./tail-spin.svg";
import "./App.css";
import generateData from "./generateData";
import generateSampleExcel from "./generateSampleExcel";

const LIMIT = 1000000;

function App() {
  const [amount, setAmount] = useState(10000);
  const [isLoading, setIsLoading] = useState(false);
  const [elapsedTime, setElapsedTime] = useState({
    dataGenerate: 0,
    sheetGenerate: 0,
    fileGenerate: 0,
  });

  const handleAmountChange = useCallback((e: ChangeEvent<HTMLInputElement>) => {
    const parsedValue = parseInt(e.target.value, 10);
    if (!isNaN(parsedValue)) {
      setAmount(Math.max(0, Math.min(LIMIT, parsedValue)));
    } else {
      setAmount(0);
    }
  }, []);

  const handleDownload = useCallback(() => {
    setIsLoading(true);
    requestAnimationFrame(() => {
      setTimeout(() => {
        const generateDataStartTime = new Date().getTime();
        const data = generateData(amount);
        const dataGenerate = new Date().getTime() - generateDataStartTime;

        const generateResultTime = generateSampleExcel(data);
        setElapsedTime({
          dataGenerate,
          ...generateResultTime,
        });
        setIsLoading(false);
      }, 20);
    });
  }, [amount]);

  return (
    <div className="App">
      <header className="App-header">
        <img src={logo} className="App-logo" alt="logo" />
        <div className="App-generate">
          <label htmlFor="amount-input">Data Size: </label>
          <input
            id="amount-input"
            className="App-generate-amount"
            type="number"
            min={0}
            max={LIMIT}
            value={amount || ""}
            onChange={handleAmountChange}
          />
        </div>
        <button
          disabled={amount === 0}
          className="App-download"
          onClick={handleDownload}
        >
          Download an Excel file
        </button>
        {elapsedTime.sheetGenerate ? (
          <>
            <h3 className="App-result-title">Elapsed Time</h3>
            <dl className="App-result">
              <div>
                <dt className="App-result-label">Generating Data</dt>
                <dd className="App-result-data">
                  {elapsedTime.dataGenerate}ms
                </dd>
              </div>
              <div>
                <dt className="App-result-label">Generating Sheet</dt>
                <dd className="App-result-data">
                  {elapsedTime.sheetGenerate}ms
                </dd>
              </div>
              <div>
                <dt className="App-result-label">Generating File</dt>
                <dd className="App-result-data">
                  {elapsedTime.fileGenerate}ms
                </dd>
              </div>
            </dl>
          </>
        ) : null}
      </header>
      {isLoading && (
        <div className="App-spinner">
          <img className="App-spinner-icon" src={spinner} alt="Loading..." />
          <p>Loading... Be patient, please.</p>
        </div>
      )}
    </div>
  );
}

export default App;
