import * as React from "react";
import { useState } from "react";
import { 
  Button, 
  Field, 
  Textarea, 
  tokens, 
  makeStyles, 
  Input,
  Spinner,
  Checkbox,
  Radio,
  RadioGroup,
} from "@fluentui/react-components";
import PropTypes from "prop-types";
import { insertText } from "../taskpane";

const useStyles = makeStyles({
  container: {
    maxWidth: "1000px",
    margin: "0 auto",
    background: "white",
    boxShadow: "0 2px 10px rgba(0, 0, 0, 0.1)",
    overflow: "visible",
    position: "relative",
    zIndex: 1,
  },
  header: {
    display: "flex",
    alignItems: "center",
    padding: "20px",
    background: "linear-gradient(135deg, #ef4444 0%, #FF6B33 100%)",
    color: "white",
    position: "relative",
  },
  headerContent: {
    marginLeft: "15px",
    flex: 1,
  },
  headerTitle: {
    fontSize: "24px",
    fontWeight: 600,
    marginBottom: "5px",
  },
  headerSubtitle: {
    opacity: 0.9,
  },
  logoutButton: {
    position: "absolute",
    top: "15px",
    right: "15px",
    background: "rgba(255, 255, 255, 0.2)",
    color: "white",
    border: "1px solid rgba(255, 255, 255, 0.3)",
    borderRadius: "4px",
    padding: "6px 12px",
    fontSize: "12px",
    fontWeight: 500,
    cursor: "pointer",
    transition: "all 0.2s ease",
    ":hover": {
      background: "rgba(255, 255, 255, 0.3)",
      borderColor: "rgba(255, 255, 255, 0.5)",
    },
  },
  mainContent: {
    padding: "20px",
    display: "grid",
    gridTemplateColumns: "1fr",
    gap: "20px",
    position: "relative",
    zIndex: 2,
  },
  aiFeature: {
    background: "#f9fafb",
    borderRadius: "8px",
    padding: "20px",
    boxShadow: "0 1px 3px rgba(0, 0, 0, 0.1)",
    transition: "transform 0.2s, box-shadow 0.2s",
    position: "relative",
    zIndex: 3,
    marginBottom: "20px",
    ":hover": {
      transform: "translateY(-2px)",
      boxShadow: "0 4px 6px rgba(0, 0, 0, 0.1)",
    },
  },
  featureTitle: {
    fontSize: "18px",
    marginBottom: "15px",
    color: "#ef4444",
    display: "flex",
    alignItems: "center",
  },
  featureIcon: {
    marginRight: "10px",
  },
  featureDescription: {
    marginBottom: "15px",
    color: "#6b7280",
  },
  inputField: {
    width: "100%",
    padding: "12px",
    border: "1px solid #e5e7eb",
    borderRadius: "6px",
    marginBottom: "15px",
    fontSize: "14px",
    resize: "vertical",
    minHeight: "100px",
    ":focus": {
      outline: "none",
      borderColor: "#ef4444",
      boxShadow: "0 0 0 3px rgba(16, 124, 16, 0.2)",
    },
  },
  outputSection: {
    marginTop: "20px",
    padding: "15px",
    background: "#f9fafb",
    borderRadius: "6px",
    borderLeft: "4px solid #ef4444",
    position: "relative",
    zIndex: 1,
  },
  outputTitle: {
    marginBottom: "10px",
    color: "#ef4444",
  },
  outputContent: {
    background: "white",
    padding: "15px",
    borderRadius: "4px",
    border: "1px solid #e5e7eb",
    minHeight: "100px",
    maxHeight: "200px",
    overflowY: "auto",
    fontFamily: "monospace",
  },
  buttonGroup: {
    marginTop: "10px",
    display: "flex",
    gap: "10px",
  },
  footer: {
    padding: "20px",
    textAlign: "center",
    background: "#f9fafb",
    color: "#6b7280",
    fontSize: "14px",
    borderTop: "1px solid #e5e7eb",
    position: "relative",
    zIndex: 1,
  },
  loadingContainer: {
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    padding: "20px",
    position: "relative",
    zIndex: 1,
  },
  exampleQueries: {
    marginTop: "10px",
    marginBottom: "15px",
  },
  exampleQuery: {
    display: "inline-block",
    margin: "5px",
    padding: "8px 12px",
    background: "#e5e7eb",
    borderRadius: "16px",
    fontSize: "12px",
    color: "#4b5563",
    cursor: "pointer",
    transition: "all 0.2s ease",
    ":hover": {
      background: "#d1d5db",
    },
  },
  debugInfo: {
    marginTop: "10px",
    padding: "10px",
    background: "#ffecb3",
    borderRadius: "4px",
    border: "1px solid #ffd54f",
  },
  optionsGrid: {
    display: "grid",
    gridTemplateColumns: "1fr 1fr",
    gap: "10px",
    marginBottom: "15px",
  },
  checkboxGroup: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    marginBottom: "15px",
  },
  infoBox: {
    padding: "10px",
    background: "#e8f4f8",
    borderRadius: "4px",
    marginBottom: "15px",
    fontSize: "14px",
    color: "#2c5282",
  },
  summaryResult: {
    background: "#f0fff4",
    padding: "15px",
    borderRadius: "4px",
    border: "1px solid #c6f6d5",
    marginTop: "15px",
  },
  summaryMetric: {
    display: "flex",
    justifyContent: "space-between",
    padding: "8px 0",
    borderBottom: "1px solid #e5e7eb",
  },
  summaryMetricLabel: {
    fontWeight: "500",
    color: "#4a5568",
  },
  summaryMetricValue: {
    fontWeight: "600",
    color: "#ef4444",
  },
});

const AIAssistance = () => {
  const [isLoading, setIsLoading] = useState(false);
  const [activeFeature, setActiveFeature] = useState(null);
  const [naturalLanguageQuery, setNaturalLanguageQuery] = useState("");
  const [formulaToExplain, setFormulaToExplain] = useState("");
  const [generatedFormula, setGeneratedFormula] = useState("");
  const [formulaExplanation, setFormulaExplanation] = useState("");
  const [formulaDebugging, setFormulaDebugging] = useState("");
  
  // Data Cleaning states
  const [cleaningOptions, setCleaningOptions] = useState({
    removeDuplicates: true,
    fillMissingValues: true,
    standardizeDates: false,
    standardizeNumbers: false,
    trimWhitespace: true,
    fixCapitalization: false,
  });
  const [cleaningResult, setCleaningResult] = useState("");
  
  // Data Summarization states
  const [summarizationRange, setSummarizationRange] = useState("current");
  const [dataSummary, setDataSummary] = useState("");
  
  // Cell Modification Demo states
  const [cellModificationText, setCellModificationText] = useState("");
  const [cellModificationResult, setCellModificationResult] = useState("");
  const [modificationOption, setModificationOption] = useState("replace");

  const styles = useStyles();

  const handleLogout = () => {
    localStorage.removeItem("anaUserAuthenticated");
    localStorage.removeItem("anaUserEmail");
    window.location.reload();
  };

  // Placeholder for AI API call
  const simulateAICall = (prompt, type) => {
    return new Promise(resolve => {
      setTimeout(() => {
        // Placeholder responses based on type
        if (type === "formula") {
          // Natural language to formula conversion
          let formula = "";
          
          if (prompt.toLowerCase().includes("average") && prompt.toLowerCase().includes("january")) {
            formula = "=AVERAGEIF(A2:A100,\"January\",B2:B100)";
          } else if (prompt.toLowerCase().includes("sum") && prompt.toLowerCase().includes("sales")) {
            formula = "=SUM(B2:B100)";
          } else if (prompt.toLowerCase().includes("count") && prompt.toLowerCase().includes("products")) {
            formula = "=COUNTIF(C2:C100,\"Product\")";
          } else if (prompt.toLowerCase().includes("max") || prompt.toLowerCase().includes("highest")) {
            formula = "=MAX(B2:B100)";
          } else if (prompt.toLowerCase().includes("min") || prompt.toLowerCase().includes("lowest")) {
            formula = "=MIN(B2:B100)";
          } else {
            formula = "=FORMULA_PLACEHOLDER(A1:Z100)";
          }
          
          resolve(formula);
        } else if (type === "explain") {
          // Formula explanation
          let explanation = "This formula ";
          
          if (prompt.includes("AVERAGE")) {
            explanation += "calculates the average of values in the specified range.";
          } else if (prompt.includes("SUM")) {
            explanation += "adds up all the values in the specified range.";
          } else if (prompt.includes("COUNT")) {
            explanation += "counts the number of cells that meet the specified criteria.";
          } else if (prompt.includes("MAX")) {
            explanation += "finds the maximum value in the specified range.";
          } else if (prompt.includes("MIN")) {
            explanation += "finds the minimum value in the specified range.";
          } else {
            explanation += "performs a calculation on the specified range.";
          }
          
          resolve({ explanation });
        } else if (type === "debug") {
          // Formula debugging
          let debugging = "";
          
          if (prompt.includes("#REF!")) {
            debugging = "Error: The formula contains an invalid cell reference. This might be because the cells you're referring to have been deleted.";
          } else if (prompt.includes("#VALUE!")) {
            debugging = "Error: The formula is using values of the wrong type. Check that all references contain numeric values when needed.";
          } else if (prompt.includes("#DIV/0!")) {
            debugging = "Error: The formula is attempting to divide by zero. Consider adding an IF statement to handle this case.";
          } else if (prompt.includes("#NAME?")) {
            debugging = "Error: The formula contains an unrecognized function name or range name. Check for typos.";
          } else if (prompt.includes("#NULL!")) {
            debugging = "Error: The formula uses an intersection of two ranges that don't intersect.";
          } else if (prompt.includes(",") && !prompt.includes(";") && prompt.includes(";") && !prompt.includes(",")) {
            debugging = "Warning: Your formula might be using the wrong delimiter for your Excel's regional settings.";
          } else if (prompt.includes("(") && !prompt.includes(")")) {
            debugging = "Error: Missing closing parenthesis in the formula.";
          } else {
            debugging = "No obvious errors detected. The formula appears to be syntactically correct.";
          }
          
          resolve({ debugging });
        } else if (type === "clean") {
          // Data cleaning
          const options = prompt;
          let result = "Data Cleaning Results:\n\n";
          
          if (options.removeDuplicates) {
            result += "• Removed 12 duplicate rows\n";
          }
          
          if (options.fillMissingValues) {
            result += "• Filled 8 missing values with appropriate data\n";
          }
          
          if (options.standardizeDates) {
            result += "• Standardized 15 date values to MM/DD/YYYY format\n";
          }
          
          if (options.standardizeNumbers) {
            result += "• Standardized 10 number formats (currency, percentages, etc.)\n";
          }
          
          if (options.trimWhitespace) {
            result += "• Trimmed unnecessary whitespace from 24 cells\n";
          }
          
          if (options.fixCapitalization) {
            result += "• Fixed capitalization in 18 text entries\n";
          }
          
          result += "\nCleaning complete! Your data is now consistent and ready for analysis.";
          
          resolve(result);
        } else if (type === "summarize") {
          // Data summarization
          const range = prompt;
          let summary = "";
          
          if (range === "current") {
            summary = "Data Summary:\n\n• Total sales: ₹2.3M\n• Highest in March (₹320K)\n• Lowest in July (₹98K)\n• Average monthly sales: ₹191.7K\n• Year-over-year growth: +12.4%\n• Top product category: Electronics (42% of sales)\n• Underperforming region: South (18% below target)";
          } else if (range === "selection") {
            summary = "Selection Summary:\n\n• 24 data points analyzed\n• Sum: ₹845K\n• Average: ₹35.2K\n• Median: ₹28.7K\n• Maximum: ₹112K (Row 14)\n• Minimum: ₹8.2K (Row 7)\n• Standard Deviation: ₹24.6K";
          } else if (range === "sheet") {
            summary = "Sheet Summary:\n\n• 120 rows of data analyzed\n• Total revenue: ₹4.8M\n• Profit margin: 22.4%\n• Best performing product: Product X (₹720K)\n• Customer segments: Enterprise (65%), SMB (25%), Consumer (10%)\n• Growth trend: Positive (+8.2% quarterly)";
          }
          
          resolve(summary);
        }
      }, 1500);
    });
  };

  const handleNaturalLanguageToFormula = async () => {
    if (!naturalLanguageQuery) return;
    
    setIsLoading(true);
    setActiveFeature("formula");
    
    try {
      // This would be replaced with an actual AI API call
      const formula = await simulateAICall(naturalLanguageQuery, "formula");
      setGeneratedFormula(formula);
    } catch (error) {
      console.error("Error generating formula:", error);
      setGeneratedFormula("Error generating formula. Please try again.");
    } finally {
      setIsLoading(false);
    }
  };

  const handleExplainFormula = async () => {
    if (!formulaToExplain) return;
    
    setIsLoading(true);
    setActiveFeature("explain");
    
    try {
      // This would be replaced with an actual AI API call
      const { explanation, debugging } = await simulateAICall(formulaToExplain, "explain");
      setFormulaExplanation(explanation);
      
      // Also check for potential issues with the formula
      const debugInfo = await simulateAICall(formulaToExplain, "debug");
      setFormulaDebugging(debugInfo.debugging);
    } catch (error) {
      console.error("Error explaining formula:", error);
      setFormulaExplanation("Error explaining formula. Please try again.");
      setFormulaDebugging("");
    } finally {
      setIsLoading(false);
    }
  };
  
  const handleCleanData = async () => {
    setIsLoading(true);
    setActiveFeature("clean");
    
    try {
      // This would be replaced with an actual AI API call
      const result = await simulateAICall(cleaningOptions, "clean");
      setCleaningResult(result);
    } catch (error) {
      console.error("Error cleaning data:", error);
      setCleaningResult("Error cleaning data. Please try again.");
    } finally {
      setIsLoading(false);
    }
  };
  
  const handleSummarizeData = async () => {
    setIsLoading(true);
    setActiveFeature("summarize");
    
    try {
      // This would be replaced with an actual AI API call
      const summary = await simulateAICall(summarizationRange, "summarize");
      setDataSummary(summary);
    } catch (error) {
      console.error("Error summarizing data:", error);
      setDataSummary("Error summarizing data. Please try again.");
    } finally {
      setIsLoading(false);
    }
  };
  
  const handleCellModification = async () => {
    if (!cellModificationText) return;
    
    setIsLoading(true);
    setActiveFeature("cellModify");
    
    try {
      // Simulate the cell modification process
      await new Promise(resolve => setTimeout(resolve, 1000));
      
      // Create a result message based on the selected option
      let result = "";
      if (modificationOption === "replace") {
        result = "✅ Successfully replaced content in selected cell(s) with:\n\n" + cellModificationText;
      } else if (modificationOption === "append") {
        result = "✅ Successfully appended to selected cell(s):\n\n" + cellModificationText;
      } else if (modificationOption === "prepend") {
        result = "✅ Successfully prepended to selected cell(s):\n\n" + cellModificationText;
      } else if (modificationOption === "format") {
        result = "✅ Successfully applied formatting to selected cell(s) with text:\n\n" + cellModificationText;
      }
      
      setCellModificationResult(result);
      
      // Actually modify the Excel cells
      // This would call the Excel API to modify the selected cells
      if (Office && Office.context && Office.context.document) {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function (asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            // Get the current selection
            const selectedText = asyncResult.value;
            let newContent = cellModificationText;
            
            // Modify the content based on the selected option
            if (modificationOption === "append") {
              // Append the new text to the existing content
              newContent = selectedText + cellModificationText;
            } else if (modificationOption === "prepend") {
              // Prepend the new text to the existing content
              newContent = cellModificationText + selectedText;
            } else if (modificationOption === "format") {
              // For format option, we'll just add some formatting markers
              // In a real implementation, this would use proper Excel formatting APIs
              newContent = "**" + cellModificationText + "**";
            }
            // For replace option, we just use the new text as is
            
            // Update the selected cells with the modified content
            Office.context.document.setSelectedDataAsync(
              newContent,
              { coercionType: Office.CoercionType.Text },
              function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                  console.error('Error:', asyncResult.error.message);
                }
              }
            );
          }
        });
      }
      
    } catch (error) {
      console.error("Error modifying cells:", error);
      setCellModificationResult("Error modifying cells. Please try again.");
    } finally {
      setIsLoading(false);
    }
  };

  const insertIntoExcel = (content) => {
    if (content) {
      insertText(content);
    }
  };

  const handleExampleClick = (example) => {
    setNaturalLanguageQuery(example);
  };
  
  const handleCleaningOptionChange = (option) => {
    setCleaningOptions({
      ...cleaningOptions,
      [option]: !cleaningOptions[option],
    });
  };

  const exampleQueries = [
    "Find average sales for January",
    "Sum total revenue for Q1",
    "Count products sold in December",
    "Find highest sales value",
    "Calculate growth percentage"
  ];

  return (
    <div className={styles.container}>
      <div className={styles.header}>
        <img src="assets/Ana_logo.png" alt="Ana Logo" width="60" height="60" />
        <div className={styles.headerContent}>
          <h1 className={styles.headerTitle}>Ana Excel Assistant</h1>
          <p className={styles.headerSubtitle}>Enhance your spreadsheets with AI-powered insights</p>
        </div>
        
        <button 
          className={styles.logoutButton}
          onClick={handleLogout}
          title="Logout"
        >
          Logout
        </button>
      </div>
      
      <div className={styles.mainContent}>
        {/* Natural Language to Formula Section */}
        <div className={styles.aiFeature}>
          <h3 className={styles.featureTitle}>
            <span className={styles.featureIcon}>🔄</span>
            Natural Language to Formula
          </h3>
          <p className={styles.featureDescription}>
            Describe what you want to calculate in plain English, and get the Excel formula.
          </p>
          
          <div className={styles.exampleQueries}>
            <p style={{ marginBottom: "8px", fontSize: "14px", color: "#6b7280" }}>Try these examples:</p>
            {exampleQueries.map((query, index) => (
              <span 
                key={index} 
                className={styles.exampleQuery}
                onClick={() => handleExampleClick(query)}
              >
                {query}
              </span>
            ))}
          </div>
          
          <Input
            className={styles.inputField}
            style={{ minHeight: "auto" }}
            value={naturalLanguageQuery}
            onChange={(e) => setNaturalLanguageQuery(e.target.value)}
            placeholder="E.g., Find average sales for January"
          />
          
          <Button 
            appearance="primary" 
            onClick={handleNaturalLanguageToFormula}
            disabled={isLoading || !naturalLanguageQuery}
            style={{backgroundColor:naturalLanguageQuery?"#ef4444":""}}
          >
            Generate Formula
          </Button>

          {generatedFormula && activeFeature === "formula" && (
            <div className={styles.outputSection}>
              <h4 className={styles.outputTitle}>Generated Formula</h4>
              <div className={styles.outputContent}>{generatedFormula}</div>
              <div className={styles.buttonGroup}>
                <Button appearance="primary" style={{backgroundColor:"#ef4444"}} onClick={() => insertIntoExcel(generatedFormula)}>
                  Insert into Excel
                </Button>
              </div>
            </div>
          )}
        </div>
        
        {/* Formula Explanation & Debugging Section */}
        <div className={styles.aiFeature}>
          <h3 className={styles.featureTitle}>
            <span className={styles.featureIcon}>🧮</span>
            Formula Explanation & Debugging
          </h3>
          <p className={styles.featureDescription}>
            Paste an Excel formula to get a plain English explanation and debugging suggestions.
          </p>
          <Input
            className={styles.inputField}
            style={{ minHeight: "auto" }}
            value={formulaToExplain}
            onChange={(e) => setFormulaToExplain(e.target.value)}
            placeholder="E.g., =SUMIFS(B2:B100,A2:A100,\January\,C2:C100,\Sales\)"
          />
          
          <Button 
            appearance="primary" 
            onClick={handleExplainFormula}
            disabled={isLoading || !formulaToExplain}
            style={{backgroundColor:formulaToExplain?"#ef4444":""}}
          >
            Explain & Debug Formula
          </Button>

          {formulaExplanation && activeFeature === "explain" && (
            <div className={styles.outputSection}>
              <h4 className={styles.outputTitle}>Formula Explanation</h4>
              <div className={styles.outputContent}>{formulaExplanation}</div>
              
              {formulaDebugging && (
                <div className={styles.debugInfo}>
                  <strong>Debugging Suggestions:</strong>
                  <p>{formulaDebugging}</p>
                </div>
              )}
              
              <div className={styles.buttonGroup}>
                <Button appearance="primary" style={{backgroundColor:"#ef4444"}} onClick={() => insertIntoExcel(formulaExplanation + (formulaDebugging ? "\n\nDebugging Suggestions:\n" + formulaDebugging : ""))}>
                  Insert Explanation
                </Button>
              </div>
            </div>
          )}
        </div>
        
        {/* Smart Data Cleaning Section */}
        <div className={styles.aiFeature}>
          <h3 className={styles.featureTitle}>
            <span className={styles.featureIcon}>🧹</span>
            Smart Data Cleaning
          </h3>
          <p className={styles.featureDescription}>
            Clean and standardize your data with one click. Select the operations you want to perform.
          </p>
          
          <div className={styles.infoBox}>
            First select your data in Excel, then choose cleaning options below.
          </div>
          
          <div className={styles.checkboxGroup}>
            <Checkbox 
              label="Remove duplicates" 
              checked={cleaningOptions.removeDuplicates}
              onChange={() => handleCleaningOptionChange('removeDuplicates')}
            />
            <Checkbox 
              label="Fill missing values" 
              checked={cleaningOptions.fillMissingValues}
              onChange={() => handleCleaningOptionChange('fillMissingValues')}
            />
            <Checkbox 
              label="Standardize dates" 
              checked={cleaningOptions.standardizeDates}
              onChange={() => handleCleaningOptionChange('standardizeDates')}
            />
            <Checkbox 
              label="Standardize numbers" 
              checked={cleaningOptions.standardizeNumbers}
              onChange={() => handleCleaningOptionChange('standardizeNumbers')}
            />
            <Checkbox 
              label="Trim whitespace" 
              checked={cleaningOptions.trimWhitespace}
              onChange={() => handleCleaningOptionChange('trimWhitespace')}
            />
            <Checkbox 
              label="Fix capitalization" 
              checked={cleaningOptions.fixCapitalization}
              onChange={() => handleCleaningOptionChange('fixCapitalization')}
            />
          </div>
          
          <Button 
            appearance="primary" 
            onClick={handleCleanData}
            style={{backgroundColor:"#ef4444"}}
            disabled={isLoading || Object.values(cleaningOptions).every(option => !option)}
          >
            Clean Data
          </Button>

          {cleaningResult && activeFeature === "clean" && (
            <div className={styles.outputSection}>
              <h4 className={styles.outputTitle}>Cleaning Results</h4>
              <div className={styles.outputContent}>{cleaningResult}</div>
              <div className={styles.buttonGroup}>
                <Button appearance="primary" style={{backgroundColor:"#ef4444"}} onClick={() => insertIntoExcel(cleaningResult)}>
                  Insert Results
                </Button>
              </div>
            </div>
          )}
        </div>
        
        {/* Data Summarization Section */}
        <div className={styles.aiFeature}>
          <h3 className={styles.featureTitle}>
            <span className={styles.featureIcon}>📊</span>
            Data Summarization
          </h3>
          <p className={styles.featureDescription}>
            Get a quick summary of your data with key metrics and insights.
          </p>
          
          <div className={styles.infoBox}>
            Select your data in Excel, then click "Summarize Data" to get insights.
          </div>
          
          <RadioGroup
            value={summarizationRange}
            onChange={(e, data) => setSummarizationRange(data.value)}
          >
            <Radio value="current" label="Current selection" />
            <Radio value="selection" label="Active table/range" />
            <Radio value="sheet" label="Entire worksheet" />
          </RadioGroup>
          
          <Button 
            appearance="primary" 
            onClick={handleSummarizeData}
            disabled={isLoading}
            style={{ marginTop: "15px", backgroundColor:"#ef4444" }}
          >
            Summarize Data
          </Button>

          {dataSummary && activeFeature === "summarize" && (
            <div className={styles.outputSection}>
              <h4 className={styles.outputTitle}>Data Summary</h4>
              <div className={styles.outputContent}>{dataSummary}</div>
              <div className={styles.buttonGroup}>
                <Button appearance="primary" style={{backgroundColor:"#ef4444"}} onClick={() => insertIntoExcel(dataSummary)}>
                  Insert Summary
                </Button>
              </div>
            </div>
          )}
        </div>
        
        {/* Cell Modification Demo Section */}
        <div className={styles.aiFeature}>
          <h3 className={styles.featureTitle}>
            <span className={styles.featureIcon}>✏️</span>
            Cell Modification Demo
          </h3>
          <p className={styles.featureDescription}>
            Demonstrate how to modify content in selected Excel cells. First select cells in Excel, then use options below.
          </p>
          
          <div className={styles.infoBox}>
            Select one or more cells in Excel, then enter text and choose how to modify the selected cells.
          </div>
          
          <Textarea
            className={styles.inputField}
            value={cellModificationText}
            onChange={(e) => setCellModificationText(e.target.value)}
            placeholder="Enter text to insert into selected cell(s)"
            style={{ minHeight: "80px" }}
          />
          
          <RadioGroup
            value={modificationOption}
            onChange={(e, data) => setModificationOption(data.value)}
            style={{ marginBottom: "15px" }}
          >
            <Radio value="replace" label="Replace cell content" />
            <Radio value="append" label="Append to existing content" />
            <Radio value="prepend" label="Prepend to existing content" />
            <Radio value="format" label="Format cells with text" />
          </RadioGroup>
          
          <Button 
            appearance="primary" 
            onClick={handleCellModification}
            disabled={isLoading || !cellModificationText}
            style={{ backgroundColor: cellModificationText ? "#ef4444" : "" }}
          >
            Modify Selected Cells
          </Button>

          {cellModificationResult && activeFeature === "cellModify" && (
            <div className={styles.outputSection}>
              <h4 className={styles.outputTitle}>Modification Result</h4>
              <div className={styles.outputContent}>{cellModificationResult}</div>
            </div>
          )}
        </div>
      </div>

      {isLoading && (
        <div className={styles.loadingContainer}>
          <Spinner label="Ana is analyzing your request..." labelPosition="below" />
        </div>
      )}

      <div className={styles.footer}>
        <p>Powered by Ana AI - Excel Intelligence</p>
      </div>
    </div>
  );
};

export default AIAssistance;
