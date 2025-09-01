import * as React from "react";
import { useState } from "react";
import { 
  Button, 
  Field, 
  Textarea, 
  tokens, 
  makeStyles, 
  Select,
  Option,
  Input,
  Spinner,
  Image
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
    background: "linear-gradient(135deg, #107C10 0%, #27AE27 100%)",
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
    gridTemplateColumns: "1fr 1fr",
    gap: "20px",
    "@media (max-width: 768px)": {
      gridTemplateColumns: "1fr",
    },
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
    ":hover": {
      transform: "translateY(-2px)",
      boxShadow: "0 4px 6px rgba(0, 0, 0, 0.1)",
    },
  },
  featureTitle: {
    fontSize: "18px",
    marginBottom: "15px",
    color: "#107C10",
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
      borderColor: "#107C10",
      boxShadow: "0 0 0 3px rgba(16, 124, 16, 0.2)",
    },
  },
  selectContainer: {
    position: "relative",
    marginBottom: "15px",
    zIndex: 1000,
  },
  selectField: {
    width: "100%",
    padding: "10px",
    border: "1px solid #e5e7eb",
    borderRadius: "6px",
    fontSize: "14px",
    backgroundColor: "white",
    position: "relative",
    zIndex: 1001,
  },
  outputSection: {
    marginTop: "20px",
    padding: "15px",
    background: "#f9fafb",
    borderRadius: "6px",
    borderLeft: "4px solid #107C10",
    position: "relative",
    zIndex: 1,
  },
  outputTitle: {
    marginBottom: "10px",
    color: "#107C10",
  },
  outputContent: {
    background: "white",
    padding: "15px",
    borderRadius: "4px",
    border: "1px solid #e5e7eb",
    minHeight: "100px",
    maxHeight: "200px",
    overflowY: "auto",
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
});

const CustomSelect = ({ className, value, onChange, children, ...props }) => {
  const styles = useStyles();
  
  return (
    <div className={styles.selectContainer}>
      <Select
        className={`${styles.selectField} ${className}`}
        value={value}
        onChange={onChange}
        {...props}
      >
        {children}
      </Select>
    </div>
  );
};

const AIAssistance = () => {
  const [isLoading, setIsLoading] = useState(false);
  const [activeFeature, setActiveFeature] = useState(null);
  const [prompt, setPrompt] = useState("");
  const [dataToAnalyze, setDataToAnalyze] = useState("");
  const [formulaToExplain, setFormulaToExplain] = useState("");
  const [analysisType, setAnalysisType] = useState("");
  const [dataFormat, setDataFormat] = useState("");
  const [insightLevel, setInsightLevel] = useState("");
  const [generatedContent, setGeneratedContent] = useState("");
  const [analysisResults, setAnalysisResults] = useState("");
  const [formulaExplanation, setFormulaExplanation] = useState("");
  const [recommendations, setRecommendations] = useState([]);

  const styles = useStyles();

  const handleLogout = () => {
    localStorage.removeItem("anaUserAuthenticated");
    localStorage.removeItem("anaUserEmail");
    window.location.reload();
  };

  const simulateAPICall = () => {
    return new Promise(resolve => {
      setTimeout(() => {
        resolve(true);
      }, 1500);
    });
  };

  const handleGenerateContent = async () => {
    if (!prompt) return;
    
    setIsLoading(true);
    setActiveFeature("generate");
    
    await simulateAPICall();
    
    const content = `Based on your request about "${prompt}", here's a comprehensive analysis:\n\n• Key insights and trends related to your topic\n• Potential data patterns and correlations\n• Recommended Excel formulas for analysis\n• Visualization suggestions for better data presentation\n• Actionable recommendations for decision-making`;
    
    setGeneratedContent(content);
    setIsLoading(false);
  };

  const handleAnalyzeData = async () => {
    if (!dataToAnalyze) return;
    
    setIsLoading(true);
    setActiveFeature("analyze");
    
    await simulateAPICall();
    
    const analysis = `Data Analysis Results:\n\n• Dataset contains ${Math.floor(dataToAnalyze.length / 10)} potential data points\n• Detected patterns suggest seasonal trends\n• Correlation coefficient: 0.${Math.floor(Math.random() * 8) + 2}\n• Recommended visualizations: Line chart, Scatter plot\n• Key insights: Data shows strong upward trend with periodic fluctuations`;
    
    setAnalysisResults(analysis);
    setIsLoading(false);
  };

  const handleExplainFormula = async () => {
    if (!formulaToExplain) return;
    
    setIsLoading(true);
    setActiveFeature("explain");
    
    await simulateAPICall();
    
    const explanation = `Formula Explanation: ${formulaToExplain}\n\n• Purpose: This formula calculates ${Math.random() > 0.5 ? 'aggregate values' : 'specific metrics'} based on your data\n• Components: Includes ${Math.floor(Math.random() * 5) + 2} functions working together\n• Usage: Best applied to ${Math.random() > 0.5 ? 'financial data' : 'operational metrics'}\n• Tips: Use with absolute references for consistent results across cells`;
    
    setFormulaExplanation(explanation);
    setIsLoading(false);
  };

  const handleGetRecommendations = async () => {
    setIsLoading(true);
    setActiveFeature("recommend");
    
    await simulateAPICall();
    
    const mockRecommendations = [
      "Consider using PivotTables for better data summarization",
      "Implement data validation rules to maintain data integrity",
      "Use conditional formatting to highlight key metrics",
      "Create dynamic charts that update with new data automatically",
      "Set up named ranges for easier formula management"
    ];
    
    setRecommendations(mockRecommendations);
    setIsLoading(false);
  };

  const insertIntoExcel = (content) => {
    if (content) {
      insertText(content);
    }
  };

  const regenerateContent = () => {
    setGeneratedContent("");
    handleGenerateContent();
  };

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
        {/* Generate Insights Section */}
        <div className={styles.aiFeature}>
          <h3 className={styles.featureTitle}>
            <span className={styles.featureIcon}>✨</span>
            Generate Insights
          </h3>
          <p className={styles.featureDescription}>
            Get AI-powered analysis and recommendations for your data.
          </p>
          <Input
            className={styles.inputField}
            style={{ minHeight: "auto" }}
            value={prompt}
            onChange={(e) => setPrompt(e.target.value)}
            placeholder="Describe your data or analysis needs..."
          />
          
          <CustomSelect
            value={analysisType}
            onChange={(e, data) => setAnalysisType(data.value)}
          >
            <Option value="">Select analysis type</Option>
            <Option value="trend">Trend Analysis</Option>
            <Option value="financial">Financial Analysis</Option>
            <Option value="statistical">Statistical Analysis</Option>
            <Option value="predictive">Predictive Modeling</Option>
            <Option value="comparative">Comparative Analysis</Option>
            <Option value="visualization">Data Visualization</Option>
            <Option value="optimization">Optimization</Option>
          </CustomSelect>
          
          <CustomSelect
            value={dataFormat}
            onChange={(e, data) => setDataFormat(data.value)}
          >
            <Option value="">Data format</Option>
            <Option value="tabular">Tabular Data</Option>
            <Option value="time-series">Time Series</Option>
            <Option value="categorical">Categorical Data</Option>
            <Option value="numerical">Numerical Data</Option>
            <Option value="mixed">Mixed Data Types</Option>
          </CustomSelect>
          
          <Button 
            appearance="primary" 
            onClick={handleGenerateContent}
            disabled={isLoading || !prompt}
          >
            Generate Insights
          </Button>

          {generatedContent && activeFeature === "generate" && (
            <div className={styles.outputSection}>
              <h4 className={styles.outputTitle}>Generated Insights</h4>
              <div className={styles.outputContent}>{generatedContent}</div>
              <div className={styles.buttonGroup}>
                <Button appearance="primary" onClick={() => insertIntoExcel(generatedContent)}>
                  Insert into Excel
                </Button>
                <Button appearance="secondary" onClick={regenerateContent}>
                  Regenerate
                </Button>
              </div>
            </div>
          )}
        </div>
        
        {/* Data Analysis Section */}
        <div className={styles.aiFeature}>
          <h3 className={styles.featureTitle}>
            <span className={styles.featureIcon}>📊</span>
            Analyze Data
          </h3>
          <p className={styles.featureDescription}>
            Paste your data for comprehensive analysis and pattern detection.
          </p>
          <Textarea
            className={styles.inputField}
            value={dataToAnalyze}
            onChange={(e) => setDataToAnalyze(e.target.value)}
            placeholder="Paste your data here (CSV, table, or description)..."
          />
          
          <CustomSelect
            value={insightLevel}
            onChange={(e, data) => setInsightLevel(data.value)}
          >
            <Option value="">Insight depth</Option>
            <Option value="basic">Basic Overview</Option>
            <Option value="detailed">Detailed Analysis</Option>
            <Option value="comprehensive">Comprehensive Report</Option>
            <Option value="executive">Executive Summary</Option>
          </CustomSelect>
          
          <Button 
            appearance="primary" 
            onClick={handleAnalyzeData}
            disabled={isLoading || !dataToAnalyze}
          >
            Analyze Data
          </Button>

          {analysisResults && activeFeature === "analyze" && (
            <div className={styles.outputSection}>
              <h4 className={styles.outputTitle}>Analysis Results</h4>
              <div className={styles.outputContent}>{analysisResults}</div>
              <div className={styles.buttonGroup}>
                <Button appearance="primary" onClick={() => insertIntoExcel(analysisResults)}>
                  Insert Analysis
                </Button>
              </div>
            </div>
          )}
        </div>

        {/* Formula Explanation Section */}
        <div className={styles.aiFeature}>
          <h3 className={styles.featureTitle}>
            <span className={styles.featureIcon}>🧮</span>
            Formula Assistant
          </h3>
          <p className={styles.featureDescription}>
            Get explanations and improvements for your Excel formulas.
          </p>
          <Input
            className={styles.inputField}
            style={{ minHeight: "auto" }}
            value={formulaToExplain}
            onChange={(e) => setFormulaToExplain(e.target.value)}
            placeholder="Enter Excel formula to explain or optimize..."
          />
          
          <Button 
            appearance="primary" 
            onClick={handleExplainFormula}
            disabled={isLoading || !formulaToExplain}
          >
            Explain Formula
          </Button>

          {formulaExplanation && activeFeature === "explain" && (
            <div className={styles.outputSection}>
              <h4 className={styles.outputTitle}>Formula Explanation</h4>
              <div className={styles.outputContent}>{formulaExplanation}</div>
              <div className={styles.buttonGroup}>
                <Button appearance="primary" onClick={() => insertIntoExcel(formulaExplanation)}>
                  Insert Explanation
                </Button>
              </div>
            </div>
          )}
        </div>

        {/* Excel Recommendations Section */}
        <div className={styles.aiFeature}>
          <h3 className={styles.featureTitle}>
            <span className={styles.featureIcon}>🚀</span>
            Excel Optimization
          </h3>
          <p className={styles.featureDescription}>
            Get personalized recommendations to optimize your Excel workflow.
          </p>
          <div style={{ textAlign: "center", padding: "30px 0" }}>
            <p>Discover ways to improve your spreadsheet efficiency and accuracy.</p>
            <Button
              appearance="primary"
              onClick={handleGetRecommendations}
              disabled={isLoading}
            >
              Get Recommendations
            </Button>
          </div>

          {recommendations.length > 0 && activeFeature === "recommend" && (
            <div className={styles.outputSection}>
              <h4 className={styles.outputTitle}>Optimization Tips</h4>
              <div className={styles.outputContent}>
                <ul>
                  {recommendations.map((rec, index) => (
                    <li key={index}>{rec}</li>
                  ))}
                </ul>
              </div>
              <div className={styles.buttonGroup}>
                <Button appearance="primary" onClick={() => insertIntoExcel(recommendations.join('\n• '))}>
                  Insert Recommendations
                </Button>
              </div>
            </div>
          )}
        </div>
      </div>

      {isLoading && (
        <div className={styles.loadingContainer}>
          <Spinner label="Ana is analyzing your data..." labelPosition="below" />
        </div>
      )}

      <div className={styles.footer}>
        <p>Powered by Ana AI - Excel Intelligence</p>
      </div>
    </div>
  );
};

export default AIAssistance;