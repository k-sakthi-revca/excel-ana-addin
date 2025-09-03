import * as React from "react";
import { useState } from "react";
import {
  Button,
  Dropdown,
  Option,
  Radio,
  RadioGroup,
  Textarea,
  Input,
  makeStyles,
} from "@fluentui/react-components";

const useStyles = makeStyles({
  container: { 
    maxWidth: "1000px", 
    margin: "0 auto", 
    background: "white", 
    boxShadow: "0 4px 16px rgba(0, 0, 0, 0.08)", 
    overflow: "visible", 
    position: "relative", 
    zIndex: 1, 
    padding: "16px 16px 24px 16px",
    borderRadius: "8px",
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  },
  section: {
    display: "flex",
    flexDirection: "column",
    gap: "10px",
    marginBottom: "20px",
  },
  sectionLabel: {
    fontSize: "16px",
    fontWeight: "600",
    color: "#333",
    marginBottom: "4px",
  },
  dropdown: {
    width: "100%",
    height: "40px",
    borderRadius: "6px",
    border: "1px solid #e0e0e0",
    transition: "all 0.2s ease",
    ":hover": {
      borderColor: "#ef4444",
    },
  },
  inputField: {
    width: "100%",
    height: "70px",
    borderRadius: "6px",
    border: "1px solid #e0e0e0",
    fontSize: "15px",
    padding: "12px",
    transition: "all 0.2s ease",
    ":focus": {
      borderColor: "#ef4444",
      boxShadow: "0 0 0 2px rgba(37, 99, 235, 0.2)",
    },
  },
  radioGroup: {
    gap: "12px",
  },
  radio: {
    ":checked": {
      color: "#ef4444 !important",
    },
  },
  runButton: {
    marginTop: "16px",
    backgroundColor: "#ef4444",
    color: "white",
    fontWeight: "600",
    width: "100%",
    height: "44px",
    borderRadius: "6px",
    fontSize: "16px",
    transition: "all 0.2s ease",
    ":hover": {
      backgroundColor: "#ef4444",
      transform: "translateY(-1px)",
      boxShadow: "0 4px 8px rgba(37, 99, 235, 0.2)",
    },
    ":active": {
      transform: "translateY(0)",
    },
  },
  outputBox: {
    padding: "16px 0",
    fontSize: "15px",
    lineHeight: "1.5",
    whiteSpace: "pre-wrap",
    borderTop: "1px solid #e5e7eb",
    marginTop: "20px",
    color: "#333",
  },
  footer: {
    marginTop: "16px",
    textAlign: "center",
    color: "#6b7280",
    fontSize: "13px",
  },
  linkContainer: {
    marginTop: "8px",
  },
  link: {
    color: "#ef4444",
    marginRight: "8px",
    textDecoration: "none",
    fontWeight: "500",
    ":hover": {
      textDecoration: "underline",
    },
  },
});

const AIAssistance = () => {
  const styles = useStyles();

  // State
  const [task, setTask] = useState("Ask");
  const [prompt, setPrompt] = useState("");
  const [contextOption, setContextOption] = useState("selection");
  const [outputTarget, setOutputTarget] = useState("currentSheet");
  const [output, setOutput] = useState("");

  const handleRun = () => {
    // Placeholder AI response
    setOutput(
      "This sheet models cash flows for a potential investment under base, management and downside cases. " +
      "Inputs include operating assumptions, purchase multiples, transaction and financing fees, and debt assumptions. " +
      "The model calculates key metrics such as IRR, MOIC, and total value to paid-in ratio."
    );
  };

  return (
    <div className={styles.container}>
      {/* Dropdown */}
      <div className={styles.section}>
        <Dropdown
          className={styles.dropdown}
          value={task}
          onOptionSelect={(e, data) => setTask(data.optionValue)}
          appearance="outline"
        >
          <Option value="Ask">Ask</Option>
          <Option value="Summarize">Summarize</Option>
          <Option value="Explain">Explain</Option>
        </Dropdown>
      </div>

      {/* Prompt Input */}
      <div className={styles.section}>
        <Input
          className={styles.inputField}
          placeholder="Enter a prompt..."
          value={prompt}
          onChange={(e, data) => setPrompt(data.value)}
          appearance="outline"
        />
      </div>

      {/* Context Options */}
      <div className={styles.section}>
        <div className={styles.sectionLabel}>Context Options</div>
        <RadioGroup
          className={styles.radioGroup}
          value={contextOption}
          onChange={(e, data) => setContextOption(data.value)}
        >
          <Radio className={styles.radio} value="selection" label="Use Selection" />
          <Radio className={styles.radio} value="currentSheet" label="Use Current Sheet" />
          <Radio className={styles.radio} value="workbook" label="Use Workbook" />
        </RadioGroup>
      </div>

      {/* Output Target */}
      <div className={styles.section}>
        <div className={styles.sectionLabel}>Output Target</div>
        <RadioGroup
          className={styles.radioGroup}
          value={outputTarget}
          onChange={(e, data) => setOutputTarget(data.value)}
        >
          <Radio className={styles.radio} value="currentSheet" label="Insert in current sheet" />
          <Radio className={styles.radio} value="newSheet" label="Insert into new sheet" />
          <Radio className={styles.radio} value="clipboard" label="Copy to clipboard" />
        </RadioGroup>
      </div>

      {/* Run Button */}
      <Button className={styles.runButton} onClick={handleRun}>
        Run
      </Button>

      {/* Output Box */}
      {output && <div className={styles.outputBox}>{output}</div>}

      {/* Links at the bottom of the output */}
      {output && (
        <div className={styles.linkContainer}>
          <a href="#" className={styles.link}>[1]</a>
          <a href="#" className={styles.link}>[2]</a>
        </div>
      )}
    </div>
  );
};

export default AIAssistance;
