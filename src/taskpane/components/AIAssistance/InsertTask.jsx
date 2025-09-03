import * as React from "react";
import {
  Button,
  Radio,
  RadioGroup,
  Input,
  Spinner,
  Text,
} from "@fluentui/react-components";
import { useStyles } from "./styles";

const InsertTask = ({
  prompt,
  setPrompt,
  contextOption,
  setContextOption,
  outputTarget,
  setOutputTarget,
  handleRun,
  isLoading,
  insertionResult
}) => {
  const styles = useStyles();

  return (
    <div className={styles.section}>
      <div className={styles.sectionLabel}>Insert Content</div>
      
      <div className={styles.contextSection}>
        <RadioGroup
          className={styles.radioGroup}
          value={contextOption}
          onChange={(e, data) => setContextOption(data.value)}
        >
          <Radio className={styles.radio} value="narrative" label="Generate Narrative" />
          <Radio className={styles.radio} value="table" label="Generate Table" />
        </RadioGroup>
      </div>
      
      <Input
        className={styles.inputField}
        placeholder={
          contextOption === "narrative" 
            ? "Describe the narrative you want to generate..." 
            : "Describe the table you want to generate..."
        }
        value={prompt}
        onChange={(e, data) => setPrompt(data.value)}
        appearance="outline"
      />
      
      <div className={styles.outputSection}>
        <div className={styles.sectionLabel}>Output Options</div>
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
      
      <Button 
        className={styles.runButton} 
        onClick={handleRun}
        disabled={isLoading}
      >
        {isLoading ? <Spinner size="tiny" /> : "Generate"}
      </Button>
      
      {insertionResult && (
        <div className={styles.insertionResult}>
          <div className={styles.insertionHeader}>
            <Text weight="semibold" size={500}>
              {insertionResult.type === "narrative" ? "Generated Narrative" : "Generated Table"}
            </Text>
          </div>
          
          <div className={styles.insertionBody}>
            {insertionResult.type === "narrative" ? (
              <div className={styles.narrativeResult}>
                <Text block>{insertionResult.data.narrative}</Text>
                <div className={styles.narrativeMeta}>
                  <Text size={200}>
                    Word count: {insertionResult.data.wordCount} • 
                    Confidence: {Math.round(insertionResult.data.confidence * 100)}%
                  </Text>
                </div>
                <div className={styles.alternativeStyles}>
                  <Text size={200}>
                    Alternative styles: {insertionResult.data.alternativeStyles.join(", ")}
                  </Text>
                </div>
              </div>
            ) : (
              <div className={styles.tableResult}>
                <table className={styles.generatedTable}>
                  <thead>
                    <tr>
                      {insertionResult.data.tableData.headers.map((header, index) => (
                        <th key={index}>{header}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {insertionResult.data.tableData.rows.map((row, rowIndex) => (
                      <tr key={rowIndex}>
                        {row.map((cell, cellIndex) => (
                          <td key={cellIndex}>{cell}</td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                  <tfoot>
                    <tr>
                      {insertionResult.data.tableData.totals.map((total, index) => (
                        <td key={index}>{total}</td>
                      ))}
                    </tr>
                  </tfoot>
                </table>
                <div className={styles.tableMeta}>
                  <Text size={200}>
                    {insertionResult.data.tableData.summary} • 
                    {insertionResult.data.dimensions.rows} rows × 
                    {insertionResult.data.dimensions.columns} columns
                  </Text>
                </div>
              </div>
            )}
            
            <Button 
              className={styles.insertButton}
              onClick={() => {
                // In a real implementation, this would insert the content into Excel
                alert(`Content would be inserted into ${outputTarget}`);
              }}
            >
              Insert into {outputTarget === "currentSheet" ? "Current Sheet" : 
                           outputTarget === "newSheet" ? "New Sheet" : "Clipboard"}
            </Button>
          </div>
        </div>
      )}
    </div>
  );
};

export default InsertTask;
