/* 
 * Mock API and data for Excel Add-in ribbon items
 * This file provides mock implementations for the various ribbon functionalities
 */

// Mock user data
export const mockUserData = {
  id: "user123",
  name: "John Doe",
  email: "john.doe@example.com",
  apiKey: "mock-api-key-12345",
  organization: "Contoso Ltd.",
  plan: "Enterprise",
  usageCredits: 1000,
  lastLogin: "2023-06-15T10:30:00Z"
};

// Mock login API
export const loginApi = {
  login: async (email, password) => {
    // Simulate network delay
    await new Promise(resolve => setTimeout(resolve, 800));
    
    // Simple validation
    if (!email || !password) {
      throw new Error("Email and password are required");
    }
    
    if (!email.includes("@")) {
      throw new Error("Invalid email format");
    }
    
    if (password.length < 6) {
      throw new Error("Password must be at least 6 characters");
    }
    
    // Always succeed for mock
    return {
      success: true,
      user: mockUserData,
      token: "mock-jwt-token-xyz"
    };
  },
  
  logout: async () => {
    // Simulate network delay
    await new Promise(resolve => setTimeout(resolve, 300));
    
    return { success: true };
  },
  
  validateToken: async (token) => {
    // Simulate network delay
    await new Promise(resolve => setTimeout(resolve, 200));
    
    if (!token) {
      return { valid: false };
    }
    
    return { 
      valid: true,
      user: mockUserData
    };
  }
};

// Mock DDQ (Due Diligence Questionnaire) Assistant API
export const ddqApi = {
  // Sample DDQ questions and responses
  sampleQuestions: [
    {
      id: "ddq1",
      category: "Financial",
      question: "What was the company's revenue for the past 3 fiscal years?",
      importance: "high"
    },
    {
      id: "ddq2",
      category: "Operations",
      question: "Describe the company's supply chain and key dependencies.",
      importance: "medium"
    },
    {
      id: "ddq3",
      category: "Legal",
      question: "Are there any pending litigation matters?",
      importance: "high"
    },
    {
      id: "ddq4",
      category: "Market",
      question: "Who are the company's main competitors?",
      importance: "medium"
    },
    {
      id: "ddq5",
      category: "Technology",
      question: "Describe the company's IT infrastructure and security protocols.",
      importance: "medium"
    }
  ],
  
  // Generate DDQ questions based on context
  generateQuestions: async (context) => {
    // Simulate network delay
    await new Promise(resolve => setTimeout(resolve, 1200));
    
    // Return sample questions with some filtering based on context
    let questions = [...ddqApi.sampleQuestions];
    
    if (context && context.toLowerCase().includes("financ")) {
      questions = questions.filter(q => q.category === "Financial");
    } else if (context && context.toLowerCase().includes("legal")) {
      questions = questions.filter(q => q.category === "Legal");
    }
    
    return {
      success: true,
      questions: questions,
      context: context || "general"
    };
  },
  
  // Answer a DDQ question based on available data
  answerQuestion: async (questionId, data) => {
    // Simulate network delay
    await new Promise(resolve => setTimeout(resolve, 1500));
    
    const mockAnswers = {
      "ddq1": "The company reported revenues of $10.2M (FY2020), $12.5M (FY2021), and $15.8M (FY2022), showing consistent growth.",
      "ddq2": "The company relies on 3 primary manufacturers in Asia and distributes through a network of 12 regional warehouses across North America and Europe.",
      "ddq3": "There is one ongoing patent infringement case (Smith v. Company) and two minor contract disputes, with total contingent liability estimated at $1.2M.",
      "ddq4": "Main competitors include XYZ Corp (35% market share), ABC Inc (22%), and regional players comprising the remaining 43% of the market.",
      "ddq5": "The company uses a hybrid cloud infrastructure with AWS and on-premises servers. Security includes SOC2 compliance, quarterly penetration testing, and a dedicated security team."
    };
    
    return {
      success: true,
      question: ddqApi.sampleQuestions.find(q => q.id === questionId) || { question: "Unknown question" },
      answer: mockAnswers[questionId] || "No data available to answer this question.",
      confidence: questionId in mockAnswers ? 0.92 : 0.4,
      sources: ["Annual Report", "Management Interviews", "Market Analysis"]
    };
  }
};

// Mock Explain Sheet API
export const explainSheetApi = {
  // Analyze and explain sheet structure and data
  explainSheet: async (sheetData) => {
    // Simulate network delay
    await new Promise(resolve => setTimeout(resolve, 2000));
    
    // Default explanation if no specific pattern is detected
    let explanation = "This appears to be a general data sheet with mixed content.";
    let insights = ["Consider adding headers for better organization"];
    let dataQuality = "medium";
    
    // Check for financial data patterns
    if (sheetData && (
        sheetData.includes("revenue") || 
        sheetData.includes("expense") || 
        sheetData.includes("profit") ||
        sheetData.includes("budget")
    )) {
      explanation = "This sheet contains financial data, likely a P&L statement or budget forecast.";
      insights = [
        "The data shows a positive trend in quarterly revenue",
        "Expenses appear to be growing at a slower rate than revenue",
        "Q3 shows the highest profit margin at approximately 22%"
      ];
      dataQuality = "high";
    }
    
    // Check for project management data
    else if (sheetData && (
        sheetData.includes("task") || 
        sheetData.includes("deadline") || 
        sheetData.includes("project") ||
        sheetData.includes("status")
    )) {
      explanation = "This sheet appears to be a project management tracker or task list.";
      insights = [
        "There are 3 tasks currently marked as 'Overdue'",
        "The 'Marketing Campaign' project has the most pending tasks",
        "Resource allocation appears imbalanced, with Team A assigned to 65% of tasks"
      ];
      dataQuality = "medium";
    }
    
    return {
      success: true,
      explanation: explanation,
      structure: {
        rowCount: 150,
        columnCount: 15,
        dataRegions: 3,
        hasFormulas: true,
        hasCharts: false
      },
      insights: insights,
      dataQuality: dataQuality,
      recommendations: [
        "Consider adding data validation to improve consistency",
        "Adding conditional formatting would highlight key trends",
        "Some columns could be consolidated for better analysis"
      ]
    };
  },
  
  // Explain specific formula or cell
  explainCell: async (cellReference, formula) => {
    // Simulate network delay
    await new Promise(resolve => setTimeout(resolve, 800));
    
    // Mock explanations for different formula types
    const explanations = {
      "VLOOKUP": "This VLOOKUP formula searches for a value in the first column of a table and returns a value in the same row from a column you specify.",
      "INDEX/MATCH": "This formula combination uses MATCH to find the position of a lookup value, and INDEX to retrieve the value at that position in a range.",
      "SUMIFS": "This SUMIFS formula adds values that meet multiple criteria, filtering data before performing the sum operation.",
      "SUMPRODUCT": "This SUMPRODUCT formula multiplies corresponding components in the given arrays, and returns the sum of those products.",
      "IF": "This IF formula performs a logical test and returns one value for TRUE and another for FALSE.",
      "default": "This formula performs a calculation on the referenced cells."
    };
    
    let explanation = explanations.default;
    
    // Check formula type
    if (formula) {
      Object.keys(explanations).forEach(key => {
        if (formula.toUpperCase().includes(key)) {
          explanation = explanations[key];
        }
      });
    }
    
    return {
      success: true,
      cellReference: cellReference,
      formula: formula,
      explanation: explanation,
      dependencies: ["B2:B10", "Sheet2!A1"],
      potentialIssues: formula && formula.length > 100 ? ["Formula is complex and may be difficult to maintain"] : []
    };
  }
};

// Mock Summarize Selection API
export const summarizeApi = {
  // Summarize selected data
  summarizeSelection: async (selectionData, options = {}) => {
    // Simulate network delay
    await new Promise(resolve => setTimeout(resolve, 1800));
    
    // Default summary
    let summary = "The selected data contains numeric values with an average of 42.5.";
    let keyPoints = ["Data ranges from 10 to 75", "Contains 25 data points"];
    let chartRecommendation = "bar";
    
    // Check for time series data
    if (selectionData && selectionData.includes("date") || selectionData.includes("month") || selectionData.includes("year")) {
      summary = "The selected data appears to be a time series, showing values over a 12-month period.";
      keyPoints = [
        "Values peak in Q3 (July-September)",
        "Year-over-year growth rate is approximately 12%",
        "December shows the lowest values consistently"
      ];
      chartRecommendation = "line";
    }
    
    // Check for categorical data
    else if (selectionData && selectionData.includes("category") || selectionData.includes("region") || selectionData.includes("product")) {
      summary = "The selected data contains categorical information across different regions or product categories.";
      keyPoints = [
        "Region 'North' has the highest total value",
        "Product category 'Electronics' represents 45% of total",
        "3 categories account for 80% of the total value"
      ];
      chartRecommendation = "pie";
    }
    
    return {
      success: true,
      summary: summary,
      keyPoints: keyPoints,
      statistics: {
        count: 25,
        mean: 42.5,
        median: 38.2,
        min: 10,
        max: 75,
        stdDev: 18.3
      },
      chartRecommendation: chartRecommendation,
      insightLevel: options.detailLevel || "medium"
    };
  },
  
  // Generate a narrative based on data
  generateNarrative: async (data, style = "business") => {
    // Simulate network delay
    await new Promise(resolve => setTimeout(resolve, 2200));
    
    const narrativeStyles = {
      "business": "The data indicates a positive trend in quarterly performance, with Q3 showing the strongest results. Year-over-year growth stands at 15%, exceeding industry averages by approximately 3 percentage points. Key drivers include new product launches and expanded market presence in the APAC region.",
      
      "academic": "Analysis of the dataset (n=25) reveals a statistically significant upward trend (p<0.05) across the observed time period. The mean value (μ=42.5) represents a 15% increase from the previous comparable period, with a standard deviation (σ=18.3) indicating moderate variability.",
      
      "simple": "The numbers are going up! They're about 15% higher than last year, which is really good. The summer months (July-September) had the best results, while December was the slowest month."
    };
    
    return {
      success: true,
      narrative: narrativeStyles[style] || narrativeStyles.business,
      wordCount: narrativeStyles[style] ? narrativeStyles[style].split(" ").length : 0,
      confidence: 0.89,
      alternativeStyles: Object.keys(narrativeStyles).filter(s => s !== style)
    };
  }
};

// Mock Insert Narrative/Table API
export const insertionApi = {
  // Generate narrative text based on data
  generateNarrative: async (data, options = {}) => {
    // Reuse the summarize API's narrative generator
    return summarizeApi.generateNarrative(data, options.style);
  },
  
  // Generate a formatted table based on data
  generateTable: async (data, options = {}) => {
    // Simulate network delay
    await new Promise(resolve => setTimeout(resolve, 1500));
    
    // Mock table data
    const mockTableData = {
      "financial": {
        headers: ["Quarter", "Revenue ($M)", "Expenses ($M)", "Profit ($M)", "Margin (%)"],
        rows: [
          ["Q1 2023", 12.4, 9.8, 2.6, "21.0%"],
          ["Q2 2023", 14.2, 10.5, 3.7, "26.1%"],
          ["Q3 2023", 15.8, 11.2, 4.6, "29.1%"],
          ["Q4 2023", 13.5, 10.8, 2.7, "20.0%"]
        ],
        summary: "Annual totals",
        totals: ["2023", 55.9, 42.3, 13.6, "24.3%"]
      },
      
      "project": {
        headers: ["Task", "Owner", "Status", "Due Date", "Completion"],
        rows: [
          ["Research", "John Smith", "Completed", "2023-03-15", "100%"],
          ["Design", "Alice Johnson", "Completed", "2023-04-10", "100%"],
          ["Development", "Team A", "In Progress", "2023-07-30", "65%"],
          ["Testing", "QA Team", "Not Started", "2023-08-15", "0%"],
          ["Deployment", "DevOps", "Not Started", "2023-09-01", "0%"]
        ],
        summary: "Project status",
        totals: ["Overall", "", "In Progress", "", "53%"]
      },
      
      "default": {
        headers: ["Category", "Value 1", "Value 2", "Value 3", "Total"],
        rows: [
          ["A", 10, 15, 20, 45],
          ["B", 12, 18, 22, 52],
          ["C", 8, 12, 16, 36],
          ["D", 14, 21, 28, 63]
        ],
        summary: "Summary",
        totals: ["Total", 44, 66, 86, 196]
      }
    };
    
    // Determine which table to return
    let tableType = "default";
    if (data) {
      if (data.toLowerCase().includes("financ") || data.toLowerCase().includes("revenue") || data.toLowerCase().includes("profit")) {
        tableType = "financial";
      } else if (data.toLowerCase().includes("project") || data.toLowerCase().includes("task") || data.toLowerCase().includes("status")) {
        tableType = "project";
      }
    }
    
    const tableData = mockTableData[tableType];
    
    return {
      success: true,
      tableData: tableData,
      formatting: {
        headerStyle: "bold",
        alternatingRows: true,
        totalsStyle: "bold",
        numberFormat: tableType === "financial" ? "currency" : "general"
      },
      dimensions: {
        rows: tableData.rows.length + 2, // +2 for header and totals
        columns: tableData.headers.length
      }
    };
  },
  
  // Insert text at a specific location
  insertTextAtLocation: async (text, location) => {
    // Simulate network delay
    await new Promise(resolve => setTimeout(resolve, 300));
    
    return {
      success: true,
      insertedText: text,
      location: location,
      timestamp: new Date().toISOString()
    };
  }
};

// Utility function to simulate Excel data retrieval
export const excelDataUtils = {
  // Get data from current selection
  getCurrentSelection: async () => {
    // Simulate network delay
    await new Promise(resolve => setTimeout(resolve, 400));
    
    // Mock selection data
    return {
      address: "B2:F15",
      values: [
        ["January", 10, 15, 25, "Region A"],
        ["February", 12, 18, 30, "Region A"],
        ["March", 15, 20, 35, "Region A"],
        ["April", 18, 22, 40, "Region B"],
        ["May", 22, 25, 47, "Region B"],
        ["June", 25, 28, 53, "Region B"],
        ["July", 30, 32, 62, "Region C"],
        ["August", 32, 35, 67, "Region C"],
        ["September", 28, 30, 58, "Region C"],
        ["October", 25, 28, 53, "Region D"],
        ["November", 20, 24, 44, "Region D"],
        ["December", 15, 20, 35, "Region D"]
      ],
      hasHeaders: true,
      dataType: "timeSeries"
    };
  },
  
  // Get data from current sheet
  getCurrentSheetData: async () => {
    // Simulate network delay
    await new Promise(resolve => setTimeout(resolve, 600));
    
    // Mock sheet data (simplified)
    return {
      name: "Financial Summary",
      usedRange: "A1:H20",
      hasCharts: true,
      hasTables: true,
      hasFormulas: true,
      dataRegions: [
        { address: "A1:E13", type: "table", name: "RevenueData" },
        { address: "G1:H13", type: "calculations" },
        { address: "A15:E20", type: "table", name: "ExpenseData" }
      ],
      sampleData: "Revenue\tQ1\tQ2\tQ3\tQ4\nProduct A\t100\t120\t150\t130\nProduct B\t80\t95\t110\t100\n..."
    };
  },
  
  // Get workbook structure
  getWorkbookStructure: async () => {
    // Simulate network delay
    await new Promise(resolve => setTimeout(resolve, 800));
    
    // Mock workbook structure
    return {
      name: "Financial Analysis 2023.xlsx",
      sheets: [
        { name: "Summary", visible: true, type: "worksheet" },
        { name: "Revenue", visible: true, type: "worksheet" },
        { name: "Expenses", visible: true, type: "worksheet" },
        { name: "Projections", visible: true, type: "worksheet" },
        { name: "Raw Data", visible: false, type: "worksheet" },
        { name: "Chart Data", visible: false, type: "worksheet" }
      ],
      activeSheet: "Summary",
      definedNames: ["AnnualTotal", "QuarterlyAverage", "GrowthRate"],
      tables: ["RevenueTable", "ExpenseTable", "ProjectionTable"],
      charts: 5
    };
  }
};
