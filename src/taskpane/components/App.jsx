import * as React from "react";
import { useState, useEffect } from "react";
import PropTypes from "prop-types";
import Header from "./Header";
import HeroList from "./HeroList";
import TextInsertion from "./TextInsertion";
import { makeStyles } from "@fluentui/react-components";
import { Ribbon24Regular, LockOpen24Regular, DesignIdeas24Regular } from "@fluentui/react-icons";
import { insertText } from "../taskpane";
import AIAssistance from "./AIAssistance/index";
import Login from "./Login";
import { loginApi } from "../mockApi";
const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
  },
});

const App = (props) => {
  const { title } = props;
  const styles = useStyles();
  const [isAuthenticated, setIsAuthenticated] = useState(false);
  const [isSignUp, setIsSignUp] = useState(false);
  const [currentTaskpane, setCurrentTaskpane] = useState("default");
  const [isLoading, setIsLoading] = useState(true);
  
  // The list items are static and won't change at runtime,
  // so this should be an ordinary const, not a part of state.
  const listItems = [
    {
      icon: <Ribbon24Regular />,
      primaryText: "Achieve more with Office integration",
    },
    {
      icon: <LockOpen24Regular />,
      primaryText: "Unlock features and functionality",
    },
    {
      icon: <DesignIdeas24Regular />,
      primaryText: "Create and visualize like a pro",
    },
  ];

  // Check for authentication and get taskpane ID on component mount
  useEffect(() => {
    const checkAuth = async () => {
      // Check if user is authenticated from localStorage
      const storedAuth = localStorage.getItem("anaUserAuthenticated");
      if (storedAuth === "true") {
        // Validate token with mock API
        const email = localStorage.getItem("anaUserEmail");
        try {
          const result = await loginApi.validateToken("mock-token");
          if (result.valid) {
            setIsAuthenticated(true);
          }
        } catch (error) {
          console.error("Token validation failed:", error);
          localStorage.removeItem("anaUserAuthenticated");
        }
      }
      
      // Get taskpane ID from Office context
      if (Office && Office.context && Office.context.ui) {
        // In a real implementation, we would get the taskpaneId from Office context
        // For mock purposes, we'll extract it from the URL if available
        const urlParams = new URLSearchParams(window.location.search);
        const taskpaneId = urlParams.get("taskpaneId") || "default";
        setCurrentTaskpane(taskpaneId);
      }
      
      setIsLoading(false);
    };
    
    checkAuth();
  }, []);

  // If still loading, show a simple loading state
  if (isLoading) {
    return <div className={styles.root}>Loading...</div>;
  }

  // If not authenticated and not on the login taskpane, redirect to login
  if (!isAuthenticated && currentTaskpane !== "TaskpaneLogin") {
    return <Login setIsAuthenticated={setIsAuthenticated} setIsSignUp={setIsSignUp} />;
  }

  // Render the appropriate component based on the taskpane ID
  const renderTaskpane = () => {
    switch (currentTaskpane) {
      case "TaskpaneLogin":
        return <Login setIsAuthenticated={setIsAuthenticated} setIsSignUp={setIsSignUp} />;
      
      case "TaskpaneDDQ":
        return <AIAssistance initialTask="DDQ" />;
      
      case "TaskpaneExplain":
        return <AIAssistance initialTask="Explain" />;
      
      case "TaskpaneSummarize":
        return <AIAssistance initialTask="Summarize" />;
      
      case "TaskpaneInsert":
        return <AIAssistance initialTask="Insert" />;
      
      case "TaskpaneSettings":
        return (
          <div className={styles.root}>
            <Header logo="assets/logo-filled.png" title="Settings" message="Configure your preferences" />
            <div style={{ padding: "20px" }}>
              <p>Settings functionality would be implemented here.</p>
            </div>
          </div>
        );
      
      default:
        return <AIAssistance />;
    }
  };

  return (
    <div className={styles.root}>
      {renderTaskpane()}
    </div>
  );
};

App.propTypes = {
  title: PropTypes.string,
};

export default App;
