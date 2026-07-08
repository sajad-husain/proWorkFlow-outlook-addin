import * as React from "react";
import { makeStyles, tokens } from "@fluentui/react-components";
import Header from "./Header";
import CreateTaskForm from "./CreateTask/CreateTaskForm";
import { useEmailContext } from "../hooks/useEmailContext";

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
    display: "flex",
    flexDirection: "column",
    backgroundColor: tokens.colorNeutralBackground1,
  },
});

const App: React.FC = () => {
  const styles = useStyles();
  const { emailData, loading } = useEmailContext();

  return (
    <div className={styles.root}>
      <Header logo="assets/logo-filled.png" />
      <CreateTaskForm emailData={emailData} loading={loading} />
    </div>
  );
};

export default App;
