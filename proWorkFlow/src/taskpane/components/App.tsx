import * as React from "react";
import { makeStyles, tokens } from "@fluentui/react-components";
import Header from "./Header";
import CreateTaskForm from "./CreateTask/CreateTaskForm";
import { useEmailContext } from "../hooks/useEmailContext";

export interface AppProps {
  title: string;
}

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
    display: "flex",
    flexDirection: "column",
    backgroundColor: tokens.colorNeutralBackground1,
  },
});

const App: React.FC<AppProps> = (props: AppProps) => {
  const styles = useStyles();
  const { emailData, loading } = useEmailContext();

  return (
    <div className={styles.root}>
      <Header logo="assets/logo-filled.png" title={props.title} />
      <CreateTaskForm emailData={emailData} loading={loading} />
    </div>
  );
};

export default App;
