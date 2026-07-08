import * as React from "react";
import AppRouter from "./Router/AppRouter";

const App: React.FC = () => {
  return (
    <div>
      <h2>Apps</h2>
      <AppRouter />
    </div>
  );
};

export default App;
