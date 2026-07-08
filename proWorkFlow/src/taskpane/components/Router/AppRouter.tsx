import React, { useState } from "react";
import Header from "../Header";
import CreateTask from "../CreateTask"; // Phase 1 mein banayenge
import EditTask from "../EditTask"; // Phase 3 mein banayenge

type Route = "create" | "edit";

export const AppRouter: React.FC = () => {
  const [currentRoute, setCurrentRoute] = useState<Route>("create");

  const renderContent = () => {
    switch (currentRoute) {
      case "create":
        return <CreateTask />;
      case "edit":
        return <EditTask />;
      default:
        return <CreateTask />;
    }
  };

  return (
    <div className="app-router">
      <Header activeRoute={currentRoute} onRouteChange={setCurrentRoute} />
      <div className="content-area">{renderContent()}</div>
    </div>
  );
};
