import React, { useState } from "react";
import Header from "../Header";
import CreateTask from "../CreateTask/CreateTask";
import EditTask from "../EditTask/EditTask";

type Route = "create" | "edit";

const AppRouter: React.FC = () => {
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
    <>
      <Header activeRoute={currentRoute} onRouteChange={setCurrentRoute} />
      <div style={{ padding: "16px" }}>{renderContent()}</div>
    </>
  );
};

export default AppRouter;
