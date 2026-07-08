import React, { useState } from "react";
import Header from "../Header";
import CreateTask from "../CreateTask/CreateTask";
import EditTask from "../EditTask/EditTask";

type Route = "create" | "edit";

const AppRouter: React.FC = () => {
  const [currentRoute, setCurrentRoute] = useState<Route>("create");

  // Route ko tab index mein convert karein
  const getTabIndex = (route: Route): number => {
    return route === "create" ? 0 : 1;
  };

  // Tab index se route mein convert karein
  const getRouteFromTab = (tabIndex: number): Route => {
    return tabIndex === 0 ? "create" : "edit";
  };

  const handleTabChange = (newTabIndex: number) => {
    const newRoute = getRouteFromTab(newTabIndex);
    setCurrentRoute(newRoute);
  };

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
      <Header currentTab={getTabIndex(currentRoute)} onTabChange={handleTabChange} />
      <div style={{ padding: "16px" }}>{renderContent()}</div>
    </>
  );
};

export default AppRouter;
