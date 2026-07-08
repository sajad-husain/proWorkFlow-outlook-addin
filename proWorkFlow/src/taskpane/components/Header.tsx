import React from "react";

interface HeaderProps {
  activeRoute: "create" | "edit";
  onRouteChange: (route: "create" | "edit") => void;
}

const Header: React.FC<HeaderProps> = ({ activeRoute, onRouteChange }) => {
  return (
    <div className="header">
      <div className="logo">
        <span>📋 ProWorkflow</span>
      </div>
      <div className="nav-tabs">
        <button
          className={`tab ${activeRoute === "create" ? "active" : ""}`}
          onClick={() => onRouteChange("create")}
        >
          ✏️ New Task
        </button>
        <button
          className={`tab ${activeRoute === "edit" ? "active" : ""}`}
          onClick={() => onRouteChange("edit")}
        >
          📝 Edit Task
        </button>
      </div>
    </div>
  );
};

export default Header;
