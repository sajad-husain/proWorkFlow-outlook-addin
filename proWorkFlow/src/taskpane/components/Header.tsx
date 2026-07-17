import React from "react";
import { AppBar, Toolbar, Typography, Tabs, Tab, Box } from "@mui/material";
import { Assignment, EditNote } from "@mui/icons-material";

interface HeaderProps {
  activeRoute?: "create" | "edit";
  onRouteChange?: (route: "create" | "edit") => void;
}

const Header: React.FC<HeaderProps> = ({ activeRoute = "create", onRouteChange }) => {
  const getTabIndex = (route: "create" | "edit"): number => {
    return route === "create" ? 0 : 1;
  };

  const handleChange = (_event: React.SyntheticEvent, newValue: number) => {
    if (onRouteChange) {
      const newRoute = newValue === 0 ? "create" : "edit";
      onRouteChange(newRoute);
    }
  };

  return (
    <AppBar position="static" elevation={0} sx={{ bgcolor: "primary.main" }}>
      <Toolbar
        sx={{
          flexDirection: "column",
          alignItems: "stretch",
          p: 0,
          minHeight: "48px", // default is 64px, we reduce
        }}
      >
        <Box
          sx={{
            display: "flex",
            alignItems: "center",
            px: 1.5,
            py: 0.5, // less padding
          }}
        >
          <Typography
            variant="subtitle1" // smaller variant
            sx={{ flex: 1, color: "white", fontWeight: 600, fontSize: "0.9rem" }}
          >
            ProWorkflow
          </Typography>
        </Box>

        <Tabs
          value={getTabIndex(activeRoute)}
          onChange={handleChange}
          sx={{
            bgcolor: "rgba(255,255,255,0.05)",
            minHeight: "36px", // default 48px, reduced
            "& .MuiTab-root": {
              color: "rgba(255,255,255,0.7)",
              minHeight: "36px",
              padding: "4px 12px", // reduced padding
              fontSize: "0.75rem", // smaller font
              textTransform: "none",
              "&.Mui-selected": {
                color: "white",
              },
            },
            "& .MuiTabs-indicator": {
              bgcolor: "white",
              height: "2px",
            },
          }}
        >
          <Tab
            icon={<Assignment sx={{ fontSize: "1rem" }} />}
            iconPosition="start"
            label="New Task"
            sx={{ minHeight: "36px" }}
          />
          <Tab
            icon={<EditNote sx={{ fontSize: "1rem" }} />}
            iconPosition="start"
            label="Edit Task"
            sx={{ minHeight: "36px" }}
          />
        </Tabs>
      </Toolbar>
    </AppBar>
  );
};

export default Header;
