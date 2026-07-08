import React from "react";
import { AppBar, Toolbar, Typography, Tabs, Tab, Box } from "@mui/material";
import { Assignment, EditNote } from "@mui/icons-material";

interface HeaderProps {
  currentTab?: number;
  onTabChange?: (newValue: number) => void;
}

const Header: React.FC<HeaderProps> = ({ currentTab = 0, onTabChange }) => {
  const handleChange = (_event: React.SyntheticEvent, newValue: number) => {
    if (onTabChange) {
      onTabChange(newValue);
    }
  };

  return (
    <AppBar position="static" elevation={0} sx={{ bgcolor: "primary.main" }}>
      <Toolbar sx={{ flexDirection: "column", alignItems: "stretch", p: 0 }}>
        <Box sx={{ display: "flex", alignItems: "center", px: 2, py: 1 }}>
          <Typography variant="h6" sx={{ flex: 1, color: "white" }}>
            ProWorkflow
          </Typography>
          <Typography variant="caption" sx={{ color: "rgba(255,255,255,0.7)" }}>
            Outlook Add-in
          </Typography>
        </Box>
        <Tabs
          value={currentTab}
          onChange={handleChange}
          sx={{
            bgcolor: "rgba(255,255,255,0.05)",
            "& .MuiTab-root": {
              color: "rgba(255,255,255,0.7)",
              "&.Mui-selected": {
                color: "white",
              },
            },
            "& .MuiTabs-indicator": {
              bgcolor: "white",
            },
          }}
        >
          <Tab
            icon={<Assignment />}
            iconPosition="start"
            label="New Task"
            sx={{ textTransform: "none" }}
          />
          <Tab
            icon={<EditNote />}
            iconPosition="start"
            label="Edit Task"
            sx={{ textTransform: "none" }}
          />
        </Tabs>
      </Toolbar>
    </AppBar>
  );
};

export default Header;
