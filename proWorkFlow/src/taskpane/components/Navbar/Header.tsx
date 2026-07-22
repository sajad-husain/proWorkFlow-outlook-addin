import React from "react";
import { AppBar, Toolbar, Typography, Tabs, Tab, Box, SxProps, Theme } from "@mui/material";
import { Assignment, EditNote } from "@mui/icons-material";

interface HeaderProps {
  activeRoute?: "create" | "edit";
  onRouteChange?: (route: "create" | "edit") => void;
  // Optional: allow overriding styles
  sx?: SxProps<Theme>;
}

const Header: React.FC<HeaderProps> = ({ activeRoute = "create", onRouteChange, sx }) => {
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
    <AppBar
      position="static"
      elevation={0}
      sx={{
        bgcolor: "white", // 🔥 white background
        borderBottom: "1px solid #e0e0e0", // subtle separator
        ...sx, // allow external overrides
      }}
    >
      <Toolbar
        sx={{
          flexDirection: "column",
          alignItems: "stretch",
          p: 0,
          minHeight: "48px",
        }}
      >
        {/* Title row */}
        <Box
          sx={{
            display: "flex",
            alignItems: "center",
            px: 1.5,
            py: 0.5,
          }}
        >
          <Typography
            variant="subtitle1"
            sx={{
              flex: 1,
              color: "primary.main", // 🔥 primary color text
              fontWeight: 600,
              fontSize: "0.9rem",
            }}
          >
            ProWorkflow
          </Typography>
        </Box>

        {/* Tabs */}
        <Tabs
          value={getTabIndex(activeRoute)}
          onChange={handleChange}
          sx={{
            bgcolor: "transparent",
            minHeight: "36px",
            "& .MuiTab-root": {
              color: "text.secondary", // 🔥 unselected – grey
              minHeight: "36px",
              padding: "4px 12px",
              fontSize: "0.75rem",
              textTransform: "none",
              "&.Mui-selected": {
                color: "primary.main", // 🔥 selected – primary
              },
              "&:hover": {
                color: "primary.main",
                bgcolor: "rgba(0, 0, 0, 0.04)",
              },
            },
            "& .MuiTabs-indicator": {
              bgcolor: "primary.main", // 🔥 indicator – primary
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
