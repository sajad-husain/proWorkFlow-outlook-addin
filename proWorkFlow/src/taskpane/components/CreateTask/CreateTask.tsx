import React from "react";
import { Box, Typography, Paper } from "@mui/material";

const CreateTask: React.FC = () => {
  return (
    <Box>
      <Typography variant="h6" gutterBottom>
        Create New Task
      </Typography>
      <Typography variant="body2" color="textSecondary">
        Task form will be here...
      </Typography>
    </Box>
  );
};

export default CreateTask;
