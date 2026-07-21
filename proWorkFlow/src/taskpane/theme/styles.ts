import { SxProps, Theme } from "@mui/material";
import type { MenuProps } from "@mui/material";

export const smallMenuProps: Partial<MenuProps> = {
  slotProps: {
    paper: {    
        sx: {
            "& .MuiMenuItem-root": {
                fontSize: "0.75rem",
                paddingTop: "4px",
                paddingBottom: "4px",
                minHeight: "auto",
            },
        }
    },
  },
};

export const smallLabelStyles = {
  "& .MuiInputLabel-root": { fontSize: "0.7rem" },
  "& .MuiInputBase-root": { fontSize: "0.75rem" },
  "& .MuiFormHelperText-root": { fontSize: "0.65rem" },
  "& .MuiSelect-select": { fontSize: "0.75rem" },
  "& .MuiOutlinedInput-input": { fontSize: "0.75rem", padding: "8px 12px" },
  "& .MuiInputLabel-shrink": { fontSize: "0.7rem" },
};

