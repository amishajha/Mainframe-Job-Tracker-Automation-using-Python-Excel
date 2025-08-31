import pandas as pd

class Tracker:
    def __init__(self, file_path):
        self.file_path = file_path
        self.df = self._load_excel()

    # --- Private Method (internal use only) ---
    def _load_excel(self):
        """Load the Excel file into a DataFrame."""
        return pd.read_excel(self.file_path)

    def _save_excel(self):
        """Save the DataFrame back to the Excel file."""
        with pd.ExcelWriter(self.file_path, engine="openpyxl", mode="w") as writer:
            self.df.to_excel(writer, index=False)

    # --- Public Method: Find a row ---
    def find_row(self, job_name, step_name):
        """Return matching rows for job and step name."""
        return self.df[(self.df["job name"] == job_name) & (self.df["step name"] == step_name)]

    # --- Public Method: Update component ---
    def update_component(self, job_name, step_name, work_pds, cobol_program, cobol_pds):
        """Update details of a component if job and step match."""
        match = self.find_row(job_name, step_name)

        if match.empty:
            print(f" No match found for Job: '{job_name}' and Step: '{step_name}'")
            return

        # Update the matched row(s)
        self.df.loc[(self.df["job name"] == job_name) & (self.df["step name"] == step_name), "work pds"] = work_pds
        self.df.loc[(self.df["job name"] == job_name) & (self.df["step name"] == step_name), "cobol program"] = cobol_program
        self.df.loc[(self.df["job name"] == job_name) & (self.df["step name"] == step_name), "cobol pds"] = cobol_pds

        # Save after update
        self._save_excel()
        print(f"Updated Job '{job_name}' Step '{step_name}' successfully!")

    # --- Extra: Reload Data (if Excel changes outside) ---
    def reload(self):
        """Reload Excel from disk."""
        self.df = self._load_excel()
        print("Excel reloaded.")


# --- Example Usage ---
tracker = Tracker("Mainframe Tracker.xlsx")

# Update component
tracker.update_component("axssz", "qqqq", "WorkPDS123", "COBOLProgA", "dea.test.cobol(COBOLProgA)")

# Optionally reload if Excel was modified outside Python
tracker.reload()
