import pandas as pd

# Path of your Excel file
file_path = "Mainframe Tracker.xlsx"

# Load Excel into DataFrame
df = pd.read_excel(file_path)

# Define required columns
required_columns = [
    "job name",
    "step name",
    "cobol program",
    "delivery job",

    "cobol pds"
  # treating this as Delivery Job PDS
]


# Validate all required columns exist
for col in required_columns:
    if col not in df.columns:
        raise ValueError(f" Missing column: {col} in Excel file")

# Add result columns if not already present
if "unit test" not in df.columns:
    df["unit test"] = ""
if "Job status" not in df.columns:
    df["Job status"] = ""

    # Ensure columns are object dtype to allow string assignment
    df["unit test"] = df["unit test"].astype("object")
    df["Job status"] = df["Job status"].astype("object")

# Validation logic
for idx, row in df.iterrows():
    component = str(row["job name"]).strip().upper()
    step_name = str(row["step name"]).strip().upper()
   
    delivery_pds = str(row["delivery job"]).strip().upper()   # Delivery Job PDS

    # Condition 1: Work PDS must contain Component AND Step Name


    # Condition 2: Delivery Job PDS must contain Job Name
    cond2 =  component in delivery_pds
   

    if cond2:
        df.at[idx, "unit test"] = "Completed"
        df.at[idx, "Job status"] = "Successful"
    else:
        df.at[idx, "unit test"] = "Pending"
        df.at[idx, "Job status"] = "pending"

# Save back to SAME Excel file (overwrite)
with pd.ExcelWriter(file_path, engine="openpyxl", mode="w") as writer:
    df.to_excel(writer, index=False)

    print("Excel updated with validation results!")
