import pandas as pd

# Path to your Excel file
file_path = "Mainframe Tracker.xlsx"

# Function to update the details of a component, but only if the step name matches
def update_component_details(component, step_name, work_pds, cobol_program, cobol_pds):
    # Load the Excel sheet into a DataFrame
    df = pd.read_excel(file_path)

    # Drop rows with missing 'job name' or 'step name' (ignore empty rows)
    df = df.dropna(subset=["job name", "step name"])
    df["job name"] = df["job name"].fillna("").astype(str)  # Convert NaN to empty string
    df["step name"] = df["step name"].fillna("").astype(str)

    component = component.strip().upper()
    step_name = step_name.strip().upper()
    df["job name"] = df["job name"].str.strip().str.upper()
    df["step name"] = df["step name"].str.strip().str.upper()


    print(f"Searching for job name: '{component}' and step name: '{step_name}'")
    print("\nFirst few rows of 'job name' and 'step name':")
    print(df[["job name", "step name"]].head())  # Show first few rows to debug

    # Find the specific row that matches the component and step name
    matching_row = df[(df["job name"] == component) & (df["step name"] == step_name)]

    list_of_val = [matching_row]
    print(type(list_of_val))
    # print(matching_row)
    if matching_row.empty:
        print(f" No matching component and step name found for '{component}' and '{step_name}'!")
        return
    
    # If match found, update the row directly
    df.loc[matching_row.index, "work pds"] = work_pds
    df.loc[matching_row.index, "cobol program"] = cobol_program
    df.loc[matching_row.index, "cobol pds"] = cobol_pds

    # Save the updated DataFrame back to the Excel file
    with pd.ExcelWriter(file_path, engine="openpyxl", mode="w") as writer:
        df.to_excel(writer, index=False)

    print(f"Component '{component}' with Step '{step_name}' updated successfully!")

# Example usage
update_component_details("123dddc", "ccccc", "WorkPDS123", "COBOLProgA", "dea.test.cobol(COBOLProgA)")
