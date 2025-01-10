import pandas as pd
from Data.outputData import OutputData


def load_excel(file_path):
    """Load Excel file and return a list of OutputData objects."""
    df = pd.read_excel(file_path)
    data_list = []

    for index, row in df.iterrows():
        # Create OutputData object from each row in the DataFrame
        output_data = OutputData(
            date_acquired=row['Date'],
            mobile_phone='0' + str(row['Mobile'])[2:],
            landline=row['Landline'],
            name=f"{row['First name']} {row['Last name']}"  # Combine first and last name
        )
        data_list.append(output_data)

    return data_list
