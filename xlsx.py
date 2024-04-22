import pandas as pd
import re

# Read the input Excel file
df = pd.read_excel('C:/Users/user/Downloads/mouneshkp.xlsx')
df['BillingDate'] = df['BillingDate'].astype(str)

# Split rows based on multiple values in the 'Airline Name/Rail/Bus/Hotel/Mis' column
df['Airline Name/Rail/Bus/Hotel/Mis'] = df['Airline Name/Rail/Bus/Hotel/Mis'].apply(lambda x: x.split('|') if isinstance(x, str) else [x])
df['Air Line Pnr'] = df['Air Line Pnr'].apply(lambda x: x.split('|') if isinstance(x, str) else [x])
df['Booking Class'] = df['Booking Class'].apply(lambda x: x.split(',') if isinstance(x, str) else [x])

expanded_rows = []

for index, row in df.iterrows():
    # Calculate the number of airlines in the current row
    num_airlines = len(row['Airline Name/Rail/Bus/Hotel/Mis'])
    
    # Divide 'Transaction_Amount_INR (Basic)' and 'K3' columns by the number of airlines
    amount_per_airline = row['Transaction_Amount_INR (Basic)'] / num_airlines
    k3_per_airline = row['K3'] / num_airlines
    
    # Create a new row for each airline
    for i in range(num_airlines):
        new_row = row.copy()  # Create a copy of the current row
        new_row['Airline Name/Rail/Bus/Hotel/Mis'] = row['Airline Name/Rail/Bus/Hotel/Mis'][i]
        new_row['Air Line Pnr'] = row['Air Line Pnr'][i] if i < len(row['Air Line Pnr']) else ''
        new_row['Booking Class'] = row['Booking Class'][i] if i < len(row['Booking Class']) else ''
        new_row['Transaction_Amount_INR (Basic)'] = amount_per_airline  # Update amount
        new_row['K3'] = k3_per_airline  # Update K3
        expanded_rows.append(new_row)

# Create the final DataFrame with expanded rows
output_df = pd.DataFrame(expanded_rows)

# Rename the columns
output_df = output_df.rename(columns={
    'Airline Name/Rail/Bus/Hotel/Mis': 'Vendor',
    'BillingDate': 'Transcation_date',
    'TicketNo': 'Ticket_Number',
    'Air Line Pnr': 'PNR',
    'Booking Class': 'Class',
    'Sector': 'Location(sector)',
    'Corporate GST_Number': 'Customer_GSTIN',
    'Invoice Number /Credit Note Number -Emt Number': 'AGENCY_INVOICE',
    'Pax Name': 'TRAVELLER_NAME'
})

# Add the 'Sr.' column based on the index
output_df['Sr.'] = output_df.index + 1

for i in output_df['Transaction_Amount_INR (Basic)']:
    if i<0:
        output_df['Transcation_type']="REFUND"
    elif i>=0:
        output_df['Transcation_type']="INVOICE"

# Add the 'Transaction_type' and 'Customer_Name' columns
 # Fill with appropriate values
output_df['Customer_Name'] = output_df['Company Name']
output_df['WORKSPACE_NAME'] = output_df['Company Name']
output_df['TRAVELLER_NAME'] = output_df['TRAVELLER_NAME'].replace({'MR': '', 'MS': ''})
def remove_prefix_suffix(name):
    # Define a regular expression pattern to match prefixes and suffixes
    pattern = r'^MR |^MS |, Jr\.|, Sr\.|, III$'

    # Use re.sub() to replace the matched pattern with an empty string
    cleaned_name = re.sub(pattern, '', name)

    return cleaned_name

# Apply the remove_prefix_suffix function to each element in 'TRAVELLER_NAME'
output_df['TRAVELLER_NAME'] = output_df['TRAVELLER_NAME'].apply(remove_prefix_suffix)



output_df['AGENCY_Name'] = 'EASE MY TRIP'
output_df['Domestic/International'] = output_df['Domestic/International'].replace({'Dom': 'Domestic', 'Intl': 'International'})


 # Fill with appropriate values

# Save the output to a new Excel file
output_df.to_excel('output.xlsx', index=False)
