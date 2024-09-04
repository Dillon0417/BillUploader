import streamlit as st
import openai, os, base64, pandas as pd, json
from pydantic import BaseModel
from PIL import Image
import openpyxl
from openpyxl import load_workbook

# Constants
openai.api_key = st.secrets["openai"]["api_key"]

# Classes and Functions
class Purchase(BaseModel):
    alcohol: str
    date: str
    store: str
    quantity: int
    price: int

class Bill(BaseModel):
    purchases: list[Purchase]

def append_df_to_excel(df, excel_file_path, sheet_name='Sheet1'):
    """
    Append a DataFrame to an existing Excel file, starting from the first truly empty row.

    Parameters:
    - df: DataFrame to append
    - excel_file_path: Path to the existing Excel file
    - sheet_name: The sheet name where the DataFrame will be appended. Default is 'Sheet1'.
    """
    # Load the existing workbook
    book = load_workbook(excel_file_path)
    
    # Access the sheet, or create it if it does not exist
    if sheet_name in book.sheetnames:
        sheet = book[sheet_name]
        
        # Find the first empty row by checking cell contents
        startrow = 0
        for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, min_col=1, max_col=2):
            if all(cell.value is None for cell in row):
                startrow = row[0].row - 1
                break
        else:
            # If no empty row found, startrow is the max_row
            startrow = sheet.max_row
    else:
        startrow = 0  # Start from the beginning if the sheet does not exist

    with pd.ExcelWriter(excel_file_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False, header=False, 
                    startrow=startrow, startcol=1)

def encode_image(image_path):
    with open(image_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode('utf-8')
  
def parse_purchases_to_dataframe(purchases_json):
    # Load the JSON string into a Python dictionary
    purchases_dict = json.loads(purchases_json)
    
    # Convert the 'purchases' list from the dictionary into a DataFrame
    df = pd.DataFrame(purchases_dict['purchases'])
    
    return df

# Streamlit Logic
def streamlit_app():
    st.title('Bill Uploader')

    # File uploader for selecting an Excel file
    excel_file = st.file_uploader("Upload an Excel file to modify", type='xlsx')
    
    if excel_file:
        # Save the uploaded Excel file temporarily
        excel_file_path = excel_file.name
        with open(excel_file_path, "wb") as f:
            f.write(excel_file.getbuffer())

        uploaded_files = st.file_uploader("Choose images...", type='jpg', accept_multiple_files=True)
        
        if uploaded_files is not None:
            all_data = []
            temp_dir = "temp_uploads"
            os.makedirs(temp_dir, exist_ok=True)
            
            for uploaded_file in uploaded_files:
                # Save uploaded file temporarily
                temp_file_path = os.path.join(temp_dir, uploaded_file.name)

                with open(temp_file_path, "wb") as f:
                    f.write(uploaded_file.read())
                encoded_image = encode_image(temp_file_path)

                try:
                    response = openai.beta.chat.completions.parse(
                        model="gpt-4o-mini",  
                        messages=[
                            {
                                'role': 'user',
                                'content': [
                                    {'type': 'text', 'text': """You are an expert at structured data extraction. From the picture of this bill, get the Alcohol Name, Date Purchased (MM/DD/YYYY), Store Name, Quantity Purchased, and Price per Bottle. Output the data into the given structure."""},
                                    {'type': 'image_url', 'image_url': {'url': f'data:image/jpeg;base64,{encoded_image}'}}
                                ]
                            }
                        ],
                        response_format=Bill
                    )
                    result = response.choices[0].message.content
                    frame = parse_purchases_to_dataframe(result)
                    all_data.append(frame)

                    # Debugging: Print the dataframe for each processed image
                    st.write(f"Extracted data for {uploaded_file.name}:")
                    st.dataframe(frame)

                    # Remove the temporary file after processing
                    os.remove(temp_file_path)

                except Exception as e:
                    st.error(f"Error processing bill text from file {uploaded_file.name}: {e}")
            
            # Combine all extracted data into a single DataFrame
            if all_data:
                combined_frame = pd.concat(all_data, ignore_index=True)
                
                st.write("Here is the extracted data from all uploaded files:")
                st.dataframe(combined_frame)
                
                # User confirmation to append data to Excel
                if st.button("Is the data correct? Click to append to Excel."):
                    append_df_to_excel(combined_frame, excel_file_path)
                    st.success("Data appended to Excel successfully!")

                    # Provide download link
                    with open(excel_file_path, "rb") as f:
                        excel_data = f.read()
                    st.download_button(label="Download modified Excel file", data=excel_data, file_name=excel_file_path)

# Main Execution
if __name__ == "__main__":
    # Run the Streamlit app
    streamlit_app()
