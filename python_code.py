#!/usr/bin/env python
# coding: utf-8

# In[1]:


pip install PyMuPDF


# In[2]:


import fitz  # PyMuPDF

input_pdf = "visitors.pdf"
output_pdf = "visitors_new.pdf"

doc = fitz.open(input_pdf)

# Reduce image quality for each page
for page in doc:
    pix = page.get_pixmap(matrix=fitz.Matrix(1, 1), alpha=False)
    page.clean_contents()  # Clean previous high-res images
    page.insert_image(page.rect, pixmap=pix)

doc.save(output_pdf)
doc.close()

print(" Compressed PDF saved as visitors_compressed.pdf")


# In[3]:


pip install pdf2image


# In[1]:


from pdf2image import convert_from_path

pdf_path = "visitors.pdf"
images = convert_from_path(pdf_path)

# Save only the first page to test
images[0].save("page1.jpg", "JPEG")
print(" First page saved as image.")


# In[2]:


from azure.ai.formrecognizer import DocumentAnalysisClient
from azure.core.credentials import AzureKeyCredential

endpoint = "https://visitors500description.cognitiveservices.azure.com/"
key = "ooooooo"

client = DocumentAnalysisClient(endpoint, AzureKeyCredential(key))

with open("page1.jpg", "rb") as f:
    poller = client.begin_analyze_document("prebuilt-document", document=f)
    result = poller.result()

print("Azure processed the image successfully!")


# In[3]:


import pandas as pd

# Create a list to hold all tables
all_tables = []

# Loop through each table in the result
for i, table in enumerate(result.tables):
    data = []

    # Extract each row of the table
    for row_idx in range(table.row_count):
        row_data = []
        for col_idx in range(table.column_count):
            # Find the cell at this row/column
            cell = next(
                (c for c in table.cells if c.row_index == row_idx and c.column_index == col_idx),
                None
            )
            row_data.append(cell.content if cell else "")
        data.append(row_data)

    # Convert the current table to DataFrame
    df = pd.DataFrame(data)
    all_tables.append((f"Table_{i+1}", df))

# Save to Excel file
with pd.ExcelWriter("output_tables.xlsx") as writer:
    for sheet_name, df in all_tables:
        df.to_excel(writer, sheet_name=sheet_name, index=False)

print(" Tables extracted and saved to output_tables.xlsx!")


# In[4]:


# Display the first table
all_tables[0][1]


# In[5]:


from pdf2image import convert_from_path
import os
import pandas as pd
from azure.ai.formrecognizer import DocumentAnalysisClient
from azure.core.credentials import AzureKeyCredential
from PIL import Image
from io import BytesIO

# Azure credentials
endpoint = "https://visitors500description.cognitiveservices.azure.com/"
key = "ooooooooooo"
client = DocumentAnalysisClient(endpoint=endpoint, credential=AzureKeyCredential(key))

# Convert full PDF to images
pdf_path = "visitors.pdf"
images = convert_from_path(pdf_path)

# List to hold all tables
combined_data = []

# Process each page
for page_num, image in enumerate(images, start=1):
    # Convert image to bytes
    img_byte_arr = BytesIO()
    image.save(img_byte_arr, format="PNG")
    img_byte_arr.seek(0)

    # Analyze the image
    poller = client.begin_analyze_document("prebuilt-layout", document=img_byte_arr)
    result = poller.result()

    for i, table in enumerate(result.tables):
        data = []

        for row_idx in range(table.row_count):
            row_data = []
            for col_idx in range(table.column_count):
                cell = next(
                    (c for c in table.cells if c.row_index == row_idx and c.column_index == col_idx),
                    None
                )
                row_data.append(cell.content if cell else "")
            data.append(row_data)

        df = pd.DataFrame(data)
        df.insert(0, "Page", page_num)
        df.insert(1, "Table", i + 1)
        combined_data.append(df)

# Combine all into one DataFrame
final_df = pd.concat(combined_data, ignore_index=True)

# Save to a single Excel sheet
final_df.to_excel("Combined_Tables_From_PDF.xlsx", index=False)

print(" All tables combined and saved to Combined_Tables_From_PDF.xlsx")


# In[6]:


# Import the required library
import pandas as pd

# Load the saved Excel file
df_loaded = pd.read_excel("Combined_Tables_From_PDF.xlsx")

# Display the first few rows
df_loaded.head()  # You can also use df_loaded.tail() to see the bottom rows


# In[ ]:




