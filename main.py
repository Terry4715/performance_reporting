#%%
import pandas as pd
from pptx import Presentation
from pptx.util import Inches

# step 1: Import data from excel
excel_file = 'dummy_data.xlsx'
raw_data = pd.read_excel(excel_file)

# step 2: Transform data - converts the return columns to percentages
return_columns = [col for col in raw_data.columns if 'rtn' in col]  # Identify columns containing return data
raw_data[return_columns] = raw_data[return_columns] * 100  # Convert to percentage
raw_data[return_columns] = raw_data[return_columns].round(2)  # Round to two decimal places
raw_data[return_columns] = raw_data[return_columns].map(lambda x: f"{x:.1f}%")  # Format as string with percentage

#%%

# step 3: Insert data into PowerPoint
prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[5])  # Adding a slide with title and content layout

title = slide.shapes.title
title.text = "Hello, World!"

# Setup table dimensions
rows, cols = raw_data.shape
left = top = Inches(2)
width = Inches(6)
height = Inches(0.8)

# Create table and populate it with data
table = slide.shapes.add_table(rows+1, cols, left, top, width, height).table

# Set column names
for col_index, col_name in enumerate(raw_data.columns):
    table.cell(0, col_index).text = col_name

# Fill table with data
for row_index, row in raw_data.iterrows():
    for col_index, item in enumerate(row):
        table.cell(row_index + 1, col_index).text = str(item)

# Save the presentation
prs.save('test.pptx')