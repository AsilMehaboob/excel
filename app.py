import os
import xlwings as xw

# Define the folder containing the Excel files
folder_path = r"C:\Users\asilm\OneDrive\Desktop\sm\excel"

# Define the path to the new logo
new_logo_path = r"C:\Users\asilm\Downloads\jpeg.jpg"

# Iterate through each file in the folder
for filename in os.listdir(folder_path):
    if filename.endswith(".xlsx") or filename.endswith(".xlsm"):
        file_path = os.path.join(folder_path, filename)
        
        # Open the workbook with xlwings
        app = xw.App(visible=False)
        workbook = app.books.open(file_path)
        
        # Iterate through each sheet in the workbook
        for sheet in workbook.sheets:
            # Clear existing header
            sheet.api.PageSetup.CenterHeaderPicture.Filename = ''
            sheet.api.PageSetup.LeftHeaderPicture.Filename = ''
            sheet.api.PageSetup.RightHeaderPicture.Filename = ''
            
            # Add the new logo to the header (you can choose left, center, or right)
            sheet.api.PageSetup.CenterHeaderPicture.Filename = new_logo_path
            sheet.api.PageSetup.CenterHeader = "&G"  # &G is a placeholder for the picture
            
        # Save and close the workbook
        workbook.save()
        workbook.close()
        app.quit()

print("Logos in headers replaced successfully.")
