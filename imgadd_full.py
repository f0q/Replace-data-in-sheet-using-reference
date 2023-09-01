import os
import fnmatch
import pandas as pd
from fuzzywuzzy import fuzz

#данный код добавляет в новый столбец адреса файлов, которые нашлись по идентификатору в столбце identificator_bbq в целевой таблице. Для чего? Чтобы потом использовать эти файлы в выгрузке в авито

# Replace with your target directory path
target_directory = 'images'
table_filename = 'products.xlsx'  # Replace with your table's filename

# Load the table into a pandas DataFrame
table_df = pd.read_excel(table_filename)

# Create a new column for the image filenames
table_df['Filenames'] = ''

# Iterate through rows in the DataFrame
for index, row in table_df.iterrows():
    specific_id_column = 'identificator_bbq'  # Replace with the column name for specific ID
    specific_word = 'размеры'  # Replace with your specific word

    specific_id = row[specific_id_column]
    
    # Search for folders with similar IDs
    matching_folders = [folder for folder in os.listdir(target_directory) if str(specific_id) in folder]
    
    best_matching_folder = None
    best_matching_score = 0
    
    for folder in matching_folders:
        matching_score = fuzz.ratio(str(specific_id), folder)
        if matching_score > best_matching_score:
            best_matching_folder = folder
            best_matching_score = matching_score
    
    if best_matching_folder:
        folder_path = os.path.join(target_directory, best_matching_folder)
        
        # Search for the JPG files with the specific word
        matching_files = [filename for filename in os.listdir(folder_path) if
                          fnmatch.fnmatch(filename, f'*{specific_word}*.jpg')]
        
        if matching_files:
            # Combine the matching filenames with full paths using the | delimiter
            full_filenames = [os.path.join(folder_path, filename) for filename in matching_files]
            combined_full_filenames = '|'.join(full_filenames)
            table_df.at[index, 'Filenames'] = combined_full_filenames
        else:
            table_df.at[index, 'Filenames'] = ''  # No matching images

# Save the updated DataFrame to a new Excel file
output_filename = 'updated_products_table.xlsx'
table_df.to_excel(output_filename, index=False)