import os
import pandas as pd

student_data = {}

def extract_data_from_folder(folder):
    spreadsheet_files = [os.path.join(folder, file) for file in os.listdir(folder) if file.endswith('.xlsx')]
    
    for file in spreadsheet_files:
        spreadsheet = pd.read_excel(file)
        
        if 'RA:' in spreadsheet.columns and 'Name:' in spreadsheet.columns:
            columns_to_extract = ['RA:', 'Name:', 'Course:', 'Email:']
            present_columns = spreadsheet.columns.tolist()
            
            selected_columns = [column for column in columns_to_extract if column in present_columns]
            
            data = spreadsheet[selected_columns].copy()
            data.dropna(subset=['RA:', 'Name:'], inplace=True)
            
            for _, row in data.iterrows():
                student_id = str(row['RA:']).split('.')[0]
                student_id = student_id.upper().replace('A', '', 1)
                name = row['Name:']
                course = row.get('Course:', '')
                
                if student_id not in student_data:
                    student_data[student_id] = {'RA:': student_id, 'Name:': name, 'Course:': course, 'Days:': 1}
                else:
                    student_data[student_id]['Days:'] += 1

main_directory = input("Enter the path to the main directory: ")

for root, dirs, files in os.walk(main_directory):
    for dir_name in dirs:
        current_folder = os.path.join(root, dir_name)
        extract_data_from_folder(current_folder)

complete_data = pd.DataFrame(student_data.values())

output_folder = os.path.join(main_directory, 'TotalHours')
os.makedirs(output_folder, exist_ok=True)

output_path = os.path.join(output_folder, 'ComplementaryHours.xlsx')
complete_data.to_excel(output_path, index=False)

print(f"Data saved to: {output_path}")