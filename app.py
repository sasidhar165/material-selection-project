from flask import Flask, render_template, request, redirect, url_for
import pandas as pd
import openpyxl as xl
import math
app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

# Function to normalize a value
def normalize_value(value, min_val, max_val):
    return (value - min_val) / (max_val - min_val)

# Function to rank materials
def rank_materials(materials_data):
    # Normalize Material Index
    min_index = materials_data['Material Index'].min()
    max_index = materials_data['Material Index'].max()
    materials_data['Normalized_Index'] = materials_data['Material Index'].apply(lambda x: normalize_value(x, min_index, max_index))

    # Normalize CO2 Equivalent
    min_co2 = materials_data['CO2 Equivalent'].min()
    max_co2 = materials_data['CO2 Equivalent'].max()
    materials_data['Normalized_CO2'] = materials_data['CO2 Equivalent'].apply(lambda x: normalize_value(x, min_co2, max_co2))

    # Normalize Cost
    min_cost = materials_data['Cost'].min()
    max_cost = materials_data['Cost'].max()
    materials_data['Normalized_Cost'] = materials_data['Cost'].apply(lambda x: normalize_value(x, min_cost, max_cost))

    # Calculate Overall Score
    materials_data['Overall_Score'] = 0.8 * materials_data['Normalized_Index'] + 0.1 * materials_data['Normalized_CO2'] + 0.1 * materials_data['Normalized_Cost']

    # Rank materials
    materials_data = materials_data.sort_values(by='Overall_Score', ascending=False)
    materials_data['Rank'] = range(1, len(materials_data) + 1)

    return materials_data

@app.route('/material_selection', methods=['GET', 'POST'])
def material_selection():
    if request.method == 'POST':
        sheet_id='1iyPgU4wUjFJchws91LxqKf9jTVDd9F6K'
        # Read data from URL
        sheet=pd.read_excel(f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx",sheet_name ='Materials_data')
        materials_data = []
        for i in range(0, 62):
            material = sheet.iloc[i,2]
            material_index = float(sheet.iloc[i,19])
            co2_equivalent = float(sheet.iloc[i,12])
            cost = float(sheet.iloc[i,18])
            materials_data.append({'Material': material, 'Material Index': material_index, 'CO2 Equivalent': co2_equivalent, 'Cost': cost})

        # Create DataFrame
        materials_df = pd.DataFrame(materials_data)

        # Get input values
        material_index_limit = float(request.form['material_index_limit'])
        co2_limit = float(request.form['co2_limit'])
        cost_limit = float(request.form['cost_limit'])

        # Filter materials based on input limits
        filtered_materials = materials_df[
            (materials_df['Material Index'] > material_index_limit) &
            (materials_df['CO2 Equivalent'] < co2_limit) &
            (materials_df['Cost'] < cost_limit)
        ]

        # Rank the filtered materials
        ranked_materials = rank_materials(filtered_materials)

        # Get selected materials
        selected_materials = ranked_materials['Material'].tolist()

        return render_template('material_result.html', selected_materials=selected_materials)

    return render_template('material_selection.html')
@app.route('/process_selection', methods=['GET', 'POST'])
def process_selection():
    if request.method == 'POST':
        sheet_id='1EUxYMwoFsIkGD690Xe32ZAU8ik1Y5XhS'
        materials = pd.read_excel(f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx",sheet_name ='Materials_data')
        materials_selection = pd.read_excel(f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx",sheet_name ='Process_compatibility_matrix')
        section_thickness_sheet = pd.read_excel(f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx",sheet_name ='Section_Thickness(mm)')
        mass_sheet = pd.read_excel(f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx",sheet_name ='Mass(kg)')
        tolerance_sheet = pd.read_excel(f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx",sheet_name ='Tolerance(mm)')
        roughness_sheet = pd.read_excel(f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx",sheet_name ='Roughness(µm)')
        batch_size_sheet = pd.read_excel(f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx",sheet_name ='Economic_batch_size(units)')

        material_name = request.form['material_name']
        section_thickness = float(request.form['section_thickness'])
        mass = float(request.form['mass'])
        tolerance = float(request.form['tolerance'])
        roughness = float(request.form['roughness'])
        batch_size = float(request.form['batch_size'])

        material_row = materials[materials['Materials'] == material_name]
        if not material_row.empty:
            material_type = material_row['Sub-Class'].iloc[0]
        else:
            return "No material of given Type. Run again and enter another material"

        selected_shaping_material = materials_selection[(materials_selection[material_type] == 1) & (materials_selection['Process_Type'] == 'Shaping')]
        selected_finishing_material = materials_selection[(materials_selection[material_type] == 1) & (materials_selection['Process_Type'] == 'Finishing')]
        selected_joining_material = materials_selection[(materials_selection[material_type] == 1) & (materials_selection['Process_Type'] == 'Joining')]

        selected_section_thickness = section_thickness_sheet[(section_thickness_sheet['Lower limit'] <= section_thickness) & (section_thickness_sheet["Upper limit"] >= section_thickness)]
        selected_shaping_mass = mass_sheet[(mass_sheet['Lower limit'] <= mass) & (mass_sheet['Upper limit'] >= mass) & (mass_sheet['Process_Type'] == 'Shaping')]
        selected_joining_mass = mass_sheet[(mass_sheet['Lower limit'] <= mass) & (mass_sheet['Upper limit'] >= mass) & (mass_sheet['Process_Type'] == 'Joining')]
        selected_shaping_tolerance = tolerance_sheet[(tolerance_sheet['Lower limit'] <= tolerance) & (tolerance_sheet['Upper limit'] >= tolerance) & (tolerance_sheet['Process_Type'] == 'Shaping')]
        selected_finishing_tolerance = tolerance_sheet[(tolerance_sheet['Lower limit'] <= tolerance) & (tolerance_sheet['Upper limit'] >= tolerance) & (tolerance_sheet['Process_Type'] == 'Finishing')]
        selected_shaping_roughness = roughness_sheet[(roughness_sheet['Lower limit'] <= roughness) & (roughness_sheet['Upper limit'] >= roughness) & (roughness_sheet["Process_Type"] == 'Shaping')]
        selected_finishing_roughness = roughness_sheet[(roughness_sheet['Lower limit'] <= roughness) & (roughness_sheet['Upper limit'] >= roughness) & (roughness_sheet["Process_Type"] == 'Finishing')]
        selected_batch_size = batch_size_sheet[(batch_size_sheet['Lower limit'] <= batch_size) & (batch_size_sheet['Upper limit'] >= batch_size)]

        selected_shaping_process = pd.concat([selected_shaping_material, selected_section_thickness, selected_shaping_mass, selected_shaping_roughness, selected_shaping_tolerance, selected_batch_size], axis=1, join='inner', keys=['', 'section', 'mass', 'roughness', 'tolerance', 'batch_size'])
        selected_shaping_process.reset_index(drop=True, inplace=True)

        selected_finishing_process = pd.concat([selected_finishing_material, selected_finishing_roughness, selected_finishing_tolerance], axis=1, join='inner', keys=['', 'roughness', 'tolerance'])
        selected_finishing_process.reset_index(drop=True, inplace=True)

        selected_joining_process = pd.concat([selected_joining_material, selected_joining_mass], axis=1, join='inner', keys=['', 'mass'])
        selected_joining_process.reset_index(drop=True, inplace=True)
        
        return render_template('process_result.html',selected_shaping_process=selected_shaping_process, selected_finishing_process=selected_finishing_process, selected_joining_process=selected_joining_process)

    return render_template('process_selection.html')

@app.route('/process_details/<process_name>')
def process_details(process_name):
    sheet_id='1EUxYMwoFsIkGD690Xe32ZAU8ik1Y5XhS'
    # Assuming you have a DataFrame containing process details
    process_details_df = pd.read_excel(f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx",sheet_name ='details_sheet')

    # Find the details for the selected process
    selected_process = process_details_df[process_details_df['process_name'] == process_name]

    if selected_process.empty:
        # Handle the case where the process is not found
        return render_template('process_not_found.html', process_name=process_name)
    else:
        # Extract the definition and comments for the selected process
        definition = selected_process['definition'].iloc[0]
        comments = selected_process['comments'].iloc[0]

        # Render a template to display the details
        return render_template('process_details.html', process_name=process_name, definition=definition, comments=comments)

