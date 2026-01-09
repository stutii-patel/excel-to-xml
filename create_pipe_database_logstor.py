import xml.etree.ElementTree as ET
import pandas as pd
import matplotlib as mpl
import matplotlib.pyplot as pl
import os
import re

# File paths
input_files = [
    {'path': 'data/VICUS_DB_LOGSTOR_Rohre.xlsx', 'out': 'data/u_wert_logstor_sheet2.xml'}
]
db_xml_file = 'data/db_pipes_original.xml'
updated_db_file = 'data/db_pipes.xml'

# Helper to format numeric values to match original XML (no .0 for integers)
def fmt_val(val):
    if pd.isna(val):
        return ""
    try:
        f_val = float(val)
        if f_val.is_integer():
            return str(int(f_val))
        return str(f_val)
    except:
        return str(val)

# Function to get the last ID from the existing XML (robustly)
def get_last_id(xml_path):
    if not os.path.exists(xml_path):
        return 1100000
    try:
        with open(xml_path, 'r', encoding='utf-8') as f:
            content = f.read()
        ids = re.findall(r'id="(\d+)"', content)
        if ids:
            return int(ids[-1])
    except Exception as e:
        print(f"Error finding last ID in {xml_path}: {e}")
    return 1100000

# Get the initial last ID
current_id_counter = get_last_id(db_xml_file)
print(f"Initial ID: {current_id_counter}")

# We will collect all new chunks to append to the final DB
all_new_pipes_chunk = ""

for file_info in input_files:
    excel_path = file_info['path']
    individual_out = file_info['out']
    
    print(f"Processing {excel_path}...")
    
    # Load the Excel file and filter for "isoflex"
    df = pd.read_excel(excel_path, sheet_name='Einzel- o Doppelrohr mit U-Wert', skiprows=1)
    df = df.dropna(subset=['Produkt', 'Außendurchmesser [mm]'], how='all')
    
    # Color mapping setup
    cmap = pl.get_cmap('turbo')
    diam_col = 'Außendurchmesser [mm]'

    # Calculate total_number per diameter block for this file
    total_numbers = [None] * len(df)
    if len(df) > 0:
        start_idx = 0
        previous_diameter = df.loc[0, diam_col]
        for i in range(1, len(df)):
            current_diameter = df.loc[i, diam_col]
            if current_diameter < previous_diameter:
                block_len = i - start_idx
                total_numbers[start_idx:i] = [block_len] * block_len
                start_idx = i
            previous_diameter = current_diameter
        block_len = len(df) - start_idx
        total_numbers[start_idx:] = [block_len] * block_len
    df['total_number'] = total_numbers

    # Generate entries for this file
    new_pipes_list = []
    prev_da = -1
    color_count = 0

    for index, row in df.iterrows():
        da = row[diam_col]
        s = row['Wandstärke [mm]']
        roughness = row.get('Rohrrauigkeit [mm]')
        product_name = str(row.get('Produkt', ''))
        manufacturer_raw = str(row.get('Hersteller', 'LOGSTOR'))
        manufacturer = manufacturer_raw.split('-')[0].upper()
        density = row.get('Dichte Rohrwand [kg/m3]')
        cp = row.get('Wärmekapazität Rohrwand [W/mK]')

        # Material detection (Force Steel for isoflex)
        material_wall_val = str(row.get('Material Rohrwand', '')).lower()
        is_plastic = "kunststoff" in material_wall_val
        if "isoflex" in product_name.lower():
            is_plastic = False


        if is_plastic:
            lambda_wall = 0.4
            material_standard = "PlasticPipe"
            cat_name = "DE: PE isoliert | EN: PE insulated"
        else:
            lambda_wall = 50
            material_standard = "EnStandard"
            cat_name = "DE: Stahl KMR | EN: Steel bonded pipe"

        total_outer_diameter = row.get('Außendurchmesser gesamt mit Isolierung und Schutzschicht [mm]')
        layout_type = str(row.get('Einzel- oder Doppelrohr', ''))
        spacing = row.get('Abstand Vor- und Rücklauf [mm]')
        UValue = row.get('U-Wert [W/mK]')
        pn_value = row.get('PN [bar]')

        if da < prev_da:
            color_count = 0
        prev_da = da
        color = mpl.colors.rgb2hex(cmap(color_count / row['total_number']), keep_alpha=False)
        color_count += 1

        # Increment ID across all files
        current_id_counter += 1
        # Order: id, color, categoryName, productName, manufacturerName
        pipe = ET.Element("NetworkPipe", id=str(current_id_counter), color=color, categoryName=cat_name)
        pipe.set("productName", product_name)
        pipe.set("manufacturerName", manufacturer)

        def add_ibk_param(parent, param_name, value, unit):
            if pd.notna(value):
                node = ET.SubElement(parent, "IBK:Parameter")
                node.set("name", param_name)
                node.set("unit", unit)
                node.text = fmt_val(value)

        add_ibk_param(pipe, "DiameterOutside", da, "mm")
        add_ibk_param(pipe, "ThicknessWall", s, "mm")
        add_ibk_param(pipe, "RoughnessWall", roughness, "mm")
        # add_ibk_param(pipe, "ThermalConductivityWall", lambda_wall, "W/mK")
        add_ibk_param(pipe, "HeatCapacityWall", cp, "J/kgK")
        add_ibk_param(pipe, "DensityWall", density, "kg/m3")
        add_ibk_param(pipe, "FixedUValue", UValue, "W/mK")
        add_ibk_param(pipe, "FixedTotalOuterDiameter", total_outer_diameter, "mm")
        add_ibk_param(pipe, "PipeSpacing", spacing, "mm")

        if pd.notna(pn_value):
            ET.SubElement(pipe, "NominalPressure").text = fmt_val(pn_value)
        ET.SubElement(pipe, "FixedUValueGiven").text = "true"
        
        if 'Einzelrohr' in layout_type:
            ET.SubElement(pipe, "PipeLayout").text = "SinglePipe"
        else:
            ET.SubElement(pipe, "PipeLayout").text = "TwinPipe"

        ET.SubElement(pipe, "PipeMaterialStandard").text = material_standard
        new_pipes_list.append(pipe)

    def prettify_str(element, level=1):
        indent = "\t" * level
        sub_indent = "\t" * (level + 1)
        # Use list of keys for attribute ordering
        attr_keys = ['id', 'color', 'categoryName', 'productName', 'manufacturerName']
        attrs_list = []
        for k in attr_keys:
            if k in element.attrib:
                attrs_list.append(f'{k}="{element.attrib[k]}"')
        attrs = " ".join(attrs_list)
        xml_str = f"{indent}<NetworkPipe {attrs}>\n"
        for sub in element:
            if sub.tag == "IBK:Parameter":
                # Ensure name then unit
                p_name = sub.attrib.get('name', '')
                p_unit = sub.attrib.get('unit', '')
                xml_str += f'{sub_indent}<IBK:Parameter name="{p_name}" unit="{p_unit}">{sub.text}</IBK:Parameter>\n'
            else:
                xml_str += f"{sub_indent}<{sub.tag}>{sub.text}</{sub.tag}>\n"
        xml_str += f"{indent}</NetworkPipe>\n"
        return xml_str

    # String chunk for this file
    file_pipes_chunk = ""
    for pipe in new_pipes_list:
        file_pipes_chunk += prettify_str(pipe)
    
    # Save individual file
    file_xml_content = '<?xml version="1.0" encoding="UTF-8" ?>\n<NetworkPipes>\n' + file_pipes_chunk + '</NetworkPipes>\n'
    with open(individual_out, 'w', encoding='utf-8') as f:
        f.write(file_xml_content)
    print(f"Created {individual_out}")
    
    # Accumulate for global merge
    all_new_pipes_chunk += file_pipes_chunk

# Total entries added count
total_added = len(re.findall(r'<NetworkPipe\b', all_new_pipes_chunk))
print(f"\nTotal entries added: {total_added}")

# Final Merge and Summary
summary = []

if os.path.exists(db_xml_file):
    with open(db_xml_file, 'r', encoding='utf-8') as f:
        db_content = f.read()
    
    # Count original entries - Use \s or [^s] to avoid matching <NetworkPipes>
    original_count = len(re.findall(r'<NetworkPipe\b', db_content))
    summary.append(f"{db_xml_file}: {original_count} entries")
    
    insertion_point = db_content.rfind('</NetworkPipes>')
    if insertion_point != -1:
        updated_content = db_content[:insertion_point] + all_new_pipes_chunk + db_content[insertion_point:]
        with open(updated_db_file, 'w', encoding='utf-8') as f:
            f.write(updated_content)
        
        # New counts for summary
        for file_info in input_files:
            with open(file_info['out'], 'r', encoding='utf-8') as f:
                c = len(re.findall(r'<NetworkPipe\b', f.read()))
                summary.append(f"{file_info['out']}: {c} entries")
        
        final_count = len(re.findall(r'<NetworkPipe\b', updated_content))
        summary.append(f"{updated_db_file}: {final_count} entries")
        
        print("\n--- Processing Summary ---")
        for line in summary:
            print(line)
    else:
        print("Error: Could not find </NetworkPipes> in original DB.")
else:
    print(f"Warning: {db_xml_file} not found.")

print("\nAll files processed successfully.")
