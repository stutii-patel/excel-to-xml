import xml.etree.ElementTree as ET
import pandas as pd
import matplotlib as mpl
import matplotlib.pyplot as pl
import os
import re
from pipe_utils import calculate_insulation_thickness

path = r'C:\Daten\2_NextCloud\VICUS\VICUS-Daten\02_Entwicklung\07_VICUS_Datenbanken\Isoplus'
xlsx_file = 'VICUS_DB_Template_Rohre_isoplus_10_2025_bearbeitet.xlsx'
file = os.path.join(path, xlsx_file)
df = pd.read_excel(file, skiprows=1)

# first id
id = 1100500

# Create the root element
root = ET.Element("NetworkPipes")

# Add each as a pipe element
cmap = pl.get_cmap('turbo')
total_number = -1
color_count = 0

# properties


# count thet total numbers per diameter block
diam_col = 'Außendurchmesser [mm]'
total_numbers = [None] * len(df)
start_idx = 0
previous_diameter = df.loc[0, diam_col]
for i in range(1, len(df)):
    current_diameter = df.loc[i, diam_col]

    # If diameter decreases → finalize previous block
    if current_diameter < previous_diameter:
        block_length = i - start_idx
        total_numbers[start_idx:i] = [block_length] * block_length
        start_idx = i  # start new block

    previous_diameter = current_diameter
# Handle the last block (end of dataframe)
block_length = len(df) - start_idx
total_numbers[start_idx:] = [block_length] * block_length
# Assign to DataFrame
df['total_number'] = total_numbers


prev_da = -1
for index, row in df.iterrows():

    if row["Material de"] == "":
        continue

    if row['Außendurchmesser [mm]'] < prev_da:
        color_count = 0        
    prev_da = row['Außendurchmesser [mm]']
    total_number = row['total_number']
    color = mpl.colors.rgb2hex(cmap(color_count / row['total_number']), keep_alpha=False)
    color_count += 1

    da = row['Außendurchmesser [mm]']
    s = row['Wandstärke [mm]']
    name = f"{da} x {s}"
    cat_name = f"DE: {row["Material de"]} | EN: {row["Material en"]}"
    
    product_name = row["Produkt"]
    product_name = product_name.replace("Stahl-Einzelrohr, ", "")
    product_name = product_name.replace("Stahl-Doppelrohr, ", "")

    pn_value = row["PN [bar]"]
    spacing = row["Abstand Vor- und Rücklauf [mm]"]
    UValue = row["U-Wert [W/mK]"]
    total_outer_diameter = row["Außendurchmesser gesamt mit Isolierung und Schutzschicht [mm]"]
    roughness = row["Rohrrauigkeit [mm]"]

    if row['Material en'] == 'PE insulated':
        material_standard = "PlasticPipe"
        rho_wall = 960
        cp = 1900
        lambda_wall = 0.4
    else:
        material_standard = "EnStandard"
        rho_wall = 7900
        cp = 480
        lambda_wall = 50

    pipe = ET.Element("NetworkPipe", id=str(id + index + 1), categoryName=cat_name, color=color)
    pipe.set("manufacturerName", "Isoplus")
    pipe.set("productName", product_name)

    ET.SubElement(pipe, "IBK:Parameter", name="DiameterOutside", unit="mm").text = str(da)
    ET.SubElement(pipe, "IBK:Parameter", name="ThicknessWall", unit="mm").text = str(s)
    ET.SubElement(pipe, "IBK:Parameter", name="RoughnessWall", unit="mm").text = str(roughness)
    ET.SubElement(pipe, "IBK:Parameter", name="DensityWall", unit="kg/m3").text = str(rho_wall)
    ET.SubElement(pipe, "IBK:Parameter", name="HeatCapacityWall", unit="J/kgK").text = str(cp)
    ET.SubElement(pipe, "IBK:Parameter", name="ThermalConductivityWall", unit="W/mK").text = str(lambda_wall)
    # ET.SubElement(pipe, "IBK:Parameter", name="ThicknessInsulation", unit="mm").text = str(s_ins_fin)
    # ET.SubElement(pipe, "IBK:Parameter", name="ThermalConductivityInsulation", unit="W/mK").text = str(row["lambda_ins"])
    
    ET.SubElement(pipe, "IBK:Parameter", name="FixedUValue", unit="W/mK").text = str(UValue)
    ET.SubElement(pipe, "IBK:Parameter", name="FixedTotalOuterDiameter", unit="mm").text = str(total_outer_diameter)
    ET.SubElement(pipe, "IBK:Parameter", name="PipeSpacing", unit="mm").text = str(spacing)

    ET.SubElement(pipe, "NominalPressure").text = str(pn_value)
    ET.SubElement(pipe, "FixedUValueGiven").text = "true"
    
    if row['Typ'] == 'Einzelrohr':
        ET.SubElement(pipe, "PipeLayout").text = "SinglePipe"
    else:
        ET.SubElement(pipe, "PipeLayout").text = "TwinPipe"

    ET.SubElement(pipe, "PipeMaterialStandard").text = material_standard 

    
    root.append(pipe)


def prettify(element, level=0):
    # This function adds new lines and indentation between XML elements
    indent = "\n" + "  " * level
    if element:
        if not element.text or not element.text.strip():
            element.text = indent + "  "
        if not element.tail or not element.tail.strip():
            element.tail = indent
    for subelement in element:
        prettify(subelement, level+1)
    if not element.tail or not element.tail.strip():
        element.tail = indent


# Use prettify to format the output with new lines and indents
prettify(root)

# Convert to a string
tree_string = ET.tostring(root, 'unicode')

# Write to a file
xml_file = xlsx_file.replace('xlsx', 'xml')
with open(os.path.join(path, xml_file), 'w') as f:
    f.write(tree_string)

# Generate the XML tree
tree = ET.ElementTree(root)

# Write the XML to a file
tree.write(os.path.join(path, "pipes_created.xml"), encoding="utf-8", xml_declaration=True)

print(f"XML file '{xml_file}' has been created with the pipe data.")


