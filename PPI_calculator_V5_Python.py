from os import getcwd, path
from pandas import read_excel, isna

def log_error(message):
    """Log errors to a file named 'error.log'"""
    log_path = path.join(getcwd(), 'error.log')
    with open(log_path, 'a') as log_file:
        log_file.write(message + '\n')

# Der Pfad zur Excel-Datei im selben Verzeichnis wie die .exe
file_path = path.join(getcwd(), 'data_PPI.xlsx')

# Überprüfen, ob die Excel-Datei vorhanden ist
if not path.isfile(file_path):
    log_error("Error: Can not find data_PPI.xlsx. Please put it into the same folder as PI_calculator")
    print("Error: Can not find data_PPI.xlsx")
    exit(1)

# Laden der Excel-Datei
try:
    data = read_excel(file_path)
except Exception as e:
    log_error(f"Error: {e}")
    print(f"Error: Can not load data: {e}")
    exit(1)

# Überprüfen der Spaltennamen
expected_columns = [
    'case', 'preauricular_sulcus', 'dorsal_pubic_pitting', 'extended_pubic_tubercle',
    'exostoses_sacroiliac_joint_margin', 'exostoses_ventral_pubic_surface',
    'lesions_ventral_pubic_surface', 'preauricular_extension_or_notch', 'corresponding_facet',
    'margo_auricularis_groove', 'pelvic_features_no', 'PPI', 'weighted_PPI'
]

missing_columns = [col for col in expected_columns if col not in data.columns]
if missing_columns:
    log_error("Error: Missing variable: Please rename all column names to the original names")
    print(f"Error: Missing variable: Please rename all column names to the original names")
    exit(1)

#lists
PPI_list = []
weighted_PPI_list = []
pelvic_features_no_list = []

# Initialisierung & no of pelvic features
count = 0
number = 0
present = 0

for x in range(len(data['case'])):
    number = 0
    present = 0
    if not isna(data['preauricular_sulcus'][x]):
        number += 1
    if not isna(data['dorsal_pubic_pitting'][x]):
        number += 1
    if not isna(data['extended_pubic_tubercle'][x]):
        number += 1
    if not isna(data['exostoses_sacroiliac_joint_margin'][x]):
        number += 1
    if not isna(data['exostoses_ventral_pubic_surface'][x]):
        number += 1
    if not isna(data['lesions_ventral_pubic_surface'][x]):
        number += 1
    if not isna(data['preauricular_extension_or_notch'][x]):
        number += 1
    if not isna(data['corresponding_facet'][x]):
        number += 1
    if not isna(data['margo_auricularis_groove'][x]):
        number += 1
    if data['preauricular_sulcus'][x] > 0:
        present = 1
    if data['dorsal_pubic_pitting'][x] > 0:
        present += 1
    if data['extended_pubic_tubercle'][x] > 0:
        present += 1
    if data['exostoses_sacroiliac_joint_margin'][x] > 0:
        present += 1
    if data['exostoses_ventral_pubic_surface'][x] > 0:
        present += 1
    if data['lesions_ventral_pubic_surface'][x] > 0:
        present += 1
    if data['preauricular_extension_or_notch'][x] > 0:
        present += 1
    if data['corresponding_facet'][x] > 0:
        present += 1
    if data['margo_auricularis_groove'][x] > 0:
        present += 1
    pelvic_features_no_list.append(present)
    if number < 3:
        PPI_list.append(None)
        weighted_PPI_list.append(None)
        continue


#PPI
    count = 0
    if not isna(data['preauricular_sulcus'][x]):
        count += (data['preauricular_sulcus'][x] / 3)
    if not isna(data['dorsal_pubic_pitting'][x]):
        count += (data['dorsal_pubic_pitting'][x] / 2)
    if not isna(data['extended_pubic_tubercle'][x]):
        count += (data['extended_pubic_tubercle'][x] / 1)
    if not isna(data['exostoses_sacroiliac_joint_margin'][x]):
        count += (data['exostoses_sacroiliac_joint_margin'][x] / 1)
    if not isna(data['exostoses_ventral_pubic_surface'][x]):
        count += (data['exostoses_ventral_pubic_surface'][x] / 1)
    if not isna(data['lesions_ventral_pubic_surface'][x]):
        count += (data['lesions_ventral_pubic_surface'][x] / 1)
    if not isna(data['preauricular_extension_or_notch'][x]):
        count += (data['preauricular_extension_or_notch'][x] / 2)
    if not isna(data['corresponding_facet'][x]):
        count += (data['corresponding_facet'][x] / 1)
    if not isna(data['margo_auricularis_groove'][x]):
        count += (data['margo_auricularis_groove'][x] / 1)

    PPI_value = (count / number) * 100
    PPI_value = round(PPI_value, 2)
    PPI_list.append(PPI_value)

    count = 0
    number = 0
    if data['preauricular_sulcus'][x] == 0:
        number = 100
    if data['preauricular_sulcus'][x] == 1:
        count = 25
        number = 100
    if data['preauricular_sulcus'][x] == 2:
        count = 50
        number = 100
    if data['preauricular_sulcus'][x] == 3:
        count = 100
        number = 100
    if data['dorsal_pubic_pitting'][x] == 0:
        number += 100
    if data['dorsal_pubic_pitting'][x] == 1:
        count += 75
        number += 100
    if data['dorsal_pubic_pitting'][x] == 2:
        count += 100
        number += 100
    if data['extended_pubic_tubercle'][x] == 0:
        number += 10
    if data['extended_pubic_tubercle'][x] == 1:
        count += 10
        number += 10
    if data['exostoses_sacroiliac_joint_margin'][x] == 0:
        number += 10
    if data['exostoses_sacroiliac_joint_margin'][x] == 1:
        count += 10
        number += 10
    if data['exostoses_ventral_pubic_surface'][x] == 0:
        number += 10
    if data['exostoses_ventral_pubic_surface'][x] == 1:
        count += 10
        number += 10
    if data['lesions_ventral_pubic_surface'][x] == 0:
        number += 10
    if data['lesions_ventral_pubic_surface'][x] == 1:
        count += 10
        number += 10
    if data['preauricular_extension_or_notch'][x] == 0:
        number += 100
    if data['preauricular_extension_or_notch'][x] == 1:
        count += 50
        number += 100
    if data['preauricular_extension_or_notch'][x] == 2:
        count += 100
        number += 100

    if number > 0:
        weighted_PPI_value = count / (number / 100)
        weighted_PPI_value = round(weighted_PPI_value, 2)
    else:
        weighted_PPI_value = None
    weighted_PPI_list.append(weighted_PPI_value)

# Speichern der Ergebnisse
data['pelvic_features_no'] = pelvic_features_no_list
data['PPI'] = PPI_list
data['weighted_PPI'] = weighted_PPI_list

output_file = path.join(getcwd(), 'data_PPI.xlsx')

try:
    data.to_excel(output_file, index=False)
except Exception as e:
    log_error(f"Error: Can not save data.xlsx - {e}; please close the data-file")
    print(f"Error: Can not save data.xlsx: {e}; please close the data-file")
    exit(1)
