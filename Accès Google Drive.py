import os
import pandas as pd
import warnings

pd.set_option('display.max_rows', 1000)
pd.set_option('display.max_columns', 1000)
pd.set_option('display.expand_frame_repr', False)

# Ignorer les avertissements liés à openpyxl
warnings.simplefilter("ignore", UserWarning)

# Chemin vers le dossier contenant les fichiers Excel
folder_path = r"G:\Drive partagés\12. Property Management BtoB\8. Reportings\Impayés"

# Liste tous les fichiers dans le dossier
files = os.listdir(folder_path)

# Filtre les fichiers Excel
excel_files = [f for f in files if f.endswith('.xlsx')]

# Trie les fichiers par date de modification (le plus récent d'abord)
latest_file = max(excel_files, key=lambda x: os.path.getmtime(os.path.join(folder_path, x)))

# Chemin complet du dernier fichier Excel
file_path = os.path.join(folder_path, latest_file)

# Chargez le dernier fichier Excel avec pandas
df = pd.read_excel(file_path)

# Affichez le contenu du DataFrame
print(df)
