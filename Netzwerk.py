from pyvis.network import Network
import pandas as pd
import matplotlib.cm as cm

# Pfad zur Excel-Tabelle
excel_file = 'Netzwerk_Sachbegriffe.xlsx'

# Name des Arbeitsblatts in der Excel-Tabelle
sheet_name = 'Tabelle1'

# Spaltennamen in der Excel-Tabelle
quelldaten_col = 'Quelldaten'
gnd_col = 'GND'

# Daten aus der Excel-Tabelle laden
df = pd.read_excel(excel_file, sheet_name=sheet_name)

# Zählen der Häufigkeit der Matches
match_counts = df[gnd_col].value_counts().to_dict()

# Netzwerk erstellen
net = Network(notebook=True, height="800px", width="100%")

# Alle Begriffe als Knoten hinzufügen
all_terms = set(df[quelldaten_col]).union(set(df[gnd_col]))
for term in all_terms:
    if term in df[quelldaten_col].unique():
        # Quelldaten-Knoten in hellrot
        color = "rgba(255, 0, 0, 1.0)"
    else:
        # GND-Knoten in blau
        # Farbskala basierend auf der Häufigkeit der Matches
        color = cm.get_cmap('Blues')(match_counts.get(term, 0) / df.shape[0])
    net.add_node(term, title=term, color=color)

# Kanten hinzufügen
for i in range(len(df)):
    source = df.loc[i, quelldaten_col]
    target = df.loc[i, gnd_col]
    net.add_edge(source, target)

# Netzwerk in HTML-Datei speichern
net.show_buttons(filter_=['physics'])
net.show("index.html")
