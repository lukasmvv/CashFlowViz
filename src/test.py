import plotly.graph_objects as go
import os
import pandas as pd

def addToLabels(labels, valList):
    """ This function adds the values in a list to another list if not already in that list """
    for v in valList:
        if v not in labels and "Unnamed" not in v and v != "None":
            labels.append(v)
    return labels

def readExcelData(path):
    """ This function reads the excel file, loops through the sheets and returns 4 lists """
    # Empty lists
    labels = []
    sources = []
    targets = []
    values = []

    # Reading excel file
    xls = pd.ExcelFile(path)

    # Getting sheet names
    sheetNames = xls.sheet_names

    # Looping through sheets
    for sheet in sheetNames:
        df = pd.read_excel(path, sheet_name=sheet, index_col="Index")

        localTargets = df.index.values.tolist()
        localSources = df.columns.tolist()

        # Checking if values should be added to labels
        labels = addToLabels(labels, localTargets)
        labels = addToLabels(labels, localSources)

        # Do nothing if first sheet
        if "None" not in localSources:
            for target in localTargets:
                for source in localSources:
                    sources.append(labels.index(source))
                    targets.append(labels.index(target))
                    values.append(df[source][target])
            

    return labels, sources, targets, values

# Path to data file
excelPath = os.path.dirname(os.path.abspath(__file__)) + "\\excelData\\data.xlsx"

# Extracting all required lists for Sankey chart
labels, sources, targets, values = readExcelData(excelPath)
print(labels)
print(sources)
print(targets)
print(values)

# Creating Sankey Viz
fig = go.Figure(data=[go.Sankey(
    node = dict(
      pad = 15,
      thickness = 20,
      line = dict(color = "black", width = 0.5),
      label = labels,
      color = "blue"
    ),
    link = dict(
      source = sources,
      target = targets,
      value = values
  ))])
fig.update_layout(title_text="Basic Sankey Diagram", font_size=10)
fig.show()