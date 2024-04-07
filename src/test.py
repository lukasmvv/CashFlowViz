import plotly.graph_objects as go
import os
import pandas as pd
import yaml
from pathlib import Path


def addToLabels(labels, valList):
    """ This function adds the values in a list to another list if not already in that list """
    for v in valList:
        if v not in labels and "Unnamed" not in v and v != "None":
            labels.append(v)
    return labels

def readExcelData():
    """ This function reads the excel file, loops through the sheets and returns 4 lists """

    # Path to data file
    excelPath = os.path.dirname(os.path.abspath(__file__)) + "\\excelData\\data.xlsx"
    colorPath = os.path.dirname(os.path.abspath(__file__)) + "\\config\\colors.yml"

    # Color config
    colorConfig = yaml.safe_load(Path(colorPath).read_text())

    # Empty lists
    labels = ["Fixed Expense", "Variable Expense"]
    sources = []
    targets = []
    values = []
    nodeColours = []
    colours = []

    # Reading excel file
    xls = pd.ExcelFile(excelPath)

    # Getting sheet names
    sheetNames = xls.sheet_names

    # Looping through sheets
    for sheet in sheetNames:
        df = pd.read_excel(excelPath, sheet_name=sheet, index_col="Index")

        localTargets = df.index.values.tolist()
        localSources = df.columns.tolist()
        

        # Checking if values should be added to labels
        labels = addToLabels(labels, localTargets)
        labels = addToLabels(labels, localSources)

        # Do nothing if first sheet
        if "None" not in localSources:
            for target in localTargets:
                for source in localSources:
                    if source == "Type":
                        pass
                    else:
                        # Adding normal nodes
                        sources.append(labels.index(source))
                        targets.append(labels.index(target))
                        values.append(df[source][target])
                        colours.append(colorConfig[df["Type"][target]])

                        # Adding fixed/variable expenses nodes
                        if df["Type"][target] == "Fixed Expense" or df["Type"][target] == "Variable Expense":
                            sources.append(labels.index(target))
                            targets.append(labels.index(df["Type"][target]))
                            values.append(df[source][target])
                            colours.append(colorConfig[df["Type"][target]])


                    

    return labels, sources, targets, values, nodeColours, colours


# Extracting all required lists for Sankey chart
labels, sources, targets, values, allColours, colours = readExcelData()
print(labels)
print(sources)
print(targets)
print(values)
print(allColours)
print(colours)

# Creating Sankey Viz
fig = go.Figure(data=[go.Sankey(
    valueformat = ".0f",
    valuesuffix = "€",
    node = dict(
      pad = 15,
      thickness = 20,
      line = dict(color = "black", width = 0.5),
      label = labels,
    #   color = allColours
    ),
    link = dict(
      source = sources,
      target = targets,
      value = values,
      color = colours
  ))])
fig.update_layout(title_text="Basic Sankey Diagram", font_size=10)
fig.show()