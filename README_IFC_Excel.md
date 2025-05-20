# IFC-CSV Excel Integration Tool

This script enables export, manual editing, and re-import of IFC model data using Excel and the `ifcopenshell` and `xlwings` libraries.

## 📄 Description

The script performs the following steps:

1. **Loads an IFC file** using `ifcopenshell`.
2. **Exports selected attributes** of IFC elements to a CSV file.
3. **Opens the CSV in Excel** using `xlwings`, allowing the user to edit the values manually.
4. **Monitors Excel** until it is closed.
5. **Re-imports the edited CSV** into the IFC model.
6. **Writes a new IFC file** with the updated values.

## ✅ Features

- Interactive workflow between BIM data and Excel
- Exports IFC attributes such as `Name`, `Description`, and `PredefinedType`
- Supports editing directly in Excel before updating the IFC model
- Automatically waits for the user to finish editing

## 📦 Requirements

- Python 3.7+
- [`ifcopenshell`](https://github.com/IfcOpenShell/IfcOpenShell)
- [`xlwings`](https://docs.xlwings.org/)
- [`ifccsv`](https://github.com/IfcOpenShell/IfcOpenShell/blob/v0.7.0/src/ifccsv/README.md)

Install dependencies using pip:

```bash
pip install ifcopenshell xlwings
```

## 🚀 Usage

1. Make sure Microsoft Excel is installed on your system.
2. Save your IFC file (e.g., `20230907_Rohdaten_Betondruckfestigkeit.ifc`) in the same directory.
3. Run the script.

```bash
python script.py
```

4. Edit the values in Excel and **close Excel** when done.
5. A new IFC file will be saved as `20230907_Rohdaten_Betondruckfestigkeit_updated.ifc`.

## 📝 Notes

- Exported attributes: `Name`, `Description`, `PredefinedType`
- CSV file used: `file.csv`
- Script waits until Excel is closed before applying changes
- Make sure not to rename or move `file.csv` while editing in Excel

## 📂 File Naming

- Input IFC file: `20230907_Rohdaten_Betondruckfestigkeit.ifc`
- Output IFC file: `20230907_Rohdaten_Betondruckfestigkeit_updated.ifc`
- Temporary CSV file: `file.csv`

## 🧑‍💻 Author

Amir Sadrnia

## 📝 License

This project is open source and available under the [MIT License](https://opensource.org/licenses/MIT).
