# Plugin Step Mover

An [XrmToolBox](https://www.xrmtoolbox.com/) plugin that moves plugin steps (`sdkmessageprocessingstep`) from one plugin type to another in Microsoft Dataverse.

Useful when migrating steps between plugin assemblies, for example moving steps from a development/branch plugin to your main plugin.

## Features

- Move or copy plugin steps between plugin types
- Cross-environment support (connect a second Dataverse environment as target)
- Auto-map source plugin types to matching target types
- Optionally add moved/copied steps to a solution
- Optionally keeps same step IDs when copying

## Usage

1. Open **Plugin Step Mover** in XrmToolBox and connect to your Dataverse environment.
2. Click **Load Plugins** to load all plugin assemblies and types.
3. Select a source assembly and plugin type, then click **Load Steps**.
4. Choose **Move** or **Copy** mode.
5. Select a target assembly and plugin type, or use **Auto Map All** to match automatically.
6. Select the steps you want to move/copy (use **Select All** / **Deselect All** as needed).
7. Optionally enable **Add to solution** and pick a target solution.
8. Click **Apply Move** (or **Apply Copy**) to execute.
