using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using McTools.Xrm.Connection;
using Microsoft.Xrm.Sdk;
using XrmToolBox.Extensibility;
using XrmToolBox.Extensibility.Interfaces;

namespace PluginStepMover;

public class PluginStepMoverControl : MultipleConnectionsPluginControlBase, IPluginMetadata
{
    private readonly StepMoveService _stepMoveService = new();

    private readonly BindingList<StepItem> _steps = new();
    private readonly BindingList<PluginAssemblyItem> _sourceAssemblies = new();
    private readonly BindingList<PluginAssemblyItem> _targetAssemblies = new();
    private readonly BindingList<PluginTypeItem> _sourcePluginTypes = new();
    private readonly BindingList<PluginTypeItem> _targetPluginTypes = new();

    private readonly List<PluginTypeItem> _allPluginTypes = new();
    private readonly List<PluginTypeItem> _allTargetPluginTypes = new();

    private readonly BindingList<SolutionItem> _solutions = new();

    private readonly DataGridView _stepsGrid = new();
    private readonly ComboBox _sourceAssemblyCombo = new();
    private readonly ComboBox _targetAssemblyCombo = new();
    private readonly ComboBox _sourceTypeCombo = new();
    private readonly ComboBox _targetTypeCombo = new();
    private readonly ComboBox _solutionCombo = new();
    private readonly CheckBox _addToSolutionCheck = new();
    private readonly CheckBox _keepSameIdsCheck = new();
    private readonly RadioButton _moveRadio = new();
    private readonly RadioButton _copyRadio = new();
    private readonly Button _loadPluginsButton = new();
    private readonly Button _autoMapButton = new();
    private readonly Button _loadButton = new();
    private readonly Button _applyButton = new();
    private readonly Button _setTargetButton = new();
    private readonly Button _connectTargetButton = new();
    private readonly Button _disconnectTargetButton = new();
    private readonly Button _selectAllButton = new();
    private readonly Button _deselectAllButton = new();
    private readonly System.Windows.Forms.Label _selectionCountLabel = new();
    private readonly System.Windows.Forms.Label _targetEnvLabel = new();
    private readonly TextBox _outputBox = new();
    private bool _isBindingPluginData;
    private bool _solutionsLoaded;

    private IOrganizationService? _targetService;

    private bool IsCrossEnvironment => _targetService != null;

    public PluginStepMoverControl()
    {
        BuildUi();
        RegisterEvents();
        UpdateOperationModeUi();
        EnableActions();
    }

    string IPluginMetadata.BackgroundColor => "White";
    string IPluginMetadata.BigImageBase64 => PluginBranding.BigImageBase64;
    string IPluginMetadata.Description => "Move or copy plugin steps between plugin types, with cross-environment support.";
    string IPluginMetadata.Name => "Plugin Step Mover";
    string IPluginMetadata.PrimaryFontColor => "Black";
    string IPluginMetadata.SecondaryFontColor => "DimGray";
    string IPluginMetadata.SmallImageBase64 => PluginBranding.SmallImageBase64;

    private void RegisterEvents()
    {
        _loadPluginsButton.Click += (_, _) => ExecuteMethod(LoadPluginTypes);
        _autoMapButton.Click += (_, _) => ExecuteMethod(AutoMapSteps);
        _loadButton.Click += (_, _) => ExecuteMethod(LoadSourceSteps);
        _applyButton.Click += (_, _) => ExecuteMethod(ApplyOperation);
        _setTargetButton.Click += (_, _) => SetTargetForSelectedSteps();
        _selectAllButton.Click += (_, _) => SetAllSelected(true);
        _deselectAllButton.Click += (_, _) => SetAllSelected(false);
        _connectTargetButton.Click += (_, _) => ConnectTargetEnvironment();
        _disconnectTargetButton.Click += (_, _) => DisconnectTargetEnvironment();
        ConnectionUpdated += (_, _) => EnableActions();

        _moveRadio.CheckedChanged += (_, _) => UpdateOperationModeUi();
        _copyRadio.CheckedChanged += (_, _) => UpdateOperationModeUi();

        _addToSolutionCheck.CheckedChanged += (_, _) =>
        {
            _solutionCombo.Enabled = _addToSolutionCheck.Checked;

            if (_addToSolutionCheck.Checked && !_solutionsLoaded && Service != null)
            {
                LoadSolutions(_targetService ?? Service);
            }

            EnableActions();
        };

        _sourceAssemblyCombo.SelectedIndexChanged += (_, _) =>
        {
            if (_isBindingPluginData)
            {
                return;
            }

            FilterSourcePluginTypes();
            if (!IsCrossEnvironment)
            {
                TryAutoSelectTargetMainFromSource();
            }

            PopulateGrid(Array.Empty<StepItem>());
            EnableActions();
        };

        _targetAssemblyCombo.SelectedIndexChanged += (_, _) =>
        {
            if (_isBindingPluginData)
            {
                return;
            }

            FilterTargetPluginTypes();
            if (!IsCrossEnvironment)
            {
                TryAutoSelectTargetFromSource();
            }

            EnableActions();
        };

        _sourceTypeCombo.SelectedIndexChanged += (_, _) =>
        {
            if (_isBindingPluginData)
            {
                return;
            }

            if (!IsCrossEnvironment)
            {
                TryAutoSelectTargetFromSource();
            }

            EnableActions();
        };
        _targetTypeCombo.SelectedIndexChanged += (_, _) => EnableActions();
        _steps.ListChanged += (_, _) =>
        {
            EnableActions();
            UpdateSelectionCount();
        };
        _stepsGrid.CurrentCellDirtyStateChanged += StepsGrid_CurrentCellDirtyStateChanged;
        _stepsGrid.CellValueChanged += (_, _) =>
        {
            EnableActions();
            UpdateSelectionCount();
        };
    }

    private void BuildUi()
    {
        Dock = DockStyle.Fill;

        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 4,
            Padding = new Padding(8)
        };
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 148));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 75));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 25));

        var configPanel = BuildConfigPanel();
        var gridContainer = BuildGridPanel();

        layout.Controls.Add(BuildToolbar(), 0, 0);
        layout.Controls.Add(configPanel, 0, 1);
        layout.Controls.Add(gridContainer, 0, 2);
        layout.Controls.Add(BuildOutputBox(), 0, 3);

        Controls.Add(layout);
    }

    private FlowLayoutPanel BuildToolbar()
    {
        var toolbarPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            WrapContents = true,
            Margin = new Padding(0, 0, 0, 4)
        };

        _loadPluginsButton.Text = "Load Plugins";
        _loadPluginsButton.AutoSize = true;

        _applyButton.Text = "Apply Move";
        _applyButton.AutoSize = true;
        _applyButton.Font = new Font(_applyButton.Font, FontStyle.Bold);
        _applyButton.Margin = new Padding(12, 3, 3, 3);

        toolbarPanel.Controls.Add(_loadPluginsButton);
        toolbarPanel.Controls.Add(_applyButton);

        return toolbarPanel;
    }

    private TableLayoutPanel BuildConfigPanel()
    {
        var configPanel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 3,
            RowCount = 1,
            Height = 144
        };
        configPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33));
        configPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 37));
        configPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30));

        configPanel.Controls.Add(BuildSourceGroup(), 0, 0);
        configPanel.Controls.Add(BuildTargetGroup(), 1, 0);
        configPanel.Controls.Add(BuildOptionsGroup(), 2, 0);

        return configPanel;
    }

    private GroupBox BuildSourceGroup()
    {
        var sourceGroup = new GroupBox
        {
            Text = "Source",
            Dock = DockStyle.Fill,
            Padding = new Padding(6, 2, 6, 4)
        };

        var sourceInner = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 5,
            AutoSize = false
        };
        sourceInner.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        for (var i = 0; i < 5; i++) sourceInner.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        var sourceMainLabel = new System.Windows.Forms.Label { Text = "Assembly:", AutoSize = true, Margin = new Padding(0, 1, 0, 0) };
        _sourceAssemblyCombo.Dock = DockStyle.Fill;
        _sourceAssemblyCombo.DropDownStyle = ComboBoxStyle.DropDownList;

        var sourceTypeLabel = new System.Windows.Forms.Label { Text = "Plugin type:", AutoSize = true, Margin = new Padding(0, 3, 0, 0) };
        _sourceTypeCombo.Dock = DockStyle.Fill;
        _sourceTypeCombo.DropDownStyle = ComboBoxStyle.DropDownList;

        var sourceButtonRow = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = new Padding(0, 4, 0, 0), Padding = new Padding(0) };
        _autoMapButton.Text = "Auto Map All";
        _autoMapButton.AutoSize = true;
        _autoMapButton.Margin = new Padding(0, 0, 4, 0);
        _loadButton.Text = "Load Steps";
        _loadButton.AutoSize = true;
        _loadButton.Margin = new Padding(0);
        sourceButtonRow.Controls.Add(_autoMapButton);
        sourceButtonRow.Controls.Add(_loadButton);

        sourceInner.Controls.Add(sourceMainLabel, 0, 0);
        sourceInner.Controls.Add(_sourceAssemblyCombo, 0, 1);
        sourceInner.Controls.Add(sourceTypeLabel, 0, 2);
        sourceInner.Controls.Add(_sourceTypeCombo, 0, 3);
        sourceInner.Controls.Add(sourceButtonRow, 0, 4);
        sourceGroup.Controls.Add(sourceInner);

        return sourceGroup;
    }

    private GroupBox BuildTargetGroup()
    {
        var targetGroup = new GroupBox
        {
            Text = "Target",
            Dock = DockStyle.Fill,
            Padding = new Padding(6, 2, 6, 4)
        };

        var targetInner = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 6,
            AutoSize = false
        };
        targetInner.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        for (var i = 0; i < 6; i++) targetInner.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        var targetMainLabel = new System.Windows.Forms.Label { Text = "Assembly:", AutoSize = true, Margin = new Padding(0, 1, 0, 0) };
        _targetAssemblyCombo.Dock = DockStyle.Fill;
        _targetAssemblyCombo.DropDownStyle = ComboBoxStyle.DropDownList;

        var targetTypeLabel = new System.Windows.Forms.Label { Text = "Plugin type:", AutoSize = true, Margin = new Padding(0, 3, 0, 0) };
        _targetTypeCombo.Dock = DockStyle.Fill;
        _targetTypeCombo.DropDownStyle = ComboBoxStyle.DropDownList;

        var targetActionRow = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = new Padding(0, 4, 0, 0), Padding = new Padding(0) };
        _setTargetButton.Text = "Assign to Selected";
        _setTargetButton.AutoSize = true;
        _setTargetButton.Margin = new Padding(0, 0, 12, 0);

        var envSeparator = new System.Windows.Forms.Label { Text = "Environment:", AutoSize = true, Padding = new Padding(0, 5, 0, 0), ForeColor = Color.DimGray };
        _targetEnvLabel.Text = "Same environment";
        _targetEnvLabel.AutoSize = true;
        _targetEnvLabel.Padding = new Padding(0, 5, 2, 0);
        _targetEnvLabel.ForeColor = Color.DimGray;
        _connectTargetButton.Text = "Connect...";
        _connectTargetButton.AutoSize = true;
        _connectTargetButton.Margin = new Padding(0, 0, 2, 0);
        _disconnectTargetButton.Text = "Disconnect";
        _disconnectTargetButton.AutoSize = true;
        _disconnectTargetButton.Visible = false;

        targetActionRow.Controls.Add(_setTargetButton);
        targetActionRow.Controls.Add(envSeparator);
        targetActionRow.Controls.Add(_targetEnvLabel);
        targetActionRow.Controls.Add(_connectTargetButton);
        targetActionRow.Controls.Add(_disconnectTargetButton);

        targetInner.Controls.Add(targetMainLabel, 0, 0);
        targetInner.Controls.Add(_targetAssemblyCombo, 0, 1);
        targetInner.Controls.Add(targetTypeLabel, 0, 2);
        targetInner.Controls.Add(_targetTypeCombo, 0, 3);
        targetInner.Controls.Add(targetActionRow, 0, 4);
        targetGroup.Controls.Add(targetInner);

        return targetGroup;
    }

    private GroupBox BuildOptionsGroup()
    {
        var optionsGroup = new GroupBox
        {
            Text = "Options",
            Dock = DockStyle.Fill,
            Padding = new Padding(6, 2, 6, 4)
        };

        var optionsInner = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 4,
            AutoSize = false
        };
        optionsInner.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        for (var i = 0; i < 4; i++) optionsInner.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        var modeRow = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = new Padding(0), Padding = new Padding(0, 1, 0, 0) };
        _moveRadio.Text = "Move";
        _moveRadio.AutoSize = true;
        _moveRadio.Checked = true;
        _moveRadio.Margin = new Padding(0, 0, 6, 0);
        _copyRadio.Text = "Copy";
        _copyRadio.AutoSize = true;
        modeRow.Controls.Add(_moveRadio);
        modeRow.Controls.Add(_copyRadio);

        _keepSameIdsCheck.Text = "Keep same IDs";
        _keepSameIdsCheck.AutoSize = true;
        _keepSameIdsCheck.Enabled = false;
        _keepSameIdsCheck.Margin = new Padding(3, 2, 0, 0);

        _addToSolutionCheck.Text = "Add to solution";
        _addToSolutionCheck.AutoSize = true;
        _addToSolutionCheck.Margin = new Padding(3, 4, 0, 0);

        _solutionCombo.Dock = DockStyle.Fill;
        _solutionCombo.DropDownStyle = ComboBoxStyle.DropDownList;
        _solutionCombo.Enabled = false;

        optionsInner.Controls.Add(modeRow, 0, 0);
        optionsInner.Controls.Add(_keepSameIdsCheck, 0, 1);
        optionsInner.Controls.Add(_addToSolutionCheck, 0, 2);
        optionsInner.Controls.Add(_solutionCombo, 0, 3);
        optionsGroup.Controls.Add(optionsInner);

        return optionsGroup;
    }

    private Panel BuildGridPanel()
    {
        var gridToolbar = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            WrapContents = false,
            Margin = new Padding(0, 4, 0, 2)
        };

        _selectAllButton.Text = "Select All";
        _selectAllButton.AutoSize = true;
        _selectAllButton.FlatStyle = FlatStyle.System;

        _deselectAllButton.Text = "Deselect All";
        _deselectAllButton.AutoSize = true;
        _deselectAllButton.FlatStyle = FlatStyle.System;

        _selectionCountLabel.Text = "0 steps selected";
        _selectionCountLabel.AutoSize = true;
        _selectionCountLabel.Padding = new Padding(8, 6, 0, 0);
        _selectionCountLabel.ForeColor = Color.DimGray;

        gridToolbar.Controls.Add(_selectAllButton);
        gridToolbar.Controls.Add(_deselectAllButton);
        gridToolbar.Controls.Add(_selectionCountLabel);

        var gridContainer = new Panel { Dock = DockStyle.Fill };

        _stepsGrid.Dock = DockStyle.Fill;
        _stepsGrid.AutoGenerateColumns = false;
        _stepsGrid.AllowUserToAddRows = false;
        _stepsGrid.AllowUserToDeleteRows = false;
        _stepsGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        _stepsGrid.MultiSelect = true;
        _stepsGrid.BackgroundColor = SystemColors.Window;
        _stepsGrid.BorderStyle = BorderStyle.FixedSingle;
        _stepsGrid.RowHeadersVisible = false;
        _stepsGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        _stepsGrid.AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle
        {
            BackColor = Color.FromArgb(245, 249, 255)
        };

        _stepsGrid.Columns.Add(new DataGridViewCheckBoxColumn
        {
            HeaderText = "",
            DataPropertyName = nameof(StepItem.Selected),
            AutoSizeMode = DataGridViewAutoSizeColumnMode.None,
            Width = 40,
            MinimumWidth = 40
        });

        _stepsGrid.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = "Step",
            DataPropertyName = nameof(StepItem.StepName),
            FillWeight = 25,
            MinimumWidth = 120,
            ReadOnly = true
        });

        _stepsGrid.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = "Message",
            DataPropertyName = nameof(StepItem.MessageName),
            FillWeight = 10,
            MinimumWidth = 70,
            ReadOnly = true
        });

        _stepsGrid.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = "Entity",
            DataPropertyName = nameof(StepItem.PrimaryEntity),
            FillWeight = 10,
            MinimumWidth = 70,
            ReadOnly = true
        });

        _stepsGrid.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = "Source Assembly",
            DataPropertyName = nameof(StepItem.SourceAssemblyName),
            FillWeight = 15,
            MinimumWidth = 100,
            ReadOnly = true
        });

        _stepsGrid.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = "Source Plugin Type",
            DataPropertyName = nameof(StepItem.SourcePluginTypeName),
            FillWeight = 20,
            MinimumWidth = 120,
            ReadOnly = true
        });

        _stepsGrid.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = "Suggested Target",
            DataPropertyName = nameof(StepItem.SuggestedTargetPluginTypeName),
            FillWeight = 20,
            MinimumWidth = 120,
            ReadOnly = true
        });

        _stepsGrid.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = "Warning",
            DataPropertyName = nameof(StepItem.MatchWarning),
            FillWeight = 15,
            MinimumWidth = 80,
            ReadOnly = true
        });

        _stepsGrid.DataSource = _steps;

        // Add grid first (Fill), then toolbar (Top) - WinForms docks in reverse add order
        gridContainer.Controls.Add(_stepsGrid);
        gridContainer.Controls.Add(gridToolbar);

        return gridContainer;
    }

    private TextBox BuildOutputBox()
    {
        _outputBox.Dock = DockStyle.Fill;
        _outputBox.Multiline = true;
        _outputBox.ScrollBars = ScrollBars.Vertical;
        _outputBox.ReadOnly = true;
        _outputBox.BackColor = Color.FromArgb(30, 30, 30);
        _outputBox.ForeColor = Color.FromArgb(220, 220, 220);
        _outputBox.Font = new Font("Consolas", 9f);
        _outputBox.BorderStyle = BorderStyle.FixedSingle;

        return _outputBox;
    }

    // --- Target Environment ---

    private void ConnectTargetEnvironment()
    {
        AddAdditionalOrganization();
    }

    protected override void ConnectionDetailsUpdated(NotifyCollectionChangedEventArgs e)
    {
        if (e.Action == NotifyCollectionChangedAction.Add && e.NewItems != null)
        {
            var detail = (ConnectionDetail)e.NewItems[0]!;
            _targetService = detail.GetCrmServiceClient();
            if (_targetService == null)
            {
                AppendOutput("Failed to connect to target environment: connection returned null.");
                return;
            }
            _targetEnvLabel.Text = detail.ConnectionName ?? detail.WebApplicationUrl ?? "Connected";
            _targetEnvLabel.ForeColor = Color.FromArgb(0, 120, 215);
            _connectTargetButton.Visible = false;
            _disconnectTargetButton.Visible = true;
            _copyRadio.Checked = true;
            _keepSameIdsCheck.Checked = true;
            UpdateOperationModeUi();
            AppendOutput($"Target environment connected: {_targetEnvLabel.Text}");
            ReloadTargetData();
        }

        EnableActions();
    }

    private void DisconnectTargetEnvironment()
    {
        // Remove all additional connections from the framework
        foreach (var detail in AdditionalConnectionDetails.ToList())
        {
            RemoveAdditionalOrganization(detail);
        }

        _targetService = null;
        _targetEnvLabel.Text = "Same environment";
        _targetEnvLabel.ForeColor = Color.DimGray;
        _connectTargetButton.Visible = true;
        _disconnectTargetButton.Visible = false;

        _keepSameIdsCheck.Checked = false;
        UpdateOperationModeUi();

        // Reload target combos from source data
        _allTargetPluginTypes.Clear();
        if (_allPluginTypes.Count > 0)
        {
            ReloadTargetCombosFromSource();
        }

        _solutionsLoaded = false;
        PopulateGrid(Array.Empty<StepItem>());
        AppendOutput("Target environment disconnected. Using same environment.");
        EnableActions();
    }

    private void ReloadTargetData()
    {
        if (_targetService == null)
        {
            return;
        }

        WorkAsync(new WorkAsyncInfo
        {
            Message = "Loading target environment plugins...",
            Work = (worker, args) =>
            {
                var assemblies = _stepMoveService.GetPluginAssemblies(_targetService);
                var types = _stepMoveService.GetPluginTypes(_targetService);
                args.Result = new Tuple<IReadOnlyList<PluginAssemblyItem>, IReadOnlyList<PluginTypeItem>>(assemblies, types);
            },
            PostWorkCallBack = args =>
            {
                if (args.Error != null)
                {
                    AppendOutput($"Failed to load target environment data: {args.Error.Message}");
                    LogError(args.Error.ToString());
                    return;
                }

                var result = (Tuple<IReadOnlyList<PluginAssemblyItem>, IReadOnlyList<PluginTypeItem>>)args.Result;
                PopulateTargetData(result.Item1, result.Item2);
                AppendOutput($"Loaded {result.Item1.Count} assemblies and {result.Item2.Count} plugin types from target environment.");
                EnableActions();
            }
        });
    }

    private void PopulateTargetData(IReadOnlyList<PluginAssemblyItem> assemblies, IReadOnlyList<PluginTypeItem> types)
    {
        _isBindingPluginData = true;

        try
        {
            _allTargetPluginTypes.Clear();
            _allTargetPluginTypes.AddRange(types);

            _targetAssemblies.RaiseListChangedEvents = false;
            _targetAssemblies.Clear();

            foreach (var assembly in assemblies)
            {
                _targetAssemblies.Add(new PluginAssemblyItem
                {
                    AssemblyId = assembly.AssemblyId,
                    Name = assembly.Name
                });
            }

            _targetAssemblies.RaiseListChangedEvents = true;
            _targetAssemblies.ResetBindings();

            _targetAssemblyCombo.DataSource = _targetAssemblies;
            _targetAssemblyCombo.DisplayMember = nameof(PluginAssemblyItem.DisplayName);
            _targetAssemblyCombo.ValueMember = nameof(PluginAssemblyItem.AssemblyId);

            if (_targetAssemblies.Count > 0)
            {
                TrySetSelectedIndex(_targetAssemblyCombo, 0);
            }

            FilterTargetPluginTypes();
        }
        finally
        {
            _isBindingPluginData = false;
        }

        // Auto-select the target assembly matching the current source assembly
        // Must be outside _isBindingPluginData block so SelectedIndexChanged fires and FilterTargetPluginTypes runs
        TryAutoSelectTargetMainFromSource();

        EnableActions();
    }

    private void ReloadTargetCombosFromSource()
    {
        _isBindingPluginData = true;

        try
        {
            _targetAssemblies.RaiseListChangedEvents = false;
            _targetAssemblies.Clear();

            foreach (var assembly in _sourceAssemblies)
            {
                _targetAssemblies.Add(new PluginAssemblyItem
                {
                    AssemblyId = assembly.AssemblyId,
                    Name = assembly.Name
                });
            }

            _targetAssemblies.RaiseListChangedEvents = true;
            _targetAssemblies.ResetBindings();

            if (_targetAssemblies.Count > 1)
            {
                TrySetSelectedIndex(_targetAssemblyCombo, 1);
            }
            else if (_targetAssemblies.Count == 1)
            {
                TrySetSelectedIndex(_targetAssemblyCombo, 0);
            }

            FilterTargetPluginTypes();
        }
        finally
        {
            _isBindingPluginData = false;
        }

        TryAutoSelectTargetMainFromSource();
        TryAutoSelectTargetFromSource();
    }

    // --- Operation Mode ---

    private void UpdateOperationModeUi()
    {
        var isCopy = _copyRadio.Checked;
        var isCrossEnv = IsCrossEnvironment;

        // Keep Same IDs: disabled only for same-env copy (can't create with same ID that already exists)
        var canKeepIds = isCrossEnv || !isCopy;
        _keepSameIdsCheck.Enabled = canKeepIds;
        if (!canKeepIds)
        {
            _keepSameIdsCheck.Checked = false;
        }

        // Update Apply button text
        _applyButton.Text = isCopy ? "Apply Copy" : "Apply Move";

        EnableActions();
    }

    // --- UI Helpers ---

    private void StepsGrid_CurrentCellDirtyStateChanged(object? sender, EventArgs e)
    {
        if (_stepsGrid.IsCurrentCellDirty)
        {
            _stepsGrid.CommitEdit(DataGridViewDataErrorContexts.Commit);
        }
    }

    private void SetTargetForSelectedSteps()
    {
        var target = _targetTypeCombo.SelectedItem as PluginTypeItem;

        if (target == null)
        {
            AppendOutput("Select a target plugin type first.");
            return;
        }

        var selectedSteps = _steps.Where(s => s.Selected).ToList();

        if (!selectedSteps.Any())
        {
            AppendOutput("Select at least one step in the grid first.");
            return;
        }

        _steps.RaiseListChangedEvents = false;

        foreach (var step in selectedSteps)
        {
            step.SuggestedTargetPluginTypeId = target.PluginTypeId;
            step.SuggestedTargetPluginTypeName = target.Name;
            step.MatchWarning = "Manually set";
        }

        _steps.RaiseListChangedEvents = true;
        _steps.ResetBindings();

        AppendOutput($"Set target to '{target.Name}' for {selectedSteps.Count} step(s).");
    }

    private void SetAllSelected(bool selected)
    {
        _steps.RaiseListChangedEvents = false;

        foreach (var step in _steps)
        {
            step.Selected = selected;
        }

        _steps.RaiseListChangedEvents = true;
        _steps.ResetBindings();
        UpdateSelectionCount();
        EnableActions();
    }

    private void UpdateSelectionCount()
    {
        var count = _steps.Count(s => s.Selected);
        var total = _steps.Count;
        _selectionCountLabel.Text = $"{count} of {total} steps selected";
    }

    // --- Data Loading ---

    private void LoadPluginTypes()
    {
        WorkAsync(new WorkAsyncInfo
        {
            Message = "Loading main plugins and plugin types...",
            Work = (worker, args) =>
            {
                var assemblies = _stepMoveService.GetPluginAssemblies(Service);
                var types = _stepMoveService.GetPluginTypes(Service);
                args.Result = new Tuple<IReadOnlyList<PluginAssemblyItem>, IReadOnlyList<PluginTypeItem>>(assemblies, types);
            },
            PostWorkCallBack = args =>
            {
                if (args.Error != null)
                {
                    AppendOutput($"Load failed: {args.Error.Message}");
                    LogError(args.Error.ToString());
                    return;
                }

                var result = (Tuple<IReadOnlyList<PluginAssemblyItem>, IReadOnlyList<PluginTypeItem>>)args.Result;
                PopulatePluginData(result.Item1, result.Item2);
                PopulateGrid(Array.Empty<StepItem>());
                AppendOutput($"Loaded {result.Item1.Count} main plugins and {result.Item2.Count} plugin types.");

                // If cross-env, also reload target data
                if (IsCrossEnvironment)
                {
                    ReloadTargetData();
                }

                EnableActions();
            }
        });
    }

    private void AutoMapSteps()
    {
        var sourceAssembly = _sourceAssemblyCombo.SelectedItem as PluginAssemblyItem;
        var targetAssembly = _targetAssemblyCombo.SelectedItem as PluginAssemblyItem;

        if (sourceAssembly == null)
        {
            AppendOutput("Select a source assembly first.");
            return;
        }

        if (targetAssembly == null)
        {
            AppendOutput("Select a target assembly first.");
            return;
        }

        var targetTypesSnapshot = _targetPluginTypes.ToList();
        var stepCountService = _targetService ?? Service;

        WorkAsync(new WorkAsyncInfo
        {
            Message = "Loading all source steps and auto-matching targets...",
            Work = (worker, args) =>
            {
                var steps = _stepMoveService.GetSteps(Service, sourceAssemblyId: sourceAssembly.AssemblyId);
                var existingStepCounts = _stepMoveService.GetStepCountsByPluginTypeIds(stepCountService, targetTypesSnapshot.Select(t => t.PluginTypeId).ToList());

                args.Result = new Tuple<IReadOnlyList<StepItem>, IReadOnlyList<PluginTypeItem>, IReadOnlyDictionary<Guid, int>>(steps, targetTypesSnapshot, existingStepCounts);
            },
            PostWorkCallBack = args =>
            {
                if (args.Error != null)
                {
                    AppendOutput($"Auto map failed: {args.Error.Message}");
                    LogError(args.Error.ToString());
                    return;
                }

                var result = (Tuple<IReadOnlyList<StepItem>, IReadOnlyList<PluginTypeItem>, IReadOnlyDictionary<Guid, int>>)args.Result;
                var steps = result.Item1;

                ApplyAutoMatchToSteps(steps, result.Item2, result.Item3);
                PopulateGrid(steps);

                var matched = steps.Count(s => s.SuggestedTargetPluginTypeId.HasValue && s.SuggestedTargetPluginTypeId.Value != Guid.Empty);
                var unmatched = steps.Count - matched;
                var withWarnings = steps.Count(s => !string.IsNullOrWhiteSpace(s.MatchWarning));
                AppendOutput($"Auto mapped {steps.Count} steps from {sourceAssembly.DisplayName}. matched={matched}, unmatched={unmatched}, warnings={withWarnings}");
                EnableActions();
            }
        });
    }

    private void LoadSourceSteps()
    {
        var sourceAssembly = _sourceAssemblyCombo.SelectedItem as PluginAssemblyItem;
        var source = _sourceTypeCombo.SelectedItem as PluginTypeItem;

        if (sourceAssembly == null)
        {
            AppendOutput("Select a source assembly first.");
            return;
        }

        if (source == null)
        {
            AppendOutput("Select a source plugin type first.");
            return;
        }

        var sourceTypeId = source.PluginTypeId;
        var sourceTypeDisplay = source.DisplayName;

        WorkAsync(new WorkAsyncInfo
        {
            Message = "Loading source plugin steps...",
            Work = (worker, args) =>
            {
                args.Result = _stepMoveService.GetSteps(Service, sourceTypeId);
            },
            PostWorkCallBack = args =>
            {
                if (args.Error != null)
                {
                    AppendOutput($"Load failed: {args.Error.Message}");
                    LogError(args.Error.ToString());
                    return;
                }

                var steps = (IReadOnlyList<StepItem>)args.Result;
                PopulateGrid(steps);
                AppendOutput($"Loaded {steps.Count} steps for source plugin type: {sourceTypeDisplay}");
                EnableActions();
            }
        });
    }

    // --- Apply Operation ---

    private void ApplyOperation()
    {
        var selectedSteps = _steps.Where(s => s.Selected).ToList();
        var addToSolution = _addToSolutionCheck.Checked;
        var solution = _solutionCombo.SelectedItem as SolutionItem;
        var isMove = _moveRadio.Checked;
        var isCrossEnv = IsCrossEnvironment;
        var keepSameIds = _keepSameIdsCheck.Checked;

        if (!selectedSteps.Any())
        {
            AppendOutput("Select at least one step before applying.");
            return;
        }

        var stepsWithoutTarget = selectedSteps.Where(s => !s.SuggestedTargetPluginTypeId.HasValue || s.SuggestedTargetPluginTypeId == Guid.Empty).ToList();
        if (stepsWithoutTarget.Any())
        {
            AppendOutput($"{stepsWithoutTarget.Count} selected step(s) have no target assigned. Use 'Assign to Selected' to assign them or deselect them.");
            return;
        }

        if (addToSolution && solution == null)
        {
            AppendOutput("Select a solution first or uncheck 'Add to solution'.");
            return;
        }

        var analysis = _stepMoveService.AnalyzeAutoMatched(selectedSteps).ToList();

        var movable = analysis.Count(a => a.CanMove);
        var blocked = analysis.Count - movable;

        if (movable == 0)
        {
            var blockReport = new StringBuilder();
            blockReport.AppendLine($"No steps can be processed. blocked={blocked}");

            foreach (var line in analysis.Where(a => !a.CanMove).Select(a => $"- {a.StepName}: {a.Reason}"))
            {
                blockReport.AppendLine(line);
            }

            AppendOutput(blockReport.ToString().TrimEnd());
            return;
        }

        // --- Confirmation dialogs ---
        var operationLabel = isMove ? "move" : "copy";
        string confirmMessage;

        if (isMove && isCrossEnv)
        {
            confirmMessage = $"WARNING: {movable} step(s) will be CREATED on the target environment and DELETED from the source environment.\n\nThis cannot be undone. Continue?";
        }
        else if (isMove)
        {
            confirmMessage = $"This will reassign {movable} step(s) to a different plugin type.\n\nSteps will no longer appear under the source plugin type. Continue?";
        }
        else if (isCrossEnv)
        {
            confirmMessage = $"{movable} step(s) will be created on the target environment.\n\nSource steps remain unchanged. Continue?";
        }
        else
        {
            confirmMessage = $"{movable} new step(s) will be created on the same environment.\n\nSource steps remain unchanged. Continue?";
        }

        var dialogIcon = isMove && isCrossEnv ? MessageBoxIcon.Warning : MessageBoxIcon.Question;
        var dialogResult = MessageBox.Show(confirmMessage, $"Confirm {operationLabel}", MessageBoxButtons.YesNo, dialogIcon);

        if (dialogResult != DialogResult.Yes)
        {
            AppendOutput($"{operationLabel} cancelled by user.");
            return;
        }

        var solutionUniqueName = solution?.UniqueName;
        var targetService = _targetService ?? Service;

        var context = new OperationContext(Service, targetService)
        {
            Mode = isMove ? OperationMode.Move : OperationMode.Copy,
            IsCrossEnvironment = isCrossEnv,
            KeepSameIds = keepSameIds
        };

        WorkAsync(new WorkAsyncInfo
        {
            Message = $"Applying {operationLabel}...",
            Work = (worker, args) =>
            {
                var moveResults = _stepMoveService.Execute(context, analysis);
                IReadOnlyList<StepMoveResult>? solutionResults = null;

                if (addToSolution && !string.IsNullOrWhiteSpace(solutionUniqueName) && moveResults.Any(r => r.Success))
                {
                    solutionResults = _stepMoveService.AddStepsToSolution(targetService, solutionUniqueName!, moveResults);
                }

                args.Result = new Tuple<IReadOnlyList<StepMoveResult>, IReadOnlyList<StepMoveResult>?>(moveResults, solutionResults);
            },
            PostWorkCallBack = args =>
            {
                if (args.Error != null)
                {
                    AppendOutput($"Apply failed: {args.Error.Message}");
                    LogError(args.Error.ToString());
                    return;
                }

                var result = (Tuple<IReadOnlyList<StepMoveResult>, IReadOnlyList<StepMoveResult>?>)args.Result;
                var moveResults = result.Item1;
                var solutionResults = result.Item2;

                var success = moveResults.Count(r => r.Success);
                var failed = moveResults.Count - success;

                var report = new StringBuilder();
                report.AppendLine($"{operationLabel} result: success={success}, failed={failed}");

                foreach (var failure in moveResults.Where(r => !r.Success))
                {
                    report.AppendLine($"- {failure.StepName}: {failure.Message}");
                }

                if (solutionResults != null)
                {
                    var solSuccess = solutionResults.Count(r => r.Success);
                    var solFailed = solutionResults.Count - solSuccess;
                    report.AppendLine($"Add to solution '{solutionUniqueName}': success={solSuccess}, failed={solFailed}");

                    foreach (var failure in solutionResults.Where(r => !r.Success))
                    {
                        report.AppendLine($"- {failure.StepName}: {failure.Message}");
                    }
                }

                AppendOutput(report.ToString().TrimEnd());
                PopulateGrid(Array.Empty<StepItem>());
            }
        });
    }

    // --- Solutions ---

    private void LoadSolutions(IOrganizationService? service = null)
    {
        service ??= _targetService ?? Service;
        if (service == null)
        {
            AppendOutput("No Dataverse connection available.");
            return;
        }

        var svc = service;

        WorkAsync(new WorkAsyncInfo
        {
            Message = "Loading unmanaged solutions...",
            Work = (worker, args) =>
            {
                args.Result = _stepMoveService.GetUnmanagedSolutions(svc);
            },
            PostWorkCallBack = args =>
            {
                if (args.Error != null)
                {
                    AppendOutput($"Load solutions failed: {args.Error.Message}");
                    LogError(args.Error.ToString());
                    return;
                }

                var solutions = (IReadOnlyList<SolutionItem>)args.Result;

                _solutions.RaiseListChangedEvents = false;
                _solutions.Clear();

                foreach (var sol in solutions)
                {
                    _solutions.Add(sol);
                }

                _solutions.RaiseListChangedEvents = true;
                _solutions.ResetBindings();

                _solutionCombo.DataSource = _solutions;
                _solutionCombo.DisplayMember = nameof(SolutionItem.DisplayName);
                _solutionCombo.ValueMember = nameof(SolutionItem.SolutionId);

                _solutionsLoaded = true;
                AppendOutput($"Loaded {solutions.Count} unmanaged solutions.");
            }
        });
    }

    // --- Grid & Combo Population ---

    private void PopulateGrid(IReadOnlyList<StepItem> items)
    {
        _steps.RaiseListChangedEvents = false;
        _steps.Clear();

        foreach (var item in items)
        {
            _steps.Add(item);
        }

        _steps.RaiseListChangedEvents = true;
        _steps.ResetBindings();
        UpdateSelectionCount();
    }

    private void PopulatePluginData(IReadOnlyList<PluginAssemblyItem> assemblies, IReadOnlyList<PluginTypeItem> types)
    {
        _isBindingPluginData = true;

        try
        {
            _allPluginTypes.Clear();
            _allPluginTypes.AddRange(types);

            _sourceAssemblies.RaiseListChangedEvents = false;
            _sourceAssemblies.Clear();

            foreach (var assembly in assemblies)
            {
                _sourceAssemblies.Add(assembly);
            }

            _sourceAssemblies.RaiseListChangedEvents = true;
            _sourceAssemblies.ResetBindings();

            _sourceAssemblyCombo.DataSource = _sourceAssemblies;
            _sourceAssemblyCombo.DisplayMember = nameof(PluginAssemblyItem.DisplayName);
            _sourceAssemblyCombo.ValueMember = nameof(PluginAssemblyItem.AssemblyId);

            if (_sourceAssemblies.Count > 0 && _sourceAssemblyCombo.SelectedIndex < 0)
            {
                TrySetSelectedIndex(_sourceAssemblyCombo, 0);
            }

            FilterSourcePluginTypes();

            // Only populate target from source if same environment
            if (!IsCrossEnvironment)
            {
                _targetAssemblies.RaiseListChangedEvents = false;
                _targetAssemblies.Clear();

                foreach (var assembly in assemblies)
                {
                    _targetAssemblies.Add(new PluginAssemblyItem
                    {
                        AssemblyId = assembly.AssemblyId,
                        Name = assembly.Name
                    });
                }

                _targetAssemblies.RaiseListChangedEvents = true;
                _targetAssemblies.ResetBindings();

                _targetAssemblyCombo.DataSource = _targetAssemblies;
                _targetAssemblyCombo.DisplayMember = nameof(PluginAssemblyItem.DisplayName);
                _targetAssemblyCombo.ValueMember = nameof(PluginAssemblyItem.AssemblyId);

                if (_targetAssemblies.Count > 1)
                {
                    TrySetSelectedIndex(_targetAssemblyCombo, 1);
                }
                else if (_targetAssemblies.Count == 1)
                {
                    TrySetSelectedIndex(_targetAssemblyCombo, 0);
                }

                FilterTargetPluginTypes();
            }
        }
        finally
        {
            _isBindingPluginData = false;
        }

        if (!IsCrossEnvironment)
        {
            TryAutoSelectTargetMainFromSource();
            TryAutoSelectTargetFromSource();
        }

        EnableActions();
    }

    private void FilterSourcePluginTypes()
    {
        var selectedAssembly = _sourceAssemblyCombo.SelectedItem as PluginAssemblyItem;
        var filtered = selectedAssembly == null
            ? new List<PluginTypeItem>()
            : _allPluginTypes.Where(t => t.AssemblyId == selectedAssembly.AssemblyId).OrderBy(t => t.Name).ToList();

        _sourcePluginTypes.RaiseListChangedEvents = false;
        _sourcePluginTypes.Clear();

        foreach (var item in filtered)
        {
            _sourcePluginTypes.Add(item);
        }

        _sourcePluginTypes.RaiseListChangedEvents = true;
        _sourcePluginTypes.ResetBindings();

        _sourceTypeCombo.DataSource = _sourcePluginTypes;
        _sourceTypeCombo.DisplayMember = nameof(PluginTypeItem.DisplayName);
        _sourceTypeCombo.ValueMember = nameof(PluginTypeItem.PluginTypeId);

        if (_sourcePluginTypes.Count > 0)
        {
            TrySetSelectedIndex(_sourceTypeCombo, 0);
        }
    }

    private void FilterTargetPluginTypes()
    {
        var selectedAssembly = _targetAssemblyCombo.SelectedItem as PluginAssemblyItem;

        // Use target-specific types when cross-env, otherwise use source types
        var typeSource = IsCrossEnvironment && _allTargetPluginTypes.Count > 0
            ? _allTargetPluginTypes
            : _allPluginTypes;

        var filtered = selectedAssembly == null
            ? new List<PluginTypeItem>()
            : typeSource.Where(t => t.AssemblyId == selectedAssembly.AssemblyId).OrderBy(t => t.Name).ToList();

        _targetPluginTypes.RaiseListChangedEvents = false;
        _targetPluginTypes.Clear();

        foreach (var item in filtered)
        {
            _targetPluginTypes.Add(item);
        }

        _targetPluginTypes.RaiseListChangedEvents = true;
        _targetPluginTypes.ResetBindings();

        _targetTypeCombo.DataSource = _targetPluginTypes;
        _targetTypeCombo.DisplayMember = nameof(PluginTypeItem.DisplayName);
        _targetTypeCombo.ValueMember = nameof(PluginTypeItem.PluginTypeId);

        if (_targetPluginTypes.Count > 0)
        {
            TrySetSelectedIndex(_targetTypeCombo, 0);
        }
    }

    // --- Auto-selection helpers ---

    private void TryAutoSelectTargetFromSource()
    {
        var sourceType = _sourceTypeCombo.SelectedItem as PluginTypeItem;
        if (sourceType == null || _targetPluginTypes.Count == 0)
        {
            return;
        }

        var suggested = GetParentByLastDot(sourceType.Name);
        if (string.IsNullOrWhiteSpace(suggested))
        {
            return;
        }

        var target = _targetPluginTypes.FirstOrDefault(t =>
            string.Equals(t.Name, suggested, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(t.Name, sourceType.Name, StringComparison.OrdinalIgnoreCase));

        if (target == null)
        {
            return;
        }

        for (var index = 0; index < _targetPluginTypes.Count; index++)
        {
            if (_targetPluginTypes[index].PluginTypeId == target.PluginTypeId)
            {
                TrySetSelectedIndex(_targetTypeCombo, index);
                break;
            }
        }
    }

    private void TryAutoSelectTargetMainFromSource()
    {
        var sourceAssembly = _sourceAssemblyCombo.SelectedItem as PluginAssemblyItem;
        if (sourceAssembly == null || _targetAssemblies.Count == 0)
        {
            return;
        }

        var suggested = GetParentByLastDot(sourceAssembly.Name);
        if (string.IsNullOrWhiteSpace(suggested))
        {
            return;
        }

        var targetAssembly = _targetAssemblies.FirstOrDefault(a =>
            string.Equals(a.Name, suggested, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(a.Name, sourceAssembly.Name, StringComparison.OrdinalIgnoreCase));

        if (targetAssembly == null)
        {
            return;
        }

        for (var index = 0; index < _targetAssemblies.Count; index++)
        {
            if (_targetAssemblies[index].AssemblyId == targetAssembly.AssemblyId)
            {
                TrySetSelectedIndex(_targetAssemblyCombo, index);
                break;
            }
        }
    }

    private static void TrySetSelectedIndex(ComboBox comboBox, int index)
    {
        if (index < 0 || comboBox.Items.Count <= index)
        {
            return;
        }

        if (comboBox.SelectedIndex == index)
        {
            return;
        }

        comboBox.SelectedIndex = index;
    }

    private static string GetParentByLastDot(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        var normalized = Regex.Replace(value.Trim(), @"\.+", ".");
        var lastDotIndex = normalized.LastIndexOf('.');
        if (lastDotIndex <= 0)
        {
            return string.Empty;
        }

        return normalized.Substring(0, lastDotIndex);
    }

    private void ApplyAutoMatchToSteps(IReadOnlyList<StepItem> steps, IReadOnlyList<PluginTypeItem> targetTypes, IReadOnlyDictionary<Guid, int> existingTargetStepCounts)
    {
        var targetTypesByName = targetTypes
            .GroupBy(t => t.Name, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

        foreach (var step in steps)
        {
            step.SuggestedTargetPluginTypeId = null;
            step.SuggestedTargetPluginTypeName = string.Empty;
            step.TargetTypeHasExistingSteps = false;
            step.TargetTypeExistingStepCount = 0;
            step.MatchWarning = string.Empty;

            var suggestedName = GetParentByLastDot(step.SourcePluginTypeName);
            PluginTypeItem? matchedType = null;

            if (!string.IsNullOrWhiteSpace(suggestedName) && targetTypesByName.TryGetValue(suggestedName, out var suggestedMatch))
            {
                matchedType = suggestedMatch;
            }
            else if (!string.IsNullOrWhiteSpace(step.SourcePluginTypeName) && targetTypesByName.TryGetValue(step.SourcePluginTypeName, out var exactMatch))
            {
                matchedType = exactMatch;
            }

            if (matchedType == null)
            {
                step.MatchWarning = "No matching target plugin type found";
                continue;
            }

            step.SuggestedTargetPluginTypeId = matchedType.PluginTypeId;
            step.SuggestedTargetPluginTypeName = matchedType.Name;

            if (existingTargetStepCounts.TryGetValue(matchedType.PluginTypeId, out var existingCount) && existingCount > 0)
            {
                step.TargetTypeHasExistingSteps = true;
                step.TargetTypeExistingStepCount = existingCount;
                step.MatchWarning = $"Target plugin type already has {existingCount} step(s)";
            }
        }
    }

    // --- Output & State ---

    private void AppendOutput(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return;
        }

        var line = $"[{DateTime.Now:HH:mm:ss}] {text}";

        if (_outputBox.TextLength > 0)
        {
            _outputBox.AppendText(Environment.NewLine);
        }

        _outputBox.AppendText(line);
        LogInfo(text);
    }

    private void EnableActions()
    {
        var hasConnection = Service != null;
        var hasSourceAssembly = _sourceAssemblyCombo.SelectedItem is PluginAssemblyItem;
        var hasTargetAssembly = _targetAssemblyCombo.SelectedItem is PluginAssemblyItem;
        var hasSource = _sourceTypeCombo.SelectedItem is PluginTypeItem;
        var hasTarget = _targetTypeCombo.SelectedItem is PluginTypeItem;
        var selectedCount = _steps.Count(s => s.Selected);
        var hasSteps = _steps.Count > 0;

        _loadPluginsButton.Enabled = hasConnection;
        _autoMapButton.Enabled = hasConnection && hasSourceAssembly && hasTargetAssembly;
        _loadButton.Enabled = hasConnection && hasSourceAssembly && hasSource;
        _applyButton.Enabled = hasConnection && selectedCount > 0;
        _setTargetButton.Enabled = hasTarget && selectedCount > 0;
        _connectTargetButton.Enabled = hasConnection;
        _selectAllButton.Enabled = hasSteps;
        _deselectAllButton.Enabled = hasSteps;
    }
}
