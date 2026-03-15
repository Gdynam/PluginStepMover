using System;
using Microsoft.Xrm.Sdk;

namespace PluginStepMover;

public enum OperationMode
{
    Move,
    Copy
}

public class OperationContext
{
    public OperationContext(IOrganizationService sourceService, IOrganizationService targetService)
    {
        SourceService = sourceService ?? throw new ArgumentNullException(nameof(sourceService));
        TargetService = targetService ?? throw new ArgumentNullException(nameof(targetService));
    }

    public OperationMode Mode { get; set; } = OperationMode.Move;
    public bool KeepSameIds { get; set; }
    public bool IsCrossEnvironment { get; set; }
    public IOrganizationService SourceService { get; }
    public IOrganizationService TargetService { get; }
}

public class StepItem
{
    public Guid StepId { get; set; }
    public bool Selected { get; set; }
    public string StepName { get; set; } = string.Empty;
    public string MessageName { get; set; } = string.Empty;
    public string PrimaryEntity { get; set; } = string.Empty;
    public int Stage { get; set; }
    public int Mode { get; set; }
    public int Rank { get; set; }
    public Guid SourcePluginTypeId { get; set; }
    public string SourcePluginTypeName { get; set; } = string.Empty;
    public string SourceAssemblyName { get; set; } = string.Empty;
    public Guid? SuggestedTargetPluginTypeId { get; set; }
    public string SuggestedTargetPluginTypeName { get; set; } = string.Empty;
    public bool TargetTypeHasExistingSteps { get; set; }
    public int TargetTypeExistingStepCount { get; set; }
    public string MatchWarning { get; set; } = string.Empty;
}

public class PluginTypeItem
{
    public Guid PluginTypeId { get; set; }
    public Guid AssemblyId { get; set; }
    public string Name { get; set; } = string.Empty;
    public string AssemblyName { get; set; } = string.Empty;

    public string DisplayName => $"{AssemblyName} :: {Name}";
}

public class PluginAssemblyItem
{
    public Guid AssemblyId { get; set; }
    public string Name { get; set; } = string.Empty;

    public string DisplayName => Name;
}

public class StepAnalyzeItem
{
    public Guid StepId { get; set; }
    public string StepName { get; set; } = string.Empty;
    public Guid SourcePluginTypeId { get; set; }
    public string SourcePluginTypeName { get; set; } = string.Empty;
    public Guid TargetPluginTypeId { get; set; }
    public string TargetPluginTypeName { get; set; } = string.Empty;
    public bool CanMove { get; set; }
    public bool HasWarning { get; set; }
    public string Reason { get; set; } = string.Empty;
}

public class StepMoveResult
{
    public Guid StepId { get; set; }
    public Guid? NewStepId { get; set; }
    public string StepName { get; set; } = string.Empty;
    public bool Success { get; set; }
    public string Message { get; set; } = string.Empty;
}

public class SolutionItem
{
    public Guid SolutionId { get; set; }
    public string UniqueName { get; set; } = string.Empty;
    public string FriendlyName { get; set; } = string.Empty;

    public string DisplayName => $"{FriendlyName} ({UniqueName})";
}
