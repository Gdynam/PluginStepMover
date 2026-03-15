using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using PluginStepMover.Models;

namespace PluginStepMover;

using StepFields = SdkMessageProcessingStep.Fields;
using ImageFields = SdkMessageProcessingStepImage.Fields;
using TypeFields = PluginType.Fields;
using AssemblyFields = PluginAssembly.Fields;
using MessageFields = SdkMessage.Fields;
using FilterFields = SdkMessageFilter.Fields;
using SolutionFields = Solution.Fields;

public class StepMoveService
{
    private const int SdkMessageProcessingStepComponentType = 92;

    public IReadOnlyList<PluginAssemblyItem> GetPluginAssemblies(IOrganizationService service)
    {
        var query = new QueryExpression(PluginAssembly.EntityLogicalName)
        {
            ColumnSet = new ColumnSet(AssemblyFields.PluginAssemblyId, AssemblyFields.Name)
        };

        return RetrieveAll(service, query)
            .Select(entity => new PluginAssemblyItem
            {
                AssemblyId = entity.Id,
                Name = entity.GetAttributeValue<string>(AssemblyFields.Name) ?? string.Empty
            })
            .Where(item => item.AssemblyId != Guid.Empty && !string.IsNullOrWhiteSpace(item.Name))
            .OrderBy(item => item.Name)
            .ToList();
    }

    public IReadOnlyList<StepItem> GetSteps(IOrganizationService service, Guid? sourcePluginTypeId = null, Guid? sourceAssemblyId = null)
    {
        var query = new QueryExpression(SdkMessageProcessingStep.EntityLogicalName)
        {
            ColumnSet = new ColumnSet(
                StepFields.SdkMessageProcessingStepId,
                StepFields.Name,
                StepFields.Stage,
                StepFields.Mode,
                StepFields.Rank,
                StepFields.PluginTypeId)
        };

        if (sourcePluginTypeId.HasValue && sourcePluginTypeId.Value != Guid.Empty)
        {
            query.Criteria.AddCondition(StepFields.PluginTypeId, ConditionOperator.Equal, sourcePluginTypeId.Value);
        }

        var messageLink = query.AddLink(SdkMessage.EntityLogicalName, StepFields.SdkMessageId, MessageFields.SdkMessageId, JoinOperator.LeftOuter);
        messageLink.Columns = new ColumnSet(MessageFields.Name);
        messageLink.EntityAlias = "message";

        var filterLink = query.AddLink(SdkMessageFilter.EntityLogicalName, StepFields.SdkMessageFilterId, FilterFields.SdkMessageFilterId, JoinOperator.LeftOuter);
        filterLink.Columns = new ColumnSet(FilterFields.PrimaryObjectTypeCode);
        filterLink.EntityAlias = "filter";

        var typeJoin = sourceAssemblyId.HasValue && sourceAssemblyId.Value != Guid.Empty
            ? JoinOperator.Inner
            : JoinOperator.LeftOuter;

        var typeLink = query.AddLink(PluginType.EntityLogicalName, StepFields.PluginTypeId, TypeFields.PluginTypeId, typeJoin);
        typeLink.Columns = new ColumnSet(TypeFields.TypeName, TypeFields.Name, TypeFields.PluginAssemblyId);
        typeLink.EntityAlias = "ptype";

        if (sourceAssemblyId.HasValue && sourceAssemblyId.Value != Guid.Empty)
        {
            typeLink.LinkCriteria.AddCondition(TypeFields.PluginAssemblyId, ConditionOperator.Equal, sourceAssemblyId.Value);
        }

        var assemblyLink = typeLink.AddLink(PluginAssembly.EntityLogicalName, TypeFields.PluginAssemblyId, AssemblyFields.PluginAssemblyId, JoinOperator.LeftOuter);
        assemblyLink.Columns = new ColumnSet(AssemblyFields.Name);
        assemblyLink.EntityAlias = "assembly";

        var records = RetrieveAll(service, query);

        return records
            .Select(MapStep)
            .Where(s => s.SourcePluginTypeId != Guid.Empty)
            .OrderBy(s => s.SourceAssemblyName)
            .ThenBy(s => s.SourcePluginTypeName)
            .ThenBy(s => s.StepName)
            .ToList();
    }

    public IReadOnlyDictionary<Guid, int> GetStepCountsByPluginTypeIds(IOrganizationService service, IReadOnlyCollection<Guid> pluginTypeIds)
    {
        if (pluginTypeIds.Count == 0)
        {
            return new Dictionary<Guid, int>();
        }

        var query = new QueryExpression(SdkMessageProcessingStep.EntityLogicalName)
        {
            ColumnSet = new ColumnSet(StepFields.PluginTypeId)
        };

        query.Criteria.AddCondition(StepFields.PluginTypeId, ConditionOperator.In, pluginTypeIds.Cast<object>().ToArray());

        var counts = new Dictionary<Guid, int>();

        foreach (var step in RetrieveAll(service, query))
        {
            var typeId = step.GetAttributeValue<EntityReference>(StepFields.PluginTypeId)?.Id ?? Guid.Empty;
            if (typeId == Guid.Empty)
            {
                continue;
            }

            counts[typeId] = counts.TryGetValue(typeId, out var current) ? current + 1 : 1;
        }

        return counts;
    }

    public IReadOnlyList<PluginTypeItem> GetPluginTypes(IOrganizationService service, Guid? assemblyId = null)
    {
        var query = new QueryExpression(PluginType.EntityLogicalName)
        {
            ColumnSet = new ColumnSet(TypeFields.PluginTypeId, TypeFields.TypeName, TypeFields.Name, TypeFields.PluginAssemblyId)
        };

        if (assemblyId.HasValue && assemblyId.Value != Guid.Empty)
        {
            query.Criteria.AddCondition(TypeFields.PluginAssemblyId, ConditionOperator.Equal, assemblyId.Value);
        }

        var assemblyLink = query.AddLink(PluginAssembly.EntityLogicalName, TypeFields.PluginAssemblyId, AssemblyFields.PluginAssemblyId, JoinOperator.LeftOuter);
        assemblyLink.Columns = new ColumnSet(AssemblyFields.Name);
        assemblyLink.EntityAlias = "assembly";

        return RetrieveAll(service, query)
            .Select(entity => new PluginTypeItem
            {
                PluginTypeId = entity.Id,
                AssemblyId = entity.GetAttributeValue<EntityReference>(TypeFields.PluginAssemblyId)?.Id ?? Guid.Empty,
                Name = GetText(entity, "ptype.typename", TypeFields.TypeName, TypeFields.Name),
                AssemblyName = GetAliasedString(entity, "assembly." + AssemblyFields.Name)
            })
            .Where(item => item.PluginTypeId != Guid.Empty && !string.IsNullOrWhiteSpace(item.Name))
            .OrderBy(item => item.AssemblyName)
            .ThenBy(item => item.Name)
            .ToList();
    }

    public IReadOnlyList<StepAnalyzeItem> AnalyzeAutoMatched(IReadOnlyCollection<StepItem> selectedSteps)
    {
        var result = new List<StepAnalyzeItem>();

        foreach (var step in selectedSteps)
        {
            var hasTarget = step.SuggestedTargetPluginTypeId.HasValue && step.SuggestedTargetPluginTypeId.Value != Guid.Empty;
            var hasWarning = !string.IsNullOrWhiteSpace(step.MatchWarning);
            var canMove = hasTarget;
            var reason = hasWarning ? step.MatchWarning : "Ready";

            if (!hasTarget)
            {
                reason = string.IsNullOrWhiteSpace(step.MatchWarning)
                    ? "No matching target plugin type found"
                    : step.MatchWarning;
            }

            if (hasTarget && step.SourcePluginTypeId == step.SuggestedTargetPluginTypeId)
            {
                canMove = false;
                reason = "Source and target are identical";
            }

            result.Add(new StepAnalyzeItem
            {
                StepId = step.StepId,
                StepName = step.StepName,
                SourcePluginTypeId = step.SourcePluginTypeId,
                SourcePluginTypeName = step.SourcePluginTypeName,
                TargetPluginTypeId = step.SuggestedTargetPluginTypeId ?? Guid.Empty,
                TargetPluginTypeName = step.SuggestedTargetPluginTypeName,
                CanMove = canMove,
                HasWarning = hasWarning,
                Reason = reason
            });
        }

        return result;
    }

    public IReadOnlyList<StepMoveResult> Execute(OperationContext context, IReadOnlyCollection<StepAnalyzeItem> analysis)
    {
        var movableItems = analysis.Where(a => a.CanMove).ToList();
        var skippedItems = analysis.Where(a => !a.CanMove).ToList();

        var results = new List<StepMoveResult>();

        // Add skipped items to results
        foreach (var item in skippedItems)
        {
            results.Add(new StepMoveResult
            {
                StepId = item.StepId,
                StepName = item.StepName,
                Success = false,
                Message = item.Reason
            });
        }

        if (!movableItems.Any())
        {
            return results;
        }

        // For same-env move: use simple update (existing behavior)
        if (context.Mode == OperationMode.Move && !context.IsCrossEnvironment)
        {
            results.AddRange(ExecuteSameEnvMove(context.TargetService, movableItems));
            return results;
        }

        // For copy or cross-env: retrieve full step entities first
        var stepIds = movableItems.Select(i => i.StepId).ToList();
        var fullSteps = RetrieveFullSteps(context.SourceService, stepIds);

        // Build caches for cross-env message/filter resolution
        var messageCache = new Dictionary<string, Guid?>(StringComparer.OrdinalIgnoreCase);
        var filterCache = new Dictionary<string, Guid?>(StringComparer.OrdinalIgnoreCase);

        foreach (var item in movableItems)
        {
            var fullStep = fullSteps.FirstOrDefault(e => e.Id == item.StepId);
            if (fullStep == null)
            {
                results.Add(new StepMoveResult
                {
                    StepId = item.StepId,
                    StepName = item.StepName,
                    Success = false,
                    Message = "Could not retrieve full step data"
                });
                continue;
            }

            try
            {
                // Resolve message/filter IDs for cross-env, or reuse for same-env
                Guid? resolvedMessageId = null;
                Guid? resolvedFilterId = null;

                if (context.IsCrossEnvironment)
                {
                    var messageName = GetMessageName(context.SourceService, fullStep);
                    if (!string.IsNullOrWhiteSpace(messageName))
                    {
                        if (!messageCache.TryGetValue(messageName, out resolvedMessageId))
                        {
                            resolvedMessageId = ResolveSdkMessageId(context.TargetService, messageName);
                            messageCache[messageName] = resolvedMessageId;
                        }
                    }

                    if (resolvedMessageId == null)
                    {
                        results.Add(new StepMoveResult
                        {
                            StepId = item.StepId,
                            StepName = item.StepName,
                            Success = false,
                            Message = $"Could not resolve SDK message '{messageName}' on target environment"
                        });
                        continue;
                    }

                    var filterRef = fullStep.GetAttributeValue<EntityReference>(StepFields.SdkMessageFilterId);
                    if (filterRef != null)
                    {
                        var primaryEntity = GetPrimaryEntity(context.SourceService, filterRef.Id);
                        var filterKey = $"{resolvedMessageId}|{primaryEntity}";

                        if (!filterCache.TryGetValue(filterKey, out resolvedFilterId))
                        {
                            resolvedFilterId = ResolveSdkMessageFilterId(context.TargetService, resolvedMessageId.Value, primaryEntity);
                            filterCache[filterKey] = resolvedFilterId;
                        }
                    }
                }

                var newId = CreateStep(
                    context.TargetService,
                    fullStep,
                    item.TargetPluginTypeId,
                    context.KeepSameIds,
                    context.IsCrossEnvironment,
                    resolvedMessageId,
                    resolvedFilterId);

                // Copy step images to the new step
                var imageCount = CopyStepImages(context.SourceService, context.TargetService, item.StepId, newId, context.IsCrossEnvironment);

                var operationLabel = context.Mode == OperationMode.Copy ? "Copied" : "Created on target";
                if (imageCount > 0)
                {
                    operationLabel += $" ({imageCount} image{(imageCount == 1 ? "" : "s")})";
                }

                results.Add(new StepMoveResult
                {
                    StepId = item.StepId,
                    NewStepId = newId,
                    StepName = item.StepName,
                    Success = true,
                    Message = operationLabel
                });

                // For cross-env move: delete from source after successful create
                if (context.Mode == OperationMode.Move && context.IsCrossEnvironment)
                {
                    try
                    {
                        DeleteStep(context.SourceService, item.StepId);
                        results[results.Count - 1].Message = "Moved to target (deleted from source)";
                    }
                    catch (Exception deleteEx)
                    {
                        results[results.Count - 1].Message = $"Created on target but FAILED to delete from source: {deleteEx.Message}";
                    }
                }
            }
            catch (Exception ex)
            {
                results.Add(new StepMoveResult
                {
                    StepId = item.StepId,
                    StepName = item.StepName,
                    Success = false,
                    Message = ex.Message
                });
            }
        }

        return results;
    }

    public IReadOnlyList<SolutionItem> GetUnmanagedSolutions(IOrganizationService service)
    {
        var query = new QueryExpression(Solution.EntityLogicalName)
        {
            ColumnSet = new ColumnSet(SolutionFields.SolutionId, SolutionFields.UniqueName, SolutionFields.FriendlyName)
        };

        query.Criteria.AddCondition(SolutionFields.IsManaged, ConditionOperator.Equal, false);
        query.Criteria.AddCondition(SolutionFields.IsVisible, ConditionOperator.Equal, true);
        query.Criteria.AddCondition(SolutionFields.UniqueName, ConditionOperator.NotEqual, "Default");

        return RetrieveAll(service, query)
            .Select(entity => new SolutionItem
            {
                SolutionId = entity.Id,
                UniqueName = entity.GetAttributeValue<string>(SolutionFields.UniqueName) ?? string.Empty,
                FriendlyName = entity.GetAttributeValue<string>(SolutionFields.FriendlyName) ?? string.Empty
            })
            .Where(s => !string.IsNullOrWhiteSpace(s.UniqueName))
            .OrderBy(s => s.FriendlyName)
            .ToList();
    }

    public IReadOnlyList<StepMoveResult> AddStepsToSolution(IOrganizationService service, string solutionUniqueName, IReadOnlyCollection<StepMoveResult> movedSteps)
    {
        var results = new List<StepMoveResult>();

        foreach (var step in movedSteps.Where(s => s.Success))
        {
            try
            {
                // Use NewStepId if available (copy/cross-env), otherwise use StepId (same-env move)
                var stepId = step.NewStepId ?? step.StepId;

                var request = new OrganizationRequest("AddSolutionComponent")
                {
                    ["ComponentId"] = stepId,
                    ["ComponentType"] = SdkMessageProcessingStepComponentType,
                    ["SolutionUniqueName"] = solutionUniqueName,
                    ["AddRequiredComponents"] = false
                };

                service.Execute(request);

                results.Add(new StepMoveResult
                {
                    StepId = stepId,
                    StepName = step.StepName,
                    Success = true,
                    Message = "Added to solution"
                });
            }
            catch (Exception ex)
            {
                results.Add(new StepMoveResult
                {
                    StepId = step.NewStepId ?? step.StepId,
                    StepName = step.StepName,
                    Success = false,
                    Message = ex.Message
                });
            }
        }

        return results;
    }

    // --- Private helpers for copy/cross-env ---

    private IReadOnlyList<StepMoveResult> ExecuteSameEnvMove(IOrganizationService service, IReadOnlyList<StepAnalyzeItem> items)
    {
        var results = new List<StepMoveResult>();

        foreach (var item in items)
        {
            try
            {
                var update = new Entity(SdkMessageProcessingStep.EntityLogicalName, item.StepId)
                {
                    [StepFields.PluginTypeId] = new EntityReference(PluginType.EntityLogicalName, item.TargetPluginTypeId)
                };

                service.Update(update);

                results.Add(new StepMoveResult
                {
                    StepId = item.StepId,
                    StepName = item.StepName,
                    Success = true,
                    Message = "Moved"
                });
            }
            catch (Exception ex)
            {
                results.Add(new StepMoveResult
                {
                    StepId = item.StepId,
                    StepName = item.StepName,
                    Success = false,
                    Message = ex.Message
                });
            }
        }

        return results;
    }

    private IReadOnlyList<Entity> RetrieveFullSteps(IOrganizationService service, IReadOnlyList<Guid> stepIds)
    {
        if (stepIds.Count == 0)
        {
            return Array.Empty<Entity>();
        }

        var query = new QueryExpression(SdkMessageProcessingStep.EntityLogicalName)
        {
            ColumnSet = new ColumnSet(true)
        };

        query.Criteria.AddCondition(StepFields.SdkMessageProcessingStepId, ConditionOperator.In, stepIds.Cast<object>().ToArray());

        return RetrieveAll(service, query);
    }

    private Guid? ResolveSdkMessageId(IOrganizationService targetService, string messageName)
    {
        var query = new QueryExpression(SdkMessage.EntityLogicalName)
        {
            ColumnSet = new ColumnSet(MessageFields.SdkMessageId),
            TopCount = 1
        };

        query.Criteria.AddCondition(MessageFields.Name, ConditionOperator.Equal, messageName);

        var result = targetService.RetrieveMultiple(query).Entities.FirstOrDefault();
        return result?.Id;
    }

    private Guid? ResolveSdkMessageFilterId(IOrganizationService targetService, Guid messageId, string primaryEntity)
    {
        if (string.IsNullOrWhiteSpace(primaryEntity))
        {
            return null;
        }

        var query = new QueryExpression(SdkMessageFilter.EntityLogicalName)
        {
            ColumnSet = new ColumnSet(FilterFields.SdkMessageFilterId),
            TopCount = 1
        };

        query.Criteria.AddCondition(FilterFields.SdkMessageId, ConditionOperator.Equal, messageId);
        query.Criteria.AddCondition(FilterFields.PrimaryObjectTypeCode, ConditionOperator.Equal, primaryEntity);

        var result = targetService.RetrieveMultiple(query).Entities.FirstOrDefault();
        return result?.Id;
    }

    private Guid CreateStep(
        IOrganizationService targetService,
        Entity sourceStep,
        Guid targetPluginTypeId,
        bool keepSameId,
        bool isCrossEnv,
        Guid? resolvedMessageId,
        Guid? resolvedFilterId)
    {
        var newStep = keepSameId
            ? new Entity(SdkMessageProcessingStep.EntityLogicalName, sourceStep.Id)
            : new Entity(SdkMessageProcessingStep.EntityLogicalName);

        // Set the target plugin type
        newStep[StepFields.PluginTypeId] = new EntityReference(PluginType.EntityLogicalName, targetPluginTypeId);

        // Set message reference
        if (isCrossEnv && resolvedMessageId.HasValue)
        {
            newStep[StepFields.SdkMessageId] = new EntityReference(SdkMessage.EntityLogicalName, resolvedMessageId.Value);
        }
        else if (sourceStep.Contains(StepFields.SdkMessageId))
        {
            newStep[StepFields.SdkMessageId] = sourceStep[StepFields.SdkMessageId];
        }

        // Set filter reference
        if (isCrossEnv && resolvedFilterId.HasValue)
        {
            newStep[StepFields.SdkMessageFilterId] = new EntityReference(SdkMessageFilter.EntityLogicalName, resolvedFilterId.Value);
        }
        else if (!isCrossEnv && sourceStep.Contains(StepFields.SdkMessageFilterId))
        {
            newStep[StepFields.SdkMessageFilterId] = sourceStep[StepFields.SdkMessageFilterId];
        }

        // Copy standard attributes
        CopyAttribute(sourceStep, newStep, StepFields.Name);
        CopyAttribute(sourceStep, newStep, StepFields.Stage);
        CopyAttribute(sourceStep, newStep, StepFields.Mode);
        CopyAttribute(sourceStep, newStep, StepFields.Rank);
        CopyAttribute(sourceStep, newStep, StepFields.FilteringAttributes);
        CopyAttribute(sourceStep, newStep, StepFields.Configuration);
        CopyAttribute(sourceStep, newStep, StepFields.Description);
        CopyAttribute(sourceStep, newStep, StepFields.SupportedDeployment);
        CopyAttribute(sourceStep, newStep, StepFields.AsyncAutoDelete);

        // Skip impersonating user for cross-env (user may not exist on target)
        if (!isCrossEnv)
        {
            CopyAttribute(sourceStep, newStep, StepFields.ImpersonatingUserId);
        }

        return targetService.Create(newStep);
    }

    private void DeleteStep(IOrganizationService sourceService, Guid stepId)
    {
        sourceService.Delete(SdkMessageProcessingStep.EntityLogicalName, stepId);
    }

    private string GetMessageName(IOrganizationService service, Entity step)
    {
        var messageRef = step.GetAttributeValue<EntityReference>(StepFields.SdkMessageId);
        if (messageRef == null)
        {
            return string.Empty;
        }

        var message = service.Retrieve(SdkMessage.EntityLogicalName, messageRef.Id, new ColumnSet(MessageFields.Name));
        return message.GetAttributeValue<string>(MessageFields.Name) ?? string.Empty;
    }

    private string GetPrimaryEntity(IOrganizationService service, Guid filterId)
    {
        var filter = service.Retrieve(SdkMessageFilter.EntityLogicalName, filterId, new ColumnSet(FilterFields.PrimaryObjectTypeCode));
        return filter.GetAttributeValue<string>(FilterFields.PrimaryObjectTypeCode) ?? string.Empty;
    }

    private static List<Entity> RetrieveAll(IOrganizationService service, QueryExpression query)
    {
        var results = new List<Entity>();
        query.PageInfo = new PagingInfo { PageNumber = 1, Count = 5000 };

        while (true)
        {
            var response = service.RetrieveMultiple(query);
            results.AddRange(response.Entities);

            if (!response.MoreRecords)
            {
                break;
            }

            query.PageInfo.PageNumber++;
            query.PageInfo.PagingCookie = response.PagingCookie;
        }

        return results;
    }

    private static void CopyAttribute(Entity source, Entity target, string attributeName)
    {
        if (source.Contains(attributeName) && source[attributeName] != null)
        {
            target[attributeName] = source[attributeName];
        }
    }

    private static StepItem MapStep(Entity entity)
    {
        return new StepItem
        {
            StepId = entity.Id,
            StepName = entity.GetAttributeValue<string>(StepFields.Name) ?? string.Empty,
            MessageName = GetAliasedString(entity, "message." + MessageFields.Name),
            PrimaryEntity = GetAliasedString(entity, "filter." + FilterFields.PrimaryObjectTypeCode),
            Stage = entity.GetAttributeValue<OptionSetValue>(StepFields.Stage)?.Value ?? 0,
            Mode = entity.GetAttributeValue<OptionSetValue>(StepFields.Mode)?.Value ?? 0,
            Rank = entity.GetAttributeValue<int?>(StepFields.Rank) ?? 0,
            SourcePluginTypeId = entity.GetAttributeValue<EntityReference>(StepFields.PluginTypeId)?.Id ?? Guid.Empty,
            SourcePluginTypeName = GetText(entity, "ptype." + TypeFields.TypeName, "ptype." + TypeFields.Name),
            SourceAssemblyName = GetAliasedString(entity, "assembly." + AssemblyFields.Name)
        };
    }

    private static string GetText(Entity entity, params string[] keys)
    {
        foreach (var key in keys)
        {
            var value = key.Contains(".", StringComparison.Ordinal)
                ? GetAliasedString(entity, key)
                : entity.GetAttributeValue<string>(key);

            if (!string.IsNullOrWhiteSpace(value))
            {
                return value;
            }
        }

        return string.Empty;
    }

    private static string GetAliasedString(Entity entity, string key)
    {
        if (!entity.Attributes.TryGetValue(key, out var raw) || raw is not AliasedValue alias || alias.Value == null)
        {
            return string.Empty;
        }

        return alias.Value.ToString()!;
    }

    private static readonly string[] ImageCopyAttributes =
    {
        ImageFields.Name, ImageFields.EntityAlias, ImageFields.MessagePropertyName,
        ImageFields.ImageType, ImageFields.Attributes1, ImageFields.Description
    };

    private int CopyStepImages(IOrganizationService sourceService, IOrganizationService targetService, Guid sourceStepId, Guid newStepId, bool isCrossEnv)
    {
        var query = new QueryExpression(SdkMessageProcessingStepImage.EntityLogicalName)
        {
            ColumnSet = new ColumnSet(ImageCopyAttributes)
        };

        query.Criteria.AddCondition(ImageFields.SdkMessageProcessingStepId, ConditionOperator.Equal, sourceStepId);

        var images = RetrieveAll(sourceService, query);
        var count = 0;

        foreach (var image in images)
        {
            var newImage = new Entity(SdkMessageProcessingStepImage.EntityLogicalName)
            {
                [ImageFields.SdkMessageProcessingStepId] = new EntityReference(SdkMessageProcessingStep.EntityLogicalName, newStepId)
            };

            foreach (var attr in ImageCopyAttributes)
            {
                CopyAttribute(image, newImage, attr);
            }

            targetService.Create(newImage);
            count++;
        }

        return count;
    }
}
