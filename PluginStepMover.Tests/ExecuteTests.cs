using System;
using System.Collections.Generic;
using System.Linq;
using FluentAssertions;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using NSubstitute;
using NSubstitute.ExceptionExtensions;
using Xunit;

namespace PluginStepMover.Tests;

public class ExecuteTests
{
    private readonly StepMoveService _sut = new();
    private readonly IOrganizationService _service = Substitute.For<IOrganizationService>();

    [Fact]
    public void Execute_SkippedItems_AppearAsFailures()
    {
        var analysis = new[]
        {
            new StepAnalyzeItem
            {
                StepId = Guid.NewGuid(),
                StepName = "Skipped",
                CanMove = false,
                Reason = "Source and target are identical"
            }
        };

        var context = CreateSameEnvMoveContext(_service);

        var result = _sut.Execute(context, analysis);

        result.Should().ContainSingle()
            .Which.Should().Match<StepMoveResult>(r =>
                r.Success == false &&
                r.Message == "Source and target are identical");
    }

    [Fact]
    public void Execute_SameEnvMove_CallsUpdateWithCorrectPluginType()
    {
        var stepId = Guid.NewGuid();
        var targetTypeId = Guid.NewGuid();
        var analysis = new[] { CreateMovableItem(stepId, targetTypeId) };
        var context = CreateSameEnvMoveContext(_service);

        var result = _sut.Execute(context, analysis);

        result.Should().ContainSingle()
            .Which.Should().Match<StepMoveResult>(r => r.Success == true && r.Message == "Moved");

        _service.Received(1).Update(Arg.Is<Entity>(e =>
            e.LogicalName == "sdkmessageprocessingstep" &&
            e.Id == stepId &&
            ((EntityReference)e["plugintypeid"]).Id == targetTypeId));
    }

    [Fact]
    public void Execute_SameEnvMove_UpdateFails_ReportsErrorAndContinues()
    {
        var failId = Guid.NewGuid();
        var successId = Guid.NewGuid();
        var targetTypeId = Guid.NewGuid();

        _service.When(s => s.Update(Arg.Is<Entity>(e => e.Id == failId)))
            .Do(_ => throw new InvalidOperationException("Update failed"));

        var analysis = new[]
        {
            CreateMovableItem(failId, targetTypeId, "FailStep"),
            CreateMovableItem(successId, targetTypeId, "SuccessStep")
        };
        var context = CreateSameEnvMoveContext(_service);

        var result = _sut.Execute(context, analysis);

        result.Should().HaveCount(2);
        result.Should().ContainSingle(r => r.StepName == "FailStep" && !r.Success && r.Message == "Update failed");
        result.Should().ContainSingle(r => r.StepName == "SuccessStep" && r.Success);
    }

    [Fact]
    public void Execute_CopyMode_CreatesStepAndReturnsNewId()
    {
        var stepId = Guid.NewGuid();
        var targetTypeId = Guid.NewGuid();
        var newStepId = Guid.NewGuid();

        var sourceService = Substitute.For<IOrganizationService>();
        var targetService = Substitute.For<IOrganizationService>();

        SetupRetrieveFullSteps(sourceService, stepId);
        targetService.Create(Arg.Any<Entity>()).Returns(newStepId);

        var analysis = new[] { CreateMovableItem(stepId, targetTypeId) };
        var context = new OperationContext(sourceService, targetService)
        {
            Mode = OperationMode.Copy,
            IsCrossEnvironment = false,
            KeepSameIds = false
        };

        var result = _sut.Execute(context, analysis);

        result.Should().ContainSingle()
            .Which.Should().Match<StepMoveResult>(r =>
                r.Success == true &&
                r.NewStepId == newStepId);
    }

    [Fact]
    public void Execute_CrossEnvMove_CreatesAndDeletesStep()
    {
        var stepId = Guid.NewGuid();
        var targetTypeId = Guid.NewGuid();
        var newStepId = Guid.NewGuid();
        var messageId = Guid.NewGuid();

        var sourceService = Substitute.For<IOrganizationService>();
        var targetService = Substitute.For<IOrganizationService>();

        var sourceStep = CreateFullStepEntity(stepId, messageId);
        SetupRetrieveFullSteps(sourceService, sourceStep);
        SetupMessageResolve(sourceService, messageId, "Create");
        SetupTargetMessageLookup(targetService, messageId);

        targetService.Create(Arg.Any<Entity>()).Returns(newStepId);

        var analysis = new[] { CreateMovableItem(stepId, targetTypeId) };
        var context = new OperationContext(sourceService, targetService)
        {
            Mode = OperationMode.Move,
            IsCrossEnvironment = true,
            KeepSameIds = false
        };

        var result = _sut.Execute(context, analysis);

        result.Should().ContainSingle()
            .Which.Success.Should().BeTrue();

        sourceService.Received(1).Delete("sdkmessageprocessingstep", stepId);
    }

    [Fact]
    public void Execute_CrossEnvMove_DeleteFails_ReportsPartialSuccess()
    {
        var stepId = Guid.NewGuid();
        var targetTypeId = Guid.NewGuid();
        var newStepId = Guid.NewGuid();
        var messageId = Guid.NewGuid();

        var sourceService = Substitute.For<IOrganizationService>();
        var targetService = Substitute.For<IOrganizationService>();

        var sourceStep = CreateFullStepEntity(stepId, messageId);
        SetupRetrieveFullSteps(sourceService, sourceStep);
        SetupMessageResolve(sourceService, messageId, "Create");
        SetupTargetMessageLookup(targetService, messageId);

        targetService.Create(Arg.Any<Entity>()).Returns(newStepId);
        sourceService.When(s => s.Delete("sdkmessageprocessingstep", stepId))
            .Do(_ => throw new InvalidOperationException("Delete failed"));

        var analysis = new[] { CreateMovableItem(stepId, targetTypeId) };
        var context = new OperationContext(sourceService, targetService)
        {
            Mode = OperationMode.Move,
            IsCrossEnvironment = true,
            KeepSameIds = false
        };

        var result = _sut.Execute(context, analysis);

        var item = result.Should().ContainSingle().Subject;
        item.Success.Should().BeTrue();
        item.Message.Should().Contain("FAILED to delete from source");
    }

    private static OperationContext CreateSameEnvMoveContext(IOrganizationService service)
    {
        return new OperationContext(service, service)
        {
            Mode = OperationMode.Move,
            IsCrossEnvironment = false,
            KeepSameIds = false
        };
    }

    private static StepAnalyzeItem CreateMovableItem(Guid stepId, Guid targetTypeId, string name = "TestStep")
    {
        return new StepAnalyzeItem
        {
            StepId = stepId,
            StepName = name,
            SourcePluginTypeId = Guid.NewGuid(),
            TargetPluginTypeId = targetTypeId,
            CanMove = true,
            Reason = "Ready"
        };
    }

    private static Entity CreateFullStepEntity(Guid stepId, Guid messageId)
    {
        var entity = new Entity("sdkmessageprocessingstep", stepId);
        entity["name"] = "TestStep";
        entity["sdkmessageid"] = new EntityReference("sdkmessage", messageId);
        entity["stage"] = new OptionSetValue(20);
        entity["mode"] = new OptionSetValue(0);
        entity["rank"] = 1;
        return entity;
    }

    private static void SetupRetrieveFullSteps(IOrganizationService service, Guid stepId)
    {
        var entity = CreateFullStepEntity(stepId, Guid.NewGuid());
        SetupRetrieveFullSteps(service, entity);
    }

    private static void SetupRetrieveFullSteps(IOrganizationService service, Entity entity)
    {
        var collection = new EntityCollection(new List<Entity> { entity });
        service.RetrieveMultiple(Arg.Is<QueryExpression>(q =>
            q.EntityName == "sdkmessageprocessingstep"))
            .Returns(collection);

        // Return empty collection for image queries
        service.RetrieveMultiple(Arg.Is<QueryExpression>(q =>
            q.EntityName == "sdkmessageprocessingstepimage"))
            .Returns(new EntityCollection());
    }

    private static void SetupMessageResolve(IOrganizationService sourceService, Guid messageId, string messageName)
    {
        var messageEntity = new Entity("sdkmessage", messageId);
        messageEntity["name"] = messageName;
        sourceService.Retrieve("sdkmessage", messageId, Arg.Any<ColumnSet>())
            .Returns(messageEntity);
    }

    private static void SetupTargetMessageLookup(IOrganizationService targetService, Guid messageId)
    {
        var targetMessage = new Entity("sdkmessage", messageId);
        var messageCollection = new EntityCollection(new List<Entity> { targetMessage });
        targetService.RetrieveMultiple(Arg.Is<QueryExpression>(q =>
            q.EntityName == "sdkmessage"))
            .Returns(messageCollection);
    }
}
