using System;
using FluentAssertions;
using Xunit;

namespace PluginStepMover.Tests;

public class AnalyzeAutoMatchedTests
{
    private readonly StepMoveService _sut = new();

    [Fact]
    public void AnalyzeAutoMatched_NoSuggestedTarget_IsBlocked()
    {
        var step = CreateStep(suggestedTargetId: null);

        var result = _sut.AnalyzeAutoMatched(new[] { step });

        result.Should().ContainSingle()
            .Which.Should().Match<StepAnalyzeItem>(r =>
                r.CanMove == false &&
                r.Reason == "No matching target plugin type found");
    }

    [Fact]
    public void AnalyzeAutoMatched_EmptySuggestedTarget_IsBlocked()
    {
        var step = CreateStep(suggestedTargetId: Guid.Empty);

        var result = _sut.AnalyzeAutoMatched(new[] { step });

        result.Should().ContainSingle()
            .Which.Should().Match<StepAnalyzeItem>(r =>
                r.CanMove == false &&
                r.Reason == "No matching target plugin type found");
    }

    [Fact]
    public void AnalyzeAutoMatched_ValidSuggestedTarget_IsReady()
    {
        var step = CreateStep(suggestedTargetId: Guid.NewGuid());

        var result = _sut.AnalyzeAutoMatched(new[] { step });

        result.Should().ContainSingle()
            .Which.Should().Match<StepAnalyzeItem>(r =>
                r.CanMove == true &&
                r.Reason == "Ready");
    }

    [Fact]
    public void AnalyzeAutoMatched_SourceEqualsTarget_IsBlocked()
    {
        var typeId = Guid.NewGuid();
        var step = CreateStep(sourcePluginTypeId: typeId, suggestedTargetId: typeId);

        var result = _sut.AnalyzeAutoMatched(new[] { step });

        result.Should().ContainSingle()
            .Which.Should().Match<StepAnalyzeItem>(r =>
                r.CanMove == false &&
                r.Reason == "Source and target are identical");
    }

    [Fact]
    public void AnalyzeAutoMatched_WithWarningAndValidTarget_CanMoveWithWarning()
    {
        var step = CreateStep(suggestedTargetId: Guid.NewGuid(), matchWarning: "Multiple matches found");

        var result = _sut.AnalyzeAutoMatched(new[] { step });

        var item = result.Should().ContainSingle().Subject;
        item.CanMove.Should().BeTrue();
        item.HasWarning.Should().BeTrue();
        item.Reason.Should().Be("Multiple matches found");
    }

    [Fact]
    public void AnalyzeAutoMatched_WithWarningAndNoTarget_IsBlockedWithWarning()
    {
        var step = CreateStep(suggestedTargetId: null, matchWarning: "Ambiguous match");

        var result = _sut.AnalyzeAutoMatched(new[] { step });

        var item = result.Should().ContainSingle().Subject;
        item.CanMove.Should().BeFalse();
        item.Reason.Should().Be("Ambiguous match");
    }

    [Fact]
    public void AnalyzeAutoMatched_SetsTargetFieldsFromSuggested()
    {
        var targetId = Guid.NewGuid();
        var step = CreateStep(suggestedTargetId: targetId, suggestedTargetName: "Suggested.Plugin");

        var result = _sut.AnalyzeAutoMatched(new[] { step });

        var item = result.Should().ContainSingle().Subject;
        item.TargetPluginTypeId.Should().Be(targetId);
        item.TargetPluginTypeName.Should().Be("Suggested.Plugin");
    }

    private static StepItem CreateStep(
        Guid? sourcePluginTypeId = null,
        Guid? suggestedTargetId = null,
        string suggestedTargetName = "TargetType",
        string matchWarning = "")
    {
        return new StepItem
        {
            StepId = Guid.NewGuid(),
            StepName = "TestStep",
            SourcePluginTypeId = sourcePluginTypeId ?? Guid.NewGuid(),
            SourcePluginTypeName = "SourceType",
            SuggestedTargetPluginTypeId = suggestedTargetId,
            SuggestedTargetPluginTypeName = suggestedTargetName,
            MatchWarning = matchWarning
        };
    }
}
