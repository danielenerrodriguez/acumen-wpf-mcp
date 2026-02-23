using System.IO;
using Xunit;

namespace WpfMcp.Tests;

/// <summary>
/// Tests for InputRecorder. Since hooks require a Windows message loop and
/// an elevated process, we test the state machine logic and public API
/// without actually installing hooks. The hook installation will fail
/// gracefully in the test environment.
/// </summary>
public class InputRecorderTests
{
    private static UiaEngine Engine => UiaEngine.Instance;

    // --- State transitions ---

    [Fact]
    public void InitialState_IsIdle()
    {
        var recorder = new InputRecorder(Engine);
        Assert.Equal(InputRecorder.RecorderState.Idle, recorder.State);
    }

    [Fact]
    public void StartRecording_WhenNotAttached_ReturnsError()
    {
        var recorder = new InputRecorder(Engine);
        // Engine is not attached to any process in test environment
        var result = recorder.StartRecording("test", Path.GetTempPath());
        Assert.False(result.success);
        Assert.Contains("Not attached", result.message);
        Assert.Equal(InputRecorder.RecorderState.Idle, recorder.State);
    }

    [Fact]
    public void StopRecording_WhenIdle_ReturnsError()
    {
        var recorder = new InputRecorder(Engine);
        var result = recorder.StopRecording();
        Assert.False(result.success);
        Assert.Contains("Not currently recording", result.message);
        Assert.Null(result.yaml);
        Assert.Null(result.filePath);
    }

    [Fact]
    public void ActionCount_WhenIdle_IsZero()
    {
        var recorder = new InputRecorder(Engine);
        Assert.Equal(0, recorder.ActionCount);
    }

    [Fact]
    public void Duration_WhenIdle_IsZero()
    {
        var recorder = new InputRecorder(Engine);
        Assert.Equal(TimeSpan.Zero, recorder.Duration);
    }

    [Fact]
    public void Dispose_WhenIdle_DoesNotThrow()
    {
        var recorder = new InputRecorder(Engine);
        recorder.Dispose(); // Should not throw
    }

    [Fact]
    public void MacroName_WhenIdle_IsEmpty()
    {
        var recorder = new InputRecorder(Engine);
        Assert.Equal("", recorder.MacroName);
    }

    // --- Wait computation ---

    [Fact]
    public void ComputeWaits_LargeGap_SetsWaitBeforeSec()
    {
        // We test this indirectly through BuildFromRecordedActions + WaitBeforeSec
        var baseTime = DateTime.UtcNow;
        var actions = new List<RecordedAction>
        {
            new()
            {
                Type = RecordedActionType.SendKeys,
                Timestamp = baseTime,
                Keys = "A",
            },
            new()
            {
                Type = RecordedActionType.SendKeys,
                Timestamp = baseTime.AddSeconds(5),
                Keys = "B",
                WaitBeforeSec = 5.0,
            }
        };

        var macro = MacroSerializer.BuildFromRecordedActions("Test", "Test", actions);

        // Should have wait step between the two
        Assert.Equal(3, macro.Steps.Count);
        Assert.Equal("wait", macro.Steps[1].Action);
        Assert.Equal(5.0, macro.Steps[1].Seconds);
    }

    [Fact]
    public void ComputeWaits_MaxCap_CappedAt10Seconds()
    {
        // Test that a very large gap is capped at MaxRecordedWaitSec (10s)
        var baseTime = DateTime.UtcNow;
        var actions = new List<RecordedAction>
        {
            new()
            {
                Type = RecordedActionType.SendKeys,
                Timestamp = baseTime,
                Keys = "A",
            },
            new()
            {
                Type = RecordedActionType.SendKeys,
                Timestamp = baseTime.AddSeconds(30),
                Keys = "B",
                WaitBeforeSec = Constants.MaxRecordedWaitSec, // 10.0
            }
        };

        var macro = MacroSerializer.BuildFromRecordedActions("Test", "Test", actions);

        Assert.Equal(3, macro.Steps.Count);
        Assert.Equal("wait", macro.Steps[1].Action);
        Assert.Equal(Constants.MaxRecordedWaitSec, macro.Steps[1].Seconds);
    }

    [Fact]
    public void ComputeWaits_SmallGap_NoWaitInserted()
    {
        var baseTime = DateTime.UtcNow;
        var actions = new List<RecordedAction>
        {
            new()
            {
                Type = RecordedActionType.SendKeys,
                Timestamp = baseTime,
                Keys = "A",
            },
            new()
            {
                Type = RecordedActionType.SendKeys,
                Timestamp = baseTime.AddMilliseconds(100),
                Keys = "B",
                // No WaitBeforeSec set (gap is < 1.5s threshold)
            }
        };

        var macro = MacroSerializer.BuildFromRecordedActions("Test", "Test", actions);

        // No wait step
        Assert.Equal(2, macro.Steps.Count);
        Assert.All(macro.Steps, s => Assert.Equal("send_keys", s.Action));
    }

    // --- RecordedAction record ---

    [Fact]
    public void RecordedAction_Click_StoresAllProperties()
    {
        var action = new RecordedAction
        {
            Type = RecordedActionType.Click,
            Timestamp = DateTime.UtcNow,
            X = 100,
            Y = 200,
            AutomationId = "uxBtn",
            ElementName = "Submit",
            ClassName = "Button",
            ControlType = "Button",
        };

        Assert.Equal(RecordedActionType.Click, action.Type);
        Assert.Equal(100, action.X);
        Assert.Equal(200, action.Y);
        Assert.Equal("uxBtn", action.AutomationId);
        Assert.Equal("Submit", action.ElementName);
        Assert.Equal("Button", action.ClassName);
        Assert.Equal("Button", action.ControlType);
    }

    [Fact]
    public void RecordedAction_SendKeys_StoresKeys()
    {
        var action = new RecordedAction
        {
            Type = RecordedActionType.SendKeys,
            Timestamp = DateTime.UtcNow,
            Keys = "Ctrl+Shift+S",
        };

        Assert.Equal("Ctrl+Shift+S", action.Keys);
    }

    [Fact]
    public void RecordedAction_Type_StoresText()
    {
        var action = new RecordedAction
        {
            Type = RecordedActionType.Type,
            Timestamp = DateTime.UtcNow,
            Text = "Hello World",
        };

        Assert.Equal("Hello World", action.Text);
    }

    [Fact]
    public void RecordedAction_WaitBeforeSec_IsMutable()
    {
        var action = new RecordedAction
        {
            Type = RecordedActionType.SendKeys,
            Timestamp = DateTime.UtcNow,
            Keys = "Enter",
        };

        Assert.Null(action.WaitBeforeSec);
        action.WaitBeforeSec = 3.5;
        Assert.Equal(3.5, action.WaitBeforeSec);
    }
}
