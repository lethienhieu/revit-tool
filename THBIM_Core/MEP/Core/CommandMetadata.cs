#if NET8_0_OR_GREATER
namespace THBIM.MEP.Core;

internal sealed record CommandMetadata(
    string InternalName,
    string DisplayName,
    string Tooltip,
    string IconFile,
    string CommandTypeName,
    string? LongDescription = null,
    bool LargeButton = false);

/// <summary>
/// A group of commands shown as a SplitButton dropdown in the ribbon.
/// The first command is the default (shown on the button face).
/// </summary>
internal sealed record SplitButtonGroup(
    string InternalName,
    string DisplayName,
    string IconFile,
    bool LargeButton,
    CommandMetadata[] Items);

/// <summary>
/// A ribbon item can be either a single command or a split-button group.
/// </summary>
internal sealed class RibbonItem
{
    public CommandMetadata? Command { get; init; }
    public SplitButtonGroup? SplitGroup { get; init; }

    public static implicit operator RibbonItem(CommandMetadata cmd) => new() { Command = cmd };
    public static implicit operator RibbonItem(SplitButtonGroup grp) => new() { SplitGroup = grp };
}

internal static class SuiteDefinition
{
    public const string TabName = "TH Tools MEP";

    public static IReadOnlyList<(string Panel, RibbonItem[] Items)> Panels { get; } =
    [
        ("LT Create",
        [
            (RibbonItem)new CommandMetadata("Bloom", "Bloom", "Extend selected MEP curves from their open end by 500mm. Loop pick until ESC.", "Bloom32.png", "THBIM.MEP.Commands.BloomCommand", LargeButton: true),
            (RibbonItem)new CommandMetadata("Tap", "Tap", "Connect branch to main with auto tee/elbow.", "Tap32.png", "THBIM.MEP.Commands.TapCommand", LargeButton: true),
            (RibbonItem)new CommandMetadata("ElbowDown45", "Elbow Down 45", "Offset MEP curve diagonally downward 45\u00B0.", "ElbowDown4532.png", "THBIM.MEP.Commands.ElbowDown45Command"),
            (RibbonItem)new CommandMetadata("ElbowRight", "Elbow Right", "Offset MEP curve to the right.", "ElbowRight32.png", "THBIM.MEP.Commands.ElbowRightCommand"),
            (RibbonItem)new CommandMetadata("ElbowDown", "Elbow Down", "Offset MEP curve downward.", "ElbowDown32.png", "THBIM.MEP.Commands.ElbowDownCommand"),
            (RibbonItem)new CommandMetadata("ElbowLeft", "Elbow Left", "Offset MEP curve to the left.", "ElbowLeft32.png", "THBIM.MEP.Commands.ElbowLeftCommand"),
            (RibbonItem)new CommandMetadata("ElbowUp", "Elbow Up", "Offset MEP curve upward.", "ElbowUp32.png", "THBIM.MEP.Commands.ElbowUpCommand")
        ]),
        ("LT Modify",
        [
            (RibbonItem)new CommandMetadata("FlipMultiple", "Flip Multiple", "Loop pick family instances to flip. ESC to finish.", "Microdesk.FlipMultiple32.png", "THBIM.MEP.Commands.FlipMultipleCommand", LargeButton: true),
            (RibbonItem)new SplitButtonGroup("RotateFittings", "Rotate", "RotateFitting32.png", LargeButton: false, Items:
            [
                new("Rotate30", "Rotate 30\u00B0", "Pick fitting to rotate 30\u00B0 incrementally. ESC to finish.", "RotateFitting32.png", "THBIM.MEP.Commands.RotateFitting30Command"),
                new("Rotate45", "Rotate 45\u00B0", "Pick fitting to rotate 45\u00B0 incrementally. ESC to finish.", "RotateFitting32.png", "THBIM.MEP.Commands.RotateFitting45Command"),
                new("Rotate60", "Rotate 60\u00B0", "Pick fitting to rotate 60\u00B0 incrementally. ESC to finish.", "RotateFitting32.png", "THBIM.MEP.Commands.RotateFitting60Command"),
                new("Rotate75", "Rotate 75\u00B0", "Pick fitting to rotate 75\u00B0 incrementally. ESC to finish.", "RotateFitting32.png", "THBIM.MEP.Commands.RotateFitting75Command"),
                new("Rotate90", "Rotate 90\u00B0", "Pick fitting to rotate 90\u00B0 incrementally. ESC to finish.", "RotateFitting32.png", "THBIM.MEP.Commands.RotateFitting90Command")
            ]),
            (RibbonItem)new CommandMetadata("DeleteSystem", "Delete System", "Traverse connected network from picked element and disconnect all.", "DeleteSystem32.png", "THBIM.MEP.Commands.DeleteSystemCommand"),
            (RibbonItem)new CommandMetadata("Disconnect", "Disconnect", "Loop pick MEP elements to disconnect. ESC to finish.", "Disconnect32.png", "THBIM.MEP.Commands.DisconnectCommand"),
            (RibbonItem)new CommandMetadata("MoveConnect", "Move Connect", "Pick reference first, then follower; follower moves and connects.", "MoveConnect32.png", "THBIM.MEP.Commands.MoveConnectCommand"),
            (RibbonItem)new CommandMetadata("MoveConnectAlign", "Mo-Co Align", "Pick reference first, then follower; follower rotates, moves, connects.", "MoveConnectAlign32.png", "THBIM.MEP.Commands.MoveConnectAlignCommand")
        ]),
        ("LT Align",
        [
            (RibbonItem)new CommandMetadata("Align3D", "Align in 3D", "Pick reference first, then follower; follower origin aligns to reference.", "AlignIn3D32.png", "THBIM.MEP.Commands.AlignIn3DCommand", LargeButton: true),
            (RibbonItem)new CommandMetadata("AlignBranch", "Align Branch", "Pick main first, then branch; branch moves onto main centerline.", "BranchAlignLite32.png", "THBIM.MEP.Commands.AlignBranchCommand"),
            (RibbonItem)new CommandMetadata("AlignBranchPlus", "Align Branch+", "Align branch to main and attempt tee/elbow connection.", "BranchAlign32.png", "THBIM.MEP.Commands.AlignBranchPlusCommand")
        ]),
        ("LT Tools",
        [
            (RibbonItem)new CommandMetadata("SumParam", "Sum Param.", "Select multiple elements to sum their numeric parameters.", "Sum32.png", "THBIM.MEP.Commands.SumParameterCommand", LargeButton: true),
            (RibbonItem)new CommandMetadata("Section", "Section", "Pick element to create section view around it.", "Section32.png", "THBIM.MEP.Commands.SectionCommand")
        ])
    ];
}
#endif
