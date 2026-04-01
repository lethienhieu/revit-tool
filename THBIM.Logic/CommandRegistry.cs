#if NET8_0_OR_GREATER
using Autodesk.Revit.UI;

namespace THBIM.Logic;

/// <summary>
/// Maps command keys (used by Loader proxies) to actual IExternalCommand types.
/// </summary>
internal static class CommandRegistry
{
    private static readonly Dictionary<string, Type> _commands = new(StringComparer.Ordinal)
    {
        // ==========================================
        // Annotations & General panel
        // ==========================================
        ["CallUICol"] = typeof(THBIM.CallUICol),
        ["CallUIColorslapher"] = typeof(THBIM.CallUIColorslapher),
        ["CallUIbubble"] = typeof(THBIM.CallUIbubble),
        ["CallLinkIDs"] = typeof(THBIM.CallLinkIDs),
        ["CallUIReName"] = typeof(THBIM.CallUIReName),
        ["CallUIOP"] = typeof(THBIM.CallUIOP),
        ["CallUIOverlap"] = typeof(THBIM.CallUIOverlap),
        ["CallUIParam"] = typeof(THBIM.CallUIParam),
        ["CallUIParaSync"] = typeof(THBIM.CallUIParaSync),
        ["CallUIFloordrop"] = typeof(THBIM.CallUIFloordrop),
        ["CallUIZone"] = typeof(THBIM.CallUIZone),
        ["CallUIPROSHEET"] = typeof(THBIM.CallUIPROSHEET),
        ["CallUIQTOPRO"] = typeof(THBIM.CallUIQTOPRO),
        ["SheetLinkCommand"] = typeof(THBIM.SheetLinkCommand),

        // Dim Tools
        ["AutoDimGrid"] = typeof(THBIM.AutoDimGrid),
        ["CombineDim"] = typeof(THBIM.CombineDim),

        // ProFilter
        ["ProFilterCommand"] = typeof(THBIM.ProFilterCommand),
        ["ProFilterByFamilyCommand"] = typeof(THBIM.ProFilterByFamilyCommand),
        ["ProFilterByTypeCommand"] = typeof(THBIM.ProFilterByTypeCommand),

        // AlignTags
        ["AlignTagsLeftCommand"] = typeof(THBIM.AlignTagsLeftCommand),
        ["AlignTagsRightCommand"] = typeof(THBIM.AlignTagsRightCommand),
        ["AlignTagsTopCommand"] = typeof(THBIM.AlignTagsTopCommand),
        ["AlignTagsBottomCommand"] = typeof(THBIM.AlignTagsBottomCommand),
        ["ArrangeTagsNoCrossCommand"] = typeof(THBIM.ArrangeTagsNoCrossCommand),
        ["LeaderAngleSettingCommand"] = typeof(THBIM.LeaderAngleSettingCommand),

        // ==========================================
        // MEP panel (TH Tools tab)
        // ==========================================
        ["SelectChanged"] = typeof(THBIM.SelectChanged),
        ["CreateClouds"] = typeof(THBIM.CreateClouds),
        ["Accept"] = typeof(THBIM.Accept),
        ["CallSP"] = typeof(THBIM.CallSP),
        ["CallUIQTO"] = typeof(THBIM.CallUIQTO),
        ["CallUISUP"] = typeof(THBIM.CallUISUP),

        // ==========================================
        // Structure panel
        // ==========================================
        ["CallUIStructureSync"] = typeof(THBIM.CallUIStructureSync),
        ["CallUIDP"] = typeof(THBIM.CallUIDP),
        ["CallUISplit"] = typeof(THBIM.CallUISplit),
        ["CallUILV"] = typeof(THBIM.CallUILV),
        ["AutoPile"] = typeof(THBIM.AutoPile),

        // ==========================================
        // TH Tools MEP tab — LT Create
        // ==========================================
        ["BloomCommand"] = typeof(THBIM.MEP.Commands.BloomCommand),
        ["TapCommand"] = typeof(THBIM.MEP.Commands.TapCommand),
        ["ElbowDown45Command"] = typeof(THBIM.MEP.Commands.ElbowDown45Command),
        ["ElbowRightCommand"] = typeof(THBIM.MEP.Commands.ElbowRightCommand),
        ["ElbowDownCommand"] = typeof(THBIM.MEP.Commands.ElbowDownCommand),
        ["ElbowLeftCommand"] = typeof(THBIM.MEP.Commands.ElbowLeftCommand),
        ["ElbowUpCommand"] = typeof(THBIM.MEP.Commands.ElbowUpCommand),

        // TH Tools MEP tab — LT Modify
        ["FlipMultipleCommand"] = typeof(THBIM.MEP.Commands.FlipMultipleCommand),
        ["RotateFitting30Command"] = typeof(THBIM.MEP.Commands.RotateFitting30Command),
        ["RotateFitting45Command"] = typeof(THBIM.MEP.Commands.RotateFitting45Command),
        ["RotateFitting60Command"] = typeof(THBIM.MEP.Commands.RotateFitting60Command),
        ["RotateFitting75Command"] = typeof(THBIM.MEP.Commands.RotateFitting75Command),
        ["RotateFitting90Command"] = typeof(THBIM.MEP.Commands.RotateFitting90Command),
        ["DeleteSystemCommand"] = typeof(THBIM.MEP.Commands.DeleteSystemCommand),
        ["DisconnectCommand"] = typeof(THBIM.MEP.Commands.DisconnectCommand),
        ["MoveConnectCommand"] = typeof(THBIM.MEP.Commands.MoveConnectCommand),
        ["MoveConnectAlignCommand"] = typeof(THBIM.MEP.Commands.MoveConnectAlignCommand),

        // TH Tools MEP tab — LT Align
        ["AlignIn3DCommand"] = typeof(THBIM.MEP.Commands.AlignIn3DCommand),
        ["AlignBranchCommand"] = typeof(THBIM.MEP.Commands.AlignBranchCommand),
        ["AlignBranchPlusCommand"] = typeof(THBIM.MEP.Commands.AlignBranchPlusCommand),

        // TH Tools MEP tab — LT Tools
        ["SumParameterCommand"] = typeof(THBIM.MEP.Commands.SumParameterCommand),
        ["SectionCommand"] = typeof(THBIM.MEP.Commands.SectionCommand),
    };

    public static Type? Get(string key) => _commands.GetValueOrDefault(key);
}
#endif
