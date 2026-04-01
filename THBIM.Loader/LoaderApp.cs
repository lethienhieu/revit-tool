using Autodesk.Revit.UI;
using System.IO;
using System.Reflection;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Linq;

namespace THBIM.Loader;

public class LoaderApp : IExternalApplication
{
    public Result OnStartup(UIControlledApplication app)
    {
        // Set up PluginManager base directory
        string asm = Assembly.GetExecutingAssembly().Location;
        string dir = Path.GetDirectoryName(asm) ?? "";
        PluginManager.Instance.BaseDir = dir;

        // Load THBIM.Logic
        try
        {
            PluginManager.Instance.Load();
            PluginManager.Instance.LogicEntry?.Initialize(app);
        }
        catch (Exception ex)
        {
            TaskDialog.Show("THBIM Load Error",
                $"Failed to load THBIM Logic:\n{ex.Message}");
        }

        // ==========================================
        // RIBBON SETUP (same structure as original App.cs)
        // ==========================================
        const string tab = "TH Tools";
        try { app.CreateRibbonTab(tab); } catch { }

        RegisterAnnotationsPanel(app, tab, asm, dir);
        RegisterMepPanel(app, tab, asm, dir);
        RegisterStructurePanel(app, tab, asm, dir);
        RegisterDevToolsPanel(app, tab, asm, dir);

        try { RegisterMepTab(app, asm, dir); } catch { }

        return Result.Succeeded;
    }

    private void RegisterAnnotationsPanel(UIControlledApplication app, string tab, string asm, string dir)
    {
        const string panelName = "Annotations & General";
        var panel = app.GetRibbonPanels(tab).Find(p => p.Name == panelName)
                   ?? app.CreateRibbonPanel(tab, panelName);

        // Icons
        string icon16Bubble = Path.Combine(dir, "Resources", "GridBubble.png");
        string icon16Dim = Path.Combine(dir, "Resources", "DIM.png");
        string icon16CombineDim = Path.Combine(dir, "Resources", "DIM.png");
        string icon16Drop = Path.Combine(dir, "Resources", "Floordrop_16.png");
        string icon32Col = Path.Combine(dir, "Resources", "Col_32.png");
        string icon16ID = Path.Combine(dir, "Resources", "IDs_16.png");
        string icon16RS = Path.Combine(dir, "Resources", "Rename_16.png");
        string icon16CS = Path.Combine(dir, "Resources", "ColorSplasher_16.png");
        string icon16PF = Path.Combine(dir, "Resources", "Profilter_16.png");
        string iconzone16 = Path.Combine(dir, "Resources", "Zone_16.png");
        string iconParam = Path.Combine(dir, "Resources", "CombineParam_16.png");
        string iconParam16 = Path.Combine(dir, "Resources", "CombineParam_16.png");
        string icon16OV = Path.Combine(dir, "Resources", "overlap_16.png");
        string iconSync16 = Path.Combine(dir, "Resources", "SyncParam_16.png");
        string iconProSheet32 = Path.Combine(dir, "Resources", "ProSheet_32.png");
        string iconQTOPRO32 = Path.Combine(dir, "Resources", "QTOPRO_32.png");
        string iconSheetLink32 = Path.Combine(dir, "Resources", "Sheetlink_32.png");
        string icon16Align = Path.Combine(dir, "Resources", "AlignTags_16.png");
        string iconAlignLeft = Path.Combine(dir, "Resources", "LeftAlign_16.png");
        string iconAlignRight = Path.Combine(dir, "Resources", "RightAlign_16.png");
        string iconAlignTop = Path.Combine(dir, "Resources", "TopAlign_16.png");
        string iconAlignBottom = Path.Combine(dir, "Resources", "BottomAlign_16.png");
        string iconArr = Path.Combine(dir, "Resources", "Arrange_16.png");

        // Proxy namespace prefix
        const string ns = "THBIM.Loader.Proxies.";

        var pbCS = new PushButtonData("ColorSplasher_Run", "Color Splasher", asm, ns + "Proxy_CallUIColorslapher") { ToolTip = "Visualize model data by automatically coloring elements.", Image = LoadIcon(icon16CS) };
        var pdAlignData = new PulldownButtonData("THBIM.AlignTags", "Align Tags") { ToolTip = "Tools to align and arrange Tags.", Image = LoadIcon(icon16Align) };
        var pbCol = new PushButtonData("ColumnDim_Run", "AutoDim & Tag\nColumn", asm, ns + "Proxy_CallUICol") { ToolTip = "Automatically dimension columns and grids.", LargeImage = LoadIcon(icon32Col) };
        var pdParamData = new PulldownButtonData("ParameterTools_Pulldown", "Parameter Tools") { ToolTip = "Tools for project parameters management.", Image = LoadIcon(iconParam) };
        var pbOV = new PushButtonData("Checkoverlap_Run", "Check overlap", asm, ns + "Proxy_CallUIOverlap") { ToolTip = "Check overlap and highlight.", Image = LoadIcon(icon16OV) };
        var pbID = new PushButtonData("LinkIDs_Run", "Linked IDS", asm, ns + "Proxy_CallLinkIDs") { ToolTip = "Display ElementId of selected element.", Image = LoadIcon(icon16ID) };
        var pbRS = new PushButtonData("RenameSheets_Run", "Rename Sheets", asm, ns + "Proxy_CallUIReName") { ToolTip = "Rename Sheet and count SheetNumber.", Image = LoadIcon(icon16RS) };
        var pdPFData = new PulldownButtonData("ProFilter_DD", "Pro Filter") { ToolTip = "Select elements by Category, Family, or Type.", Image = LoadIcon(icon16PF) };
        var pbBubble = new PushButtonData("GridBubble_Run", "GridBubble", asm, ns + "Proxy_CallUIbubble") { ToolTip = "Show/Hide Grid/Level heads.", Image = LoadIcon(icon16Bubble) };
        var pbProSheet = new PushButtonData("ProSheets_Run", "ProSheets", asm, ns + "Proxy_CallUIPROSHEET") { ToolTip = "Export View/Sheets to PDF, DWG.", LargeImage = LoadIcon(iconProSheet32) };
        var pbQTOPRO = new PushButtonData("QTOPRO_Run", "QTOPRO", asm, ns + "Proxy_CallUIQTOPRO") { ToolTip = "Extract QTO from Revit directly to Excel.", LargeImage = LoadIcon(iconQTOPRO32) };
        var pbSheetLink = new PushButtonData("SheetLink_Run", "SheetLink", asm, ns + "Proxy_SheetLinkCommand") { ToolTip = "Open the SheetLink window.", LargeImage = LoadIcon(iconSheetLink32) };
        var pdDimGroup = new PulldownButtonData("DimensionTools_Pulldown", "Dim Tools") { ToolTip = "Dimensioning helper tools.", Image = LoadIcon(icon16Dim) };
        var pbDrop = new PushButtonData("FloorDrop_Run", "Detect Floor Drop", asm, ns + "Proxy_CallUIFloordrop") { ToolTip = "Calculate floor drop values.", Image = LoadIcon(icon16Drop) };
        var pbzone = new PushButtonData("Zone_Run", "Set Zone", asm, ns + "Proxy_CallUIZone") { ToolTip = "Set up Zone", Image = LoadIcon(iconzone16) };

        panel.AddItem(pbCol);
        panel.AddSeparator();
        panel.AddItem(pbProSheet);
        panel.AddSeparator();
        panel.AddItem(pbSheetLink);
        panel.AddSeparator();
        panel.AddItem(pbQTOPRO);
        panel.AddSeparator();

        IList<RibbonItem> stackFilter = panel.AddStackedItems(pbRS, pbID, pdPFData);
        if (stackFilter.Count > 2 && stackFilter[2] is PulldownButton pdPF)
        {
            pdPF.AddPushButton(new PushButtonData("PF_Category", "By Category", asm, ns + "Proxy_ProFilterCommand") { ToolTip = "Filter elements by Category" });
            pdPF.AddPushButton(new PushButtonData("PF_Family", "By Family", asm, ns + "Proxy_ProFilterByFamilyCommand") { ToolTip = "Filter elements by Family name" });
            pdPF.AddPushButton(new PushButtonData("PF_Type", "By Type", asm, ns + "Proxy_ProFilterByTypeCommand") { ToolTip = "Filter elements by Type name" });
        }
        panel.AddSeparator();

        IList<RibbonItem> stackDim = panel.AddStackedItems(pbBubble, pdDimGroup, pbDrop);
        if (stackDim.Count > 1 && stackDim[1] is PulldownButton pdDim)
        {
            pdDim.AddPushButton(new PushButtonData("AutoDimGrid_Sub", "Auto Dim Grid", asm, ns + "Proxy_AutoDimGrid") { ToolTip = "Automatically create dimensions for grids.", Image = LoadIcon(icon16Dim) });
            pdDim.AddPushButton(new PushButtonData("CombineDim_Sub", "Combine Dim", asm, ns + "Proxy_CombineDim") { ToolTip = "Combine multiple collinear dimensions into one.", Image = LoadIcon(icon16CombineDim) });
        }
        panel.AddSeparator();

        IList<RibbonItem> stackAlign = panel.AddStackedItems(pbCS, pdAlignData, pbzone);
        if (stackAlign.Count > 1 && stackAlign[1] is PulldownButton pullDown)
        {
            pullDown.AddPushButton(new PushButtonData("THBIM.AlignTags.Left", "Left", asm, ns + "Proxy_AlignTagsLeftCommand") { LargeImage = LoadIcon(iconAlignLeft) });
            pullDown.AddPushButton(new PushButtonData("THBIM.AlignTags.Right", "Right", asm, ns + "Proxy_AlignTagsRightCommand") { LargeImage = LoadIcon(iconAlignRight) });
            pullDown.AddPushButton(new PushButtonData("THBIM.AlignTags.Top", "Top", asm, ns + "Proxy_AlignTagsTopCommand") { LargeImage = LoadIcon(iconAlignTop) });
            pullDown.AddPushButton(new PushButtonData("THBIM.AlignTags.Bottom", "Bottom", asm, ns + "Proxy_AlignTagsBottomCommand") { LargeImage = LoadIcon(iconAlignBottom) });
            pullDown.AddPushButton(new PushButtonData("THBIM.AlignTags.Autoarrange", "Autoarrange", asm, ns + "Proxy_ArrangeTagsNoCrossCommand") { LargeImage = LoadIcon(iconArr) });
            pullDown.AddSeparator();
            pullDown.AddPushButton(new PushButtonData("THBIM.AlignTags.LeaderAngle", "Leader Angle Setting", asm, ns + "Proxy_LeaderAngleSettingCommand") { ToolTip = "Set leader elbow angle for align commands" });
        }

        panel.AddSeparator();
        IList<RibbonItem> stackParam = panel.AddStackedItems(pdParamData, pbOV);
        if (stackParam.Count > 0 && stackParam[0] is PulldownButton pdParam)
        {
            pdParam.AddPushButton(new PushButtonData("CombineParam_Sub", "Combine Param", asm, ns + "Proxy_CallUIParam") { Image = LoadIcon(iconParam16), LargeImage = LoadIcon(iconParam16) });
            pdParam.AddPushButton(new PushButtonData("SyncParam_Sub", "Sync Parameter", asm, ns + "Proxy_CallUIParaSync") { Image = LoadIcon(iconSync16), LargeImage = LoadIcon(iconSync16) });
        }

        ColorizePanel(tab, panelName, "#CE93D8");
    }

    private void RegisterMepPanel(UIControlledApplication app, string tab, string asm, string dir)
    {
        const string panelName = "MEP";
        var panel = app.GetRibbonPanels(tab).Find(p => p.Name == panelName) ?? app.CreateRibbonPanel(tab, panelName);
        const string ns = "THBIM.Loader.Proxies.";

        string iconOP_32 = Path.Combine(dir, "Resources", "OPENING.png");
        string iconSL_16 = Path.Combine(dir, "Resources", "select_16.png");
        string iconCLD16 = Path.Combine(dir, "Resources", "cloud_16.png");
        string iconAC16 = Path.Combine(dir, "Resources", "accept_16.png");
        string iconSP16 = Path.Combine(dir, "Resources", "SplitPipe_16.png");
        string iconBOQ16 = Path.Combine(dir, "Resources", "BOQ_16.png");
        string iconSUP16 = Path.Combine(dir, "Resources", "hanger_16.png");

        var pbOP = new PushButtonData("OPENING_Run", "Create OPENING", asm, ns + "Proxy_CallUIOP") { ToolTip = "Open the interface to create openings (sleeves).", LargeImage = LoadIcon(iconOP_32) };
        var pbSL = new PushButtonData("Select_Run", "Select Changed", asm, ns + "Proxy_SelectChanged") { ToolTip = "Select all changed openings (NeedsReview=true).", Image = LoadIcon(iconSL_16) };
        var pbCLD = new PushButtonData("CreateClouds_Run", "Create Clouds", asm, ns + "Proxy_CreateClouds") { ToolTip = "Create revision clouds around changed openings.", Image = LoadIcon(iconCLD16) };
        var pbAC = new PushButtonData("AcceptChanged_Run", "Accept Changed", asm, ns + "Proxy_Accept") { ToolTip = "Accept changes and update baseline.", Image = LoadIcon(iconAC16) };
        var pbSP = new PushButtonData("SplitPipe_Run", "Split Pipe", asm, ns + "Proxy_CallSP") { ToolTip = "Divide the pipe length according to the product length.", Image = LoadIcon(iconSP16) };
        var pbQTO = new PushButtonData("MechanicalQTO_Run", "Mechanical QTO", asm, ns + "Proxy_CallUIQTO") { ToolTip = "Select MEP elements for quantity takeoff.", Image = LoadIcon(iconBOQ16) };
        var pbSUP = new PushButtonData("Hanger_Run", "Hanger", asm, ns + "Proxy_CallUISUP") { ToolTip = "Generates hangers and supports for Pipes and Ducts.", Image = LoadIcon(iconSUP16) };

        panel.AddItem(pbOP);
        panel.AddStackedItems(pbSL, pbCLD, pbAC);
        panel.AddSeparator();
        panel.AddStackedItems(pbSP, pbQTO, pbSUP);

        ColorizePanel(tab, panelName, "#8FC67A");
    }

    private void RegisterStructurePanel(UIControlledApplication app, string tab, string asm, string dir)
    {
        const string panelName = "Structure";
        var panel = app.GetRibbonPanels(tab).Find(p => p.Name == panelName) ?? app.CreateRibbonPanel(tab, panelName);
        const string ns = "THBIM.Loader.Proxies.";

        string iconDP32 = Path.Combine(dir, "Resources", "Droppanel_32.png");
        string iconSPLIT16 = Path.Combine(dir, "Resources", "Splitcol_16.png");
        string iconLV16 = Path.Combine(dir, "Resources", "LevelRehost_16.png");
        string iconATP16 = Path.Combine(dir, "Resources", "AutoPile_16.png");

        var pbDP = new PushButtonData("AutoDropPanel_Run", "Create DropPanel", asm, ns + "Proxy_CallUIDP") { ToolTip = "Automatically create Drop Panel or Pile Cap.", LargeImage = LoadIcon(iconDP32) };
        var pbSplitCol = new PushButtonData("SplitColumnl_Run", "Split Columnl", asm, ns + "Proxy_CallUISplit") { ToolTip = "Split columns by floor levels.", Image = LoadIcon(iconSPLIT16) };
        var pbLV = new PushButtonData("LevelRehost_Run", "Level Rehost", asm, ns + "Proxy_CallUILV") { ToolTip = "Move elements to new level while keeping 3D position.", Image = LoadIcon(iconLV16) };
        var pbATP = new PushButtonData("AutoPile_Run", "create Pile", asm, ns + "Proxy_AutoPile") { ToolTip = "Automatically create Piling.", Image = LoadIcon(iconATP16) };

        panel.AddItem(pbDP);
        panel.AddSeparator();
        panel.AddStackedItems(pbATP, pbSplitCol, pbLV);

        ColorizePanel(tab, panelName, "#E38888");
    }

    private void RegisterDevToolsPanel(UIControlledApplication app, string tab, string asm, string dir)
    {
        const string panelName = "Dev";
        var panel = app.CreateRibbonPanel(tab, panelName);

        string iconReload = Path.Combine(dir, "Resources", "update_32.png");
        var pbReload = new PushButtonData("Reload_Run", "Reload\nPlugin", asm,
            "THBIM.Loader.ReloadCommand")
        {
            ToolTip = "Hot-reload THBIM Logic without restarting Revit.",
            LargeImage = File.Exists(iconReload) ? new BitmapImage(new Uri(iconReload)) : null,
            Image = File.Exists(iconReload) ? new BitmapImage(new Uri(iconReload)) : null
        };
        panel.AddItem(pbReload);

        ColorizePanel(tab, panelName, "#90CAF9");
    }

    private void RegisterMepTab(UIControlledApplication app, string asm, string dir)
    {
        const string mepTab = "TH Tools MEP";
        try { app.CreateRibbonTab(mepTab); } catch { }
        string mepIconDir = Path.Combine(dir, "Resources", "MEP");
        const string ns = "THBIM.Loader.Proxies.";

        // --- LT Create ---
        var pCreate = app.CreateRibbonPanel(mepTab, "LT Create");
        pCreate.AddItem(MepBtn("Bloom", "Bloom", asm, ns + "Proxy_BloomCommand", mepIconDir, "Bloom32.png"));
        pCreate.AddItem(MepBtn("Tap", "Tap", asm, ns + "Proxy_TapCommand", mepIconDir, "Tap32.png"));
        pCreate.AddItem(MepBtn("ElbowDown45", "Elbow Down 45", asm, ns + "Proxy_ElbowDown45Command", mepIconDir, "ElbowDown4532.png"));
        pCreate.AddItem(MepBtn("ElbowRight", "Elbow Right", asm, ns + "Proxy_ElbowRightCommand", mepIconDir, "ElbowRight32.png"));
        pCreate.AddItem(MepBtn("ElbowDown", "Elbow Down", asm, ns + "Proxy_ElbowDownCommand", mepIconDir, "ElbowDown32.png"));
        pCreate.AddItem(MepBtn("ElbowLeft", "Elbow Left", asm, ns + "Proxy_ElbowLeftCommand", mepIconDir, "ElbowLeft32.png"));
        pCreate.AddItem(MepBtn("ElbowUp", "Elbow Up", asm, ns + "Proxy_ElbowUpCommand", mepIconDir, "ElbowUp32.png"));
        ColorizePanel(mepTab, "LT Create", "#64B5F6");

        // --- LT Modify ---
        var pModify = app.CreateRibbonPanel(mepTab, "LT Modify");
        pModify.AddItem(MepBtn("FlipMultiple", "Flip Multiple", asm, ns + "Proxy_FlipMultipleCommand", mepIconDir, "Microdesk.FlipMultiple32.png"));
        var rotateData = new SplitButtonData("RotateFittings", "Rotate");
        var rotateSplit = pModify.AddItem(rotateData) as SplitButton;
        if (rotateSplit != null)
        {
            string rotIcon = Path.Combine(mepIconDir, "RotateFitting32.png");
            if (File.Exists(rotIcon)) { var img = new BitmapImage(new Uri(rotIcon)); rotateSplit.Image = img; rotateSplit.LargeImage = img; }
            rotateSplit.AddPushButton(MepBtn("Rotate30", "Rotate 30\u00B0", asm, ns + "Proxy_RotateFitting30Command", mepIconDir, "RotateFitting32.png"));
            rotateSplit.AddPushButton(MepBtn("Rotate45", "Rotate 45\u00B0", asm, ns + "Proxy_RotateFitting45Command", mepIconDir, "RotateFitting32.png"));
            rotateSplit.AddPushButton(MepBtn("Rotate60", "Rotate 60\u00B0", asm, ns + "Proxy_RotateFitting60Command", mepIconDir, "RotateFitting32.png"));
            rotateSplit.AddPushButton(MepBtn("Rotate75", "Rotate 75\u00B0", asm, ns + "Proxy_RotateFitting75Command", mepIconDir, "RotateFitting32.png"));
            rotateSplit.AddPushButton(MepBtn("Rotate90", "Rotate 90\u00B0", asm, ns + "Proxy_RotateFitting90Command", mepIconDir, "RotateFitting32.png"));
        }
        pModify.AddItem(MepBtn("DeleteSystem", "Delete System", asm, ns + "Proxy_DeleteSystemCommand", mepIconDir, "DeleteSystem32.png"));
        pModify.AddItem(MepBtn("Disconnect", "Disconnect", asm, ns + "Proxy_DisconnectCommand", mepIconDir, "Disconnect32.png"));
        pModify.AddItem(MepBtn("MoveConnect", "Move Connect", asm, ns + "Proxy_MoveConnectCommand", mepIconDir, "MoveConnect32.png"));
        pModify.AddItem(MepBtn("MoveConnectAlign", "Mo-Co Align", asm, ns + "Proxy_MoveConnectAlignCommand", mepIconDir, "MoveConnectAlign32.png"));
        ColorizePanel(mepTab, "LT Modify", "#FFB74D");

        // --- LT Align ---
        var pAlign = app.CreateRibbonPanel(mepTab, "LT Align");
        pAlign.AddItem(MepBtn("Align3D", "Align in 3D", asm, ns + "Proxy_AlignIn3DCommand", mepIconDir, "AlignIn3D32.png"));
        pAlign.AddItem(MepBtn("AlignBranch", "Align Branch", asm, ns + "Proxy_AlignBranchCommand", mepIconDir, "BranchAlignLite32.png"));
        pAlign.AddItem(MepBtn("AlignBranchPlus", "Align Branch+", asm, ns + "Proxy_AlignBranchPlusCommand", mepIconDir, "BranchAlign32.png"));
        ColorizePanel(mepTab, "LT Align", "#81C784");

        // --- LT Tools ---
        var pTools = app.CreateRibbonPanel(mepTab, "LT Tools");
        pTools.AddItem(MepBtn("SumParam", "Sum Param.", asm, ns + "Proxy_SumParameterCommand", mepIconDir, "Sum32.png"));
        pTools.AddItem(MepBtn("Section", "Section", asm, ns + "Proxy_SectionCommand", mepIconDir, "Section32.png"));
        ColorizePanel(mepTab, "LT Tools", "#BA68C8");
    }

    public Result OnShutdown(UIControlledApplication app)
    {
        try { PluginManager.Instance.Shutdown(); } catch { }
        return Result.Succeeded;
    }

    // ==========================================
    // HELPER METHODS
    // ==========================================

    private static PushButtonData MepBtn(string name, string text, string asm, string className, string iconDir, string iconFile)
    {
        var btn = new PushButtonData(name, text, asm, className);
        string path = Path.Combine(iconDir, iconFile);
        if (File.Exists(path))
        {
            var img = new BitmapImage(new Uri(path));
            btn.Image = img;
            btn.LargeImage = img;
        }
        return btn;
    }

    private static BitmapImage? LoadIcon(string path)
    {
        return File.Exists(path) ? new BitmapImage(new Uri(path)) : null;
    }

    private static void ColorizePanel(string tabName, string panelName, string hexColor)
    {
        Autodesk.Windows.RibbonControl ribbon = Autodesk.Windows.ComponentManager.Ribbon;
        var targetTab = ribbon.Tabs.FirstOrDefault(t => t.Name == tabName || t.Title == tabName);
        if (targetTab == null) return;
        var targetPanel = targetTab.Panels.FirstOrDefault(p => p.Source.Title == panelName);
        if (targetPanel != null)
        {
            var converter = new BrushConverter();
            var brush = (Brush)converter.ConvertFromString(hexColor);
            targetPanel.CustomPanelTitleBarBackground = brush;
        }
    }
}
