using Autodesk.Revit.UI;
using System;
using System.IO;
using System.Reflection;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Linq;
using System.Collections.Generic;

namespace THBIM
{
    public class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication app)
        {
            const string tab = "TH Tools";
            
            // Generate Tab
            try { app.CreateRibbonTab(tab); } catch { }

            // ==========================================
            // ANNOTATIONS & GENERAL PANEL
            // ==========================================
            const string panelAnnotName = "Annotations & General";
            var panelAnnot = app.GetRibbonPanels(tab).Find(p => p.Name == panelAnnotName)
                        ?? app.CreateRibbonPanel(tab, panelAnnotName);

            string asm = Assembly.GetExecutingAssembly().Location;
            string dir = Path.GetDirectoryName(asm) ?? "";

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

            var pbCS = new PushButtonData("ColorSplasher_Run", "Color Splasher", asm, "THBIM.CallUIColorslapher") { ToolTip = "Visualize model data by automatically coloring elements.", Image = File.Exists(icon16CS) ? new BitmapImage(new Uri(icon16CS)) : null };
            var pdAlignData = new PulldownButtonData("THBIM.AlignTags", "Align Tags") { ToolTip = "Tools to align and arrange Tags.", Image = File.Exists(icon16Align) ? new BitmapImage(new Uri(icon16Align)) : null };
            var pbCol = new PushButtonData("ColumnDim_Run", "AutoDim & Tag\nColumn", asm, "THBIM.CallUICol") { ToolTip = "Automatically dimension columns and grids.", LargeImage = File.Exists(icon32Col) ? new BitmapImage(new Uri(icon32Col)) : null };
            var pdParamData = new PulldownButtonData("ParameterTools_Pulldown", "Parameter Tools") { ToolTip = "Tools for project parameters management.", Image = File.Exists(iconParam) ? new BitmapImage(new Uri(iconParam)) : null };
            var pbOV = new PushButtonData("Checkoverlap_Run", "Check overlap", asm, "THBIM.CallUIOverlap") { ToolTip = "Check overlap and highlight.", Image = File.Exists(icon16OV) ? new BitmapImage(new Uri(icon16OV)) : null };
            var pbID = new PushButtonData("LinkIDs_Run", "Linked IDS", asm, "THBIM.CallLinkIDs") { ToolTip = "Display ElementId of selected element.", Image = File.Exists(icon16ID) ? new BitmapImage(new Uri(icon16ID)) : null };
            var pbRS = new PushButtonData("RenameSheets_Run", "Rename Sheets", asm, "THBIM.CallUIReName") { ToolTip = "Rename Sheet and count SheetNumber.", Image = File.Exists(icon16RS) ? new BitmapImage(new Uri(icon16RS)) : null };
            var pdPFData = new PulldownButtonData("ProFilter_DD", "Pro Filter") { ToolTip = "Select elements by Category, Family, or Type.", Image = File.Exists(icon16PF) ? new BitmapImage(new Uri(icon16PF)) : null };
            var pbBubble = new PushButtonData("GridBubble_Run", "GridBubble", asm, "THBIM.CallUIbubble") { ToolTip = "Show/Hide Grid/Level heads.", Image = File.Exists(icon16Bubble) ? new BitmapImage(new Uri(icon16Bubble)) : null };
            var pbProSheet = new PushButtonData("ProSheets_Run", "ProSheets", asm, "THBIM.CallUIPROSHEET") { ToolTip = "Export View/Sheets to PDF, DWG.", LargeImage = File.Exists(iconProSheet32) ? new BitmapImage(new Uri(iconProSheet32)) : null };
            var pbQTOPRO = new PushButtonData("QTOPRO_Run", "QTOPRO", asm, "THBIM.CallUIQTOPRO") { ToolTip = "Extract QTO from Revit directly to Excel, complete with revision tracking and custom profile management.", LargeImage = File.Exists(iconQTOPRO32) ? new BitmapImage(new Uri(iconQTOPRO32)) : null };
            var pbSheetLink = new PushButtonData("SheetLink_Run", "SheetLink", asm, "THBIM.SheetLinkCommand") { ToolTip = "Open the SheetLink window.", LargeImage = File.Exists(iconSheetLink32) ? new BitmapImage(new Uri(iconSheetLink32)) : null };
            var pdDimGroup = new PulldownButtonData("DimensionTools_Pulldown", "Dim Tools") { ToolTip = "Dimensioning helper tools.", Image = File.Exists(icon16Dim) ? new BitmapImage(new Uri(icon16Dim)) : null };
            var pbDrop = new PushButtonData("FloorDrop_Run", "Detect Floor Drop", asm, "THBIM.CallUIFloordrop") { ToolTip = "Calculate floor drop values.", Image = File.Exists(icon16Drop) ? new BitmapImage(new Uri(icon16Drop)) : null };
            var pbzone = new PushButtonData("Zone_Run", "Set Zone", asm, "THBIM.CallUIZone") { ToolTip = "Set up Zone or somethings", Image = File.Exists(iconzone16) ? new BitmapImage(new Uri(iconzone16)) : null };

            panelAnnot.AddItem(pbCol);
            panelAnnot.AddSeparator();
            panelAnnot.AddItem(pbProSheet);
            panelAnnot.AddSeparator();
            panelAnnot.AddItem(pbSheetLink);
            panelAnnot.AddSeparator();
            panelAnnot.AddItem(pbQTOPRO);
            panelAnnot.AddSeparator();
            IList<RibbonItem> stackFilter = panelAnnot.AddStackedItems(pbRS, pbID, pdPFData);
            if (stackFilter.Count > 2 && stackFilter[2] is PulldownButton pdPF)
            {
                pdPF.AddPushButton(new PushButtonData("PF_Category", "By Category", asm, "THBIM.ProFilterCommand") { ToolTip = "Filter elements by Category" });
                pdPF.AddPushButton(new PushButtonData("PF_Family", "By Family", asm, "THBIM.ProFilterByFamilyCommand") { ToolTip = "Filter elements by Family name" });
                pdPF.AddPushButton(new PushButtonData("PF_Type", "By Type", asm, "THBIM.ProFilterByTypeCommand") { ToolTip = "Filter elements by Type name" });
            }
            panelAnnot.AddSeparator();

            IList<RibbonItem> stackDim = panelAnnot.AddStackedItems(pbBubble, pdDimGroup, pbDrop);
            if (stackDim.Count > 1 && stackDim[1] is PulldownButton pdDim)
            {
                pdDim.AddPushButton(new PushButtonData("AutoDimGrid_Sub", "Auto Dim Grid", asm, "THBIM.AutoDimGrid") { ToolTip = "Automatically create dimensions for grids.", Image = File.Exists(icon16Dim) ? new BitmapImage(new Uri(icon16Dim)) : null });
                pdDim.AddPushButton(new PushButtonData("CombineDim_Sub", "Combine Dim", asm, "THBIM.CombineDim") { ToolTip = "Combine multiple collinear dimensions into one.", Image = File.Exists(icon16CombineDim) ? new BitmapImage(new Uri(icon16CombineDim)) : null });
            }
            panelAnnot.AddSeparator();

            IList<RibbonItem> stackAlign = panelAnnot.AddStackedItems(pbCS, pdAlignData, pbzone);
            if (stackAlign.Count > 1 && stackAlign[1] is PulldownButton pullDown)
            {
                pullDown.AddPushButton(new PushButtonData("THBIM.AlignTags.Left", "Left", asm, "THBIM.AlignTagsLeftCommand") { LargeImage = File.Exists(iconAlignLeft) ? new BitmapImage(new Uri(iconAlignLeft)) : null });
                pullDown.AddPushButton(new PushButtonData("THBIM.AlignTags.Right", "Right", asm, "THBIM.AlignTagsRightCommand") { LargeImage = File.Exists(iconAlignRight) ? new BitmapImage(new Uri(iconAlignRight)) : null });
                pullDown.AddPushButton(new PushButtonData("THBIM.AlignTags.Top", "Top", asm, "THBIM.AlignTagsTopCommand") { LargeImage = File.Exists(iconAlignTop) ? new BitmapImage(new Uri(iconAlignTop)) : null });
                pullDown.AddPushButton(new PushButtonData("THBIM.AlignTags.Bottom", "Bottom", asm, "THBIM.AlignTagsBottomCommand") { LargeImage = File.Exists(iconAlignBottom) ? new BitmapImage(new Uri(iconAlignBottom)) : null });
                pullDown.AddPushButton(new PushButtonData("THBIM.AlignTags.Autoarrange", "Autoarrange", asm, "THBIM.ArrangeTagsNoCrossCommand") { LargeImage = File.Exists(iconArr) ? new BitmapImage(new Uri(iconArr)) : null });
                pullDown.AddSeparator();
                pullDown.AddPushButton(new PushButtonData("THBIM.AlignTags.LeaderAngle", "Leader Angle Setting", asm, "THBIM.LeaderAngleSettingCommand") { ToolTip = "Set leader elbow angle for align commands" });
            }

            panelAnnot.AddSeparator();
            IList<RibbonItem> stackParam = panelAnnot.AddStackedItems(pdParamData, pbOV);
            if (stackParam.Count > 0 && stackParam[0] is PulldownButton pdParam)
            {
                pdParam.AddPushButton(new PushButtonData("CombineParam_Sub", "Combine Param", asm, "THBIM.CallUIParam") { Image = File.Exists(iconParam16) ? new BitmapImage(new Uri(iconParam16)) : null, LargeImage = File.Exists(iconParam16) ? new BitmapImage(new Uri(iconParam16)) : null });
                pdParam.AddPushButton(new PushButtonData("SyncParam_Sub", "Sync Parameter", asm, "THBIM.CallUIParaSync") { Image = File.Exists(iconSync16) ? new BitmapImage(new Uri(iconSync16)) : null, LargeImage = File.Exists(iconSync16) ? new BitmapImage(new Uri(iconSync16)) : null });
            }

            ColorizePanel(tab, panelAnnotName, "#CE93D8");


            // ==========================================
            // MEP PANEL
            // ==========================================
            const string panelMEPName = "MEP";
            var panelMEP = app.GetRibbonPanels(tab).Find(p => p.Name == panelMEPName) ?? app.CreateRibbonPanel(tab, panelMEPName);

            try { SleeveChangeUpdater.Register(app); } catch {} // Register MEP sleeve updater

            string iconOP_32 = Path.Combine(dir, "Resources", "OPENING.png");
            string iconSL_16 = Path.Combine(dir, "Resources", "select_16.png");
            string iconCLD16 = Path.Combine(dir, "Resources", "cloud_16.png");
            string iconAC16 = Path.Combine(dir, "Resources", "accept_16.png");
            string iconSP16 = Path.Combine(dir, "Resources", "SplitPipe_16.png");
            string iconBOQ16 = Path.Combine(dir, "Resources", "BOQ_16.png");
            string iconSUP16 = Path.Combine(dir, "Resources", "hanger_16.png");

            var pbOP = new PushButtonData("OPENING_Run", "Create OPENING", asm, "THBIM.CallUIOP") { ToolTip = "Open the interface to create openings (sleeves).", LargeImage = File.Exists(iconOP_32) ? new BitmapImage(new Uri(iconOP_32)) : null };
            var pbSL = new PushButtonData("Select_Run", "Select Changed", asm, "THBIM.SelectChanged") { ToolTip = "Select all changed openings (NeedsReview=true) in the current view/model.", Image = File.Exists(iconSL_16) ? new BitmapImage(new Uri(iconSL_16)) : null };
            var pbCLD = new PushButtonData("CreateClouds_Run", "Create Clouds", asm, "THBIM.CreateClouds") { ToolTip = "Create revision clouds around changed openings in the active view.", Image = File.Exists(iconCLD16) ? new BitmapImage(new Uri(iconCLD16)) : null };
            var pbAC = new PushButtonData("AcceptChanged_Run", "Accept Changed", asm, "THBIM.Accept") { ToolTip = "Accept changes and update baseline for all changed openings..", Image = File.Exists(iconAC16) ? new BitmapImage(new Uri(iconAC16)) : null };
            var pbSP = new PushButtonData("SplitPipe_Run", "Split Pipe", asm, "THBIM.CallSP") { ToolTip = "Divide the pipe length according to the product length.", Image = File.Exists(iconSP16) ? new BitmapImage(new Uri(iconSP16)) : null };
            var pbQTO = new PushButtonData("MechanicalQTO_Run", "Mechanical QTO", asm, "THBIM.CallUIQTO") { ToolTip = "Select MEP elements in the model for quantity takeoff", Image = File.Exists(iconBOQ16) ? new BitmapImage(new Uri(iconBOQ16)) : null };
            var pbSUP = new PushButtonData("Hanger_Run", "Hanger", asm, "THBIM.CallUISUP") { ToolTip = "Generates hangers and supports for Pipes and Ducts.", Image = File.Exists(iconSUP16) ? new BitmapImage(new Uri(iconSUP16)) : null };

            panelMEP.AddItem(pbOP);          
            panelMEP.AddStackedItems(pbSL, pbCLD, pbAC);
            panelMEP.AddSeparator();
            panelMEP.AddStackedItems(pbSP, pbQTO, pbSUP);

            ColorizePanel(tab, panelMEPName, "#8FC67A");


            // ==========================================
            // STRUCTURE PANEL
            // ==========================================
            const string panelStructName = "Structure";
            var panelStruct = app.GetRibbonPanels(tab).Find(p => p.Name == panelStructName) ?? app.CreateRibbonPanel(tab, panelStructName);

            string iconDP32= Path.Combine(dir, "Resources", "Droppanel_32.png");
            string iconSPLIT16 = Path.Combine(dir, "Resources", "Splitcol_16.png");
            string iconLV16 = Path.Combine(dir, "Resources", "LevelRehost_16.png");
            string iconATP16 = Path.Combine(dir, "Resources", "AutoPile_16.png");

            var pbDP = new PushButtonData("AutoDropPanel_Run", "Create DropPanel", asm, "THBIM.CallUIDP") { ToolTip = "Automatically create Drop Panel or Pile Cap.", LargeImage = File.Exists(iconDP32) ? new BitmapImage(new Uri(iconDP32)) : null };
            var pbSplitCol = new PushButtonData("SplitColumnl_Run", "Split Columnl", asm, "THBIM.CallUISplit") { ToolTip = "Split columns by floor levels.", Image = File.Exists(iconSPLIT16) ? new BitmapImage(new Uri(iconSPLIT16)) : null };
            var pbLV = new PushButtonData("LevelRehost_Run", "Level Rehost", asm, "THBIM.CallUILV") { ToolTip = "Move elements to new level while keeping 3D position.", Image = File.Exists(iconLV16) ? new BitmapImage(new Uri(iconLV16)) : null };
            var pbATP = new PushButtonData("AutoPile_Run", "create Pile", asm, "THBIM.AutoPile") { ToolTip = "Automatically create Piling.", Image = File.Exists(iconATP16) ? new BitmapImage(new Uri(iconATP16)) : null };

            panelStruct.AddItem(pbDP);
            panelStruct.AddSeparator();
            panelStruct.AddStackedItems(pbATP, pbSplitCol, pbLV);

            ColorizePanel(tab, panelStructName, "#E38888");

            // ==========================================
            // TH TOOLS MEP TAB (MEP Accelerator)
            // ==========================================
            try { RegisterMepTab(app, asm, dir); } catch { }

            return Result.Succeeded;
        }

        private void RegisterMepTab(UIControlledApplication app, string asm, string dir)
        {
            const string mepTab = "TH Tools MEP";
            try { app.CreateRibbonTab(mepTab); } catch { }
            string mepIconDir = Path.Combine(dir, "Resources", "MEP");

            // --- LT Create ---
            var pCreate = app.CreateRibbonPanel(mepTab, "LT Create");
            pCreate.AddItem(MepBtn("Bloom", "Bloom", asm, "THBIM.MEP.Commands.BloomCommand", mepIconDir, "Bloom32.png"));
            pCreate.AddItem(MepBtn("Tap", "Tap", asm, "THBIM.MEP.Commands.TapCommand", mepIconDir, "Tap32.png"));
            pCreate.AddItem(MepBtn("ElbowDown45", "Elbow Down 45", asm, "THBIM.MEP.Commands.ElbowDown45Command", mepIconDir, "ElbowDown4532.png"));
            pCreate.AddItem(MepBtn("ElbowRight", "Elbow Right", asm, "THBIM.MEP.Commands.ElbowRightCommand", mepIconDir, "ElbowRight32.png"));
            pCreate.AddItem(MepBtn("ElbowDown", "Elbow Down", asm, "THBIM.MEP.Commands.ElbowDownCommand", mepIconDir, "ElbowDown32.png"));
            pCreate.AddItem(MepBtn("ElbowLeft", "Elbow Left", asm, "THBIM.MEP.Commands.ElbowLeftCommand", mepIconDir, "ElbowLeft32.png"));
            pCreate.AddItem(MepBtn("ElbowUp", "Elbow Up", asm, "THBIM.MEP.Commands.ElbowUpCommand", mepIconDir, "ElbowUp32.png"));
            ColorizePanel(mepTab, "LT Create", "#64B5F6");

            // --- LT Modify ---
            var pModify = app.CreateRibbonPanel(mepTab, "LT Modify");
            pModify.AddItem(MepBtn("FlipMultiple", "Flip Multiple", asm, "THBIM.MEP.Commands.FlipMultipleCommand", mepIconDir, "Microdesk.FlipMultiple32.png"));
            var rotateData = new SplitButtonData("RotateFittings", "Rotate");
            var rotateSplit = pModify.AddItem(rotateData) as SplitButton;
            if (rotateSplit != null)
            {
                string rotIcon = Path.Combine(mepIconDir, "RotateFitting32.png");
                if (File.Exists(rotIcon)) { var img = new BitmapImage(new Uri(rotIcon)); rotateSplit.Image = img; rotateSplit.LargeImage = img; }
                rotateSplit.AddPushButton(MepBtn("Rotate30", "Rotate 30°", asm, "THBIM.MEP.Commands.RotateFitting30Command", mepIconDir, "RotateFitting32.png"));
                rotateSplit.AddPushButton(MepBtn("Rotate45", "Rotate 45°", asm, "THBIM.MEP.Commands.RotateFitting45Command", mepIconDir, "RotateFitting32.png"));
                rotateSplit.AddPushButton(MepBtn("Rotate60", "Rotate 60°", asm, "THBIM.MEP.Commands.RotateFitting60Command", mepIconDir, "RotateFitting32.png"));
                rotateSplit.AddPushButton(MepBtn("Rotate75", "Rotate 75°", asm, "THBIM.MEP.Commands.RotateFitting75Command", mepIconDir, "RotateFitting32.png"));
                rotateSplit.AddPushButton(MepBtn("Rotate90", "Rotate 90°", asm, "THBIM.MEP.Commands.RotateFitting90Command", mepIconDir, "RotateFitting32.png"));
            }
            pModify.AddItem(MepBtn("DeleteSystem", "Delete System", asm, "THBIM.MEP.Commands.DeleteSystemCommand", mepIconDir, "DeleteSystem32.png"));
            pModify.AddItem(MepBtn("Disconnect", "Disconnect", asm, "THBIM.MEP.Commands.DisconnectCommand", mepIconDir, "Disconnect32.png"));
            pModify.AddItem(MepBtn("MoveConnect", "Move Connect", asm, "THBIM.MEP.Commands.MoveConnectCommand", mepIconDir, "MoveConnect32.png"));
            pModify.AddItem(MepBtn("MoveConnectAlign", "Mo-Co Align", asm, "THBIM.MEP.Commands.MoveConnectAlignCommand", mepIconDir, "MoveConnectAlign32.png"));
            ColorizePanel(mepTab, "LT Modify", "#FFB74D");

            // --- LT Align ---
            var pAlign = app.CreateRibbonPanel(mepTab, "LT Align");
            pAlign.AddItem(MepBtn("Align3D", "Align in 3D", asm, "THBIM.MEP.Commands.AlignIn3DCommand", mepIconDir, "AlignIn3D32.png"));
            pAlign.AddItem(MepBtn("AlignBranch", "Align Branch", asm, "THBIM.MEP.Commands.AlignBranchCommand", mepIconDir, "BranchAlignLite32.png"));
            pAlign.AddItem(MepBtn("AlignBranchPlus", "Align Branch+", asm, "THBIM.MEP.Commands.AlignBranchPlusCommand", mepIconDir, "BranchAlign32.png"));
            ColorizePanel(mepTab, "LT Align", "#81C784");

            // --- LT Tools ---
            var pTools = app.CreateRibbonPanel(mepTab, "LT Tools");
            pTools.AddItem(MepBtn("SumParam", "Sum Param.", asm, "THBIM.MEP.Commands.SumParameterCommand", mepIconDir, "Sum32.png"));
            pTools.AddItem(MepBtn("Section", "Section", asm, "THBIM.MEP.Commands.SectionCommand", mepIconDir, "Section32.png"));
            ColorizePanel(mepTab, "LT Tools", "#BA68C8");
        }

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

        public Result OnShutdown(UIControlledApplication app)
        {
            try { SleeveChangeUpdater.Unregister(); } catch {}
            return Result.Succeeded;
        }

        public void ColorizePanel(string tabName, string panelName, string hexColor)
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
}