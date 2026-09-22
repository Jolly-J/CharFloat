# WPS 宿主对象模型能力地图（2.1.0-p0p1 / p5 / mcp-sweep）

> ⚠️ **本报告未由原子代理写完**：它在测绘中途被中断（WPS 进程于 16:10 崩溃、Excel 组件掉线）。
> 已完成的探测数据保存在 [06-wps-api-map-raw-excel.json](06-wps-api-map-raw-excel.json)，下面的成员清单由协调方从该数据**自动重建**。
> **Word / PPT 两组件未测绘。**

## 附：表格（ET）对象模型成员清单（由原始探测数据重建）

> 本节由协调方从 `06-wps-api-map-raw-excel.json` **自动重建**。原子代理在测绘中途被中断（WPS 于 16:10 崩溃、Excel 组件掉线），报告正文未写完，但探测数据已落盘，故此处保留可用的部分。
> 覆盖 **48** 个表达式；**Word / PPT 两组件未测绘**（中断前未开始）。

### 概览

| 表达式 | 成员数 | 方法数 | 属性数 |
|---|---:|---:|---:|
| `wb` | 226 | 87 | 139 |
| `wb.Worksheets.Item("Data")` | 135 | 62 | 73 |
| `wb.Worksheets.Item("Probe")` | 135 | 62 | 73 |
| `wb.Worksheets.Item("Data").Range("A1:B2")` | 215 | 114 | 101 |
| `wb.Worksheets.Item("Data").Cells` | 0 | 0 | 0 |
| `wb.Worksheets.Item("Data").Rows` | 0 | 0 | 0 |
| `wb.Worksheets.Item("Data").Columns` | 0 | 0 | 0 |
| `wb.Worksheets.Item("Data").ListObjects.Item(1)` | 0 | 0 | 0 |
| `wb.Worksheets.Item("Data").ListObjects` | 0 | 0 | 0 |
| `wb.Worksheets.Item("Probe").PivotTables()` | 0 | 0 | 0 |
| `wb.PivotCaches()` | 0 | 0 | 0 |
| `wb.Worksheets.Item("Data").Range("D2:D11").FormatConditions.Item(1)` | 0 | 0 | 0 |
| `wb.Worksheets.Item("Data").Range("D2:D11").FormatConditions` | 0 | 0 | 0 |
| `wb.Names` | 0 | 0 | 0 |
| `wb.Names.Item(1)` | 0 | 0 | 0 |
| `wb.Worksheets.Item("Data").Range("B2:B11").Validation` | 0 | 0 | 0 |
| `wb.Worksheets.Item("Data").Range("A1").Comment` | 0 | 0 | 0 |
| `wb.Worksheets.Item("Data").Range("A2").CommentThreaded` | 0 | 0 | 0 |
| `wb.Worksheets.Item("Data").Comments` | 0 | 0 | 0 |
| `wb.Worksheets.Item("Data").CommentsThreaded` | 0 | 0 | 0 |
| `wb.Worksheets.Item("Data").Sort` | 0 | 0 | 0 |
| `wb.Worksheets.Item("Data").Sort.SortFields` | 0 | 0 | 0 |
| `wb.Worksheets.Item("Data").AutoFilter` | 0 | 0 | 0 |
| `wb.Worksheets.Item("Data").AutoFilter.Filters` | 0 | 0 | 0 |
| `wb.Worksheets.Item("Data").PageSetup` | 0 | 0 | 0 |
| `wb.Worksheets.Item("Probe").Shapes.Item(1)` | 0 | 0 | 0 |
| `wb.Worksheets.Item("Probe").Shapes` | 0 | 0 | 0 |
| `wb.Worksheets.Item("Probe").ChartObjects().Item(1)` | 0 | 0 | 0 |
| `wb.Worksheets.Item("Probe").ChartObjects().Item(1).Chart` | 0 | 0 | 0 |
| `wb.Worksheets.Item("Data").QueryTables` | 0 | 0 | 0 |
| `wb.Connections` | 0 | 0 | 0 |
| `app.Windows.Item(1)` | 0 | 0 | 0 |
| `app.Windows.Item(1).Panes` | 0 | 0 | 0 |
| `wb.Worksheets.Item("Data").Tab` | 0 | 0 | 0 |
| `wb.Worksheets.Item("Data").Outline` | 0 | 0 | 0 |
| `wb.Worksheets.Item("Data").Hyperlinks` | 0 | 0 | 0 |
| `wb.Worksheets.Item("Data").Protection` | 0 | 0 | 0 |
| `wb.Worksheets.Item("Data").CustomProperties` | 0 | 0 | 0 |
| `wb.Worksheets.Item("Data").NamedSheetViews` | 0 | 0 | 0 |
| `wb.Worksheets.Item("Data").Names` | 0 | 0 | 0 |
| `wb.Styles` | 0 | 0 | 0 |
| `wb.Worksheets.Item("Data").UsedRange` | 0 | 0 | 0 |
| `wb.Worksheets.Item("Data").Range("A1").Font` | 0 | 0 | 0 |
| `wb.Worksheets.Item("Data").Range("A1").Interior` | 0 | 0 | 0 |
| `wb.Worksheets.Item("Data").Range("A1").Borders` | 0 | 0 | 0 |
| `wb.Worksheets.Item("Data").Range("A1").NumberFormat` | 0 | 0 | 0 |
| `app.ActiveWindow` | 0 | 0 | 0 |
| `app.CommandBars` | 0 | 0 | 0 |

### 关键对象成员明细

#### `wb`（226 个成员）

- **方法（87）**：Activate、ChangeFileAccess、Close、NewWindow、ForwardMailer、SetLinkOnData、LinkInfo、LinkSources、Route、MergeWorkbook、CheckIn、Dummy17、HighlightChangesOptions、__PrintOut、PrintPreview、_ProtectSharing、RefreshAll、ReplyAll、RemoveUser、__SaveAs、ChangeLink、CheckInWithVersion、Unprotect、UnprotectSharing、UpdateFromFile、Post、Reply、PivotTableWizard、AcceptAllChanges、FollowHyperlink、AddToFavorites、_PrintOut、WebPagePreview、sblt、Dummy16、SaveAsBinaryString、_SaveAs、CanCheckIn、SendForReview、PublishToPBI、DeleteNumberFormat、ReplyWithChanges、SendMail、EndReview、ExportAsFixedFormat、ApplyTheme、SetPasswordEncryptionOptions、ResetColors、_ExportAsFixedFormat、BreakLink、Protect、Dummy26、Save、Colors、EnableConnections、XmlImport、XmlImportXml、PivotCaches、SaveAsXMLData、LockServerFile、GetWorkflowTasks、GetWorkflowTemplates、PrintOut、SaveCopyAs、UpdateLink、RecheckSmartTags、SaveAs、RejectAllChanges、CreateForecastSheet、ProtectSharing、OpenLinks、RemoveDocumentInformation、ExclusiveAccess、_Protect、ConvertComments、SaveAsUrl、PurgeChangeHistoryNow、PublishToDocs、Dummy27、GetWorkbookEx、ToggleFormsDesign、LookUpInDocs、ReloadAs、SendFaxOverInternet、InvalidateRightsInfo、RunAutoMacros、SendMailer
- **属性（139）**：AcceptLabelsInFormulas(boolean)、ActiveChart(object)、ActiveSheet(object)、Author(string)、AutoUpdateFrequency(number)、AutoUpdateSaveChanges(object)、ChangeHistoryDuration(number)、CodeName(string)、CommandBars(object,186)、Comments(string)、ConflictResolution(number)、CreateBackup(boolean)、CustomDocumentProperties(object,3)、BuiltinDocumentProperties(object,31)、DialogSheets(object,0)、DisplayDrawingObjects(number)、DocumentInspectors(object,1)、FullName(string)、HasRoutingSlip(boolean)、Modules(object,0)、MultiUserEditing(boolean)、Keywords(string)、Path(string)、ProtectWindows(boolean)、ProtectStructure(boolean)、Charts(object,0)、Routed(boolean)、RoutingSlip(object)、Saved(boolean)、Sheets(object,2)、Title(string)、Parent(object)、SmartDocument(object)、OnSave(string)、PasswordEncryptionKeyLength(number)、ShowPivotTableFieldList(boolean)、UpdateRemoteReferences(boolean)、UserStatus(object)、CustomViews(object,0)、Windows(object,1)、Worksheets(object,2)、WriteReservedBy(string)、ListChangesOnNewSheet(boolean)、_ReadOnlyRecommended(boolean)、PublishObjects(object,0)、HTMLProject(object)、CalculationVersion(number)、_CodeName(string)、WebOptions(object)、UserControl(boolean)、JSProject(object)、OnSheetActivate(string)、ServerViewableItems(object,1)、Container(object)、ReadOnly(boolean)、FileFormat(number)、OnSheetDeactivate(string)、WriteReserved(boolean)、PersonalViewPrintSettings(boolean)、SensitivityLabel(object)、SaveLinkValues(boolean)、HasPassword(boolean)、SharedWorkspace(object)、IconSets(object,20)、Application(object)、Password(string)、PrecisionAsDisplayed(boolean)、WritePassword(string)、Subject(string)、RemovePersonalInformation(boolean)、PasswordEncryptionAlgorithm(string)、DefaultPivotTableStyle(string)、UseWholeCellCriteria(boolean)、Excel4IntlMacroSheets(object,1)、PasswordEncryptionFileProperties(boolean)、ReadOnlyRecommended(boolean)、UpdateLinks(number)、SmartTagOptions(object)、Excel4MacroSheets(object,0)、HighlightChangesOnScreen(boolean)、XmlNamespaces(object,0)、PersonalViewListSettings(boolean)、XmlMaps(object,0)、PasswordEncryptionProvider(string)、DocumentLibraryVersions(object,1)、InactiveListBorderVisible(boolean)、ForceFullCalculation(boolean)、DisplayInkComments(boolean)、ContentTypeProperties(object,1)、KeepChangeHistory(boolean)、EnvelopeVisible(boolean)、Styles(object,49)、Sync(object)、FullNameURLEncoded(string)、TemplateRemoveExtData(boolean)、Model(object)、TableStyles(object,284)、CheckCompatibility(boolean)、HasVBProject(object)、Date1904(boolean)、DefaultSlicerStyle(string)、CustomXMLParts(object,3)、CaseSensitive(boolean)、Research(object)、DefaultTableStyle(string)、EncryptionProvider(string)、Signatures(object,1)、IsAddin(boolean)、ActiveSlicer(object)、ConnectionsDisabled(boolean)、ShowPivotChartActiveFields(boolean)、Queries(object,0)、IsInplace(boolean)、Permission(object,0)、DefaultTimelineStyle(object)、AccuracyVersion(number)、Final(boolean)、UseWildcards(boolean)、RevisionNumber(number)、PivotTables(object)、Creator(number)、HasMailer(boolean)、WorkIdentity(string)、Mailer(object)、Name(string)、AutoSaveOn(boolean)、SlicerCaches(object,0)、VBASigned(boolean)、Excel8CompatibilityMode(boolean)、DoNotPromptForConvert(boolean)、ExternalCodeServiceTimeout(number)、Connections(object,0)、Theme(object)、ServerPolicy(object,1)、ChartDataPointTrack(boolean)、ShowConflictHistory(boolean)、Names(object,2)、EnableAutoRecover(boolean)、CompatibilityVersion(number)

#### `wb.Worksheets.Item("Data").Range("A1:B2")`（215 个成员）

- **方法（114）**：Address、AdvancedFilter、ApplyNames、ApplyOutlineStyles、AutoComplete、AutoFill、_AutoFilter、AutoFit、AutoOutline、Calculate、Parse、Characters、_ClearContents、DataSeries、ClearFormats、ClearNotes、Ungroup、ShowErrors、ColumnDifferences、CopyFromRecordset、DataTypeToText、RemoveControls、AllocateChanges、_Default、TogglePythonMarshalMode、DialogBox、FindPrevious、FillLeft、SetPhonetic、FunctionWizard、AddCommentThreaded、ClearOutline、Cut、Insert、Dirty、Merge、Copy、UnMerge、Run、UpdatePictureInCellAlternativeText、Sort、NoteText、_PrintOut、CreatePublisher、Offset、FillUp、_PasteSpecial、Range、RemoveSubtotal、_Replace、AutoFormat、Select、Show、_Sort、SortSpecial、ConvertToLinkedDataType、CopyPicture、GetRangeEx、FlashFill、FindNext、ClearContents、Value、FillDown、_BorderAround、AddComment、FillRight、Speak、NavigateArrow、PasteSpecial、PlacePictureOverCells、PastePictureInCell、RowDifferences、SpecialCells、__PrintOut、PrintPreview、_ExportAsFixedFormat、GoalSeek、TextToColumns、ClearHyperlinks、SetCellDataTypeFromCell、AddressLocal、Clear、BorderAround、Find、ExportAsFixedFormat、ShowDependents、PrintOut、ShowCard、Item、Table、RefreshLinkedDataType、Activate、ClearComments、Group、Justify、CalculateRowMajorOrder、Consolidate、SubscribeTo、RemoveDuplicates、CreateNames、End、CheckSpelling、AutoFilter、Delete、InsertPictureInCell、ResetContents、EditionOptions、Resize、Replace、ListNames、ShowPrecedents、DiscardChanges、InsertIndent、Subtotal
- **属性（101）**：Creator(number)、Areas(object,1)、Borders(object,8)、Cells(object,4)、Previous(object)、Interior(object)、Columns(object,2)、Count(number)、Errors(object)、MergeCells(boolean)、CurrentArray(object,0)、Formula2Local(object)、Font(object)、EntireColumn(object,2)、FormulaLabel(number)、FormulaHidden(boolean)、SmartTags(object,1)、PivotTable(object)、HorizontalAlignment(number)、Precedents(object)、ID(object)、ListHeaderRows(number)、Hyperlinks(object,0)、LocationInTable(object)、Locked(boolean)、HasArray(boolean)、HasFormula(boolean)、Summary(object)、Worksheet(object)、MergeArea(object)、Name(object)、Next(object,1)、DirectPrecedents(object)、SpillingToRange(object)、Text(object)、FormulaR1C1(object)、NumberFormatLocal(string)、Left(number)、PageBreak(number)、FormulaArray(object)、Height(number)、Parent(object)、Hidden(object)、PivotField(object)、Row(number)、RowHeight(number)、Rows(object,2)、PivotItem(object)、FormulaR1C1Local(object)、OutlineLevel(object)、ShowDetail(object)、VerticalAlignment(number)、CountLarge(number)、CurrentRegion(object,77)、Style(object)、XPath(object)、Formula2(object)、ListObject(object)、Top(number)、UseStandardHeight(boolean)、UseStandardWidth(object)、RealValue2(object)、Value2(object)、RangeEx(object)、Width(number)、ColumnWidth(number)、FormatConditions(object,0)、Phonetics(object,1)、Orientation(number)、PivotCell(object)、Phonetic(object)、Formula2R1C1(object)、ReadingOrder(number)、PrefixCharacter(string)、ServerActions(object,1)、AddIndent(number)、Column(number)、CellControl(object)、SparklineGroups(object,0)、MDX(string)、DisplayFormat(object)、Formula(object)、HasRichDataType(object)、WrapText(boolean)、EntireRow(object,2)、IndentLevel(number)、AllowEdit(boolean)、SavedAsArray(object)、Application(object)、SoundNote(object)、CommentThreaded(object)、Validation(object)、HasSpill(boolean)、LinkedDataTypeState(object)、FormulaLocal(object)、NumberFormat(string)、Formula2R1C1Local(object)、QueryTable(object)、ShrinkToFit(boolean)、SpillParent(object)、Comment(object)

#### `wb.Worksheets.Item("Data")`（135 个成员）

- **方法（62）**：Activate、_Protect、Unprotect、SetBackgroundPicture、ChartObjects、CheckSpelling、ClearArrows、Move、_SaveAs、Drawings、_CheckSpelling、DrawingObjects、Evaluate、_Evaluate、ResetAllPageBreaks、GroupBoxes、Lines、GroupObjects、ListBoxes、OLEObjects、EditBoxes、Delete、Buttons、Labels、Hide、Ovals、Paste、Pictures、Rectangles、ShowDataForm、Spinners、TextBoxes、PivotTableWizard、CheckBoxes、_PrintOut、CircleInvalid、ExportToPNG、ShowAllData、Copy、Select、DropDowns、Scenarios、XmlMapQuery、_PasteSpecial、XmlDataQuery、Calculate、_ExportAsFixedFormat、Range、PasteSpecial、ExportAsFixedFormat、__SaveAs、ClearCircles、Show、__PrintOut、PrintPreview、OptionButtons、ScrollBars、PrintOut、Arcs、Protect、PivotTables、SaveAs
- **属性（73）**：Application(object)、Parent(object)、CodeName(string)、_CodeName(string)、Index(number)、Name(string)、Next(object)、OnSheetActivate(string)、PageSetup(object)、ProtectDrawingObjects(boolean)、Shapes(object,1)、TransitionExpEval(boolean)、AutoFilterMode(boolean)、Visible(number)、Cells(object,17179869184)、AutoFilter(object)、ConsolidationOptions(object)、ConsolidationSources(object)、DisplayAutomaticPageBreaks(boolean)、OnSheetDeactivate(string)、OnCalculate(string)、EnableOutlining(boolean)、_DisplayRightToLeft(number)、Scripts(object,0)、EnablePivotTable(boolean)、FilterMode(boolean)、ConsolidationFunction(number)、SmartTags(object,1)、ListObjects(object,1)、ScrollArea(string)、Tab(object)、Names(object,1)、OnData(string)、Previous(object)、Outline(object)、EnableSelection(number)、MailEnvelope(object)、ProtectContents(boolean)、StandardHeight(number)、StandardWidth(number)、QueryTables(object,0)、DisplayRightToLeft(boolean)、_Sort(object)、HPageBreaks(object,0)、DisplayPageBreaks(boolean)、Protection(object)、Type(number)、CustomProperties(object,0)、Rows(object,1048576)、VPageBreaks(object,0)、TransitionFormEntry(boolean)、UsedRange(object,77)、CircularReference(object)、Hyperlinks(object,0)、ProtectionMode(boolean)、Creator(number)、WorksheetEx(object)、CommentsThreaded(object,1)、DefaultButton(object)、OnDoubleClick(string)、DialogFrame(object)、Columns(object,16384)、OnEntry(string)、Focus(object)、PrintedCommentPages(number)、EnableFormatConditionsCalculation(boolean)、NamedSheetViews(object,1)、_AutoFilter(object)、ProtectScenarios(boolean)、Comments(object,1)、EnableAutoFilter(boolean)、EnableCalculation(boolean)、Sort(object)

#### `wb.Worksheets.Item("Probe")`（135 个成员）

- **方法（62）**：Activate、_Protect、Unprotect、SetBackgroundPicture、ChartObjects、CheckSpelling、ClearArrows、Move、_SaveAs、Drawings、_CheckSpelling、DrawingObjects、Evaluate、_Evaluate、ResetAllPageBreaks、GroupBoxes、Lines、GroupObjects、ListBoxes、OLEObjects、EditBoxes、Delete、Buttons、Labels、Hide、Ovals、Paste、Pictures、Rectangles、ShowDataForm、Spinners、TextBoxes、PivotTableWizard、CheckBoxes、_PrintOut、CircleInvalid、ExportToPNG、ShowAllData、Copy、Select、DropDowns、Scenarios、XmlMapQuery、_PasteSpecial、XmlDataQuery、Calculate、_ExportAsFixedFormat、Range、PasteSpecial、ExportAsFixedFormat、__SaveAs、ClearCircles、Show、__PrintOut、PrintPreview、OptionButtons、ScrollBars、PrintOut、Arcs、Protect、PivotTables、SaveAs
- **属性（73）**：Application(object)、Parent(object)、CodeName(string)、_CodeName(string)、Index(number)、Name(string)、Next(object)、OnSheetActivate(string)、PageSetup(object)、ProtectDrawingObjects(boolean)、Shapes(object,5)、TransitionExpEval(boolean)、AutoFilterMode(boolean)、Visible(number)、Cells(object,17179869184)、AutoFilter(object)、ConsolidationOptions(object)、ConsolidationSources(object)、DisplayAutomaticPageBreaks(boolean)、OnSheetDeactivate(string)、OnCalculate(string)、EnableOutlining(boolean)、_DisplayRightToLeft(number)、Scripts(object,0)、EnablePivotTable(boolean)、FilterMode(boolean)、ConsolidationFunction(number)、SmartTags(object,1)、ListObjects(object,0)、ScrollArea(string)、Tab(object)、Names(object,0)、OnData(string)、Previous(object)、Outline(object)、EnableSelection(number)、MailEnvelope(object)、ProtectContents(boolean)、StandardHeight(number)、StandardWidth(number)、QueryTables(object,0)、DisplayRightToLeft(boolean)、_Sort(object)、HPageBreaks(object,0)、DisplayPageBreaks(boolean)、Protection(object)、Type(number)、CustomProperties(object,0)、Rows(object,1048576)、VPageBreaks(object,1)、TransitionFormEntry(boolean)、UsedRange(object,10)、CircularReference(object)、Hyperlinks(object,1)、ProtectionMode(boolean)、Creator(number)、WorksheetEx(object)、CommentsThreaded(object,1)、DefaultButton(object)、OnDoubleClick(string)、DialogFrame(object)、Columns(object,16384)、OnEntry(string)、Focus(object)、PrintedCommentPages(number)、EnableFormatConditionsCalculation(boolean)、NamedSheetViews(object,1)、_AutoFilter(object)、ProtectScenarios(boolean)、Comments(object,0)、EnableAutoFilter(boolean)、EnableCalculation(boolean)、Sort(object)

### 其余表达式（成员较少）

- `wb.Worksheets.Item("Data").Cells`：0 个成员
- `wb.Worksheets.Item("Data").Rows`：0 个成员
- `wb.Worksheets.Item("Data").Columns`：0 个成员
- `wb.Worksheets.Item("Data").ListObjects.Item(1)`：0 个成员
- `wb.Worksheets.Item("Data").ListObjects`：0 个成员
- `wb.Worksheets.Item("Probe").PivotTables()`：0 个成员
- `wb.PivotCaches()`：0 个成员
- `wb.Worksheets.Item("Data").Range("D2:D11").FormatConditions.Item(1)`：0 个成员
- `wb.Worksheets.Item("Data").Range("D2:D11").FormatConditions`：0 个成员
- `wb.Names`：0 个成员
- `wb.Names.Item(1)`：0 个成员
- `wb.Worksheets.Item("Data").Range("B2:B11").Validation`：0 个成员
- `wb.Worksheets.Item("Data").Range("A1").Comment`：0 个成员
- `wb.Worksheets.Item("Data").Range("A2").CommentThreaded`：0 个成员
- `wb.Worksheets.Item("Data").Comments`：0 个成员
- `wb.Worksheets.Item("Data").CommentsThreaded`：0 个成员
- `wb.Worksheets.Item("Data").Sort`：0 个成员
- `wb.Worksheets.Item("Data").Sort.SortFields`：0 个成员
- `wb.Worksheets.Item("Data").AutoFilter`：0 个成员
- `wb.Worksheets.Item("Data").AutoFilter.Filters`：0 个成员
- `wb.Worksheets.Item("Data").PageSetup`：0 个成员
- `wb.Worksheets.Item("Probe").Shapes.Item(1)`：0 个成员
- `wb.Worksheets.Item("Probe").Shapes`：0 个成员
- `wb.Worksheets.Item("Probe").ChartObjects().Item(1)`：0 个成员
- `wb.Worksheets.Item("Probe").ChartObjects().Item(1).Chart`：0 个成员
- `wb.Worksheets.Item("Data").QueryTables`：0 个成员
- `wb.Connections`：0 个成员
- `app.Windows.Item(1)`：0 个成员
- `app.Windows.Item(1).Panes`：0 个成员
- `wb.Worksheets.Item("Data").Tab`：0 个成员
- `wb.Worksheets.Item("Data").Outline`：0 个成员
- `wb.Worksheets.Item("Data").Hyperlinks`：0 个成员
- `wb.Worksheets.Item("Data").Protection`：0 个成员
- `wb.Worksheets.Item("Data").CustomProperties`：0 个成员
- `wb.Worksheets.Item("Data").NamedSheetViews`：0 个成员
- `wb.Worksheets.Item("Data").Names`：0 个成员
- `wb.Styles`：0 个成员
- `wb.Worksheets.Item("Data").UsedRange`：0 个成员
- `wb.Worksheets.Item("Data").Range("A1").Font`：0 个成员
- `wb.Worksheets.Item("Data").Range("A1").Interior`：0 个成员
- `wb.Worksheets.Item("Data").Range("A1").Borders`：0 个成员
- `wb.Worksheets.Item("Data").Range("A1").NumberFormat`：0 个成员
- `app.ActiveWindow`：0 个成员
- `app.CommandBars`：0 个成员

### 中断证据

最后一次探测返回：`HTTP 422: {"success":false,"error":"加载项断开，执行结果未知，请先读回确认"}`
