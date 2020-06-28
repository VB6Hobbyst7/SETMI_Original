Public Class Main

#Region "ArcGIS"

    Public Function CreateTable(ByVal Path As String) As ESRI.ArcGIS.Geodatabase.ITable
        Dim GPUtilities As ESRI.ArcGIS.Geoprocessing.IGPUtilities = New ESRI.ArcGIS.Geoprocessing.GPUtilities
        Return GPUtilities.OpenTableFromString(Path) 'Opens a table, this is used to load .xls files and to populate other database tables JBB
    End Function

    Public Function CreateRasterDataset(ByVal Path As String) As ESRI.ArcGIS.Geodatabase.IRasterDataset
        Dim GPUtilities As ESRI.ArcGIS.Geoprocessing.IGPUtilities = New ESRI.ArcGIS.Geoprocessing.GPUtilities
        Return GPUtilities.OpenRasterDatasetFromString(Path) 'Opens a new raster dataset with a given path name JBB
    End Function

    Public Sub AddRasterDatasetRasterToMap(ByVal RasterDataset As ESRI.ArcGIS.Geodatabase.IRasterDataset, ByVal Map As ESRI.ArcGIS.Carto.IMap)
        Dim RasterLayer As ESRI.ArcGIS.Carto.IRasterLayer = New ESRI.ArcGIS.Carto.RasterLayerClass
        RasterLayer.CreateFromDataset(RasterDataset)
        Map.AddLayer(RasterLayer)
    End Sub

    Public Function CreateIntersectRaster(ByVal InputRasterPath() As String, ByVal InputRasters() As ESRI.ArcGIS.Geodatabase.IRasterDataset, ByVal IntersectRasterPath As String) As ESRI.ArcGIS.Geodatabase.IRasterDataset
        Dim MapAlgebraOperation As ESRI.ArcGIS.SpatialAnalyst.IMapAlgebraOp = New ESRI.ArcGIS.SpatialAnalyst.RasterMapAlgebraOpClass
        Dim Environment As ESRI.ArcGIS.GeoAnalyst.IRasterAnalysisEnvironment = MapAlgebraOperation
        Dim WorkspaceFactory As ESRI.ArcGIS.Geodatabase.IWorkspaceFactory = New ESRI.ArcGIS.DataSourcesRaster.RasterWorkspaceFactory
        Dim OutWorkspace As ESRI.ArcGIS.Geodatabase.IWorkspace = WorkspaceFactory.OpenFromFile(IO.Path.GetDirectoryName(IntersectRasterPath), 0)
        Environment.OutWorkspace = OutWorkspace 'Set the output workspace JBB
        Environment.OutSpatialReference = CType(CreateRasterDataset(InputRasterPath(InputRasterPath.Count - 1)), ESRI.ArcGIS.Geodatabase.IGeoDataset).SpatialReference 'CType converts data types JBB, I think this creates a new raster and then grabs the spatial reference based on the Wilting point image JBB
        Environment.SetExtent(ESRI.ArcGIS.GeoAnalyst.esriRasterEnvSettingEnum.esriRasterEnvMinOf) 'Set raster extent as min of all input rasters JBB
        Environment.SetCellSize(ESRI.ArcGIS.GeoAnalyst.esriRasterEnvSettingEnum.esriRasterEnvMinOf) 'Set raster cell size as min of all input rasters JBB

        Dim Operation As New System.Text.StringBuilder("SetNull(") 'Starts building the code for intersecting rastser JBB
        For P = 0 To InputRasterPath.Count - 1
            InputRasters(P) = CreateRasterDataset(InputRasterPath(P)) 'Creates a raster dataset for each input raster JBB
            MapAlgebraOperation.BindRaster(InputRasters(P), "Raster" & P) 'delete the line below from + sign 'Binds the name "RasterP" (where P is the input raster number) to the newly created input raster databases JBB
            'Operation.Append(" Con( IsNull( [" & "Raster" & P & "] ) , 1 , 0 ) +  Con( [" & "Raster" & P & "] < 0 , 1 , 0 ) ")
            'Operation.Append(" Con( IsNull( """ & IO.Path.GetFileName(InputRasterPath(P)) & """ ) , 1 , 0 )  ")
            Operation.Append(" Con( IsNull( [" & "Raster" & P & "] ) , 1 , 0 )  ") 'Adds logic for next raster JBB
            If Not P = InputRasters.Count - 1 Then Operation.Append("+ ") 'Adds logic for next raster JBB
        Next
        Operation.Append("> 0 , 1 )") 'Finishes building code for intersecting raster, if all rasters are intersecting the value is one otherwise it is Null JBB

        'If ExistsArcGISFile(IntersectRasterPath) Then DeleteArcGISFile(IntersectRasterPath)

        Dim OutRaster As ESRI.ArcGIS.Geodatabase.IRaster = MapAlgebraOperation.Execute(Operation.ToString) 'Executes the intersection string that was just created JBB
        Dim Save As ESRI.ArcGIS.Geodatabase.ISaveAs2 = OutRaster 'Save the output intersecting image JBB
        Save.SaveAs(IO.Path.GetFileName(IntersectRasterPath), OutWorkspace, "IMAGINE Image") 'Save as .img JBB
        Dim IntersectRaster = CreateRasterDataset(IntersectRasterPath) 'Create a raster dataset from the new intersect raster 'JBB
        Return IntersectRaster
    End Function


    Public Function GetRasterCursorIterations(ByVal IntersectRaster As ESRI.ArcGIS.Geodatabase.IRasterDataset) As Long
        Dim IntersectRaster2 As ESRI.ArcGIS.DataSourcesRaster.IRaster2 = CType(CType(IntersectRaster.CreateDefaultRaster, ESRI.ArcGIS.Geodatabase.IRaster), ESRI.ArcGIS.DataSourcesRaster.IRaster2)
        Dim IntersectRasterCursor As ESRI.ArcGIS.Geodatabase.IRasterCursor = IntersectRaster2.CreateCursorEx(Nothing)
        Dim I As Long = 0
        Do
            I += 1
        Loop While IntersectRasterCursor.Next = True
        Return I
    End Function

    Public Sub ExtractRaster(ByVal InputRasterPath As String, ByRef InputRaster As ESRI.ArcGIS.Geodatabase.IRasterDataset, ByVal IntersectRasterPath As String, ByVal IntersectRaster As ESRI.ArcGIS.Geodatabase.IRasterDataset)
        Dim ExtractionOperation As ESRI.ArcGIS.SpatialAnalyst.IExtractionOp = New ESRI.ArcGIS.SpatialAnalyst.RasterExtractionOp 'References the RasterExtractionOP interface JBB
        Dim WorkspaceFactory As ESRI.ArcGIS.Geodatabase.IWorkspaceFactory = New ESRI.ArcGIS.DataSourcesRaster.RasterWorkspaceFactory 'Sets a workspace for extraction JBB
        '*****TESTING
        'Dim OutWorkspace As ESRI.ArcGIS.Geodatabase.IWorkspace = WorkspaceFactory.OpenFromFile(IO.Path.GetDirectoryName(IntersectRasterPath), 0) 'Sets an output workspace for the extraction JBB
        Dim FileName As String = IO.Path.ChangeExtension(IO.Path.GetTempFileName, ".img")
        Dim OutWorkspace As ESRI.ArcGIS.Geodatabase.IWorkspace = WorkspaceFactory.OpenFromFile(IO.Path.GetDirectoryName(FileName), 0) 'Sets an output workspace for the extraction JBB
        '******END TESTING
        Dim Environment As ESRI.ArcGIS.GeoAnalyst.IRasterAnalysisEnvironment = ExtractionOperation 'Sets the extraction environment JBB
        Environment.OutWorkspace = OutWorkspace 'Assigns the output workspace JBB
        Environment.OutSpatialReference = CType(IntersectRaster, ESRI.ArcGIS.Geodatabase.IGeoDataset).SpatialReference 'Grabs the spatial reference from the Intersect Raster JBB
        Dim ExtentProvider As Object = CType(CType(IntersectRaster, ESRI.ArcGIS.Geodatabase.IGeoDataset).Extent, ESRI.ArcGIS.Geometry.IEnvelope) 'From ESRI All Geometry objects have an envelope defined by the XMin, XMax, YMin, and YMax of the object JBB
        Environment.SetExtent(ESRI.ArcGIS.GeoAnalyst.esriRasterEnvSettingEnum.esriRasterEnvValue, ExtentProvider, IntersectRaster) 'Set the extent for further operations, intersect raster envelope appearst to be the extent JBB
        Environment.SetCellSize(ESRI.ArcGIS.GeoAnalyst.esriRasterEnvSettingEnum.esriRasterEnvMinOf) 'Cell size is minimum of input file cells JBB
        Dim Geo = CType(InputRaster, ESRI.ArcGIS.Geodatabase.IGeoDataset)
        Dim OutRaster As ESRI.ArcGIS.Geodatabase.IRaster = ExtractionOperation.Raster(CType(InputRaster, ESRI.ArcGIS.Geodatabase.IGeoDataset), CType(IntersectRaster, ESRI.ArcGIS.Geodatabase.IGeoDataset)) 'Extracts the data from the input raster to a new raster matching the intersect raster JBB
        Dim Save As ESRI.ArcGIS.Geodatabase.ISaveAs2 = OutRaster 'Saves the output extracted raster as a temp file JBB
        '*****TESTING
        'Dim FileName As String = IO.Path.ChangeExtension(IO.Path.GetTempFileName, ".img") 'Saves the output extracted raster as a temp file JBB
        '******END TESTING
        Save.SaveAs(FileName, OutWorkspace, "IMAGINE Image") 'Saves the output extracted raster as a temp file with .img extension JBB
        InputRaster = CreateRasterDataset(FileName) 'Creates a raster data set for the extracted raster JBB
    End Sub

    Public Function TransformRaster(ByVal Path As String, ByVal RasterDataset As ESRI.ArcGIS.Geodatabase.IRasterDataset, ByVal gcsType As Integer) As ESRI.ArcGIS.DataSourcesRaster.IRaster2
        Dim ExtractionOperation As ESRI.ArcGIS.SpatialAnalyst.IExtractionOp = New ESRI.ArcGIS.SpatialAnalyst.RasterExtractionOp
        Dim Environment As ESRI.ArcGIS.GeoAnalyst.IRasterAnalysisEnvironment = ExtractionOperation
        Dim WorkspaceFactory As ESRI.ArcGIS.Geodatabase.IWorkspaceFactory = New ESRI.ArcGIS.DataSourcesRaster.RasterWorkspaceFactory
        'Testing
        'Dim OutWorkspace As ESRI.ArcGIS.Geodatabase.IWorkspace = WorkspaceFactory.OpenFromFile(IO.Path.GetDirectoryName(Path), 0)
        Dim TransformFileName As String = IO.Path.ChangeExtension(IO.Path.GetTempFileName, ".img")
        Dim OutWorkspace As ESRI.ArcGIS.Geodatabase.IWorkspace = WorkspaceFactory.OpenFromFile(IO.Path.GetDirectoryName(TransformFileName), 0)
        'end testing
        Environment.OutWorkspace = OutWorkspace
        Dim SpatialReferenceFactory As ESRI.ArcGIS.Geometry.ISpatialReferenceFactory = Activator.CreateInstance(Type.GetTypeFromProgID("esriGeometry.SpatialReferenceEnvironment"))
        Dim GeographicCoordinateSystem As ESRI.ArcGIS.Geometry.IGeographicCoordinateSystem = SpatialReferenceFactory.CreateGeographicCoordinateSystem(gcsType)
        Environment.OutSpatialReference = GeographicCoordinateSystem
        Dim TransformedRaster = ExtractionOperation.Raster(CType(RasterDataset, ESRI.ArcGIS.Geodatabase.IGeoDataset), CType(RasterDataset, ESRI.ArcGIS.Geodatabase.IGeoDataset))
        Dim Save As ESRI.ArcGIS.Geodatabase.ISaveAs2 = TransformedRaster
        'testing
        'Dim TransformFileName As String = IO.Path.ChangeExtension(IO.Path.GetTempFileName, ".img")
        'end testing
        Save.SaveAs(TransformFileName, OutWorkspace, "IMAGINE Image")
        Dim TransformRaster2 As ESRI.ArcGIS.DataSourcesRaster.IRaster2 = CType(CreateRasterDataset(TransformFileName).CreateDefaultRaster, ESRI.ArcGIS.DataSourcesRaster.IRaster2)
        Return TransformRaster2
    End Function

    Public Function ExistsArcGISFile(ByVal Path As String)
        Try
            If IO.File.Exists(Path) Then
                Dim GPUtilities As ESRI.ArcGIS.Geoprocessing.IGPUtilities = New ESRI.ArcGIS.Geoprocessing.GPUtilities
                Dim DataElement = Nothing
                Try
                    DataElement = GPUtilities.MakeDataElement(Path, Nothing, Nothing)
                Catch ex As Exception
                    Throw
                End Try
                Dim GPValue = GPUtilities.MakeGPValueFromObject(DataElement)
                Return GPUtilities.Exists(GPValue)
            Else
                Return False
            End If
        Catch
            Return False
        End Try
        'Return False
    End Function

    Public Function ExistsArcGISFile(ByVal Paths As List(Of String))
        Dim Exists As Boolean = True
        For I = 0 To Paths.Count - 1
            If Not ExistsArcGISFile(Paths(I)) Then
                MsgBox(Paths(I) & " does not exist.")
                Exists = False
                Exit For
            End If
        Next
        Return Exists
    End Function

    Public Sub DeleteArcGISFile(ByVal Path As String)
        Dim GPUtilities As ESRI.ArcGIS.Geoprocessing.IGPUtilities = New ESRI.ArcGIS.Geoprocessing.GPUtilities
        Dim DataElement = GPUtilities.MakeDataElement(Path, Nothing, Nothing)
        Dim GPValue = GPUtilities.MakeGPValueFromObject(DataElement)
        GPUtilities.Delete(GPValue)
    End Sub

    Public Sub AddRastersIntoListBox(ListBox As Windows.Forms.ListBox, Text As String)
        Dim OpenFileDialog As ESRI.ArcGIS.CatalogUI.IGxDialog = New ESRI.ArcGIS.CatalogUI.GxDialog
        OpenFileDialog.Title = "Load " & Text & " images"
        OpenFileDialog.AllowMultiSelect = True
        Dim Filter As ESRI.ArcGIS.Catalog.IGxObjectFilter = New ESRI.ArcGIS.Catalog.GxFilterRasterDatasets
        OpenFileDialog.ObjectFilter = Filter
        Dim List As ESRI.ArcGIS.Catalog.IEnumGxObject = Nothing
        Dim GxObject As ESRI.ArcGIS.Catalog.IGxDataset

        If Not OpenFileDialog.DoModalOpen(Me.Handle, List) Then Exit Sub

        GxObject = List.Next
        Do
            Dim Path As String = GxObject.Dataset.Workspace.PathName & GxObject.Dataset.Name
            If Not ListBox.Items.Contains(Path) Then ListBox.Items.Add(Path)
            GxObject = List.Next
        Loop Until GxObject Is Nothing

        System.Runtime.InteropServices.Marshal.ReleaseComObject(OpenFileDialog)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(Filter)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(List)
    End Sub

    Public Sub RemoveRastersFromListBox(ListBox As Windows.Forms.ListBox)
        Dim Indices = ListBox.SelectedIndices
        For Index = Indices.Count - 1 To 0 Step -1
            ListBox.Items.RemoveAt(Indices(Index))
        Next
    End Sub

    Public Function FileExists(ListBox As Windows.Forms.ListBox, List As List(Of String), Index As Integer)
        Dim Exists = True

        For I = 0 To ListBox.Items.Count - 1
            If Not ExistsArcGISFile(List(I)) Then
                TabControl1.SelectedIndex = Index
                ListBox.ClearSelected()
                ListBox.SelectedIndex = ListBox.Items.IndexOf(List(I))
                MsgBox(List(I) & " does not exist.")
                Exists = False
                Exit For
            End If
        Next

        Return Exists
    End Function

    Public Function FileExists(TextBox As Windows.Forms.TextBox, Index As Integer)
        Dim Exists = True

        If Not TextBox.Text = "" Then
            If Not ExistsArcGISFile(TextBox.Text) Then
                TabControl1.SelectedIndex = Index
                TextBox.SelectAll()
                MsgBox(TextBox.Text & " does not exist.")
                Exists = False
            End If
        End If

        Return Exists
    End Function

#End Region

#Region "Project"

#Region "Load"

    Private Sub Main_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.AutoScaleMode = Windows.Forms.AutoScaleMode.Dpi
        '
        'CalculationCoordinatesGrid
        '
        CalculationCoordinatesGrid = New System.Windows.Forms.DataGridView()
        CType(CalculationCoordinatesGrid, System.ComponentModel.ISupportInitialize).BeginInit()
        TabPage10.Controls.Add(CalculationCoordinatesGrid)
        CalculationCoordinatesGrid.AllowUserToAddRows = False
        CalculationCoordinatesGrid.AllowUserToDeleteRows = False
        CalculationCoordinatesGrid.AllowUserToResizeRows = False
        CalculationCoordinatesGrid.BackgroundColor = System.Drawing.SystemColors.Window
        CalculationCoordinatesGrid.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        CalculationCoordinatesGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        CalculationCoordinatesGrid.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        'CalculationCoordinatesGrid.Location = New System.Drawing.Point(16, 85)
        CalculationCoordinatesGrid.Location = New System.Drawing.Point(16, 91) 'Modified by JBB
        CalculationCoordinatesGrid.Margin = New System.Windows.Forms.Padding(13, 3, 13, 13)
        CalculationCoordinatesGrid.Name = "CalculationCoordinatesGrid"
        CalculationCoordinatesGrid.RowHeadersVisible = False
        CalculationCoordinatesGrid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        'CalculationCoordinatesGrid.Size = New System.Drawing.Size(402, 107)
        CalculationCoordinatesGrid.Size = New System.Drawing.Size(741, 236) 'Modified by JBB
        CalculationCoordinatesGrid.TabIndex = 85
        CType(CalculationCoordinatesGrid, System.ComponentModel.ISupportInitialize).EndInit()

        NewProject_Click(Nothing, Nothing)

        'Brazil()
        'Cibola_air()
        'Cibola()
        'Sample_2()
        'AGU_2013_point()
        'AGU_2013_gridded()
        'Sample_Iowa_Landsat_hybrid()
        'Sample_Iowa_MODIS_TSEB()
        'Sample_Iowa_MODIS_hybrid()
    End Sub

    Public Class DataGridViewColumns

        Shared Sub AddNormal(ByVal DataGridView As Windows.Forms.DataGridView, ByVal ColumnName As String, Optional ByVal DataType As DataTypes = DataTypes.typeString, Optional ByVal Read_Only As Boolean = False)
            DataGridView.Columns.Add(ColumnName, ColumnName)
            Dim NewColumnDataType As String = ""
            Select Case DataType
                Case DataTypes.typeBoolean
                    NewColumnDataType = "Boolean"
                Case DataTypes.typeByte
                    NewColumnDataType = "Byte"
                Case DataTypes.typeChar
                    NewColumnDataType = "Char"
                Case DataTypes.typeDate
                    NewColumnDataType = "DateTime"
                Case DataTypes.typeDecimal
                    NewColumnDataType = "Decimal"
                Case DataTypes.typeDouble
                    NewColumnDataType = "Double"
                Case DataTypes.typeInteger
                    NewColumnDataType = "Int32"
                Case DataTypes.typeLong
                    NewColumnDataType = "Int64"
                Case DataTypes.typeObject
                    NewColumnDataType = "Object"
                Case DataTypes.typeSByte
                    NewColumnDataType = "SByte"
                Case DataTypes.typeShort
                    NewColumnDataType = "Int16"
                Case DataTypes.typeSingle
                    NewColumnDataType = "Single"
                Case DataTypes.typeString
                    NewColumnDataType = "String"
                Case DataTypes.typeUInteger
                    NewColumnDataType = "UInt32"
                Case DataTypes.typeULong
                    NewColumnDataType = "UInt64"
                Case DataTypes.typeUShort
                    NewColumnDataType = "UInt16"
            End Select
            DataGridView.Columns(ColumnName).ValueType = Type.GetType("System." & NewColumnDataType)
            DataGridView.Columns(ColumnName).ReadOnly = Read_Only
        End Sub

        Shared Sub AddCombo(ByVal DataGridView As Windows.Forms.DataGridView, ByVal ColumnName As String, ByVal Items As Array, Optional ByVal BindingList As ArrayList = Nothing)
            Dim NewColumn As New Windows.Forms.DataGridViewComboBoxColumn

            NewColumn.Name = ColumnName

            For Each Item In Items
                NewColumn.Items.Add(Item)
            Next

            NewColumn.FlatStyle = Windows.Forms.FlatStyle.Flat

            Dim Graphics As Drawing.Graphics = Drawing.Graphics.FromHwnd(New System.IntPtr)
            Dim Width As Integer = 9
            For Each Item In Items
                Width = Math.Max(Width, Graphics.MeasureString(Item, DataGridView.DefaultCellStyle.Font).ToSize.Width)
            Next
            NewColumn.DropDownWidth = Math.Max(Width + 15, NewColumn.DropDownWidth)

            If Not BindingList Is Nothing Then NewColumn.DataSource = BindingList

            DataGridView.Columns.Add(NewColumn)
        End Sub

        Shared Sub AddCheckbox(ByVal DataGridView As Windows.Forms.DataGridView, ByVal ColumnName As String)
            Dim NewColumn As New Windows.Forms.DataGridViewCheckBoxColumn
            NewColumn.Name = ColumnName
            DataGridView.Columns.Add(NewColumn)
        End Sub

        Enum DataTypes
            typeBoolean = 1
            typeByte = 2
            typeChar = 3
            typeDate = 4
            typeDecimal = 5
            typeDouble = 6
            typeInteger = 7
            typeLong = 8
            typeObject = 9
            typeSByte = 10
            typeShort = 11
            typeSingle = 12
            typeString = 13
            typeUInteger = 14
            typeULong = 15
            typeUShort = 16
        End Enum

    End Class

    Sub Cibola_air()
        Dim Path As String = "E:\Cibola\Data_for_paper_3\Data_Chapman_paper\airborne_version_1\"
        EnergyBalanceBox.SelectedIndex = 1
        ETExtrapolationBox.SelectedIndex = 1
        WaterBalanceBox.SelectedIndex = 0
        ImageSourceBox.SelectedIndex = 0
        MultispectralList.Items.Add(Path & "reflectance_05-17-2008.img")
        Dim RasterDataset = CreateRasterDataset(MultispectralList.Items(0))
        Dim RasterBandCollection As ESRI.ArcGIS.DataSourcesRaster.IRasterBandCollection = RasterDataset
        For Item = 1 To RasterBandCollection.Count
            RedIndex.Items.Add(Item)
            GreenIndex.Items.Add(Item)
            BlueIndex.Items.Add(Item)
            NIRIndex.Items.Add(Item)
            MIR1Index.Items.Add(Item)
            MIR2Index.Items.Add(Item)
        Next
        System.Runtime.InteropServices.Marshal.ReleaseComObject(RasterBandCollection)
        SurfaceTemperatureList.Items.Add(Path & "temp_05-17-2008.img")
        CoverClassificationList.Items.Add(Path & "landuse_05-18-2008.img")
        Dim Raster As ESRI.ArcGIS.DataSourcesRaster.IRaster2 = CType(CreateRasterDataset(CoverClassificationList.Items(0)).CreateDefaultRaster, ESRI.ArcGIS.DataSourcesRaster.IRaster2)
        For F = 0 To Raster.AttributeTable.Fields.FieldCount - 1
            CoverSelectionBox.Items.Add(Raster.AttributeTable.Fields.Field(F).AliasName)
            If Raster.AttributeTable.Fields.Field(F).AliasName.Contains("Name") Or Raster.AttributeTable.Fields.Field(F).AliasName.Contains("NAME") Or Raster.AttributeTable.Fields.Field(F).AliasName.Contains("name") Then CoverSelectionBox.Text = Raster.AttributeTable.Fields.Field(F).AliasName
        Next
        System.Runtime.InteropServices.Marshal.ReleaseComObject(Raster)
        NIRIndex.SelectedIndex = 0
        RedIndex.SelectedIndex = 1
        MIR1Index.SelectedIndex = 2
        WeatherTableText.Text = Path & "airborne_Cibola_data_EB.xlsx\Sheet1$"
        CoverPropertiesText.Text = Path & "airborne_Cibola_data_EB.xlsx\Sheet4$"
        OutputDirectoryTextEnergy.Text = "E:\Cibola\Data_for_paper_3\Data_Chapman_paper\results_airborne_1"
        OutputDirectoryTextWater.Text = "E:\Cibola\Data_for_paper_3\Data_Chapman_paper\results_airborne_1"
        OutputImagesCheckAllEnergy_Click(New Object, New System.EventArgs)
        OutputImagesCheckAllWater_Click(New Object, New System.EventArgs)
        Dim Rows = {2, 3, 4}
        For Row = 0 To 3
            If Not Rows.Contains(Row) Then
                CoverSelectionGrid.Rows(Row).Cells(3).Value = False
            Else
                CoverSelectionGrid.Rows(Row).Cells(2).Value = "Basin"
            End If
        Next
        TimeZoneBox.SelectedIndex = TimeZoneBox.Items.IndexOf("Mountain Standard Time")
        Dim Outs = {3, 4, 6, 9, 11}
        For Item = 0 To OutputImagesBoxEnergy.Items.Count - 1
            If Not Outs.Contains(Item) Then OutputImagesBoxEnergy.SetItemChecked(Item, False)
        Next
        FieldCapacityText.Text = "D:\SETMI\USA Field Capacity.img"
        WiltingPointText.Text = "D:\SETMI\USA Wilting Point.img"
        WeatherTableGrid.Rows(6).Cells(1).Value = WeatherVariables(7)
    End Sub

    Sub Cibola()
        Dim Path As String = "E:\Cibola\Data_for_paper_3\Data_Chapman_paper\version_1_data\"
        EnergyBalanceBox.SelectedIndex = 1
        ETExtrapolationBox.SelectedIndex = 1
        WaterBalanceBox.SelectedIndex = 0
        ImageSourceBox.SelectedIndex = 1
        MultispectralList.Items.Add(Path & "reflectance_v1_05-17-2008 10-04.img")
        Dim RasterDataset = CreateRasterDataset(MultispectralList.Items(0))
        Dim RasterBandCollection As ESRI.ArcGIS.DataSourcesRaster.IRasterBandCollection = RasterDataset
        For Item = 1 To RasterBandCollection.Count
            RedIndex.Items.Add(Item)
            GreenIndex.Items.Add(Item)
            BlueIndex.Items.Add(Item)
            NIRIndex.Items.Add(Item)
            MIR1Index.Items.Add(Item)
            MIR2Index.Items.Add(Item)
        Next
        System.Runtime.InteropServices.Marshal.ReleaseComObject(RasterBandCollection)
        SurfaceTemperatureList.Items.Add(Path & "temp_05-17-2008 10-04.img")
        CoverClassificationList.Items.Add(Path & "Landuse_05-18-2008 10-00.img")
        Dim Raster As ESRI.ArcGIS.DataSourcesRaster.IRaster2 = CType(CreateRasterDataset(CoverClassificationList.Items(0)).CreateDefaultRaster, ESRI.ArcGIS.DataSourcesRaster.IRaster2)
        For F = 0 To Raster.AttributeTable.Fields.FieldCount - 1
            CoverSelectionBox.Items.Add(Raster.AttributeTable.Fields.Field(F).AliasName)
            If Raster.AttributeTable.Fields.Field(F).AliasName.Contains("Name") Or Raster.AttributeTable.Fields.Field(F).AliasName.Contains("NAME") Or Raster.AttributeTable.Fields.Field(F).AliasName.Contains("name") Then CoverSelectionBox.Text = Raster.AttributeTable.Fields.Field(F).AliasName
        Next
        System.Runtime.InteropServices.Marshal.ReleaseComObject(Raster)
        NIRIndex.SelectedIndex = 0
        RedIndex.SelectedIndex = 1
        MIR1Index.SelectedIndex = 2
        WeatherTableText.Text = Path & "Cibola_data_EB.xlsx\Sheet1$"
        CoverPropertiesText.Text = Path & "Cibola_data_EB.xlsx\Sheet4$"
        OutputDirectoryTextEnergy.Text = "E:\Cibola\Data_for_paper_3\Data_Chapman_paper\results_version_1_2"
        'OutputDirectoryTextWater.Text = "E:\Cibola\Data_for_paper_3\Data_Chapman_paper\results_version_1"
        OutputImagesCheckAllEnergy_Click(New Object, New System.EventArgs)
        OutputImagesCheckAllWater_Click(New Object, New System.EventArgs)
        Dim Rows = {2, 3, 4}
        For Row = 0 To 3
            If Not Rows.Contains(Row) Then
                CoverSelectionGrid.Rows(Row).Cells(3).Value = False
            Else
                CoverSelectionGrid.Rows(Row).Cells(2).Value = "Basin"
            End If
        Next
        TimeZoneBox.SelectedIndex = TimeZoneBox.Items.IndexOf("Mountain Standard Time")
        Dim Outs = {3, 4, 6, 9, 11}
        For Item = 0 To OutputImagesBoxEnergy.Items.Count - 1
            If Not Outs.Contains(Item) Then OutputImagesBoxEnergy.SetItemChecked(Item, False)
        Next
        'FieldCapacityText.Text = "D:\SETMI\USA Field Capacity.img"
        'WiltingPointText.Text = "D:\SETMI\USA Wilting Point.img"
        'WeatherTableGrid.Rows(6).Cells(1).Value = WeatherVariables(7)
    End Sub

    Sub Brazil()
        Dim Path As String = "E:\Brazil Study\SETMI_analysis\images_included\"
        EnergyBalanceBox.SelectedIndex = 1
        ETExtrapolationBox.SelectedIndex = 1
        WaterBalanceBox.SelectedIndex = 0
        ImageSourceBox.SelectedIndex = 1
        MultispectralList.Items.Add(Path & "reflectance_173_06-22-2010.img")
        Dim RasterDataset = CreateRasterDataset(MultispectralList.Items(0))
        Dim RasterBandCollection As ESRI.ArcGIS.DataSourcesRaster.IRasterBandCollection = RasterDataset
        For Item = 1 To RasterBandCollection.Count
            RedIndex.Items.Add(Item)
            GreenIndex.Items.Add(Item)
            BlueIndex.Items.Add(Item)
            NIRIndex.Items.Add(Item)
            MIR1Index.Items.Add(Item)
            MIR2Index.Items.Add(Item)
        Next
        System.Runtime.InteropServices.Marshal.ReleaseComObject(RasterBandCollection)
        SurfaceTemperatureList.Items.Add(Path & "temp_173_06-22-2010.img")
        CoverClassificationList.Items.Add("E:\Brazil Study\landuse\Landuse_new_06-22-2010.img")
        Dim Raster As ESRI.ArcGIS.DataSourcesRaster.IRaster2 = CType(CType(CreateRasterDataset(CoverClassificationList.Items(0)), ESRI.ArcGIS.Geodatabase.IRasterDataset2).CreateFullRaster, ESRI.ArcGIS.DataSourcesRaster.IRaster2)
        For F = 0 To Raster.AttributeTable.Fields.FieldCount - 1
            CoverSelectionBox.Items.Add(Raster.AttributeTable.Fields.Field(F).AliasName)
            If Raster.AttributeTable.Fields.Field(F).AliasName.Contains("Name") Or Raster.AttributeTable.Fields.Field(F).AliasName.Contains("NAME") Or Raster.AttributeTable.Fields.Field(F).AliasName.Contains("name") Then CoverSelectionBox.Text = Raster.AttributeTable.Fields.Field(F).AliasName
        Next
        System.Runtime.InteropServices.Marshal.ReleaseComObject(Raster)
        NIRIndex.SelectedIndex = 3
        RedIndex.SelectedIndex = 2
        MIR1Index.SelectedIndex = 4
        WeatherTableText.Text = "E:\Brazil Study\weather_data\Acarau_data_EB.xls\Sheet1$"
        OutputDirectoryTextEnergy.Text = "E:\Brazil Study\SETMI_analysis\results"
        OutputImagesCheckAllEnergy_Click(New Object, New System.EventArgs)
        Dim Rows = {0, 1, 2, 3, 4, 5, 6, 7, 8, 10, 11, 13, 14, 15, 18, 19, 21, 22, 23, 28}
        For Row = 0 To CoverSelectionGrid.Rows.Count - 1
            If Not Rows.Contains(Row) Then
                CoverSelectionGrid.Rows(Row).Cells(3).Value = False
            Else
                CoverSelectionGrid.Rows(Row).Cells(2).Value = "Basin"
            End If
        Next
        TimeZoneBox.SelectedIndex = TimeZoneBox.Items.IndexOf("Central Brazilian Standard Time")
        Dim Outs = {3, 4, 5, 6, 7, 9, 11}
        For Item = 0 To OutputImagesBoxEnergy.Items.Count - 1
            If Not Outs.Contains(Item) Then OutputImagesBoxEnergy.SetItemChecked(Item, False)
        Next
        FieldCapacityText.Text = "D:\SETMI\USA Field Capacity.img"
        WiltingPointText.Text = "D:\SETMI\USA Wilting Point.img"
        WeatherTableGrid.Rows(6).Cells(1).Value = WeatherVariables(7)
    End Sub

    Sub Sample()
        Dim Path As String = "D:\SETMI\SETMI_old\"
        EnergyBalanceBox.SelectedIndex = 1
        ETExtrapolationBox.SelectedIndex = 0
        WaterBalanceBox.SelectedIndex = 0
        ImageSourceBox.SelectedIndex = 0
        MultispectralList.Items.Add(Path & "3band_07-12-2008.img")
        Dim RasterDataset = CreateRasterDataset(MultispectralList.Items(0))
        Dim RasterBandCollection As ESRI.ArcGIS.DataSourcesRaster.IRasterBandCollection = RasterDataset
        For Item = 1 To RasterBandCollection.Count
            RedIndex.Items.Add(Item)
            GreenIndex.Items.Add(Item)
            BlueIndex.Items.Add(Item)
            NIRIndex.Items.Add(Item)
            MIR1Index.Items.Add(Item)
            MIR2Index.Items.Add(Item)
        Next
        System.Runtime.InteropServices.Marshal.ReleaseComObject(RasterBandCollection)
        SurfaceTemperatureList.Items.Add(Path & "temp_k_07-12-2008.img")
        CoverClassificationList.Items.Add(Path & "bearax08_class_filt_2_resamp_07-12-2008.img")
        Dim Raster As ESRI.ArcGIS.DataSourcesRaster.IRaster2 = CType(CreateRasterDataset(CoverClassificationList.Items(0)).CreateDefaultRaster, ESRI.ArcGIS.DataSourcesRaster.IRaster2)
        For F = 0 To Raster.AttributeTable.Fields.FieldCount - 1
            CoverSelectionBox.Items.Add(Raster.AttributeTable.Fields.Field(F).AliasName)
            If Raster.AttributeTable.Fields.Field(F).AliasName.Contains("Name") Or Raster.AttributeTable.Fields.Field(F).AliasName.Contains("NAME") Or Raster.AttributeTable.Fields.Field(F).AliasName.Contains("name") Then CoverSelectionBox.Text = Raster.AttributeTable.Fields.Field(F).AliasName
        Next
        System.Runtime.InteropServices.Marshal.ReleaseComObject(Raster)
        NIRIndex.SelectedIndex = 0
        RedIndex.SelectedIndex = 1
        GreenIndex.SelectedIndex = 2
        WeatherTableText.Text = Path & "Bearax_data_EB.xlsx\Sheet1$"
        CoverPropertiesText.Text = Path & "Bearax_data_EB.xlsx\Sheet2$"
        OutputDirectoryTextEnergy.Text = "D:\SETMI\Output"
        OutputDirectoryTextWater.Text = "D:\SETMI\Output"
        OutputImagesCheckAllEnergy_Click(New Object, New System.EventArgs)
        OutputImagesCheckAllWater_Click(New Object, New System.EventArgs)
        Dim Rows = {0, 1, 2, 3, 4, 5, 8, 9, 10, 11, 12}
        For Row = 0 To 12
            If Rows.Contains(Row) Then
                CoverSelectionGrid.Rows(Row).Cells(3).Value = False
            Else
                CoverSelectionGrid.Rows(Row).Cells(2).Value = "Not Irrigated"
            End If
        Next
        CoverSelectionGrid.Rows(6).Cells(2).Value = "Sprinkler"
        TimeZoneBox.SelectedIndex = TimeZoneBox.Items.IndexOf("Central Standard Time")
        For Item = 0 To OutputImagesBoxEnergy.Items.Count - 1
            If Not Item = 5 Then OutputImagesBoxEnergy.SetItemChecked(Item, False)
        Next
        FieldCapacityText.Text = "D:\SETMI\USA Field Capacity.img"
        WiltingPointText.Text = "D:\SETMI\USA Wilting Point.img"
    End Sub

    Sub Sample_2()
        Dim Path As String = "D:\SETMI\SETMI_old\"
        EnergyBalanceBox.SelectedIndex = 1
        ETExtrapolationBox.SelectedIndex = 0
        WaterBalanceBox.SelectedIndex = 0
        ImageSourceBox.SelectedIndex = 1
        MultispectralList.Items.Add("D:\AGU_2012_PVID\Landsat_images\PVID_subset\3banb_pvid_131_2008_tucson_sub_05-10-2008.img")
        Dim RasterDataset = CreateRasterDataset(MultispectralList.Items(0))
        Dim RasterBandCollection As ESRI.ArcGIS.DataSourcesRaster.IRasterBandCollection = RasterDataset
        For Item = 1 To RasterBandCollection.Count
            RedIndex.Items.Add(Item)
            GreenIndex.Items.Add(Item)
            BlueIndex.Items.Add(Item)
            NIRIndex.Items.Add(Item)
            MIR1Index.Items.Add(Item)
            MIR2Index.Items.Add(Item)
        Next
        System.Runtime.InteropServices.Marshal.ReleaseComObject(RasterBandCollection)
        SurfaceTemperatureList.Items.Add("D:\AGU_2012_PVID\Landsat_images\PVID_subset\pvid_thermal_c_131_2008_corrected_tucson_sub_05-10-2008.img")
        CoverClassificationList.Items.Add("D:\AGU_2012_PVID\landuse\Modified_2\n1_05-01-2008.img")
        Dim Raster As ESRI.ArcGIS.DataSourcesRaster.IRaster2 = CType(CreateRasterDataset(CoverClassificationList.Items(0)).CreateDefaultRaster, ESRI.ArcGIS.DataSourcesRaster.IRaster2)
        For F = 0 To Raster.AttributeTable.Fields.FieldCount - 1
            CoverSelectionBox.Items.Add(Raster.AttributeTable.Fields.Field(F).AliasName)
            If Raster.AttributeTable.Fields.Field(F).AliasName.Contains("Name") Or Raster.AttributeTable.Fields.Field(F).AliasName.Contains("NAME") Or Raster.AttributeTable.Fields.Field(F).AliasName.Contains("name") Then CoverSelectionBox.Text = Raster.AttributeTable.Fields.Field(F).AliasName
        Next
        System.Runtime.InteropServices.Marshal.ReleaseComObject(Raster)
        NIRIndex.SelectedIndex = 0
        RedIndex.SelectedIndex = 1
        GreenIndex.SelectedIndex = 2
        WeatherTableText.Text = "D:\AGU_2012_PVID\weather_data\PVID_data_EB.xlsx\Sheet1$"
        CoverPropertiesText.Text = "D:\AGU_2012_PVID\weather_data\Bearax_data_EB.xlsx\Sheet2$"
        OutputDirectoryTextEnergy.Text = "D:\AGU_2012_PVID\TSEB_results"
        OutputDirectoryTextWater.Text = "D:\SETMI\Output"
        OutputImagesCheckAllEnergy_Click(New Object, New System.EventArgs)
        OutputImagesCheckAllWater_Click(New Object, New System.EventArgs)
        Dim Rows = {0, 1, 2, 3, 6, 7, 8, 15, 16, 18}
        For Row = 0 To CoverSelectionGrid.Rows.Count - 1
            If Not Rows.Contains(Row) Then
                CoverSelectionGrid.Rows(Row).Cells(3).Value = False
            Else
                CoverSelectionGrid.Rows(Row).Cells(2).Value = "Basin"
            End If
        Next
        CoverSelectionGrid.Rows(2).Cells(1).Value = "Wheat"
        CoverSelectionGrid.Rows(3).Cells(1).Value = "Wheat"
        CoverSelectionGrid.Rows(6).Cells(1).Value = "Grass"
        CoverSelectionGrid.Rows(7).Cells(1).Value = "Agriculture"
        CoverSelectionGrid.Rows(8).Cells(1).Value = "Bare Soil"
        CoverSelectionGrid.Rows(15).Cells(1).Value = "Grass"
        CoverSelectionGrid.Rows(18).Cells(1).Value = "Agriculture"
        TimeZoneBox.SelectedIndex = TimeZoneBox.Items.IndexOf("Pacific Standard Time")
        Dim Output = {3, 5, 7, 10}
        For Item = 0 To OutputImagesBoxEnergy.Items.Count - 1
            If Not Output.Contains(Item) Then OutputImagesBoxEnergy.SetItemChecked(Item, False)
        Next
        FieldCapacityText.Text = "D:\SETMI\USA Field Capacity.img"
        WiltingPointText.Text = "D:\SETMI\USA Wilting Point.img"
        WeatherTemperatureList.Items.Add("D:\AGU_2012_PVID\gridded_weather\Temperature1_05_10_2008.img")
        WeatherHumidityList.Items.Add("D:\AGU_2012_PVID\gridded_weather\Humidity1_05_10_2008.img")
        WeatherWindSpeedList.Items.Add("D:\AGU_2012_PVID\gridded_weather\Wind Speed1_05_10_2008.img")
        WeatherRadiationList.Items.Add("D:\AGU_2012_PVID\gridded_weather\Solar Radiation1_05_10_2008.img")
    End Sub

    Sub AGU_2013_gridded()
        'Dim Path As String = "D:\SETMI\SETMI_old\"
        EnergyBalanceBox.SelectedIndex = 1
        ETExtrapolationBox.SelectedIndex = 1
        'WaterBalanceBox.SelectedIndex = 0
        ImageSourceBox.SelectedIndex = 1
        MultispectralList.Items.Add("E:\AGU_2013_PVID\Landsat_images\PVID_subset\3banb_pvid_019_2008_tucson_sub_01-19-2008 09-58.img")
        MultispectralList.Items.Add("E:\AGU_2013_PVID\Landsat_images\PVID_subset\3banb_pvid_042_2008_tucson_sub_02-11-2008 09-58.img")
        MultispectralList.Items.Add("E:\AGU_2013_PVID\Landsat_images\PVID_subset\3banb_pvid_058_2008_tucson_sub_02-27-2008 09-58.img")
        MultispectralList.Items.Add("E:\AGU_2013_PVID\Landsat_images\PVID_subset\3banb_pvid_067_2008_tucson_sub_03-07-2008 09-58.img")
        MultispectralList.Items.Add("E:\AGU_2013_PVID\Landsat_images\PVID_subset\3banb_pvid_083_2008_tucson_sub_03-23-2008 09-58.img")
        MultispectralList.Items.Add("E:\AGU_2013_PVID\Landsat_images\PVID_subset\3banb_pvid_099_2008_tucson_sub_04-08-2008 09-58.img")
        MultispectralList.Items.Add("E:\AGU_2013_PVID\Landsat_images\PVID_subset\3banb_pvid_115_2008_tucson_sub_04-24-2008 09-58.img")
        MultispectralList.Items.Add("E:\AGU_2013_PVID\Landsat_images\PVID_subset\3banb_pvid_131_2008_tucson_sub_05-10-2008 09-58.img")
        MultispectralList.Items.Add("E:\AGU_2013_PVID\Landsat_images\PVID_subset\3banb_pvid_138_2008_tucson_sub_05-17-2008 10-04.img")
        MultispectralList.Items.Add("E:\AGU_2013_PVID\Landsat_images\PVID_subset\3banb_pvid_147_2008_tucson_sub_05-26-2008 09-58.img")
        MultispectralList.Items.Add("E:\AGU_2013_PVID\Landsat_images\PVID_subset\3banb_pvid_163_2008_tucson_sub_06-11-2008 09-58.img")
        MultispectralList.Items.Add("E:\AGU_2013_PVID\Landsat_images\PVID_subset\3banb_pvid_170_2008_tucson_sub_06-18-2008 10-03.img")
        MultispectralList.Items.Add("E:\AGU_2013_PVID\Landsat_images\PVID_subset\3banb_pvid_195_2008_tucson_sub_07-13-2008 09-58.img")
        MultispectralList.Items.Add("E:\AGU_2013_PVID\Landsat_images\PVID_subset\3banb_pvid_211_2008_tucson_sub_07-29-2008 09-57.img")
        MultispectralList.Items.Add("E:\AGU_2013_PVID\Landsat_images\PVID_subset\3banb_pvid_218_2008_tucson_sub_08-05-2008 09-58.img")
        MultispectralList.Items.Add("E:\AGU_2013_PVID\Landsat_images\PVID_subset\3banb_pvid_234_2008_tucson_sub_08-21-2008 09-58.img")
        MultispectralList.Items.Add("E:\AGU_2013_PVID\Landsat_images\PVID_subset\3banb_pvid_259_2008_tucson_sub_09-15-2008 09-58.img")
        MultispectralList.Items.Add("E:\AGU_2013_PVID\Landsat_images\PVID_subset\3banb_pvid_275_2008_tucson_sub_10-01-2008 09-58.img")
        MultispectralList.Items.Add("E:\AGU_2013_PVID\Landsat_images\PVID_subset\3banb_pvid_291_2008_tucson_sub_10-17-2008 09-58.img")
        MultispectralList.Items.Add("E:\AGU_2013_PVID\Landsat_images\PVID_subset\3banb_pvid_314_2008_tucson_sub_11-09-2008 09-58.img")
        MultispectralList.Items.Add("E:\AGU_2013_PVID\Landsat_images\PVID_subset\3banb_pvid_323_2008_tucson_sub_11-18-2008 09-58.img")
        Dim RasterDataset = CreateRasterDataset(MultispectralList.Items(0))
        Dim RasterBandCollection As ESRI.ArcGIS.DataSourcesRaster.IRasterBandCollection = RasterDataset
        For Item = 1 To RasterBandCollection.Count
            RedIndex.Items.Add(Item)
            GreenIndex.Items.Add(Item)
            BlueIndex.Items.Add(Item)
            NIRIndex.Items.Add(Item)
            MIR1Index.Items.Add(Item)
            MIR2Index.Items.Add(Item)
        Next
        System.Runtime.InteropServices.Marshal.ReleaseComObject(RasterBandCollection)
        SurfaceTemperatureList.Items.Add("E:\AGU_2013_PVID\Landsat_images\PVID_subset\pvid_thermal_c_019_2008_corrected_tucson_sub_01-19-2008 09-58.img")
        SurfaceTemperatureList.Items.Add("E:\AGU_2013_PVID\Landsat_images\PVID_subset\pvid_thermal_c_042_2008_corrected_tucson_sub_02-11-2008 09-58.img")
        SurfaceTemperatureList.Items.Add("E:\AGU_2013_PVID\Landsat_images\PVID_subset\pvid_thermal_c_058_2008_corrected_tucson_sub_02-27-2008 09-58.img")
        SurfaceTemperatureList.Items.Add("E:\AGU_2013_PVID\Landsat_images\PVID_subset\pvid_thermal_c_067_2008_corrected_tucson_sub_03-07-2008 09-58.img")
        SurfaceTemperatureList.Items.Add("E:\AGU_2013_PVID\Landsat_images\PVID_subset\pvid_thermal_c_083_2008_corrected_tucson_sub_03-23-2008 09-58.img")
        SurfaceTemperatureList.Items.Add("E:\AGU_2013_PVID\Landsat_images\PVID_subset\pvid_thermal_c_099_2008_corrected_tucson_sub_04-08-2008 09-58.img")
        SurfaceTemperatureList.Items.Add("E:\AGU_2013_PVID\Landsat_images\PVID_subset\pvid_thermal_c_115_2008_corrected_tucson_sub_04-24-2008 09-58.img")
        SurfaceTemperatureList.Items.Add("E:\AGU_2013_PVID\Landsat_images\PVID_subset\pvid_thermal_c_131_2008_corrected_tucson_sub_05-10-2008 09-58.img")
        SurfaceTemperatureList.Items.Add("E:\AGU_2013_PVID\Landsat_images\PVID_subset\pvid_thermal_c_138_2008_corrected_tucson_sub_05-17-2008 10-04.img")
        SurfaceTemperatureList.Items.Add("E:\AGU_2013_PVID\Landsat_images\PVID_subset\pvid_thermal_c_147_2008_corrected_tucson_sub_05-26-2008 09-58.img")
        SurfaceTemperatureList.Items.Add("E:\AGU_2013_PVID\Landsat_images\PVID_subset\pvid_thermal_c_163_2008_corrected_tucson_sub_06-11-2008 09-58.img")
        SurfaceTemperatureList.Items.Add("E:\AGU_2013_PVID\Landsat_images\PVID_subset\pvid_thermal_c_170_2008_corrected_tucson_sub_06-18-2008 10-03.img")
        SurfaceTemperatureList.Items.Add("E:\AGU_2013_PVID\Landsat_images\PVID_subset\pvid_thermal_c_195_2008_corrected_tucson_sub_07-13-2008 09-58.img")
        SurfaceTemperatureList.Items.Add("E:\AGU_2013_PVID\Landsat_images\PVID_subset\pvid_thermal_c_211_2008_corrected_tucson_sub_07-29-2008 09-57.img")
        SurfaceTemperatureList.Items.Add("E:\AGU_2013_PVID\Landsat_images\PVID_subset\pvid_thermal_c_218_2008_corrected_tucson_sub_08-05-2008 09-58.img")
        SurfaceTemperatureList.Items.Add("E:\AGU_2013_PVID\Landsat_images\PVID_subset\pvid_thermal_c_234_2008_corrected_tucson_sub_08-21-2008 09-58.img")
        SurfaceTemperatureList.Items.Add("E:\AGU_2013_PVID\Landsat_images\PVID_subset\pvid_thermal_c_259_2008_corrected_tucson_sub_09-15-2008 09-58.img")
        SurfaceTemperatureList.Items.Add("E:\AGU_2013_PVID\Landsat_images\PVID_subset\pvid_thermal_c_275_2008_corrected_tucson_sub_10-01-2008 09-58.img")
        SurfaceTemperatureList.Items.Add("E:\AGU_2013_PVID\Landsat_images\PVID_subset\pvid_thermal_c_291_2008_corrected_tucson_sub_10-17-2008 09-58.img")
        SurfaceTemperatureList.Items.Add("E:\AGU_2013_PVID\Landsat_images\PVID_subset\pvid_thermal_c_314_2008_corrected_tucson_sub_11-09-2008 09-58.img")
        SurfaceTemperatureList.Items.Add("E:\AGU_2013_PVID\Landsat_images\PVID_subset\pvid_thermal_c_323_2008_corrected_tucson_sub_11-18-2008 09-58.img")
        CoverClassificationList.Items.Add("E:\AGU_2013_PVID\landuse\Modified_2\Landuse_05-01-2008 10-00.img")
        Dim Raster As ESRI.ArcGIS.DataSourcesRaster.IRaster2 = CType(CreateRasterDataset(CoverClassificationList.Items(0)).CreateDefaultRaster, ESRI.ArcGIS.DataSourcesRaster.IRaster2)
        For F = 0 To Raster.AttributeTable.Fields.FieldCount - 1
            CoverSelectionBox.Items.Add(Raster.AttributeTable.Fields.Field(F).AliasName)
            If Raster.AttributeTable.Fields.Field(F).AliasName.Contains("Name") Or Raster.AttributeTable.Fields.Field(F).AliasName.Contains("NAME") Or Raster.AttributeTable.Fields.Field(F).AliasName.Contains("name") Then CoverSelectionBox.Text = Raster.AttributeTable.Fields.Field(F).AliasName
        Next
        System.Runtime.InteropServices.Marshal.ReleaseComObject(Raster)
        NIRIndex.SelectedIndex = 0
        RedIndex.SelectedIndex = 1
        GreenIndex.SelectedIndex = 2
        WeatherTableText.Text = "E:\AGU_2013_PVID\weather_data\PVID_data_EB.xlsx\gridded$"
        'CoverPropertiesText.Text = "E:\AGU_2013_PVID\weather_data\Bearax_data_EB.xlsx\Sheet2$"
        OutputDirectoryTextEnergy.Text = "E:\AGU_2013_PVID\testing_new_results_gridded_2014"
        OutputDirectoryTextWater.Text = "E:\SETMI\Output"
        OutputImagesCheckAllEnergy_Click(New Object, New System.EventArgs)
        OutputImagesCheckAllWater_Click(New Object, New System.EventArgs)
        Dim Rows = {0, 1, 2, 3, 4, 5, 6, 7, 9, 12, 13, 14}
        For Row = 0 To CoverSelectionGrid.Rows.Count - 1
            If Not Rows.Contains(Row) Then
                CoverSelectionGrid.Rows(Row).Cells(3).Value = False
            Else
                CoverSelectionGrid.Rows(Row).Cells(2).Value = "Basin"
            End If
        Next
        CoverSelectionGrid.Rows(2).Cells(1).Value = "Wheat"
        CoverSelectionGrid.Rows(3).Cells(1).Value = "Wheat"
        CoverSelectionGrid.Rows(4).Cells(1).Value = "Watermelon"
        CoverSelectionGrid.Rows(5).Cells(1).Value = "Grass"
        CoverSelectionGrid.Rows(6).Cells(1).Value = "Agriculture"
        CoverSelectionGrid.Rows(7).Cells(1).Value = "Bare Soil"
        CoverSelectionGrid.Rows(9).Cells(1).Value = "Agriculture"
        CoverSelectionGrid.Rows(12).Cells(1).Value = "Grass"
        CoverSelectionGrid.Rows(13).Cells(1).Value = "Grass"
        CoverSelectionGrid.Rows(14).Cells(1).Value = "Citrus"
        TimeZoneBox.SelectedIndex = TimeZoneBox.Items.IndexOf("Pacific Standard Time")
        Dim Output = {3, 4, 6, 9, 11}
        For Item = 0 To OutputImagesBoxEnergy.Items.Count - 1
            If Not Output.Contains(Item) Then OutputImagesBoxEnergy.SetItemChecked(Item, False)
        Next
        'FieldCapacityText.Text = "D:\SETMI\USA Field Capacity.img"
        'WiltingPointText.Text = "D:\SETMI\USA Wilting Point.img"
        WeatherTemperatureList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\tmp_019_18_01-19-2008 09-58.img")
        WeatherTemperatureList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\tmp_042_18_02-11-2008 09-58.img")
        WeatherTemperatureList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\tmp_058_18_02-27-2008 09-58.img")
        WeatherTemperatureList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\tmp_067_18_03-07-2008 09-58.img")
        WeatherTemperatureList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\tmp_083_18_03-23-2008 09-58.img")
        WeatherTemperatureList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\tmp_099_18_04-08-2008 09-58.img")
        WeatherTemperatureList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\tmp_115_18_04-24-2008 09-58.img")
        WeatherTemperatureList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\tmp_131_18_05-10-2008 09-58.img")
        WeatherTemperatureList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\tmp_138_18_05-17-2008 10-04.img")
        WeatherTemperatureList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\tmp_147_18_05-26-2008 09-58.img")
        WeatherTemperatureList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\tmp_163_18_06-11-2008 09-58.img")
        WeatherTemperatureList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\tmp_170_18_06-18-2008 10-03.img")
        WeatherTemperatureList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\tmp_195_18_07-13-2008 09-58.img")
        WeatherTemperatureList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\tmp_211_18_07-29-2008 09-57.img")
        WeatherTemperatureList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\tmp_218_18_08-05-2008 09-58.img")
        WeatherTemperatureList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\tmp_234_18_08-21-2008 09-58.img")
        WeatherTemperatureList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\tmp_259_18_09-15-2008 09-58.img")
        WeatherTemperatureList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\tmp_275_18_10-01-2008 09-58.img")
        WeatherTemperatureList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\tmp_291_18_10-17-2008 09-58.img")
        WeatherTemperatureList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\tmp_314_18_11-09-2008 09-58.img")
        WeatherTemperatureList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\tmp_323_18_11-18-2008 09-58.img")

        WeatherHumidityList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\spfh_019_18_01-19-2008 09-58.img")
        WeatherHumidityList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\spfh_042_18_02-11-2008 09-58.img")
        WeatherHumidityList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\spfh_058_18_02-27-2008 09-58.img")
        WeatherHumidityList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\spfh_067_18_03-07-2008 09-58.img")
        WeatherHumidityList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\spfh_083_18_03-23-2008 09-58.img")
        WeatherHumidityList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\spfh_099_18_04-08-2008 09-58.img")
        WeatherHumidityList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\spfh_115_18_04-24-2008 09-58.img")
        WeatherHumidityList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\spfh_131_18_05-10-2008 09-58.img")
        WeatherHumidityList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\spfh_138_18_05-17-2008 10-04.img")
        WeatherHumidityList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\spfh_147_18_05-26-2008 09-58.img")
        WeatherHumidityList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\spfh_163_18_06-11-2008 09-58.img")
        WeatherHumidityList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\spfh_170_18_06-18-2008 10-03.img")
        WeatherHumidityList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\spfh_195_18_07-13-2008 09-58.img")
        WeatherHumidityList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\spfh_211_18_07-29-2008 09-57.img")
        WeatherHumidityList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\spfh_218_18_08-05-2008 09-58.img")
        WeatherHumidityList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\spfh_234_18_08-21-2008 09-58.img")
        WeatherHumidityList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\spfh_259_18_09-15-2008 09-58.img")
        WeatherHumidityList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\spfh_275_18_10-01-2008 09-58.img")
        WeatherHumidityList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\spfh_291_18_10-17-2008 09-58.img")
        WeatherHumidityList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\spfh_314_18_11-09-2008 09-58.img")
        WeatherHumidityList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\spfh_323_18_11-18-2008 09-58.img")

        WeatherWindSpeedList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\wind_019_18_01-19-2008 09-58.img")
        WeatherWindSpeedList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\wind_042_18_02-11-2008 09-58.img")
        WeatherWindSpeedList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\wind_058_18_02-27-2008 09-58.img")
        WeatherWindSpeedList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\wind_067_18_03-07-2008 09-58.img")
        WeatherWindSpeedList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\wind_083_18_03-23-2008 09-58.img")
        WeatherWindSpeedList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\wind_099_18_04-08-2008 09-58.img")
        WeatherWindSpeedList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\wind_115_18_04-24-2008 09-58.img")
        WeatherWindSpeedList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\wind_131_18_05-10-2008 09-58.img")
        WeatherWindSpeedList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\wind_138_18_05-17-2008 10-04.img")
        WeatherWindSpeedList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\wind_147_18_05-26-2008 09-58.img")
        WeatherWindSpeedList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\wind_163_18_06-11-2008 09-58.img")
        WeatherWindSpeedList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\wind_170_18_06-18-2008 10-03.img")
        WeatherWindSpeedList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\wind_195_18_07-13-2008 09-58.img")
        WeatherWindSpeedList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\wind_211_18_07-29-2008 09-57.img")
        WeatherWindSpeedList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\wind_218_18_08-05-2008 09-58.img")
        WeatherWindSpeedList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\wind_234_18_08-21-2008 09-58.img")
        WeatherWindSpeedList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\wind_259_18_09-15-2008 09-58.img")
        WeatherWindSpeedList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\wind_275_18_10-01-2008 09-58.img")
        WeatherWindSpeedList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\wind_291_18_10-17-2008 09-58.img")
        WeatherWindSpeedList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\wind_314_18_11-09-2008 09-58.img")
        WeatherWindSpeedList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\wind_323_18_11-18-2008 09-58.img")

        WeatherRadiationList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\dswrf_019_18_01-19-2008 09-58.img")
        WeatherRadiationList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\dswrf_042_18_02-11-2008 09-58.img")
        WeatherRadiationList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\dswrf_058_18_02-27-2008 09-58.img")
        WeatherRadiationList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\dswrf_067_18_03-07-2008 09-58.img")
        WeatherRadiationList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\dswrf_083_18_03-23-2008 09-58.img")
        WeatherRadiationList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\dswrf_099_18_04-08-2008 09-58.img")
        WeatherRadiationList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\dswrf_115_18_04-24-2008 09-58.img")
        WeatherRadiationList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\dswrf_131_18_05-10-2008 09-58.img")
        WeatherRadiationList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\dswrf_138_18_05-17-2008 10-04.img")
        WeatherRadiationList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\dswrf_147_18_05-26-2008 09-58.img")
        WeatherRadiationList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\dswrf_163_18_06-11-2008 09-58.img")
        WeatherRadiationList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\dswrf_170_18_06-18-2008 10-03.img")
        WeatherRadiationList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\dswrf_195_18_07-13-2008 09-58.img")
        WeatherRadiationList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\dswrf_211_18_07-29-2008 09-57.img")
        WeatherRadiationList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\dswrf_218_18_08-05-2008 09-58.img")
        WeatherRadiationList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\dswrf_234_18_08-21-2008 09-58.img")
        WeatherRadiationList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\dswrf_259_18_09-15-2008 09-58.img")
        WeatherRadiationList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\dswrf_275_18_10-01-2008 09-58.img")
        WeatherRadiationList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\dswrf_291_18_10-17-2008 09-58.img")
        WeatherRadiationList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\dswrf_314_18_11-09-2008 09-58.img")
        WeatherRadiationList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\dswrf_323_18_11-18-2008 09-58.img")

        WeatherPressureList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\pres_019_18_01-19-2008 09-58.img")
        WeatherPressureList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\pres_042_18_02-11-2008 09-58.img")
        WeatherPressureList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\pres_058_18_02-27-2008 09-58.img")
        WeatherPressureList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\pres_067_18_03-07-2008 09-58.img")
        WeatherPressureList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\pres_083_18_03-23-2008 09-58.img")
        WeatherPressureList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\pres_099_18_04-08-2008 09-58.img")
        WeatherPressureList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\pres_115_18_04-24-2008 09-58.img")
        WeatherPressureList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\pres_131_18_05-10-2008 09-58.img")
        WeatherPressureList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\pres_138_18_05-17-2008 10-04.img")
        WeatherPressureList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\pres_147_18_05-26-2008 09-58.img")
        WeatherPressureList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\pres_163_18_06-11-2008 09-58.img")
        WeatherPressureList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\pres_170_18_06-18-2008 10-03.img")
        WeatherPressureList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\pres_195_18_07-13-2008 09-58.img")
        WeatherPressureList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\pres_211_18_07-29-2008 09-57.img")
        WeatherPressureList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\pres_218_18_08-05-2008 09-58.img")
        WeatherPressureList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\pres_234_18_08-21-2008 09-58.img")
        WeatherPressureList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\pres_259_18_09-15-2008 09-58.img")
        WeatherPressureList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\pres_275_18_10-01-2008 09-58.img")
        WeatherPressureList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\pres_291_18_10-17-2008 09-58.img")
        WeatherPressureList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\pres_314_18_11-09-2008 09-58.img")
        WeatherPressureList.Items.Add("E:\AGU_2013_PVID\gridded_data\projected\pres_323_18_11-18-2008 09-58.img")

        'WeatherETInstantaneousList.Items.Add("E:\AGU_2013_PVID\reference_ET\raster_projected\hourly\ETrc_hr_18_01-19-2008 09-58.img")
        'WeatherETInstantaneousList.Items.Add("E:\AGU_2013_PVID\reference_ET\raster_projected\hourly\ETrc_hr_18_02-11-2008 09-58.img")
        'WeatherETInstantaneousList.Items.Add("E:\AGU_2013_PVID\reference_ET\raster_projected\hourly\ETrc_hr_18_02-27-2008 09-58.img")
        'WeatherETInstantaneousList.Items.Add("E:\AGU_2013_PVID\reference_ET\raster_projected\hourly\ETrc_hr_18_03-07-2008 09-58.img")
        'WeatherETInstantaneousList.Items.Add("E:\AGU_2013_PVID\reference_ET\raster_projected\hourly\ETrc_hr_18_03-23-2008 09-58.img")
        'WeatherETInstantaneousList.Items.Add("E:\AGU_2013_PVID\reference_ET\raster_projected\hourly\ETrc_hr_18_04-08-2008 09-58.img")
        'WeatherETInstantaneousList.Items.Add("E:\AGU_2013_PVID\reference_ET\raster_projected\hourly\ETrc_hr_18_04-24-2008 09-58.img")
        'WeatherETInstantaneousList.Items.Add("E:\AGU_2013_PVID\reference_ET\raster_projected\hourly\ETrc_hr_18_05-10-2008 09-58.img")
        'WeatherETInstantaneousList.Items.Add("E:\AGU_2013_PVID\reference_ET\raster_projected\hourly\ETrc_hr_18_05-17-2008 10-04.img")
        'WeatherETInstantaneousList.Items.Add("E:\AGU_2013_PVID\reference_ET\raster_projected\hourly\ETrc_hr_18_05-26-2008 09-58.img")
        'WeatherETInstantaneousList.Items.Add("E:\AGU_2013_PVID\reference_ET\raster_projected\hourly\ETrc_hr_18_06-11-2008 09-58.img")
        'WeatherETInstantaneousList.Items.Add("E:\AGU_2013_PVID\reference_ET\raster_projected\hourly\ETrc_hr_18_06-18-2008 10-03.img")
        'WeatherETInstantaneousList.Items.Add("E:\AGU_2013_PVID\reference_ET\raster_projected\hourly\ETrc_hr_18_07-13-2008 09-58.img")
        'WeatherETInstantaneousList.Items.Add("E:\AGU_2013_PVID\reference_ET\raster_projected\hourly\ETrc_hr_18_07-29-2008 09-57.img")
        'WeatherETInstantaneousList.Items.Add("E:\AGU_2013_PVID\reference_ET\raster_projected\hourly\ETrc_hr_18_08-05-2008 09-58.img")
        'WeatherETInstantaneousList.Items.Add("E:\AGU_2013_PVID\reference_ET\raster_projected\hourly\ETrc_hr_18_08-21-2008 09-58.img")
        'WeatherETInstantaneousList.Items.Add("E:\AGU_2013_PVID\reference_ET\raster_projected\hourly\ETrc_hr_18_09-15-2008 09-58.img")
        'WeatherETInstantaneousList.Items.Add("E:\AGU_2013_PVID\reference_ET\raster_projected\hourly\ETrc_hr_18_10-01-2008 09-58.img")
        'WeatherETInstantaneousList.Items.Add("E:\AGU_2013_PVID\reference_ET\raster_projected\hourly\ETrc_hr_18_10-17-2008 09-58.img")
        'WeatherETInstantaneousList.Items.Add("E:\AGU_2013_PVID\reference_ET\raster_projected\hourly\ETrc_hr_18_11-09-2008 09-58.img")
        'WeatherETInstantaneousList.Items.Add("E:\AGU_2013_PVID\reference_ET\raster_projected\hourly\ETrc_hr_18_11-18-2008 09-58.img")

        'WeatherETDailyReferenceList.Items.Add("E:\AGU_2013_PVID\reference_ET\raster_projected\daily_new\019_ref_proj_01-19-2008 09-58.img")
        'WeatherETDailyReferenceList.Items.Add("E:\AGU_2013_PVID\reference_ET\raster_projected\daily_new\042_ref_proj_02-11-2008 09-58.img")
        'WeatherETDailyReferenceList.Items.Add("E:\AGU_2013_PVID\reference_ET\raster_projected\daily_new\058_ref_proj_02-27-2008 09-58.img")
        'WeatherETDailyReferenceList.Items.Add("E:\AGU_2013_PVID\reference_ET\raster_projected\daily_new\067_ref_proj_03-07-2008 09-58.img")
        'WeatherETDailyReferenceList.Items.Add("E:\AGU_2013_PVID\reference_ET\raster_projected\daily_new\083_ref_proj_03-23-2008 09-58.img")
        'WeatherETDailyReferenceList.Items.Add("E:\AGU_2013_PVID\reference_ET\raster_projected\daily_new\099_ref_proj_04-08-2008 09-58.img")
        'WeatherETDailyReferenceList.Items.Add("E:\AGU_2013_PVID\reference_ET\raster_projected\daily_new\115_ref_proj_04-24-2008 09-58.img")
        'WeatherETDailyReferenceList.Items.Add("E:\AGU_2013_PVID\reference_ET\raster_projected\daily_new\131_ref_proj_05-10-2008 09-58.img")
        'WeatherETDailyReferenceList.Items.Add("E:\AGU_2013_PVID\reference_ET\raster_projected\daily_new\138_ref_proj_05-17-2008 10-04.img")
        'WeatherETDailyReferenceList.Items.Add("E:\AGU_2013_PVID\reference_ET\raster_projected\daily_new\147_ref_proj_05-26-2008 09-58.img")
        'WeatherETDailyReferenceList.Items.Add("E:\AGU_2013_PVID\reference_ET\raster_projected\daily_new\163_ref_proj_06-11-2008 09-58.img")
        'WeatherETDailyReferenceList.Items.Add("E:\AGU_2013_PVID\reference_ET\raster_projected\daily_new\170_ref_proj_06-18-2008 10-03.img")
        'WeatherETDailyReferenceList.Items.Add("E:\AGU_2013_PVID\reference_ET\raster_projected\daily_new\195_ref_proj_07-13-2008 09-58.img")
        'WeatherETDailyReferenceList.Items.Add("E:\AGU_2013_PVID\reference_ET\raster_projected\daily_new\211_ref_proj_07-29-2008 09-57.img")
        'WeatherETDailyReferenceList.Items.Add("E:\AGU_2013_PVID\reference_ET\raster_projected\daily_new\218_ref_proj_08-05-2008 09-58.img")
        'WeatherETDailyReferenceList.Items.Add("E:\AGU_2013_PVID\reference_ET\raster_projected\daily_new\234_ref_proj_08-21-2008 09-58.img")
        'WeatherETDailyReferenceList.Items.Add("E:\AGU_2013_PVID\reference_ET\raster_projected\daily_new\259_ref_proj_09-15-2008 09-58.img")
        'WeatherETDailyReferenceList.Items.Add("E:\AGU_2013_PVID\reference_ET\raster_projected\daily_new\275_ref_proj_10-01-2008 09-58.img")
        'WeatherETDailyReferenceList.Items.Add("E:\AGU_2013_PVID\reference_ET\raster_projected\daily_new\291_ref_proj_10-17-2008 09-58.img")
        'WeatherETDailyReferenceList.Items.Add("E:\AGU_2013_PVID\reference_ET\raster_projected\daily_new\314_ref_proj_11-09-2008 09-58.img")
        'WeatherETDailyReferenceList.Items.Add("E:\AGU_2013_PVID\reference_ET\raster_projected\daily_new\323_ref_proj_11-18-2008 09-58.img")

    End Sub

    Sub AGU_2013_point()
        'Dim Path As String = "D:\SETMI\SETMI_old\"
        EnergyBalanceBox.SelectedIndex = 1
        ETExtrapolationBox.SelectedIndex = 1
        'WaterBalanceBox.SelectedIndex = 0
        ImageSourceBox.SelectedIndex = 1
        'MultispectralList.Items.Add("E:\agu_2013_pvid\Landsat_images\PVID_subset\3banb_pvid_019_2008_tucson_sub_01-19-2008 09-58.img")
        'MultispectralList.Items.Add("E:\agu_2013_pvid\Landsat_images\PVID_subset\3banb_pvid_042_2008_tucson_sub_02-11-2008 09-58.img")
        'MultispectralList.Items.Add("E:\agu_2013_pvid\Landsat_images\PVID_subset\3banb_pvid_058_2008_tucson_sub_02-27-2008 09-58.img")
        'MultispectralList.Items.Add("E:\agu_2013_pvid\Landsat_images\PVID_subset\3banb_pvid_067_2008_tucson_sub_03-07-2008 09-58.img")
        'MultispectralList.Items.Add("E:\agu_2013_pvid\Landsat_images\PVID_subset\3banb_pvid_083_2008_tucson_sub_03-23-2008 09-58.img")
        'MultispectralList.Items.Add("E:\agu_2013_pvid\Landsat_images\PVID_subset\3banb_pvid_099_2008_tucson_sub_04-08-2008 09-58.img")
        'MultispectralList.Items.Add("E:\agu_2013_pvid\Landsat_images\PVID_subset\3banb_pvid_115_2008_tucson_sub_04-24-2008 09-58.img")
        MultispectralList.Items.Add("E:\agu_2013_pvid\Landsat_images\PVID_subset\3banb_pvid_131_2008_tucson_sub_05-10-2008 09-58.img")
        'MultispectralList.Items.Add("E:\agu_2013_pvid\Landsat_images\PVID_subset\3banb_pvid_138_2008_tucson_sub_05-17-2008 10-04.img")
        'MultispectralList.Items.Add("E:\agu_2013_pvid\Landsat_images\PVID_subset\3banb_pvid_147_2008_tucson_sub_05-26-2008 09-58.img")
        'MultispectralList.Items.Add("E:\agu_2013_pvid\Landsat_images\PVID_subset\3banb_pvid_163_2008_tucson_sub_06-11-2008 09-58.img")
        'MultispectralList.Items.Add("E:\agu_2013_pvid\Landsat_images\PVID_subset\3banb_pvid_170_2008_tucson_sub_06-18-2008 10-03.img")
        'MultispectralList.Items.Add("E:\agu_2013_pvid\Landsat_images\PVID_subset\3banb_pvid_195_2008_tucson_sub_07-13-2008 09-58.img")
        'MultispectralList.Items.Add("E:\agu_2013_pvid\Landsat_images\PVID_subset\3banb_pvid_211_2008_tucson_sub_07-29-2008 09-57.img")
        'MultispectralList.Items.Add("E:\agu_2013_pvid\Landsat_images\PVID_subset\3banb_pvid_218_2008_tucson_sub_08-05-2008 09-58.img")
        'MultispectralList.Items.Add("E:\agu_2013_pvid\Landsat_images\PVID_subset\3banb_pvid_234_2008_tucson_sub_08-21-2008 09-58.img")
        'MultispectralList.Items.Add("E:\agu_2013_pvid\Landsat_images\PVID_subset\3banb_pvid_259_2008_tucson_sub_09-15-2008 09-58.img")
        'MultispectralList.Items.Add("E:\agu_2013_pvid\Landsat_images\PVID_subset\3banb_pvid_275_2008_tucson_sub_10-01-2008 09-58.img")
        'MultispectralList.Items.Add("E:\agu_2013_pvid\Landsat_images\PVID_subset\3banb_pvid_291_2008_tucson_sub_10-17-2008 09-58.img")
        'MultispectralList.Items.Add("E:\agu_2013_pvid\Landsat_images\PVID_subset\3banb_pvid_314_2008_tucson_sub_11-09-2008 09-58.img")
        'MultispectralList.Items.Add("E:\agu_2013_pvid\Landsat_images\PVID_subset\3banb_pvid_323_2008_tucson_sub_11-18-2008 09-58.img")
        Dim RasterDataset = CreateRasterDataset(MultispectralList.Items(0))
        Dim RasterBandCollection As ESRI.ArcGIS.DataSourcesRaster.IRasterBandCollection = RasterDataset
        For Item = 1 To RasterBandCollection.Count
            RedIndex.Items.Add(Item)
            GreenIndex.Items.Add(Item)
            BlueIndex.Items.Add(Item)
            NIRIndex.Items.Add(Item)
            MIR1Index.Items.Add(Item)
            MIR2Index.Items.Add(Item)
        Next
        System.Runtime.InteropServices.Marshal.ReleaseComObject(RasterBandCollection)
        SurfaceTemperatureList.Items.Add("E:\agu_2013_pvid\Landsat_images\PVID_subset\pvid_thermal_c_019_2008_corrected_tucson_sub_01-19-2008 09-58.img")
        SurfaceTemperatureList.Items.Add("E:\agu_2013_pvid\Landsat_images\PVID_subset\pvid_thermal_c_042_2008_corrected_tucson_sub_02-11-2008 09-58.img")
        SurfaceTemperatureList.Items.Add("E:\agu_2013_pvid\Landsat_images\PVID_subset\pvid_thermal_c_058_2008_corrected_tucson_sub_02-27-2008 09-58.img")
        SurfaceTemperatureList.Items.Add("E:\agu_2013_pvid\Landsat_images\PVID_subset\pvid_thermal_c_067_2008_corrected_tucson_sub_03-07-2008 09-58.img")
        SurfaceTemperatureList.Items.Add("E:\agu_2013_pvid\Landsat_images\PVID_subset\pvid_thermal_c_083_2008_corrected_tucson_sub_03-23-2008 09-58.img")
        SurfaceTemperatureList.Items.Add("E:\agu_2013_pvid\Landsat_images\PVID_subset\pvid_thermal_c_099_2008_corrected_tucson_sub_04-08-2008 09-58.img")
        SurfaceTemperatureList.Items.Add("E:\agu_2013_pvid\Landsat_images\PVID_subset\pvid_thermal_c_115_2008_corrected_tucson_sub_04-24-2008 09-58.img")
        SurfaceTemperatureList.Items.Add("E:\agu_2013_pvid\Landsat_images\PVID_subset\pvid_thermal_c_131_2008_corrected_tucson_sub_05-10-2008 09-58.img")
        SurfaceTemperatureList.Items.Add("E:\agu_2013_pvid\Landsat_images\PVID_subset\pvid_thermal_c_138_2008_corrected_tucson_sub_05-17-2008 10-04.img")
        SurfaceTemperatureList.Items.Add("E:\agu_2013_pvid\Landsat_images\PVID_subset\pvid_thermal_c_147_2008_corrected_tucson_sub_05-26-2008 09-58.img")
        SurfaceTemperatureList.Items.Add("E:\agu_2013_pvid\Landsat_images\PVID_subset\pvid_thermal_c_163_2008_corrected_tucson_sub_06-11-2008 09-58.img")
        SurfaceTemperatureList.Items.Add("E:\agu_2013_pvid\Landsat_images\PVID_subset\pvid_thermal_c_170_2008_corrected_tucson_sub_06-18-2008 10-03.img")
        SurfaceTemperatureList.Items.Add("E:\agu_2013_pvid\Landsat_images\PVID_subset\pvid_thermal_c_195_2008_corrected_tucson_sub_07-13-2008 09-58.img")
        SurfaceTemperatureList.Items.Add("E:\agu_2013_pvid\Landsat_images\PVID_subset\pvid_thermal_c_211_2008_corrected_tucson_sub_07-29-2008 09-57.img")
        SurfaceTemperatureList.Items.Add("E:\agu_2013_pvid\Landsat_images\PVID_subset\pvid_thermal_c_218_2008_corrected_tucson_sub_08-05-2008 09-58.img")
        SurfaceTemperatureList.Items.Add("E:\agu_2013_pvid\Landsat_images\PVID_subset\pvid_thermal_c_234_2008_corrected_tucson_sub_08-21-2008 09-58.img")
        SurfaceTemperatureList.Items.Add("E:\agu_2013_pvid\Landsat_images\PVID_subset\pvid_thermal_c_259_2008_corrected_tucson_sub_09-15-2008 09-58.img")
        SurfaceTemperatureList.Items.Add("E:\agu_2013_pvid\Landsat_images\PVID_subset\pvid_thermal_c_275_2008_corrected_tucson_sub_10-01-2008 09-58.img")
        SurfaceTemperatureList.Items.Add("E:\agu_2013_pvid\Landsat_images\PVID_subset\pvid_thermal_c_291_2008_corrected_tucson_sub_10-17-2008 09-58.img")
        SurfaceTemperatureList.Items.Add("E:\agu_2013_pvid\Landsat_images\PVID_subset\pvid_thermal_c_314_2008_corrected_tucson_sub_11-09-2008 09-58.img")
        SurfaceTemperatureList.Items.Add("E:\agu_2013_pvid\Landsat_images\PVID_subset\pvid_thermal_c_323_2008_corrected_tucson_sub_11-18-2008 09-58.img")
        CoverClassificationList.Items.Add("E:\agu_2013_pvid\landuse\Modified_2\Landuse_05-01-2008 10-00.img")
        Dim Raster As ESRI.ArcGIS.DataSourcesRaster.IRaster2 = CType(CreateRasterDataset(CoverClassificationList.Items(0)).CreateDefaultRaster, ESRI.ArcGIS.DataSourcesRaster.IRaster2)
        For F = 0 To Raster.AttributeTable.Fields.FieldCount - 1
            CoverSelectionBox.Items.Add(Raster.AttributeTable.Fields.Field(F).AliasName)
            If Raster.AttributeTable.Fields.Field(F).AliasName.Contains("Name") Or Raster.AttributeTable.Fields.Field(F).AliasName.Contains("NAME") Or Raster.AttributeTable.Fields.Field(F).AliasName.Contains("name") Then CoverSelectionBox.Text = Raster.AttributeTable.Fields.Field(F).AliasName
        Next
        System.Runtime.InteropServices.Marshal.ReleaseComObject(Raster)
        NIRIndex.SelectedIndex = 0
        RedIndex.SelectedIndex = 1
        GreenIndex.SelectedIndex = 2
        WeatherTableText.Text = "E:\agu_2013_pvid\weather_data\PVID_data_EB.xlsx\Point$"
        'CoverPropertiesText.Text = "E:\agu_2013_pvid\weather_data\Bearax_data_EB.xlsx\Sheet2$"
        OutputDirectoryTextEnergy.Text = "E:\agu_2013_pvid\testing_new_results_point_2014_refine"
        'OutputDirectoryTextWater.Text = "E:\SETMI\Output"
        OutputImagesCheckAllEnergy_Click(New Object, New System.EventArgs)
        'OutputImagesCheckAllWater_Click(New Object, New System.EventArgs)
        Dim Rows = {0, 1, 2, 3, 4, 5, 6, 7, 9, 12, 13, 14}
        For Row = 0 To CoverSelectionGrid.Rows.Count - 1
            If Not Rows.Contains(Row) Then
                CoverSelectionGrid.Rows(Row).Cells(3).Value = False
            Else
                CoverSelectionGrid.Rows(Row).Cells(2).Value = "Basin"
            End If
        Next
        CoverSelectionGrid.Rows(2).Cells(1).Value = "Wheat"
        CoverSelectionGrid.Rows(3).Cells(1).Value = "Wheat"
        CoverSelectionGrid.Rows(4).Cells(1).Value = "Agriculture"
        CoverSelectionGrid.Rows(5).Cells(1).Value = "Grass"
        CoverSelectionGrid.Rows(6).Cells(1).Value = "Watermelon"
        CoverSelectionGrid.Rows(7).Cells(1).Value = "Bare Soil"
        CoverSelectionGrid.Rows(9).Cells(1).Value = "Agriculture"
        CoverSelectionGrid.Rows(12).Cells(1).Value = "Grass"
        CoverSelectionGrid.Rows(13).Cells(1).Value = "Grass"
        CoverSelectionGrid.Rows(14).Cells(1).Value = "Citrus"
        TimeZoneBox.SelectedIndex = TimeZoneBox.Items.IndexOf("Pacific Standard Time")
        Dim Output = {3, 4, 6, 9, 11}
        For Item = 0 To OutputImagesBoxEnergy.Items.Count - 1
            If Not Output.Contains(Item) Then OutputImagesBoxEnergy.SetItemChecked(Item, False)
        Next
        'FieldCapacityText.Text = "D:\SETMI\USA Field Capacity.img"
        'WiltingPointText.Text = "D:\SETMI\USA Wilting Point.img"


    End Sub

    Sub Sample_Iowa_MODIS_TSEB()
        EnergyBalanceBox.SelectedIndex = 1
        ETExtrapolationBox.SelectedIndex = 1
        'WaterBalanceBox.SelectedIndex = 0
        'DataAssimilationBox.SelectedIndex = 0
        ImageSourceBox.SelectedIndex = 1
        'MultispectralList.Items.Add("D:\SETMI\SETMI_old\Iowa 2002\3band_thermal_landuse_SM\landsat_3band_reflectance_07-17-2002.img")
        MultispectralList.Items.Add("E:\ASC_2014_data\MODIS_data\analysis_data\reflectance_06-23-2002.img")
        MultispectralList.Items.Add("E:\ASC_2014_data\MODIS_data\analysis_data\reflectance_07-01-2002.img")
        MultispectralList.Items.Add("E:\ASC_2014_data\MODIS_data\analysis_data\reflectance_07-08-2002.img")
        MultispectralList.Items.Add("E:\ASC_2014_data\MODIS_data\analysis_data\reflectance_07-16-2002.img")
        Dim RasterDataset = CreateRasterDataset(MultispectralList.Items(0))
        Dim RasterBandCollection As ESRI.ArcGIS.DataSourcesRaster.IRasterBandCollection = RasterDataset
        For Item = 1 To RasterBandCollection.Count
            RedIndex.Items.Add(Item)
            GreenIndex.Items.Add(Item)
            BlueIndex.Items.Add(Item)
            NIRIndex.Items.Add(Item)
            MIR1Index.Items.Add(Item)
            MIR2Index.Items.Add(Item)
        Next
        System.Runtime.InteropServices.Marshal.ReleaseComObject(RasterBandCollection)
        'SurfaceTemperatureList.Items.Add("D:\SETMI\SETMI_old\Iowa 2002\3band_thermal_landuse_SM\temp_07-17-2002.img")
        SurfaceTemperatureList.Items.Add("E:\ASC_2014_data\MODIS_data\analysis_data\temp_fine_with_delta_nan2_06-23-2002.img")
        SurfaceTemperatureList.Items.Add("E:\ASC_2014_data\MODIS_data\analysis_data\temp_fine_with_delta_nan2_07-01-2002.img")
        SurfaceTemperatureList.Items.Add("E:\ASC_2014_data\MODIS_data\analysis_data\temp_fine_with_delta_nan2_07-08-2002.img")
        SurfaceTemperatureList.Items.Add("E:\ASC_2014_data\MODIS_data\analysis_data\temp_fine_with_delta_nan2_07-16-2002.img")
        CoverClassificationList.Items.Add("E:\ASC_2014_data\MODIS_data\analysis_data\landuse_07-01-2002.img")
        Dim Raster As ESRI.ArcGIS.DataSourcesRaster.IRaster2 = CType(CreateRasterDataset(CoverClassificationList.Items(0)).CreateDefaultRaster, ESRI.ArcGIS.DataSourcesRaster.IRaster2)
        For F = 0 To Raster.AttributeTable.Fields.FieldCount - 1
            CoverSelectionBox.Items.Add(Raster.AttributeTable.Fields.Field(F).AliasName)
            If Raster.AttributeTable.Fields.Field(F).AliasName.Contains("Name") Or Raster.AttributeTable.Fields.Field(F).AliasName.Contains("NAME") Or Raster.AttributeTable.Fields.Field(F).AliasName.Contains("name") Then CoverSelectionBox.Text = Raster.AttributeTable.Fields.Field(F).AliasName
        Next
        System.Runtime.InteropServices.Marshal.ReleaseComObject(Raster)
        NIRIndex.SelectedIndex = 0
        RedIndex.SelectedIndex = 1
        'GreenIndex.SelectedIndex = 2
        WeatherTableText.Text = "E:\ASC_2014_data\MODIS_data\Iowa_weather_data.xls\Sheet4$"
        'CoverPropertiesText.Text = "W:\SETMI_Visual_Studio\Iowa_data_for_Hybrid_ET_paper_one\cropProps.xlsx\Cover$"
        OutputDirectoryTextEnergy.Text = "E:\ASC_2014_data\MODIS_data\results_TSEB"
        'OutputDirectoryTextWater.Text = "W:\SETMI_Visual_Studio\Iowa_data_for_Hybrid_ET_paper_one\results"
        OutputImagesCheckAllEnergy_Click(New Object, New System.EventArgs)
        'OutputImagesCheckAllWater_Click(New Object, New System.EventArgs)
        'WeatherETDailyActualList.Items.Add("W:\SETMI_Visual_Studio\Iowa_data_for_Hybrid_ET_paper_one\ET_maps\TSMET_06-23-2002.img")
        'WeatherETDailyActualList.Items.Add("W:\SETMI_Visual_Studio\Iowa_data_for_Hybrid_ET_paper_one\ET_maps\TSMET_07-01-2002.img")
        'WeatherETDailyActualList.Items.Add("W:\SETMI_Visual_Studio\Iowa_data_for_Hybrid_ET_paper_one\ET_maps\TSMET_07-08-2002.img")
        'WeatherETDailyActualList.Items.Add("W:\SETMI_Visual_Studio\Iowa_data_for_Hybrid_ET_paper_one\ET_maps\TSMET_07-16-2002.img")
        Dim Rows = {2, 4}
        For Row = 0 To CoverSelectionGrid.Rows.Count - 1
            If Not Rows.Contains(Row) Then
                CoverSelectionGrid.Rows(Row).Cells(3).Value = False
            Else
                CoverSelectionGrid.Rows(Row).Cells(2).Value = "Not Irrigated"
            End If
        Next
        TimeZoneBox.SelectedIndex = TimeZoneBox.Items.IndexOf("Central Standard Time")
        Dim Output = {3, 4, 6, 9, 11}
        For Item = 0 To OutputImagesBoxEnergy.Items.Count - 1
            If Not Output.Contains(Item) Then OutputImagesBoxEnergy.SetItemChecked(Item, False)
        Next

        'Dim Output_2 = {15}
        'For Item = 0 To OutputImagesBoxWater.Items.Count - 1
        '    If Not Output_2.Contains(Item) Then OutputImagesBoxWater.SetItemChecked(Item, False)
        'Next
        'FieldCapacityText.Text = "W:\SETMI_Visual_Studio\Iowa_data_for_Hybrid_ET_paper_one\3band_thermal_landuse_SM\fc_map_07-01-2002.img"
        'WiltingPointText.Text = "W:\SETMI_Visual_Studio\Iowa_data_for_Hybrid_ET_paper_one\3band_thermal_landuse_SM\pwp_map_07-01-2002.img"
    End Sub

    Sub Sample_Iowa_MODIS_hybrid()
        'EnergyBalanceBox.SelectedIndex = 1
        'ETExtrapolationBox.SelectedIndex = 0
        WaterBalanceBox.SelectedIndex = 0
        DataAssimilationBox.SelectedIndex = 0
        ImageSourceBox.SelectedIndex = 1
        'MultispectralList.Items.Add("D:\SETMI\SETMI_old\Iowa 2002\3band_thermal_landuse_SM\landsat_3band_reflectance_07-17-2002.img")
        MultispectralList.Items.Add("E:\ASC_2014_data\MODIS_data\analysis_data\reflectance_06-23-2002.img")
        MultispectralList.Items.Add("E:\ASC_2014_data\MODIS_data\analysis_data\reflectance_07-01-2002.img")
        MultispectralList.Items.Add("E:\ASC_2014_data\MODIS_data\analysis_data\reflectance_07-08-2002.img")
        'MultispectralList.Items.Add("E:\ASC_2014_data\MODIS_data\analysis_data\reflectance_07-16-2002.img")
        Dim RasterDataset = CreateRasterDataset(MultispectralList.Items(0))
        Dim RasterBandCollection As ESRI.ArcGIS.DataSourcesRaster.IRasterBandCollection = RasterDataset
        For Item = 1 To RasterBandCollection.Count
            RedIndex.Items.Add(Item)
            GreenIndex.Items.Add(Item)
            BlueIndex.Items.Add(Item)
            NIRIndex.Items.Add(Item)
            MIR1Index.Items.Add(Item)
            MIR2Index.Items.Add(Item)
        Next
        System.Runtime.InteropServices.Marshal.ReleaseComObject(RasterBandCollection)
        'SurfaceTemperatureList.Items.Add("D:\SETMI\SETMI_old\Iowa 2002\3band_thermal_landuse_SM\temp_07-17-2002.img")
        'SurfaceTemperatureList.Items.Add("E:\ASC_2014_data\MODIS_data\analysis_data\temp_fine_with_delta_nan2_06-23-2002.img")
        'SurfaceTemperatureList.Items.Add("E:\ASC_2014_data\MODIS_data\analysis_data\temp_fine_with_delta_nan2_07-01-2002.img")
        'SurfaceTemperatureList.Items.Add("E:\ASC_2014_data\MODIS_data\analysis_data\temp_fine_with_delta_nan2_07-08-2002.img")
        'SurfaceTemperatureList.Items.Add("E:\ASC_2014_data\MODIS_data\analysis_data\temp_fine_with_delta_nan2_07-16-2002.img")
        CoverClassificationList.Items.Add("E:\ASC_2014_data\MODIS_data\analysis_data\landuse_07-01-2002.img")
        Dim Raster As ESRI.ArcGIS.DataSourcesRaster.IRaster2 = CType(CreateRasterDataset(CoverClassificationList.Items(0)).CreateDefaultRaster, ESRI.ArcGIS.DataSourcesRaster.IRaster2)
        For F = 0 To Raster.AttributeTable.Fields.FieldCount - 1
            CoverSelectionBox.Items.Add(Raster.AttributeTable.Fields.Field(F).AliasName)
            If Raster.AttributeTable.Fields.Field(F).AliasName.Contains("Name") Or Raster.AttributeTable.Fields.Field(F).AliasName.Contains("NAME") Or Raster.AttributeTable.Fields.Field(F).AliasName.Contains("name") Then CoverSelectionBox.Text = Raster.AttributeTable.Fields.Field(F).AliasName
        Next
        System.Runtime.InteropServices.Marshal.ReleaseComObject(Raster)
        NIRIndex.SelectedIndex = 0
        RedIndex.SelectedIndex = 1
        'GreenIndex.SelectedIndex = 2
        WeatherTableText.Text = "E:\ASC_2014_data\MODIS_data\cropProps.xlsx\meteo$"
        CoverPropertiesText.Text = "E:\ASC_2014_data\MODIS_data\cropProps.xlsx\Cover$"
        'OutputDirectoryTextEnergy.Text = "W:\SETMI_Visual_Studio\Iowa_data_for_Hybrid_ET_paper_one\results"
        OutputDirectoryTextWater.Text = "E:\ASC_2014_data\MODIS_data\results_Hybrid"
        OutputImagesCheckAllEnergy_Click(New Object, New System.EventArgs)
        OutputImagesCheckAllWater_Click(New Object, New System.EventArgs)
        WeatherETDailyActualList.Items.Add("E:\ASC_2014_data\MODIS_data\results_TSEB\SETMI Two_Source ET 06-23-2002.img")
        WeatherETDailyActualList.Items.Add("E:\ASC_2014_data\MODIS_data\results_TSEB\SETMI Two_Source ET 07-01-2002.img")
        WeatherETDailyActualList.Items.Add("E:\ASC_2014_data\MODIS_data\results_TSEB\SETMI Two_Source ET 07-08-2002.img")
        'WeatherETDailyActualList.Items.Add("E:\ASC_2014_data\MODIS_data\results_TSEB\TSMET_07-16-2002.img")
        Dim Rows = {2, 4}
        For Row = 0 To CoverSelectionGrid.Rows.Count - 1
            If Not Rows.Contains(Row) Then
                CoverSelectionGrid.Rows(Row).Cells(3).Value = False
            Else
                CoverSelectionGrid.Rows(Row).Cells(2).Value = "Not Irrigated"
            End If
        Next
        TimeZoneBox.SelectedIndex = TimeZoneBox.Items.IndexOf("Central Standard Time")
        Dim Output = {15}
        For Item = 0 To OutputImagesBoxEnergy.Items.Count - 1
            If Not Output.Contains(Item) Then OutputImagesBoxEnergy.SetItemChecked(Item, False)
        Next

        Dim Output_2 = {15}
        For Item = 0 To OutputImagesBoxWater.Items.Count - 1
            If Not Output_2.Contains(Item) Then OutputImagesBoxWater.SetItemChecked(Item, False)
        Next
        FieldCapacityText.Text = "E:\ASC_2014_data\MODIS_data\analysis_data\fc_map_07-01-2002.img"
        WiltingPointText.Text = "E:\ASC_2014_data\MODIS_data\analysis_data\pwp_map_07-01-2002.img"
    End Sub

    Sub Sample_Iowa_Landsat_hybrid()
        'EnergyBalanceBox.SelectedIndex = 1
        'ETExtrapolationBox.SelectedIndex = 0
        WaterBalanceBox.SelectedIndex = 0
        DataAssimilationBox.SelectedIndex = 0
        ImageSourceBox.SelectedIndex = 1
        'MultispectralList.Items.Add("D:\SETMI\SETMI_old\Iowa 2002\3band_thermal_landuse_SM\landsat_3band_reflectance_07-17-2002.img")
        MultispectralList.Items.Add("W:\SETMI_Visual_Studio\Iowa_data_for_Hybrid_ET_paper_one\3band_thermal_landuse_SM\landsat_3band_reflectance_06-23-2002.img")
        MultispectralList.Items.Add("W:\SETMI_Visual_Studio\Iowa_data_for_Hybrid_ET_paper_one\3band_thermal_landuse_SM\landsat_3band_reflectance_07-01-2002.img")
        MultispectralList.Items.Add("W:\SETMI_Visual_Studio\Iowa_data_for_Hybrid_ET_paper_one\3band_thermal_landuse_SM\landsat_3band_reflectance_07-08-2002.img")
        MultispectralList.Items.Add("W:\SETMI_Visual_Studio\Iowa_data_for_Hybrid_ET_paper_one\3band_thermal_landuse_SM\landsat_3band_reflectance_07-16-2002.img")
        Dim RasterDataset = CreateRasterDataset(MultispectralList.Items(0))
        Dim RasterBandCollection As ESRI.ArcGIS.DataSourcesRaster.IRasterBandCollection = RasterDataset
        For Item = 1 To RasterBandCollection.Count
            RedIndex.Items.Add(Item)
            GreenIndex.Items.Add(Item)
            BlueIndex.Items.Add(Item)
            NIRIndex.Items.Add(Item)
            MIR1Index.Items.Add(Item)
            MIR2Index.Items.Add(Item)
        Next
        System.Runtime.InteropServices.Marshal.ReleaseComObject(RasterBandCollection)
        'SurfaceTemperatureList.Items.Add("D:\SETMI\SETMI_old\Iowa 2002\3band_thermal_landuse_SM\temp_07-17-2002.img")
        SurfaceTemperatureList.Items.Add("W:\SETMI_Visual_Studio\Iowa_data_for_Hybrid_ET_paper_one\3band_thermal_landuse_SM\temp_06-23-2002.img")
        SurfaceTemperatureList.Items.Add("W:\SETMI_Visual_Studio\Iowa_data_for_Hybrid_ET_paper_one\3band_thermal_landuse_SM\temp_07-01-2002.img")
        SurfaceTemperatureList.Items.Add("W:\SETMI_Visual_Studio\Iowa_data_for_Hybrid_ET_paper_one\3band_thermal_landuse_SM\temp_07-08-2002.img")
        SurfaceTemperatureList.Items.Add("W:\SETMI_Visual_Studio\Iowa_data_for_Hybrid_ET_paper_one\3band_thermal_landuse_SM\temp_07-16-2002.img")
        CoverClassificationList.Items.Add("W:\SETMI_Visual_Studio\Iowa_data_for_Hybrid_ET_paper_one\3band_thermal_landuse_SM\landuse_07-01-2002.img")
        Dim Raster As ESRI.ArcGIS.DataSourcesRaster.IRaster2 = CType(CreateRasterDataset(CoverClassificationList.Items(0)).CreateDefaultRaster, ESRI.ArcGIS.DataSourcesRaster.IRaster2)
        For F = 0 To Raster.AttributeTable.Fields.FieldCount - 1
            CoverSelectionBox.Items.Add(Raster.AttributeTable.Fields.Field(F).AliasName)
            If Raster.AttributeTable.Fields.Field(F).AliasName.Contains("Name") Or Raster.AttributeTable.Fields.Field(F).AliasName.Contains("NAME") Or Raster.AttributeTable.Fields.Field(F).AliasName.Contains("name") Then CoverSelectionBox.Text = Raster.AttributeTable.Fields.Field(F).AliasName
        Next
        System.Runtime.InteropServices.Marshal.ReleaseComObject(Raster)
        NIRIndex.SelectedIndex = 0
        RedIndex.SelectedIndex = 1
        GreenIndex.SelectedIndex = 2
        WeatherTableText.Text = "W:\SETMI_Visual_Studio\Iowa_data_for_Hybrid_ET_paper_one\cropProps.xlsx\meteo$"
        CoverPropertiesText.Text = "W:\SETMI_Visual_Studio\Iowa_data_for_Hybrid_ET_paper_one\cropProps.xlsx\Cover$"
        'OutputDirectoryTextEnergy.Text = "W:\SETMI_Visual_Studio\Iowa_data_for_Hybrid_ET_paper_one\results"
        OutputDirectoryTextWater.Text = "W:\SETMI_Visual_Studio\Iowa_data_for_Hybrid_ET_paper_one\results"
        OutputImagesCheckAllEnergy_Click(New Object, New System.EventArgs)
        OutputImagesCheckAllWater_Click(New Object, New System.EventArgs)
        WeatherETDailyActualList.Items.Add("W:\SETMI_Visual_Studio\Iowa_data_for_Hybrid_ET_paper_one\ET_maps\TSMET_06-23-2002.img")
        WeatherETDailyActualList.Items.Add("W:\SETMI_Visual_Studio\Iowa_data_for_Hybrid_ET_paper_one\ET_maps\TSMET_07-01-2002.img")
        WeatherETDailyActualList.Items.Add("W:\SETMI_Visual_Studio\Iowa_data_for_Hybrid_ET_paper_one\ET_maps\TSMET_07-08-2002.img")
        WeatherETDailyActualList.Items.Add("W:\SETMI_Visual_Studio\Iowa_data_for_Hybrid_ET_paper_one\ET_maps\TSMET_07-16-2002.img")
        Dim Rows = {2, 4}
        For Row = 0 To CoverSelectionGrid.Rows.Count - 1
            If Not Rows.Contains(Row) Then
                CoverSelectionGrid.Rows(Row).Cells(3).Value = False
            Else
                CoverSelectionGrid.Rows(Row).Cells(2).Value = "Not Irrigated"
            End If
        Next
        TimeZoneBox.SelectedIndex = TimeZoneBox.Items.IndexOf("Central Standard Time")
        Dim Output = {15}
        For Item = 0 To OutputImagesBoxEnergy.Items.Count - 1
            If Not Output.Contains(Item) Then OutputImagesBoxEnergy.SetItemChecked(Item, False)
        Next

        Dim Output_2 = {15}
        For Item = 0 To OutputImagesBoxWater.Items.Count - 1
            If Not Output_2.Contains(Item) Then OutputImagesBoxWater.SetItemChecked(Item, False)
        Next
        FieldCapacityText.Text = "W:\SETMI_Visual_Studio\Iowa_data_for_Hybrid_ET_paper_one\3band_thermal_landuse_SM\fc_map_07-01-2002.img"
        WiltingPointText.Text = "W:\SETMI_Visual_Studio\Iowa_data_for_Hybrid_ET_paper_one\3band_thermal_landuse_SM\pwp_map_07-01-2002.img"
    End Sub

#End Region

    Dim ProjectSaved As Boolean = False
    Dim ProjectSavePath As String = ""
    Dim NewSaveString As String = ""

    Private Sub NewProject_Click(sender As System.Object, e As System.EventArgs) Handles NewProject.Click
        If Not ExitWithoutSaving() Then Exit Sub

        Dim List = GetSortedListOfControls(Me)

        For Each Control In List
            Select Case Control.GetType().FullName
                Case "System.Windows.Forms.TextBox"
                    Dim C = CType(Control, Windows.Forms.TextBox)
                    C.Text = ""
                Case "System.Windows.Forms.ComboBox"
                    Dim C = CType(Control, Windows.Forms.ComboBox)
                    C.Text = ""
                    C.Items.Clear()
                Case "System.Windows.Forms.ListBox"
                    Dim C = CType(Control, Windows.Forms.ListBox)
                    C.Items.Clear()
                Case "System.Windows.Forms.CheckedListBox"
                    Dim C = CType(Control, Windows.Forms.CheckedListBox)
                    C.Items.Clear()
                Case "System.Windows.Forms.DataGridView"
                    Dim C = CType(Control, Windows.Forms.DataGridView)
                    C.Columns.Clear()
            End Select
        Next

        For Item = 0 To [Enum].GetNames(GetType(EnergyBalance)).Count - 1
            EnergyBalanceBox.Items.Add([Enum].GetName(GetType(EnergyBalance), Item).Replace("_", " "))
        Next
        For Item = 0 To [Enum].GetNames(GetType(ETExtrapolation)).Count - 1
            ETExtrapolationBox.Items.Add([Enum].GetName(GetType(ETExtrapolation), Item).Replace("_", " "))
        Next
        For Item = 0 To [Enum].GetNames(GetType(WaterBalance)).Count - 1
            WaterBalanceBox.Items.Add([Enum].GetName(GetType(WaterBalance), Item).Replace("_", " "))
        Next
        'WaterBalanceBox.SelectedIndex = 0
        For Item = 0 To [Enum].GetNames(GetType(DataAssimilation)).Count - 1
            DataAssimilationBox.Items.Add([Enum].GetName(GetType(DataAssimilation), Item).Replace("_", " "))
        Next
        For Item = 0 To [Enum].GetNames(GetType(ImageSource)).Count - 1
            ImageSourceBox.Items.Add([Enum].GetName(GetType(ImageSource), Item).Replace("_", " "))
        Next

        'Additional Options Added by JBB
        For Item = 0 To [Enum].GetNames(GetType(TSMInitialTemperature)).Count - 1 'added by JBB copying code above
            TSMInitialTemperatureBox.Items.Add([Enum].GetName(GetType(TSMInitialTemperature), Item).Replace("_", " "))
        Next
        For Item = 0 To [Enum].GetNames(GetType(ETReferenceType)).Count - 1 'added by JBB copying code above
            ETReferenceBox.Items.Add([Enum].GetName(GetType(ETReferenceType), Item).Replace("_", " "))
        Next
        For Item = 0 To [Enum].GetNames(GetType(EffectivePrecipType)).Count - 1 'added by JBB copying code above
            EffectivePreciptiationBox.Items.Add([Enum].GetName(GetType(EffectivePrecipType), Item).Replace("_", " "))
        Next
        For Item = 0 To [Enum].GetNames(GetType(SoilHeatFlux)).Count - 1 'added by JBB copying code above
            SoilHeatFluxBox.Items.Add([Enum].GetName(GetType(SoilHeatFlux), Item).Replace("_", " "))
        Next
        For Item = 0 To [Enum].GetNames(GetType(WindAdjustment)).Count - 1 'added by JBB copying code above
            WindAdjustMethodBox.Items.Add([Enum].GetName(GetType(WindAdjustment), Item).Replace("_", " "))
        Next
        For Item = 0 To [Enum].GetNames(GetType(KcbType)).Count - 1 'added by JBB copying code above
            KcbTypeBox.Items.Add([Enum].GetName(GetType(KcbType), Item).Replace("_", " "))
        Next
        For Item = 0 To [Enum].GetNames(GetType(FcType)).Count - 1 'added by JBB copying code above
            FcTypeBox.Items.Add([Enum].GetName(GetType(FcType), Item).Replace("_", " "))
        Next

        For Item = 0 To [Enum].GetNames(GetType(FgMethod)).Count - 1 'added by JBB copying code above
            FgMethodBox.Items.Add([Enum].GetName(GetType(FgMethod), Item).Replace("_", " "))
        Next
        For Item = 0 To [Enum].GetNames(GetType(HcMethod)).Count - 1 'added by JBB copying code above
            HcMethodBox.Items.Add([Enum].GetName(GetType(HcMethod), Item).Replace("_", " "))
        Next
        For Item = 0 To [Enum].GetNames(GetType(FcMethod)).Count - 1 'added by JBB copying code above
            FcMethodBox.Items.Add([Enum].GetName(GetType(FcMethod), Item).Replace("_", " "))
        Next

        For Item = 0 To [Enum].GetNames(GetType(ClumpingD)).Count - 1 'added by JBB copying code above
            ClumpingMethodBox.Items.Add([Enum].GetName(GetType(ClumpingD), Item).Replace("_", " "))
        Next
        For Item = 0 To [Enum].GetNames(GetType(SAVIForecast)).Count - 1 'added by JBB copying code above
            SAVIForecastMethodBox.Items.Add([Enum].GetName(GetType(SAVIForecast), Item).Replace("_", " "))
        Next
        For Item = 0 To [Enum].GetNames(GetType(DPLimit)).Count - 1 'added by JBB copying code above
            DPLimitMethodBox.Items.Add([Enum].GetName(GetType(DPLimit), Item).Replace("_", " "))
        Next
        For Item = 0 To [Enum].GetNames(GetType(WBPntMethod)).Count - 1 'added by JBB copying code above
            WBPntLocMethodBox.Items.Add([Enum].GetName(GetType(WBPntMethod), Item).Replace("_", " "))
        Next
        For Item = 0 To [Enum].GetNames(GetType(WBTeMethod)).Count - 1 'added by JBB copying code above
            WBTeMethodBox.Items.Add([Enum].GetName(GetType(WBTeMethod), Item).Replace("_", " "))
        Next
        For Item = 0 To [Enum].GetNames(GetType(WBSkinMethod)).Count - 1 'added by JBB copying code above
            SkinLyrMethodBox.Items.Add([Enum].GetName(GetType(WBSkinMethod), Item).Replace("_", " "))
        Next


        'Dim OutputImageListE = Split("Leaf Area Index (LAI),Soil Adjusted Vegetation Index (SAVI),Net Solar Radiation (Rn),Aerodynamic Resistance to Heat Transfer (rah),Soil Heat Flux (G),Latent Flux (LE),Normalized Difference Vegetation Index (NDVI),Optimized Soil Adjusted Vegetation Index (OSAVI),Albedo (α),Aerodynamic Temperature (To),Sensible Heat Flux (H),Evapotranspiration (ET)", ",")'Commented out by JBB
        ' Dim OutputImageListE = Split("Leaf Area Index (LAI),Soil Adjusted Vegetation Index (SAVI),Net Solar Radiation (Rn),Aerodynamic Resistance to Heat Transfer (rah),Soil Heat Flux (G),Latent Flux (LE),Normalized Difference Vegetation Index (NDVI),Optimized Soil Adjusted Vegetation Index (OSAVI),Albedo (α),Aerodynamic Temperature (To),Sensible Heat Flux (H),Evapotranspiration (ET),Fraction of Cover (fc),Priestly Taylor Coefficient (αPT),Obukhov Length (Lo),Canopy Height (Hc),Wind Speed (U),Latent Flux Canopy (LEc),Latent Flux Soil (LEs),Canopy Resistance (Rc),Fraction of Green Vegetation (Fg)", ",") 'Added by JBB
        Dim OutputImageListE = Split("Leaf Area Index (LAI),Soil Adjusted Vegetation Index (SAVI),Net Radiation (Rn),Aerodynamic Resistance to Heat Transfer (rah),Soil Heat Flux (G),Latent Flux (LE),Normalized Difference Vegetation Index (NDVI),Optimized Soil Adjusted Vegetation Index (OSAVI),Albedo (α),Aerodynamic Temperature (To),Sensible Heat Flux (H),Evapotranspiration (ET),Fraction of Cover (fc),Priestly Taylor Coefficient (αPT),Obukhov Length (Lo),Canopy Height (Hc),Wind Speed (U),Latent Flux Canopy (LEc),Latent Flux Soil (LEs),Canopy Resistance (Rc),Fraction of Green Leaf Area (Fg),Net Radiation Canopy(Rnc),Net Radiation Soil (Rns),Temperature Canopy (Tc),Temperature Soil (Ts),Temperature Canopy Air Space (Tac),Friction Velocity (Ustar)", ",") 'Added by JBB
        Array.Sort(OutputImageListE)
        For Each Item In OutputImageListE
            OutputImagesBoxEnergy.Items.Add(Item)
        Next
        'Dim OutputImageListW = Split("Basal Cover Coefficient (Kb),Evaporation Coefficient (Ke),Water Stress Coefficient (Ks),Cover Evapotranspiration (ETc),Water Stressed Cover Evapotranspiration (ETcAdjusted)", ",")'Commented out by JBB
        '*************Added by JBB To Include More Options ******************************
        'Dim OutputImageListW = Split("Basal Cover Coefficient (Kb),Evaporation Coefficient (Ke),Water Stress Coefficient (Ks),Cover Evapotranspiration (ETc),Water Stressed Cover Evapotranspiration (ETcAdjusted),Soil Water Depletion (Dr),Deep Percolation (DP),Total Available Water (TAW),Fraction of Cover WB (Fc),Required Gross Irrigation Depth (Ireq),Available Water Above MAD (AW),EFC DOY (DOYefc),Initiation DOY (DOYini),Termination DOY (DOYterm),Water Content in Lower Layer (ThetavL),Water Content in Whole Profile (ThetavProfile),xTotal Water Stress Adjusted ETc (SumETcAdj),xTotal Precipitation (SumP),xTotal Runoff (SumRO),xTotal Net Irrigation (SumInet),yRoot Zone Depth (Zri)", ",")
        Dim OutputImageListW = Split("Basal Crop Coefficient (Kb),Evaporation Coefficient (Ke),Water Stress Coefficient (Ks),Unstressed Cover Evapotranspiration (ETc),Water Stressed Cover Evapotranspiration (ETcAdjusted),Excess Infiltration (Fexc),Soil Water Depletion (Dr),System Deep Percolation (DP),Root Zone Deep Percolation (DPr),Total Available Water (TAW),Fraction of Cover WB (Fc),Required Gross Irrigation Depth (Ireq),Available Water Above MAD (AW),EFC DOY (DOYefc),Initiation DOY (DOYini),Termination DOY (DOYterm),Water Content in Lower Layer (ThetavL),Water Content in Whole Profile (ThetavProfile),Total Water Stress Adjusted ETc (SumETcAdj),Total Precipitation (SumP),Total Runoff (SumRO),Total Net Irrigation (SumInet),Root Zone Depth (Zri),Total Root Zone Deep Percolation (SumDPr),Total System Deep Percolation (SumDP),Total Excess Infiltration (SumFexc),Runoff (RO),Integrated Stressed Basal Crop Coefficient (SumKtKsKcb),Depth of Skin Evaporation (DREW),Depth Evaporated (De),Net Irrigation (Inet)", ",")
        '******************End Added by JBB
        Array.Sort(OutputImageListW)
        For Each Item In OutputImageListW
            OutputImagesBoxWater.Items.Add(Item)
        Next

        For Each TimeZone In TimeZoneInfo.GetSystemTimeZones
            TimeZoneBox.Items.Add(TimeZone.Id)
        Next
        DataGridViewColumns.AddNormal(CalculationCoordinatesGrid, "X", DataGridViewColumns.DataTypes.typeDouble)
        DataGridViewColumns.AddNormal(CalculationCoordinatesGrid, "Y", DataGridViewColumns.DataTypes.typeDouble)
        DataGridViewColumns.AddNormal(CalculationCoordinatesGrid, "Projection", DataGridViewColumns.DataTypes.typeString)
        For I = 0 To CalculationCoordinatesGrid.Columns.Count - 1
            CalculationCoordinatesGrid.Columns(I).ReadOnly = True
            CalculationCoordinatesGrid.Columns(I).SortMode = Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Next
        'CalculationCoordinatesGrid.Columns(2).Width = 182'Commented out by JBB
        CalculationCoordinatesGrid.Columns(0).Width = 150 'added by JBB copied from above
        CalculationCoordinatesGrid.Columns(1).Width = 150 'added by JBB copied from above
        CalculationCoordinatesGrid.Columns(2).Width = 500 'added by JBB copied from above

        ProjectSaved = False
        ProjectSavePath = ""

        NewSaveString = CreateSaveString()
    End Sub

    Private Sub OpenProject_Click(sender As System.Object, e As System.EventArgs) Handles OpenProject.Click
        If Not ExitWithoutSaving() Then Exit Sub

        Dim OpenFileDialog As New Windows.Forms.OpenFileDialog
        OpenFileDialog.Filter = "SETMI Files | *.setmi"
        OpenFileDialog.RestoreDirectory = True
        OpenFileDialog.Title = "Open SETMI file location"
        OpenFileDialog.Multiselect = False

        If Not OpenFileDialog.ShowDialog = Windows.Forms.DialogResult.OK Then Exit Sub

        NewProject_Click(Nothing, Nothing)
        ProjectSavePath = OpenFileDialog.FileName
        ProjectSaved = True

        Dim Lines = ReadSaveString().Split(New String() {Environment.NewLine}, StringSplitOptions.None)
        Dim ErrorCount As Integer = 0
        For I = 0 To Lines.Length - 2
            Try
                Dim Line = Lines(I).Split("★")
                Dim Control = Me.Controls.Find(Line(1), True)(0)

                Select Case Line(0)
                    Case "TextBox"
                        Dim C = CType(Control, Windows.Forms.TextBox)
                        C.Text = Line(2)
                    Case "ComboBox"
                        Dim C = CType(Control, Windows.Forms.ComboBox)
                        For J = 3 To Line.Length - 1
                            If Not C.Items.Contains(Line(J)) Then C.Items.Add(Line(J))
                        Next
                        C.Text = Line(2)
                    Case "ListBox"
                        Dim C = CType(Control, Windows.Forms.ListBox)
                        For J = 2 To Line.Length - 1
                            If Not C.Items.Contains(Line(J)) Then C.Items.Add(Line(J))
                        Next
                    Case "CheckedListBox"
                        Dim C = CType(Control, Windows.Forms.CheckedListBox)
                        For J = 2 To Line.Length - 1
                            Dim Item = Line(J).Split("✯")
                            Dim Index = C.Items.IndexOf(Item(0))
                            If Index > -1 Then C.SetItemCheckState(Index, Item(1))
                        Next
                    Case "DataGridView"
                        Dim C = CType(Control, Windows.Forms.DataGridView)
                        Dim Widths = Line(2).Split("✯")
                        For J = 0 To C.ColumnCount - 2
                            C.Columns(J).Width = Widths(J)
                        Next
                        Dim Rows = Line(3).Split("✯")
                        Dim CI = IIf(Line(1) = "CalculationCoordinatesGrid", 0, 1)
                        For J = 0 To Rows.Length - 2
                            Dim Col = Rows(J).Split("☆")
                            For L = CI To Col.Length - 2
                                C(L, J).Value = Col(L)
                            Next
                        Next
                End Select
            Catch
                ErrorCount += 1
            End Try
        Next
        If ErrorCount > 0 Then MsgBox("There were some problems reloading data (" & ErrorCount & ").")
    End Sub

    Private Sub SaveProject_Click(sender As System.Object, e As System.EventArgs) Handles SaveProject.Click
        If Not IO.File.Exists(ProjectSavePath) Then
            SaveAsProject_Click(Nothing, Nothing)
            Exit Sub
        End If

        Using File As New IO.Compression.GZipStream(New IO.FileStream(ProjectSavePath, IO.FileMode.Create, IO.FileAccess.ReadWrite, IO.FileShare.None), IO.Compression.CompressionMode.Compress)
            Using Writer As New IO.StreamWriter(File, System.Text.Encoding.Unicode)
                Writer.Write(CreateSaveString)
            End Using
        End Using

        ProjectSaved = True
    End Sub

    Private Sub SaveAsProject_Click(sender As System.Object, e As System.EventArgs) Handles SaveAsProject.Click
        Dim SaveFileDialog As New Windows.Forms.SaveFileDialog
        SaveFileDialog.Filter = "SETMI Files | *.setmi"
        SaveFileDialog.RestoreDirectory = True
        SaveFileDialog.Title = "Save SETMI file location as"

        If Not SaveFileDialog.ShowDialog = Windows.Forms.DialogResult.OK Then Exit Sub
        ProjectSavePath = SaveFileDialog.FileName

        IO.File.Create(ProjectSavePath).Dispose()

        SaveProject_Click(Nothing, Nothing)
    End Sub

    Private Sub ExitProject_Click(sender As System.Object, e As System.EventArgs) Handles ExitProject.Click
        Me.Close()
    End Sub

    Private Sub GetInnerControls(ByVal ChildControls As System.Windows.Forms.Control.ControlCollection, List As List(Of Windows.Forms.Control))
        For Each Control As Windows.Forms.Control In ChildControls
            List.Add(Control)
            If Control.HasChildren Then GetInnerControls(Control.Controls, List)
        Next
    End Sub

    Private Sub ImageSourceBox_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ImageSourceBox.SelectedIndexChanged
        Dim ImageSource As ImageSource = DirectCast([Enum].Parse(GetType(ImageSource), ImageSourceBox.SelectedItem.ToString.Replace(" ", "_")), ImageSource)

        Dim Enabled = (ImageSource = Functions.ImageSource.MODIS)
        ZenithList.Enabled = Enabled
        ZenithAdd.Enabled = Enabled
        ZenithRemove.Enabled = Enabled
        ZenithList.BackColor = If(Enabled, Drawing.Color.White, Drawing.Color.LightGray)
    End Sub

    Private Sub WaterBalanceBox_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles WaterBalanceBox.SelectedIndexChanged
        Dim WaterBalance As WaterBalance = DirectCast([Enum].Parse(GetType(WaterBalance), WaterBalanceBox.SelectedItem.ToString.Replace(" ", "_")), WaterBalance)
        '***This toggles the view in the water balance box based on the selected point or grid model JBB
        Select Case WaterBalance
            Case Functions.WaterBalance.Crop_Coefficient_Grid
                Label23.Text = "        Output Images"
                OutputImagesBoxWater.Visible = True
                OutputImagesCheckAllWater.Visible = True
                OutputImagesUncheckAllWater.Visible = True
                CalculationCoordinatesGrid.Visible = False
                CalculationCoordinatesAdd.Visible = False
                CalculationCoordinatesRemove.Visible = False
            Case Functions.WaterBalance.Crop_Coefficient_Point
                Label23.Text = "        Calculation Coordinates"
                OutputImagesBoxWater.Visible = False
                OutputImagesCheckAllWater.Visible = False
                OutputImagesUncheckAllWater.Visible = False
                CalculationCoordinatesGrid.Visible = True
                CalculationCoordinatesAdd.Visible = True
                CalculationCoordinatesRemove.Visible = True
        End Select
    End Sub

    Private Function GetSortedListOfControls(Control As Windows.Forms.Control) As List(Of Windows.Forms.Control)
        Dim List As New List(Of Windows.Forms.Control)
        GetInnerControls(Control.Controls, List)

        Dim ControlName = {"System.Windows.Forms.TextBox", "System.Windows.Forms.ListBox", "System.Windows.Forms.ComboBox", "System.Windows.Forms.CheckedListBox", "System.Windows.Forms.DataGridView"}

        Dim SortedList As New List(Of Windows.Forms.Control)
        For I = 0 To ControlName.Length - 1
            SortedList.AddRange(List.Where(Function(C) C.GetType().FullName = ControlName(I)))
        Next

        Return SortedList
    End Function

    Private Function CreateSaveString()
        CalculationTextEnergy.Clear()
        CalculationTextWater.Clear()

        Dim SB As New System.Text.StringBuilder

        Dim List = GetSortedListOfControls(Me)
        For Each Control In List
            Select Case Control.GetType().FullName
                Case "System.Windows.Forms.TextBox"
                    Dim C = CType(Control, Windows.Forms.TextBox)
                    SB.AppendLine("TextBox★" & C.Name & "★" & C.Text)
                Case "System.Windows.Forms.ComboBox"
                    Dim C = CType(Control, Windows.Forms.ComboBox)
                    SB.Append("ComboBox★" & C.Name & "★" & C.Text)
                    For Each Item In C.Items
                        SB.Append("★" & Item)
                    Next
                    SB.AppendLine()
                Case "System.Windows.Forms.ListBox"
                    Dim C = CType(Control, Windows.Forms.ListBox)
                    SB.Append("ListBox★" & C.Name)
                    For Each Item In C.Items
                        SB.Append("★" & Item)
                    Next
                    SB.AppendLine()
                Case "System.Windows.Forms.CheckedListBox"
                    Dim C = CType(Control, Windows.Forms.CheckedListBox)
                    SB.Append("CheckedListBox★" & C.Name)
                    For I = 0 To C.Items.Count - 1
                        SB.Append("★" & C.Items(I) & "✯" & C.GetItemCheckState(I))
                    Next
                    SB.AppendLine()
                Case "System.Windows.Forms.DataGridView"
                    Dim C = CType(Control, Windows.Forms.DataGridView)
                    SB.Append("DataGridView★" & C.Name & "★")
                    For Each Column As Windows.Forms.DataGridViewColumn In C.Columns
                        SB.Append(Column.Width & "✯")
                    Next
                    SB.Append("★")
                    For Each Row As Windows.Forms.DataGridViewRow In C.Rows
                        For Col = 0 To C.ColumnCount - 1
                            SB.Append(Row.Cells(Col).Value & "☆")
                        Next
                        SB.Append("✯")
                    Next
                    SB.AppendLine()
            End Select
        Next
        'SB = SB.Replace("✯★", "★").Replace("☆★", "★")

        Return SB.ToString
    End Function

    Private Function ReadSaveString() As String
        Dim SaveString As String = ""

        If IO.File.Exists(ProjectSavePath) Then
            Using File As New IO.Compression.GZipStream(New IO.FileStream(ProjectSavePath, IO.FileMode.Open, IO.FileAccess.ReadWrite, IO.FileShare.None), IO.Compression.CompressionMode.Decompress)
                Using Reader As New IO.StreamReader(File, System.Text.Encoding.Unicode)
                    SaveString = Reader.ReadToEnd
                End Using
            End Using
        End If

        Return SaveString
    End Function

    Private Function ExitWithoutSaving()
        Dim SaveString = CreateSaveString()
        If SaveString = ReadSaveString() Then
            Return True
        ElseIf NewSaveString = "" Then
            Return True
        ElseIf SaveString = NewSaveString Then
            Return True
        Else
            Select Case MsgBox("Would you like to save your changes?", MsgBoxStyle.YesNoCancel, "Save Changes")
                Case MsgBoxResult.Yes
                    SaveProject_Click(Nothing, Nothing)
                    Return True
                Case MsgBoxResult.No
                    Return True
                Case MsgBoxResult.Cancel
                    Return False
                Case Else
                    Return False
            End Select
        End If
    End Function

    Private Sub Main_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If Not ExitWithoutSaving() Then Exit Sub
    End Sub

#End Region

#Region "Surface"

    Private Sub MultispectralAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MultispectralAdd.Click
        Dim GetBands As Boolean = IIf(MultispectralList.Items.Count = 0, True, False)

        Dim OpenFileDialog As ESRI.ArcGIS.CatalogUI.IGxDialog = New ESRI.ArcGIS.CatalogUI.GxDialog
        OpenFileDialog.Title = "Load multispectral image(s)"
        OpenFileDialog.AllowMultiSelect = True
        Dim Filter As ESRI.ArcGIS.Catalog.IGxObjectFilter = New ESRI.ArcGIS.Catalog.GxFilterRasterDatasets
        OpenFileDialog.ObjectFilter = Filter
        Dim List As ESRI.ArcGIS.Catalog.IEnumGxObject = Nothing
        Dim GxObject As ESRI.ArcGIS.Catalog.IGxDataset

        If Not OpenFileDialog.DoModalOpen(Me.Handle, List) Then Exit Sub

        GxObject = List.Next
        Do
            Dim Path As String = GxObject.Dataset.Workspace.PathName & GxObject.Dataset.Name
            If Not MultispectralList.Items.Contains(Path) Then MultispectralList.Items.Add(Path)
            GxObject = List.Next
        Loop Until GxObject Is Nothing

        If GetBands = True Then
            Dim RasterDataset = CreateRasterDataset(MultispectralList.Items(0))
            Dim RasterBandCollection As ESRI.ArcGIS.DataSourcesRaster.IRasterBandCollection = RasterDataset
            For Item = 1 To RasterBandCollection.Count
                RedIndex.Items.Add(Item)
                GreenIndex.Items.Add(Item)
                BlueIndex.Items.Add(Item)
                NIRIndex.Items.Add(Item)
                MIR1Index.Items.Add(Item)
                MIR2Index.Items.Add(Item)
            Next
            System.Runtime.InteropServices.Marshal.ReleaseComObject(RasterDataset)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(RasterBandCollection)
        End If

        System.Runtime.InteropServices.Marshal.ReleaseComObject(OpenFileDialog)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(Filter)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(List)
    End Sub

    Private Sub MultispectralRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MultispectralRemove.Click
        Dim Indices = MultispectralList.SelectedIndices
        For Index = Indices.Count - 1 To 0 Step -1
            MultispectralList.Items.RemoveAt(Indices(Index))
        Next
        If MultispectralList.Items.Count = 0 Then
            RedIndex.Items.Clear()
            GreenIndex.Items.Clear()
            MIR1Index.Items.Clear()
            RedIndex.Text = ""
            GreenIndex.Text = ""
            MIR1Index.Text = ""
        End If
    End Sub

    Private Sub SurfaceTemperatureAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SurfaceTemperatureAdd.Click
        AddRastersIntoListBox(SurfaceTemperatureList, "surface temperature")
    End Sub

    Private Sub SurfaceTemperatureRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SurfaceTemperatureRemove.Click
        RemoveRastersFromListBox(SurfaceTemperatureList)
    End Sub

    Private Sub VegetationHeightAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VegetationHeightAdd.Click
        AddRastersIntoListBox(VegetationHeightList, "vegetatation height")
    End Sub

    Private Sub VegetationHeightRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VegetationHeightRemove.Click
        RemoveRastersFromListBox(VegetationHeightList)
    End Sub

    Private Sub LAIAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LAIAdd.Click
        AddRastersIntoListBox(LAIList, "leaf area index")
    End Sub

    Private Sub LAIRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LAIRemove.Click
        RemoveRastersFromListBox(LAIList)
    End Sub

    Private Sub ZenithAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ZenithAdd.Click
        AddRastersIntoListBox(ZenithList, "zenith angle")
    End Sub

    Private Sub ZenithRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ZenithRemove.Click
        RemoveRastersFromListBox(ZenithList)
    End Sub
    '******Added by JBB
    Private Sub FcAdd_Click(sender As Object, e As EventArgs) Handles FcAdd.Click
        AddRastersIntoListBox(FcList, "fraction of cover")
    End Sub
    Private Sub FcRemove_Click(sender As Object, e As EventArgs) Handles FcRemove.Click
        RemoveRastersFromListBox(FcList)
    End Sub

    '******End Added by JBB

    Private Sub FieldCapacityAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FieldCapacityAdd.Click
        Dim OpenFileDialog As ESRI.ArcGIS.CatalogUI.IGxDialog = New ESRI.ArcGIS.CatalogUI.GxDialog
        OpenFileDialog.Title = "Load upper layer field capacity image"
        OpenFileDialog.AllowMultiSelect = False
        Dim Filter As ESRI.ArcGIS.Catalog.IGxObjectFilter = New ESRI.ArcGIS.Catalog.GxFilterRasterDatasets
        OpenFileDialog.ObjectFilter = Filter
        Dim List As ESRI.ArcGIS.Catalog.IEnumGxObject = Nothing

        If Not OpenFileDialog.DoModalOpen(Me.Handle, List) Then Exit Sub

        Dim FileInfo As ESRI.ArcGIS.Catalog.IGxObject = List.Next
        FieldCapacityText.Text = FileInfo.FullName

        System.Runtime.InteropServices.Marshal.ReleaseComObject(OpenFileDialog)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(Filter)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(List)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(FileInfo)
    End Sub

    Private Sub WiltingPointAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WiltingPointAdd.Click
        Dim OpenFileDialog As ESRI.ArcGIS.CatalogUI.IGxDialog = New ESRI.ArcGIS.CatalogUI.GxDialog
        OpenFileDialog.Title = "Load upper layer wilting point image"
        OpenFileDialog.AllowMultiSelect = False
        Dim Filter As ESRI.ArcGIS.Catalog.IGxObjectFilter = New ESRI.ArcGIS.Catalog.GxFilterRasterDatasets
        OpenFileDialog.ObjectFilter = Filter
        Dim List As ESRI.ArcGIS.Catalog.IEnumGxObject = Nothing

        If Not OpenFileDialog.DoModalOpen(Me.Handle, List) Then Exit Sub

        Dim FileInfo As ESRI.ArcGIS.Catalog.IGxObject = List.Next
        WiltingPointText.Text = FileInfo.FullName

        System.Runtime.InteropServices.Marshal.ReleaseComObject(OpenFileDialog)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(Filter)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(List)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(FileInfo)
    End Sub

    '***added by JBB
    Private Sub InitialSoilMoistureAdd_Click(sender As Object, e As EventArgs) Handles InitialSoilMoistureAdd.Click
        Dim OpenFileDialog As ESRI.ArcGIS.CatalogUI.IGxDialog = New ESRI.ArcGIS.CatalogUI.GxDialog
        OpenFileDialog.Title = "Load upper layer initial volumetric soil moisture image"
        OpenFileDialog.AllowMultiSelect = False
        Dim Filter As ESRI.ArcGIS.Catalog.IGxObjectFilter = New ESRI.ArcGIS.Catalog.GxFilterRasterDatasets
        OpenFileDialog.ObjectFilter = Filter
        Dim List As ESRI.ArcGIS.Catalog.IEnumGxObject = Nothing

        If Not OpenFileDialog.DoModalOpen(Me.Handle, List) Then Exit Sub

        Dim FileInfo As ESRI.ArcGIS.Catalog.IGxObject = List.Next
        InitialSoilMoistureText.Text = FileInfo.FullName

        System.Runtime.InteropServices.Marshal.ReleaseComObject(OpenFileDialog)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(Filter)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(List)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(FileInfo)
    End Sub

    Private Sub FieldCapacityAdd2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FieldCapacityAdd2.Click
        Dim OpenFileDialog As ESRI.ArcGIS.CatalogUI.IGxDialog = New ESRI.ArcGIS.CatalogUI.GxDialog
        OpenFileDialog.Title = "Load second layer field capacity image"
        OpenFileDialog.AllowMultiSelect = False
        Dim Filter As ESRI.ArcGIS.Catalog.IGxObjectFilter = New ESRI.ArcGIS.Catalog.GxFilterRasterDatasets
        OpenFileDialog.ObjectFilter = Filter
        Dim List As ESRI.ArcGIS.Catalog.IEnumGxObject = Nothing

        If Not OpenFileDialog.DoModalOpen(Me.Handle, List) Then Exit Sub

        Dim FileInfo As ESRI.ArcGIS.Catalog.IGxObject = List.Next
        FieldCapacityText2.Text = FileInfo.FullName

        System.Runtime.InteropServices.Marshal.ReleaseComObject(OpenFileDialog)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(Filter)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(List)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(FileInfo)
    End Sub

    Private Sub FieldCapacityAdd3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FieldCapacityAdd3.Click
        Dim OpenFileDialog As ESRI.ArcGIS.CatalogUI.IGxDialog = New ESRI.ArcGIS.CatalogUI.GxDialog
        OpenFileDialog.Title = "Load third layer field capacity image"
        OpenFileDialog.AllowMultiSelect = False
        Dim Filter As ESRI.ArcGIS.Catalog.IGxObjectFilter = New ESRI.ArcGIS.Catalog.GxFilterRasterDatasets
        OpenFileDialog.ObjectFilter = Filter
        Dim List As ESRI.ArcGIS.Catalog.IEnumGxObject = Nothing

        If Not OpenFileDialog.DoModalOpen(Me.Handle, List) Then Exit Sub

        Dim FileInfo As ESRI.ArcGIS.Catalog.IGxObject = List.Next
        FieldCapacityText3.Text = FileInfo.FullName

        System.Runtime.InteropServices.Marshal.ReleaseComObject(OpenFileDialog)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(Filter)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(List)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(FileInfo)
    End Sub

    Private Sub WiltingPointAdd2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WiltingPointAdd2.Click
        Dim OpenFileDialog As ESRI.ArcGIS.CatalogUI.IGxDialog = New ESRI.ArcGIS.CatalogUI.GxDialog
        OpenFileDialog.Title = "Load second layer wilting point image"
        OpenFileDialog.AllowMultiSelect = False
        Dim Filter As ESRI.ArcGIS.Catalog.IGxObjectFilter = New ESRI.ArcGIS.Catalog.GxFilterRasterDatasets
        OpenFileDialog.ObjectFilter = Filter
        Dim List As ESRI.ArcGIS.Catalog.IEnumGxObject = Nothing

        If Not OpenFileDialog.DoModalOpen(Me.Handle, List) Then Exit Sub

        Dim FileInfo As ESRI.ArcGIS.Catalog.IGxObject = List.Next
        WiltingPointText2.Text = FileInfo.FullName

        System.Runtime.InteropServices.Marshal.ReleaseComObject(OpenFileDialog)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(Filter)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(List)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(FileInfo)
    End Sub

    Private Sub WiltingPointAdd3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WiltingPointAdd3.Click
        Dim OpenFileDialog As ESRI.ArcGIS.CatalogUI.IGxDialog = New ESRI.ArcGIS.CatalogUI.GxDialog
        OpenFileDialog.Title = "Load third layer wilting point image"
        OpenFileDialog.AllowMultiSelect = False
        Dim Filter As ESRI.ArcGIS.Catalog.IGxObjectFilter = New ESRI.ArcGIS.Catalog.GxFilterRasterDatasets
        OpenFileDialog.ObjectFilter = Filter
        Dim List As ESRI.ArcGIS.Catalog.IEnumGxObject = Nothing

        If Not OpenFileDialog.DoModalOpen(Me.Handle, List) Then Exit Sub

        Dim FileInfo As ESRI.ArcGIS.Catalog.IGxObject = List.Next
        WiltingPointText3.Text = FileInfo.FullName

        System.Runtime.InteropServices.Marshal.ReleaseComObject(OpenFileDialog)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(Filter)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(List)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(FileInfo)
    End Sub

    Private Sub InitialSoilMoistureAdd2_Click(sender As Object, e As EventArgs) Handles InitialSoilMoistureAdd2.Click
        Dim OpenFileDialog As ESRI.ArcGIS.CatalogUI.IGxDialog = New ESRI.ArcGIS.CatalogUI.GxDialog
        OpenFileDialog.Title = "Load second layer initial volumetric soil moisture image"
        OpenFileDialog.AllowMultiSelect = False
        Dim Filter As ESRI.ArcGIS.Catalog.IGxObjectFilter = New ESRI.ArcGIS.Catalog.GxFilterRasterDatasets
        OpenFileDialog.ObjectFilter = Filter
        Dim List As ESRI.ArcGIS.Catalog.IEnumGxObject = Nothing

        If Not OpenFileDialog.DoModalOpen(Me.Handle, List) Then Exit Sub

        Dim FileInfo As ESRI.ArcGIS.Catalog.IGxObject = List.Next
        InitialSoilMoistureText2.Text = FileInfo.FullName

        System.Runtime.InteropServices.Marshal.ReleaseComObject(OpenFileDialog)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(Filter)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(List)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(FileInfo)
    End Sub

    Private Sub InitialSoilMoistureAdd3_Click(sender As Object, e As EventArgs) Handles InitialSoilMoistureAdd3.Click
        Dim OpenFileDialog As ESRI.ArcGIS.CatalogUI.IGxDialog = New ESRI.ArcGIS.CatalogUI.GxDialog
        OpenFileDialog.Title = "Load third layer initial volumetric soil moisture image"
        OpenFileDialog.AllowMultiSelect = False
        Dim Filter As ESRI.ArcGIS.Catalog.IGxObjectFilter = New ESRI.ArcGIS.Catalog.GxFilterRasterDatasets
        OpenFileDialog.ObjectFilter = Filter
        Dim List As ESRI.ArcGIS.Catalog.IEnumGxObject = Nothing

        If Not OpenFileDialog.DoModalOpen(Me.Handle, List) Then Exit Sub

        Dim FileInfo As ESRI.ArcGIS.Catalog.IGxObject = List.Next
        InitialSoilMoistureText3.Text = FileInfo.FullName

        System.Runtime.InteropServices.Marshal.ReleaseComObject(OpenFileDialog)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(Filter)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(List)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(FileInfo)
    End Sub

    Private Sub LayerThicknessAdd_Click(sender As Object, e As EventArgs) Handles LayerThicknessAdd.Click
        Dim OpenFileDialog As ESRI.ArcGIS.CatalogUI.IGxDialog = New ESRI.ArcGIS.CatalogUI.GxDialog
        OpenFileDialog.Title = "Load upper layer thickness image"
        OpenFileDialog.AllowMultiSelect = False
        Dim Filter As ESRI.ArcGIS.Catalog.IGxObjectFilter = New ESRI.ArcGIS.Catalog.GxFilterRasterDatasets
        OpenFileDialog.ObjectFilter = Filter
        Dim List As ESRI.ArcGIS.Catalog.IEnumGxObject = Nothing

        If Not OpenFileDialog.DoModalOpen(Me.Handle, List) Then Exit Sub

        Dim FileInfo As ESRI.ArcGIS.Catalog.IGxObject = List.Next
        LayerThicknessText.Text = FileInfo.FullName

        System.Runtime.InteropServices.Marshal.ReleaseComObject(OpenFileDialog)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(Filter)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(List)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(FileInfo)
    End Sub
    Private Sub LayerThicknessAdd2_Click(sender As Object, e As EventArgs) Handles LayerThicknessAdd2.Click
        Dim OpenFileDialog As ESRI.ArcGIS.CatalogUI.IGxDialog = New ESRI.ArcGIS.CatalogUI.GxDialog
        OpenFileDialog.Title = "Load second layer thickness image"
        OpenFileDialog.AllowMultiSelect = False
        Dim Filter As ESRI.ArcGIS.Catalog.IGxObjectFilter = New ESRI.ArcGIS.Catalog.GxFilterRasterDatasets
        OpenFileDialog.ObjectFilter = Filter
        Dim List As ESRI.ArcGIS.Catalog.IEnumGxObject = Nothing

        If Not OpenFileDialog.DoModalOpen(Me.Handle, List) Then Exit Sub

        Dim FileInfo As ESRI.ArcGIS.Catalog.IGxObject = List.Next
        LayerThicknessText2.Text = FileInfo.FullName

        System.Runtime.InteropServices.Marshal.ReleaseComObject(OpenFileDialog)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(Filter)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(List)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(FileInfo)
    End Sub
    Private Sub LayerThicknessAdd3_Click(sender As Object, e As EventArgs) Handles LayerThicknessAdd3.Click
        Dim OpenFileDialog As ESRI.ArcGIS.CatalogUI.IGxDialog = New ESRI.ArcGIS.CatalogUI.GxDialog
        OpenFileDialog.Title = "Load third layer thickness image"
        OpenFileDialog.AllowMultiSelect = False
        Dim Filter As ESRI.ArcGIS.Catalog.IGxObjectFilter = New ESRI.ArcGIS.Catalog.GxFilterRasterDatasets
        OpenFileDialog.ObjectFilter = Filter
        Dim List As ESRI.ArcGIS.Catalog.IEnumGxObject = Nothing

        If Not OpenFileDialog.DoModalOpen(Me.Handle, List) Then Exit Sub

        Dim FileInfo As ESRI.ArcGIS.Catalog.IGxObject = List.Next
        LayerThicknessText3.Text = FileInfo.FullName

        System.Runtime.InteropServices.Marshal.ReleaseComObject(OpenFileDialog)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(Filter)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(List)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(FileInfo)
    End Sub

    Private Sub HydraulicConductivityAdd_Click(sender As Object, e As EventArgs) Handles HydraulicConductivityAdd.Click
        Dim OpenFileDialog As ESRI.ArcGIS.CatalogUI.IGxDialog = New ESRI.ArcGIS.CatalogUI.GxDialog
        OpenFileDialog.Title = "Load upper layer saturated hydraulic conductivity image"
        OpenFileDialog.AllowMultiSelect = False
        Dim Filter As ESRI.ArcGIS.Catalog.IGxObjectFilter = New ESRI.ArcGIS.Catalog.GxFilterRasterDatasets
        OpenFileDialog.ObjectFilter = Filter
        Dim List As ESRI.ArcGIS.Catalog.IEnumGxObject = Nothing

        If Not OpenFileDialog.DoModalOpen(Me.Handle, List) Then Exit Sub

        Dim FileInfo As ESRI.ArcGIS.Catalog.IGxObject = List.Next
        HydraulicConductivityText.Text = FileInfo.FullName

        System.Runtime.InteropServices.Marshal.ReleaseComObject(OpenFileDialog)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(Filter)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(List)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(FileInfo)
    End Sub

    Private Sub HydraulicConductivityAdd2_Click(sender As Object, e As EventArgs) Handles HydraulicConductivityAdd2.Click
        Dim OpenFileDialog As ESRI.ArcGIS.CatalogUI.IGxDialog = New ESRI.ArcGIS.CatalogUI.GxDialog
        OpenFileDialog.Title = "Load second layer saturated hydraulic conductivity image"
        OpenFileDialog.AllowMultiSelect = False
        Dim Filter As ESRI.ArcGIS.Catalog.IGxObjectFilter = New ESRI.ArcGIS.Catalog.GxFilterRasterDatasets
        OpenFileDialog.ObjectFilter = Filter
        Dim List As ESRI.ArcGIS.Catalog.IEnumGxObject = Nothing

        If Not OpenFileDialog.DoModalOpen(Me.Handle, List) Then Exit Sub

        Dim FileInfo As ESRI.ArcGIS.Catalog.IGxObject = List.Next
        HydraulicConductivityText2.Text = FileInfo.FullName

        System.Runtime.InteropServices.Marshal.ReleaseComObject(OpenFileDialog)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(Filter)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(List)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(FileInfo)
    End Sub

    Private Sub HydraulicConductivityAdd3_Click(sender As Object, e As EventArgs) Handles HydraulicConductivityAdd3.Click
        Dim OpenFileDialog As ESRI.ArcGIS.CatalogUI.IGxDialog = New ESRI.ArcGIS.CatalogUI.GxDialog
        OpenFileDialog.Title = "Load third layer saturated hydraulic conductivity image"
        OpenFileDialog.AllowMultiSelect = False
        Dim Filter As ESRI.ArcGIS.Catalog.IGxObjectFilter = New ESRI.ArcGIS.Catalog.GxFilterRasterDatasets
        OpenFileDialog.ObjectFilter = Filter
        Dim List As ESRI.ArcGIS.Catalog.IEnumGxObject = Nothing

        If Not OpenFileDialog.DoModalOpen(Me.Handle, List) Then Exit Sub

        Dim FileInfo As ESRI.ArcGIS.Catalog.IGxObject = List.Next
        HydraulicConductivityText3.Text = FileInfo.FullName

        System.Runtime.InteropServices.Marshal.ReleaseComObject(OpenFileDialog)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(Filter)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(List)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(FileInfo)
    End Sub

    Private Sub SaturatedWaterContentAdd_Click(sender As Object, e As EventArgs) Handles SaturatedWaterContentAdd.Click
        Dim OpenFileDialog As ESRI.ArcGIS.CatalogUI.IGxDialog = New ESRI.ArcGIS.CatalogUI.GxDialog
        OpenFileDialog.Title = "Load upper layer water content at saturation image"
        OpenFileDialog.AllowMultiSelect = False
        Dim Filter As ESRI.ArcGIS.Catalog.IGxObjectFilter = New ESRI.ArcGIS.Catalog.GxFilterRasterDatasets
        OpenFileDialog.ObjectFilter = Filter
        Dim List As ESRI.ArcGIS.Catalog.IEnumGxObject = Nothing

        If Not OpenFileDialog.DoModalOpen(Me.Handle, List) Then Exit Sub

        Dim FileInfo As ESRI.ArcGIS.Catalog.IGxObject = List.Next
        SaturatedWaterContentText.Text = FileInfo.FullName

        System.Runtime.InteropServices.Marshal.ReleaseComObject(OpenFileDialog)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(Filter)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(List)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(FileInfo)
    End Sub

    Private Sub SaturatedWaterContentAdd2_Click(sender As Object, e As EventArgs) Handles SaturatedWaterContentAdd2.Click
        Dim OpenFileDialog As ESRI.ArcGIS.CatalogUI.IGxDialog = New ESRI.ArcGIS.CatalogUI.GxDialog
        OpenFileDialog.Title = "Load second layer water content at saturation image"
        OpenFileDialog.AllowMultiSelect = False
        Dim Filter As ESRI.ArcGIS.Catalog.IGxObjectFilter = New ESRI.ArcGIS.Catalog.GxFilterRasterDatasets
        OpenFileDialog.ObjectFilter = Filter
        Dim List As ESRI.ArcGIS.Catalog.IEnumGxObject = Nothing

        If Not OpenFileDialog.DoModalOpen(Me.Handle, List) Then Exit Sub

        Dim FileInfo As ESRI.ArcGIS.Catalog.IGxObject = List.Next
        SaturatedWaterContentText2.Text = FileInfo.FullName

        System.Runtime.InteropServices.Marshal.ReleaseComObject(OpenFileDialog)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(Filter)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(List)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(FileInfo)
    End Sub

    Private Sub SaturatedWaterContentAdd3_Click(sender As Object, e As EventArgs) Handles SaturatedWaterContentAdd3.Click
        Dim OpenFileDialog As ESRI.ArcGIS.CatalogUI.IGxDialog = New ESRI.ArcGIS.CatalogUI.GxDialog
        OpenFileDialog.Title = "Load third layer water content at saturation image"
        OpenFileDialog.AllowMultiSelect = False
        Dim Filter As ESRI.ArcGIS.Catalog.IGxObjectFilter = New ESRI.ArcGIS.Catalog.GxFilterRasterDatasets
        OpenFileDialog.ObjectFilter = Filter
        Dim List As ESRI.ArcGIS.Catalog.IEnumGxObject = Nothing

        If Not OpenFileDialog.DoModalOpen(Me.Handle, List) Then Exit Sub

        Dim FileInfo As ESRI.ArcGIS.Catalog.IGxObject = List.Next
        SaturatedWaterContentText3.Text = FileInfo.FullName

        System.Runtime.InteropServices.Marshal.ReleaseComObject(OpenFileDialog)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(Filter)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(List)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(FileInfo)
    End Sub
    Private Sub REWAdd_Click(sender As Object, e As EventArgs) Handles REWAdd.Click 'Added by JBB copying code (as with others added by JBB) from above 2/9/2018
        Dim OpenFileDialog As ESRI.ArcGIS.CatalogUI.IGxDialog = New ESRI.ArcGIS.CatalogUI.GxDialog
        OpenFileDialog.Title = "Load readily evaporable water image"
        OpenFileDialog.AllowMultiSelect = False
        Dim Filter As ESRI.ArcGIS.Catalog.IGxObjectFilter = New ESRI.ArcGIS.Catalog.GxFilterRasterDatasets
        OpenFileDialog.ObjectFilter = Filter
        Dim List As ESRI.ArcGIS.Catalog.IEnumGxObject = Nothing

        If Not OpenFileDialog.DoModalOpen(Me.Handle, List) Then Exit Sub

        Dim FileInfo As ESRI.ArcGIS.Catalog.IGxObject = List.Next
        REWText.Text = FileInfo.FullName

        System.Runtime.InteropServices.Marshal.ReleaseComObject(OpenFileDialog)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(Filter)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(List)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(FileInfo)
    End Sub

    Private Sub DPLimitAdd_Click(sender As Object, e As EventArgs) Handles DPLimitAdd.Click 'Added by JBB copying code (as with others added by JBB) from above 2/9/2018
        Dim OpenFileDialog As ESRI.ArcGIS.CatalogUI.IGxDialog = New ESRI.ArcGIS.CatalogUI.GxDialog
        OpenFileDialog.Title = "Load constant deep percolation limit image"
        OpenFileDialog.AllowMultiSelect = False
        Dim Filter As ESRI.ArcGIS.Catalog.IGxObjectFilter = New ESRI.ArcGIS.Catalog.GxFilterRasterDatasets
        OpenFileDialog.ObjectFilter = Filter
        Dim List As ESRI.ArcGIS.Catalog.IEnumGxObject = Nothing

        If Not OpenFileDialog.DoModalOpen(Me.Handle, List) Then Exit Sub

        Dim FileInfo As ESRI.ArcGIS.Catalog.IGxObject = List.Next
        DPLimitText.Text = FileInfo.FullName

        System.Runtime.InteropServices.Marshal.ReleaseComObject(OpenFileDialog)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(Filter)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(List)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(FileInfo)
    End Sub
    '***end added by JBB
#End Region

#Region "Cover"

    Private CoverPointIndex As New CoverPointIndex
    Private WBOutputDateIndex As New WBOutputDateIndex 'Added by JBB for Water Balance Output Date Table JBB
    Private WBPointOutputIndex As New WBPointOutputIndex 'added by JBB copying the above lines 2/1/2018
    'Private CoverVariables = {"", "Cover Name", "Kcb Initial", "Kcb Mid", "Kcb End", "Period Initial (days)", "Period Development (days)", "Period Mid (days)", "Period End (days)", "Cover Maximum Height (m)", "Cover Minimum Height (m)", "Root Maximum Depth (m)", "Start Date (day of year)", "p"} ' Commented out by JBB
    'Private CoverVariables = {"", "Cover Name", "Kcb Initial", "Kcb Mid", "Kcb End", "Period Initial (days)", "Period Development (days)", "Period Mid (days)", "Period End (days)", "Cover Maximum Height (m)", "Cover Minimum Height (m)", "Root Maximum Depth (m)", "Start Date (day of year)", "p", "Curve Number", "Percent Effective Precip.", "Depth of Evaporative Layer (m)", "DOY Start Water Balance", "DOY End Water Balance", "DOY Initiation Earliest", "DOY Initiation Latest", "DOY EFC Earliest", "DOY EFC Latest", "DOY Termination Earliest", "DOY Termination Latest", "Root Minimum Depth (m)", "Kc Max", "Kcb Off Season", "Weight for Assimilation", "MAD (%)", "Target Depth Above MAD (mm)", "Irrigation App Eff (%)", "Weight for Depletion", "Weight for Evapd Depth", "Weight for Lwr Lyr ThetaV", "False End SAVI", "False End SAVI DOY", "False Peak SAVI", "False Peak SAVI Max DOY", "xWB Start Year"} 'Added by JBB
    Private CoverVariables = {"", "Cover Name", "Cover Maximum Height (m)", "Cover Minimum Height (m)", "Root Maximum Depth (m)", "p", "Curve Number", "Percent Effective Precip.", "Depth of Evaporative Layer (m)", "DOY Start Water Balance", "DOY End Water Balance", "DOY Initiation Earliest", "DOY Initiation Latest", "DOY EFC Earliest", "DOY EFC Latest", "DOY Termination Earliest", "DOY Termination Latest", "Root Minimum Depth (m)", "Kc Max", "Kcb Off Season", "Weight for Assimilation", "MAD (%)", "Target Depth Above MAD (mm)", "Irrigation App Eff (%)", "Weight for Depletion", "Weight for Evapd Depth", "Weight for Lwr Lyr ThetaV", "Weight for Skin Evaporated Depth", "False End SAVI", "False End SAVI DOY", "False Peak SAVI", "False Peak SAVI Max DOY", "WB Start Year", "False End SAVI CGDD", "False Peak SAVI Max CGDD", "GDD Base Temperature (C)", "GDD Maximum Temperature (C)", "Temp. Stress Base (C)", "Temp. Stress Upper Limit (C)", "Ratio of Kcb:Kcbmax to Initiate Late Period", "xSavi at Max Biomass"} 'Added by JBB
    Private WtrBalOutDate = {"Output Date"} 'Added by JBB for Water Balance Output Table JBB
    Private WBPntOutLocs = {"X Coordinate", "Y Coordinate"} 'Added by JBB for inputting tabular coordinates for the WB, mimmicking the above lines.
    Private CoverValues As New List(Of Integer)

    Private Sub CoverClassificationAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CoverClassificationAdd.Click
        Dim OpenFileDialog As ESRI.ArcGIS.CatalogUI.IGxDialog = New ESRI.ArcGIS.CatalogUI.GxDialog
        OpenFileDialog.Title = "Load cover classification image(s)"
        OpenFileDialog.AllowMultiSelect = True
        Dim Filter As ESRI.ArcGIS.Catalog.IGxObjectFilter = New ESRI.ArcGIS.Catalog.GxFilterRasterDatasets
        OpenFileDialog.ObjectFilter = Filter
        Dim List As ESRI.ArcGIS.Catalog.IEnumGxObject = Nothing
        Dim GxObject As ESRI.ArcGIS.Catalog.IGxDataset

        If Not OpenFileDialog.DoModalOpen(Me.Handle, List) Then Exit Sub

        GxObject = List.Next
        Do
            Dim Path As String = GxObject.Dataset.Workspace.PathName & GxObject.Dataset.Name
            If Not CoverClassificationList.Items.Contains(Path) Then CoverClassificationList.Items.Add(Path)
            If CoverClassificationList.Items.Count = 1 Then
                Dim Raster As ESRI.ArcGIS.DataSourcesRaster.IRaster2 = CType(CreateRasterDataset(Path).CreateDefaultRaster, ESRI.ArcGIS.DataSourcesRaster.IRaster2)
                For F = 0 To Raster.AttributeTable.Fields.FieldCount - 1
                    CoverSelectionBox.Items.Add(Raster.AttributeTable.Fields.Field(F).AliasName)
                    If Raster.AttributeTable.Fields.Field(F).AliasName.Contains("Name") Or Raster.AttributeTable.Fields.Field(F).AliasName.Contains("NAME") Or Raster.AttributeTable.Fields.Field(F).AliasName.Contains("name") Then CoverSelectionBox.Text = Raster.AttributeTable.Fields.Field(F).AliasName
                Next
                System.Runtime.InteropServices.Marshal.ReleaseComObject(Raster)
            End If
            GxObject = List.Next
        Loop Until GxObject Is Nothing

        System.Runtime.InteropServices.Marshal.ReleaseComObject(OpenFileDialog)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(Filter)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(List)
    End Sub

    Private Sub CoverClassificationRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CoverClassificationRemove.Click
        Dim Indices = CoverClassificationList.SelectedIndices
        For Index = Indices.Count - 1 To 0 Step -1
            CoverClassificationList.Items.RemoveAt(Indices(Index))
        Next
        If CoverClassificationList.Items.Count = 0 Then
            CoverSelectionBox.Items.Clear()
            CoverSelectionBox.Text = ""
            CoverSelectionGrid.Columns.Clear()
        End If
    End Sub

    Private Sub CoverSelectionBox_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CoverSelectionBox.TextChanged 'Populates the datatable under the cover selection box JBB
        If CoverClassificationList.Items.Count > 0 Then
            Try
                Dim CoverList As New List(Of String)
                For Item = 0 To [Enum].GetNames(GetType(Cover)).Count - 1
                    CoverList.Add([Enum].GetName(GetType(Cover), Item).Replace("_", " "))
                Next
                CoverList.Sort()
                CoverList.Insert(0, "")

                Dim IrrigationMethodList As New List(Of String)
                For Item = 0 To [Enum].GetNames(GetType(IrrigationMethod)).Count - 1
                    IrrigationMethodList.Add([Enum].GetName(GetType(IrrigationMethod), Item).Replace("_", " "))
                Next
                IrrigationMethodList.Sort()
                IrrigationMethodList.Insert(0, "")

                'Added by JBB by mimicking the code above
                Dim KcbVIList As New List(Of String)
                For Item = 0 To [Enum].GetNames(GetType(KcbVI)).Count - 1
                    KcbVIList.Add([Enum].GetName(GetType(KcbVI), Item))
                Next
                KcbVIList.Insert(0, "")

                Dim FcVIList As New List(Of String)
                For Item = 0 To [Enum].GetNames(GetType(FcVI)).Count - 1
                    FcVIList.Add([Enum].GetName(GetType(FcVI), Item))
                Next
                FcVIList.Insert(0, "")
                'End added by JBB

                CoverSelectionGrid.Columns.Clear()
                DataGridViewColumns.AddNormal(CoverSelectionGrid, "Cover", DataGridViewColumns.DataTypes.typeString, True)
                DataGridViewColumns.AddCombo(CoverSelectionGrid, "Cover Classification", CoverList.ToArray)
                DataGridViewColumns.AddCombo(CoverSelectionGrid, "Irrigation Method", IrrigationMethodList.ToArray)
                DataGridViewColumns.AddCheckbox(CoverSelectionGrid, "Include")
                DataGridViewColumns.AddNormal(CoverSelectionGrid, "α Leaf VIS", DataGridViewColumns.DataTypes.typeDouble)
                DataGridViewColumns.AddNormal(CoverSelectionGrid, "α Leaf NIR", DataGridViewColumns.DataTypes.typeDouble)
                DataGridViewColumns.AddNormal(CoverSelectionGrid, "α Leaf TIR", DataGridViewColumns.DataTypes.typeDouble)
                DataGridViewColumns.AddNormal(CoverSelectionGrid, "α Dead VIS", DataGridViewColumns.DataTypes.typeDouble)
                DataGridViewColumns.AddNormal(CoverSelectionGrid, "α Dead NIR", DataGridViewColumns.DataTypes.typeDouble)
                DataGridViewColumns.AddNormal(CoverSelectionGrid, "α Dead TIR", DataGridViewColumns.DataTypes.typeDouble)
                DataGridViewColumns.AddNormal(CoverSelectionGrid, "fg", DataGridViewColumns.DataTypes.typeDouble)
                'DataGridViewColumns.AddNormal(CoverSelectionGrid, "Hc max", DataGridViewColumns.DataTypes.typeDouble)'Commented out by JBB, because the label was in the wrong order
                DataGridViewColumns.AddNormal(CoverSelectionGrid, "Hc min", DataGridViewColumns.DataTypes.typeDouble)
                DataGridViewColumns.AddNormal(CoverSelectionGrid, "Hc max", DataGridViewColumns.DataTypes.typeDouble) 'Added by JBB because the label was in the wrong order
                DataGridViewColumns.AddNormal(CoverSelectionGrid, "s", DataGridViewColumns.DataTypes.typeDouble)
                DataGridViewColumns.AddNormal(CoverSelectionGrid, "Wc", DataGridViewColumns.DataTypes.typeDouble)
                DataGridViewColumns.AddNormal(CoverSelectionGrid, "LAI", DataGridViewColumns.DataTypes.typeDouble)
                DataGridViewColumns.AddNormal(CoverSelectionGrid, "Refl Soil VIS", DataGridViewColumns.DataTypes.typeDouble) 'Added by JBB
                DataGridViewColumns.AddNormal(CoverSelectionGrid, "Refl Soil NIR", DataGridViewColumns.DataTypes.typeDouble) 'Added by JBB
                DataGridViewColumns.AddNormal(CoverSelectionGrid, "ε Soil TIR", DataGridViewColumns.DataTypes.typeDouble) 'Added by JBB
                DataGridViewColumns.AddNormal(CoverSelectionGrid, "Ag", DataGridViewColumns.DataTypes.typeDouble) 'Added by JBB
                DataGridViewColumns.AddNormal(CoverSelectionGrid, "D", DataGridViewColumns.DataTypes.typeDouble) 'Added by JBB
                DataGridViewColumns.AddNormal(CoverSelectionGrid, "αPT Ini", DataGridViewColumns.DataTypes.typeDouble) 'Added by JBB
                DataGridViewColumns.AddNormal(CoverSelectionGrid, "Rc Ini", DataGridViewColumns.DataTypes.typeDouble) 'Added by JBB
                DataGridViewColumns.AddNormal(CoverSelectionGrid, "Rc Max", DataGridViewColumns.DataTypes.typeDouble) 'Added by JBB
                'DataGridViewColumns.AddNormal(CoverSelectionGrid, "Max Iter", DataGridViewColumns.DataTypes.typeInteger) 'Added by JBB
                DataGridViewColumns.AddCombo(CoverSelectionGrid, "Fc Veg. Index", FcVIList.ToArray) 'Added by JBB by mimicking code above
                DataGridViewColumns.AddNormal(CoverSelectionGrid, "Fc Min VI", DataGridViewColumns.DataTypes.typeDouble) 'Added by JBB
                DataGridViewColumns.AddNormal(CoverSelectionGrid, "Fc Max VI", DataGridViewColumns.DataTypes.typeDouble) 'Added by JBB
                DataGridViewColumns.AddNormal(CoverSelectionGrid, "Fc Expon.", DataGridViewColumns.DataTypes.typeDouble) 'Added by JBB
                DataGridViewColumns.AddCombo(CoverSelectionGrid, "Kcb Veg. Index", KcbVIList.ToArray) 'Added by JBB by mimicking code above
                DataGridViewColumns.AddNormal(CoverSelectionGrid, "Kcbrf Slope", DataGridViewColumns.DataTypes.typeDouble) 'Added by JBB
                DataGridViewColumns.AddNormal(CoverSelectionGrid, "Kcbrf Incpt", DataGridViewColumns.DataTypes.typeDouble) 'Added by JBB
                DataGridViewColumns.AddNormal(CoverSelectionGrid, "Kcbrf Min", DataGridViewColumns.DataTypes.typeDouble) 'Added by JBB
                DataGridViewColumns.AddNormal(CoverSelectionGrid, "Kcbrf Max", DataGridViewColumns.DataTypes.typeDouble) 'Added by JBB

                For I = 0 To 1
                    CoverSelectionGrid.Columns(I).Width = 110
                Next
                For I = 2 To CoverSelectionGrid.Columns.Count - 1
                    CoverSelectionGrid.Columns(I).Width = 55
                Next
                For I = 0 To CoverSelectionGrid.Columns.Count - 1
                    CoverSelectionGrid.Columns(I).SortMode = Windows.Forms.DataGridViewColumnSortMode.NotSortable
                Next

                Dim Raster As ESRI.ArcGIS.DataSourcesRaster.IRaster2 = CType(CreateRasterDataset(CoverClassificationList.Items(0)).CreateDefaultRaster, ESRI.ArcGIS.DataSourcesRaster.IRaster2)
                Dim QueryFilter As ESRI.ArcGIS.Geodatabase.IQueryFilter = New ESRI.ArcGIS.Geodatabase.QueryFilter
                QueryFilter.AddField(CoverSelectionBox.Text)
                Dim Cursor As ESRI.ArcGIS.Geodatabase.ICursor = Raster.AttributeTable.Search(QueryFilter, True)
                Dim Row As ESRI.ArcGIS.Geodatabase.IRow = Cursor.NextRow()
                Dim List As New List(Of Object)
                Dim Index As Integer = Raster.AttributeTable.Fields.FindField(CoverSelectionBox.Text)
                CoverValues.Clear()
                Do While Not Row Is Nothing
                    List.Add(Row.Value(Index))
                    CoverValues.Add(Row.Value(1))
                    Row = Cursor.NextRow()
                Loop
                System.Runtime.InteropServices.Marshal.ReleaseComObject(Raster)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(QueryFilter)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(Cursor)

                For L = 0 To List.Count - 1
                    Dim CoverClass As String = ""
                    For Each Item In CoverList
                        If Not CoverClass = List(L) Then
                            If Item.Contains(List(L)) Then CoverClass = Item
                            If CoverClass = "" And List(L).ToString.Contains(Item) Then CoverClass = Item
                        End If
                    Next
                    CoverSelectionGrid.Rows.Add({List(L), CoverClass, "", True})
                Next


                'CoverSelectionGrid.Rows(I).Cells(4).Value = 0.83    '   α Leaf VIS
                'CoverSelectionGrid.Rows(I).Cells(5).Value = 0.35    '   α Leaf NIR
                'CoverSelectionGrid.Rows(I).Cells(6).Value = 0.97    '   α Leaf TIR
                'CoverSelectionGrid.Rows(I).Cells(7).Value = 0.97    '   α Dead VIS
                'CoverSelectionGrid.Rows(I).Cells(8).Value = 0.97    '   α Dead NIR
                'CoverSelectionGrid.Rows(I).Cells(9).Value = 0.97    '   α Dead TIR
                'CoverSelectionGrid.Rows(I).Cells(10).Value = 1.0    '   fg
                'CoverSelectionGrid.Rows(I).Cells(11).Value = -999   '   Hc min
                'CoverSelectionGrid.Rows(I).Cells(12).Value = -999   '   Hc max
                'CoverSelectionGrid.Rows(I).Cells(13).Value = 0.05   '   s
                'CoverSelectionGrid.Rows(I).Cells(14).Value = 0      '   w
                'CoverSelectionGrid.Rows(I).Cells(15).Value = 0      '   LAI measured

                For I = 0 To CoverSelectionGrid.Rows.Count - 1
                    If CoverSelectionGrid.Rows(I).Cells(1).Value <> "" Then
                        Dim Cover As Cover = DirectCast([Enum].Parse(GetType(Cover), CoverSelectionGrid.Rows(I).Cells(1).Value.Replace(" ", "_")), Cover)
                        Dim output = CoverProperties(Cover)
                        For J = 4 To 32
                            CoverSelectionGrid.Rows(I).Cells(J).Value = output(J - 4)
                        Next
                    End If
                Next
            Catch
            End Try
        End If
    End Sub

    Private Sub CoverSelectionGrid_CellValueChanged(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles CoverSelectionGrid.CellValueChanged
        If e.ColumnIndex = 1 Then
            Dim I = e.RowIndex

            If CoverSelectionGrid.Rows(I).Cells(1).Value <> "" Then
                Dim Cover As Cover = DirectCast([Enum].Parse(GetType(Cover), CoverSelectionGrid.Rows(I).Cells(1).Value.Replace(" ", "_")), Cover)
                Dim Output = CoverProperties(Cover)
                For J = 4 To 32 'Changed by JBB
                    CoverSelectionGrid.Rows(I).Cells(J).Value = Output(J - 4)
                Next
            End If
        End If
    End Sub

    Private Sub CoverSelectionGrid_DataError(sender As Object, e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles CoverSelectionGrid.DataError
        MsgBox("The value entered was not a number. Please try again.", MsgBoxStyle.Exclamation, "Not Numeric Value")
    End Sub

    Private Sub CoverPropertiesGrid_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles CoverPropertiesGrid.CellValueChanged
        CoverPointIndex.Initialize()
        For Row = 0 To CoverPropertiesGrid.RowCount - 1
            Select Case CoverPropertiesGrid.Rows(Row).Cells(1).Value 'the index numbers were updated by JBB when Curve Number, PctPeff, and Ze were added as inputs
                Case CoverVariables(1)
                    CoverPointIndex.MaximumCoverHeight = Row
                Case CoverVariables(2)
                    CoverPointIndex.MinimumCoverHeight = Row
                Case CoverVariables(3)
                    CoverPointIndex.CoverName = Row
                Case CoverVariables(4) 'Added by JBB
                    CoverPointIndex.CurveNumber = Row 'Added by JBB
                Case CoverVariables(5) 'Added by JBB
                    CoverPointIndex.EvaporativeDepth = Row 'Added by JBB
                Case CoverVariables(6) 'Added by JBB
                    CoverPointIndex.DOYefcMin = Row 'Added by JBB
                Case CoverVariables(7) 'Added by JBB
                    CoverPointIndex.DOYefcMax = Row 'Added by JBB
                Case CoverVariables(8) 'Added by JBB
                    CoverPointIndex.DOYEndWB = Row 'Added by JBB
                Case CoverVariables(9) 'Added by JBB
                    CoverPointIndex.DOYiniMin = Row 'Added by JBB
                Case CoverVariables(10) 'Added by JBB
                    CoverPointIndex.DOYiniMax = Row 'Added by JBB
                Case CoverVariables(11) 'Added by JBB
                    CoverPointIndex.DOYStartWB = Row 'Added by JBB
                Case CoverVariables(12) 'Added by JBB
                    CoverPointIndex.DOYtermMin = Row 'Added by JBB
                Case CoverVariables(13) 'Added by JBB
                    CoverPointIndex.DOYtermMax = Row 'Added by JBB
                Case CoverVariables(14) 'added by JBB
                    CoverPointIndex.FalseEndSAVI = Row 'Added by JBB

                Case CoverVariables(15) 'added by JBB
                    CoverPointIndex.FalseEndSAVIGDD = Row 'Added by JBB

                Case CoverVariables(16) 'added by JBB
                    CoverPointIndex.FalseEndSAVIDOY = Row 'Added by JBB
                Case CoverVariables(17) 'added by JBB
                    CoverPointIndex.FalsePeakSAVI = Row 'Added by JBB

                Case CoverVariables(18) 'added by JBB
                    CoverPointIndex.FalsePeakSAVIGDD = Row 'Added by JBB

                Case CoverVariables(19) 'added by JBB
                    CoverPointIndex.FalsePeakSAVIDOY = Row 'Added by JBB

                Case CoverVariables(20) 'added by JBB
                    CoverPointIndex.GDDBase = Row 'Added by JBB
                Case CoverVariables(21) 'added by JBB
                    CoverPointIndex.GDDMaxTemp = Row 'Added by JBB

                Case CoverVariables(22) 'added by JBB
                    CoverPointIndex.ApplicationEfficiency = Row 'added by JBB
                Case CoverVariables(23) 'Added by JBB
                    CoverPointIndex.KcMax = Row 'Added by JBB
                'Case CoverVariables(20) '(4)
                '    CoverPointIndex.KcbEnd = Row
                'Case CoverVariables(21) '(5)
                '    CoverPointIndex.KcbInitial = Row
                'Case CoverVariables(22) '(6)
                '    CoverPointIndex.KcbMid = Row
                Case CoverVariables(24) 'Added by JBB
                    CoverPointIndex.KcbOffSeason = Row 'Added by JBB
                Case CoverVariables(25) 'added by JBB
                    CoverPointIndex.MAD = Row 'Added by JBB
                Case CoverVariables(26) '(7)
                    CoverPointIndex.P = Row
                Case CoverVariables(27) 'Added by JBB
                    CoverPointIndex.PercentEffective = Row 'Added by JBB
                'Case CoverVariables(27) '(8)
                '    CoverPointIndex.PeriodDevelopment = Row
                'Case CoverVariables(28) '(9)
                '    CoverPointIndex.PeriodEnd = Row
                'Case CoverVariables(29) '(10)
                '    CoverPointIndex.PeriodInitial = Row
                'Case CoverVariables(30) '(11)
                '    CoverPointIndex.PeriodMid = Row
                Case CoverVariables(28) '(12)
                    CoverPointIndex.RatioKcbLate = Row 'Aded by JBB Copying others
                Case CoverVariables(29) '(12)
                    CoverPointIndex.MaximumRootDepth = Row
                Case CoverVariables(30) '(12)
                    CoverPointIndex.MinimumRootDepth = Row
                'Case CoverVariables(33) '(13)
                '    CoverPointIndex.DateInitial = Row
                Case CoverVariables(31) 'added by JBB
                    CoverPointIndex.TargetDepthAboveMAD = Row 'Added by JBB
                Case CoverVariables(32) '(12)
                    CoverPointIndex.StressBaseTemp = Row 'Aded by JBB Copying others
                Case CoverVariables(33) '(12)
                    CoverPointIndex.StressMaxTemp = Row 'Aded by JBB Copying others
                Case CoverVariables(34) 'added by JBB
                    CoverPointIndex.YearStartWB = Row 'added by JBB
                Case CoverVariables(35) 'added by JBB
                    CoverPointIndex.WeightForAssimilation = Row 'added by JBB
                Case CoverVariables(36) 'added by JBB
                    CoverPointIndex.WeightForDepletion = Row 'added by JBB
                Case CoverVariables(37) 'added by JBB
                    CoverPointIndex.WeightForEvaporatedDepth = Row 'added by JBB
                Case CoverVariables(38) 'added by JBB
                    CoverPointIndex.WeightForThetaVL = Row 'added by JBB
                Case CoverVariables(39)
                    CoverPointIndex.WeightForSkinEvap = Row 'added by JBB
                Case CoverVariables(40)
                    CoverPointIndex.xBiomassMaxSavi = Row 'added by JBB copying others
            End Select
        Next
    End Sub

    Private Sub CoverPropertiesAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CoverPropertiesAdd.Click
        Dim OpenFileDialog As ESRI.ArcGIS.CatalogUI.IGxDialog = New ESRI.ArcGIS.CatalogUI.GxDialog
        OpenFileDialog.Title = "Load cover properties parameter table"
        OpenFileDialog.AllowMultiSelect = False
        Dim Filter As ESRI.ArcGIS.Catalog.IGxObjectFilter = New ESRI.ArcGIS.Catalog.GxFilterTables
        OpenFileDialog.ObjectFilter = Filter
        Dim List As ESRI.ArcGIS.Catalog.IEnumGxObject = Nothing

        If Not OpenFileDialog.DoModalOpen(Me.Handle, List) Then Exit Sub
        Dim GxDataset As ESRI.ArcGIS.Catalog.IGxDataset = List.Next
        CoverPropertiesText.Text = GxDataset.DatasetName.WorkspaceName.PathName & "\" & GxDataset.DatasetName.Name

        System.Runtime.InteropServices.Marshal.ReleaseComObject(OpenFileDialog)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(Filter)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(List)
    End Sub

    Private Sub CoverPropertiesTableText_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CoverPropertiesText.TextChanged
        Try
            Dim CoverPropertiesTable = CreateTable(CoverPropertiesText.Text)

            Dim Variables = CoverVariables
            Array.Sort(Variables)

            CoverPropertiesGrid.Columns.Clear()
            DataGridViewColumns.AddNormal(CoverPropertiesGrid, "Column Name", DataGridViewColumns.DataTypes.typeString, True)
            DataGridViewColumns.AddCombo(CoverPropertiesGrid, "Cover Property", Variables)
            'CoverPropertiesGrid.Columns(0).Width = 175
            'CoverPropertiesGrid.Columns(1).Width = 250
            CoverPropertiesGrid.Columns(0).Width = 361 'Modified by JBB 2/1/2018
            CoverPropertiesGrid.Columns(1).Width = 386 'Modified by JBB 2/1/2018


            For Col = 0 To CoverPropertiesTable.Fields.FieldCount - 1
                CoverPropertiesGrid.Rows.Add({CoverPropertiesTable.Fields.Field(Col).Name, "", 1})
            Next

            For R = 0 To CoverPropertiesGrid.RowCount - 1 'This list has been updated by JBB beyond the specific comments
                Dim FieldName As String = CoverPropertiesGrid.Rows(R).Cells(0).Value.ToString 'Indices updated when CN, Ze, and PctPeff were added as inputs by JBB
                If FieldName.Contains("Maximum Cover") Or FieldName.Contains("HcMax") Or FieldName.Contains("Max Crop") Or FieldName.Contains("MaxCrop") Or FieldName.Contains("Maximum Crop") Then
                    CoverPropertiesGrid.Rows(R).Cells(1).Value = Variables(1)
                ElseIf FieldName.Contains("Minimum Cover") Or FieldName.Contains("HcMin") Or FieldName.Contains("Min Crop") Or FieldName.Contains("MinCrop") Or FieldName.Contains("Maximum Crop") Then
                    CoverPropertiesGrid.Rows(R).Cells(1).Value = Variables(2)
                ElseIf FieldName.Contains("Name") Then
                    CoverPropertiesGrid.Rows(R).Cells(1).Value = Variables(3)
                ElseIf FieldName.Contains("CN") Or FieldName.Contains("Curve") Then 'added by JBB
                    CoverPropertiesGrid.Rows(R).Cells(1).Value = Variables(4) 'Added by JBB
                ElseIf FieldName.Contains("Ze") Or FieldName.Contains("Evaporative") Or FieldName.Contains("vapD") Then 'added by JBB
                    CoverPropertiesGrid.Rows(R).Cells(1).Value = Variables(5) 'Added by JBB
                    'ElseIf FieldName.Contains("Kcb3") Then
                    '    CoverPropertiesGrid.Rows(R).Cells(1).Value = Variables(16) '(4)
                    'ElseIf FieldName.Contains("Kcb1") Then
                    '   CoverPropertiesGrid.Rows(R).Cells(1).Value = Variables(17) '(5)
                    'ElseIf FieldName.Contains("Kcb2") Then
                    '   CoverPropertiesGrid.Rows(R).Cells(1).Value = Variables(18) '(6)
                ElseIf FieldName.Contains("EFC Min") Or FieldName.Contains("EFCMin") Then 'added by JBB
                    CoverPropertiesGrid.Rows(R).Cells(1).Value = Variables(6) 'Added by JBB
                ElseIf FieldName.Contains("EFC Max") Or FieldName.Contains("EFCMax") Then 'added by JBB
                    CoverPropertiesGrid.Rows(R).Cells(1).Value = Variables(7) 'Added by JBB
                ElseIf FieldName.Contains("DOYEnd") Or FieldName.Contains("DOY End") Or FieldName.Contains("End DOY") Or FieldName.Contains("EndDOY") Or FieldName.Contains("EndDay") Or FieldName.Contains("End Day") Then 'added by JBB
                    CoverPropertiesGrid.Rows(R).Cells(1).Value = Variables(8) 'Added by JBB
                ElseIf FieldName.Contains("Ini Min") Or FieldName.Contains("IniMin") Or FieldName.Contains("Min Initiation") Then 'added by JBB
                    CoverPropertiesGrid.Rows(R).Cells(1).Value = Variables(9) 'Added by JBB
                ElseIf FieldName.Contains("Ini Max") Or FieldName.Contains("IniMax") Or FieldName.Contains("Max Initiation") Then 'added by JBB
                    CoverPropertiesGrid.Rows(R).Cells(1).Value = Variables(10) 'Added by JBB
                ElseIf FieldName.Contains("DOYStart") Or FieldName.Contains("DOY Start") Or FieldName.Contains("Start DOY") Or FieldName.Contains("StartDOY") Or FieldName.Contains("StartDay") Or FieldName.Contains("Start Day") Then 'added by JBB
                    CoverPropertiesGrid.Rows(R).Cells(1).Value = Variables(11) 'Added by JBB
                ElseIf FieldName.Contains("Term Min") Or FieldName.Contains("TermMin") Or FieldName.Contains("Min Termination") Then 'added by JBB
                    CoverPropertiesGrid.Rows(R).Cells(1).Value = Variables(12) 'Added by JBB
                ElseIf FieldName.Contains("Term Max") Or FieldName.Contains("TermMax") Or FieldName.Contains("Max Termination") Then 'added by JBB
                    CoverPropertiesGrid.Rows(R).Cells(1).Value = Variables(13) 'Added by JBB
                ElseIf FieldName = "False End SAVI" Or FieldName = "Projected End SAVI" Or FieldName = "Forecasted End SAVI" Or FieldName = "FalseEndSAVI" Then 'added by JBB
                    CoverPropertiesGrid.Rows(R).Cells(1).Value = Variables(14) 'Added by JBB

                ElseIf FieldName.Contains("nd") And FieldName.Contains("GDD") Then 'added by JBB
                    CoverPropertiesGrid.Rows(R).Cells(1).Value = Variables(15) 'Added by JBB

                ElseIf FieldName.Contains("nd") And FieldName.Contains("SAVI") And FieldName.Contains("DOY") Then 'added by JBB
                    CoverPropertiesGrid.Rows(R).Cells(1).Value = Variables(16) 'Added by JBB
                ElseIf FieldName = "False Peak SAVI" Or FieldName = "Projected Peak SAVI" Or FieldName = "Forecasted Peak SAVI" Or FieldName = "FalsePeakSAVI" Then 'added by JBB
                    CoverPropertiesGrid.Rows(R).Cells(1).Value = Variables(17) 'Added by JBB

                ElseIf FieldName.Contains("eak") And FieldName.Contains("GDD") Then 'added by JBB
                    CoverPropertiesGrid.Rows(R).Cells(1).Value = Variables(18) 'Added by JBB

                ElseIf FieldName.Contains("eak") And FieldName.Contains("SAVI") And FieldName.Contains("DOY") Then 'added by JBB
                    CoverPropertiesGrid.Rows(R).Cells(1).Value = Variables(19) 'Added by JBB

                ElseIf FieldName.Contains("ase") And FieldName.Contains("GDD") Then 'added by JBB
                    CoverPropertiesGrid.Rows(R).Cells(1).Value = Variables(20) 'Added by JBB
                ElseIf FieldName.Contains("ax") And FieldName.Contains("GDD") Then 'added by JBB
                    CoverPropertiesGrid.Rows(R).Cells(1).Value = Variables(21) 'Added by JBB
                ElseIf FieldName.Contains("up") And FieldName.Contains("GDD") Then 'added by JBB
                    CoverPropertiesGrid.Rows(R).Cells(1).Value = Variables(21) 'Added by JBB
                ElseIf FieldName.Contains("App") Or FieldName.Contains("Effi") Then 'added by JBB
                    CoverPropertiesGrid.Rows(R).Cells(1).Value = Variables(22) 'Added by JBB
                ElseIf FieldName.Contains("Kcmax") Then 'added by JBB
                    CoverPropertiesGrid.Rows(R).Cells(1).Value = Variables(23) 'Added by JBB
                ElseIf FieldName.Contains("off") Or FieldName.Contains("Off") Then 'added by JBB
                    CoverPropertiesGrid.Rows(R).Cells(1).Value = Variables(24) 'Added by JBB
                ElseIf FieldName = "MAD" Or FieldName.Contains("Allow") Then 'added by JBB
                    CoverPropertiesGrid.Rows(R).Cells(1).Value = Variables(25) 'Added by JBB
                ElseIf FieldName = "P" Or FieldName = "p" Or FieldName.Contains("RAW") Then
                    CoverPropertiesGrid.Rows(R).Cells(1).Value = Variables(26) '(7)
                ElseIf FieldName.Contains("Pct") Or FieldName.Contains("Effe") Or FieldName.Contains("Perce") Or FieldName.Contains("Peff") Then 'added by JBB
                    CoverPropertiesGrid.Rows(R).Cells(1).Value = Variables(27) 'Added by JBB
                    'ElseIf FieldName.Contains("P2") Then
                    '    CoverPropertiesGrid.Rows(R).Cells(1).Value = Variables(24) '(8)
                    'ElseIf FieldName.Contains("P4") Then
                    '    CoverPropertiesGrid.Rows(R).Cells(1).Value = Variables(25) '(9)
                    'ElseIf FieldName.Contains("P1") Then
                    '    CoverPropertiesGrid.Rows(R).Cells(1).Value = Variables(26) '(10)
                    'ElseIf FieldName.Contains("P3") Then
                    '    CoverPropertiesGrid.Rows(R).Cells(1).Value = Variables(27) '(11)
                ElseIf FieldName.Contains("atio") Or (FieldName.Contains("ate") And FieldName.Contains("To")) Then 'Added by JBB Copying other code
                    CoverPropertiesGrid.Rows(R).Cells(1).Value = Variables(28) 'Added by JBB Copying other code
                ElseIf FieldName.Contains("Max Root") Or FieldName.Contains("MaxRoot") Or FieldName.Contains("Zrmax") Or FieldName.Contains("Zrx") Then
                    CoverPropertiesGrid.Rows(R).Cells(1).Value = Variables(29) '(12)
                ElseIf FieldName.Contains("Min Root") Or FieldName.Contains("MinRoot") Or FieldName.Contains("Zrmin") Or FieldName.Contains("Zrn") Then
                    CoverPropertiesGrid.Rows(R).Cells(1).Value = Variables(30)
                    'ElseIf FieldName.Contains("Start") Then
                    '    CoverPropertiesGrid.Rows(R).Cells(1).Value = Variables(30)
                ElseIf FieldName.Contains("Targ") Then
                    CoverPropertiesGrid.Rows(R).Cells(1).Value = Variables(31)
                ElseIf (FieldName.Contains("st") And FieldName.Contains("b")) Or (FieldName.Contains("st") And FieldName.Contains("B")) Or (FieldName.Contains("St") And FieldName.Contains("B")) Then 'Added by JBB Copying other code
                    CoverPropertiesGrid.Rows(R).Cells(1).Value = Variables(32) 'Added by JBB Copying other code
                ElseIf (FieldName.Contains("st") And FieldName.Contains("ax")) Or (FieldName.Contains("St") And FieldName.Contains("ax")) Then 'Added by JBB Copying other code
                    CoverPropertiesGrid.Rows(R).Cells(1).Value = Variables(233) 'Added by JBB Copying other code
                ElseIf (FieldName.Contains("st") And FieldName.Contains("Up")) Or (FieldName.Contains("St") And FieldName.Contains("Up")) Or (FieldName.Contains("st") And FieldName.Contains("up")) Then 'Added by JBB Copying other code
                    CoverPropertiesGrid.Rows(R).Cells(1).Value = Variables(233) 'Added by JBB Copying other code
                ElseIf FieldName.Contains("Year") Then
                    CoverPropertiesGrid.Rows(R).Cells(1).Value = Variables(34)
                ElseIf FieldName.Contains("ET") Then
                    CoverPropertiesGrid.Rows(R).Cells(1).Value = Variables(35)
                ElseIf FieldName.Contains("Dr") Or FieldName.Contains("Depl") Then
                    CoverPropertiesGrid.Rows(R).Cells(1).Value = Variables(36)
                ElseIf FieldName.Contains("W") And FieldName.Contains("vap") Then
                    CoverPropertiesGrid.Rows(R).Cells(1).Value = Variables(37)
                ElseIf FieldName.Contains("Low") Or FieldName.Contains("Theta") Or FieldName.Contains("VL") Then
                    CoverPropertiesGrid.Rows(R).Cells(1).Value = Variables(38)
                ElseIf FieldName.Contains("Drew") Or FieldName.Contains("kin") Then
                    CoverPropertiesGrid.Rows(R).Cells(1).Value = Variables(39)
                ElseIf FieldName.Contains("BM") Or FieldName.Contains("bio") Then
                    CoverPropertiesGrid.Rows(R).Cells(1).Value = Variables(40)
                End If
            Next
        Catch ex As Exception
        End Try
    End Sub

    Private Sub WaterBalanceOutputDateAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WaterBalanceOutputDateAdd.Click 'Added by JBB, this allows for a table to be input to determine which dates to output images from the water balance, it is crude, but functional.
        Dim OpenFileDialog As ESRI.ArcGIS.CatalogUI.IGxDialog = New ESRI.ArcGIS.CatalogUI.GxDialog 'JBB mimmiced the CoverPropertiesAdd_Click sub JBB
        OpenFileDialog.Title = "Load water balance calculation date output table"
        OpenFileDialog.AllowMultiSelect = False
        Dim Filter As ESRI.ArcGIS.Catalog.IGxObjectFilter = New ESRI.ArcGIS.Catalog.GxFilterTables
        OpenFileDialog.ObjectFilter = Filter
        Dim List As ESRI.ArcGIS.Catalog.IEnumGxObject = Nothing

        If Not OpenFileDialog.DoModalOpen(Me.Handle, List) Then Exit Sub
        Dim GxDataset As ESRI.ArcGIS.Catalog.IGxDataset = List.Next
        WaterBalanceOutputDateText.Text = GxDataset.DatasetName.WorkspaceName.PathName & "\" & GxDataset.DatasetName.Name

        System.Runtime.InteropServices.Marshal.ReleaseComObject(OpenFileDialog)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(Filter)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(List)
    End Sub
    Private Sub WaterBalanceOutputDateText_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles WaterBalanceOutputDateText.TextChanged 'Added by JBB Following CoverPropertiesTableText_TextChanged Sub JBB
        Try
            Dim WaterBalacneOutputDateTable = CreateTable(WaterBalanceOutputDateText.Text)

            Dim OutDates = WtrBalOutDate
            Array.Sort(OutDates)

            WaterBalanceOutputDateGrid.Columns.Clear()
            DataGridViewColumns.AddNormal(WaterBalanceOutputDateGrid, "Column Name", DataGridViewColumns.DataTypes.typeString, True)
            DataGridViewColumns.AddCombo(WaterBalanceOutputDateGrid, "Output Property", OutDates)
            WaterBalanceOutputDateGrid.Columns(0).Width = 361 'Modified by JBB 2/1/2018
            WaterBalanceOutputDateGrid.Columns(1).Width = 406 'Modified by JBB 2/1/2018

            For Col = 0 To WaterBalacneOutputDateTable.Fields.FieldCount - 1
                WaterBalanceOutputDateGrid.Rows.Add({WaterBalacneOutputDateTable.Fields.Field(Col).Name, "", 1})
            Next

            For R = 0 To WaterBalanceOutputDateGrid.RowCount - 1
                Dim FieldName As String = WaterBalanceOutputDateGrid.Rows(R).Cells(0).Value.ToString
                If FieldName.Contains("ate") Then
                    WaterBalanceOutputDateGrid.Rows(R).Cells(1).Value = WtrBalOutDate(0)
                Else
                    'There is no Else at this time since the water balance output date table is a one column table JBB
                End If
            Next
        Catch ex As Exception
        End Try
    End Sub

    Private Sub WaterBalanceOutputDateGrid_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles WaterBalanceOutputDateGrid.CellValueChanged 'Added by JBB to handle the new water balance output date files JBB
        WBOutputDateIndex.Initialize()
        For Row = 0 To WaterBalanceOutputDateGrid.RowCount - 1
            Select Case WaterBalanceOutputDateGrid.Rows(Row).Cells(1).Value
                Case WtrBalOutDate(0)
                    WBOutputDateIndex.WBOutDate = Row
            End Select
        Next
    End Sub




    Private Sub WBPointOutputButton_Click(sender As Object, e As EventArgs) Handles WBPointOutputButton.Click
        Dim OpenFileDialog As ESRI.ArcGIS.CatalogUI.IGxDialog = New ESRI.ArcGIS.CatalogUI.GxDialog 'JBB copying the WaterBalanceOutputDateAdd_Click sub JBB
        OpenFileDialog.Title = "Load water balance point output location table"
        OpenFileDialog.AllowMultiSelect = False
        Dim Filter As ESRI.ArcGIS.Catalog.IGxObjectFilter = New ESRI.ArcGIS.Catalog.GxFilterTables
        OpenFileDialog.ObjectFilter = Filter
        Dim List As ESRI.ArcGIS.Catalog.IEnumGxObject = Nothing

        If Not OpenFileDialog.DoModalOpen(Me.Handle, List) Then Exit Sub
        Dim GxDataset As ESRI.ArcGIS.Catalog.IGxDataset = List.Next
        WBPointOutputText.Text = GxDataset.DatasetName.WorkspaceName.PathName & "\" & GxDataset.DatasetName.Name

        System.Runtime.InteropServices.Marshal.ReleaseComObject(OpenFileDialog)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(Filter)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(List)
    End Sub

    Private Sub WBPointOutputText_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles WBPointOutputText.TextChanged 'Added by JBB copied from WaterBalanceOutputDateText_TextChanged Sub JBB
        Try
            Dim WBPointOutputTable = CreateTable(WBPointOutputText.Text)

            Dim OutPnts = WBPntOutLocs
            Array.Sort(OutPnts)

            WBPointOutputGrid.Columns.Clear()
            DataGridViewColumns.AddNormal(WBPointOutputGrid, "Column Name", DataGridViewColumns.DataTypes.typeString, True)
            DataGridViewColumns.AddCombo(WBPointOutputGrid, "Location Property", OutPnts)
            WBPointOutputGrid.Columns(0).Width = 361 'Modified by JBB 2/1/2018
            WBPointOutputGrid.Columns(1).Width = 406 'Modified by JBB 2/1/2018

            For Col = 0 To WBPointOutputTable.Fields.FieldCount - 1
                WBPointOutputGrid.Rows.Add({WBPointOutputTable.Fields.Field(Col).Name, "", 1})
            Next

            For R = 0 To WBPointOutputGrid.RowCount - 1
                Dim FieldName As String = WBPointOutputGrid.Rows(R).Cells(0).Value.ToString
                If FieldName.Contains("x") Or FieldName.Contains("X") Or FieldName.Contains("on") Or FieldName.Contains("ast") Then
                    WBPointOutputGrid.Rows(R).Cells(1).Value = WBPntOutLocs(0)
                ElseIf FieldName.Contains("y") Or FieldName.Contains("Y") Or FieldName.Contains("at") Or FieldName.Contains("ort") Then
                    WBPointOutputGrid.Rows(R).Cells(1).Value = WBPntOutLocs(1)
                Else
                    'There is no Else at this time JBB
                End If
            Next
        Catch ex As Exception
        End Try
    End Sub

    Private Sub WBPointOutputGrid_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles WBPointOutputGrid.CellValueChanged 'Added by JBB to handle the new water balance output point files copied from WaterBalanceOutputDateGrid_CellValueChanged, which was probably a copy of one of the others JBB
        WBPointOutputIndex.Initialize()
        For Row = 0 To WBPointOutputGrid.RowCount - 1
            Select Case WBPointOutputGrid.Rows(Row).Cells(1).Value
                Case WBPntOutLocs(0)
                    WBPointOutputIndex.Xcoord = Row
                Case WBPntOutLocs(1)
                    WBPointOutputIndex.Ycoord = Row
            End Select
        Next
    End Sub

#End Region

#Region "Weather"

    Private WeatherPointIndex As New WeatherPointIndex
    ' Private WeatherVariables = {"", "Actual Vapor Pressure (kPa)", "Air Temperature (C)", "Air Temperature Reference Height (m)", "Anemometer Reference Height (m)", "Atmospheric Pressure (kPa)", "Cover Height (m)", "Cover Name", "Instantaneous Short Reference Evapotranspiration (mm/h)", "Irrigation (mm)", "Precipitation (mm)", "Record Date (mm/dd/yyyy)", "Relative Humidity (%)", "Short Reference Evapotranspiration (mm/day)", "Solar Radiation (W/m^2)", "Wind Speed (m/s)"} 'Commented out by JBB
    'Private WeatherVariables = {"", "Actual Vapor Pressure (kPa)", "Air Temperature (C)", "Air Temperature Reference Height (m)", "Anemometer Reference Height (m)", "Atmospheric Pressure (kPa)", "Cover Height (m)", "Cover Name", "Instantaneous Short Reference Evapotranspiration (mm/h)", "Irrigation (mm)", "Precipitation (mm)", "Record Date (mm/dd/yyyy)", "Relative Humidity (%)", "Short Reference Evapotranspiration (mm/day)", "Solar Radiation (W/m^2)", "Wind Speed (m/s)", "Wind Speed Fetch Height (m)", "Daily Available Energy (W/m^2)", "Daily Maximum Temperature (C)", "Daily Minimum Temperature (C)", "xEvaporation Scaling Factor", "xFetch Distance for Anemometer (m)", "xFetch Distance for Cover (m)", "xIBL Height Optional (m)", "xRegional Cover Height (m)"} 'Added by JBB
    Private WeatherVariables = {"", "Air Temperature Measurement Height (m)", "Cover Name", "Daily Available Energy (W/m^2)", "Daily Irrigation (mm)", "Daily Maximum Temperature (C)", "Daily Minimum Temperature (C)", "Daily Precipitation (mm)", "Daily Reference Evapotranspiration (mm/day)", "Evaporation Scaling Factor", "Fetch Distance for Modeled Cover (m)", "Fetch Distance for Weather Station (m)", "Instantaneous Actual Vapor Pressure (kPa)", "Instantaneous Air Temperature (C)", "Instantaneous Atmospheric Pressure (kPa)", "Instantaneous Reference Evapotranspiration (mm/h)", "Instantaneous Solar Radiation (W/m^2)", "Instantaneous Wind Speed (m/s)", "Internal Boundary Layer Height Optional (m)", "Measurement Area Cover Height (m)", "Date (mm/dd/yyyy)", "Regional Cover Height (m)", "Daily Average 2m Wind Speed (m/s)", "Daily Minimum Relative Humidity (%)", "Wind Speed Measurement Height (m)", "Monthly Avg. Short Reference ET - ETo (mm/d)", "Fraction of Precip. or Irrig. for the Day"} 'Added by JBB
    Private Sub WeatherTableAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WeatherTableAdd.Click
        Dim OpenFileDialog As ESRI.ArcGIS.CatalogUI.IGxDialog = New ESRI.ArcGIS.CatalogUI.GxDialog
        OpenFileDialog.Title = "Load weather parameter table"
        OpenFileDialog.AllowMultiSelect = False
        Dim Filter As ESRI.ArcGIS.Catalog.IGxObjectFilter = New ESRI.ArcGIS.Catalog.GxFilterTables
        OpenFileDialog.ObjectFilter = Filter
        Dim List As ESRI.ArcGIS.Catalog.IEnumGxObject = Nothing

        If Not OpenFileDialog.DoModalOpen(Me.Handle, List) Then Exit Sub
        Dim GxDataset As ESRI.ArcGIS.Catalog.IGxDataset = List.Next
        WeatherTableText.Text = GxDataset.DatasetName.WorkspaceName.PathName & "\" & GxDataset.DatasetName.Name

        System.Runtime.InteropServices.Marshal.ReleaseComObject(OpenFileDialog)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(Filter)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(List)
    End Sub

    Private Sub WeatherTableText_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles WeatherTableText.TextChanged
        Try
            If WeatherTableText.Text = "" Then Exit Sub
            Dim WeatherTable = CreateTable(WeatherTableText.Text)

            Dim Variables = WeatherVariables
            Array.Sort(Variables)

            WeatherTableGrid.Columns.Clear()
            DataGridViewColumns.AddNormal(WeatherTableGrid, "Column Name", DataGridViewColumns.DataTypes.typeString, True)
            DataGridViewColumns.AddCombo(WeatherTableGrid, "Weather Variable", Variables)
            DataGridViewColumns.AddNormal(WeatherTableGrid, "Multiplier", DataGridViewColumns.DataTypes.typeDouble)
            'WeatherTableGrid.Columns(0).Width = 175
            'WeatherTableGrid.Columns(1).Width = 175
            'WeatherTableGrid.Columns(2).Width = 75

            WeatherTableGrid.Columns(0).Width = 220 'Modified by JBB 2/1/2018
            WeatherTableGrid.Columns(1).Width = 413 'Modified by JBB 2/1/2018
            WeatherTableGrid.Columns(2).Width = 70 'Modified by JBB 2/1/2018



            For Col = 0 To WeatherTable.Fields.FieldCount - 1
                WeatherTableGrid.Rows.Add({WeatherTable.Fields.Field(Col).Name, "", 1})
            Next

            For Row = 0 To WeatherTableGrid.RowCount - 1 'This list has been updated by JBB beyond the specific comments
                Dim FieldName As String = WeatherTableGrid.Rows(Row).Cells(0).Value.ToString

                If FieldName.Contains("Zt") Then
                    WeatherTableGrid.Rows(Row).Cells(1).Value = Variables(1)
                ElseIf FieldName.Contains("Cover") Or FieldName.Contains("cover") Or FieldName.Equals("Crop") Or FieldName.equals("crop") Or FieldName.Contains("Name") Then
                    WeatherTableGrid.Rows(Row).Cells(1).Value = Variables(2)
                ElseIf FieldName.Contains("Ava") Or FieldName.Contains("Rn") Then 'Added by JBB
                    WeatherTableGrid.Rows(Row).Cells(1).Value = Variables(3)
                ElseIf FieldName.Contains("U2day") Or (FieldName.Contains("Sp") And fieldname.contains("da")) Then 'added by JBB copying others
                    WeatherTableGrid.Rows(Row).Cells(1).Value = Variables(4) 'added by JBB copying others
                ElseIf FieldName.Contains("Irri") Or FieldName.Contains("irri") Or FieldName = "I" Then
                    WeatherTableGrid.Rows(Row).Cells(1).Value = Variables(5)
                ElseIf FieldName.Contains("Tmax") Or FieldName.Contains("MaxT") Or FieldName.Contains("Max T") Then
                    WeatherTableGrid.Rows(Row).Cells(1).Value = Variables(6)
                ElseIf FieldName.Contains("RH") Or FieldName.Contains("um") Then 'added by JBB copying others
                    WeatherTableGrid.Rows(Row).Cells(1).Value = Variables(7) 'added by JBB copying others
                ElseIf FieldName.Contains("Tmin") Or FieldName.Contains("MinT") Or FieldName.Contains("Min T") Then
                    WeatherTableGrid.Rows(Row).Cells(1).Value = Variables(8)
                ElseIf FieldName.Contains("Prec") Or FieldName.Contains("prec") Or FieldName.Contains("Pcp") Or FieldName = "P" Then
                    WeatherTableGrid.Rows(Row).Cells(1).Value = Variables(9)
                ElseIf FieldName.Contains("Da") And FieldName.Contains("ET") Then
                    WeatherTableGrid.Rows(Row).Cells(1).Value = Variables(10)
                ElseIf FieldName.Contains("Date") Or FieldName.Contains("date") Then
                    WeatherTableGrid.Rows(Row).Cells(1).Value = Variables(11)
                ElseIf FieldName.Contains("Ke") Or FieldName.Contains("Scal") Or FieldName.Contains("Evap") Or FieldName.Contains("Adj") Then
                    WeatherTableGrid.Rows(Row).Cells(1).Value = Variables(12)
                ElseIf (FieldName.Contains("Dist") And FieldName.Contains("Cov")) Or (FieldName.Contains("Dist") And FieldName.Contains("Mod")) Then
                    WeatherTableGrid.Rows(Row).Cells(1).Value = Variables(13)
                ElseIf (FieldName.Contains("fetch") And FieldName.Contains("M") And FieldName.Contains("H")) Or (FieldName.Contains("H") And FieldName.Contains("W") And FieldName.Contains("fetch")) Then 'Added by JBB
                    WeatherTableGrid.Rows(Row).Cells(1).Value = Variables(14)
                ElseIf FieldName.Contains("rac") Or FieldName.Contains("fb") Or FieldName.Contains("Fb") Then 'added by JBB copying others
                    WeatherTableGrid.Rows(Row).Cells(1).Value = Variables(15) 'added by JBB copying others
                ElseIf FieldName.Contains("Vap") Or FieldName.Contains("vap") Or FieldName.Contains("Ea") Then
                    WeatherTableGrid.Rows(Row).Cells(1).Value = Variables(16)
                ElseIf FieldName.Contains("Ta") Then
                    WeatherTableGrid.Rows(Row).Cells(1).Value = Variables(17)
                ElseIf FieldName.Contains("Atm") Or FieldName.Contains("atm") Then
                    WeatherTableGrid.Rows(Row).Cells(1).Value = Variables(18)
                ElseIf FieldName.Contains("Inst") And FieldName.Contains("ET") Then
                    WeatherTableGrid.Rows(Row).Cells(1).Value = Variables(19)
                ElseIf FieldName.Contains("Solar") Or FieldName.Contains("solar") Or FieldName.Contains("Rs") Then
                    WeatherTableGrid.Rows(Row).Cells(1).Value = Variables(20)
                ElseIf FieldName.Contains("Spd") Or FieldName.Contains("Speed") Or FieldName.Contains("U") Then
                    WeatherTableGrid.Rows(Row).Cells(1).Value = Variables(21)
                ElseIf FieldName.Contains("IBL") Then
                    WeatherTableGrid.Rows(Row).Cells(1).Value = Variables(22)
                ElseIf (FieldName.Contains("Fetch") And FieldName.Contains("Meas")) Or (FieldName.Contains("Fetch") And FieldName.Contains("W")) Then 'Added by JBB
                    WeatherTableGrid.Rows(Row).Cells(1).Value = Variables(23)
                ElseIf FieldName.Contains("nth") And FieldName.Contains("ET") Then 'added by JBB copying others
                    WeatherTableGrid.Rows(Row).Cells(1).Value = Variables(24) 'added by JBB copying others
                ElseIf FieldName.Contains("Reg") Or FieldName.Contains("reg") Then
                    WeatherTableGrid.Rows(Row).Cells(1).Value = Variables(25)
                ElseIf FieldName.Contains("Anem") Or FieldName.Contains("anem") Or FieldName.Contains("Zw") Then
                    WeatherTableGrid.Rows(Row).Cells(1).Value = Variables(26)
                    'ElseIf FieldName.Contains("Height") Or FieldName.Contains("height") Or FieldName.Contains("Hc") Then
                    '    WeatherTableGrid.Rows(Row).Cells(1).Value = Variables(6)
                End If
            Next
        Catch ex As Exception
        End Try
    End Sub

    Private Sub WeatherTableGrid_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles WeatherTableGrid.CellValueChanged
        WeatherPointIndex.Initialize()
        For Row = 0 To WeatherTableGrid.RowCount - 1
            Select Case WeatherTableGrid.Rows(Row).Cells(1).Value
                Case WeatherVariables(1)
                    WeatherPointIndex.AirTemperatureReferenceHeight = Row
                Case WeatherVariables(2)
                    WeatherPointIndex.CoverName = Row
                Case WeatherVariables(3) 'Added by JBB copying other code
                    WeatherPointIndex.DailyAvailableEnergy = Row
                Case WeatherVariables(4)
                    WeatherPointIndex.xTypWindSpeed2m = Row 'added by JBB copying other code
                Case WeatherVariables(5)
                    WeatherPointIndex.Irrigation = Row
                Case WeatherVariables(6)
                    WeatherPointIndex.DailyMaximumTemperature = Row
                Case WeatherVariables(7) 'Added by JBB copying other code
                    WeatherPointIndex.RelativeHumidity = Row 'added by JBB copying other code
                Case WeatherVariables(8)
                    WeatherPointIndex.DailyMinimumTemperature = Row
                Case WeatherVariables(9)
                    WeatherPointIndex.Precipitation = Row
                Case WeatherVariables(10)
                    WeatherPointIndex.ShortReferenceEvapotranspiration = Row
                Case WeatherVariables(11)
                    WeatherPointIndex.RecordDate = Row
                Case WeatherVariables(12) 'Added by JBB
                    WeatherPointIndex.xEvaporationScalingFactor = Row 'Added by JBB
                Case WeatherVariables(13) 'added by JBB
                    WeatherPointIndex.xFetchDistanceCover = Row 'added by JBB for wind adjust
                Case WeatherVariables(14) 'added by JBB
                    WeatherPointIndex.xFetchDistanceAnemometer = Row 'added by JBB for wind adjust
                Case WeatherVariables(15) 'Added by JBB copying other code
                    WeatherPointIndex.xFractionWetToday = Row 'added by JBB copying other code
                Case WeatherVariables(16)
                    WeatherPointIndex.ActualVaporPressure = Row
                Case WeatherVariables(17)
                    WeatherPointIndex.AirTemperature = Row
                Case WeatherVariables(18)
                    WeatherPointIndex.AtmosphericPressure = Row
                Case WeatherVariables(19)
                    WeatherPointIndex.InstantaneousShortReferenceEvapotranspiration = Row
                Case WeatherVariables(20)
                    WeatherPointIndex.SolarRadiation = Row
                Case WeatherVariables(21)
                    WeatherPointIndex.WindSpeed = Row
                Case WeatherVariables(22) 'added by JBB
                    WeatherPointIndex.xIBLHeightOptional = Row 'added by JBB for wind adjust
                Case WeatherVariables(23) 'Added by JBB
                    WeatherPointIndex.WindFetchHeight = Row 'Added by JBB
                Case WeatherVariables(24) 'Added by JBB copying other code
                    WeatherPointIndex.xMonthlyEto = Row 'added by JBB copying other code
                Case WeatherVariables(25) 'added by JBB <-- when adding don't forget to add variable name to the Weathervariables declaration line a ways up in the code
                    WeatherPointIndex.xRegionalCoverHeight = Row 'added by JBB for wind adjust
                Case WeatherVariables(26)
                    WeatherPointIndex.AnemometerReferenceHeight = Row
                    'Case WeatherVariables(6)
                    '    WeatherPointIndex.CoverHeight = Row
            End Select
        Next
    End Sub

    Private Sub WeatherTemperatureAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WeatherTemperatureAdd.Click
        AddRastersIntoListBox(WeatherTemperatureList, "temperature")
    End Sub

    Private Sub WeatherTemperatureRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WeatherTemperatureRemove.Click
        RemoveRastersFromListBox(WeatherTemperatureList)
    End Sub

    Private Sub WeatherHumidityAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WeatherHumidityAdd.Click
        AddRastersIntoListBox(WeatherHumidityList, "humidity")
    End Sub

    Private Sub WeatherHumidityRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WeatherHumidityRemove.Click
        RemoveRastersFromListBox(WeatherHumidityList)
    End Sub

    Private Sub WeatherPressureAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WeatherPressureAdd.Click
        AddRastersIntoListBox(WeatherPressureList, "pressure")
    End Sub

    Private Sub WeatherPressureRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WeatherPressureRemove.Click
        RemoveRastersFromListBox(WeatherPressureList)
    End Sub

    Private Sub WeatherWindSpeedAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WeatherWindSpeedAdd.Click
        AddRastersIntoListBox(WeatherWindSpeedList, "wind speed")
    End Sub

    Private Sub WeatherWindSpeedRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WeatherWindSpeedRemove.Click
        RemoveRastersFromListBox(WeatherWindSpeedList)
    End Sub

    Private Sub WeatherRadiationAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WeatherRadiationAdd.Click
        AddRastersIntoListBox(WeatherRadiationList, "solar radiation")
    End Sub

    Private Sub WeatherRadiationRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WeatherRadiationRemove.Click
        'RemoveRastersFromListBox(sender)'Commented out by JBB, this line didn't work
        RemoveRastersFromListBox(WeatherRadiationList) 'Added by JBB
    End Sub

    Private Sub WeatherETDailyActualAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WeatherETDailyActualAdd.Click
        AddRastersIntoListBox(WeatherETDailyActualList, "actual evapotranspiration")
    End Sub

    Private Sub WeatherETDailyActualRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WeatherETDailyActualRemove.Click
        RemoveRastersFromListBox(WeatherETDailyActualList)
    End Sub

    '**********Added by JBB*********************
    Private Sub WeatherPrecipitationAdd_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WeatherPrecipitationAdd.Click
        AddRastersIntoListBox(WeatherPrecipitationDailyList, "actual precipitation")
    End Sub

    Private Sub WeatherPrecipitationRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WeatherPrecipitationRemove.Click
        RemoveRastersFromListBox(WeatherPrecipitationDailyList)
    End Sub

    Private Sub WeatherIrrigationAdd_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WeatherIrrigationAdd.Click
        AddRastersIntoListBox(WeatherDailyIrrigationList, "actual irrigation")
    End Sub

    Private Sub WeatherIrrigationRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WeatherIrrigationRemove.Click
        RemoveRastersFromListBox(WeatherDailyIrrigationList)
    End Sub
    Private Sub WeatherDepletionAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WeatherDepletionAdd.Click
        AddRastersIntoListBox(WeatherDailyDepletionList, "actual depletion")
    End Sub

    Private Sub WeatherDepletionRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WeatherDepletionRemove.Click
        RemoveRastersFromListBox(WeatherDailyDepletionList)
    End Sub
    Private Sub WeatherThetaVAdd_Click(sender As Object, e As EventArgs) Handles WeatherThetaVAdd.Click
        AddRastersIntoListBox(WeatherDailyThetaVList, "actual lower layer thetav")
    End Sub

    Private Sub WeatherThetaVRemove_Click(sender As Object, e As EventArgs) Handles WeatherThetaVRemove.Click
        RemoveRastersFromListBox(WeatherDailyThetaVList)
    End Sub

    Private Sub WeatherEvaporatedDepthAdd_Click(sender As Object, e As EventArgs) Handles WeatherEvaporatedDepthAdd.Click
        AddRastersIntoListBox(WeatherDailyEvaporatedDepthList, "actual evaporated depth")
    End Sub

    Private Sub WeatherEvaporatedDepthRemove_Click(sender As Object, e As EventArgs) Handles WeatherEvaporatedDepthRemove.Click
        RemoveRastersFromListBox(WeatherDailyEvaporatedDepthList)
    End Sub

    Private Sub WeatherSkinEvapDepthAdd_Click(sender As Object, e As EventArgs) Handles WeatherSkinEvapDepthAdd.Click
        AddRastersIntoListBox(WeatherDailySkinEvapDepthList, "actual evaporated depth from skin layer")
    End Sub
    Private Sub WeatherSkinEvapDepthRemove_Click(sender As Object, e As EventArgs) Handles WeatherSkinEvapDepthRemove.Click
        RemoveRastersFromListBox(WeatherDailySkinEvapDepthList)
    End Sub
    '************End Added by JBB ***************************

    Private Sub WeatherETDailyReferenceAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WeatherETDailyReferenceAdd.Click
        AddRastersIntoListBox(WeatherETDailyReferenceList, "reference evapotranspiration")
    End Sub

    Private Sub WeatherETDailyReferenceRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WeatherETDailyReferenceRemove.Click
        RemoveRastersFromListBox(WeatherETDailyReferenceList)
    End Sub

    Private Sub WeatherETInstantaneousAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WeatherETInstantaneousAdd.Click
        AddRastersIntoListBox(WeatherETInstantaneousList, "instantaneous evapotranspiration")
    End Sub

    Private Sub WeatherETInstantaneousRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WeatherETInstantaneousRemove.Click
        RemoveRastersFromListBox(WeatherETInstantaneousList)
    End Sub

#End Region

#Region "Calculation"
    Dim Abort As Boolean = True

#Region "Energy"

    Private Sub OutputDirectoryAddEnergy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OutputDirectoryAddEnergy.Click
        Dim OpenFileDialog As ESRI.ArcGIS.CatalogUI.IGxDialog = New ESRI.ArcGIS.CatalogUI.GxDialog
        OpenFileDialog.Title = "Choose output location"
        OpenFileDialog.AllowMultiSelect = False
        Dim Filter As ESRI.ArcGIS.Catalog.IGxObjectFilter = New ESRI.ArcGIS.Catalog.GxFilterBasicTypes
        OpenFileDialog.ObjectFilter = Filter
        Dim List As ESRI.ArcGIS.Catalog.IEnumGxObject = Nothing

        If Not OpenFileDialog.DoModalOpen(Me.Handle, List) Then Exit Sub

        Dim FileInfo As ESRI.ArcGIS.Catalog.IGxObject = List.Next
        OutputDirectoryTextEnergy.Text = FileInfo.FullName

        System.Runtime.InteropServices.Marshal.ReleaseComObject(OpenFileDialog)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(Filter)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(List)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(FileInfo)
    End Sub

    Private Sub OutputImagesCheckAllEnergy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OutputImagesCheckAllEnergy.Click
        For Item = 0 To OutputImagesBoxEnergy.Items.Count - 1
            OutputImagesBoxEnergy.SetItemChecked(Item, True)
        Next
    End Sub

    Private Sub OutputImagesUncheckAllEnergy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OutputImagesUncheckAllEnergy.Click
        For Item = 0 To OutputImagesBoxEnergy.Items.Count - 1
            OutputImagesBoxEnergy.SetItemChecked(Item, False)
        Next
    End Sub

    Private Sub RunEnergy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RunEnergy.Click
        If EnergyBalanceBox.Text = "" Then : TabControl1.SelectedIndex = 0 : EnergyBalanceBox.Focus() : MsgBox("Please select a model.")
        ElseIf ImageSourceBox.Text = "" Then : TabControl1.SelectedIndex = 0 : ImageSourceBox.Focus() : MsgBox("Please select an image source.")
        ElseIf TimeZoneBox.Text = "" Then : TabControl1.SelectedIndex = 0 : TimeZoneBox.Focus() : MsgBox("Please select a time zone.")
        ElseIf MultispectralList.Items.Count = 0 Then : TabControl1.SelectedIndex = 1 : MultispectralAdd.Focus() : MsgBox("Please add at least one multispectral image.")
        ElseIf SurfaceTemperatureList.Items.Count = 0 Then : TabControl1.SelectedIndex = 1 : SurfaceTemperatureAdd.Focus() : MsgBox("Please add at least one surface temperature image.")
        ElseIf CoverClassificationList.Items.Count = 0 Then : TabControl1.SelectedIndex = 2 : CoverClassificationAdd.Focus() : MsgBox("Please add at least one cover classification image.")
        Else
            Dim MultispectralDate As New List(Of DateTime)
            Dim SurfaceTemperatureDate As New List(Of DateTime)
            Dim CoverClassificationDate As New List(Of DateTime)
            Dim VegetationHeightDate As New List(Of DateTime)
            Dim LAIDate As New List(Of DateTime)
            Dim ZenithDate As New List(Of DateTime)
            Dim FcDate As New List(Of DateTime) 'added by JBB
            Dim WeatherGrid As New WeatherGrid : WeatherGrid.Clear()
            Try
                For I = 0 To MultispectralList.Items.Count - 1 : MultispectralDate.Add(GetDateFromPath(MultispectralList.Items(I))) : Next
                For I = 0 To SurfaceTemperatureList.Items.Count - 1 : SurfaceTemperatureDate.Add(GetDateFromPath(SurfaceTemperatureList.Items(I))) : Next
                For I = 0 To CoverClassificationList.Items.Count - 1 : CoverClassificationDate.Add(GetDateFromPath(CoverClassificationList.Items(I))) : Next
                For I = 0 To VegetationHeightList.Items.Count - 1 : VegetationHeightDate.Add(GetDateFromPath(VegetationHeightList.Items(I))) : Next
                For I = 0 To LAIList.Items.Count - 1 : LAIDate.Add(GetDateFromPath(LAIList.Items(I))) : Next
                For I = 0 To ZenithList.Items.Count - 1 : ZenithDate.Add(GetDateFromPath(ZenithList.Items(I))) : Next
                For I = 0 To FcList.Items.Count - 1 : FcDate.Add(GetDateFromPath(FcList.Items(I))) : Next 'added by JBB
            Catch ex As Exception
                TabControl1.SelectedIndex = 1
                MultispectralAdd.Focus()
                MsgBox("All image file names must end with acquisition date stamp, " & DateString & " HH-MM.") 'JBB added the HH-MM portion to match the actual input req's.
                Exit Sub
            End Try
            Dim MultispectralImage As New List(Of String)
            Dim SurfaceTemperatureImage As New List(Of String)
            Dim CoverClassificationImage As New List(Of String)
            Dim VegetationHeightImage As New List(Of String)
            Dim LAIImage As New List(Of String)
            Dim ZenithImage As New List(Of String)
            Dim FcImage As New List(Of String) 'added by JBB
            For I = 0 To MultispectralDate.Count - 1
                If SurfaceTemperatureDate.Contains(MultispectralDate(I)) Then
                    MultispectralImage.Add(Format(MultispectralDate(I), "yyyyMMdd") & MultispectralList.Items(I))
                    SurfaceTemperatureImage.Add(Format(MultispectralDate(I), "yyyyMMdd") & SurfaceTemperatureList.Items(SurfaceTemperatureDate.IndexOf(MultispectralDate(I))))
                End If
            Next
            For I = 0 To CoverClassificationDate.Count - 1 : CoverClassificationImage.Add(Format(CoverClassificationDate(I), "yyyyMMdd") & CoverClassificationList.Items(I)) : Next
            For I = 0 To VegetationHeightDate.Count - 1 : VegetationHeightImage.Add(Format(VegetationHeightDate(I), "yyyyMMdd") & VegetationHeightList.Items(I)) : Next
            For I = 0 To LAIDate.Count - 1 : LAIImage.Add(Format(LAIDate(I), "yyyyMMdd") & LAIList.Items(I)) : Next
            For I = 0 To ZenithDate.Count - 1 : ZenithImage.Add(Format(ZenithDate(I), "yyyyMMdd") & ZenithList.Items(I)) : Next
            For I = 0 To FcDate.Count - 1 : FcImage.Add(Format(FcDate(I), "yyyyMMdd") & FcList.Items(I)) : Next 'added by JBB
            If MultispectralImage.Count = 0 Then
                TabControl1.SelectedIndex = 1
                MultispectralAdd.Focus()
                MsgBox("No overlapping image dates between multispectral and surface temperature images.")
                Exit Sub
            End If
            MultispectralImage.Sort()
            SurfaceTemperatureImage.Sort()
            CoverClassificationImage.Sort()
            VegetationHeightImage.Sort()
            LAIImage.Sort()
            ZenithImage.Sort()
            FcImage.Sort() 'added by JBB
            For I = 0 To MultispectralImage.Count - 1
                MultispectralImage(I) = MultispectralImage(I).Remove(0, 8)
                SurfaceTemperatureImage(I) = SurfaceTemperatureImage(I).Remove(0, 8)
                Dim RecordDate As DateTime = GetDateFromPath(MultispectralImage(I))
                Dim Index As Integer
                Index = GetSameDateImageIndex(RecordDate, WeatherTemperatureList.Items.Cast(Of [String])().ToList())
                If Index > -1 Then WeatherGrid.Temperature.Add(WeatherTemperatureList.Items(Index))
                Index = GetSameDateImageIndex(RecordDate, WeatherHumidityList.Items.Cast(Of [String])().ToList())
                If Index > -1 Then WeatherGrid.SpecificHumidity.Add(WeatherHumidityList.Items(Index))
                Index = GetSameDateImageIndex(RecordDate, WeatherPressureList.Items.Cast(Of [String])().ToList())
                If Index > -1 Then WeatherGrid.Pressure.Add(WeatherPressureList.Items(Index))
                Index = GetSameDateImageIndex(RecordDate, WeatherWindSpeedList.Items.Cast(Of [String])().ToList())
                If Index > -1 Then WeatherGrid.WindSpeed.Add(WeatherWindSpeedList.Items(Index))
                Index = GetSameDateImageIndex(RecordDate, WeatherRadiationList.Items.Cast(Of [String])().ToList())
                If Index > -1 Then WeatherGrid.Radiation.Add(WeatherRadiationList.Items(Index))
                Index = GetSameDateImageIndex(RecordDate, WeatherETDailyReferenceList.Items.Cast(Of [String])().ToList())
                If Index > -1 Then WeatherGrid.ETDailyReference.Add(WeatherETDailyReferenceList.Items(Index))
                Index = GetSameDateImageIndex(RecordDate, WeatherETInstantaneousList.Items.Cast(Of [String])().ToList())
                If Index > -1 Then WeatherGrid.ETInstantaneous.Add(WeatherETInstantaneousList.Items(Index))
            Next
            For I = 0 To CoverClassificationImage.Count - 1 : CoverClassificationImage(I) = CoverClassificationImage(I).Remove(0, 8) : Next
            For I = 0 To VegetationHeightImage.Count - 1 : VegetationHeightImage(I) = VegetationHeightImage(I).Remove(0, 8) : Next
            For I = 0 To LAIImage.Count - 1 : LAIImage(I) = LAIImage(I).Remove(0, 8) : Next
            For I = 0 To ZenithImage.Count - 1 : ZenithImage(I) = ZenithImage(I).Remove(0, 8) : Next
            For I = 0 To FcImage.Count - 1 : FcImage(I) = FcImage(I).Remove(0, 8) : Next 'added by JBB
            If Not ExistsArcGISFile(MultispectralImage) Then Exit Sub
            If Not ExistsArcGISFile(SurfaceTemperatureImage) Then Exit Sub
            If Not ExistsArcGISFile(CoverClassificationImage) Then Exit Sub
            If Not ExistsArcGISFile(VegetationHeightImage) Then Exit Sub
            If Not ExistsArcGISFile(LAIImage) Then Exit Sub
            If Not ExistsArcGISFile(ZenithImage) Then Exit Sub
            If Not ExistsArcGISFile(FcImage) Then Exit Sub 'added by JBB
            If Not ExistsArcGISFile(WeatherGrid.AllValues) Then Exit Sub

            Abort = False
            CalculateEnergyBalance(MultispectralImage, SurfaceTemperatureImage, CoverClassificationImage, VegetationHeightImage, LAIImage, ZenithImage, FcImage, WeatherGrid)
        End If
        MsgBox("Energy Balance Calculations Are Complete") 'Added by JBB
    End Sub

    Private Sub ExitRunEnergy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitRunEnergy.Click
        If Abort = True Then Exit Sub 'Exits is the abort button is clicked on the form JBB
        CalculationTextEnergy.AppendText(vbNewLine & "Operation aborted..." & Now)
        Abort = True
    End Sub

    ' Private Sub CalculateEnergyBalance(ByVal MultispectralImages As List(Of String), ByVal SurfaceTemperatureImages As List(Of String), ByVal CoverClassificationImages As List(Of String), ByVal VegetationHeightImages As List(Of String), ByVal LAIImages As List(Of String), ByVal ZenithImages As List(Of String), ByVal WeatherGrid As WeatherGrid)
    Private Sub CalculateEnergyBalance(ByVal MultispectralImages As List(Of String), ByVal SurfaceTemperatureImages As List(Of String), ByVal CoverClassificationImages As List(Of String), ByVal VegetationHeightImages As List(Of String), ByVal LAIImages As List(Of String), ByVal ZenithImages As List(Of String), ByVal FcImages As List(Of String), ByVal WeatherGrid As WeatherGrid) 'Added by JBB
        Dim Timer As New Stopwatch : Timer.Start() 'Start timing the run JBB
        Dim NegFlag As Boolean = False
        ProgressAllEnergy.Maximum = MultispectralImages.Count * 2 + 2 : ProgressAllEnergy.Minimum = 0 : ProgressAllEnergy.Step = 1 : ProgressAllEnergy.Value = 0 'Set up the progress bar JBB
        ProgressPartEnergy.Minimum = 0 : ProgressPartEnergy.Step = 1 : ProgressPartEnergy.Value = 0 'Set up the lower progress bar JBB
        CalculationTextEnergy.Clear() : CalculationTextEnergy.AppendText("Determining intersecting area and output raster properties..." & Now) : Windows.Forms.Application.DoEvents() 'Print text to the progress text box JBB

        Dim InputRasterPath As New List(Of String) 'Defines the input raster path list JBB
        InputRasterPath.AddRange(MultispectralImages) 'Grabs path from the multispectral image list previously populated JBB
        InputRasterPath.AddRange(SurfaceTemperatureImages) 'Grabs path from the TIR image list previously populated JBB
        InputRasterPath.AddRange(CoverClassificationImages) 'Grabs path from the cover image list previously populated JBB
        InputRasterPath.AddRange(VegetationHeightImages) 'Grabs path from the crop height image list previously populated JBB
        InputRasterPath.AddRange(LAIImages) 'Grabs path from the LAI image list previously populated JBB
        InputRasterPath.AddRange(ZenithImages) 'Grabs path from the Zenith image list previously populated JBB
        InputRasterPath.AddRange(WeatherGrid.AllValues) 'Grabs paths from the weather grid files image list previously populated JBB
        InputRasterPath.AddRange(FcImages) 'added by JBB


        For Each File In IO.Directory.GetFiles(IO.Path.GetTempPath) 'Unsure what this was supposed to do, looks like it was to delete temp files, maybe a debug relic JBB
            Try
                'IO.File.Delete(File)
            Catch ex As Exception
            End Try
        Next

        Dim IntersectRasterPath As String = IO.Path.ChangeExtension(IO.Path.GetTempFileName, ".img") 'Gets a temp file name for the intersect image JBB
        Dim InputRasters(InputRasterPath.Count - 1) As ESRI.ArcGIS.Geodatabase.IRasterDataset 'Defines a raster dataset for the Input rasters JBB
        Dim IntersectRaster = CreateIntersectRaster(InputRasterPath.ToArray, InputRasters, IntersectRasterPath) 'Creates the intersect raster JBB
        Dim IntersectRasterValue As ESRI.ArcGIS.Geodatabase.IRasterValue = New ESRI.ArcGIS.Geodatabase.RasterValue 'From ESRI "This is the entry point to load or update raster data in the geodatabase." JBB
        IntersectRasterValue.RasterDataset = IntersectRaster 'Sets the Intersect Raster as the Raster for RasterValue JBB
        CalculationTextEnergy.AppendText(vbNewLine & "Succeeded at " & Now & vbNewLine & "Creating temporary intersecting datasets...") : Windows.Forms.Application.DoEvents()
        If Abort = True Then Exit Sub 'Exits if the abort button is clicked on the form JBB

        ProgressPartEnergy.Maximum = GetRasterCursorIterations(IntersectRaster) 'Sets lower progress bar to have number of iterations as in intesect cursor JBB

        For P = 0 To InputRasterPath.Count - 1
            CalculationTextEnergy.AppendText(vbNewLine & "   For " & IO.Path.GetFileName(InputRasterPath(P)) & "..." & Now) : Windows.Forms.Application.DoEvents() 'Update progress text JBB
            ExtractRaster(InputRasterPath(P), InputRasters(P), IntersectRasterPath, IntersectRaster) 'Extract the input rasters to the intersect raster JBB
            CalculationTextEnergy.AppendText(vbNewLine & "   Succeeded at " & Now) : Windows.Forms.Application.DoEvents() 'Update the progress text JBB
            If Abort = True Then Exit Sub 'Exits is the abort button is clicked on the form JBB
        Next
        CalculationTextEnergy.AppendText(vbNewLine & "   For transform raster dataset..." & Now) : Windows.Forms.Application.DoEvents() 'Update progress text JBB
        'Dim TransformRaster2 = TransformRaster(IntersectRasterPath, IntersectRaster, ESRI.ArcGIS.Geometry.esriSRGeoCSType.esriSRGeoCS_NAD1983) 'Makes a raster that is a Transform of the intersect raster into NAD1983 JBB commented out by JBB
        Dim TransformRaster2 = TransformRaster(IntersectRasterPath, IntersectRaster, ESRI.ArcGIS.Geometry.esriSRGeoCSType.esriSRGeoCS_WGS1984) 'Makes a raster that is a Transform of the intersect raster into WGS84 JBB added by JBB, this is used in TSEB for pixel lat and lon
        CalculationTextEnergy.AppendText(vbNewLine & "   Succeeded at " & Now) : ProgressAllEnergy.PerformStep() : Windows.Forms.Application.DoEvents() 'Update progress text JBB

        Dim OutputRasterNames As New List(Of String) 'Defines a list of output raster names JBB
        For I = 0 To OutputImagesBoxEnergy.Items.Count - 1 'populates the list of output raster names JBB
            If OutputImagesBoxEnergy.GetItemChecked(I) Then OutputRasterNames.Add(OutputImagesBoxEnergy.Items.Item(I))
        Next
        Dim WorkspaceFactory As ESRI.ArcGIS.Geodatabase.IWorkspaceFactory = New ESRI.ArcGIS.DataSourcesRaster.RasterWorkspaceFactoryClass() 'from ESRI: use a workspacefactory  "when you need to create a new workspace, connect to an existing workspace or find information about a workspace" JBB
        Dim Workspace As ESRI.ArcGIS.Geodatabase.IRasterWorkspace2 = CType(WorkspaceFactory.OpenFromFile(OutputDirectoryTextEnergy.Text, 0), ESRI.ArcGIS.DataSourcesRaster.IRasterWorkspace) 'From ESRI "A Workspace is a container of spatial and non-spatial datasets such as feature classes, raster datasets and tables. It provides methods to instantiate existing datasets and to create new datasets." JBB
        Dim OutputFileFormat = "IMAGINE Image" 'Output Files will be .img JBB

        Dim EnergyBalanceModel As EnergyBalance = DirectCast([Enum].Parse(GetType(EnergyBalance), EnergyBalanceBox.SelectedItem.ToString.Replace(" ", "_")), EnergyBalance)
        Dim ImageSource As ImageSource = DirectCast([Enum].Parse(GetType(ImageSource), ImageSourceBox.SelectedItem.ToString.Replace(" ", "_")), ImageSource)
        Dim WindAdjustment As WindAdjustment = DirectCast([Enum].Parse(GetType(WindAdjustment), WindAdjustMethodBox.SelectedItem.ToString.Replace(" ", "_")), WindAdjustment) 'Added by JBB
        Dim FgMethod As FgMethod = DirectCast([Enum].Parse(GetType(FgMethod), FgMethodBox.SelectedItem.ToString.Replace(" ", "_")), FgMethod) 'added by JBB
        Dim HcMethod As HcMethod = DirectCast([Enum].Parse(GetType(HcMethod), HcMethodBox.SelectedItem.ToString.Replace(" ", "_")), HcMethod) 'added by JBB
        Dim FcMethod As FcMethod = DirectCast([Enum].Parse(GetType(FcMethod), FcMethodBox.SelectedItem.ToString.Replace(" ", "_")), FcMethod) 'added by JBB for Fc Peak
        Dim ClumpDMethod As ClumpingD = DirectCast([Enum].Parse(GetType(ClumpingD), ClumpingMethodBox.SelectedItem.ToString.Replace(" ", "_")), ClumpingD) 'added by JBB for clumping method mimmicking code from the above lines
        Dim TSMInitialTemperature As TSMInitialTemperature = DirectCast([Enum].Parse(GetType(TSMInitialTemperature), TSMInitialTemperatureBox.SelectedItem.ToString.Replace(" ", "_")), TSMInitialTemperature) 'Added by JBB for PM
        Dim ETextrapolation As ETExtrapolation = DirectCast([Enum].Parse(GetType(ETExtrapolation), ETExtrapolationBox.SelectedItem.ToString.Replace(" ", "_")), ETExtrapolation) 'added by JBB
        Dim WeatherTable = CreateTable(WeatherTableText.Text) 'Creates a table for the input weather .xls table JBB
        Dim WeatherPoint As New WeatherPoint 'Weather point is the tabulated weather data JBB
        Dim WeatherGridIndex As New WeatherGridIndex 'Weather grid index is used to pair the gridded weather data with the calculation date JBB
        'Dim WeatherOffset As Integer = MultispectralImages.Count * 2 + CoverClassificationImages.Count + VegetationHeightImages.Count + LAIImages.Count + ZenithImages.Count 'Finds the image index where weather images start JBB
        Dim WeatherOffset As Integer = MultispectralImages.Count * 2 + CoverClassificationImages.Count + VegetationHeightImages.Count + LAIImages.Count + ZenithImages.Count + FcImages.Count 'Finds the image index where weather images start JBB


        Dim BandIndex(5) As Integer 'Defines an array for the band indexes as input in the SETMI GUI, The code below populates that array, -1 means not included JBB
        BandIndex(0) = -1 : If RedIndex.Text <> "" Then BandIndex(0) = RedIndex.Text - 1
        BandIndex(1) = -1 : If GreenIndex.Text <> "" Then BandIndex(1) = GreenIndex.Text - 1
        BandIndex(2) = -1 : If BlueIndex.Text <> "" Then BandIndex(2) = BlueIndex.Text - 1
        BandIndex(3) = -1 : If NIRIndex.Text <> "" Then BandIndex(3) = NIRIndex.Text - 1
        BandIndex(4) = -1 : If MIR1Index.Text <> "" Then BandIndex(4) = MIR1Index.Text - 1
        BandIndex(5) = -1 : If MIR2Index.Text <> "" Then BandIndex(5) = MIR2Index.Text - 1

        Dim M As Integer = 0
        For M = 0 To MultispectralImages.Count - 1 'Loops through Multispectral images JBB
            CalculationTextEnergy.AppendText(vbNewLine & "Calculating output for " & MultispectralImages(M) & "..." & Now) : Windows.Forms.Application.DoEvents() 'Updates the progress text box JBB
            ProgressPartEnergy.Value = 0 'Initializes the lower progress bar.

            Dim RecordDate As DateTime = GetDateFromPath(MultispectralImages(M)) 'Grabs the date-time from the image file name JBB
            Dim ReferenceLongitude As Single = 15 * (TimeZoneInfo.FindSystemTimeZoneById(TimeZoneBox.SelectedItem).BaseUtcOffset.TotalHours) 'Grabs the longitude of the center of the time zone from the SETMI GUI JBB
            Dim ReferenceTimeZone As Integer = (TimeZoneInfo.FindSystemTimeZoneById(TimeZoneBox.SelectedItem).BaseUtcOffset.TotalHours) 'Added by JBB for solar noon calculation JBB
            Dim CoverClassificationIndex As Integer = GetNearestDateImageIndex(MultispectralImages(M), CoverClassificationImages) 'Pairs a cover classification image with the multispectral image JBB
            Dim VegetationHeightIndex As Integer = GetSameDateImageIndex(MultispectralImages(M), VegetationHeightImages) 'Pairs a vegetation height image with the multispectral image, must be the same date JBB
            Dim LAIIndex As Integer = GetSameDateImageIndex(MultispectralImages(M), LAIImages)
            Dim LAICount As Integer = LAIImages.Count 'added by JBB for LAIPeak for Fg
            Dim HcCount As Integer = VegetationHeightImages.Count ' added by JBB for peak cover height calc
            Dim FcIndex As Integer = GetSameDateImageIndex(MultispectralImages(M), FcImages) 'added by JBB for peak Fc
            Dim FcCount As Integer = FcImages.Count 'added by JBB for FcPeak
            Dim ZenithIndex As Integer = GetSameDateImageIndex(MultispectralImages(M), ZenithImages)
            WeatherPoint.Populate(WeatherTable, WeatherPointIndex, RecordDate) 'Populates a table with the input weather .xls table JBB
            WeatherGridIndex.Initialize(RecordDate, WeatherGrid, WeatherOffset) 'This finds the index of the weather grid images, if exist, otherwise the indeces are -1 JBB

            Dim IntersectRasterBand As ESRI.ArcGIS.DataSourcesRaster.IRasterBandCollection = IntersectRaster ' from ESRI: "Provides access to members that control a collection of RasterBands" JBB
            Dim IntersectRaster2 As ESRI.ArcGIS.DataSourcesRaster.IRaster2 = CType(CType(IntersectRaster.CreateDefaultRaster, ESRI.ArcGIS.Geodatabase.IRaster), ESRI.ArcGIS.DataSourcesRaster.IRaster2) 'from ESRI: The CreateDefaultRaster method creates a Raster that has a square cell size and contains only three raster bands if the dataset has more than three bands.JBB
            Dim IntersectRasterCursor As ESRI.ArcGIS.Geodatabase.IRasterCursor = IntersectRaster2.CreateCursorEx(Nothing) 'Creates a cursor for the intersect raster JBB

            Dim InRasterBand(InputRasters.Count - 1) As ESRI.ArcGIS.DataSourcesRaster.IRasterBandCollection 'Declares a raster band collection for the input rasters JBB
            Dim InRaster2(InputRasters.Count - 1) As ESRI.ArcGIS.DataSourcesRaster.IRaster2 'from ESRI: "Provides access to members that control a raster" JBB
            Dim InRasterCursor(InputRasters.Count - 1) As ESRI.ArcGIS.Geodatabase.IRasterCursor 'Declares a cursor for the input rasters JBB
            For R = 0 To InputRasters.Count - 1 'Populates the input raster band, raster control, and cursor JBB
                InRasterBand(R) = InputRasters(R)
                InRaster2(R) = CType(InputRasters(R), ESRI.ArcGIS.Geodatabase.IRasterDataset2).CreateFullRaster
                InRasterCursor(R) = InRaster2(R).CreateCursorEx(Nothing)
            Next

            Dim OutputRasters(OutputRasterNames.Count - 1) As ESRI.ArcGIS.Geodatabase.IRasterDataset 'Declares a raster dataset for output rasters JBB
            For I = 0 To OutputImagesBoxEnergy.CheckedItems.Count - 1 'Names, creates, and deletes previous versions of selected output rasters JBB
                Dim FileName As String = "SETMI_" & EnergyBalanceModel.ToString & "_" & OutputImagesBoxEnergy.CheckedItems(I).ToString.Split("(")(1).Replace(")", "") & "_" & Format(RecordDate, "MM-dd-yyyy_HH-mm") & ".img"

                If ExistsArcGISFile(OutputDirectoryTextEnergy.Text & "\" & FileName) Then DeleteArcGISFile(OutputDirectoryTextEnergy.Text & "\" & FileName)
                Dim OutputRasterDataset = Workspace.CreateRasterDataset(Name:=FileName,
                                                                        Format:=OutputFileFormat,
                                                                        Origin:=IntersectRasterValue.Extent.LowerLeft,
                                                                        columnCount:=IntersectRasterValue.Extent.Width / IntersectRasterValue.RasterStorageDef.CellSize.X,
                                                                        RowCount:=IntersectRasterValue.Extent.Height / IntersectRasterValue.RasterStorageDef.CellSize.Y,
                                                                        cellSizeX:=IntersectRasterValue.RasterStorageDef.CellSize.X,
                                                                        cellSizeY:=IntersectRasterValue.RasterStorageDef.CellSize.Y,
                                                                        numBands:=1,
                                                                        PixelType:=ESRI.ArcGIS.Geodatabase.rstPixelType.PT_FLOAT,
                                                                        SpatialReference:=IntersectRasterValue.Extent.SpatialReference,
                                                                        Permanent:=True)
                OutputRasters(I) = OutputRasterDataset
            Next

            Dim OutRasterBand(OutputRasters.Count - 1) As ESRI.ArcGIS.DataSourcesRaster.IRasterBandCollection 'Creates a raster band collection for the output rasters JBB
            Dim OutRaster2(OutputRasters.Count - 1) As ESRI.ArcGIS.DataSourcesRaster.IRaster2 'Creates a Raster control for output rasters JBB
            Dim OutRasterCursor(OutputRasters.Count - 1) As ESRI.ArcGIS.Geodatabase.IRasterCursor 'Creaters a cursor for the output rasters JBB
            Dim OutRasterEdit(OutputRasters.Count - 1) As ESRI.ArcGIS.DataSourcesRaster.IRasterEdit 'from ESRI "Provides access to members that control pixel block level editing operations" JBB
            For R = 0 To OutputRasters.Count - 1 'Populates the items declared above JBB
                OutRasterBand(R) = OutputRasters(R)
                OutRaster2(R) = CType(CType(OutputRasters(R).CreateDefaultRaster, ESRI.ArcGIS.Geodatabase.IRaster), ESRI.ArcGIS.DataSourcesRaster.IRaster2)
                OutRasterCursor(R) = OutRaster2(R).CreateCursorEx(Nothing)
                OutRasterEdit(R) = OutRaster2(R)
            Next
            If Abort = True Then Exit Sub 'Exits is the abort button is clicked on the form JBB
            ProgressAllEnergy.PerformStep() : Windows.Forms.Application.DoEvents() 'Advance the upper progress bar JBB

            Do 'Loop through input raster cursor JBB
                Dim IntersectPixels As System.Array = CType(CType(IntersectRasterCursor.PixelBlock, ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3).PixelData(0), System.Array) 'Makes a pixel array for the intersect raster JBB
                Dim MultispectralPixelBlock As ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3 = InRasterCursor(M).PixelBlock 'Makes a pixel block for the multispectral image "M" from ESRI: "The PixelBlock object contains a pixel array that can be read from a raster or a raster band." JBB
                Dim MultispectralPixels(MultispectralPixelBlock.Planes - 1) As System.Array 'Makes an array for pixels from the multipsectral image "M" dimensioned for each plane or band JBB
                For I = 0 To MultispectralPixelBlock.Planes - 1 'Populates the pixel array for the multispectral images JBB
                    MultispectralPixels(I) = CType(MultispectralPixelBlock.PixelData(I), System.Array)
                Next
                Dim Offset = MultispectralImages.Count 'Offset sets the index for TIR images JBB
                Dim TemperaturePixels As System.Array = CType(CType(InRasterCursor(Offset + M).PixelBlock, ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3).PixelData(0), System.Array) 'Creates an array for TIR pixels JBB
                Offset += SurfaceTemperatureImages.Count 'Offset advanced to set index for cover images JBB
                Dim CoverPixels As System.Array = CType(CType(InRasterCursor(Offset + CoverClassificationIndex).PixelBlock, ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3).PixelData(0), System.Array) 'Creates an array for cover pixels JBB
                Offset += CoverClassificationImages.Count 'Offset advanced to set index for veg height images JBB
                Dim VegetationHeightPixels As System.Array = Nothing 'creates an empty array for vegetation height pixels JBB
                If VegetationHeightIndex > -1 Then VegetationHeightPixels = CType(CType(InRasterCursor(Offset + VegetationHeightIndex).PixelBlock, ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3).PixelData(0), System.Array) 'Only populates the veg height pixel array if there is a veg height image provided JBB

                'Find Max Cover Height

                Dim HcPeakPixels As System.Array = Nothing 'added by JBB for peak Hc
                Dim HcPeakPastPixels As System.Array = Nothing 'added by JBB for peak Hc
                If HcCount > 0 Then
                    HcPeakPastPixels = CType(CType(InRasterCursor(Offset + 0).PixelBlock, ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3).PixelData(0), System.Array) 'added by JB
                    For LL = 1 To HcCount - 1
                        HcPeakPixels = CType(CType(InRasterCursor(Offset + LL).PixelBlock, ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3).PixelData(0), System.Array) 'added by JBB for peak Hc
                        Dim ColumnCount1 As Integer = MultispectralPixelBlock.Width
                        For I = 0 To MultispectralPixelBlock.Height * MultispectralPixelBlock.Width - 1
                            Dim Row1 As Integer = Int(I / ColumnCount1) 'I is the index of a pixel, Row is the index of the corresponding row (this works) JBB
                            Dim Col1 As Integer = I - ColumnCount1 * Row1 'Col is the index of the column for pixl I (this works) JBB
                            If IntersectPixels.GetValue(Col1, Row1) = 1 Then 'If this pixel is an intersecting pixel continue JBB
                                Dim Hcnow As Single = HcPeakPixels.GetValue(Col1, Row1)
                                Dim PeakHc As Single = HcPeakPastPixels.GetValue(Col1, Row1)
                                If Hcnow > PeakHc Then
                                    HcPeakPastPixels(Col1, Row1) = Hcnow
                                End If
                            End If
                        Next
                    Next
                End If

                Offset += VegetationHeightImages.Count 'Offset advanced to set index for LAI images JBB
                Dim LAIPixels As System.Array = Nothing 'creates an empty array for LAI pixels JBB
                If LAIIndex > -1 Then LAIPixels = CType(CType(InRasterCursor(Offset + LAIIndex).PixelBlock, ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3).PixelData(0), System.Array) 'Only populates the LAI pixel array if there is an LAI image provided JBB

                'Find Peak Input LAI

                Dim LAIPeakPixels As System.Array = Nothing 'added by JBB for peak LAI
                Dim LAIPeakPastPixels As System.Array = Nothing 'added by JBB for peak LAI
                If LAICount > 0 Then
                    LAIPeakPastPixels = CType(CType(InRasterCursor(Offset + 0).PixelBlock, ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3).PixelData(0), System.Array) 'added by JB
                    For LL = 1 To LAICount - 1
                        LAIPeakPixels = CType(CType(InRasterCursor(Offset + LL).PixelBlock, ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3).PixelData(0), System.Array) 'added by JBB for peak LAI
                        Dim ColumnCount1 As Integer = MultispectralPixelBlock.Width
                        For I = 0 To MultispectralPixelBlock.Height * MultispectralPixelBlock.Width - 1
                            Dim Row1 As Integer = Int(I / ColumnCount1) 'I is the index of a pixel, Row is the index of the corresponding row (this works) JBB
                            Dim Col1 As Integer = I - ColumnCount1 * Row1 'Col is the index of the column for pixl I (this works) JBB
                            If IntersectPixels.GetValue(Col1, Row1) = 1 Then 'If this pixel is an intersecting pixel continue JBB
                                Dim LAInow As Single = LAIPeakPixels.GetValue(Col1, Row1)
                                Dim PeakLAI As Single = LAIPeakPastPixels.GetValue(Col1, Row1)
                                If LAInow > PeakLAI Then
                                    LAIPeakPastPixels(Col1, Row1) = LAInow
                                End If
                            End If
                        Next
                    Next
                End If

                Offset += LAIImages.Count 'Offset advanced to set index for fc images JBB
                Dim ZenithPixels As System.Array = Nothing 'creates an empty array for Zenith pixels JBB
                If ZenithIndex > -1 Then ZenithPixels = CType(CType(InRasterCursor(Offset + ZenithIndex).PixelBlock, ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3).PixelData(0), System.Array) 'Only populates the Zenith pixel array if there is a Zenith image provided JBB
                Offset += ZenithImages.Count 'Offset advanced to set index for fc images JBB
                Dim FcPixels As System.Array = Nothing 'creates an empty array for fc pixels JBB
                If FcIndex > -1 Then FcPixels = CType(CType(InRasterCursor(Offset + FcIndex).PixelBlock, ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3).PixelData(0), System.Array) 'Only populates the fc pixel array if there is an fc image provided JBB

                'Find Peak Input Fc

                Dim FcPeakPixels As System.Array = Nothing 'added by JBB for peak Fc
                Dim FcPeakPastPixels As System.Array = Nothing 'added by JBB for peak Fc
                If FcCount > 0 Then
                    FcPeakPastPixels = CType(CType(InRasterCursor(Offset + 0).PixelBlock, ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3).PixelData(0), System.Array) 'added by JB
                    For LL = 1 To FcCount - 1
                        FcPeakPixels = CType(CType(InRasterCursor(Offset + LL).PixelBlock, ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3).PixelData(0), System.Array) 'added by JBB for peak fc
                        Dim ColumnCount1 As Integer = MultispectralPixelBlock.Width
                        For I = 0 To MultispectralPixelBlock.Height * MultispectralPixelBlock.Width - 1
                            Dim Row1 As Integer = Int(I / ColumnCount1) 'I is the index of a pixel, Row is the index of the corresponding row (this works) JBB
                            Dim Col1 As Integer = I - ColumnCount1 * Row1 'Col is the index of the column for pixl I (this works) JBB
                            If IntersectPixels.GetValue(Col1, Row1) = 1 Then 'If this pixel is an intersecting pixel continue JBB
                                Dim Fcnow As Single = FcPeakPixels.GetValue(Col1, Row1)
                                Dim PeakFc As Single = FcPeakPastPixels.GetValue(Col1, Row1)
                                If Fcnow > PeakFc Then
                                    FcPeakPastPixels(Col1, Row1) = Fcnow
                                End If
                            End If
                        Next
                    Next
                End If



                Dim WeatherTemperaturePixels As System.Array = Nothing 'creates an empty array for Temp pixels JBB
                If WeatherGridIndex.Temperature > -1 Then WeatherTemperaturePixels = CType(CType(InRasterCursor(WeatherGridIndex.Temperature).PixelBlock, ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3).PixelData(0), System.Array) 'Only populates the Temperature pixel array if there is a Temperature image provided JBB
                Dim WeatherHumidityPixels As System.Array = Nothing 'creates an empty array for Humidity pixels JBB
                If WeatherGridIndex.SpecificHumidity > -1 Then WeatherHumidityPixels = CType(CType(InRasterCursor(WeatherGridIndex.SpecificHumidity).PixelBlock, ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3).PixelData(0), System.Array) 'Only populates the humidity pixel array if there is a humidity image provided JBB
                Dim WeatherPressurePixels As System.Array = Nothing 'creates an empty array for Pressure pixels JBB
                If WeatherGridIndex.Pressure > -1 Then WeatherPressurePixels = CType(CType(InRasterCursor(WeatherGridIndex.Pressure).PixelBlock, ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3).PixelData(0), System.Array) 'Only populates the pressure pixel array if there is a pressure image provided JBB
                Dim WeatherWindSpeedPixels As System.Array = Nothing 'creates an empty array for Wind Speed pixels JBB
                If WeatherGridIndex.WindSpeed > -1 Then WeatherWindSpeedPixels = CType(CType(InRasterCursor(WeatherGridIndex.WindSpeed).PixelBlock, ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3).PixelData(0), System.Array) 'Only populates the wind speed pixel array if there is a wind speed image provided JBB
                Dim WeatherRadiationPixels As System.Array = Nothing 'creates an empty array for Radiation pixels JBB
                If WeatherGridIndex.Radiation > -1 Then WeatherRadiationPixels = CType(CType(InRasterCursor(WeatherGridIndex.Radiation).PixelBlock, ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3).PixelData(0), System.Array) 'Only populates the radiation pixel array if there is a radiation image provided JBB
                Dim WeatherETDailyReferencePixels As System.Array = Nothing 'creates an empty array for Daily ETo pixels JBB
                If WeatherGridIndex.ETDailyReference > -1 Then WeatherETDailyReferencePixels = CType(CType(InRasterCursor(WeatherGridIndex.ETDailyReference).PixelBlock, ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3).PixelData(0), System.Array) 'Only populates the DailyETo pixel array if there is a DailyETo image provided JBB
                Dim WeatherETInstantaneousPixels As System.Array = Nothing 'creates an empty array for Instantaneoud ETo pixels JBB
                If WeatherGridIndex.ETInstantaneous > -1 Then WeatherETInstantaneousPixels = CType(CType(InRasterCursor(WeatherGridIndex.ETInstantaneous).PixelBlock, ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3).PixelData(0), System.Array) 'Only populates the InstETo pixel array if there is an InstETo image provided JBB

                Dim OutRasterPixelBlock(OutputRasters.Count - 1) As ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3 'Declares a pixel block for the output rasters JBB
                Dim OutPixels(OutputRasters.Count - 1) As System.Array 'Declares and output pixel array JBB
                For I = 0 To OutputRasters.Count - 1 'Populates the output pixel array and pixel block JBB
                    OutRasterPixelBlock(I) = OutRasterCursor(I).PixelBlock
                    OutPixels(I) = CType(OutRasterPixelBlock(I).PixelData(0), System.Array)
                Next

                Dim ColumnCount As Integer = MultispectralPixelBlock.Width 'Finds width of the multispectral image JBB
                'The line below allows calculations to be made for different pixels in parallel JBB
                'Threading.Tasks.Parallel.For(0, MultispectralPixelBlock.Height * MultispectralPixelBlock.Width, _
                'Sub(I)
                ' For I = 0 To MultispectralPixelBlock.Height * MultispectralPixelBlock.Width 'Added by JBB for single threaded processing for debugging
                '  Dim Row As Integer = Int(I / ColumnCount) 'I is the index of a pixel, Row is the index of the corresponding row (this works) JBB
                ' Dim Col As Integer = I - ColumnCount * Row 'Col is the index of the column for pixl I (this works) JBB

                For Row As Long = 0 To MultispectralPixelBlock.Height - 1 'Loop through images JBB added by JBB to follow single thread in water balance
                    For Col As Long = 0 To MultispectralPixelBlock.Width - 1 'Loop through images JBB added by JBB to follow single thread in water balance

                        If IntersectPixels.GetValue(Col, Row) = 1 Then 'If this pixel is an intersecting pixel continue JBB
                            Dim CoverIndex As Integer = CoverValues.IndexOf(CoverPixels.GetValue(Col, Row)) 'Grabs the cover index JBB
                            If CoverIndex > -1 Then 'If there is cover, there is no Else for this if JBB
                                Dim f 'Looks like a debug relic JBB
                                If CoverIndex = 4 Then 'Looks like a debug relic JBB
                                    f = 4 'Looks like a debug relic JBB
                                End If 'Looks like a debug relic JBB

                                Dim IWeather As Integer = WeatherPoint.CoverName.IndexOf(CoverSelectionGrid.Rows(CoverIndex).Cells(0).Value) 'Finds the index for the weather data for the selected cover JBB
                                If CoverSelectionGrid.Rows(CoverIndex).Cells(3).Value = True And IWeather > -1 Then 'Checks if the cover has been selected as "included" in the cover tab of the GUI JBB
                                    Dim Cover As Cover = DirectCast([Enum].Parse(GetType(Cover), CoverSelectionGrid.Rows(CoverIndex).Cells(1).Value.Replace(" ", "_")), Cover) 'Grabs the cover name JBB
                                    Dim Red As Single = 0 : If BandIndex(0) > -1 Then Red = MultispectralPixels(BandIndex(0)).GetValue(Col, Row) 'grabs the red reflectance, if it has been selected as being provided JBB
                                    Dim Green As Single = 0 : If BandIndex(1) > -1 Then Green = MultispectralPixels(BandIndex(1)).GetValue(Col, Row) 'grabs the green reflectance, if it has been selected as being provided JBB
                                    Dim Blue As Single = 0 : If BandIndex(2) > -1 Then Blue = MultispectralPixels(BandIndex(2)).GetValue(Col, Row) 'grabs the blue reflectance, if it has been selected as being provided JBB
                                    Dim NIR As Single = 0 : If BandIndex(3) > -1 Then NIR = MultispectralPixels(BandIndex(3)).GetValue(Col, Row) 'grabs the NIR reflectance, if it has been selected as being provided JBB
                                    Dim MIR1 As Single = 0 : If BandIndex(4) > -1 Then MIR1 = MultispectralPixels(BandIndex(4)).GetValue(Col, Row) 'grabs the MIR1 reflectance, if it has been selected as being provided JBB
                                    Dim MIR2 As Single = 0 : If BandIndex(5) > -1 Then MIR2 = MultispectralPixels(BandIndex(5)).GetValue(Col, Row) 'grabs the MIR2 reflectance, if it has been selected as being provided JBB
                                    Dim Temperature As Single = TemperaturePixels.GetValue(Col, Row) + 273.15 'Grabs temperature from the TIR image and converts to Kelvin JBB

                                    ''Debugging only
                                    'Temperature = 283.15
                                    'MsgBox("Fix Debugging Temp")
                                    ''End Debugging
                                    Dim Bioproperties As New Bioproperties 'Bioproperties is a class of variables declared as singles JBB
                                    Bioproperties.AlphaLeafVIS = CoverSelectionGrid.Rows(CoverIndex).Cells(4).Value 'absorptivity leaf vis pulled from input table JBB
                                    Bioproperties.AlphaLeafNIR = CoverSelectionGrid.Rows(CoverIndex).Cells(5).Value 'absorptivity leaf NIR pulled from input table JBB
                                    Bioproperties.AlphaLeafTIR = CoverSelectionGrid.Rows(CoverIndex).Cells(6).Value 'absorptivity leaf TIR pulled from input table JBB
                                    Bioproperties.AlphaDeadVIS = CoverSelectionGrid.Rows(CoverIndex).Cells(7).Value 'absorptivity dead leaf vis pulled from input table JBB'Before 1/10/17 was nir
                                    Bioproperties.AlphaDeadNIR = CoverSelectionGrid.Rows(CoverIndex).Cells(8).Value 'absorptivity dead leaf NIR pulled from input table JBB'before 1/10/17 was tir
                                    Bioproperties.AlphaDeadTIR = CoverSelectionGrid.Rows(CoverIndex).Cells(9).Value 'absorptivity dead leaf TIR pulled from input table JBB'before 1/10/17 was vis
                                    Bioproperties.fg = CoverSelectionGrid.Rows(CoverIndex).Cells(10).Value 'Grabs fg from the input table in the GUI, fg is fraction of green vs. dead leaves JBB
                                    Bioproperties.HcMin = CoverSelectionGrid.Rows(CoverIndex).Cells(11).Value ' Grabs minimum crop height from the GUI JBB
                                    Bioproperties.HcMax = CoverSelectionGrid.Rows(CoverIndex).Cells(12).Value 'Grabs maximum crop height from the GUI JBB
                                    Bioproperties.s = CoverSelectionGrid.Rows(CoverIndex).Cells(13).Value 'used for clumping,  JBB
                                    Bioproperties.Wc = CoverSelectionGrid.Rows(CoverIndex).Cells(14).Value 'Wc is canopy width and is used in clumping equations JBB
                                    Bioproperties.LAITable = CoverSelectionGrid.Rows(CoverIndex).Cells(15).Value 'Grabs LAI from the table, but default for the table is -999,so this must be a relic JBB
                                    Bioproperties.EmissSoilVIS = CoverSelectionGrid.Rows(CoverIndex).Cells(16).Value 'Grabs soil emissivity in VIS from table, added by JBB
                                    Bioproperties.EmissSoilNIR = CoverSelectionGrid.Rows(CoverIndex).Cells(17).Value 'Grabs soil emissivity in NIR from table, added by JBB
                                    Bioproperties.EmissSoilTIR = CoverSelectionGrid.Rows(CoverIndex).Cells(18).Value 'Grabs soil emissivity in TIR from table ,added by JBB
                                    Bioproperties.SoilHtFlxAg = CoverSelectionGrid.Rows(CoverIndex).Cells(19).Value 'Grabs Soil Heat Flux : Rn Ratio from table, added by JBB
                                    'Added by JBB to reduce the hardcoded values, etc.
                                    Bioproperties.ClumpD = CoverSelectionGrid.Rows(CoverIndex).Cells(20).Value 'Grabs Wc/H Ratio from table, added by JBB
                                    Bioproperties.aPTInput = CoverSelectionGrid.Rows(CoverIndex).Cells(21).Value 'Grabs input aPT from table, added by JBB
                                    Bioproperties.RcIniInput = CoverSelectionGrid.Rows(CoverIndex).Cells(22).Value 'Grabs input initial Rc from table, added by JBB
                                    Bioproperties.RcMaxInput = CoverSelectionGrid.Rows(CoverIndex).Cells(23).Value 'Grabs input maximum Rc from table, added by JBB
                                    'Bioproperties.MaxIterInput = CoverSelectionGrid.Rows(CoverIndex).Cells(24).Value 'Grabs input stability itteration limit from table, added by JBB
                                    Bioproperties.FcVIInput = DirectCast([Enum].Parse(GetType(FcVI), CoverSelectionGrid.Rows(CoverIndex).Cells(24).Value.Replace(" ", "_")), FcVI) 'Grabs the VI for Fc copied by JBB from code above for cover
                                    Bioproperties.MinVIInput = CoverSelectionGrid.Rows(CoverIndex).Cells(25).Value 'Grabs input Fc Min NDVI from table, added by JBB
                                    Bioproperties.MaxVIInput = CoverSelectionGrid.Rows(CoverIndex).Cells(26).Value 'Grabs input Fc Max NDVI from table, added by JBB
                                    Bioproperties.ExpVIInput = CoverSelectionGrid.Rows(CoverIndex).Cells(27).Value 'Grabs input Fc NDVI exponent from table, added by JBB
                                    'if some classes come with tmeprature = 0 , ignore '<-- Refering to TIR temp, in deg C JBB
                                    If TemperaturePixels.GetValue(Col, Row) = 0.0 Then
                                        Exit Sub
                                    End If
                                    Dim LAI As Single = 0
                                    If LAIIndex > -1 Then LAI = LAIPixels.GetValue(Col, Row) 'If LAI image exists then grab LAI from it JBB
                                    Dim ViewZenith As Single = 0 'This should be in degrees, but it was accidentally used in the clumping function as though it were radians JBB
                                    If ZenithIndex > -1 Then ViewZenith = Math.Abs(ZenithPixels.GetValue(Col, Row) - 65) 'If a zenith angle image exists then get zenith from it JBB <-- This should be in degrees though it was used in the clumping function as radians in some places JBB
                                    Dim AirTemperature As Single = 0 'in k
                                    If WeatherGridIndex.Temperature > -1 Then : AirTemperature = WeatherTemperaturePixels.GetValue(Col, Row) : Else : AirTemperature = WeatherPoint.AirTemperature(IWeather) + 273.15 : End If 'If an air temp grid is available then get temp from it, else get tabular temp and convert to K JBB
                                    Dim Rs As Single = 0 ' in W/m2
                                    If WeatherGridIndex.Radiation > -1 Then : Rs = WeatherRadiationPixels.GetValue(Col, Row) : Else : Rs = WeatherPoint.SolarRadiation(IWeather) : End If ' if a solar radiation grid exists then get Rs from it, else get tabular Rs JBB

                                    Dim WindSpeed As Single = 0 ' in m/s
                                    If WeatherGridIndex.WindSpeed > -1 Then : WindSpeed = WeatherWindSpeedPixels.GetValue(Col, Row) : Else : WindSpeed = WeatherPoint.WindSpeed(IWeather) : End If ' If a wind speed grid exists then get WS from it, else get tabular WS JBB
                                    Dim WindHeight As Single = WeatherPoint.AnemometerReferenceHeight(IWeather) 'grab wind height from form JBB
                                    '**** added by JBB for wind height
                                    Dim XfAnemometer As Single = WeatherPoint.xFetchDistanceAnemometer(IWeather) 'Grabs weather station fetch distance
                                    Dim XfCover As Single = WeatherPoint.xFetchDistanceCover(IWeather) 'Grabs cover fetch distance
                                    Dim ZIBLforced As Single = WeatherPoint.xIBLHeightOptional(IWeather) 'Grabs forced internal boundary layer Height
                                    Dim HcRegional As Single = WeatherPoint.xRegionalCoverHeight(IWeather) 'Grabs Regional Cover Height
                                    '********
                                    Dim TemperatureHeight As Single = WeatherPoint.AirTemperatureReferenceHeight(IWeather) 'grab temp height from form JBB
                                    Dim AvailableEnergy As Single = WeatherPoint.DailyAvailableEnergy(IWeather) 'grab input daily Rn - G added by JBB
                                    Dim WindSpeed2m As Single = calcU2(WindSpeed, WindHeight, 1) ' kept as it is, this option is kept for future applications if needed'<--Gridded wind is probably usually 2m JBB
                                    Dim DailyReferenceET As Single = 0 ' in mm/day
                                    If WeatherGridIndex.ETDailyReference > -1 Then : DailyReferenceET = WeatherETDailyReferencePixels.GetValue(Col, Row) : Else : DailyReferenceET = WeatherPoint.ETDailyReference(IWeather) : End If 'If a daily ETo image exists then get ETo from it, else get tabular ETo JBB
                                    Dim InstantaneousET As Single = 0 ' in mm/day
                                    If WeatherGridIndex.ETInstantaneous > -1 Then : InstantaneousET = WeatherETInstantaneousPixels.GetValue(Col, Row) : Else : InstantaneousET = WeatherPoint.ETInstantaneous(IWeather) : End If 'If an instantaneous ETo image exists then get ETo from it, else get tabular EToinst JBB
                                    Dim Pressure As Single = 0 ' in kPa if from tables or Pa if from gridded and then converted to mb
                                    If WeatherGridIndex.Pressure > -1 Then : Pressure = WeatherPressurePixels.GetValue(Col, Row) * 0.01 : Else : Pressure = WeatherPoint.AtmosphericPressure(IWeather) * 10 : End If 'If a gridded pressure file exists then grab pressure from it and convert from Pa to mb, else  grab the tabular pressure and convert form kPa to mb JBB
                                    Dim PeakLAI As Single = -999 'added by JBB for peak LAI for Fg
                                    Dim PeakHc As Single = -999 'added by JBB for peak Hc
                                    Dim PeakFc As Single = -999 'added by JBB for peak Fc

                                    If FgMethod = Functions.FgMethod.Using_Past_LAI Then
                                        If LAICount > 0 Then PeakLAI = LAIPeakPastPixels.GetValue(Col, Row) 'added by JBB for peak lai for Fg
                                    End If
                                    If HcMethod = Functions.HcMethod.Using_Past_Hc Then
                                        If HcCount > 0 Then PeakHc = HcPeakPastPixels.GetValue(Col, Row) 'added by JBB for peak Hc
                                    End If

                                    Dim Ea As Single = 0 ' in kPa and then needs to be converted to mb
                                    If WeatherGridIndex.SpecificHumidity > -1 Then
                                        Dim SpecificHumidity As Single = 0
                                        SpecificHumidity = WeatherHumidityPixels.GetValue(Col, Row) ' (kg/kg)
                                        Ea = calcActualVaporPressure(SpecificHumidity, Pressure) 'Calcs vapor pressure in mb from specific humidity and barometric pressure JBB
                                    Else
                                        Ea = WeatherPoint.ActualVaporPressure(IWeather) * 10 'If no humidity image then grabs tabular Ea in kPa and converts to mb JBB
                                    End If

                                    Dim Tr As Single = 1 - 373.15 / AirTemperature 'Eq 3.24a from Brutseart 1982
                                    Dim Es As Single = 1013.25 * Math.Exp(13.3185 * Tr - 1.976 * Tr ^ 2 - 0.6445 * Tr ^ 3 - 0.1299 * Tr ^ 4) 'Eq 3.24a from Brutseart 1982 JBB

                                    'Dim Albedo As Single = calcAlbedo(Red, NIR) ' Brest and Goward 1987 for vegetation, they say 0.526*RED + 0.474*NIR for non-veg. They say Veg is when NIR/RED >2.0. JBB
                                    Dim Albedo As Single = calcAlbedo(Red, NIR, Green) ' Brest and Goward 1987 for vegetation, Updated to include Green by JBB
                                    Dim SAVI As Single = calcSAVI(Red, NIR) 'Calc VI's JBB
                                    Dim oSAVI As Single = calcOSAVI(Red, NIR) 'Calc VI's JBB
                                    Dim NDVI As Single = calcNDVI(Red, NIR) 'Calc VI's JBB
                                    Dim NDWI As Single = calcNDWI(MIR1, NIR) 'Calc VI's JBB

                                    '#####NOTE LAI PEAK IS APPLIED AFTER THIS
                                    Dim FractionOfCover As Single
                                    If FcIndex > -1 Then FractionOfCover = Limit(FcPixels.GetValue(Col, Row), FcoverMin, 0.95) 'If fc image exists then grab fc from it JBB
                                    If FcIndex < 0 Then 'logic added by JBB
                                        'FractionOfCover = calcFc(Cover, NDVI, SAVI, LAI) 'Uses ' From Choudhury et al 1994 as cited by Li et al 2005 eq. 1 JBB
                                        FractionOfCover = calcFc(Cover, NDVI, SAVI, LAI, Bioproperties) 'Uses ' From Choudhury et al 1994 as cited by Li et al 2005 eq. 1 JBB Updated by JBB 1/31/2018
                                    End If 'logic added by JBB
                                    If FcMethod = Functions.FcMethod.Using_Past_Fc Then 'added byJBB for peak FC
                                        If FcCount > 0 Then PeakFc = Limit(FcPeakPastPixels.GetValue(Col, Row), FcoverMin, 0.95) 'added byJBB for peak FC
                                    End If 'added byJBB for peak FC

                                    'this is to avoid some pixels with negative or zero values
                                    If SAVI <= 0 Or NDVI <= 0 Or oSAVI <= 0 Then
                                        If NegFlag = False Then 'added by JBB to allow Landsat 7 images to run
                                            NegFlag = True 'prevents this from showing up many times
                                            'MsgBox("A negative vegetation index was computed, check output for invalid pixels, e.g. NDVI, SAVI, oSAVI ≤ 0")'moved this out of the calcs to be less obtrusive
                                        End If
                                        'Exit Sub
                                    End If
                                    Select Case ImageSource
                                        Case Functions.ImageSource.Airborne
                                            'If LAIIndex < 0 Then LAI = calcLAILandSAT(NDVI, NDWI, SAVI, oSAVI, Cover, FractionOfCover)'Commented out by JBB
                                            If LAIIndex < 0 Then LAI = calcLAIAir(NDVI, SAVI, oSAVI, Cover, FractionOfCover) 'added by JBB
                                        Case Functions.ImageSource.Landsat
                                            If LAIIndex < 0 Then LAI = calcLAILandSAT(NDVI, NDWI, SAVI, oSAVI, Cover, FractionOfCover)
                                        Case Functions.ImageSource.Unmanned_Aircraft 'Added by JBB 3/1/2018 copied from above for airborne 
                                            If LAIIndex < 0 Then LAI = calcLAIUAV(NDVI, SAVI, oSAVI, Cover, FractionOfCover) 'Added by JBB 3/1/2018 copied from above for airborne 
                                    End Select
                                    If LAI < 0 Then LAI = calcLAITable(Cover, Bioproperties)
                                    LAI = Limit(LAI, LAIMin, 8)

                                    If PeakLAI > 0 Then 'added by JBB to calculate Fraction of Green LAI from LAI and peak LAI
                                        PeakLAI = Limit(PeakLAI, LAIMin, 8)
                                        Dim fg_int As Single = LAI / PeakLAI
                                        fg_int = Limit(fg_int, 0, 1)
                                        Bioproperties.fg = fg_int
                                        LAI = Math.Max(LAI, PeakLAI) 'Now Set LAI = Peak LAI
                                    End If
                                    ' this is in case we have estimated/measured LAI like in Transitional Forest savana
                                    'FractionOfCover = calcFcbasedonLAI(Cover, LAI) 'Commented out by JBB because it was overwriting FractionOfCover to always be Fcmin = 0.0631, since no LAI image was provided JBB

                                    '======== Here we estimate vegetation height ============
                                    '#######Note peak Fc is calculated after this!
                                    Dim MaximumCoverHeight As Single = 0
                                    If VegetationHeightIndex > -1 Then 'First priority is an image of Vegetation Height, Probably from LIDAR JBB
                                        MaximumCoverHeight = VegetationHeightPixels.GetValue(Col, Row)
                                    Else
                                        Select Case ImageSource 'Next priority is using a VI, in this case OSAVI JBB
                                            Case Functions.ImageSource.Airborne
                                                MaximumCoverHeight = calcHcAir(NDVI, NDWI, SAVI, oSAVI, FractionOfCover, Cover)
                                            Case Functions.ImageSource.Landsat
                                                MaximumCoverHeight = calcHcLandSAT(NDVI, NDWI, SAVI, oSAVI, FractionOfCover, Cover)
                                            Case Functions.ImageSource.Unmanned_Aircraft 'Added by JBB 3/1/2018 copied from above for airborne and Landsat
                                                MaximumCoverHeight = calcHcUAV(NDVI, NDWI, SAVI, oSAVI, FractionOfCover, Cover) 'Added by JBB 3/1/2018 copied from above for airborne and Landsat
                                        End Select
                                    End If
                                    If MaximumCoverHeight < 0 Then MaximumCoverHeight = calcHcTable(Cover, FractionOfCover, Bioproperties) 'If no cover height is available from an image, then calculate late from Hcmin and Hcmax JBB
                                    'This made the crop shorter if it was > wind height - 0.5 m, a more appropriate method is to estimate the wind height over the canopy using the log law. JBB
                                    'If MaximumCoverHeight > 0 Then MaximumCoverHeight = Limit(MaximumCoverHeight, 0.1, Math.Max(WindHeight - 0.5, TemperatureHeight - 0.5)) 'If cover height is greater than wind height - 0.5 m then it is limited to Zwind-0.5m commented out by JBB
                                    '****Added by JBB *****Assumes Neutral Conditions*****
                                    Dim HWindFetch As Single = WeatherPoint.WindFetchHeight(IWeather) 'Grabs Input Fetch Height near Weather Station
                                    Dim WindHeightAct As Single = WindHeight 'Preserve Acutal Wind Height JBB
                                    Dim TempHeightAct As Single = TemperatureHeight 'Preserve Actual Temperature Height JBB (note the height is adjusted below, but the temperature measurement isn't, unlike the windspeed, the assumpition is that a similar air temp would be measured at the same height over any actively transpiring canopy (perhaps a bad assumption) added by JBB

                                    If PeakHc > 0 Then
                                        MaximumCoverHeight = Math.Max(MaximumCoverHeight, PeakHc) 'Makes Cover Height the Maximum of past (if input) and current added by JBB
                                    End If

                                    If PeakFc > 0 Then
                                        FractionOfCover = Math.Max(FractionOfCover, PeakFc) 'Makes fraction of cover the peak of max past (if input) and current added by JBB
                                    End If

                                    If MaximumCoverHeight > 0 Then MaximumCoverHeight = Limit(MaximumCoverHeight, 0.1, 116) 'Limits the crop height to 0.1m to 116m, the height of the tallest known redwood JBB, this logic step of Hc >0 is overwritten later, where all heights are limited to be >=0.1
                                    'If MaximumCoverHeight > WindHeight - 1.0 Then WindHeight = MaximumCoverHeight + 1.0 'if the maximum cover height is too close to the wind height, or taller then adjust the wind height commented out by JBB

                                    'Logic added on 8/10/2016 to allow for measurements over the site of interest
                                    If WindAdjustment = WindAdjustment.No_Adjustment Then
                                        WindHeight = WindHeightAct
                                        TemperatureHeight = TempHeightAct
                                    Else
                                        WindHeight = (WindHeightAct - HWindFetch) + MaximumCoverHeight 'Added by JBB maintains wind speed height at same height over the canopy as the original measurement. added by JBB 
                                        TemperatureHeight = (TempHeightAct - HWindFetch) + MaximumCoverHeight ' Added by JBB maintains same dist above canopy as oringinal measurement,  but temp is not adjusted as wind speed is added by JBB
                                    End If

                                    '****Added by JBB *****
                                    '===== end of vegetation height calculations=========

                                    Dim Rn As Single = 0 'Declare energy balance coefficients JBB
                                    Dim G As Single = 0
                                    Dim Rah As Single = 0
                                    Dim H As Single = 0
                                    Dim ET As Single = 0
                                    'Dim Taerodynamic As Double = 0
                                    Dim Taerodynamic As Single = 0 'Modified by JBB 2/1/2018
                                    Dim LE As Single = 0
                                    Dim PT_Out As Single = 0 'added by JBB
                                    Dim Rcanopy_Out As Single = 0 'added by JBB for PM
                                    Dim Lo_Out As Single = 0 'added by JBB
                                    Dim U_Out As Single = 0 'added by JBB
                                    Dim LEc_Out As Single = 0 'added by JBB
                                    Dim LEs_Out As Single = 0 'added by JBB
                                    Dim fg_Out As Single = 0 'added by JBB
                                    Dim Ustar_Out As Single = 0 'added by JBB
                                    Dim Rnc_Out As Single = 0 ' added by JBB
                                    Dim Rns_Out As Single = 0 'added by JBB
                                    Dim Tc_Out As Single = 0 'added by JBB
                                    Dim Ts_Out As Single = 0 'added by JBB
                                    Dim Tac_Out As Single = 0 'added by JBB
                                    Dim To_Out As Single = 0 'added by JBB
                                    Dim Albedo_Out As Single = 0 'added by JBB

                                    Select Case EnergyBalanceModel
                                        Case EnergyBalance.One_Layer '************I did not review the one layer code JBB*******************
                                            ''This was the code originally in SETMI, Burdette rewrote it on 12/04/2017 following Chavez et al. 2005
                                            'Dim Zom As Single = 0.123 * MaximumCoverHeight
                                            'Dim Zoh As Single = 0.1 * Zom
                                            'Dim D As Single = 0.67 * MaximumCoverHeight
                                            'Dim Cp As Single = 1157.46
                                            'Dim Ustarn As Single = Math.Max(WindSpeed * K / Math.Log((WindHeight - D) / Zom), 0.01)
                                            'Select Case Cover
                                            '    Case Functions.Cover.Corn, Functions.Cover.Soybean
                                            '        Taerodynamic = 273.15 + calcTaerodynamic(Temperature - 273.15, AirTemperature - 273.15, WindSpeed, LAI)
                                            '    Case Functions.Cover.Tamarisk, Functions.Cover.Dead_Tamarisk
                                            '        'Taerodynamic = 273.15 + calctaerodynamic((temp - 273.15), TaM(0, DD), uM(0, DD), lai)
                                            '        Taerodynamic = Math.Min(273.15 + 1.48 * (Temperature - 273.15) - 0.48 * (AirTemperature - 273.15) - 2.47 * LAI - 0.68 * WindSpeed + 8.32, 273.15)
                                            '        Dim Tair_Tamarisk As Single = Math.Min(273.15 + (26.6 * Math.Log(Taerodynamic - 273.15) - 62.9), 273.15)
                                            '    Case Functions.Cover.Alfalfa
                                            '    Case Functions.Cover.Cotton
                                            '    Case Functions.Cover.Grass
                                            '    Case Functions.Cover.Wheat
                                            '    Case Functions.Cover.Agriculture
                                            'End Select
                                            'Rn = calcRn(SAVI, NDVI, Albedo, AirTemperature, Temperature, Rs, Ea, Esoil, Eveg, Sigma, RecordDate.Month, Cover, FractionOfCover)
                                            'G = calcG(Rn, LAI, Cover)
                                            'Dim Rahn As Single = Math.Log((WindHeight - D) / Zom) * Math.Log((WindHeight - D) / Zoh) / WindSpeed / K ^ 2
                                            'Dim Resistances As Resistances_Output
                                            'If (Taerodynamic - AirTemperature) > 0 Then
                                            '    Resistances = calcRaMO_Unstable_OLM(AirTemperature, Taerodynamic, WindSpeed, MaximumCoverHeight, Rahn, Ustarn, Zom, Zoh, D, K, Cp, WindHeight, Gravity)
                                            '    'rahYmYh = RaMO_Hatim(Ta, Taerodynamic, u, hc)
                                            'Else
                                            '    Resistances = calcRaMO_Stable_OLM(AirTemperature, Taerodynamic, WindSpeed, MaximumCoverHeight, Rahn, Ustarn, Zom, Zoh, D, K, Cp, WindHeight, Gravity)
                                            '    'rahYmYh = RaMO_stable_Hatim(Ta, Taerodynamic, u, hc)
                                            'End If
                                            'Rah = Resistances.Rah
                                            'If Resistances.Rah <> 0 Then
                                            '    H = Cp * (Taerodynamic - AirTemperature) / Resistances.Rah
                                            '    If (H + G) < Rn Then LE = Rn - H - G
                                            'End If

                                            '****************************************************************************************************
                                            '****************************************************************************************************
                                            '***                     Start JBB Code                                                           *** 
                                            '****************************************************************************************************
                                            '****************************************************************************************************

                                            '***Equations are from Chavez et al. 2005, Comparing aircraft-based remotely sensed energy balance fluxes with eddy covariance tower data using heat flux source area functions. J. of Hydrometeorology 6:923-940, unless otherwise specified

                                            'This is how Chavez et al. 2005 did it, I am following the TSEB methodology
                                            'Dim Zom As Single = 0.123 * MaximumCoverHeight 'eq. 4
                                            'Dim Zoh As Single = 0.1 * Zom 'eq. 5
                                            'Dim D As Single = 0.67 * MaximumCoverHeight 'eq.6

                                            '---Copied from TSEB code:
                                            '**** This method should be good for many agricultural crops, it is the default
                                            Dim Zom As Single = MaximumCoverHeight / 8 'From Li et al 2005 p 888, 8 is within the range given by Brutseart on pp 113 and 114 JBB
                                            Dim Zoh As Single = Zom / 7 'From Li et al 2005 p 888, FAO56 uses Zoh=0.1Zom JBB
                                            Dim D As Single = 2 / 3 * MaximumCoverHeight 'From Li et al 2005 p 888, same as Brutseart 1982 Eq. 5.4 JBB
                                            'Added by JBB to allow for Hc/Wc to be an input ratio rather than computed ratio 2/1/2018
                                            Dim Drat As Single
                                            If ClumpDMethod = ClumpingD.Input_Height_to_Width_Ratio Then
                                                Drat = Bioproperties.ClumpD
                                            Else 'default is use Wc
                                                Drat = MaximumCoverHeight / Bioproperties.Wc
                                            End If
                                            'End Hc/Wc additions
                                            '**no limits are needed on Zom or Zoh because the limits on hc, lai, and fc already impose a limit JBB
                                            If FractionOfCover <= FcoverMin Or LAI <= LAIMin Then ' mainly for bare soils
                                                'Zom = 0.1 / 8
                                                Zom = 0.005  'recent update from Martha's code'Brutseart 2005 says 0.0001-0.0005 for mud flats and 0.008 is the smallest for short grass, this is probably okay
                                                D = 0
                                                Zoh = Zom / 7 'added by JBB because otherwise Zoh is from the equations above!
                                            ElseIf Cover = Functions.Cover.Water Then 'for water surface
                                                Zom = 0.00035 'Brutseart 2005 says 0.0001-0.0006 for large water surfaces
                                                D = 0
                                                Zoh = Zom / 7 'added by JBB because otherwise Zoh is from the equations above!
                                                'Else
                                                '**** The logic below was created to make a class of broad leaf plants, it should be verified by the user before selecting one such cover
                                            ElseIf Cover = Functions.Cover.Cottonwood Or Cover = Functions.Cover.Desert_Shrubs Or Cover = Functions.Cover.Mesquite Or Cover = Functions.Cover.Upland_Bushes Or Cover = Functions.Cover.Citrus Or Cover = Functions.Cover.Guava Or Cover = Functions.Cover.Coconut Or Cover = Functions.Cover.Vineyards Or Cover = Functions.Cover.Graviola Or Cover = Functions.Cover.Sapoti Or Cover = Functions.Cover.Acerola Or Cover = Functions.Cover.Sabia Or Cover = Functions.Cover.Mango Or Cover = Functions.Cover.Cashew_Giant Or Cover = Functions.Cover.Cashew_Early Or Cover = Functions.Cover.Mata Or Cover = Functions.Cover.Papaya Or Cover = Functions.Cover.Lemon Or Cover = Functions.Cover.Tangerine Or Cover = Functions.Cover.Orange Or Cover = Functions.Cover.Passion_Fruit Or Cover = Functions.Cover.Banana_Apple Or Cover = Functions.Cover.Mixed_Forest Or Cover = Functions.Cover.Transitional_Forest Or Cover = Functions.Cover.Agro_Forestry_Areas Or Cover = Functions.Cover.Broad_Leaved_Forest Or Cover = Functions.Cover.Transitional_Woodland_Shrub Or Cover = Functions.Cover.Sparsely_vegetated_areas Then
                                                '***** The code below is based on PyTSEB and may be subject to their open source agreements: 
                                                '       It is Copyrighted by Hector Nieto and contributors. It is covered by the GNU General Public License see https://github.com/hectornieto/pyTSEB
                                                'JBB pulled equations from the source literature, Schaudt and Dickinson 2000
                                                'Dim LambdaRaupach As Single = FractionOfCover * MaximumCoverHeight / Bioproperties.Wc
                                                Dim LambdaRaupach As Single = FractionOfCover * Drat
                                                Zom = SchaudtZom(MaximumCoverHeight, LAI, LambdaRaupach)
                                                D = SchaudtD(MaximumCoverHeight, LAI, LambdaRaupach)
                                                Zoh = Zom / 7 'following  Li et al 2005 and pyTSEB (kinda)
                                                '**** End material from PyTESB
                                                '**** The logic below was created to make a class of needle leaf plants, it should be verified by the user before selecting one such cover
                                            ElseIf Cover = Functions.Cover.Conifer Or Cover = Functions.Cover.Dead_Tamarisk Or Cover = Functions.Cover.Tamarisk Or Cover = Functions.Cover.Coniferous_Forest Or Cover = Functions.Cover.Eastern_Red_Cedar Then 'Added by JBB for Micheal Neale's Halsey Project
                                                '***** The code below is based on PyTSEB and may be subject to their open source agreements: 
                                                '       It is Copyrighted by Hector Nieto and contributors. It is covered by the GNU General Public License see https://github.com/hectornieto/pyTSEB
                                                'JBB pulled equations from the source literature, Schaudt and Dickinson 2000
                                                'Dim LambdaRaupach As Single = 2 / Math.PI * FractionOfCover * MaximumCoverHeight / Bioproperties.Wc
                                                Dim LambdaRaupach As Single = 2 / Math.PI * FractionOfCover * Drat
                                                Zom = SchaudtZom(MaximumCoverHeight, LAI, LambdaRaupach)
                                                D = SchaudtD(MaximumCoverHeight, LAI, LambdaRaupach)
                                                Zoh = Zom / 7 'following  Li et al 2005 and pyTSEB (kinda)
                                            Else
                                                'This is the way SETMI did D and Zom, it looked like a modification of Raupach 1994, I commented it out in favor of the methods above JBB

                                                'Dim ufact As Single = 0.36 - 0.264 * Math.Exp(-15.1 * cd * LAI) 'Looks kind of like Raupach 1994, but not quite JBB
                                                'Dim xn As Single = cd * LAI / (2 * ufact ^ 2)
                                                'Dim dispdh As Single = 0.7 - (1 / (5 * xn) * (1 - Math.Exp(-3.3 * xn)))
                                                'If dispdh < 0 Then dispdh = 0
                                                'Dim Zomdh As Single = (1 - dispdh) * Math.Exp(-0.4 / ufact)
                                                'D = dispdh * MaximumCoverHeight
                                                'Zom = Math.Max(Zomdh * MaximumCoverHeight, 0.005) 
                                            End If

                                            '***Moved up by JBB
                                            Dim PixelLatitude As Single = 0
                                            Dim PixelLongitude As Single = 0
                                            TransformRaster2.PixelToMap(InRasterCursor(M).TopLeft.X + Col, InRasterCursor(M).TopLeft.Y + Row, PixelLongitude, PixelLatitude)
                                            '***Moved up by JBB
                                            '********added by JBB **************
                                            'Cacluate dact and ZomAct for wind fetch, which is assumed to be grass and can thus use the simplified relationships of the ASCE Std. Eq. JBB <-- Now changed to follow Lie et al for consistency
                                            'Dim DAct As Single = HWindFetch * 0.67 'JBB ASCE Std. Eq. B14
                                            Dim DAct As Single = HWindFetch * 2 / 3 'From Li et al 2005 p 888, same as Brutseart 1982 Eq. 5.4 JBB
                                            'Dim ZomAct As Single = 0.123 * HWindFetch 'JBB ASCE Std. Eq. B14
                                            Dim Zomact As Single = HWindFetch / 8 'From Li et al 2005 p 888, 8 is within the range given by Brutseart on pp 113 and 114 JBB
                                            '****Allen and Wright 1997
                                            Dim Cr As Single = 0.2 'Allen and Wright 1997 p.30
                                            Dim OmegaEarth As Single = Math.PI * 2 / 86400 'speed of earth rad/s
                                            Dim PhiRad As Single = PixelLatitude * Math.PI / 180 'latitude in rad
                                            Dim ZIBLw As Single = 0
                                            Dim ZIBLv As Single = 0
                                            If ZIBLforced > 0 Then 'This allows us to force the IBL to be at a given height, i.e. a mixing height (see METRIC Methods Paper, Allen et al...)
                                                ZIBLw = ZIBLforced
                                                ZIBLv = ZIBLforced
                                            Else
                                                Dim ZIBLx As Single = Cr * WindSpeed * K / (Math.Log((WindHeightAct - DAct) / Zomact) * 2 * OmegaEarth * Math.Abs(Math.Sin(PhiRad))) 'Allen and Wright 1997 Eq 17
                                                ZIBLw = DAct + 0.33 * (Zomact ^ 0.125) * (XfAnemometer ^ 0.875) 'Allen and Wright 1997 Eq. 14
                                                ZIBLv = D + 0.33 * (Zom ^ 0.125) * (XfCover ^ 0.875) 'Allen and Wright 1997 Eq. 14
                                                If ZIBLw > ZIBLx Then ZIBLw = ZIBLx 'applies upper bound on Zibl
                                                If ZIBLv > ZIBLx Then ZIBLv = ZIBLx
                                            End If
                                            Dim Dr As Single = HcRegional * 2 / 3
                                            Dim Zomr As Single = HcRegional / 8
                                            '***End Allen and Wright 1997
                                            Dim y_m_act As Single = 0
                                            Dim y_m_new As Single = 0
                                            Dim WindSpeed_Meas As Single = WindSpeed
                                            '***************End added by JBB, more related code is in the stability loop

                                            'This section is slightly modified from the TSEB
                                            'Compute adjusted wind speed, neutral regardless of stability
                                            '**************************************************************************************************
                                            '*********  This would need to be added in the iteration loop if stability was adjusted!!!!  ******
                                            '**************************************************************************************************
                                            Select Case WindAdjustment
                                                Case WindAdjustment.No_Adjustment
                                                    WindSpeed = WindSpeed 'No adjustment
                                                Case Else 'No stability terms included - like neutral conditions - common in reference ET calculation adjustments JBB
                                                    ' WindSpeed = WindSpeed_Meas * ((Math.Log((WindHeight - D) / Zom)) / (Math.Log((WindHeightAct - DAct) / Zomact))) 'Added by JBB Adjusts windspeed measured over some other surface to equivalent measured over Cover of Interest Following ASCE Std. Eq. B14 This method is erroneous because the friction velocity is not the same over the two different surfaces
                                                    If WindHeightAct < ZIBLw And WindHeight < ZIBLv Then 'Both heights below IBLs use Allen and Wright 1997 Eq. 13
                                                        WindSpeed = WindSpeed_Meas * (Math.Log((ZIBLw - DAct) / Zomact) * Math.Log((ZIBLv - Dr) / Zomr) * Math.Log((WindHeight - D) / Zom)) / (Math.Log((WindHeightAct - DAct) / Zomact) * Math.Log((ZIBLw - Dr) / Zomr) * Math.Log((ZIBLv - D) / Zom)) 'Allen and Wright 1997 Eq. 13
                                                    ElseIf WindHeightAct >= ZIBLw And WindHeight < ZIBLv Then 'Meas above IBL, use Allen and Wright 1997 Eq. 18
                                                        WindSpeed = WindSpeed_Meas * (Math.Log((ZIBLv - Dr) / Zomr) * Math.Log((WindHeight - D) / Zom)) / (Math.Log((WindHeightAct - Dr) / Zomr) * Math.Log((ZIBLv - D) / Zom)) 'Allen and Wright 1997 Eq. 18
                                                    ElseIf WindHeight < ZIBLw And WindHeight >= ZIBLv Then 'Crop wind height above IBL, use Allen and Wright 1997 Eq. 19
                                                        WindSpeed = WindSpeed_Meas * (Math.Log((ZIBLw - DAct) / Zomact) * Math.Log((WindHeight - Dr) / Zomr)) / (Math.Log((WindHeightAct - DAct) / Zomact) * Math.Log((ZIBLw - Dr) / Zomr)) 'Allen and Wright 1997 Eq. 19
                                                    Else 'Both heights above IBLs, use Allen and Wright 1997 Eq 20
                                                        WindSpeed = WindSpeed_Meas * (Math.Log((WindHeight - Dr) / Zomr)) / (Math.Log((WindHeightAct - Dr) / Zomr)) 'Allen and Wright 1997 Eq. 20
                                                    End If
                                            End Select
                                            '**********End Code Added by JBB ***************                                            
                                            '---End copied from TSEB Code

                                            'Initial estimate of Rah assuming neutral conditions
                                            Rah = Math.Log((WindHeight - D) / Zom) * Math.Log((WindHeight - D) / Zoh) / (WindSpeed * K * K) 'Eq. 3 - Assume Wind Height is close enough to Temp Height
                                            'Initial estiamte of Ustar assuming neutral conditions
                                            Dim Ustar As Single = WindSpeed * K / Math.Log((WindHeight - D) / Zom) 'Eq. 11
                                            'Initial estimate of Taero
                                            Taerodynamic = 0.534 * (Temperature - 273.15) + 0.39 * (AirTemperature - 273.15) + 0.224 * LAI - 0.192 * WindSpeed + 1.67 + 273.15 'Eq. 26, Looks like it is in C
                                            Dim WW = calcW(AirTemperature, Ea, Pressure, 0, 0)

                                            Dim CpaRhoa As Single = WW.CpRho 'Following methods already used for TSEB
                                            H = CpaRhoa * (Taerodynamic - AirTemperature) / Rah 'Eq. 2
                                            Dim Lmo As Single
                                            Dim Rah_prev As Single  'for iteration
                                            Dim Psih As Single = 0
                                            Dim Psim As Single = 0
                                            Dim XX As Single = 0
                                            Dim XXh As Single = 0

                                            For ii = 1 To 100 'follow pp. 929-30
                                                Rah_prev = Rah 'save previous Rah
                                                Lmo = -(Ustar ^ 3) * AirTemperature * CpaRhoa / (Gravity * K * H) 'Eq. 8
                                                If Lmo < 0 Then 'If stable
                                                    XX = (1 - 16 * (WindHeight - D) / Lmo) ^ 0.25 'eq. 10
                                                    'XXh = (1 - 16 * (TemperatureHeight - D) / Lmo) ^ 0.25 'eq. 10
                                                    Psih = 2 * Math.Log((1 + XX ^ 2) / 2) 'eq. 9
                                                    Psim = 2 * Math.Log((1 + XX) / 2) + Math.Log((1 + (XX ^ 2)) / 2) - 2 * Math.Atan(XX) + Math.PI / 2 'Eq. 13 and Bruseart 1982 Eq. 4.50
                                                Else
                                                    H = -9999 'doesn't work for stable
                                                    Exit For
                                                End If
                                                Ustar = WindSpeed * K / (Math.Log((WindHeight - D) / Zom) - Psim * ((WindHeight - D) / Lmo) + Psim * Zom / Lmo) 'Eq. 12
                                                Rah = (Math.Log((WindHeight - D) / Zoh) - Psih * ((WindHeight - D) / Lmo) + Psih * Zoh / Lmo) / (Ustar * K) 'Eq. 7, this equiation seems odd and doesn't use air temp meas. height
                                                'Update Taerodynamic here if wind speed adjustement is modified to include stability
                                                H = CpaRhoa * (Taerodynamic - AirTemperature) / Rah 'Eq. 2
                                                If Math.Abs(Rah - Rah_prev) < 0.0001 Then Exit For
                                                If ii = 100 Then H = -8888 'error if no convergence
                                            Next
                                            Dim Mnth As Integer = Month(RecordDate)
                                            Esoil = Bioproperties.EmissSoilTIR
                                            Eveg = Bioproperties.AlphaLeafTIR
                                            Rn = calcRn(SAVI, NDVI, Albedo, AirTemperature, Temperature, Rs, Ea, Esoil, Eveg, Sigma, Mnth, Cover, FractionOfCover) 'Eq. 16, uses fraction of cover previously difined
                                            G = ((0.3324 - 0.024 * LAI) * (0.8155 - 0.3032 * Math.Log(LAI))) * Rn 'Eq. 25
                                            LE = Rn - G - H 'Eq. 20

                                            '***Copied from TSEB
                                            Lo_Out = Lmo
                                            U_Out = WindSpeed
                                            Ustar_Out = Ustar 'added by JBB
                                            To_Out = Taerodynamic - 273.15 'added by JBB
                                            Albedo_Out = Albedo 'added by JBB

                                        '****************************************************************************************************
                                        '****************************************************************************************************
                                        '***                     End JBB Code                                                             *** 
                                        '****************************************************************************************************
                                        '****************************************************************************************************
                                        Case EnergyBalance.Two_Source '**********************Start TSEB Code**************JBB
                                            Dim Cveg As Single = 0.5 'Cveg is related to leaf angle distribution, note: 0.5 is spherical JBB
                                            'Select Case Cover 'commented out by JBB since there are many other places where a spherical LAD is assumed. JBB
                                            '    Case Cover.Corn
                                            '        Cveg = 0.5 '<-- 0.5 is good for Corn, which is near spherical JBB
                                            '    Case Cover.Soybean
                                            '        'Cveg = 0.5 '<--Commented out by JBB as Soy should not be 1 since it is heliotrophic JBB
                                            '        Cveg = 1 'Added by JBB, Soy should not be 1 since it is heliotrophic JBB
                                            'End Select

                                            Dim L_MO As Single = -80 'Seed Monin Obukov Length, assumes unstable situation JBB

                                            '***Moved up by JBB
                                            Dim PixelLatitude As Single = 0
                                            Dim PixelLongitude As Single = 0
                                            TransformRaster2.PixelToMap(InRasterCursor(M).TopLeft.X + Col, InRasterCursor(M).TopLeft.Y + Row, PixelLongitude, PixelLatitude)
                                            '***Moved up by JBB

                                            ' note that minimum canopy height is set to 0.1 even for bare soil
                                            MaximumCoverHeight = Math.Max(MaximumCoverHeight, 0.1)
                                            '**** This method should be good for many agricultural crops, it is the default
                                            Dim Zom As Single = MaximumCoverHeight / 8 'From Li et al 2005 p 888, 8 is within the range given by Brutseart on pp 113 and 114 JBB
                                            Dim Zoh As Single = Zom / 7 'From Li et al 2005 p 888, FAO56 uses Zoh=0.1Zom JBB
                                            Dim D As Single = 2 / 3 * MaximumCoverHeight 'From Li et al 2005 p 888, same as Brutseart 1982 Eq. 5.4 JBB
                                            '**no limits are needed on Zom or Zoh because the limits on hc, lai, and fc already impose a limit JBB
                                            'Added by JBB to allow for Hc/Wc to be an input ratio rather than computed ratio 2/1/2018
                                            Dim Drat As Single
                                            If ClumpDMethod = ClumpingD.Input_Height_to_Width_Ratio Then
                                                Drat = Bioproperties.ClumpD
                                            Else 'default is use Wc
                                                Drat = MaximumCoverHeight / Bioproperties.Wc
                                            End If
                                            'End Hc/Wc additions

                                            ' Other method for getting Zom, Zoh, and D ' Martha's code '<--- It actually doesn't calculate Zoh, so that is still based on Li et al 2005
                                            Dim cd As Single = 0.2
                                            '***********************
                                            '***********************
                                            '***********************
                                            '***********************
                                            '***********************TESTING ONLY DELETE!!!!
                                            'LAI = 6.06943778548838
                                            'MaximumCoverHeight = 2.34734610641277
                                            'MaximumCoverHeight = 0.1
                                            'LAI = 0.1
                                            'FractionOfCover = FcoverMin
                                            'AirTemperature = 24.5 + 273.15
                                            'Temperature = 23 + 273.15
                                            'L_MO = 80
                                            ''***********************
                                            ''***********************
                                            ''***********************
                                            ''***********************

                                            If FractionOfCover <= FcoverMin Or LAI <= LAIMin Then ' mainly for bare soils
                                                'Zom = 0.1 / 8
                                                Zom = 0.005  'recent update from Martha's code'Brutseart 2005 says 0.0001-0.0005 for mud flats and 0.008 is the smallest for short grass, this is probably okay
                                                D = 0
                                                Zoh = Zom / 7 'added by JBB because otherwise Zoh is from the equations above!
                                            ElseIf Cover = Functions.Cover.Water Then 'for water surface
                                                Zom = 0.00035 'Brutseart 2005 says 0.0001-0.0006 for large water surfaces
                                                D = 0
                                                Zoh = Zom / 7 'added by JBB because otherwise Zoh is from the equations above!
                                                'Else
                                                '**** The logic below was created to make a class of broad leaf plants, it should be verified by the user before selecting one such cover
                                            ElseIf Cover = Functions.Cover.Cottonwood Or Cover = Functions.Cover.Desert_Shrubs Or Cover = Functions.Cover.Mesquite Or Cover = Functions.Cover.Upland_Bushes Or Cover = Functions.Cover.Citrus Or Cover = Functions.Cover.Guava Or Cover = Functions.Cover.Coconut Or Cover = Functions.Cover.Vineyards Or Cover = Functions.Cover.Graviola Or Cover = Functions.Cover.Sapoti Or Cover = Functions.Cover.Acerola Or Cover = Functions.Cover.Sabia Or Cover = Functions.Cover.Mango Or Cover = Functions.Cover.Cashew_Giant Or Cover = Functions.Cover.Cashew_Early Or Cover = Functions.Cover.Mata Or Cover = Functions.Cover.Papaya Or Cover = Functions.Cover.Lemon Or Cover = Functions.Cover.Tangerine Or Cover = Functions.Cover.Orange Or Cover = Functions.Cover.Passion_Fruit Or Cover = Functions.Cover.Banana_Apple Or Cover = Functions.Cover.Mixed_Forest Or Cover = Functions.Cover.Transitional_Forest Or Cover = Functions.Cover.Agro_Forestry_Areas Or Cover = Functions.Cover.Broad_Leaved_Forest Or Cover = Functions.Cover.Transitional_Woodland_Shrub Or Cover = Functions.Cover.Sparsely_vegetated_areas Then
                                                '***** The code below is based on PyTSEB and may be subject to their open source agreements: 
                                                '       It is Copyrighted by Hector Nieto and contributors. It is covered by the GNU General Public License see https://github.com/hectornieto/pyTSEB
                                                'JBB pulled equations from the source literature, Schaudt and Dickinson 2000
                                                'Dim LambdaRaupach As Single = FractionOfCover * MaximumCoverHeight / Bioproperties.Wc
                                                Dim LambdaRaupach As Single = FractionOfCover * Drat
                                                Zom = SchaudtZom(MaximumCoverHeight, LAI, LambdaRaupach)
                                                D = SchaudtD(MaximumCoverHeight, LAI, LambdaRaupach)
                                                Zoh = Zom / 7 'following  Li et al 2005 and pyTSEB (kinda)
                                                '**** End material from PyTESB
                                                '**** The logic below was created to make a class of needle leaf plants, it should be verified by the user before selecting one such cover
                                            ElseIf Cover = Functions.Cover.Conifer Or Cover = Functions.Cover.Dead_Tamarisk Or Cover = Functions.Cover.Tamarisk Or Cover = Functions.Cover.Coniferous_Forest Or Cover = Functions.Cover.Eastern_Red_Cedar Or Cover = Functions.Cover.Conifer Then 'Added by JBB for Micheal Neale's Halsey Project
                                                '***** The code below is based on PyTSEB and may be subject to their open source agreements: 
                                                '       It is Copyrighted by Hector Nieto and contributors. It is covered by the GNU General Public License see https://github.com/hectornieto/pyTSEB
                                                'JBB pulled equations from the source literature, Schaudt and Dickinson 2000
                                                'Dim LambdaRaupach As Single = 2 / Math.PI * FractionOfCover * MaximumCoverHeight / Bioproperties.Wc
                                                Dim LambdaRaupach As Single = 2 / Math.PI * FractionOfCover * Drat
                                                Zom = SchaudtZom(MaximumCoverHeight, LAI, LambdaRaupach)
                                                D = SchaudtD(MaximumCoverHeight, LAI, LambdaRaupach)
                                                Zoh = Zom / 7 'following  Li et al 2005 and pyTSEB (kinda)
                                            Else
                                                'This is the way SETMI did D and Zom, it looked like a modification of Raupach 1994, I commented it out in favor of the methods above JBB

                                                'Dim ufact As Single = 0.36 - 0.264 * Math.Exp(-15.1 * cd * LAI) 'Looks kind of like Raupach 1994, but not quite JBB
                                                'Dim xn As Single = cd * LAI / (2 * ufact ^ 2)
                                                'Dim dispdh As Single = 0.7 - (1 / (5 * xn) * (1 - Math.Exp(-3.3 * xn)))
                                                'If dispdh < 0 Then dispdh = 0
                                                'Dim Zomdh As Single = (1 - dispdh) * Math.Exp(-0.4 / ufact)
                                                'D = dispdh * MaximumCoverHeight
                                                'Zom = Math.Max(Zomdh * MaximumCoverHeight, 0.005) 
                                            End If

                                            '********added by JBB **************
                                            'Cacluate dact and ZomAct for wind fetch, which is assumed to be grass and can thus use the simplified relationships of the ASCE Std. Eq. JBB <-- Now changed to follow Lie et al for consistency
                                            'Dim DAct As Single = HWindFetch * 0.67 'JBB ASCE Std. Eq. B14
                                            Dim DAct As Single = HWindFetch * 2 / 3 'From Li et al 2005 p 888, same as Brutseart 1982 Eq. 5.4 JBB
                                            'Dim ZomAct As Single = 0.123 * HWindFetch 'JBB ASCE Std. Eq. B14
                                            Dim Zomact As Single = HWindFetch / 8 'From Li et al 2005 p 888, 8 is within the range given by Brutseart on pp 113 and 114 JBB
                                            '****Allen and Wright 1997
                                            Dim Cr As Single = 0.2 'Allen and Wright 1997 p.30
                                            Dim OmegaEarth As Single = Math.PI * 2 / 86400 'speed of earth rad/s
                                            Dim PhiRad As Single = PixelLatitude * Math.PI / 180 'latitude in rad
                                            Dim ZIBLw As Single = 0
                                            Dim ZIBLv As Single = 0
                                            If ZIBLforced > 0 Then 'This allows us to force the IBL to be at a given height, i.e. a mixing height (see METRIC Methods Paper, Allen et al...)
                                                ZIBLw = ZIBLforced
                                                ZIBLv = ZIBLforced
                                            Else
                                                Dim ZIBLx As Single = Cr * WindSpeed * K / (Math.Log((WindHeightAct - DAct) / Zomact) * 2 * OmegaEarth * Math.Abs(Math.Sin(PhiRad))) 'Allen and Wright 1997 Eq 17
                                                ZIBLw = DAct + 0.33 * (Zomact ^ 0.125) * (XfAnemometer ^ 0.875) 'Allen and Wright 1997 Eq. 14
                                                ZIBLv = D + 0.33 * (Zom ^ 0.125) * (XfCover ^ 0.875) 'Allen and Wright 1997 Eq. 14
                                                If ZIBLw > ZIBLx Then ZIBLw = ZIBLx 'applies upper bound on Zibl
                                                If ZIBLv > ZIBLx Then ZIBLv = ZIBLx
                                            End If
                                            Dim Dr As Single = HcRegional * 2 / 3
                                            Dim Zomr As Single = HcRegional / 8
                                            '***End Allen and Wright 1997
                                            Dim y_m_act As Single = 0
                                            Dim y_m_new As Single = 0
                                            Dim WindSpeed_Meas As Single = WindSpeed
                                            '***************End added by JBB, more related code is in the stability loop

                                            Dim WPT = calcW(AirTemperature, Ea, Pressure, Es, Tr) 'Calculates the PT parameters JBB, the PT weighting coeff is overwritten before it is used later.
                                            Dim SunZenith = calcSunZenith(RecordDate, PixelLatitude, PixelLongitude, ReferenceLongitude)  'in degrees 'Calculates solar zenith for each pixel using the timestamp from the file name JBB
                                            'Dim Clump_Sun = calcZenithClumping(NDVI, SunZenith, ViewZenith, LAI, MaximumCoverHeight, Cveg, FractionOfCover, Cover, Bioproperties) 'Calculates the clumping factors, see Kustas and Norman 2000 JBB
                                            Dim Clump_Sun = calcZenithClumping(NDVI, SunZenith, ViewZenith, LAI, MaximumCoverHeight, Cveg, FractionOfCover, Cover, Bioproperties, ClumpDMethod) 'Calculates the clumping factors, see Kustas and Norman 2000 JBB added by JBB
                                            Dim Ftheta = calcFtheta(Clump_Sun.ClumpView, LAI, Cveg, ViewZenith) 'See  Kustas & Norman 2000 JBB
                                            Dim RnCoefficients = calcRnCoefficients(Rs, Clump_Sun.ClumpSun, LAI, SunZenith, Cover, FractionOfCover, Bioproperties, Pressure) 'Finds TauSolar,TauThermal,AlbSoil,AlbCanopy,and fclear JBB
                                            'Dim Ag = calcAg(NDVI) 'Commented out by JBB
                                            Dim Ag = calcAg2(PixelLongitude, ReferenceTimeZone, RecordDate.DayOfYear, RecordDate.Hour + RecordDate.Minute / 60, Bioproperties) 'Added by JBB Ag is the ratio of G to Rn and is implented here using a solar noon calcluation JBB
                                            Dim TInitial As New TInitial_Output
                                            TInitial = calcTInitial(Temperature, AirTemperature, Ftheta, LAI, FractionOfCover) 'Gets initial Tsoil, Tcanopy, and Tac values JBB, jbb added bare soil component
                                            If Temperature = 273.15 Then
                                                Temperature = Temperature 'This appears to be a debugging artifact JBB
                                            End If
                                            Dim EnergyComponents As New EnergyComponents_Output
                                            Dim Resistances As New Resistances_Output
                                            Dim Stability As Integer = 0
                                            Dim InitialHcanopy As Boolean = True 'Added by JBB to tell TSEB to calculate the initial Hcanopy for the first iteration
                                            For Counter = 1 To 100 'Increased from 50 to 100 by JBB on 8/17/2016
                                                'If Counter > 1 Then InitialHcanopy = False 'The logic statement in calcRaMO_TSM_unstable, wasn't making the change on a global level, the initial estimate is only needed the first iteration of each pixel added by JBB 03/08/2016 This ended up only partially fixing the problem

                                                '********added by JBB **************
                                                If L_MO < 0 Then 'unstable
                                                    'Dim Xact As Single = (1 - 16 * (WindHeightAct - DAct) / L_MO) ^ 0.25 'Similar to Eq. 4.22a in Goudriaan 1977, but the 16 is a 15 in his equation, this is actually from Brutseart P. 70. JBB
                                                    'Dim Xnew As Single = (1 - 16 * (WindHeight - D) / L_MO) ^ 0.25 ' !!!!!!!!! this is actually from Brutseart P. 70, Zt is the height of temp meas, it is not used in Line 913 as it should be!!!!!!!! JBB
                                                    'y_m_act = 2 * Math.Log((1 + Xact) / 2) + Math.Log((1 + Xact ^ 2) / 2) - 2 * Math.Atan(Xact) + Math.PI / 2 'Brutseart 1982 Eq. 4.50 JBB
                                                    'y_m_new = 2 * Math.Log((1 + Xnew) / 2) + Math.Log((1 + Xnew ^ 2) / 2) - 2 * Math.Atan(Xnew) + Math.PI / 2 'Brutseart 1982 Eq. 4.50 JBB

                                                    '*****Using the stability equations of Brutseart 2005, which was learned from PyTSEB, though this wind adjustment is not in PyTSEB
                                                    '****therefore it may be subject to the general public license agreement of PyTSEB see:https://github.com/hectornieto/pyTSEB
                                                    y_m_act = CalcY_MUnstable(WindHeightAct, DAct, L_MO)
                                                    y_m_new = CalcY_MUnstable(WindHeight, D, L_MO)
                                                    '*****End Code influenced by PyTSEB
                                                Else 'Neutral or Stable conditions, STABLE SHOULD ACTUALLY HAVE A CORRECTION!!!!!
                                                    'y_m_act = 0
                                                    'y_m_act = 0

                                                    '*****Using the stability equations of Brutseart 2005, which was learned from PyTSEB, though this wind adjustment is not in PyTSEB
                                                    '****therefore it may be subject to the general public license agreement of PyTSEB see:https://github.com/hectornieto/pyTSEB
                                                    y_m_act = CalcY_MorHStable(WindHeightAct, DAct, L_MO)
                                                    y_m_new = CalcY_MorHStable(WindHeight, D, L_MO)
                                                    '*****End Code influenced by PyTSEB

                                                End If
                                                '****Zomact and dact could be calculated like is done above!!!!, also the stability terms could be added!!!!!!!!!******
                                                Select Case WindAdjustment
                                                    Case WindAdjustment.With_Stability_Terms
                                                        '    WindSpeed = WindSpeed_Meas * ((Math.Log((WindHeight - D) / Zom) - y_m_new) / (Math.Log((WindHeightAct - DAct) / Zomact)) - y_m_act) 'Added by JBB Adjusts windspeed measured over some other surface to equivalent measured over Cover of Interest Following ASCE Std. Eq. B14 +  stability terms This method is erroneous because the friction velocity is not the same over the two different surfaces
                                                        WindSpeed = -999 'Doesn't work with stability terms, they are not in Allen and Wright, but we could develop that in the future!!! JBB
                                                    Case WindAdjustment.Without_Stability_Terms 'No stability terms included - like neutral conditions - common in reference ET calculation adjustments JBB
                                                        ' WindSpeed = WindSpeed_Meas * ((Math.Log((WindHeight - D) / Zom)) / (Math.Log((WindHeightAct - DAct) / Zomact))) 'Added by JBB Adjusts windspeed measured over some other surface to equivalent measured over Cover of Interest Following ASCE Std. Eq. B14 This method is erroneous because the friction velocity is not the same over the two different surfaces
                                                        If WindHeightAct < ZIBLw And WindHeight < ZIBLv Then 'Both heights below IBLs use Allen and Wright 1997 Eq. 13
                                                            WindSpeed = WindSpeed_Meas * (Math.Log((ZIBLw - DAct) / Zomact) * Math.Log((ZIBLv - Dr) / Zomr) * Math.Log((WindHeight - D) / Zom)) / (Math.Log((WindHeightAct - DAct) / Zomact) * Math.Log((ZIBLw - Dr) / Zomr) * Math.Log((ZIBLv - D) / Zom)) 'Allen and Wright 1997 Eq. 13
                                                        ElseIf WindHeightAct >= ZIBLw And WindHeight < ZIBLv Then 'Meas above IBL, use Allen and Wright 1997 Eq. 18
                                                            WindSpeed = WindSpeed_Meas * (Math.Log((ZIBLv - Dr) / Zomr) * Math.Log((WindHeight - D) / Zom)) / (Math.Log((WindHeightAct - Dr) / Zomr) * Math.Log((ZIBLv - D) / Zom)) 'Allen and Wright 1997 Eq. 18
                                                        ElseIf WindHeight < ZIBLw And WindHeight >= ZIBLv Then 'Crop wind height above IBL, use Allen and Wright 1997 Eq. 19
                                                            WindSpeed = WindSpeed_Meas * (Math.Log((ZIBLw - DAct) / Zomact) * Math.Log((WindHeight - Dr) / Zomr)) / (Math.Log((WindHeightAct - DAct) / Zomact) * Math.Log((ZIBLw - Dr) / Zomr)) 'Allen and Wright 1997 Eq. 19
                                                        Else 'Both heights above IBLs, use Allen and Wright 1997 Eq 20
                                                            WindSpeed = WindSpeed_Meas * (Math.Log((WindHeight - Dr) / Zomr)) / (Math.Log((WindHeightAct - Dr) / Zomr)) 'Allen and Wright 1997 Eq. 20
                                                        End If
                                                    Case Else
                                                        WindSpeed = WindSpeed 'No adjustment
                                                End Select
                                                '**********End Code Added by JBB ***************
                                                '****added by JBB for PM
                                                Dim TiniMethod As Integer = 0 'Defaults to PT
                                                Select Case TSMInitialTemperature
                                                    Case TSMInitialTemperature.Priestly_Taylor
                                                        TiniMethod = 0
                                                    Case TSMInitialTemperature.Penman_Monteith
                                                        TiniMethod = 1
                                                End Select

                                                '***** end added by JBB


                                                ':::if in case of stable conditions
                                                If L_MO > 0 Then
                                                    'EnergyComponents = calcRaMO_Stable_TSM(Resistances, AirTemperature, MaximumCoverHeight, L_MO, WindSpeed, FractionOfCover, Clump_Sun.Clump0, LAI, TInitial, Temperature, Ftheta, Rs, Ea, RecordDate.Month, Zom, Zoh, D, WindHeight, TemperatureHeight, Ag, Bioproperties.s, RnCoefficients, WPT, Cover) 'commented out by JBB
                                                    EnergyComponents = calcRaMO_Stable_TSM(Resistances, AirTemperature, MaximumCoverHeight, L_MO, WindSpeed, FractionOfCover, Clump_Sun.Clump0, LAI, TInitial, Temperature, Ftheta, Rs, Ea, RecordDate.Month, Zom, Zoh, D, WindHeight, TemperatureHeight, Ag, Bioproperties.s, RnCoefficients, WPT, Cover, Bioproperties, TiniMethod, Es) 'added by JBB
                                                    Stability = EnergyComponents.Stability 'EnergyComponents.Stability is never changed, those lines are all commented out, so this doesn't do anything JBB
                                                    'ElseIf Math.Abs((WindHeight - D) / L_MO) < 0.00001 Then '<--- It appears that this has been commented for some time as the calcRaMO_Neutral_TSM function doesn't match up well with the unstable and the stable functions JBB
                                                    '    EnergyComponents = calcRaMO_Neutral_TSM(Resistances, WindSpeed, WindHeight, TemperatureHeight, L_MO, Temperature, D, Ftheta, Ag, AirTemperature, calcTInitial(Temperature, AirTemperature, Ftheta).Tsoil, Rs, Ea, RecordDate, NDVI, Albedo, SAVI, LAI, Bioproperties.s, FractionOfCover, Zoh, Zom, MaximumCoverHeight, Clump_Sun.Clump0, WPT, RnCoefficients, Cover)
                                                    '    Stability = 1
                                                    '::: in case of unstable conditions
                                                ElseIf L_MO < 0 Then
                                                    'EnergyComponents = calcRaMO_Unstable_TSM(Resistances, WindHeight, TemperatureHeight, Clump_Sun.Clump0, Ag, Ea, TInitial, Rs, RecordDate.Month, RnCoefficients, MaximumCoverHeight, Ftheta, Temperature, AirTemperature, FractionOfCover, Bioproperties.s, Zoh, LAI, D, L_MO, WindSpeed, Zom, WPT, Cover, Albedo)'commented out by JBB
                                                    EnergyComponents = calcRaMO_Unstable_TSM(Resistances, WindHeight, TemperatureHeight, Clump_Sun.Clump0, Ag, Ea, TInitial, Rs, RecordDate.Month, RnCoefficients, MaximumCoverHeight, Ftheta, Temperature, AirTemperature, FractionOfCover, Bioproperties.s, Zoh, LAI, D, L_MO, WindSpeed, Zom, WPT, Cover, Albedo, Bioproperties, InitialHcanopy, TiniMethod, Es) 'added by JBB
                                                End If

                                                EnergyComponents.HTotal = EnergyComponents.Hcanopy + EnergyComponents.Hsoil
                                                EnergyComponents.ETotal = EnergyComponents.EVsoil + EnergyComponents.Ecanopy


                                                If Stability > 0 Then Exit For 'EnergyComponents.Stability is never changed, those lines are all commented out, so this doesn't do anything JBB

                                                'Dim LV As Single = 2501300 - 2366 * (AirTemperature - 273.15) 'This is the latent heat of vaporization in J/kg, but I don't know where the equation is from JBB
                                                Dim LV As Single = 2501000 - 2361 * (AirTemperature - 273.15) 'This is the latent heat of vaporization in J/kg, from Eq. 10 of Ham (2005) added by JBB
                                                Dim L_MO_new As Single = -WPT.CpRho * AirTemperature * Resistances.Ustar ^ 3 / (K * Gravity * (EnergyComponents.HTotal + 0.61 * AirTemperature * WPT.Cp2 * (EnergyComponents.ETotal) / LV)) 'Rearrangement of EQ 4.25 in Brutseart 1982, see also Ham 2005 latent heat of vaporization is included because E should be in mass flux units JBB

                                                If Math.Abs(L_MO - L_MO_new) >= 0.001 Then 'If the Monin Obukov Length changes by less than 0.001 then stop itterating as a solution has been found JBB
                                                    L_MO = L_MO_new
                                                    Lo_Out = L_MO 'for output added by JBB
                                                Else
                                                    Exit For
                                                End If
                                            Next 'Loop for Counter (50 itterations) JBB
                                            PT_Out = EnergyComponents.PstTlr 'added by JBB to output PT
                                            Rcanopy_Out = EnergyComponents.Rcpy
                                            U_Out = WindSpeed 'added by JBB to output windspeed for debug
                                            ustar_Out = Resistances.Ustar
                                            If U_Out < 0.001 Then 'debugging
                                                U_Out = U_Out
                                            End If
                                            G = EnergyComponents.GTotal
                                            Rn = EnergyComponents.RnTotal
                                            H = EnergyComponents.HTotal
                                            LE = Rn - G - H 'This is almost exactly the same as Etotal from the Energy Components, it is off in the 5th decimal or so JBB
                                            LEc_Out = EnergyComponents.Ecanopy ' added by JBB
                                            LEs_Out = EnergyComponents.EVsoil 'added by JBB
                                            fg_Out = Bioproperties.fg 'added by JBB
                                            Ustar_Out = Resistances.Ustar 'added by JBB
                                            Rns_Out = EnergyComponents.RnSoil 'added by JBB
                                            Rnc_Out = EnergyComponents.RnCanopy 'added by JBB
                                            Tc_Out = TInitial.Tcanopy - 273.15 'added by JBB
                                            Ts_Out = TInitial.Tsoil - 273.15 'added by JBB
                                            Tac_Out = TInitial.Tac - 273.15 'added by JBB

                                            If (H) > 1000 Then 'This seems to be a debug artifact JBB
                                                Rn = Rn
                                            End If
                                            If Cover = Functions.Cover.Bare_Soil Then 'This seems to be a debug artifact JBB
                                                Rn = Rn
                                            End If '**********************END TSEB Code**************JBB
                                            'Case EnergyBalance.SEBAL_Idaho '**********************I did not review the SEBAL_Idaho Code**************JBB


                                            '    Dim RlDown As Single = 0
                                            '    Dim RlUp As Single = 0
                                            '    Dim Emissivity As Single = 0

                                            '    Dim LMax As Single = 0
                                            '    Dim Lin As Single = 0
                                            '    Dim QCalMax As Single = 0
                                            '    Dim QCalMin As Single = 0
                                            '    Dim Lλ 'This is an unused variable JBB

                                            '    Rn = (1 - Albedo) * Rs + RlDown - RlUp - (1 - Emissivity) * RlDown
                                            '    LE = Rn - G - H
                                    End Select

                                    ET = calcET(G, Rn, LE, AirTemperature, AvailableEnergy, InstantaneousET, DailyReferenceET, ETextrapolation)

                                    'Writing to output rasters'<---Outputs pixel by pixel into the ....pixels arrays, as we are inside a loop that loops through all pixels JBB
                                    Dim CountOut As Integer = 0
                                    If OutputImagesBoxEnergy.GetItemChecked(0) Then OutPixels(CountOut).SetValue(Clean(Rah), {Col, Row}) : CountOut += 1 'CountOut +=1 is just incrementing the output image index e.g. from the Rah image to the Taero, and so on JBB
                                    If OutputImagesBoxEnergy.GetItemChecked(1) Then OutPixels(CountOut).SetValue(Clean(To_Out), {Col, Row}) : CountOut += 1 'modified by JBB 2/1/2018

                                    '@Ashish; "Albedo_Out" out was used but it is outside the scope of Energy Balance Switch Case statement and the values weren't updateing, it resulted in zeros 
                                    ' irrespective of input NIR, Red & Green values.
                                    'If OutputImagesBoxEnergy.GetItemChecked(2) Then OutPixels(CountOut).SetValue(Clean(Albedo_Out), {Col, Row}) : CountOut += 1 'modified by JBB 2/1/2018
                                    If OutputImagesBoxEnergy.GetItemChecked(2) Then OutPixels(CountOut).SetValue(Clean(Albedo), {Col, Row}) : CountOut += 1 'modified by JBB 2/1/2018
                                    If OutputImagesBoxEnergy.GetItemChecked(3) Then OutPixels(CountOut).SetValue(Clean(MaximumCoverHeight), {Col, Row}) : CountOut += 1 'added by JBB
                                    If OutputImagesBoxEnergy.GetItemChecked(4) Then OutPixels(CountOut).SetValue(Clean(Rcanopy_Out), {Col, Row}) : CountOut += 1 'added by JBB
                                    If OutputImagesBoxEnergy.GetItemChecked(5) Then OutPixels(CountOut).SetValue(Clean(ET), {Col, Row}) : CountOut += 1
                                    If OutputImagesBoxEnergy.GetItemChecked(6) Then OutPixels(CountOut).SetValue(Clean(FractionOfCover), {Col, Row}) : CountOut += 1 'added by JBB
                                    If OutputImagesBoxEnergy.GetItemChecked(7) Then OutPixels(CountOut).SetValue(Clean(fg_Out), {Col, Row}) : CountOut += 1 'added by JBB
                                    If OutputImagesBoxEnergy.GetItemChecked(8) Then OutPixels(CountOut).SetValue(Clean(Ustar_Out), {Col, Row}) : CountOut += 1 'added by JBB
                                    If OutputImagesBoxEnergy.GetItemChecked(9) Then OutPixels(CountOut).SetValue(Clean(LE), {Col, Row}) : CountOut += 1
                                    If OutputImagesBoxEnergy.GetItemChecked(10) Then OutPixels(CountOut).SetValue(Clean(LEc_Out), {Col, Row}) : CountOut += 1 'added by JBB
                                    If OutputImagesBoxEnergy.GetItemChecked(11) Then OutPixels(CountOut).SetValue(Clean(LEs_Out), {Col, Row}) : CountOut += 1 'added by JBB
                                    If OutputImagesBoxEnergy.GetItemChecked(12) Then OutPixels(CountOut).SetValue(Clean(LAI), {Col, Row}) : CountOut += 1
                                    If OutputImagesBoxEnergy.GetItemChecked(13) Then OutPixels(CountOut).SetValue(Clean(Rn), {Col, Row}) : CountOut += 1
                                    If OutputImagesBoxEnergy.GetItemChecked(14) Then OutPixels(CountOut).SetValue(Clean(Rnc_Out), {Col, Row}) : CountOut += 1 'added by JBB
                                    If OutputImagesBoxEnergy.GetItemChecked(15) Then OutPixels(CountOut).SetValue(Clean(Rns_Out), {Col, Row}) : CountOut += 1 'added by JBB
                                    If OutputImagesBoxEnergy.GetItemChecked(16) Then OutPixels(CountOut).SetValue(Clean(NDVI), {Col, Row}) : CountOut += 1
                                    If OutputImagesBoxEnergy.GetItemChecked(17) Then OutPixels(CountOut).SetValue(Clean(Lo_Out), {Col, Row}) : CountOut += 1 'added by JBB
                                    If OutputImagesBoxEnergy.GetItemChecked(18) Then OutPixels(CountOut).SetValue(Clean(oSAVI), {Col, Row}) : CountOut += 1
                                    If OutputImagesBoxEnergy.GetItemChecked(19) Then OutPixels(CountOut).SetValue(Clean(PT_Out), {Col, Row}) : CountOut += 1 'added by JBB
                                    If OutputImagesBoxEnergy.GetItemChecked(20) Then OutPixels(CountOut).SetValue(Clean(H), {Col, Row}) : CountOut += 1
                                    If OutputImagesBoxEnergy.GetItemChecked(21) Then OutPixels(CountOut).SetValue(Clean(SAVI), {Col, Row}) : CountOut += 1
                                    If OutputImagesBoxEnergy.GetItemChecked(22) Then OutPixels(CountOut).SetValue(Clean(G), {Col, Row}) : CountOut += 1
                                    If OutputImagesBoxEnergy.GetItemChecked(23) Then OutPixels(CountOut).SetValue(Clean(Tc_Out), {Col, Row}) : CountOut += 1 'added by JBB
                                    If OutputImagesBoxEnergy.GetItemChecked(24) Then OutPixels(CountOut).SetValue(Clean(Tac_Out), {Col, Row}) : CountOut += 1 'added by JBB
                                    If OutputImagesBoxEnergy.GetItemChecked(25) Then OutPixels(CountOut).SetValue(Clean(Ts_Out), {Col, Row}) : CountOut += 1 'added by JBB
                                    If OutputImagesBoxEnergy.GetItemChecked(26) Then OutPixels(CountOut).SetValue(Clean(U_Out), {Col, Row}) : CountOut += 1 'added by JBB

                                End If 'End if for check if the cover is included JBB
                            End If 'End if for check on cover >-1
                        Else 'This is if the pixel is not an intersecting pixel JBB
                            Dim CountOut As Integer = 0
                            For Output = 0 To 26
                                If OutputImagesBoxEnergy.GetItemChecked(Output) Then OutPixels(CountOut).SetValue(Single.MinValue, {Col, Row}) : CountOut += 1
                            Next
                        End If 'End if for check on whether cells are intersecting JBB
                        'End Sub) 'End of threading statment JBB 'This was commented out by JBB for single threaded debugging
                    Next 'Added by JBB for single threaded processing for debugging when using col and row loops like water balance
                Next 'Added by JBB for single threaded processing for debugging
                IntersectRasterCursor.Next() 'Advance the intersect cursor JBB
                For i = 0 To OutputRasters.Count - 1
                    OutRasterPixelBlock(i).PixelData(0) = OutPixels(i) 'Copies the OutPixels array into the OutRasterPixelBlock
                    OutRasterEdit(i).Write(CType(OutRasterCursor(i).TopLeft, ESRI.ArcGIS.Geodatabase.IPnt), OutRasterPixelBlock(i)) 'from ESRI: the .Write function "Writes a PixelBlock starting at a given Top-Left corner." JBB
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(OutRasterPixelBlock(i)) 'Releases the pixel block JBB
                Next
                For i = 1 To InputRasters.Count - 1 'Advances the input cursor JBB
                    InRasterCursor(i).Next()
                Next
                For i = 0 To OutputRasters.Count - 1 'Advances the output cursor JBB
                    OutRasterCursor(i).Next()
                Next

                System.Runtime.InteropServices.Marshal.ReleaseComObject(MultispectralPixelBlock) 'Releases the multispectral pixel block
                ProgressPartEnergy.PerformStep() 'Updates the lower progress bar JBB
                If Abort = True Then Exit Sub : Windows.Forms.Application.DoEvents() 'Exits is the abort button is clicked on the form JBB
            Loop While InRasterCursor(0).Next = True 'End of Loop through input raster cursor JBB
            ProgressAllEnergy.PerformStep() : Windows.Forms.Application.DoEvents() 'Updates the upper progress bar JBB

            CalculationTextEnergy.AppendText(vbNewLine & "   Calculating output image statistics...") : Windows.Forms.Application.DoEvents() 'Updates the progress text JBB
            For I = 0 To OutputRasters.Count - 1 'Loop through output images JBB
                Dim Geoprocessor As New ESRI.ArcGIS.Geoprocessor.Geoprocessor 'From ESRI "The geoprocessor is a helper object that simplifies the task of executing tools. Toolboxes define the set of tools available for the geoprocessor." JBB

                Geoprocessor.AddOutputsToMap = False
                Dim SetRasterProperties As New ESRI.ArcGIS.DataManagementTools.SetRasterProperties
                SetRasterProperties.in_raster = OutputRasters(I)
                SetRasterProperties.nodata = "1 -3.40282346638529E+38"
                Geoprocessor.Execute(SetRasterProperties, Nothing) 'I believe that this is taking the no data number and changing it to null or taking a null value and changing ti to a really large negative JBB

                Geoprocessor.AddOutputsToMap = True 'This helps following output images be added to the current map JBB
                Dim CalculateStatistics As New ESRI.ArcGIS.DataManagementTools.CalculateStatistics()
                CalculateStatistics.in_raster_dataset = OutputRasters(I)
                CalculateStatistics.ignore_values = Single.MinValue
                Geoprocessor.Execute(CalculateStatistics, Nothing) 'I think this helps the image not display the null pixes JBB

                System.Runtime.InteropServices.Marshal.ReleaseComObject(OutputRasters(I))
                System.Runtime.InteropServices.Marshal.ReleaseComObject(OutRaster2(I))
                System.Runtime.InteropServices.Marshal.ReleaseComObject(OutRasterBand(I))
                System.Runtime.InteropServices.Marshal.ReleaseComObject(OutRasterCursor(I))
                System.Runtime.InteropServices.Marshal.ReleaseComObject(OutRasterEdit(I))
            Next

            CalculationTextEnergy.AppendText(vbNewLine & "Succeeded at " & Now) : Windows.Forms.Application.DoEvents() 'Update Progress text JBB
        Next 'Advances the loop through Multispectral images JBB

        CalculationTextEnergy.AppendText(vbNewLine & "   Deleting temporary datasets...") : Windows.Forms.Application.DoEvents() 'Update progress text JBB
        For I = 0 To InputRasters.Count - 1 'deletes the input rasters rasterdatasets JBB
            Dim Path As String = InputRasters(I).CompleteName
            DeleteArcGISFile(Path)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(InputRasters(I))
        Next
        System.Runtime.InteropServices.Marshal.ReleaseComObject(IntersectRaster) 'Releases other itmes
        System.Runtime.InteropServices.Marshal.ReleaseComObject(IntersectRasterValue)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(TransformRaster2)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(WeatherTable)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(Workspace)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(WorkspaceFactory)
        Timer.Stop() : ProgressAllEnergy.PerformStep() : CalculationTextEnergy.AppendText(vbNewLine & "Elasped time (" & Timer.Elapsed.ToString & ").") 'Final update of text JBB
        If NegFlag = True Then
            MsgBox("A negative vegetation index was computed, check output for invalid pixels, e.g. NDVI, SAVI, oSAVI ≤ 0")
        End If
    End Sub

#End Region

#Region "Water Balance"

    Private Sub OutputDirectoryAddWater_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OutputDirectoryAddWater.Click
        Dim OpenFileDialog As ESRI.ArcGIS.CatalogUI.IGxDialog = New ESRI.ArcGIS.CatalogUI.GxDialog
        OpenFileDialog.Title = "Choose output location"
        OpenFileDialog.AllowMultiSelect = False
        Dim Filter As ESRI.ArcGIS.Catalog.IGxObjectFilter = New ESRI.ArcGIS.Catalog.GxFilterBasicTypes
        OpenFileDialog.ObjectFilter = Filter
        Dim List As ESRI.ArcGIS.Catalog.IEnumGxObject = Nothing

        If Not OpenFileDialog.DoModalOpen(Me.Handle, List) Then Exit Sub

        Dim FileInfo As ESRI.ArcGIS.Catalog.IGxObject = List.Next
        OutputDirectoryTextWater.Text = FileInfo.FullName
    End Sub

    Private Sub OutputImagesCheckAllWater_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OutputImagesCheckAllWater.Click
        For Item = 0 To OutputImagesBoxWater.Items.Count - 1
            OutputImagesBoxWater.SetItemChecked(Item, True)
        Next
    End Sub

    Private Sub OutputImagesUncheckAllWater_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OutputImagesUncheckAllWater.Click
        For Item = 0 To OutputImagesBoxWater.Items.Count - 1
            OutputImagesBoxWater.SetItemChecked(Item, False)
        Next
    End Sub

    Private Sub CalculationCoordinatesAdd_Click(sender As System.Object, e As System.EventArgs) Handles CalculationCoordinatesAdd.Click
        Me.Hide()
        MsgBox("Select point on the map for water balance point output, once selected, the SETMI widow will return.") 'Added by JBB 2/8/2018
        Dim UID As New ESRI.ArcGIS.esriSystem.UID
        UID.Value = "SETMI.Coordinatetool"

        Dim CommandItem As ESRI.ArcGIS.Framework.ICommandItem = SETMItool.ArcApplication.Document.CommandBars.Find(UID)
        Dim Command As ESRI.ArcGIS.SystemUI.ICommand = CommandItem.Command
        Dim CoordinateData As Coordinatetool = CType(Command, Coordinatetool)

        CommandItem.Execute()
    End Sub

    Private Sub CalculationCoordinatesRemove_Click(sender As System.Object, e As System.EventArgs) Handles CalculationCoordinatesRemove.Click
        For Each Row In CalculationCoordinatesGrid.SelectedRows
            CalculationCoordinatesGrid.Rows.Remove(Row)
        Next
    End Sub

    Private Sub RunWater_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RunWater.Click
        For I = 0 To CoverSelectionGrid.Rows.Count - 1
            If CoverSelectionGrid.Rows(I).Cells(3).Value = True And CoverSelectionGrid.Rows(I).Cells(2).Value = "" Then
                TabControl1.SelectedIndex = 2
                CoverSelectionGrid.ClearSelection()
                CoverSelectionGrid.Rows(I).Cells(2).Selected = True
                MsgBox("Please select an irrigation method for " & CoverSelectionGrid.Rows(I).Cells(0).Value & ".")
                Exit Sub
            End If
        Next
        If WaterBalanceBox.Text = "" Then : TabControl1.SelectedIndex = 0 : WaterBalanceBox.Focus() : MsgBox("Please select a model.")
        ElseIf CoverClassificationList.Items.Count = 0 Then : TabControl1.SelectedIndex = 2 : CoverClassificationAdd.Focus() : MsgBox("Please add at least one cover classification image.")
        Else
            Dim MultispectralDate As New List(Of DateTime)
            Dim CoverClassificationDate As New List(Of DateTime)
            Dim WeatherGrid As New WeatherGrid : WeatherGrid.Clear()
            '**********Added by JBB**************'Allows for input of irrigation and precipitation images
            Dim IrrigationDate As New List(Of DateTime)
            Dim PrecipitationDate As New List(Of DateTime)
            Dim DepletionDate As New List(Of DateTime)
            Dim ThetaVDate As New List(Of DateTime)
            Dim EvaporatedDepthDate As New List(Of DateTime)
            Dim SkinEvapDepthDate As New List(Of DateTime)
            '*************End Added by JBB *************************
            Try
                For I = 0 To MultispectralList.Items.Count - 1 : MultispectralDate.Add(GetDateFromPath(MultispectralList.Items(I))) : Next
                For I = 0 To CoverClassificationList.Items.Count - 1 : CoverClassificationDate.Add(GetDateFromPath(CoverClassificationList.Items(I))) : Next
                '**********Added by JBB**************'Allows for input of irrigation and precipitation images
                For I = 0 To WeatherPrecipitationDailyList.Items.Count - 1 : PrecipitationDate.Add(GetDateFromPath(WeatherPrecipitationDailyList.Items(I))) : Next
                For I = 0 To WeatherDailyIrrigationList.Items.Count - 1 : IrrigationDate.Add(GetDateFromPath(WeatherDailyIrrigationList.Items(I))) : Next
                For I = 0 To WeatherDailyDepletionList.Items.Count - 1 : DepletionDate.Add(GetDateFromPath(WeatherDailyDepletionList.Items(I))) : Next
                For I = 0 To WeatherDailyThetaVList.Items.Count - 1 : ThetaVDate.Add(GetDateFromPath(WeatherDailyThetaVList.Items(I))) : Next
                For I = 0 To WeatherDailyEvaporatedDepthList.Items.Count - 1 : EvaporatedDepthDate.Add(GetDateFromPath(WeatherDailyEvaporatedDepthList.Items(I))) : Next
                For I = 0 To WeatherDailySkinEvapDepthList.Items.Count - 1 : SkinEvapDepthDate.Add(GetDateFromPath(WeatherDailySkinEvapDepthList.Items(I))) : Next 'Added by JBB copying others
                '*************End Added by JBB *************************
            Catch ex As Exception
                TabControl1.SelectedIndex = 1
                MultispectralAdd.Focus()
                MsgBox("All image file names must end with acquisition date stamp, " & DateString & ".")
                Exit Sub
            End Try
            Dim MultispectralImage As New List(Of String)
            Dim CoverClassificationImage As New List(Of String)
            '**********Added by JBB**************'Allows for input of irrigation and precipitation images
            'Dim PrecipitationImage As New List(Of String)
            'Dim IrrigationImage As New List(Of String)
            '********End Added by JBB **************
            For I = 0 To MultispectralDate.Count - 1
                MultispectralImage.Add(Format(MultispectralDate(I), "yyyyMMdd") & MultispectralList.Items(I)) 'this sticks the date infront of the path for a chronological sort, the date is removed later JBB
            Next
            For I = 0 To CoverClassificationDate.Count - 1
                CoverClassificationImage.Add(Format(CoverClassificationDate(I), "yyyyMMdd") & CoverClassificationList.Items(I)) 'this sticks the date infront of the path for a chronological sort, the date is removed later JBB
            Next
            '**********Added by JBB**************'Allows for input of irrigation and precipitation images
            For I = 0 To PrecipitationDate.Count - 1
                WeatherGrid.Precipitation.Add(Format(PrecipitationDate(I), "yyyyMMdd") & WeatherPrecipitationDailyList.Items(I)) 'this sticks the date infront of the path for a chronological sort, the date is removed later JBB
            Next
            For I = 0 To IrrigationDate.Count - 1
                WeatherGrid.Irrigation.Add(Format(IrrigationDate(I), "yyyyMMdd") & WeatherDailyIrrigationList.Items(I)) 'this sticks the date infront of the path for a chronological sort, the date is removed later JBB
            Next
            For I = 0 To DepletionDate.Count - 1
                WeatherGrid.Depletion.Add(Format(DepletionDate(I), "yyyyMMdd") & WeatherDailyDepletionList.Items(I)) 'this sticks the date infront of the path for a chronological sort, the date is removed later JBB
            Next
            For I = 0 To ThetaVDate.Count - 1
                WeatherGrid.ThetaV.Add(Format(ThetaVDate(I), "yyyyMMdd") & WeatherDailyThetaVList.Items(I)) 'this sticks the date infront of the path for a chronological sort, the date is removed later JBB
            Next
            For I = 0 To EvaporatedDepthDate.Count - 1
                WeatherGrid.EvaporatedDepth.Add(Format(EvaporatedDepthDate(I), "yyyyMMdd") & WeatherDailyEvaporatedDepthList.Items(I)) 'this sticks the date infront of the path for a chronological sort, the date is removed later JBB
            Next
            For I = 0 To SkinEvapDepthDate.Count - 1 'added by JBB copying others
                WeatherGrid.SkinEvapDepth.Add(Format(SkinEvapDepthDate(I), "yyyyMMdd") & WeatherDailySkinEvapDepthList.Items(I)) 'this sticks the date infront of the path for a chronological sort, the date is removed later JBB
            Next


            WeatherGrid.Precipitation.Sort() 'Sort in Chronological Order JBB
            WeatherGrid.Irrigation.Sort() 'Sort in Chronological Order JBB
            WeatherGrid.Depletion.Sort() 'Sort in Chronological Order JBB
            WeatherGrid.ThetaV.Sort() 'sort in chrono order JBB
            WeatherGrid.EvaporatedDepth.Sort() 'sort in chrono order JBB
            WeatherGrid.SkinEvapDepth.Sort() 'copied from others
            '********End Added by JBB **************
            MultispectralImage.Sort() 'Sort in Chronological Order JBB
            CoverClassificationImage.Sort() 'Sort in Chronological Order JBB
            For I = 0 To MultispectralImage.Count - 1
                MultispectralImage(I) = MultispectralImage(I).Remove(0, 8) 'This removes the time stamp from before the path after the chronological sort has been done JBB
                Dim RecordDate As DateTime = GetDateFromPath(MultispectralImage(I))
                Dim Index As Integer = GetSameDateImageIndex(RecordDate, WeatherETDailyActualList.Items.Cast(Of [String])().ToList())
                If Index > -1 Then WeatherGrid.ETDailyActual.Add(WeatherETDailyActualList.Items(Index))
            Next
            For I = 0 To CoverClassificationImage.Count - 1
                CoverClassificationImage(I) = CoverClassificationImage(I).Remove(0, 8) 'This removes the time stamp from before the path after the chronological sort has been done JBB
            Next
            '******************Added by JBB ********************************
            For I = 0 To WeatherGrid.Precipitation.Count - 1
                WeatherGrid.Precipitation(I) = WeatherGrid.Precipitation(I).Remove(0, 8) 'This removes the time stamp from before the path after the chronological sort has been done JBB
            Next
            For I = 0 To WeatherGrid.Irrigation.Count - 1
                WeatherGrid.Irrigation(I) = WeatherGrid.Irrigation(I).Remove(0, 8) 'This removes the time stamp from before the path after the chronological sort has been done JBB
            Next
            For I = 0 To WeatherGrid.Depletion.Count - 1
                WeatherGrid.Depletion(I) = WeatherGrid.Depletion(I).Remove(0, 8) 'This removes the time stamp from before the path after the chronological sort has been done JBB
            Next
            For I = 0 To WeatherGrid.ThetaV.Count - 1
                WeatherGrid.ThetaV(I) = WeatherGrid.ThetaV(I).Remove(0, 8) 'This removes the time stamp from before the path after the chronological sort has been done JBB
            Next
            For I = 0 To WeatherGrid.EvaporatedDepth.Count - 1
                WeatherGrid.EvaporatedDepth(I) = WeatherGrid.EvaporatedDepth(I).Remove(0, 8) 'This removes the time stamp from before the path after the chronological sort has been done JBB
            Next
            For I = 0 To WeatherGrid.SkinEvapDepth.Count - 1 'added by JBB copying others
                WeatherGrid.SkinEvapDepth(I) = WeatherGrid.SkinEvapDepth(I).Remove(0, 8) 'This removes the time stamp from before the path after the chronological sort has been done JBB
            Next
            '*****************End Added by JBB *******************************

            If Not FileExists(MultispectralList, MultispectralImage, 1) Then Exit Sub
            If Not FileExists(CoverClassificationList, CoverClassificationImage, 1) Then Exit Sub
            If Not FileExists(FieldCapacityText, 1) Then Exit Sub
            If Not FileExists(WiltingPointText, 1) Then Exit Sub
            If Not ExistsArcGISFile(WeatherGrid.ETDailyActual) Then Exit Sub
            '***Added by JBB***
            If Not ExistsArcGISFile(WeatherGrid.Precipitation) Then Exit Sub
            If Not ExistsArcGISFile(WeatherGrid.Irrigation) Then Exit Sub
            If Not ExistsArcGISFile(WeatherGrid.Depletion) Then Exit Sub
            If Not ExistsArcGISFile(WeatherGrid.ThetaV) Then Exit Sub
            If Not ExistsArcGISFile(WeatherGrid.EvaporatedDepth) Then Exit Sub
            If Not FileExists(InitialSoilMoistureText, 1) Then Exit Sub
            If Not ExistsArcGISFile(WeatherGrid.SkinEvapDepth) Then Exit Sub 'added by JBB copying others
            If Not FileExists(FieldCapacityText2, 1) Then Exit Sub 'added by JBB copying others
            If Not FileExists(FieldCapacityText3, 1) Then Exit Sub 'added by JBB copying others
            If Not FileExists(WiltingPointText2, 1) Then Exit Sub 'added by JBB copying others
            If Not FileExists(WiltingPointText3, 1) Then Exit Sub 'added by JBB copying others
            If Not FileExists(InitialSoilMoistureText2, 1) Then Exit Sub 'added by JBB copying others
            If Not FileExists(InitialSoilMoistureText3, 1) Then Exit Sub 'added by JBB copying others
            If Not FileExists(LayerThicknessText, 1) Then Exit Sub 'added by JBB copying others
            If Not FileExists(LayerThicknessText2, 1) Then Exit Sub 'added by JBB copying others
            If Not FileExists(LayerThicknessText3, 1) Then Exit Sub 'added by JBB copying others
            If Not FileExists(HydraulicConductivityText, 1) Then Exit Sub 'added by JBB copying others
            If Not FileExists(HydraulicConductivityText2, 1) Then Exit Sub 'added by JBB copying others
            If Not FileExists(HydraulicConductivityText3, 1) Then Exit Sub 'added by JBB copying others
            If Not FileExists(SaturatedWaterContentText, 1) Then Exit Sub 'added by JBB copying others
            If Not FileExists(SaturatedWaterContentText2, 1) Then Exit Sub 'added by JBB copying others
            If Not FileExists(SaturatedWaterContentText3, 1) Then Exit Sub 'added by JBB copying others
            If Not FileExists(REWText, 1) Then Exit Sub 'added by JBB copying others
            If Not FileExists(DPLimitText, 1) Then Exit Sub 'added by JBB copying others
            '***End Added by JBB****

            Abort = False
            CalculateWaterBalance(MultispectralImage, CoverClassificationImage, WeatherGrid)
        End If
        MsgBox("Water Balance Calculations Are Complete") 'Added by JBB
    End Sub

    Private Sub ExitRunWater_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitRunWater.Click
        If Abort = True Then Exit Sub 'Allows user to abort the model execution by clicking the little X on the form control JBB
        CalculationTextWater.AppendText(vbNewLine & "Operation aborted..." & Now) 'Updates the progress box JBB
        Abort = True 'The abort Boolean appears in the water balance subroutines as atrigger to exit sub. JBB
    End Sub

    Private Sub CalculateWaterBalance(ByVal MultispectralImages As List(Of String), ByVal CoverClassificationImages As List(Of String), ByVal WeatherGrid As WeatherGrid)
        Dim FlgImageCntEarly As Boolean = False 'added by JBB to prompt when too few early images have been added for savi log interp
        Dim FlgImageCntEarly2 As Boolean = False 'added by JBB to prompt when too few early images have been added for savi log interp
        Dim FlgImageCntLate As Boolean = False 'added by JBB to prompt when too few late images have been added for savi log interp
        Dim FlgFalseSAVI As Boolean = False 'added by JBB to prompt that false SAVI values will be in play
        Dim FlgNoInputYr As Boolean = False 'added by JBB to prompt that no input year was 
        Dim FlgDepth1 As Boolean = False 'added by JBB to notify if the top layer depth was extended
        Dim CntForceFakeSAVI As Long = 0 'added by jbb to prompt if the false SAVI was forced for a pixel
        Dim NegFlag As Boolean = False 'added to find pixels that have negative vegetation indices.
        Dim MaxSAVIsEarly As Boolean = False 'added to check if peak SAVI images were forced into early season
        Dim FlgVI As Boolean = False 'added by JBB to check that Fc and Veg Indices match.
        Dim Timer As New Stopwatch : Timer.Start()
        Dim Ilyr1 As Integer = 0 'added by JBB to allow for different soil inputs
        Dim Ilyr2 As Integer = 1 'added by JBB to allow for different soil inputs
        Dim ILyr3 As Integer = 1 'added by JBB to allow for different soil inputs
        Dim IKs1 As Integer = 1
        Dim IKs2 As Integer = 1
        Dim IKs3 As Integer = 1
        Dim IVWCs1 As Integer = 1
        Dim IVWCs2 As Integer = 1
        Dim IVWCs3 As Integer = 1
        Dim Irew As Integer = 1
        Dim Idpl As Integer = 1

        ProgressAllWater.Maximum = 3 : ProgressAllWater.Minimum = 0 : ProgressAllWater.Step = 1 : ProgressAllWater.Value = 0 'Sets up top progress bar JBB
        ProgressPartWater.Minimum = 0 : ProgressPartWater.Step = 1 : ProgressPartWater.Value = 0 'Sets up lower progress bar JBB
        CalculationTextWater.Clear() : CalculationTextWater.AppendText("Determining intersecting area and output raster properties..." & Now) : Windows.Forms.Application.DoEvents() 'updates progress text box JBB

        Dim InputRasterPath As New List(Of String) 'Makes a list of input raster file paths JBB
        InputRasterPath.AddRange(MultispectralImages.ToList) 'Grab shortwave image path JBB
        InputRasterPath.AddRange(CoverClassificationImages.ToList) 'Grab cover classification image path JBB
        InputRasterPath.AddRange(WeatherGrid.AllValues) 'Grab weather grid image paths JBB <-- This is needed to grab ET data from the input ETdaily Grid for assimilation JBB
        'InputRasterPath.AddRange({FieldCapacityText.Text, WiltingPointText.Text}) 'Grab FC and WP file paths JBB commented out by JBB
        'InputRasterPath.AddRange({FieldCapacityText.Text, WiltingPointText.Text, InitialSoilMoistureText.Text}) 'Grab FC and WP and thetav file paths JBB added by JBB

        'InputRasterPath.AddRange({FieldCapacityText.Text, FieldCapacityText2.Text, FieldCapacityText3.Text, WiltingPointText.Text, WiltingPointText2.Text, WiltingPointText3.Text, InitialSoilMoistureText.Text, InitialSoilMoistureText2.Text, InitialSoilMoistureText3.Text, LayerThicknessText.Text, LayerThicknessText2.Text, LayerThicknessText3.Text, HydraulicConductivityText.Text, HydraulicConductivityText2.Text, HydraulicConductivityText3.Text, SaturatedWaterContentText.Text, SaturatedWaterContentText2.Text, SaturatedWaterContentText3.Text, REWText.Text}) 'Grab FC and WP and thetav file paths updated by JBB copying others 2/9/2018

        InputRasterPath.AddRange({FieldCapacityText.Text, WiltingPointText.Text, InitialSoilMoistureText.Text, LayerThicknessText.Text}) 'Grab FC and WP and thetav file paths updated by JBB copying others 2/9/2018
        '****Added by JBB based on code above
        'Second and Thrid depth input are optional, but must all be provided (except Ksat and Sat water content)
        If FieldCapacityText2.Text <> "" And WiltingPointText2.Text <> "" And InitialSoilMoistureText2.Text <> "" And LayerThicknessText2.Text <> "" Then
            InputRasterPath.AddRange({FieldCapacityText2.Text, WiltingPointText2.Text, InitialSoilMoistureText2.Text, LayerThicknessText2.Text})
            Ilyr1 = Ilyr1 - 4
            Ilyr2 = Ilyr2 - 1
        End If
        If FieldCapacityText3.Text <> "" And WiltingPointText3.Text <> "" And InitialSoilMoistureText3.Text <> "" And LayerThicknessText3.Text <> "" Then
            InputRasterPath.AddRange({FieldCapacityText3.Text, WiltingPointText3.Text, InitialSoilMoistureText3.Text, LayerThicknessText3.Text})
            Ilyr1 = Ilyr1 - 4
            Ilyr2 = Ilyr2 - 4
            ILyr3 = ILyr3 - 1
        End If
        If HydraulicConductivityText.Text <> "" Then
            InputRasterPath.AddRange({HydraulicConductivityText.Text})
            Ilyr1 = Ilyr1 - 1
            If Ilyr2 < 1 Then Ilyr2 = Ilyr2 - 1
            If ILyr3 < 1 Then ILyr3 = ILyr3 - 1
            IKs1 = IKs1 - 1
        End If
        If HydraulicConductivityText2.Text <> "" Then
            InputRasterPath.AddRange({HydraulicConductivityText2.Text})
            Ilyr1 = Ilyr1 - 1
            Ilyr2 = Ilyr2 - 1
            If ILyr3 < 1 Then ILyr3 = ILyr3 - 1
            IKs1 = IKs1 - 1
            IKs2 = IKs2 - 1
        End If
        If HydraulicConductivityText3.Text <> "" Then
            InputRasterPath.AddRange({HydraulicConductivityText3.Text})
            Ilyr1 = Ilyr1 - 1
            Ilyr2 = Ilyr2 - 1
            ILyr3 = ILyr3 - 1
            IKs1 = IKs1 - 1
            IKs2 = IKs2 - 1
            IKs3 = IKs3 - 1
        End If
        If SaturatedWaterContentText.Text <> "" Then
            InputRasterPath.AddRange({SaturatedWaterContentText.Text})
            Ilyr1 = Ilyr1 - 1
            Ilyr2 = Ilyr2 - 1
            ILyr3 = ILyr3 - 1
            If IKs1 < 1 Then IKs1 = IKs1 - 1
            If IKs3 < 1 Then IKs2 = IKs2 - 1
            If IKs3 < 1 Then IKs3 = IKs3 - 1
            IVWCs1 = IVWCs1 - 1
        End If
        If SaturatedWaterContentText2.Text <> "" Then
            InputRasterPath.AddRange({SaturatedWaterContentText2.Text})
            Ilyr1 = Ilyr1 - 1
            Ilyr2 = Ilyr2 - 1
            ILyr3 = ILyr3 - 1
            If IKs1 < 1 Then IKs1 = IKs1 - 1
            If IKs3 < 1 Then IKs2 = IKs2 - 1
            If IKs3 < 1 Then IKs3 = IKs3 - 1
            IVWCs1 = IVWCs1 - 1
            IVWCs2 = IVWCs2 - 1
        End If
        If SaturatedWaterContentText3.Text <> "" Then
            InputRasterPath.AddRange({SaturatedWaterContentText3.Text})
            Ilyr1 = Ilyr1 - 1
            If Ilyr2 < 1 Then Ilyr2 = Ilyr2 - 1
            If ILyr3 < 1 Then ILyr3 = ILyr3 - 1
            If IKs1 < 1 Then IKs1 = IKs1 - 1
            If IKs3 < 1 Then IKs2 = IKs2 - 1
            If IKs3 < 1 Then IKs3 = IKs3 - 1
            IVWCs1 = IVWCs1 - 1
            IVWCs2 = IVWCs2 - 1
            IVWCs3 = IVWCs3 - 1
        End If
        If REWText.Text <> "" Then
            InputRasterPath.AddRange({REWText.Text})
            Ilyr1 = Ilyr1 - 1
            If Ilyr2 < 1 Then Ilyr2 = Ilyr2 - 1
            If ILyr3 < 1 Then ILyr3 = ILyr3 - 1
            If IKs1 < 1 Then IKs1 = IKs1 - 1
            If IKs3 < 1 Then IKs2 = IKs2 - 1
            If IKs3 < 1 Then IKs3 = IKs3 - 1
            If IVWCs1 < 1 Then IVWCs1 = IVWCs1 - 1
            If IVWCs2 < 1 Then IVWCs2 = IVWCs2 - 1
            If IVWCs3 < 1 Then IVWCs3 = IVWCs3 - 1
            Irew = Irew - 1
        End If
        If DPLimitText.Text <> "" Then
            InputRasterPath.AddRange({DPLimitText.Text})
            Ilyr1 = Ilyr1 - 1
            If Ilyr2 < 1 Then Ilyr2 = Ilyr2 - 1
            If ILyr3 < 1 Then ILyr3 = ILyr3 - 1
            If IKs1 < 1 Then IKs1 = IKs1 - 1
            If IKs3 < 1 Then IKs2 = IKs2 - 1
            If IKs3 < 1 Then IKs3 = IKs3 - 1
            If IVWCs1 < 1 Then IVWCs1 = IVWCs1 - 1
            If IVWCs2 < 1 Then IVWCs2 = IVWCs2 - 1
            If IVWCs3 < 1 Then IVWCs3 = IVWCs3 - 1
            If Irew < 1 Then Irew = Irew - 1
            Idpl = Idpl - 1
        End If
        '***end added by JBB

        Dim IntersectRasterPath As String = IO.Path.ChangeExtension(IO.Path.GetTempFileName, ".img") 'Creates a temp .img file for the intersect raster path JBB
        Dim InputRasters(InputRasterPath.Count - 1) As ESRI.ArcGIS.Geodatabase.IRasterDataset 'IRasterDataset has been superceeded by IRasterDataset3 JBB
        Dim IntersectRaster = CreateIntersectRaster(InputRasterPath.ToArray, InputRasters, IntersectRasterPath) 'Creates a raster dataset for the intersect raster JBB
        Dim IntersectRasterValue As ESRI.ArcGIS.Geodatabase.IRasterValue = New ESRI.ArcGIS.Geodatabase.RasterValue 'From ESRI "This is the entry point to load or update raster data in the geodatabase." JBB
        IntersectRasterValue.RasterDataset = IntersectRaster 'Sets the Intersect Raster as the Raster for RasterValue JBB
        CalculationTextWater.AppendText(vbNewLine & "Succeeded at " & Now) : CalculationTextWater.AppendText(vbNewLine & "Creating temporary intersecting datasets...") : Windows.Forms.Application.DoEvents() 'Update Progress Text Box JBB
        If Abort = True Then Exit Sub 'Exits is the abort button is clicked on the form JBB

        ProgressPartWater.Maximum = GetRasterCursorIterations(IntersectRaster) 'Advance the lower progress bar JBB

        For P = 0 To InputRasterPath.Count - 1
            CalculationTextWater.AppendText(vbNewLine & "   For " & IO.Path.GetFileName(InputRasterPath(P)) & "..." & Now) : Windows.Forms.Application.DoEvents() 'Update Progress Text Box JBB
            ExtractRaster(InputRasterPath(P), InputRasters(P), IntersectRasterPath, IntersectRaster) 'Extracts the input rasters into temporary rasters matching the intersect raster JBB
            CalculationTextWater.AppendText(vbNewLine & "   Succeeded at " & Now) : Windows.Forms.Application.DoEvents() 'Update Progress Text Box JBB
            If Abort = True Then Exit Sub 'Exits is the abort button is clicked on the form JBB
        Next
        ProgressAllWater.PerformStep() 'Advance upper progress bar JBB

        Dim OutputRasterNames As New List(Of String)
        For I = 0 To OutputImagesBoxWater.Items.Count - 1
            If OutputImagesBoxWater.GetItemChecked(I) Then OutputRasterNames.Add(OutputImagesBoxWater.Items.Item(I)) 'List selected output raster files JBB
        Next
        Dim WorkspaceFactory As ESRI.ArcGIS.Geodatabase.IWorkspaceFactory = New ESRI.ArcGIS.DataSourcesRaster.RasterWorkspaceFactoryClass() 'In other places RasterWorkspaceFactory is used, the...Class() may be a newer version JBB
        Dim Workspace As ESRI.ArcGIS.Geodatabase.IRasterWorkspace2 = CType(WorkspaceFactory.OpenFromFile(OutputDirectoryTextWater.Text, 0), ESRI.ArcGIS.DataSourcesRaster.IRasterWorkspace) 'Defines a workspace for output files JBB

        Dim WaterBalanceModel As WaterBalance = DirectCast([Enum].Parse(GetType(WaterBalance), WaterBalanceBox.SelectedItem.ToString.Replace(" ", "_")), WaterBalance) 'Grabs the selected water balance method Point/Grid JBB
        Dim ImageSource As ImageSource = DirectCast([Enum].Parse(GetType(ImageSource), ImageSourceBox.SelectedItem.ToString.Replace(" ", "_")), ImageSource) 'Grabs image source i.e. aerial, Landsat, MODIS JBB
        Dim AssimilationMethod As DataAssimilation = DirectCast([Enum].Parse(GetType(DataAssimilation), DataAssimilationBox.SelectedItem.ToString.Replace(" ", "_")), DataAssimilation) 'Grabs assimilation method e.g. none, single weighting, etc. JBB
        '***Start Code Added by JBB****
        Dim ETReference As ETReferenceType = DirectCast([Enum].Parse(GetType(ETReferenceType), ETReferenceBox.SelectedItem.ToString.Replace(" ", "_")), ETReferenceType) 'Grabs ETr type (tall or short) JBB
        Dim EffectivePrecipitationType As EffectivePrecipType = DirectCast([Enum].Parse(GetType(EffectivePrecipType), EffectivePreciptiationBox.SelectedItem.ToString.Replace(" ", "_")), EffectivePrecipType) 'Grabs Peff method, CN or % JBB
        Dim KcbMethod As KcbType = DirectCast([Enum].Parse(GetType(KcbType), KcbTypeBox.SelectedItem.ToString.Replace(" ", "_")), KcbType) 'grabs Kcb progression type JBB
        Dim FcMethodWB As FcType = DirectCast([Enum].Parse(GetType(FcType), FcTypeBox.SelectedItem.ToString.Replace(" ", "_")), FcType) 'Grabs Fc Method for WB
        Dim WBPntLocMeth As WBPntMethod = DirectCast([Enum].Parse(GetType(WBPntMethod), WBPntLocMethodBox.SelectedItem.ToString.Replace(" ", "_")), WBPntMethod) 'Grabs input point location method JBB, copied from FcType above, which is copied from others above
        Dim WBTeMethod As WBTeMethod = DirectCast([Enum].Parse(GetType(WBTeMethod), WBTeMethodBox.SelectedItem.ToString.Replace(" ", "_")), WBTeMethod) 'Grabs input point location method JBB, copied from FcType above, which is copied from others above
        Dim WBSkinEvapMethod As WBSkinMethod = DirectCast([Enum].Parse(GetType(WBSkinMethod), SkinLyrMethodBox.SelectedItem.ToString.Replace(" ", "_")), WBSkinMethod)
        Dim KcbVIForecastMeth As SAVIForecast = DirectCast([Enum].Parse(GetType(SAVIForecast), SAVIForecastMethodBox.SelectedItem.ToString.Replace(" ", "_")), SAVIForecast) ' copied from others above by JBB
        Dim DPMethod As DPLimit = DirectCast([Enum].Parse(GetType(DPLimit), DPLimitMethodBox.SelectedItem.ToString.Replace(" ", "_")), DPLimit) 'Grabs input point location method JBB, copied from others above

        '***Check If options were selected
        If ETReference <> ETReferenceType.Short_Grass And ETReference <> ETReferenceType.Tall_Alfalfa Then MsgBox("No ET Reference Type Selected, Program Will not Run Properly")
        'If ETReference = ETReferenceType.Tall_Alfalfa Then MsgBox("Caution: Not all cover options work with tall reference ET")
        If EffectivePrecipitationType <> EffectivePrecipType.Percent_Effective And EffectivePrecipitationType <> EffectivePrecipType.SCS_Curve_Number Then MsgBox("No Effective Precipitation Method Selected, Effective Precip. set to zero")
        If KcbMethod <> KcbType.Fitted_Curve And KcbMethod <> KcbType.VI_Interpolation And KcbMethod <> KcbType.VI_Log_Regression Then MsgBox("No Kcb Progression Method Selected, Model May Not Run Properly")
        '***End Code Added by JBB****
        Dim CoverTable = CreateTable(CoverPropertiesText.Text) 'Creates a table from the input cover properties .xls table JBB
        Dim CoverPoint As New CoverPoint
        CoverPoint.Populate(CoverTable, CoverPointIndex) 'Populates the cover datatable, populate is a SETMI function JBB
        Dim WeatherTable = CreateTable(WeatherTableText.Text) 'Creates a table from the input weather .xls table JBB
        Dim WeatherPoint As New WeatherPoint 'Weather point is the tabulated weather data JBB
        Dim WeatherGridIndex As New WeatherGridIndex 'Weather grid index is used to pair the gridded weather data with the calculation date JBB
        Dim WeatherOffset As Integer = MultispectralImages.Count + CoverClassificationImages.Count 'Used to know the index of the first weather grid image JBB
        '******Start JBB Added Code ******************
        'For precip, irrigation and depletion images
        Dim PrecipitationOffset As Integer = WeatherOffset + WeatherGrid.ETDailyActual.Count 'Finds index of first precip image JBB
        Dim IrrigationOffset As Integer = PrecipitationOffset + WeatherGrid.Precipitation.Count 'Finds index of first irrigation image JBB
        Dim DepletionOffset As Integer = IrrigationOffset + WeatherGrid.Irrigation.Count 'Finds the index of the first depletion image JBB
        Dim DepletionCount As Integer = WeatherGrid.Depletion.Count 'added by JBB to count thetav images
        Dim ThetaVOffset As Integer = DepletionOffset + WeatherGrid.Depletion.Count 'finds index of first thetav image JBB
        Dim ThetaVCount As Integer = WeatherGrid.ThetaV.Count 'added by JBB to count thetav images
        Dim EvaporatedDepthOffset As Integer = ThetaVOffset + WeatherGrid.ThetaV.Count 'finds index of first evaporated depth image JBB
        Dim EvaporatedDepthCount As Integer = WeatherGrid.EvaporatedDepth.Count 'added by JBB to count evaporated depth images
        Dim SkinEvapDepthOffset As Integer = EvaporatedDepthOffset + WeatherGrid.EvaporatedDepth.Count 'added by JBB copying others
        Dim SkinEvapDepthCount As Integer = WeatherGrid.SkinEvapDepth.Count 'added by JBB copying others
        'Four output stuff
        Dim WBOutDateTbl = CreateTable(WaterBalanceOutputDateText.Text) 'Added by JBB for water balance output date table JBB
        Dim WBOutputDates As New WBOutputDates
        Dim WBOutDateCount As Integer = WBOutputDates.CountRows(WBOutDateTbl, WBOutputDateIndex) 'Counts the number of rows in WBOutDateTbl JBB
        Dim WBOutDatesARR(WBOutDateCount - 1) As Date

        Dim WBOutCursor As ESRI.ArcGIS.Geodatabase.ICursor = WBOutDateTbl.Search(Nothing, True)
        Dim WBOutRow As ESRI.ArcGIS.Geodatabase.IRow = WBOutCursor.NextRow()
        Dim Count As Integer = 0
        Do While Not WBOutRow Is Nothing
            If Not IsDBNull(WBOutRow.Value(0)) Then
                If WBOutputDateIndex.WBOutDate > -1 Then WBOutDatesARR(Count) = CleanNull(WBOutRow.Value(WBOutputDateIndex.WBOutDate))
                Count += 1
            End If
            WBOutRow = WBOutCursor.NextRow()
        Loop
        System.Runtime.InteropServices.Marshal.ReleaseComObject(WBOutCursor)


        'Four output point locations table stuff, copied from above code for Output Date Table, which would be copied from elsewhere JBB
        Dim WBPntOutTbl = CreateTable(WBPointOutputText.Text) 'Added by JBB for water balance output date table JBB
        Dim WBPointOutputs As New WBPointOutputs
        Dim WBPntOutCount As Integer = WBPointOutputs.CountRows(WBPntOutTbl, WBPointOutputIndex) 'Counts the number of rows in WBOutDateTbl JBB
        Dim WBPntOutLocsARR(WBPntOutCount - 1, 1) As Double 'FOLLOWING THE DECLARATION FOR SELECTED POINTS FURTHER BELOW (JBB)

        Dim WBPntOutCursor As ESRI.ArcGIS.Geodatabase.ICursor = WBPntOutTbl.Search(Nothing, True)
        Dim WBPntOutRow As ESRI.ArcGIS.Geodatabase.IRow = WBPntOutCursor.NextRow()
        Dim PntCount As Integer = 0
        Do While Not WBPntOutRow Is Nothing
            If Not IsDBNull(WBPntOutRow.Value(0)) Then
                If WBPointOutputIndex.Xcoord > -1 Then WBPntOutLocsARR(PntCount, 0) = CleanNull(WBPntOutRow.Value(WBPointOutputIndex.Xcoord))
                If WBPointOutputIndex.Ycoord > -1 Then WBPntOutLocsARR(PntCount, 1) = CleanNull(WBPntOutRow.Value(WBPointOutputIndex.Ycoord))
                PntCount += 1
            End If
            WBPntOutRow = WBPntOutCursor.NextRow()
        Loop
        System.Runtime.InteropServices.Marshal.ReleaseComObject(WBPntOutCursor)

        '*******End JBB Added Code ************

        Dim CoverClassificationIndex(MultispectralImages.Count - 1) As Integer
        For M = 0 To MultispectralImages.Count - 1
            CoverClassificationIndex(M) = GetNearestDateImageIndex(MultispectralImages(M), CoverClassificationImages) 'Pairs cover image of nearest date to the multispectral image JBB
        Next

        Dim IntersectRasterBand As ESRI.ArcGIS.DataSourcesRaster.IRasterBandCollection = IntersectRaster 'Define a raster band collection for the intersecting raster JBB
        Dim IntersectRaster2 As ESRI.ArcGIS.DataSourcesRaster.IRaster2 = CType(CType(IntersectRaster.CreateDefaultRaster, ESRI.ArcGIS.Geodatabase.IRaster), ESRI.ArcGIS.DataSourcesRaster.IRaster2) 'Why is CreateDefaultRaster needed and why is .IRaster and .IRaster2 both used JBB
        Dim IntersectRasterCursor As ESRI.ArcGIS.Geodatabase.IRasterCursor = IntersectRaster2.CreateCursorEx(Nothing) 'Creates a raster cursor for the intersect raster, which allows the image to be processed 128 pixel rows at a time JBB

        Dim InRasterBand(InputRasters.Count - 1) As ESRI.ArcGIS.DataSourcesRaster.IRasterBandCollection 'Define raster band collections for each input image JBB
        Dim InRaster2(InputRasters.Count - 1) As ESRI.ArcGIS.DataSourcesRaster.IRaster2 'Sets an interface for working with input rasters JBB
        Dim InRasterCursor(InputRasters.Count - 1) As ESRI.ArcGIS.Geodatabase.IRasterCursor 'Creates a raster cursor for the input rasters, which allows the images to be processed 128 pixel rows at a time JBB
        For R = 0 To InputRasters.Count - 1
            InRasterBand(R) = InputRasters(R) 'Copies Input Rasters into the input raster band collection JBB
            InRaster2(R) = CType(CType(InputRasters(R).CreateDefaultRaster, ESRI.ArcGIS.Geodatabase.IRaster), ESRI.ArcGIS.DataSourcesRaster.IRaster2) 'Why is CreateDefaultRaster needed and why is .IRaster and .IRaster2 both used JBB
            InRasterCursor(R) = InRaster2(R).CreateCursorEx(Nothing) 'Sets the input raster cursor with initial values of nothing JBB
        Next

        'Dim OutputRasters(OutputRasterNames.Count - 1) As ESRI.ArcGIS.Geodatabase.IRasterDataset 'Makes a raster dataset for the output rasters JBB'This block was commented out by JBB
        'For I = 0 To OutputImagesBoxWater.CheckedItems.Count - 1 'Loops through selected output files
        '    Dim FileName As String = "SETMI " & WaterBalanceModel.ToString & " " & OutputImagesBoxWater.CheckedItems(I).ToString.Split("(")(1).Replace(")", "") & " " & GetDateFromPath(MultispectralImages(0)).Year & ".img" 'Creates file name string for output files JBB


        '    If ExistsArcGISFile(OutputDirectoryTextWater.Text & "\" & FileName) Then DeleteArcGISFile(OutputDirectoryTextWater.Text & "\" & FileName) 'Deletes existing output files JBB
        '    Dim OutputRasterDataset = Workspace.CreateRasterDataset(Name:=FileName, _
        '                                                            Format:="IMAGINE Image", _
        '                                                            Origin:=IntersectRasterValue.Extent.LowerLeft, _
        '                                                            columnCount:=IntersectRasterValue.Extent.Width / IntersectRasterValue.RasterStorageDef.CellSize.X, _
        '                                                            RowCount:=IntersectRasterValue.Extent.Height / IntersectRasterValue.RasterStorageDef.CellSize.Y, _
        '                                                            cellSizeX:=IntersectRasterValue.RasterStorageDef.CellSize.X, _
        '                                                            cellSizeY:=IntersectRasterValue.RasterStorageDef.CellSize.Y, _
        '                                                            numBands:=1, _
        '                                                            PixelType:=ESRI.ArcGIS.Geodatabase.rstPixelType.PT_FLOAT, _
        '                                                            SpatialReference:=IntersectRasterValue.Extent.SpatialReference, _
        '                                                            Permanent:=True) 'Creates and sets up output raster files JBB
        '    OutputRasters(I) = OutputRasterDataset 'Puts newly created output rasters into the output rasters raster dataset JBB
        'Next

        'Dim OutRasterBand(OutputRasters.Count - 1) As ESRI.ArcGIS.DataSourcesRaster.IRasterBandCollection 'Defines a raster band collection for the output rasters JBB
        'Dim OutRaster2(OutputRasters.Count - 1) As ESRI.ArcGIS.DataSourcesRaster.IRaster2 'Defines a raster interface for output rasters JBB
        'Dim OutRasterCursor(OutputRasters.Count - 1) As ESRI.ArcGIS.Geodatabase.IRasterCursor 'Defines a raster cursor four output rasters JBB
        'Dim OutRasterEdit(OutputRasters.Count - 1) As ESRI.ArcGIS.DataSourcesRaster.IRasterEdit 'Defines a raster edit interface - allows editing of a raster -  for output rasters JBB
        'For R = 0 To OutputRasters.Count - 1
        '    OutRasterBand(R) = OutputRasters(R) 'Copies Output Rasters into the output raster band collection JBB
        '    OutRaster2(R) = CType(CType(OutputRasters(R).CreateDefaultRaster, ESRI.ArcGIS.Geodatabase.IRaster), ESRI.ArcGIS.DataSourcesRaster.IRaster2) 'Why is CreateDefaultRaster needed and why is .IRaster and .IRaster2 both used JBB
        '    OutRasterCursor(R) = OutRaster2(R).CreateCursorEx(Nothing) 'Sets the output raster cursor with initial values of nothing JBB
        '    OutRasterEdit(R) = OutRaster2(R) 'Copies output rasters into the raster edit thing JBB
        'Next
        '****Start JBB Added Code
        '**** output debug raster
        'For multiple output dbg rasters
        Dim OutputRastersDbg(OutputRasterNames.Count - 1, WBOutDateCount - 1) As ESRI.ArcGIS.Geodatabase.IRasterDataset 'Makes a raster dataset for the output debug raster JBB
        For I = 0 To OutputRasterNames.Count - 1 'Loops Through Selected Output Variables
            For R = 0 To WBOutDateCount - 1 'Loops Through Selecte Output Dates
                Dim RecordDateDbg As Date = WBOutDatesARR(R)
                Dim FileNameDbg As String = "SETMI_" & WaterBalanceModel.ToString & OutputImagesBoxWater.CheckedItems(I).ToString.Split("(")(1).Replace(")", "") & " -Dbg- " & Format(RecordDateDbg, "MM-dd-yyyy HH-mm") & ".img" 'Creates file name string for output debug file JBB
                If ExistsArcGISFile(OutputDirectoryTextWater.Text & "\" & FileNameDbg) Then DeleteArcGISFile(OutputDirectoryTextWater.Text & "\" & FileNameDbg) 'Deletes existing output Debug file JBB

                Dim OutputRasterDatasetDbg = Workspace.CreateRasterDataset(Name:=FileNameDbg,
                                                            Format:="IMAGINE Image",
                                                            Origin:=IntersectRasterValue.Extent.LowerLeft,
                                                            columnCount:=IntersectRasterValue.Extent.Width / IntersectRasterValue.RasterStorageDef.CellSize.X,
                                                            RowCount:=IntersectRasterValue.Extent.Height / IntersectRasterValue.RasterStorageDef.CellSize.Y,
                                                            cellSizeX:=IntersectRasterValue.RasterStorageDef.CellSize.X,
                                                            cellSizeY:=IntersectRasterValue.RasterStorageDef.CellSize.Y,
                                                            numBands:=1,
                                                            PixelType:=ESRI.ArcGIS.Geodatabase.rstPixelType.PT_FLOAT,
                                                            SpatialReference:=IntersectRasterValue.Extent.SpatialReference,
                                                            Permanent:=True) 'Creates and sets up output Debug raster files JBB
                OutputRastersDbg(I, R) = OutputRasterDatasetDbg 'Puts newly created output debug raster into the output rasters raster dataset JBB
            Next
        Next

        Dim OutRasterBandDbg(OutputRasterNames.Count - 1, WBOutDateCount - 1) As ESRI.ArcGIS.DataSourcesRaster.IRasterBandCollection 'Defines a raster band collection for the output Debug raster JBB
        Dim OutRaster2Dbg(OutputRasterNames.Count - 1, WBOutDateCount - 1) As ESRI.ArcGIS.DataSourcesRaster.IRaster2  'Defines a raster interface for output Debug raster JBB
        Dim OutRasterCursorDbg(OutputRasterNames.Count - 1, WBOutDateCount - 1) As ESRI.ArcGIS.Geodatabase.IRasterCursor 'Defines a raster cursor four output Debug raster JBB
        Dim OutRasterEditDbg(OutputRasterNames.Count - 1, WBOutDateCount - 1) As ESRI.ArcGIS.DataSourcesRaster.IRasterEdit 'Defines a raster edit interface - allows editing of a raster -  for output Debug raster JBB

        For I = 0 To OutputRasterNames.Count - 1 'Loop Through Selected Output Variables JBB
            For R = 0 To WBOutDateCount - 1 'Loop Through Desired Output Dates JBB
                OutRasterBandDbg(I, R) = OutputRastersDbg(I, R) 'Defines a raster band collection for the output Debug raster JBB
                OutRaster2Dbg(I, R) = CType(CType(OutputRastersDbg(I, R).CreateDefaultRaster, ESRI.ArcGIS.Geodatabase.IRaster), ESRI.ArcGIS.DataSourcesRaster.IRaster2) 'Defines a raster interface for output Debug raster JBB
                OutRasterCursorDbg(I, R) = OutRaster2Dbg(I, R).CreateCursorEx(Nothing) 'Defines a raster cursor four output Debug raster JBB
                OutRasterEditDbg(I, R) = OutRaster2Dbg(I, R) 'Defines a raster edit interface - allows editing of a raster -  for output Debug raster JBB
            Next
        Next
        '****End JBB Added Code

        If Abort = True Then Exit Sub : ProgressAllWater.PerformStep() 'Exits is the abort button is clicked on the form JBB - !!!! Was Line 2951 !!!!

        Dim SelectedLatitude As New List(Of Double)
        Dim SelectedLongitude As New List(Of Double)
        If WaterBalanceModel = WaterBalance.Crop_Coefficient_Point Then 'Grabs coordinates for the selected calculation points JBB
            If WBPntLocMeth = WBPntMethod.Both_Manual_and_Tabular Or WBPntLocMeth = WBPntMethod.Manually_Selected Then '*****Added by JBB to add tabular coordinates copied from code above*****
                For Each Row As System.Windows.Forms.DataGridViewRow In CalculationCoordinatesGrid.Rows
                    SelectedLongitude.Add(Row.Cells(0).Value)
                    SelectedLatitude.Add(Row.Cells(1).Value)
                Next
            End If
            If WBPntLocMeth = WBPntMethod.Both_Manual_and_Tabular Or WBPntLocMeth = WBPntMethod.Tabular_Input Then '*****Added by JBB to add tabular coordinates copied from code above*****
                '*****Added by JBB to add tabular coordinates copied from code above*****
                For JJJ = 0 To PntCount - 1
                    SelectedLongitude.Add(WBPntOutLocsARR(JJJ, 0))
                    SelectedLatitude.Add(WBPntOutLocsARR(JJJ, 1))
                Next JJJ
                '*****End Added by JBB*****
            End If
        End If

        Dim DebugPath As String = OutputDirectoryTextWater.Text & "\SETMI Water Balance Output" 'Creates a new folder in the output directory for Debug files, I don't think any files ever get put in it JBB
        If Not IO.Directory.Exists(DebugPath) Then IO.Directory.CreateDirectory(DebugPath) 'Only creates the debug folder if it doesn't already exist JBB
        Dim PixelLength As Integer = ((IntersectRasterValue.Extent.Width / IntersectRasterValue.RasterStorageDef.CellSize.X) * (IntersectRasterValue.Extent.Height / IntersectRasterValue.RasterStorageDef.CellSize.Y)).ToString.Length 'Calculates the number of pixels in the extent JBB
        Do 'Loop through input raster Cursor
            Dim IntersectPixelBlock As ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3 = IntersectRasterCursor.PixelBlock 'Defines and sets a pixelblock array for the intersect raster cursor JBB
            Dim MultispectralPixelBlock(MultispectralImages.Count - 1) As ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3 'Defines a pixelblock array for the input multispectral rasters JBB
            Dim CoverClassificationPixelBlock(MultispectralImages.Count - 1) As ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3 'Defines a pixelblock array for the input cover rasters JBB
            For m = 0 To MultispectralImages.Count - 1
                MultispectralPixelBlock(m) = InRasterCursor(m).PixelBlock() 'Sets the input multispectral raster pixel block JBB
                CoverClassificationPixelBlock(m) = InRasterCursor(MultispectralImages.Count + CoverClassificationIndex(m)).PixelBlock() 'Sets the cover raster pixel block pairing the cover images with the multispectral images JBB
            Next
            '***Below commented out by JBB
            'Dim FieldCapacityPixPixelBlock As ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3 = InRasterCursor(InputRasterPath.Count - 2).PixelBlock 'Defines and sets a pixel block for field capacity JBB
            'Dim WiltingPointPixPixelBlock As ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3 = InRasterCursor(InputRasterPath.Count - 1).PixelBlock 'Defines and sets a pixel block for wilting point JBB
            '***added by JBB copying code above, etc.
            'Dim FieldCapacityPixPixelBlock As ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3 = InRasterCursor(InputRasterPath.Count - 3).PixelBlock 'Defines and sets a pixel block for field capacity JBB
            'Dim WiltingPointPixPixelBlock As ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3 = InRasterCursor(InputRasterPath.Count - 2).PixelBlock 'Defines and sets a pixel block for wilting point JBB
            'Dim InitialSoilMoisturePixPixelBlock As ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3 = InRasterCursor(InputRasterPath.Count - 1).PixelBlock 'Defines and sets a pixel block for initial soil water content JBB

            Dim FieldCapacityPixPixelBlock As ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3 = InRasterCursor(InputRasterPath.Count - 4 + Ilyr1).PixelBlock 'Defines and sets a pixel block for field capacity JBB modified by HBB 2/9/2018
            Dim WiltingPointPixPixelBlock As ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3 = InRasterCursor(InputRasterPath.Count - 3 + Ilyr1).PixelBlock 'Defines and sets a pixel block for wilting point JBB modified by HBB 2/9/2018
            Dim InitialSoilMoisturePixPixelBlock As ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3 = InRasterCursor(InputRasterPath.Count - 2 + Ilyr1).PixelBlock 'Defines and sets a pixel block for initial soil water content JBB modified by HBB 2/9/2018
            Dim LayerThicknessPixPixelBlock As ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3 = InRasterCursor(InputRasterPath.Count - 1 + Ilyr1).PixelBlock 'added by JBB copied from code above

            Dim FieldCapacityPixels2 As System.Array
            Dim WiltingPointPixels2 As System.Array
            Dim InitialSoilMoisturePixels2 As System.Array
            Dim LayerThicknessPixels2 As System.Array
            Dim FieldCapacityPixels3 As System.Array
            Dim WiltingPointPixels3 As System.Array
            Dim InitialSoilMoisturePixels3 As System.Array
            Dim LayerThicknessPixels3 As System.Array
            Dim HydraulicConductivityPixels As System.Array
            Dim HydraulicConductivityPixels2 As System.Array
            Dim HydraulicConductivityPixels3 As System.Array
            Dim SaturatedWaterContentPixels As System.Array
            Dim SaturatedWaterContentPixels2 As System.Array
            Dim SaturatedWaterContentPixels3
            Dim REWPixels As System.Array
            Dim DPlimitPixels As System.Array

            If Ilyr2 < 1 Then
                Dim FieldCapacityPixPixelBlock2 As ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3 = InRasterCursor(InputRasterPath.Count - 4 + Ilyr2).PixelBlock 'Defines and sets a pixel block for field capacity JBB added by JBB copied from code above
                Dim WiltingPointPixPixelBlock2 As ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3 = InRasterCursor(InputRasterPath.Count - 3 + Ilyr2).PixelBlock 'Defines and sets a pixel block for wilting point JBB added by JBB copied from code above
                Dim InitialSoilMoisturePixPixelBlock2 As ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3 = InRasterCursor(InputRasterPath.Count - 2 + Ilyr2).PixelBlock 'Defines and sets a pixel block for initial soil water content JBB added by JBB copied from code above
                Dim LayerThicknessPixPixelBlock2 As ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3 = InRasterCursor(InputRasterPath.Count - 1 + Ilyr2).PixelBlock 'added by JBB copied from code above
                FieldCapacityPixels2 = CType(FieldCapacityPixPixelBlock2.PixelData(0), System.Array) 'added by JBB copied from code above
                WiltingPointPixels2 = CType(WiltingPointPixPixelBlock2.PixelData(0), System.Array) 'added by JBB copied from code above
                InitialSoilMoisturePixels2 = CType(InitialSoilMoisturePixPixelBlock2.PixelData(0), System.Array) 'added by JBB copied from code above
                LayerThicknessPixels2 = CType(LayerThicknessPixPixelBlock2.PixelData(0), System.Array) 'added by JBB copied from code above
            End If
            If ILyr3 < 1 Then
                Dim FieldCapacityPixPixelBlock3 As ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3 = InRasterCursor(InputRasterPath.Count - 4 + ILyr3).PixelBlock 'Defines and sets a pixel block for field capacity JBB added by JBB copied from code above
                Dim WiltingPointPixPixelBlock3 As ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3 = InRasterCursor(InputRasterPath.Count - 3 + ILyr3).PixelBlock 'Defines and sets a pixel block for wilting point JBB added by JBB copied from code above
                Dim InitialSoilMoisturePixPixelBlock3 As ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3 = InRasterCursor(InputRasterPath.Count - 2 + ILyr3).PixelBlock 'Defines and sets a pixel block for initial soil water content JBB added by JBB copied from code above
                Dim LayerThicknessPixPixelBlock3 As ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3 = InRasterCursor(InputRasterPath.Count - 1 + ILyr3).PixelBlock 'added by JBB copied from code above
                FieldCapacityPixels3 = CType(FieldCapacityPixPixelBlock3.PixelData(0), System.Array) 'added by JBB copied from code above
                WiltingPointPixels3 = CType(WiltingPointPixPixelBlock3.PixelData(0), System.Array) 'added by JBB copied from code above
                InitialSoilMoisturePixels3 = CType(InitialSoilMoisturePixPixelBlock3.PixelData(0), System.Array) 'added by JBB copied from code above
                LayerThicknessPixels3 = CType(LayerThicknessPixPixelBlock3.PixelData(0), System.Array) 'added by JBB copied from code above
            End If

            If IKs1 < 1 Then
                Dim HydraulicConductivityPixPixelBlock As ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3 = InRasterCursor(InputRasterPath.Count - 1 + IKs1).PixelBlock 'added by JBB copied from code above
                HydraulicConductivityPixels = CType(HydraulicConductivityPixPixelBlock.PixelData(0), System.Array) 'added by JBB copied from code above
            End If
            If IKs2 < 1 Then
                Dim HydraulicConductivityPixPixelBlock2 As ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3 = InRasterCursor(InputRasterPath.Count - 1 + IKs2).PixelBlock 'added by JBB copied from code above
                HydraulicConductivityPixels2 = CType(HydraulicConductivityPixPixelBlock2.PixelData(0), System.Array) 'added by JBB copied from code above
            End If
            If IKs3 < 1 Then
                Dim HydraulicConductivityPixPixelBlock3 As ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3 = InRasterCursor(InputRasterPath.Count - 1 + IKs3).PixelBlock 'added by JBB copied from code above
                HydraulicConductivityPixels3 = CType(HydraulicConductivityPixPixelBlock3.PixelData(0), System.Array) 'added by JBB copied from code above
            End If
            If IVWCs1 < 1 Then
                Dim SaturatedWaterContentPixPixelBlock As ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3 = InRasterCursor(InputRasterPath.Count - 1 + IVWCs1).PixelBlock 'added by JBB copied from code above
                SaturatedWaterContentPixels = CType(SaturatedWaterContentPixPixelBlock.PixelData(0), System.Array) 'added by JBB copied from code above
            End If
            If IVWCs2 < 1 Then
                Dim SaturatedWaterContentPixPixelBlock2 As ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3 = InRasterCursor(InputRasterPath.Count - 1 + IVWCs2).PixelBlock 'added by JBB copied from code above
                SaturatedWaterContentPixels2 = CType(SaturatedWaterContentPixPixelBlock2.PixelData(0), System.Array) 'added by JBB copied from code above
            End If
            If IVWCs3 < 1 Then
                Dim SaturatedWaterContentPixPixelBlock3 As ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3 = InRasterCursor(InputRasterPath.Count - 1 + IVWCs3).PixelBlock 'added by JBB copied from code above
                SaturatedWaterContentPixels3 = CType(SaturatedWaterContentPixPixelBlock3.PixelData(0), System.Array) 'added by JBB copied from code above
            End If
            If Irew < 1 Then
                Dim REWPixPixelBlock As ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3 = InRasterCursor(InputRasterPath.Count - 1 + Irew).PixelBlock 'added by JBB copied from code above
                REWPixels = CType(REWPixPixelBlock.PixelData(0), System.Array) 'added by JBB copied from code above
            End If
            If Idpl < 1 Then
                Dim DPlimitPixPixelBlock As ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3 = InRasterCursor(InputRasterPath.Count - 1 + Idpl).PixelBlock 'added by JBB copied from code above
                DPlimitPixels = CType(DPlimitPixPixelBlock.PixelData(0), System.Array) 'added by JBB copied from code above
            End If
            '***End added by JBB

            Dim IntersectPixels As System.Array = CType(IntersectPixelBlock.PixelData(0), System.Array) 'Creates an array matching the intersect image pixel block JBB
            Dim MultispectralPixels(MultispectralImages.Count - 1, MultispectralPixelBlock(0).Planes - 1) As System.Array 'Defines an array for the multispectral image pixel blocks, from ESRI "The Planes argument returns the number of bands in the PixelBlock"  JBB
            Dim CoverPixels(MultispectralImages.Count - 1) As System.Array 'Defines an array for the cover image pixel blocks JBB
            For m = 0 To MultispectralImages.Count - 1
                For P = 0 To MultispectralPixelBlock(m).Planes - 1
                    MultispectralPixels(m, P) = CType(MultispectralPixelBlock(m).PixelData(P), System.Array) 'Populates multispectral image array JBB
                Next
                CoverPixels(m) = CType(CoverClassificationPixelBlock(m).PixelData(0), System.Array) 'Populates cover image array JBB
            Next
            Dim FieldCapacityPixels As System.Array = CType(FieldCapacityPixPixelBlock.PixelData(0), System.Array) 'Defines and creates an array matching the field capacity image pixel block JBB
            Dim WiltingPointPixels As System.Array = CType(WiltingPointPixPixelBlock.PixelData(0), System.Array) 'Defines and creates an array matching the wilting point image pixel block JBB
            Dim InitialSoilMoisturePixels As System.Array = CType(InitialSoilMoisturePixPixelBlock.PixelData(0), System.Array) 'added by JBB for initial soil moisture JBB
            Dim LayerThicknessPixels As System.Array = CType(LayerThicknessPixPixelBlock.PixelData(0), System.Array) 'added by JBB copied from code above

            Dim WeatherETDailyActualPixels(Math.Max(WeatherGrid.ETDailyActual.Count - 1, 0)) As System.Array 'Defines and creates an array matching the ET image pixel blocks JBB
            For W = 0 To WeatherGrid.ETDailyActual.Count - 1 'Populates the array matching the ET image pixel blocks JBB
                WeatherETDailyActualPixels(W) = CType(CType(InRasterCursor(WeatherOffset + W).PixelBlock, ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3).PixelData(0), System.Array) 'Original Code Commented out by JBB because it was grabbing Field Capacity instead of daily ET JBB
                'WeatherETDailyActualPixels(W) = CType(CType(InRasterCursor(WeatherOffset + W).PixelBlock, ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3).PixelData(0), System.Array) 'Adjusted code by JBB
            Next
            '********Start JBB Added Code
            'For irrigation, depletion, and precipitation images
            Dim WeatherPrecipitationDailyPixels(Math.Max(WeatherGrid.Precipitation.Count - 1, 0)) As System.Array 'Defines and creates an array matching the ET image pixel blocks JBB
            For W = 0 To WeatherGrid.Precipitation.Count - 1 'Populates the array matching the precip image pixel blocks JBB
                WeatherPrecipitationDailyPixels(W) = CType(CType(InRasterCursor(PrecipitationOffset + W).PixelBlock, ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3).PixelData(0), System.Array) 'Original Code Commented out by JBB because it was grabbing Field Capacity instead of daily ET JBB
            Next
            Dim WeatherIrrigationDailyPixels(Math.Max(WeatherGrid.Irrigation.Count - 1, 0)) As System.Array 'Defines and creates an array matching the ET image pixel blocks JBB
            For W = 0 To WeatherGrid.Irrigation.Count - 1 'Populates the array matching the irrig. image pixel blocks JBB
                WeatherIrrigationDailyPixels(W) = CType(CType(InRasterCursor(IrrigationOffset + W).PixelBlock, ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3).PixelData(0), System.Array) 'Original Code Commented out by JBB because it was grabbing Field Capacity instead of daily ET JBB
            Next
            Dim WeatherDepletionDailyPixels(Math.Max(WeatherGrid.Depletion.Count - 1, 0)) As System.Array 'Defines and creates an array matching the ET image pixel blocks JBB
            For W = 0 To WeatherGrid.Depletion.Count - 1 'Populates the array matching the depl. image pixel blocks JBB
                WeatherDepletionDailyPixels(W) = CType(CType(InRasterCursor(DepletionOffset + W).PixelBlock, ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3).PixelData(0), System.Array) 'Original Code Commented out by JBB because it was grabbing Field Capacity instead of daily ET JBB
            Next
            Dim WeatherThetaVPixels(Math.Max(WeatherGrid.ThetaV.Count - 1, 0)) As System.Array 'Defines and creates an array matching the ET image pixel blocks JBB
            For W = 0 To WeatherGrid.ThetaV.Count - 1 'Populates the array matching the depl. image pixel blocks JBB
                WeatherThetaVPixels(W) = CType(CType(InRasterCursor(ThetaVOffset + W).PixelBlock, ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3).PixelData(0), System.Array) 'Original Code Commented out by JBB because it was grabbing Field Capacity instead of daily ET JBB
            Next
            Dim WeatherEvaporatedDepthPixels(Math.Max(WeatherGrid.EvaporatedDepth.Count - 1, 0)) As System.Array 'Defines and creates an array matching the ET image pixel blocks JBB
            For W = 0 To WeatherGrid.EvaporatedDepth.Count - 1 'Populates the array matching the depl. image pixel blocks JBB
                WeatherEvaporatedDepthPixels(W) = CType(CType(InRasterCursor(EvaporatedDepthOffset + W).PixelBlock, ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3).PixelData(0), System.Array) 'Original Code Commented out by JBB because it was grabbing Field Capacity instead of daily ET JBB
            Next

            Dim WeatherSkinEvapDepthPixels(Math.Max(WeatherGrid.SkinEvapDepth.Count - 1, 0)) As System.Array 'Defines and creates an array matching the ET image pixel blocks JBB
            For W = 0 To WeatherGrid.SkinEvapDepth.Count - 1 'Populates the array matching the depl. image pixel blocks JBB
                WeatherSkinEvapDepthPixels(W) = CType(CType(InRasterCursor(SkinEvapDepthOffset + W).PixelBlock, ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3).PixelData(0), System.Array) 'added by JBB copying others 2/9/2018
            Next


            'For the multiple test debug images
            Dim OutRasterPixelBlockDbg(OutputRasterNames.Count - 1, WBOutDateCount - 1) As ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3
            Dim OutPixelsDbg(OutputRasterNames.Count - 1, WBOutDateCount - 1) As System.Array

            For I = 0 To OutputRasterNames.Count - 1
                For R = 0 To WBOutDateCount - 1 'MultispectralImages.Count - 1
                    OutRasterPixelBlockDbg(I, R) = OutRasterCursorDbg(I, R).PixelBlock
                    OutPixelsDbg(I, R) = CType(OutRasterPixelBlockDbg(I, R).PixelData(0), System.Array)
                Next
            Next

            ''For the output images
            ''copied from energy balance
            'Dim OutRasterPixelBlock(OutputRasters.Count - 1) As ESRI.ArcGIS.DataSourcesRaster.IPixelBlock3'This Block Commented out by JBB
            'Dim OutPixels(OutputRasters.Count - 1) As System.Array
            'For i = 0 To OutputRasters.Count - 1
            '    OutRasterPixelBlock(i) = OutRasterCursor(i).PixelBlock
            '    OutPixels(i) = CType(OutRasterPixelBlock(i).PixelData(0), System.Array)
            'Next
            '***********End JBB Added Code
            For Row As Long = 0 To MultispectralPixelBlock(0).Height - 1 'Loop through images JBB
                For Col As Long = 0 To MultispectralPixelBlock(0).Width - 1 'Loop through images JBB

                    If IntersectPixels.GetValue(Col, Row) = 1 Then 'If the all images intersect then proceed JBB
                        Dim CoverIndex As Integer = CoverValues.IndexOf(CoverPixels(0).GetValue(Col, Row)) 'CoverValues is the raster integer value for the covers, This line determines what the cover for the pixel is JBB
                        If CoverIndex > -1 Then 'Cover index will be -1 if the string: CoverPixels(0).GetValue(Col, Row) doesn't exist JBB
                            If CoverSelectionGrid.Rows(CoverIndex).Cells(3).Value = True Then 'If the cover is selected as "included" on the SETMI Cover Tab JBB
                                Dim PixelLatitude As Single = 0 'Define pixel lat JBB
                                Dim PixelLongitude As Single = 0 'Define pixel lat JBB
                                IntersectRaster2.PixelToMap(InRasterCursor(0).TopLeft.X + Col, InRasterCursor(0).TopLeft.Y + Row, PixelLongitude, PixelLatitude) 'From ESRI PixelToMap "Converts a location (column, row) in pixel space into map space" JBB

                                Dim PixelIndex As Integer = -1
                                Try
                                    For S = 0 To SelectedLatitude.Count - 1
                                        If Math.Abs(PixelLatitude - SelectedLatitude(S)) < IntersectRasterValue.RasterStorageDef.CellSize.X / 2 Then 'If the center of the pixel and the selected calc point are less than one half of a pixel height apart then the latitude is correct JBB
                                            If Math.Abs(PixelLongitude - SelectedLongitude(S)) < IntersectRasterValue.RasterStorageDef.CellSize.Y / 2 Then 'If the center of the pixel and the selected calc point are less than one half of a pixel width apart then the longitude is correct JBB
                                                PixelIndex = S 'Sets the index equal to the row of the selected pixel, the first selected pixel is 0, and so on JBB
                                                Exit For
                                            End If
                                        End If
                                    Next
                                Catch ex As Exception
                                    MsgBox(ex.Message) 'If there is an error it flashes a message box with the exception, not sure what the message would be JBB
                                End Try


                                If PixelIndex > -1 Then 'If the selected pixel has been found'<--Uncomment for single pixel JBB
                                    PixelIndex = PixelIndex 'for debugging
                                End If

                                Dim Output As New System.Text.StringBuilder
                                Dim Cover As Cover = DirectCast([Enum].Parse(GetType(Cover), CoverSelectionGrid.Rows(CoverIndex).Cells(1).Value.Replace(" ", "_")), Cover) 'Grabs cover JBB
                                Dim IrrigationMethod As IrrigationMethod = DirectCast([Enum].Parse(GetType(IrrigationMethod), CoverSelectionGrid.Rows(CoverIndex).Cells(2).Value.Replace(" ", "_")), IrrigationMethod) 'Grabs irrigation method JBB
                                'Added by JBB
                                Dim FcVI As FcVI = DirectCast([Enum].Parse(GetType(FcVI), CoverSelectionGrid.Rows(CoverIndex).Cells(24).Value.Replace(" ", "_")), KcbVI) 'Grabs Kcb Veg Index, if inputting Kcb relationship, added by JBB copying code above for irrigation method
                                Dim MinVIInput As Single = CoverSelectionGrid.Rows(CoverIndex).Cells(25).Value 'Grabs input Fc Min NDVI from table, added by JBB
                                Dim MaxVIInput As Single = CoverSelectionGrid.Rows(CoverIndex).Cells(26).Value 'Grabs input Fc Max NDVI from table, added by JBB
                                Dim ExpVIInput As Single = CoverSelectionGrid.Rows(CoverIndex).Cells(27).Value 'Grabs input Fc NDVI exponent from table, added by JBB
                                Dim KcbVI As KcbVI = DirectCast([Enum].Parse(GetType(KcbVI), CoverSelectionGrid.Rows(CoverIndex).Cells(28).Value.Replace(" ", "_")), KcbVI) 'Grabs Kcb Veg Index, if inputting Kcb relationship, added by JBB copying code above for irrigation method
                                Dim UsedKcbVI As AllKcbVI 'added2/8/2018
                                Dim KcbrfSlope As Single = CoverSelectionGrid.Rows(CoverIndex).Cells(29).Value 'Grabs input Fc NDVI exponent from table, added by JBB
                                Dim KcbrfIntercept As Single = CoverSelectionGrid.Rows(CoverIndex).Cells(30).Value 'Grabs input Fc Min SAVI from table, added by JBB
                                Dim MinKcbrfInput As Single = CoverSelectionGrid.Rows(CoverIndex).Cells(31).Value 'Grabs input Fc Max SAVI from table, added by JBB
                                Dim MaxKcbrfInput As Single = CoverSelectionGrid.Rows(CoverIndex).Cells(32).Value 'Grabs input Fc SAVI exponent 

                                If FlgVI = False Then
                                    If (KcbVI = KcbVI.Hardcoded And FcVI <> FcVI.Hardcoded) Or (KcbVI = KcbVI.SAVI And FcVI <> FcVI.SAVI) Or (KcbVI = KcbVI.NDVI And FcVI <> FcVI.NDVI) Then
                                        MsgBox("Input Kcb Veg. Index and Input Fc Veg. Index must be the same for water balance computations. The Kcb Veg Index will be used. Fc solutions will be erroneous unless it is set to default values by setting at least one Fc input parameter less than or equal to zero.")
                                        FlgVI = True
                                    End If
                                End If
                                If KcbVI = KcbVI.Hardcoded Then 'added 2/8/2018
                                    UsedKcbVI = FindKcbVI(Cover, ETReference)
                                Else
                                    If KcbVI = KcbVI.SAVI Then UsedKcbVI = AllKcbVI.SAVI
                                    If KcbVI = KcbVI.NDVI Then UsedKcbVI = AllKcbVI.NDVI
                                End If

                                'End Added by JBB
                                Dim ICover As Integer = CoverPoint.CoverName.IndexOf(CoverSelectionGrid.Rows(CoverIndex).Cells(0).Value) 'Finds the cover index number from the SETMI interface JBB
                                Dim MultispectralDate(MultispectralImages.Count - 1) As DateTime 'Defines an array to store the image dates, one item for each image JBB
                                Dim Red(MultispectralImages.Count - 1) As Single 'Defines an array for RED reflectance, one item for each image  JBB
                                Dim NIR(MultispectralImages.Count - 1) As Single 'Defines an array for NIR reflectance, one item for each image  JBB
                                Dim DateReflectance As New List(Of Integer) 'Defines and array to store the DOY for each image JBB
                                Dim KcbReflectance(MultispectralImages.Count - 1) As Single 'Makes an array for Kcbrf, one item for each image  JBB
                                Dim FractionOfCover(MultispectralImages.Count - 1) As Single 'Makes and array for Fraction of cover, one item for each image JBB
                                Dim YearStartWB As Integer = CoverPoint.YearStartWB(ICover) 'Grabs the year of the first year of the WB from input JBB added by JBB
                                If YearStartWB < 1900 Then
                                    YearStartWB = GetDateFromPath(MultispectralImages(0)).Year
                                End If
                                Dim SAVIs(MultispectralImages.Count - 1) As Single 'Makes an array for SAVI, one item for each image for SAVI Log Interp JBB
                                '******degbug only
                                Dim NDVIs(MultispectralImages.Count - 1) As Single 'Added by JBB could be used for Fc later
                                '******end debug only

                                For m = 0 To MultispectralImages.Count - 1
                                    MultispectralDate(m) = GetDateFromPath(MultispectralImages(m)) 'Grabs the date from the image name, includes the time JBB
                                    'DateReflectance.Add(MultispectralDate(m).Subtract(DateSerial(MultispectralDate(m).Year - 1, 12, 31)).TotalDays) 'Finds the DOY for each image (this works!) JBB
                                    DateReflectance.Add(MultispectralDate(m).Subtract(DateSerial(YearStartWB - 1, 12, 31)).TotalDays) 'Finds the DOY for each image modified by JBB to allow for more than one year
                                    If RedIndex.Text <> "" Then Red(m) = MultispectralPixels(m, RedIndex.Text - 1).GetValue(Col, Row) 'Red Index is the red band selection box on the Surface tab JBB
                                    If NIRIndex.Text <> "" Then NIR(m) = MultispectralPixels(m, NIRIndex.Text - 1).GetValue(Col, Row) 'NIR Index is the NIR band selection box on the Surface tab JBB
                                    Dim SAVI As Single = calcSAVI(Red(m), NIR(m)) 'Calc SAVI JBB
                                    SAVIs(m) = SAVI 'added by JBB for the SAVI Log Interp Kcb JBB
                                    Dim NDVI As Single = calcNDVI(Red(m), NIR(m)) 'Calc NDVI JBB
                                    NDVIs(m) = NDVI
                                    '*********************************************************************************
                                    '*********************************************************************************
                                    If UsedKcbVI = AllKcbVI.NDVI Then SAVIs(m) = NDVI 'overwrites SAVI with NDVI for Kcb, etc.
                                    '*********************************************************************************
                                    '*********************************************************************************
                                    'KcbReflectance(m) = calcKcbReflectance(Cover, SAVI) 'Corn and Soybeans match Hatim's Dissertation JBB commented out by JBB
                                    KcbReflectance(m) = calcKcbReflectance(Cover, SAVI, ETReference, NDVI, KcbVI, KcbrfSlope, KcbrfIntercept, MaxKcbrfInput, MinKcbrfInput) 'added by JBB to include alfalfa reference JBB
                                    'Dim LAI As Single ' not used here but the function is used in the energy balance and it requires LAI
                                    'FractionOfCover(m) = calcFc(Cover, NDVI, SAVI, LAI) 'Calcs Fc based on NDVI, other parameters don't matter for corn or soybean JBB
                                    FractionOfCover(m) = calcFcWB(Cover, NDVI, SAVI, KcbVI, MaxVIInput, MinVIInput, ExpVIInput, 0.99, 0) 'Added by JBB
                                    If SAVI <= 0 Or NDVI <= 0 Then
                                        If NegFlag = False Then 'added by JBB to flag negative VIs
                                            NegFlag = True
                                        End If
                                    End If
                                Next

                                'Dim KcbIni As Single = CoverPoint.KcbInitial(ICover) 'Grabs Kcb's and Stage Lengths from cover tab on form JBB 'not used anymore, but available for future use JBB
                                'Dim KcbMid As Single = CoverPoint.KcbMid(ICover) 'Grabs Kcb's and Stage Lengths from cover tab on form JBB 'not used anymore, but available for future use JBB
                                'Dim KcbEnd As Single = CoverPoint.KcbEnd(ICover) 'Grabs Kcb's and Stage Lengths from cover tab on form JBB 'not used anymore, but available for future use JBB
                                'Dim LIni As Integer = CoverPoint.PeriodInitial(ICover) 'Grabs Kcb's and Stage Lengths from cover tab on form JBB 'not used anymore, but available for future use JBB
                                'Dim LDev As Integer = CoverPoint.PeriodDevelopment(ICover) 'Grabs Kcb's and Stage Lengths from cover tab on form JBB 'not used anymore, but available for future use JBB
                                'Dim LMid As Integer = CoverPoint.PeriodMid(ICover) 'Grabs Kcb's and Stage Lengths from cover tab on form JBB 'not used anymore, but available for future use JBB
                                'Dim LLate As Integer = CoverPoint.PeriodEnd(ICover) 'Grabs Kcb's and Stage Lengths from cover tab on form JBB 'not used anymore, but available for future use JBB
                                'Dim DateIni As Integer = CoverPoint.DateInitial(ICover) 'Grabs input start date from the cover tab on form JBB 'not used anymore, but available for future use JBB
                                Dim MinimumCoverHeight As Single = Limit(CoverPoint.MinimumCoverHeight(ICover), 0.1, 20) 'Min height is grabbed from the cover tab but can't be less than 0.1, which is from FAO56, but previous to 10/1/2018, this was 0.2, that value may have come from Jensen and Allen 2016 or FAO56, but JBB could not find it JBB
                                Dim MaximumCoverHeight As Single = Limit(CoverPoint.MaximumCoverHeight(ICover), 0.1, 20) 'Max height is grabbed from the cover tab but can't be less than 0.1, which is from FAO56, but previous to 10/1/2018, this was 0.2, that value may have come from Jensen and Allen 2016 or FAO56, but JBB could not find it JBB
                                Dim Zrmax As Single = CoverPoint.MaximumRootDepth(ICover) 'fRom Table 22 in FAO 56, Grabs input maximum root depth from the cover tab on the form JBB
                                Dim ZrmaxIn As Single = Zrmax
                                'Dim Zrmin As single = 0.1 'Min root depth is set at 0.1 m JBB Commented out by JBB and moved a few lines down in code JBB
                                Dim Ptab As Single = CoverPoint.P(ICover) 'Ptable from table 22 in FAO56, is grabbed from the cover tab on the form JBB
                                Dim CurveNumber As Single = CoverPoint.CurveNumber(ICover) 'Curve Number added by JBB
                                Dim AssimilationWeight As Single = IIf(CoverPointIndex.WeightForAssimilation = -1, 0.78, CoverPoint.WeightForAssimilation(ICover)) 'assimilation weight added by JBB, if none entered it uses 0.78 from Hatim's Dissertation 
                                Dim DepletionWeight As Single = IIf(CoverPointIndex.WeightForDepletion = -1, 1, CoverPoint.WeightForDepletion(ICover)) 'Assimilation weight for depletion maps added by JBB if none it uses 1, or assumed to be the truth.
                                Dim ThetaVWeight As Single = IIf(CoverPointIndex.WeightForThetaVL = -1, 1, CoverPoint.WeightForThetaVL(ICover)) 'assimilation weight for lower layer theta V, assumed to be 1 if not entered JBB
                                Dim EvaporatedDepthWeight As Single = IIf(CoverPointIndex.WeightForEvaporatedDepth = -1, 1, CoverPoint.WeightForEvaporatedDepth(ICover)) 'assimilation weight for evaporated depth, assumed to be 1 if not entered JBB
                                Dim SkinEvapWeight As Single = IIf(CoverPointIndex.WeightForSkinEvap = -1, 1, CoverPoint.WeightForSkinEvap(ICover)) 'assimilation weight for skin evaporated depth, assumed to be 1 if not entered JBB, added by JBB copying others
                                Dim PctPeff As Single = CoverPoint.PercentEffective(ICover) ' percent of precip that is effective added by JBB
                                Dim MAD As Single = CoverPoint.MAD(ICover) 'MAD added by JBB
                                Dim TargetDepthAboveMAD As Single = CoverPoint.TargetDepthAboveMAD(ICover) 'Irrigation threshold above MAD JBB
                                Dim AppEfficiency As Single = CoverPoint.ApplicationEfficiency(ICover) 'Irrigation application efficiency JBB
                                'Commented out by JBB - now computed as arrays several lines lower
                                'Dim θFC As Single ' = FieldCapacityPixels.GetValue(Col, Row) / 1000 'FC in m/m, FC raster is input as mm/m JBB
                                'Dim θWP As Single '= WiltingPointPixels.GetValue(Col, Row) / 1000 'WP in m/m, WP raster  is input as mm/m JBB
                                'Dim θvLini As Single '= Math.Max(Math.Min(InitialSoilMoisturePixels.GetValue(Col, Row) / 1000, θFC), θWP) 'initial thetav in m/m, input raster is mm/m limited to not exceed FC JBB
                                'Dim FCL As Single = 0 'added by JBB for incrementing DrLast
                                Dim WPL As Single = 0 'added by JBB for incrementing DrLast
                                Dim DelZrL As Single = 0 'added by JBB for incrementing DrLast
                                Dim DrLwrLast As Single = 0 'added by JBB for incrementing DrLast
                                Dim FCrPrev As Single = 0 'm/m added by JBB for termination
                                Dim WPrPrev As Single = 0 'm/m added by JBB for termination

                                '****Added by JBB, much is copied from above code
                                Dim ArrFC(3) As Single  'mm/m 'layered field capacity
                                Dim ArrWP(3) As Single  'mm/m ' layered wilting point
                                Dim ArrThtaIni(3) As Single  'mm/m 'layered initial water content
                                Dim ArrThick(3) As Single  'm 'layer thickness
                                Dim ArrDpth(3) As Single 'm 'layer bottom depth
                                Dim ArrSWS(3) As Single 'mm soil water storage in each layer
                                Dim ArrSWSFC(3) As Single 'mm
                                Dim ARRSWSsat(3) As Single 'mm
                                Dim DpthMax As Single 'm maximum soil depth
                                Dim ArrOldThta(3) As Single 'm/m
                                Dim ArrThtaWt(3) As Single '-

                                Dim ArrTau(3) As Single 'AquaCrop drainage parameter (Raes et al 2017 AquaCrop V.6 Manual, FAO, Rome)
                                Dim ArrDelZr(3) As Single 'm layer thickeness within Rootzone
                                Dim ArrDelZL(3) As Single 'm layer thickness below the rootzone
                                Dim ArrDelZLprev(3) As Single 'm previous lower layer thicknesses
                                Dim ArrThta(3) As Single 'm/m water content in the lower soil layers
                                Dim ArrThtaPrev(3) As Single 'm/m water content at beginning of period
                                Dim Ksat(3) As Single  'mm/day, layered satuated hydraulic conductivity
                                Dim VWCsat(3) As Single  'mm/m, layered satuated water content
                                Dim ArrDrain(3) As Single 'This is the drainage function following (Raes et al 2017 AquaCrop V.6 Manual, FAO, Rome)
                                Dim ArrDP(3) As Single 'This is an array of drainage
                                Dim DPr As Single = 0 'Deep perc from the root zone
                                Dim DPtot As Single = 0 'total deep perc
                                Dim LyrCnt As Integer = 1

                                '****Input variables
                                ArrFC(1) = FieldCapacityPixels.GetValue(Col, Row) / 1000 'FC in m/m, FC raster is input as mm/m JBB
                                ArrWP(1) = WiltingPointPixels.GetValue(Col, Row) / 1000 'WP in m/m, WP raster  is input as mm/m JBB
                                If IVWCs1 < 1 Then
                                    VWCsat(1) = Limit(SaturatedWaterContentPixels.GetValue(Col, Row) / 1000, ArrFC(1), 1)
                                Else
                                    VWCsat(1) = ArrFC(1)
                                End If
                                ArrThtaIni(1) = Math.Max(Math.Min(InitialSoilMoisturePixels.GetValue(Col, Row) / 1000, VWCsat(1)), ArrWP(1)) 'initial thetav in m/m, input raster is mm/m limited to not exceed FC JBB
                                'Added to maintain variable names previously used in the code (this will help compute TEW and REW)
                                'θFC = ArrFC(1)
                                'θWP = ArrWP(1)
                                'θvLini = ArrThtaIni(1)

                                ArrThick(1) = LayerThicknessPixels.GetValue(Col, Row)
                                If Ilyr2 < 1 Then
                                    ArrFC(2) = FieldCapacityPixels2.GetValue(Col, Row) / 1000
                                    ArrWP(2) = WiltingPointPixels2.GetValue(Col, Row) / 1000
                                    If IVWCs2 < 1 Then
                                        VWCsat(2) = Limit(SaturatedWaterContentPixels2.GetValue(Col, Row) / 1000, ArrFC(2), 1)
                                    Else
                                        VWCsat(2) = ArrFC(2)
                                    End If
                                    ArrThtaIni(2) = Math.Max(Math.Min(InitialSoilMoisturePixels2.GetValue(Col, Row) / 1000, VWCsat(2)), ArrWP(2))
                                    ArrThick(2) = LayerThicknessPixels2.GetValue(Col, Row)
                                    LyrCnt = 2
                                End If
                                If ILyr3 < 1 Then
                                    ArrFC(3) = FieldCapacityPixels3.GetValue(Col, Row) / 1000
                                    ArrWP(3) = WiltingPointPixels3.GetValue(Col, Row) / 1000
                                    If IVWCs3 < 1 Then
                                        VWCsat(3) = Limit(SaturatedWaterContentPixels3.GetValue(Col, Row) / 1000, ArrFC(3), 1)
                                    Else
                                        VWCsat(3) = ArrFC(3)
                                    End If
                                    ArrThtaIni(3) = Math.Max(Math.Min(InitialSoilMoisturePixels3.GetValue(Col, Row) / 1000, VWCsat(3)), ArrWP(3))
                                    ArrThick(3) = LayerThicknessPixels3.GetValue(Col, Row)
                                    LyrCnt = 3
                                End If
                                'Total thickness

                                If IKs1 < 1 Then Ksat(1) = HydraulicConductivityPixels.GetValue(Col, Row)
                                If IKs2 < 1 Then Ksat(2) = HydraulicConductivityPixels2.GetValue(Col, Row)
                                If IKs3 < 1 Then Ksat(3) = HydraulicConductivityPixels3.GetValue(Col, Row)

                                Dim DPupper As Single
                                If DPMethod = DPLimit.Constant_Limit Then DPupper = DPlimitPixels.GetValue(Col, Row) 'added by JBB copying code from rew
                                '***Added by JBB
                                Dim DOYStartWB As Integer = CoverPoint.DOYStartWB(ICover) 'Grabs the starting day of year for the water balance JBB
                                Dim DOYEndWB As Integer = CoverPoint.DOYEndWB(ICover) 'Grabs the ending day of year for the water balance JBB
                                Dim DOYiniMin As Integer = CoverPoint.DOYiniMin(ICover) 'Grabs the earliest day of year for kcb inititation JBB
                                Dim DOYiniMax As Integer = CoverPoint.DOYiniMax(ICover) 'Grabs the latest day of year for kcb inititation JBB
                                Dim DOYefcMin As Integer = CoverPoint.DOYefcMin(ICover) 'Grabs the earliest day of year for kcb effective full cover JBB
                                Dim DOYefcMax As Integer = CoverPoint.DOYefcMax(ICover) 'Grabs the latest day of year for kcb effective full cover JBB
                                Dim DOYtermMin As Integer = CoverPoint.DOYtermMin(ICover) 'Grabs the earliest day of year kcb termination JBB
                                Dim DOYtermMax As Integer = CoverPoint.DOYtermMax(ICover) 'Grabs the latest day of year for kcb termination JBB
                                Dim KcMaxInput As Single = CoverPoint.KcMax(ICover) 'Grabs the Input Kcmax added by JBB
                                Dim KcbOffSeason As Single = CoverPoint.KcbOffSeason(ICover) ' Grabs the Kcb for the non-growing season JBB
                                Dim FalseEndSAVI As Single = CoverPoint.FalseEndSAVI(ICover) 'grabs false end season savi
                                Dim FalseEndSAVIDOY As Integer = CoverPoint.FalseEndSAVIDOY(ICover) 'grabs doy of false end savi
                                Dim FalsePeakSAVI As Single = CoverPoint.FalsePeakSAVI(ICover) 'Grabs false peak SAVI
                                Dim FalsePeakSAVIDOY As Integer = CoverPoint.FalsePeakSAVIDOY(ICover) 'Grabs false peak SAVI doy
                                Dim FalseEndSAVIGDD As Single = CoverPoint.FalseEndSAVIGDD(ICover) 'grabs GDD of false end savi
                                Dim FalsePeakSAVIGDD As Single = CoverPoint.FalsePeakSAVIGDD(ICover) 'Grabs false peak SAVI GDD
                                Dim GDDBase As Single = CoverPoint.GDDBase(ICover) 'Grabs false peak SAVI GDD copied from others
                                Dim GDDMaxTemp As Single = CoverPoint.GDDMaxTemp(ICover) 'Grabs false peak SAVI GDD copied from others
                                Dim RatioLate As Single = CoverPoint.RatioKcbLate(ICover) 'Copied from late
                                Dim TstressBase As Single = CoverPoint.StressBaseTemp(ICover) 'Copied from Above, for Kst following Campos et al. 2018
                                Dim TstressMax As Single = CoverPoint.StressMaxTemp(ICover) 'Copied from Above, for Kst following Campos et al. 2018
                                Dim BMMaxSAVI As Single = CoverPoint.xBiomassMaxSAVI(ICover) 'Copied from above for Campos et al. 2018 method by JBB

                                If CoverPointIndex.xBiomassMaxSavi < 0 Then 'Copied from code below by JBB 8/17/2018 sets a default value
                                    BMMaxSAVI = 0.4
                                End If

                                Dim DOYini As Integer = DOYStartWB 'added by JBB
                                Dim DOYdev As Integer = DOYEndWB 'added by JBB
                                Dim DOYefc As Integer = DOYEndWB 'added by JBB
                                Dim DOYterm As Integer = DOYEndWB 'added by JBB
                                Dim DOYlate As Integer = DOYEndWB 'added by JBB
                                Dim HcIni As Single = 0 'Used in calculating Kc adjustments for ETo
                                Dim HcDev As Single = 0 'Used in calculating Kc adjustments for ETo
                                Dim HcEFC As Single = 0 'Used in calculating Kc adjustments for ETo
                                Dim U2LimitedIni As Single = 2 ' Hard coded to make no adjustment (FAO56)'over written later if data provided
                                Dim RHminLimitedIni As Single = 45 ' Hard coded to make no adjustment (FAO56) 'over written later if data provided
                                Dim U2LimitedDev As Single = 2 ' Hard coded to make no adjustment (FAO56)'over written later if data provided
                                Dim RHminLimitedDev As Single = 45 ' Hard coded to make no adjustment (FAO56) 'over written later if data provided
                                Dim U2LimitedMid As Single = 2 ' Hard coded to make no adjustment (FAO56)'over written later if data provided
                                Dim RHminLimitedMid As Single = 45 ' Hard coded to make no adjustment (FAO56) 'over written later if data provided
                                Dim U2LimitedLate As Single = 2 ' Hard coded to make no adjustment (FAO56)'over written later if data provided
                                Dim RHminLimitedLate As Single = 45 ' Hard coded to make no adjustment (FAO56) 'over written later if data provided
                                '***End added by JBB
                                'Dim Ze As single = CoverPoint.EvaporativeDepth(ICover) '= 0.1 'm 'Sets the evaporative layer depth to 0.1m, Isidro Campos Likes 0.05m better JBB commented out by JBB
                                Dim Ze As Single = CoverPoint.EvaporativeDepth(ICover) 'added by JBB so that Ze can be an input
                                '****Compute Zrmin and limit first layer thickness to include it
                                Dim Zrlim As Single = Ze * (ArrFC(1) - 0.5 * ArrWP(1)) / (ArrFC(1) - ArrWP(1)) 'Depth needed to extract TEW from Zr
                                'Dim Zrmin As Single = Math.Max(CoverPoint.MinimumRootDepth(ICover), Ze * (θFC - 0.5 * θWP) / (θFC - θWP)) 'grabs input Zrmin but won't let it be less than Ze adjusted for the tew equation so the Dr can equal De added by JBB
                                Dim Zrmin As Single = Math.Max(CoverPoint.MinimumRootDepth(ICover), Zrlim) 'grabs input Zrmin but won't let it be less than Ze adjusted for the tew equation so the Dr can equal De added by JBB
                                If ArrThick(1) < Zrlim Then 'Evaporative layer is all in the top layer
                                    If ArrThick(2) > 0 Then 'if the second layer exists
                                        ArrThick(2) = ArrThick(2) + ArrThick(1) - Zrlim
                                    End If
                                    ArrThick(1) = Zrlim
                                    FlgDepth1 = True 'Triggers a message later
                                End If
                                '*****Compute layer depths, lowest is limited to Zrmax
                                ArrDpth(0) = 0
                                For iii = 1 To LyrCnt
                                    ArrDpth(iii) = ArrDpth(iii - 1) + ArrThick(iii)
                                    If ArrDpth(iii) > Zrmax Then
                                        ArrDpth(iii) = Zrmax
                                        ArrThick(iii) = Zrmax - ArrDpth(iii - 1)
                                    End If
                                    DpthMax = ArrDpth(iii)
                                    'ArrTau(iii) = Limit(0.0866 * Ksat(iii) ^ 0.35, 0, 1) * (VWCsat(iii) - ArrFC(iii)) 'Raes et al 2017. Aqua Crop 6.0, this is the constant part of this formula, no need to compute it each time it is used, commented out since it is only used in the root zone presently.
                                Next
                                Zrmax = DpthMax '<-- roots grow as though reaching full length and get cut short when hitting limiting layer
                                'Dim DateDev As Integer = DateIni + LIni 'Original code commented out by JBB, Finds the date of the beginning of development stage, following Eq. 66 in FAO 56 this should be + LIni-1 JBB
                                'Dim DateDev As Integer = DateIni + LIni - 1 'Matches FAO 56 Eq. 66 JBB'not used anymore, but available for future use JBB
                                'Dim DateMid As Integer = DateDev + LDev 'Finds the date of the beginning of middle crop stage JBB 'not used anymore, but available for future use JBB
                                'Dim DateLate As Integer = DateMid + LMid 'Finds the data of the beginning of late stage JBB 'not used anymore, but available for future use JBB
                                ''Dim DateEnd As Integer = DateLate + LLate - 1 'Crop growth ends, Finds the last date of the growing season, commented out by JBB, see note for Lini JBB
                                'Dim DateEnd As Integer = DateLate + LLate  'Matches FAO 56 Eq. 66 JBB 'not used anymore, but available for future use JBB
                                Dim TEWo As Single = 1000 * (ArrFC(1) - 0.5 * ArrWP(1)) * Ze 'mm 'FAO 56 Eq 73 JBB ' Will be adjusted later by JBB
                                'Dim REW As single = 0.42 * TEW 'mm (Curve fit to data in FAO 56 Table 19)'original code commented out by JBB <--Actually this is curve fit with zero intercept, avg of all is 0.44*TEW, linear curve fit is REW = 0.31*TEW+2.3 Rsqr = 0.946, based on avg TEW and Avg REW JBB
                                'Dim REW As Single = Math.Max(0.31 * TEW + 2.3, 0.53 * TEW) ' add by JBB in mm, based on the curve fit described above, the max limit is to keep REW from becoming larger than the largest ratio of REW to TEW in Table 19 of FAO 56 JBB
                                Dim REWo As Single 'Will be adjusted later JBB
                                If REWText.Text <> "" Then 'Logic follows other exaplies in the code
                                    REWo = REWPixels.GetValue(Col, Row) 'added by JBB copying code from above
                                Else
                                    REWo = Limit((0.31 * TEWo + 2.3 * Ze / 0.1), 0.39 * TEWo, 0.52 * TEWo) ' add by JBB in mm, based on the curve fit described above, the max and min limits are to keep REW from exceeding the range of ratios of REW to TEW in Table 19 of FAO 56 The divide by 0.1 and mult by Ze are because Table 19 is for 10cm Ze JBB
                                End If
                                Dim TEW As Single ' = TEWo 'Initial condition added by JBB
                                Dim REW As Single '= REWo 'Initial condition added by JBB
                                Dim EvapSat As Single = (VWCsat(1) - ArrFC(1)) * Ze * 1000 'Saturated water content in the evap layer - used to conserve mass later

                                Dim De As Single = 0 'mm 'depth of evaporation, See FAO 56 Eq. 77 JBB
                                Dim DREW As Single = 0 'mm Skin evaporation, ASCE MOP 70, ed. 2, 2016 Eq. 9-29 added by JBB
                                Dim Dr As Single = 0 'mm 'Depletion JBB
                                Dim DrLast As Single = 0 'mm ' This is Dr,i-1 , see FAO 56 EQ. 85 JBB commented out by JBB
                                Dim DrLastOut As Single = 0 'mm for output added by JBB
                                Dim TAWLower As Single = 0 'mm this will help indicate how weighting works in lower root zone
                                Dim WBClosureRZ As Single = 0 'mm
                                Dim WBClosureFull As Single = 0 'mm
                                Dim DelDrlastFromDelZr As Single = 0 'mm depletion added because of root growth
                                Dim DrLimit As Single = 0 'For limiting Dr
                                Dim CR As Single = 0 'Capillary Rise into rootzone assumed to be 0 JBB
                                Dim DP As Single = 0 'out of root zone Added by JBB 
                                Dim RequiredGrossIrrigation As Single = 0 'Added by JBB, Currently This Isn't Implemented Yet
                                Dim ZriPrevious As Single = 0 'm added by JBB
                                Dim HcTabPrevious As Single = 0 'm added by JBB
                                ' Dim θvLPrevious As Single = 0 'm/m thetav of lower soil layer added by JBB
                                Dim FwPrevious As Single = 0 'added by JBB
                                Dim ThtvL As Single = 0 'm/m thetav of lower soil layer added by JBB
                                Dim DPL As Single = 0 'mm deep percolation beyond the lower end of the water balance added by JBB
                                ' Dim CRL As Single = 0 'Capillary Rise into the lower soil layer assumed to be 0 JBB
                                Dim ProfileAvgθv As Single = 0 'mm/m profile average water content JBB 
                                Dim ProfileAvgThtavPrev As Single = 0 'mm/m profile average water content at the begninning of the time stamp
                                Dim SumETcAdjOut As Single = 0 'mm of ETc adj and assimilated total added by JBB
                                Dim SumPcpOut As Single = 0 'mm total Precip for output added by JBB
                                Dim SumInetOut As Single = 0 'mm total net irrigation for output added by JBB
                                Dim SumROOut As Single = 0 'mm total Runoff for output added by JBB
                                Dim ROOut As Single = 0 'For Output
                                Dim SumKcOut As Single = 0 'unit is Kcb days I suppose
                                Dim SumDPOut As Single = 0
                                Dim SumDPLOut As Single = 0
                                Dim SumFexcOut As Single = 0
                                Dim SumFexc As Single = 0
                                Dim ArrThtaPrevOut(3) As Single
                                Dim SumETrOut As Single = 0
                                Dim SumAssimExcOut As Single = 0
                                Dim SumAssimExcLOut As Single = 0
                                Dim SumAddFromDelZr As Single = 0 'that depletion added by the growing root zone
                                Dim DrIni As Single = 0
                                Dim ProfileThetaIni As Single = 0
                                Dim CumWBClosureFull As Single = 0
                                Dim CumWBClosureRZ As Single = 0
                                Dim Kt As Single = 0 'For Transpiration from Evap Layer

                                ''Added by JBB to handle h for Kcmax correctly JBB - Not Used Any more Kc max is now a manual input!
                                'Dim DaysCropHeight As New List(Of Integer)
                                ''This is set up to do a stair step average stage crop height for each stage, rather than a smooth interpolation following the definintion in FAO 56 of Kcbmax
                                'DaysCropHeight.Add(DateIni) 'AddFAO Kcb Dates
                                'DaysCropHeight.Add(DateDev - 1)
                                'DaysCropHeight.Add(DateDev)
                                'DaysCropHeight.Add(DateEnd)
                                'DaysCropHeight.Sort()
                                'Dim CropHeights(DaysCropHeight.Count - 1) As single
                                'CropHeights(DaysCropHeight.IndexOf(DateIni)) = MinimumCoverHeight 'Average height ini stage is minimum cover height
                                'CropHeights(DaysCropHeight.IndexOf(DateDev - 1)) = MinimumCoverHeight 'Average height ini stage is minimum cover height
                                'CropHeights(DaysCropHeight.IndexOf(DateDev)) = MaximumCoverHeight 'Average max height in development is max cover height
                                'CropHeights(DaysCropHeight.IndexOf(DateEnd)) = MaximumCoverHeight 'Average height after Effect. Cov. is max cover height
                                ''End Added by JBB


                                '****JBB Commented out the additions of FAO Kcb dates and Kcbs so this array is only Kcbrf now!!! 
                                Dim Days As New List(Of Integer) 'Declares a list of days that will include FAO Kcb dates and image dates JBB
                                For Each Day As Integer In DateReflectance : Days.Add(Day) : Next 'Adds image DOY's to the list JBB
                                'If Not Days.Contains(DateIni) Then Days.Add(DateIni) 'Adds Planting date to the list JBB
                                'If Not Days.Contains(DateDev) Then Days.Add(DateDev) 'Adds date of the beginning of development stage to the list JBB
                                'If Not Days.Contains(DateMid) Then Days.Add(DateMid) 'Adds date of the beginning of mid stage to the list JBB
                                'If Not Days.Contains(DateLate) Then Days.Add(DateLate) 'Adds date of the beginning of late stage to the list JBB
                                'If Not Days.Contains(DateEnd) Then Days.Add(DateEnd) 'Adds date of harvest to the list JBB
                                Days.Sort() 'Sorts the list of kcb dates in ascending order JBB
                                Dim Kcbs(Days.Count - 1) As Single 'Defines an array of kcbs that matches the Days list JBB
                                'Kcbs(Days.IndexOf(DateIni)) = KcbIni 'Adds KcbIni to the planting date JBB
                                'Kcbs(Days.IndexOf(DateDev)) = KcbIni 'Adds KcbIni to the first day of development stage JBB
                                'Kcbs(Days.IndexOf(DateMid)) = KcbMid 'Adds KcbMid to the first day of mid stage JBB
                                'Kcbs(Days.IndexOf(DateLate)) = KcbMid 'Adds KcbMid to the first day of late stage JBB
                                'Kcbs(Days.IndexOf(DateEnd)) = KcbEnd 'Adds KcbEnd to the harvest date JBB

                                Dim EarlyImageCount As Integer = 0 'added by JBB
                                Dim LateImageFlag As Boolean = False 'added by JBB
                                Dim DOYFirst As Integer 'added by JBB
                                Dim MFirst As Integer 'added by JBB
                                Dim SSq As Single = 0 'added by JBB
                                Dim SSqMin As Single = 99999 'added by JBB
                                Dim KcbInt As Single = 0 'added by JBB

                                For m = 0 To MultispectralDate.Count - 1 'grabs reflectance based Kcb for each image JBB
                                    Kcbs(Days.IndexOf(DateReflectance(m))) = KcbReflectance(m) 'grabs reflectance based Kcb for each image JBB
                                    '****Added by JBB *****
                                    If DateReflectance(m) >= DOYiniMin And DateReflectance(m) <= DOYefcMax Then
                                        EarlyImageCount = EarlyImageCount + 1
                                        If EarlyImageCount = 1 Then
                                            DOYFirst = DateReflectance(m) 'This is only used if only one image is between the estimated EFC and Ini range JBB
                                            MFirst = m 'This is used in Kcb curve fit
                                        End If
                                    End If
                                    'If DateReflectance(m) >= DOYefcMin And DateReflectance(m) >= DOYtermMax Then LateImageFlag = True 'I may need to think this through more!
                                    '***End Added by JBB***
                                Next

                                '***********Added by JBB for irrigation, precip and depletion images****************
                                Dim DateIrrigation As New List(Of Integer)
                                Dim DatePrecipitation As New List(Of Integer)
                                Dim DateDepletion As New List(Of Integer)
                                Dim DateThetaV As New List(Of Integer)
                                Dim DateEvaporatedDepth As New List(Of Integer)
                                Dim DateSkinEvapDepth As New List(Of Integer) 'added by JBB copying others 2/9/2018

                                Dim PrecipitationDate(WeatherGrid.Precipitation.Count - 1) As DateTime
                                Dim IrrigationDate(WeatherGrid.Irrigation.Count - 1) As DateTime
                                Dim DepletionDate(WeatherGrid.Depletion.Count - 1) As DateTime
                                Dim ThetaVDate(WeatherGrid.ThetaV.Count - 1) As DateTime
                                Dim EvaporatedDepthDate(WeatherGrid.EvaporatedDepth.Count - 1) As DateTime
                                Dim SkinEvapDepthDate(WeatherGrid.SkinEvapDepth.Count - 1) As DateTime 'added by JBB copying others 2/9/2018
                                For m = 0 To WeatherGrid.Precipitation.Count - 1
                                    PrecipitationDate(m) = GetDateFromPath(WeatherGrid.Precipitation(m)) 'Grabs the date from the image name, includes the time JBB
                                    'DatePrecipitation.Add(PrecipitationDate(m).Subtract(DateSerial(PrecipitationDate(m).Year - 1, 12, 31)).TotalDays) 'Finds the DOY for each image (this works!) JBB
                                    DatePrecipitation.Add(PrecipitationDate(m).Subtract(DateSerial(YearStartWB - 1, 12, 31)).TotalDays) 'Finds the DOY for each image (this works!) JBB
                                Next
                                For m = 0 To WeatherGrid.Irrigation.Count - 1
                                    IrrigationDate(m) = GetDateFromPath(WeatherGrid.Irrigation(m)) 'Grabs the date from the image name, includes the time JBB
                                    'DateIrrigation.Add(IrrigationDate(m).Subtract(DateSerial(IrrigationDate(m).Year - 1, 12, 31)).TotalDays) 'Finds the DOY for each image (this works!) JBB
                                    DateIrrigation.Add(IrrigationDate(m).Subtract(DateSerial(YearStartWB - 1, 12, 31)).TotalDays) 'Finds the DOY for each image (this works!) JBB
                                Next
                                For m = 0 To WeatherGrid.Depletion.Count - 1
                                    DepletionDate(m) = GetDateFromPath(WeatherGrid.Depletion(m)) 'Grabs the date from the image name, includes the time JBB
                                    'DateDepletion.Add(DepletionDate(m).Subtract(DateSerial(DepletionDate(m).Year - 1, 12, 31)).TotalDays) 'Finds the DOY for each image (this works!) JBB
                                    DateDepletion.Add(DepletionDate(m).Subtract(DateSerial(YearStartWB - 1, 12, 31)).TotalDays) 'Finds the DOY for each image (this works!) JBB
                                Next
                                For m = 0 To WeatherGrid.ThetaV.Count - 1
                                    ThetaVDate(m) = GetDateFromPath(WeatherGrid.ThetaV(m)) 'Grabs the date from the image name, includes the time JBB
                                    'DateThetaV.Add(ThetaVDate(m).Subtract(DateSerial(ThetaVDate(m).Year - 1, 12, 31)).TotalDays) 'Finds the DOY for each image (this works!) JBB
                                    DateThetaV.Add(ThetaVDate(m).Subtract(DateSerial(YearStartWB - 1, 12, 31)).TotalDays) 'Finds the DOY for each image (this works!) JBB
                                Next
                                For m = 0 To WeatherGrid.EvaporatedDepth.Count - 1
                                    EvaporatedDepthDate(m) = GetDateFromPath(WeatherGrid.EvaporatedDepth(m)) 'Grabs the date from the image name, includes the time JBB
                                    'DateEvaporatedDepth.Add(EvaporatedDepthDate(m).Subtract(DateSerial(EvaporatedDepthDate(m).Year - 1, 12, 31)).TotalDays) 'Finds the DOY for each image (this works!) JBB
                                    DateEvaporatedDepth.Add(EvaporatedDepthDate(m).Subtract(DateSerial(YearStartWB - 1, 12, 31)).TotalDays) 'Finds the DOY for each image (this works!) JBB
                                Next

                                For m = 0 To WeatherGrid.SkinEvapDepth.Count - 1 'Added by JBB 2/9/2018 copying others
                                    SkinEvapDepthDate(m) = GetDateFromPath(WeatherGrid.SkinEvapDepth(m)) 'Grabs the date from the image name, includes the time JBB
                                    DateSkinEvapDepth.Add(SkinEvapDepthDate(m).Subtract(DateSerial(YearStartWB - 1, 12, 31)).TotalDays) 'Finds the DOY for each image (this works!) JBB
                                Next

                                Dim DaysPrecipitation As New List(Of Integer)
                                Dim DaysIrrigation As New List(Of Integer)
                                Dim DaysDepletion As New List(Of Integer)
                                Dim DaysThetaV As New List(Of Integer)
                                Dim DaysEvaporatedDepth As New List(Of Integer)
                                Dim DaysSkinEvapDepth As New List(Of Integer)
                                For Each Day As Integer In DatePrecipitation : DaysPrecipitation.Add(Day) : Next 'Adds image DOY's to the list JBB
                                For Each Day As Integer In DateIrrigation : DaysIrrigation.Add(Day) : Next 'Adds image DOY's to the list JBB
                                For Each Day As Integer In DateDepletion : DaysDepletion.Add(Day) : Next 'Adds image doy's to the list
                                For Each Day As Integer In DateThetaV : DaysThetaV.Add(Day) : Next 'Adds image doy's to the list
                                For Each Day As Integer In DateEvaporatedDepth : DaysEvaporatedDepth.Add(Day) : Next 'Adds image doy's to the list added 3/1/2018 coppied from above by JBB
                                For Each Day As Integer In DateSkinEvapDepth : DaysSkinEvapDepth.Add(Day) : Next 'Adds image doy's to the list added 2/9/2018 coppied from above by JBB

                                DaysPrecipitation.Sort() 'Sorts the list of dates in ascending order JBB
                                DaysIrrigation.Sort() 'Sorts the list of dates in ascending order JBB
                                DaysDepletion.Sort() 'Sorts the list of dates in ascending order JBB
                                DaysThetaV.Sort() 'Sorts the list of dates in ascending order JBB
                                DaysEvaporatedDepth.Sort() 'Sorts the list of dates in ascending order JBB
                                DaysSkinEvapDepth.Sort() 'Sorts the list of dates in ascending order JBB 'copied from above code

                                Dim Irrigations(DaysIrrigation.Count - 1) As Single
                                Dim Precipitations(DaysPrecipitation.Count - 1) As Single
                                Dim Depletions(DaysDepletion.Count - 1) As Single
                                Dim ThetaVs(DaysThetaV.Count - 1) As Single
                                Dim EvaporatedDepths(DaysEvaporatedDepth.Count - 1) As Single
                                Dim SkinEvapDepths(DaysSkinEvapDepth.Count - 1) As Single 'added by JBB copied from above code

                                For m = 0 To PrecipitationDate.Count - 1 'grabs precip for each image JBB
                                    Precipitations(DaysPrecipitation.IndexOf(DatePrecipitation(m))) = WeatherPrecipitationDailyPixels(m).GetValue(Col, Row)
                                Next
                                For m = 0 To IrrigationDate.Count - 1 'grabs irrigation for each image JBB
                                    Irrigations(DaysIrrigation.IndexOf(DateIrrigation(m))) = WeatherIrrigationDailyPixels(m).GetValue(Col, Row)
                                Next
                                For m = 0 To DepletionDate.Count - 1 'grabs depletion for each image JBB
                                    Depletions(DaysDepletion.IndexOf(DateDepletion(m))) = WeatherDepletionDailyPixels(m).GetValue(Col, Row)
                                Next
                                For m = 0 To ThetaVDate.Count - 1 'grabs thetav for each image JBB
                                    ThetaVs(DaysThetaV.IndexOf(DateThetaV(m))) = WeatherThetaVPixels(m).GetValue(Col, Row)
                                Next
                                For m = 0 To EvaporatedDepthDate.Count - 1 'grabs evaporateddepth for each image JBB
                                    EvaporatedDepths(DaysEvaporatedDepth.IndexOf(DateEvaporatedDepth(m))) = WeatherEvaporatedDepthPixels(m).GetValue(Col, Row)
                                Next

                                For m = 0 To SkinEvapDepthDate.Count - 1 'grabs skin evaporated depth for each image JBB copied from above code
                                    SkinEvapDepths(DaysSkinEvapDepth.IndexOf(DateSkinEvapDepth(m))) = WeatherSkinEvapDepthPixels(m).GetValue(Col, Row)
                                Next
                                '******************End added by JBB*********************************

                                'WeatherPoint.Populate(WeatherTable, WeatherPointIndex, MultispectralDate(0).Year) 'Makes a table for output water balance data for a selected point, populate a SETMI function JBB
                                WeatherPoint.Populate(WeatherTable, WeatherPointIndex, YearStartWB) 'Makes a table for output water balance data for a selected point, populate a SETMI function JBB
                                Dim CoverName As String = [Enum].GetName(GetType(Cover), Cover).Replace("_", " ") 'Grabs cover name and formats it for output JBB

                                Dim OutputImages(MultispectralImages.Count - 1) As WaterBalance_ImageOverpassOutput 'Makes an array for output data for each multispectral image, top list on the output .csv file JBB
                                'Dim OutputSeason(DateSerial(MultispectralDate(0).Year, 12, 31).DayOfYear - 1) As WaterBalance_SeasonOutput 'Makes an array for output data for each day of the year, lower list on the output .csv file JBB commented out by JBB
                                Dim OutputSeason(DateSerial(YearStartWB, 12, 31).DayOfYear + DateSerial(YearStartWB + 1, 12, 31).DayOfYear - 1) As WaterBalance_SeasonOutput 'Makes an array for output data for each day of the year, lower list on the output .csv file JBB commented out by JBB 'Runs for two calendar years
                                'Dim OutputSeason(DateSerial(MultispectralDate(0).Year, 12, 31).DayOfYear) As WaterBalance_SeasonOutput 'Makes an array for output data for each day of the year, lower list on the output .csv file JBB added by JBB
                                Dim OutputDepletions(DepletionCount - 1) As WaterBalance_InputDepletionOutput 'makes and array to output data on delpetion map dates JBB
                                Dim OutputForecasts(1) As WaterBalance_FalseKcbOutput 'Makes array to output false SAVI Peak and End DOY, etc. copied from other tabular input code, JBB 2/8/2018
                                Dim OutputThetaVs(ThetaVCount - 1) As WaterBalance_InputThetaVOutput 'makes and array to output data on lower layer theta v map dates JBB
                                Dim OutputEvaporatedDepths(EvaporatedDepthCount - 1) As WaterBalance_InputEvaporatedDepthOutput 'makes and array to output data on evaporated depth map dates JBB
                                Dim OutputSkinEvapDepths(SkinEvapDepthCount - 1) As WaterBalance_InputSkinEvapDepthOutput 'Added by JBB 2/9/2018 copied from code above
                                Dim OutputLayerInfo(LyrCnt - 1) As WaterBalance_LayerOutput 'added by JBB 3/21/2018 copied from other code
                                'Dim BaseDate As DateTime = DateSerial(MultispectralDate(0).Year, 1, 1) 'Original code commented out by JBB The base date is January 1 of the study year JBB
                                'Dim BaseDate As DateTime = DateSerial(MultispectralDate(0).Year - 1, 12, 31) 'New Code by JBB, The base date is Dec 31 of the year prior to the study year JBB
                                Dim BaseDate As DateTime = DateSerial(YearStartWB - 1, 12, 31) 'New Code by JBB, The base date is Dec 31 of the year prior to the study year JBB
                                'For DoY = 0 To DateIni - 1 'Loops from beginning of year (actually 12/31 of prev year) to beginning of season JBB commented out by JBB
                                For DoY = 1 To DOYStartWB - 1 'added by JBB
                                    Dim RecordDate = BaseDate.AddDays(DoY) 'Finds the date from the DOY by adding it to 12/31 of the previous year JBB
                                    OutputSeason(DoY - 1) = New WaterBalance_SeasonOutput(RecordDate, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0) 'Populates the row with zeros if the date is before the season start JBB
                                Next
                                For DoY = DOYEndWB To OutputSeason.Length '- 1 'loops from end of season to end of the year JBB
                                    Dim RecordDate = BaseDate.AddDays(DoY) 'Finds the date from the DOY by adding it to 12/31 of the previous year JBB
                                    OutputSeason(DoY - 1) = New WaterBalance_SeasonOutput(RecordDate, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0) 'Populates the row with zeros if the date is after season end JBB
                                Next
                                OutputForecasts(0) = New WaterBalance_FalseKcbOutput(BaseDate.AddDays(1), -999, -999, -999) 'Initial Values copying code from around here 2/8/2018 JBB
                                OutputForecasts(1) = New WaterBalance_FalseKcbOutput(BaseDate.AddDays(1), -999, -999, -999) 'Initial Values copying code from around here 2/8/2018 JBB

                                '*****loopthough and minimize sum of squares for DOYini and DOYefc added by JBB
                                Dim GDD(DOYEndWB) As Single 'added by JBB for SAVI Log Interp
                                Dim SAVIGDDs(MultispectralImages.Count - 1) As Single 'Makes an array for GDD, one item for each image for SAVI Log Interp JBB
                                Dim MaxSAVI As Single = 0 'added by JBB for SAVI log Interp
                                Dim MaxSAVIDOY As Integer = 0 'added by JBB for SAVI log Interp
                                Dim MaxSAVIGDD As Single = 0 'added by JBB for SAVI log regression 2/8/2018
                                Dim MaxSaviIndx As Integer = 0 'added by JBB for SAVI log Interp
                                Dim FlgMaxSAVI As Boolean = False 'added by JBB for SAVI log regression 2/8/2018
                                Dim FlgMaxSAVIZero As Boolean = False 'added by JBB for SAVI log regression 2/8/2018
                                Dim FlgEndSAVIZero As Boolean = False 'added by JBB for SAVI log Regression 2/8/2018
                                Dim OldMaxSAVIIndx As Integer = 0 'added by JBB if all peak images are late in the season
                                Dim DailyKcb(DOYEndWB) As Single 'added by JBB for SAVI log Interp
                                Dim DailyHcETo(DOYEndWB) As Single
                                Dim DailySAVI(DOYEndWB) As Single 'added by JBB for SAVI log Interp
                                Dim MaxKcb As Single = 0
                                Dim EFCFlg As Boolean = False
                                Dim INIFlg As Boolean = False
                                Dim FlgOutLyr As Boolean = False 'added by JBB to only output layer info once
                                Dim DOYMaxBM As Integer = -1

                                '***********************************************************************************
                                '*****     Start Kcb Section
                                If KcbMethod = KcbType.Fitted_Curve Then
                                    If EarlyImageCount = 0 Then
                                        DOYini = Math.Round((DOYiniMin + DOYiniMax) / 2) 'if no image is between the input inititation and efc take average max and min doy's
                                        DOYefc = Math.Round((DOYefcMin + DOYefcMax) / 2) 'if no image is between the input inititation and efc take average max and min doy's
                                    ElseIf EarlyImageCount = 1 Then 'moves whichever of DOYini and DOYefc is closet to the image date 
                                        DOYini = Math.Round((DOYiniMin + DOYiniMax) / 2) 'setup initial value
                                        DOYefc = Math.Round((DOYefcMin + DOYefcMax) / 2) 'setup initial value
                                        If DOYFirst - DOYini < DOYefc - DOYFirst Then 'loop and allow doy ini to change, otherwise only change doy efc
                                            For II = DOYiniMin To DOYiniMax
                                                'SSq = 0
                                                'For KK = 0 To MultispectralDate.Count - 1
                                                'If DateReflectance(KK) <= DOYefcMax Then
                                                'KcbInt = calcKcbPolynomial(Cover, DateReflectance(KK), II, DOYefc, 365, ETReference, KcbOffSeason)
                                                KcbInt = calcKcbPolynomial(Cover, DateReflectance(MFirst), II, DOYefc, 366, ETReference, KcbOffSeason)
                                                'SSq = SSq + (KcbReflectance(KK) - KcbInt) ^ 2
                                                SSq = SSq + (KcbReflectance(MFirst) - KcbInt) ^ 2
                                                'End If
                                                'Next KK
                                                If SSq < SSqMin Then 'a new best fit has been found
                                                    SSqMin = SSq
                                                    DOYini = II
                                                End If
                                            Next II
                                        Else 'loop and allow doy efc to change
                                            For JJ = DOYefcMin To DOYefcMax
                                                SSq = 0
                                                'For KK = 0 To MultispectralDate.Count - 1
                                                'If DateReflectance(KK) <= DOYefcMax Then
                                                'KcbInt = calcKcbPolynomial(Cover, DateReflectance(KK), DOYini, JJ, 365, ETReference, KcbOffSeason)
                                                KcbInt = calcKcbPolynomial(Cover, DateReflectance(MFirst), DOYini, JJ, 366, ETReference, KcbOffSeason)
                                                'SSq = SSq + (KcbReflectance(KK) - KcbInt) ^ 2
                                                SSq = SSq + (KcbReflectance(MFirst) - KcbInt) ^ 2
                                                'End If
                                                'Next KK
                                                If SSq < SSqMin Then 'a new best fit has been found
                                                    SSqMin = SSq
                                                    DOYefc = JJ
                                                End If
                                            Next JJ
                                        End If
                                    Else 'if two + images allow DOYini and DOYefc to change
                                        For II = DOYiniMin To DOYiniMax
                                            For JJ = DOYefcMin To DOYefcMax
                                                SSq = 0
                                                For KK = 0 To MultispectralDate.Count - 1
                                                    If DateReflectance(KK) <= DOYefcMax And DateReflectance(KK) >= DOYiniMin Then
                                                        KcbInt = calcKcbPolynomial(Cover, DateReflectance(KK), II, JJ, 366, ETReference, KcbOffSeason)
                                                        SSq = SSq + (KcbReflectance(KK) - KcbInt) ^ 2
                                                    End If
                                                Next KK
                                                If SSq < SSqMin Then 'a new best fit has been found
                                                    SSqMin = SSq
                                                    DOYini = II
                                                    DOYefc = JJ
                                                End If
                                            Next JJ
                                        Next II
                                    End If

                                    'find if there is an image between EFC and termination
                                    For m = 0 To MultispectralDate.Count - 1 'grabs reflectance based Kcb for each image JBB
                                        If DateReflectance(m) > DOYefc And DateReflectance(m) <= DOYtermMax Then LateImageFlag = True
                                    Next

                                    SSqMin = 99999
                                    If LateImageFlag = False Then
                                        DOYterm = Math.Round((DOYtermMin + DOYtermMax) / 2) 'if no image is between the input efc and termination take average max and min doy's
                                    Else 'loop and allow doy term to change
                                        For JJ = DOYtermMin To DOYtermMax
                                            SSq = 0
                                            For KK = 0 To MultispectralDate.Count - 1
                                                If DateReflectance(KK) > DOYefc And DateReflectance(KK) <= DOYtermMax Then
                                                    KcbInt = calcKcbPolynomial(Cover, DateReflectance(KK), DOYini, DOYefc, JJ, ETReference, KcbOffSeason)
                                                    SSq = SSq + (KcbReflectance(KK) - KcbInt) ^ 2
                                                End If
                                            Next KK
                                            If SSq < SSqMin Then 'a new best fit has been found
                                                SSqMin = SSq
                                                DOYterm = JJ
                                            End If
                                        Next JJ
                                    End If

                                    'Find max Kcb
                                    Dim KcbInt2 As Single = 0
                                    For DoY = DOYStartWB To DOYEndWB
                                        KcbInt2 = calcKcbPolynomial(Cover, DoY, DOYini, DOYefc, DOYterm, ETReference, KcbOffSeason)
                                        DailyKcb(DoY) = KcbInt2
                                        If KcbInt2 > MaxKcb Then MaxKcb = KcbInt2
                                    Next

                                ElseIf KcbMethod = KcbType.VI_Log_Regression Then 'Do the log Interp method like Isidro
                                    '******************This is programmed in terms of SAVI, however, if NDVI is used for the kcbrf
                                    '*********either input, or hardcoded, then SAVI is overwritten to be NDVI and the code below
                                    '*********while referring to SAVI is really using NDVI
                                    '*******************************************************************************************
                                    DOYterm = DOYEndWB 'sets initial value of end of calculations
                                    Dim FalseSAVICnt As Integer = 0
                                    If FalseEndSAVI > 0 Then FalseSAVICnt = FalseSAVICnt + 1 'count false savi times
                                    If FalsePeakSAVI > 0 Then FalseSAVICnt = FalseSAVICnt + 1 'count false savi times
                                    For DoY = DOYStartWB To DOYEndWB 'added by JBB
                                        'Dim IWeather2 As Integer = WeatherPoint.RecordDate.IndexOf(DateSerial(MultispectralDate(0).Year - 1, 12, 31).AddDays(DoY)) 'IWeather will be the index of the current doy in the weatherpoint
                                        Dim IWeather2 As Integer = WeatherPoint.RecordDate.IndexOf(DateSerial(YearStartWB - 1, 12, 31).AddDays(DoY)) 'IWeather will be the index of the current doy in the weatherpoint
                                        If IWeather2 > -1 Then
                                            'If DoY >= DOYiniMin Then 'Start calculating GDD '<-- Not needed
                                            Dim Tmax As Single = WeatherPoint.DailyMaximumTemperature(IWeather2) 'Grabs Daily Tmax from the weather point table JBB
                                            Dim Tmin As Single = WeatherPoint.DailyMinimumTemperature(IWeather2) 'Grabs Daily Tmin from the weather point table JBB
                                            Dim Tlo As Single ' = 10 'GDD Lower Threshold C
                                            Dim Thi As Single ' = 30 'GDD Upper Threshold C
                                            If CoverPointIndex.GDDBase < 0 Then 'Added by JBB followign code used in other places.
                                                Tlo = 10 'Default
                                            Else
                                                Tlo = GDDBase
                                            End If
                                            If CoverPointIndex.GDDMaxTemp < 0 Then 'Added by JBB followign code used in other places.
                                                Thi = 30 'Default
                                            Else
                                                Thi = GDDMaxTemp
                                            End If
                                            Tmax = Math.Min(Math.Max(Tmax, Tlo), Thi) 'Limits Tmax to not drop below Tlo or go above Thi
                                            ' Tmin = Math.Max(Tmin, Tlo) 'see https://ndawn.ndsu.nodak.edu/help-corn-growing-degree-days.html
                                            Tmin = Math.Min(Math.Max(Tmin, Tlo), Thi) 'Limits Tmin to not drop below Tlo or go above Thi <-- most people seem to do it this way, e.g. Sharma and Irmak, et al. Cover Crops 2017, J. Irr. Drain. Engr. Paper I.
                                            Dim GDDi As Single = (Tmax + Tmin) / 2 - Tlo 'GDD Limited to be no less than zero
                                            If DoY = DOYStartWB Then 'first day
                                                GDD(DoY) = GDDi
                                            Else 'cumulative if not first day
                                                GDD(DoY) = GDD(DoY - 1) + GDDi
                                            End If

                                            For m = 0 To MultispectralImages.Count - 1
                                                If DateReflectance(m) = DoY Then
                                                    SAVIGDDs(m) = GDD(DoY)
                                                End If
                                            Next
                                            'End If
                                        End If
                                    Next

                                    For m = 0 To MultispectralImages.Count - 1
                                        If SAVIs(m) > MaxSAVI Then 'Finds the Maximum SAVI and its image index
                                            MaxSAVI = SAVIs(m)
                                            MaxSaviIndx = m
                                            MaxSAVIDOY = DateReflectance(m)
                                            MaxSAVIGDD = SAVIGDDs(m) 'added 2/8/2018
                                        End If
                                    Next
                                    OldMaxSAVIIndx = MaxSaviIndx 'added by JBB 01/20/2017 to push peak images to end of season
                                    If FalseEndSAVI > 0 Or FalsePeakSAVI > 0 And FlgFalseSAVI = False Then
                                        'MsgBox("Tabular False Peak and/or End SAVI value(s) > 0 and will be used in calculations")'msgbox moved to be less obtrusive JBB
                                        FlgFalseSAVI = True 'Keeps this message from reappearing many times
                                    End If

                                    If MaxSaviIndx > MultispectralImages.Count - 3 And FlgImageCntLate = False Then
                                        'MsgBox("Too Few Late Images for SAVI Log Interpolation, Program will only use false peak and end SAVI if > 0")'msgbox moved to be less obtrusive JBB
                                        FlgImageCntLate = True 'Keeps this message from reappearing many times
                                    End If

                                    'Allows for either doy or gdd based false peak savi cuttoff
                                    If KcbVIForecastMeth = SAVIForecast.Day_of_Year And MaxSAVIDOY > FalsePeakSAVIDOY Then
                                        FlgMaxSAVI = True
                                    End If
                                    If KcbVIForecastMeth = SAVIForecast.Growing_Degree_Days And MaxSAVIGDD > FalsePeakSAVIGDD Then
                                        FlgMaxSAVI = True
                                    End If
                                    'Allows for either doy or gdd based false peak savi cuttoff
                                    If KcbVIForecastMeth = SAVIForecast.Day_of_Year And FalsePeakSAVIDOY > 0 Then
                                        FlgMaxSAVIZero = True
                                    End If
                                    If KcbVIForecastMeth = SAVIForecast.Growing_Degree_Days And FalsePeakSAVIGDD > 0 Then
                                        FlgMaxSAVIZero = True
                                    End If
                                    'Allows for either doy or gdd based false peak savi cuttoff
                                    If KcbVIForecastMeth = SAVIForecast.Day_of_Year And FalseEndSAVIDOY > 0 Then
                                        FlgEndSAVIZero = True
                                    End If
                                    If KcbVIForecastMeth = SAVIForecast.Growing_Degree_Days And FalseEndSAVIGDD > 0 Then
                                        FlgEndSAVIZero = True
                                    End If

                                    If MaxSaviIndx < 1 Then
                                        If FlgImageCntEarly = False Then
                                            'MsgBox("Too Few Early Images for SAVI Log Interpolation, Program will not run properly")'msgbox moved to be less obtrusive JBB
                                            FlgImageCntEarly = True 'Keeps this message from reappearing many times
                                        End If
                                        'ElseIf MaxSaviIndx = 1 And FalsePeakSAVI > 0 Then
                                    ElseIf MaxSaviIndx = 1 And FlgMaxSAVI = True And FalsePeakSAVI > 0 Then
                                        'ElseIf MaxSaviIndx = 1 And MaxSAVIDOY > FalsePeakSAVIDOY And FalsePeakSAVI > 0 Then
                                        If FlgImageCntEarly2 = False Then
                                            'MsgBox("Too Few Early Images for SAVI Log Interpolation with Peak SAVI, Program will not run properly")'msgbox moved to be less obtrusive JBB
                                            FlgImageCntEarly2 = True 'Keeps this message from reappearing many times
                                        End If
                                        'ElseIf MaxSaviIndx > MultispectralImages.Count - 3 And FlgImageCntLate = False Then
                                        '    MsgBox("Too Few Late Images for SAVI Log Interpolation, Program will Crash")
                                        '    FlgImageCntLate = True 'Keeps this message from reappearing many times
                                    Else
                                        Dim SlopeEarly As Single
                                        Dim IncptEarly As Single
                                        Dim RsqEarly As Single
                                        Dim DOYFalsePeak As Integer = -999
                                        Dim GDDFalsePeak As Single = 0
                                        Dim MaxSAVIOffset As Integer = 1 'by default max savi is peak season
                                        Dim SAVIPeakOut As Single
                                        Dim DOYPeakOut As Single
                                        If FalsePeakSAVI > 0 Then 'Finds false peak savi
                                            If FlgMaxSAVI = True Then 'Allows for DOY and GDD'based limiting
                                                'If MaxSAVIDOY > FalsePeakSAVIDOY Then
                                                Dim LnEarly(MaxSaviIndx - 1) As Single
                                                Dim GDDEarly(MaxSaviIndx - 1) As Single
                                                For m = 0 To MaxSaviIndx - 1
                                                    LnEarly(m) = Math.Log(Math.Max(SAVIs(m), 0.05)) 'The 0.05 prevents a ln of zero
                                                    GDDEarly(m) = SAVIGDDs(m)
                                                Next
                                                SlopeEarly = Slp(GDDEarly, LnEarly)
                                                IncptEarly = Incpt(GDDEarly, LnEarly)
                                                RsqEarly = Rsqr(GDDEarly, LnEarly)
                                            Else
                                                Dim LnEarly(MaxSaviIndx) As Single
                                                Dim GDDEarly(MaxSaviIndx) As Single
                                                For m = 0 To MaxSaviIndx
                                                    LnEarly(m) = Math.Log(Math.Max(SAVIs(m), 0.05)) 'The 0.05 prevents a ln of zero
                                                    GDDEarly(m) = SAVIGDDs(m)
                                                Next
                                                SlopeEarly = Slp(GDDEarly, LnEarly)
                                                IncptEarly = Incpt(GDDEarly, LnEarly)
                                                RsqEarly = Rsqr(GDDEarly, LnEarly)
                                            End If
                                            For DoY = DOYStartWB To DOYEndWB 'find the projected EFC doy
                                                DailySAVI(DoY) = Math.Exp(SlopeEarly * GDD(DoY) + IncptEarly)
                                                If DailySAVI(DoY) > FalsePeakSAVI Then
                                                    DOYFalsePeak = DoY 'This day will be the false peak, but will actually be in the late season since it exceeded peak SAVI
                                                    GDDFalsePeak = GDD(DoY) 'added 2/8/2018 JBB
                                                    Exit For
                                                End If
                                            Next
                                        Else
                                            If KcbVIForecastMeth = SAVIForecast.Day_of_Year Then 'This logic added 2/8/2018 to allow for GDD to be used in forecasting
                                                '***Added on 01/20/2017 JBB to cutoff peak images late in the season from affecting early season curve
                                                If FalsePeakSAVIDOY > 0 And MaxSaviIndx > 1 Then
                                                    For m = OldMaxSAVIIndx To 2 Step -1
                                                        If DateReflectance(m) > FalsePeakSAVIDOY Then
                                                            MaxSaviIndx = m - 1
                                                            MaxSAVIsEarly = True 'will prompt a msg box later
                                                        End If
                                                    Next
                                                End If
                                            ElseIf KcbVIForecastMeth = SAVIForecast.Growing_Degree_Days Then
                                                If FalsePeakSAVIGDD > 0 And MaxSaviIndx > 1 Then
                                                    For m = OldMaxSAVIIndx To 2 Step -1
                                                        If SAVIGDDs(m) > FalsePeakSAVIGDD Then
                                                            MaxSaviIndx = m - 1
                                                            MaxSAVIsEarly = True 'will prompt a msg box later
                                                        End If
                                                    Next
                                                End If
                                            End If
                                            '***End Added
                                            Dim LnEarly(MaxSaviIndx) As Single
                                            Dim GDDEarly(MaxSaviIndx) As Single

                                            For m = 0 To MaxSaviIndx
                                                LnEarly(m) = Math.Log(Math.Max(SAVIs(m), 0.05)) 'The 0.05 prevents a ln of zero
                                                GDDEarly(m) = SAVIGDDs(m)
                                            Next
                                            SlopeEarly = Slp(GDDEarly, LnEarly)
                                            IncptEarly = Incpt(GDDEarly, LnEarly)
                                            RsqEarly = Rsqr(GDDEarly, LnEarly)
                                        End If
                                        'This will include the max image in both early and late seasons in at least some conditions
                                        If MaxSAVIDOY > DOYFalsePeak And DOYFalsePeak > 0 Then
                                            FalseSAVICnt = FalseSAVICnt + 1
                                            MaxSAVIOffset = 0 'allows max savi to be late season
                                        ElseIf MaxSaviIndx = MultispectralImages.Count - 1 And FlgEndSAVIZero = True And FalsePeakSAVI <= 0 Then 'Will create a fake peak SAVI the day after peak for the declining portion of the curve, if no images are input after the max savi and if no false peak savi is input and if a false end savi is input
                                            'ElseIf MaxSaviIndx = MultispectralImages.Count - 1 And FalseEndSAVI > 0 And FalsePeakSAVI <= 0 Then 'Will create a fake peak SAVI the day after peak for the declining portion of the curve, if no images are input after the max savi and if no false peak savi is input and if a false end savi is input
                                            FalseSAVICnt = FalseSAVICnt + 1
                                        End If

                                        Dim SlopeLate As Single = 0 'If no images past the max then we don't calculate the drop, will cap at Kcbmax or something
                                        Dim IncptLate As Single = 0
                                        Dim RsqLate As Single = 0
                                        Dim SAVILimit As Single = MaxSAVI '0.68 

                                        If FalsePeakSAVI > MaxSAVI Then
                                            SAVILimit = FalsePeakSAVI
                                        ElseIf KcbVIForecastMeth = SAVIForecast.Day_of_Year And MaxSAVIDOY > DOYFalsePeak And DOYFalsePeak > 0 Then 'DOYFalsePeak will be engaged whether DoY or GDD are used
                                            SAVILimit = MaxSAVI * 1.001 'prevents log of a zero or negative number
                                            '****Added on 01/20/2017 JBB to cutoff peak images late in the season from affecting early season curve
                                        ElseIf KcbVIForecastMeth = SAVIForecast.Growing_Degree_Days And MaxSAVIGDD > GDDFalsePeak And GDDFalsePeak > 0 Then 'gddFalsePeak will be engaged whether DoY or GDD are used
                                            SAVILimit = MaxSAVI * 1.001 'prevents log of a zero or negative number'!!!!could DOYFlasePeak capture same thing that FlgMaxSAVIZero does?
                                        ElseIf FlgMaxSAVIZero = True And OldMaxSAVIIndx <> MaxSaviIndx Then
                                            'ElseIf FalsePeakSAVIDOY > 0 And OldMaxSAVIIndx <> MaxSaviIndx Then
                                            SAVILimit = MaxSAVI * 1.001 'prevents log of a zero or negative number
                                            '***end added
                                        End If

                                        Dim LnLate(MultispectralImages.Count - MaxSaviIndx - 2 + FalseSAVICnt) As Single
                                        Dim GDDLate(MultispectralImages.Count - MaxSaviIndx - 2 + FalseSAVICnt) As Single
                                        For m = MaxSaviIndx + MaxSAVIOffset To MultispectralImages.Count - 1
                                            LnLate(m - MaxSaviIndx - MaxSAVIOffset) = Math.Log(-Math.Log(SAVIs(m) / SAVILimit))
                                            GDDLate(m - MaxSaviIndx - MaxSAVIOffset) = SAVIGDDs(m)
                                        Next
                                        If FalsePeakSAVI > 0 Then
                                            If DOYFalsePeak < -99 Then 'Counts times that a false peak doy was forced
                                                DOYFalsePeak = FalsePeakSAVIDOY
                                                GDDFalsePeak = FalsePeakSAVIGDD  'added 2/8/2018
                                                CntForceFakeSAVI = CntForceFakeSAVI + 1
                                            End If
                                            LnLate(MultispectralImages.Count - MaxSaviIndx - MaxSAVIOffset) = Math.Log(-Math.Log(FalsePeakSAVI * 0.999 / SAVILimit)) 'The 0.999 PREVENTS A LOG OF A NEGATIVE  OR ZERO NUMBER
                                            If KcbVIForecastMeth = SAVIForecast.Day_of_Year Then 'added 2/8/2018
                                                GDDLate(MultispectralImages.Count - MaxSaviIndx - MaxSAVIOffset) = GDD(DOYFalsePeak)
                                            ElseIf KcbVIForecastMeth = SAVIForecast.Growing_Degree_Days Then 'added 2/8/2018
                                                GDDLate(MultispectralImages.Count - MaxSaviIndx - MaxSAVIOffset) = GDDFalsePeak 'added 2/8/2018
                                            End If 'added 2/8/2018
                                            DOYPeakOut = DOYFalsePeak 'added 2/8/2018 for output
                                            SAVIPeakOut = FalsePeakSAVI * 0.999 'added 2/8/2018 for output
                                            If FalseEndSAVI > 0 Then
                                                LnLate(MultispectralImages.Count - MaxSaviIndx - MaxSAVIOffset + 1) = Math.Log(-Math.Log(FalseEndSAVI / SAVILimit))
                                                If KcbVIForecastMeth = SAVIForecast.Day_of_Year Then 'added 2/8/2018
                                                    GDDLate(MultispectralImages.Count - MaxSaviIndx - MaxSAVIOffset + 1) = GDD(FalseEndSAVIDOY)
                                                ElseIf KcbVIForecastMeth = SAVIForecast.Growing_Degree_Days Then 'added 2/8/2018
                                                    GDDLate(MultispectralImages.Count - MaxSaviIndx - MaxSAVIOffset + 1) = FalseEndSAVIGDD 'added 2/8/2018
                                                End If 'added 2/8/2018
                                            End If
                                        ElseIf MaxSaviIndx = MultispectralImages.Count - 1 And FalseEndSAVI > 0 Then 'applies a fake savi just a hair smaller than the max the day after the max for the end of season
                                            LnLate(MultispectralImages.Count - MaxSaviIndx - MaxSAVIOffset) = Math.Log(-Math.Log(MaxSAVI * 0.999 / SAVILimit)) 'The 0.999 PREVENTS A LOG OF A NEGATIVE  OR ZERO NUMBER
                                            GDDLate(MultispectralImages.Count - MaxSaviIndx - MaxSAVIOffset) = GDD(MaxSAVIDOY + 1) 'applies it the day after peak
                                            'If FalseEndSAVI > 0 Then'This was redundant
                                            LnLate(MultispectralImages.Count - MaxSaviIndx - MaxSAVIOffset + 1) = Math.Log(-Math.Log(FalseEndSAVI / SAVILimit))
                                            If KcbVIForecastMeth = SAVIForecast.Day_of_Year Then 'added 2/8/2018
                                                GDDLate(MultispectralImages.Count - MaxSaviIndx - MaxSAVIOffset + 1) = GDD(FalseEndSAVIDOY)
                                            ElseIf KcbVIForecastMeth = SAVIForecast.Growing_Degree_Days Then 'added 2/8/2018
                                                GDDLate(MultispectralImages.Count - MaxSaviIndx - MaxSAVIOffset + 1) = FalseEndSAVIGDD 'added 2/8/2018
                                            End If 'added 2/8/2018
                                            'End If
                                            DOYPeakOut = MaxSAVIDOY + 1 'added 2/8/2018 for output
                                            SAVIPeakOut = MaxSAVI * 0.999 'added 2/8/2018 for output
                                        ElseIf FalseEndSAVI > 0 Then
                                            LnLate(MultispectralImages.Count - MaxSaviIndx - MaxSAVIOffset) = Math.Log(-Math.Log(FalseEndSAVI / SAVILimit)) 'The 0.05 PREVENTS A LOG OF A NEGATIVE  OR ZERO NUMBER
                                            If KcbVIForecastMeth = SAVIForecast.Day_of_Year Then 'added 2/8/2018
                                                GDDLate(MultispectralImages.Count - MaxSaviIndx - MaxSAVIOffset) = GDD(FalseEndSAVIDOY)
                                            ElseIf KcbVIForecastMeth = SAVIForecast.Growing_Degree_Days Then 'added 2/8/2018
                                                GDDLate(MultispectralImages.Count - MaxSaviIndx - MaxSAVIOffset) = FalseEndSAVIGDD 'added 2/8/2018
                                            End If 'added 2/8/2018
                                        End If

                                        SlopeLate = Slp(GDDLate, LnLate)
                                        IncptLate = Incpt(GDDLate, LnLate)
                                        RsqLate = Rsqr(GDDLate, LnLate)

                                        For DoY = DOYStartWB To DOYEndWB 'added by JBB
                                            DailySAVI(DoY) = Math.Min(Math.Exp(SlopeEarly * GDD(DoY) + IncptEarly), SAVILimit * Math.Exp(-Math.Exp(SlopeLate * GDD(DoY) + IncptLate)))
                                            DailyKcb(DoY) = Math.Max(calcKcbReflectance(Cover, DailySAVI(DoY), ETReference, DailySAVI(DoY), KcbVI, KcbrfSlope, KcbrfIntercept, MaxKcbrfInput, MinKcbrfInput), KcbOffSeason) 'SAVI is input for NDVI, because the interpolation code was written in terms of SAVI, if NDVI is used for the Kcbrf, SAVI is overwritten as NDVI to avoid recoding the entire Log Regression Section for generic VI.
                                            If DailyKcb(DoY) > MaxKcb Then MaxKcb = DailyKcb(DoY) 'finds max Kcb
                                            If INIFlg = False And DailyKcb(DoY) > KcbOffSeason Then
                                                DOYini = DoY
                                                INIFlg = True
                                            End If
                                        Next
                                        Dim MaxKcbEFC As Single = DailyKcb(DOYini)
                                        For doy = DOYini + 1 To DOYEndWB
                                            If doy < DOYEndWB And INIFlg = True Then 'And EFCFlg = False Then
                                                If DailyKcb(doy) > MaxKcbEFC Then 'DailyKcb(doy - 1) Then 'And DailyKcb(doy) > DailyKcb(doy + 1) Then
                                                    DOYefc = doy
                                                    MaxKcbEFC = DailyKcb(doy) 'efc happens at peak, but first of peak if plateau
                                                    EFCFlg = True
                                                End If
                                            End If
                                        Next
                                        For doy = DOYini + 1 To DOYEndWB
                                            If EFCFlg = True Then
                                                If DailyKcb(doy) <= KcbOffSeason Then
                                                    DOYterm = doy - 1
                                                    Exit For
                                                End If
                                            End If
                                        Next
                                        For doy = DOYefc To DOYterm
                                            If DailySAVI(doy) < BMMaxSAVI Then
                                                DOYMaxBM = doy - 1
                                                Exit For
                                            End If
                                        Next
                                        'added for output
                                        Dim FalseDate1 As DateTime
                                        Dim FalseDate2 As DateTime
                                        Dim FalseGDD1 As Single
                                        Dim FalseGDD2 As Single
                                        Dim FalseKcbrf As Single
                                        If KcbVIForecastMeth = SAVIForecast.Day_of_Year Then
                                            FalseDate1 = BaseDate.AddDays(DOYPeakOut)
                                            FalseDate2 = BaseDate.AddDays(FalseEndSAVIDOY)
                                            FalseGDD1 = GDD(DOYPeakOut)
                                            If FalseEndSAVIDOY > 0 Then
                                                FalseGDD2 = GDD(FalseEndSAVIDOY)
                                            Else
                                                FalseGDD2 = 0
                                            End If
                                        ElseIf KcbVIForecastMeth = SAVIForecast.Growing_Degree_Days Then
                                                For doy = DOYStartWB To DOYEndWB
                                                If GDD(doy) > FalseEndSAVIGDD Then
                                                    FalseDate2 = BaseDate.AddDays(doy)
                                                    Exit For
                                                End If
                                            Next
                                            FalseDate1 = BaseDate.AddDays(DOYFalsePeak)
                                            FalseGDD1 = GDDFalsePeak
                                            FalseGDD2 = FalseEndSAVIGDD
                                        End If
                                        FalseKcbrf = calcKcbReflectance(Cover, SAVIPeakOut, ETReference, SAVIPeakOut, KcbVI, KcbrfSlope, KcbrfIntercept, MaxKcbrfInput, MinKcbrfInput)
                                        OutputForecasts(0) = New WaterBalance_FalseKcbOutput(FalseDate1, FalseGDD1, FalseKcbrf, SAVILimit)
                                        FalseKcbrf = calcKcbReflectance(Cover, FalseEndSAVI, ETReference, FalseEndSAVI, KcbVI, KcbrfSlope, KcbrfIntercept, MaxKcbrfInput, MinKcbrfInput)
                                        OutputForecasts(1) = New WaterBalance_FalseKcbOutput(FalseDate2, FalseGDD2, FalseKcbrf, SAVILimit)
                                    End If
                                ElseIf KcbMethod = KcbType.VI_Interpolation Then 'Finds peak Kcb
                                    Dim KcbInt3 As Single = 0
                                    For DoY = DOYStartWB To DOYEndWB 'added by JBB
                                        DailySAVI(DoY) = calcKcbInterpolation(DoY, Days.ToArray, SAVIs, KcbOffSeason, Cover, ETReference, KcbVI, KcbrfSlope, KcbrfIntercept, MaxKcbrfInput, MinKcbrfInput) 'Iterpolates the Kcbs JBB
                                        If DailySAVI(DoY) < -99 Then 'acts like it did when Kcbrf was computed in caclkcbinterpolation
                                            KcbInt3 = KcbOffSeason
                                        Else
                                            KcbInt3 = calcKcbReflectance(Cover, DailySAVI(DoY), ETReference, DailySAVI(DoY), KcbVI, KcbrfSlope, KcbrfIntercept, MaxKcbrfInput, MinKcbrfInput)
                                        End If
                                        DailyKcb(DoY) = KcbInt3
                                        If KcbInt3 > MaxKcb Then
                                            MaxKcb = KcbInt3 'finds max KcbWaterBalance_InputEvaporatedDepthOutput
                                            DOYefc = DoY 'EFC is on max kcb day
                                        End If
                                        If INIFlg = False And KcbInt3 > KcbOffSeason Then
                                            DOYini = DoY
                                            INIFlg = True
                                        End If
                                    Next
                                    For DoY = DOYefc To DOYEndWB
                                        'calcKcbInterpolation(DoY, Days.ToArray, SAVIs, KcbOffSeason, Cover, ETReference, KcbVI, KcbrfSlope, KcbrfIntercept, MaxKcbrfInput, MinKcbrfInput) <= KcbOffSeason Then
                                        If DailyKcb(DoY) <= KcbOffSeason Then
                                            DOYterm = DoY - 1
                                            Exit For
                                        End If
                                    Next
                                End If

                                If ETReference = ETReferenceType.Short_Grass = True Then 'Allows Selection of Short or Tall Reference ET, This logic was added by JBB
                                    Dim FcInt As Single = 0
                                    DailyHcETo(DOYStartWB - 1) = 0 'initial Hcprevious
                                    For DoY = DOYStartWB To DOYEndWB
                                        DailyHcETo(DoY) = Math.Max(Math.Min(MinimumCoverHeight + (MaximumCoverHeight - MinimumCoverHeight) * (DailyKcb(DoY) - KcbOffSeason) / (MaxKcb - KcbOffSeason), MaximumCoverHeight), MinimumCoverHeight) 'Crop height from tabular heights calculated as a linear relationship with Kcb JBB
                                        If DailyHcETo(DoY) < DailyHcETo(DoY - 1) And DoY <= DOYterm Then DailyHcETo(DoY) = DailyHcETo(DoY - 1) 'keeps crop from shrinking until end of season added by JBB
                                    Next
                                    For DoY = DOYini To DOYefc 'Find the development date
                                        If FcMethodWB = FcType.Vegetation_Index And KcbMethod <> KcbType.Fitted_Curve Then
                                            FcInt = calcFcWB(Cover, DailySAVI(DoY), DailySAVI(DoY), KcbVI, MaxVIInput, MinVIInput, ExpVIInput, 0.99, 0) 'We use DailySAVI for both NDVI and SAVI, because that is where the NDVI will be saved if that index is selected, sorry for the confusion, JBB
                                            If FcInt > 0.1 Then
                                                DOYdev = DoY
                                                Exit For
                                            End If
                                        Else 'Default is FAO 56
                                            Dim KcMinInt As Single = 0
                                            Dim KcMaxInt As Single = 0
                                            KcMaxInt = Math.Max(KcMaxInput, DailyKcb(DoY) + 0.05) 'Added by JBB Doesn't include Climate Adjustment as it will later on, probably will keep it this way because the climate adjustment is dependent on this computation!
                                            KcMinInt = KcbOffSeason 'Limit(KcbOffSeason, 0.15, 0.2)
                                            Dim HcInt As Single = DailyHcETo(DoY) ' Math.Max(Math.Min(MinimumCoverHeight + (MaximumCoverHeight - MinimumCoverHeight) * (DailyKcb(DoY) - KcbOffSeason) / (MaxKcb - KcbOffSeason), MaximumCoverHeight), MinimumCoverHeight) 'Crop height from tabular heights calculated as a linear relationship with Kcb JBB
                                            'If DailyKcb(DoY) >= KcMinInt Then FcInt = Math.Max(Math.Min((Math.Max(DailyKcb(DoY) - KcMinInt, 0.01) / Math.Max(KcMaxInt - KcMinInt, 0.01)) ^ (1 + 0.5 * HcInt), 0.99), 0) 'added by JBB to prevent undefined # if numerator goes negative this matches related text in FAO56 Eq 76
                                            FcInt = Math.Max(Math.Min((Math.Max(DailyKcb(DoY) - KcMinInt, 0.01) / Math.Max(KcMaxInt - KcMinInt, 0.01)) ^ (1 + 0.5 * HcInt), 0.99), 0) 'added by JBB to prevent undefined # if numerator goes negative this matches related text in FAO56 Eq 76
                                            If FcInt >= 0.1 Then
                                                DOYdev = DoY
                                                Exit For
                                            End If
                                        End If
                                    Next
                                    For DoY = DOYefc To DOYterm 'Find the development date
                                        If DailyKcb(DoY) <= RatioLate * MaxKcb Then
                                            DOYlate = DoY
                                            Exit For
                                        End If
                                    Next
                                    Dim CntAvg As Integer = 0
                                    For DoY = DOYini To DOYdev - 1
                                        HcIni = HcIni + DailyHcETo(DoY) 'Math.Max(Math.Min(MinimumCoverHeight + (MaximumCoverHeight - MinimumCoverHeight) * (DailyKcb(DoY) - KcbOffSeason) / (MaxKcb - KcbOffSeason), MaximumCoverHeight), MinimumCoverHeight) 'Crop height from tabular heights calculated as a linear relationship with Kcb JBB
                                        CntAvg = CntAvg + 1
                                    Next
                                    HcIni = Limit(HcIni / CntAvg, 0.1, 20) 'Follows upper limit as in ASCE MOP 70 ed2 , p. 275 for upper, FAO56 eq. 62 and eq. 70 for lower
                                    CntAvg = 0
                                    For DoY = DOYdev To DOYefc - 1
                                        HcDev = HcDev + DailyHcETo(DoY) 'Math.Max(Math.Min(MinimumCoverHeight + (MaximumCoverHeight - MinimumCoverHeight) * (DailyKcb(DoY) - KcbOffSeason) / (MaxKcb - KcbOffSeason), MaximumCoverHeight), MinimumCoverHeight) 'Crop height from tabular heights calculated as a linear relationship with Kcb JBB
                                        CntAvg = CntAvg + 1
                                    Next
                                    HcDev = Limit(HcDev / CntAvg, 0.1, 20) 'Follows upper limit as in ASCE MOP 70 ed2 , p. 275 for upper, FAO56 eq. 62 and eq. 70 for lower
                                    CntAvg = 0
                                    HcEFC = Limit(MaximumCoverHeight, 0.1, 20) 'Follows upper limit as in ASCE MOP 70 ed2 , p. 275 for upper, FAO56 eq. 62 and eq. 70 for lower, max happens at efc and is assumed to remain constant the remainder of the season
                                    'Compute  U2 and RHmin for each period
                                    If DOYdev - DOYini > 0 Then
                                        RHminLimitedIni = 0
                                        U2LimitedIni = 0
                                        CntAvg = 0
                                        For DoY = DOYini To DOYdev - 1 'added by JBB
                                            'This code grabbing input data is copied form elsewhere
                                            Dim IWeather2 As Integer = WeatherPoint.RecordDate.IndexOf(DateSerial(YearStartWB - 1, 12, 31).AddDays(DoY)) 'IWeather will be the index of the current doy in the weatherpoint
                                            If IWeather2 > -1 Then
                                                Dim U2Day As Single = WeatherPoint.xTypWindSpeed2m(IWeather2) 'Grabs wind speed the weather point table JBB
                                                Dim RHminDay As Single = WeatherPoint.RelativeHumidity(IWeather2) 'Grabs Daily RHmin from the weather point table JBB
                                                RHminLimitedIni = RHminLimitedIni + RHminDay
                                                U2LimitedIni = U2LimitedIni + U2Day
                                                CntAvg = CntAvg + 1
                                            Else 'If either U2 or RHmin is not input for any day, default values are used for this period
                                                RHminLimitedIni = 45
                                                U2LimitedIni = 2
                                                CntAvg = 1
                                                Exit For
                                            End If
                                        Next
                                        RHminLimitedIni = Limit(RHminLimitedIni / CntAvg, 20, 80) 'FAO56 Eq. 70, though I don't think it is mentioned in ASCE MOP 70 ed.2
                                        U2LimitedIni = Limit(U2LimitedIni / CntAvg, 1, 6) 'FAO56 Eq. 70, though I don't think it is mentioned in ASCE MOP 70 ed.2
                                    End If
                                    If DOYefc - DOYdev > 0 Then
                                        RHminLimitedDev = 0
                                        U2LimitedDev = 0
                                        CntAvg = 0
                                        For DoY = DOYdev To DOYefc - 1 'added by JBB
                                            'This code grabbing input data is copied form elsewhere
                                            Dim IWeather2 As Integer = WeatherPoint.RecordDate.IndexOf(DateSerial(YearStartWB - 1, 12, 31).AddDays(DoY)) 'IWeather will be the index of the current doy in the weatherpoint
                                            If IWeather2 > -1 Then
                                                Dim U2Day As Single = WeatherPoint.xTypWindSpeed2m(IWeather2) 'Grabs wind speed the weather point table JBB
                                                Dim RHminDay As Single = WeatherPoint.RelativeHumidity(IWeather2) 'Grabs Daily RHmin from the weather point table JBB
                                                RHminLimitedDev = RHminLimitedDev + RHminDay
                                                U2LimitedDev = U2LimitedDev + U2Day
                                                CntAvg = CntAvg + 1
                                            Else 'If either U2 or RHmin is not input for any day, default values are used for this period
                                                RHminLimitedDev = 45
                                                U2LimitedDev = 2
                                                CntAvg = 1
                                                Exit For
                                            End If
                                        Next
                                        RHminLimitedDev = Limit(RHminLimitedDev / CntAvg, 20, 80) 'FAO56 Eq. 70, though I don't think it is mentioned in ASCE MOP 70 ed.2
                                        U2LimitedDev = Limit(U2LimitedDev / CntAvg, 1, 6) 'FAO56 Eq. 70, though I don't think it is mentioned in ASCE MOP 70 ed.2
                                    End If
                                    If DOYlate - DOYefc > 0 Then
                                        RHminLimitedMid = 0
                                        U2LimitedMid = 0
                                        CntAvg = 0
                                        For DoY = DOYefc To DOYlate - 1 'added by JBB
                                            'This code grabbing input data is copied form elsewhere
                                            Dim IWeather2 As Integer = WeatherPoint.RecordDate.IndexOf(DateSerial(YearStartWB - 1, 12, 31).AddDays(DoY)) 'IWeather will be the index of the current doy in the weatherpoint
                                            If IWeather2 > -1 Then
                                                Dim U2Day As Single = WeatherPoint.xTypWindSpeed2m(IWeather2) 'Grabs wind speed the weather point table JBB
                                                Dim RHminDay As Single = WeatherPoint.RelativeHumidity(IWeather2) 'Grabs Daily RHmin from the weather point table JBB
                                                RHminLimitedMid = RHminLimitedMid + RHminDay
                                                U2LimitedMid = U2LimitedMid + U2Day
                                                CntAvg = CntAvg + 1
                                            Else 'If either U2 or RHmin is not input for any day, default values are used for this period
                                                RHminLimitedMid = 45
                                                U2LimitedMid = 2
                                                CntAvg = 1
                                                Exit For
                                            End If
                                        Next
                                        RHminLimitedMid = Limit(RHminLimitedMid / CntAvg, 20, 80) 'FAO56 Eq. 70, though I don't think it is mentioned in ASCE MOP 70 ed.2
                                        U2LimitedMid = Limit(U2LimitedMid / CntAvg, 1, 6) 'FAO56 Eq. 70, though I don't think it is mentioned in ASCE MOP 70 ed.2
                                    End If
                                    If DOYterm - DOYlate >= 0 Then
                                        RHminLimitedLate = 0
                                        U2LimitedLate = 0
                                        CntAvg = 0
                                        For DoY = DOYlate To DOYterm  'added by JBB
                                            'This code grabbing input data is copied form elsewhere
                                            Dim IWeather2 As Integer = WeatherPoint.RecordDate.IndexOf(DateSerial(YearStartWB - 1, 12, 31).AddDays(DoY)) 'IWeather will be the index of the current doy in the weatherpoint
                                            If IWeather2 > -1 Then
                                                Dim U2Day As Single = WeatherPoint.xTypWindSpeed2m(IWeather2) 'Grabs wind speed the weather point table JBB
                                                Dim RHminDay As Single = WeatherPoint.RelativeHumidity(IWeather2) 'Grabs Daily RHmin from the weather point table JBB
                                                RHminLimitedLate = RHminLimitedLate + RHminDay
                                                U2LimitedLate = U2LimitedLate + U2Day
                                                CntAvg = CntAvg + 1
                                            Else 'If either U2 or RHmin is not input for any day, default values are used for this period
                                                RHminLimitedLate = 45
                                                U2LimitedLate = 2
                                                CntAvg = 1
                                                Exit For
                                            End If
                                        Next
                                        RHminLimitedLate = Limit(RHminLimitedLate / CntAvg, 20, 80) 'FAO56 Eq. 70, though I don't think it is mentioned in ASCE MOP 70 ed.2
                                        U2LimitedLate = Limit(U2LimitedLate / CntAvg, 1, 6) 'FAO56 Eq. 70, though I don't think it is mentioned in ASCE MOP 70 ed.2
                                    End If
                                    Dim KcbEndInt As Single = DailyKcb(DOYterm)
                                    Dim DailyKcboAdjMid As Single
                                    Dim DailyKcboAdjEnd As Single
                                    If MaxKcb < 0.4 Then
                                        DailyKcboAdjMid = 0.001 * (RHminLimitedMid - 45) 'ASCE MOP 70 EQ 10-15b, which is for Kcend, but I think it makes sense to use it here too
                                    Else
                                        DailyKcboAdjMid = (0.04 * (U2LimitedMid - 2) - 0.004 * (RHminLimitedMid - 45)) * ((HcEFC / 3) ^ 0.3)
                                    End If
                                    If KcbEndInt < 0.4 Then
                                        DailyKcboAdjEnd = 0.001 * (RHminLimitedLate - 45) 'ASCE MOP 70 EQ 10-15b, which is for Kcend, but I think it makes sense to use it here too
                                    Else
                                        DailyKcboAdjEnd = (0.04 * (U2LimitedLate - 2) - 0.004 * (RHminLimitedLate - 45)) * ((HcEFC / 3) ^ 0.3) 'Use HcEFF here too assuming that crop doesn't shrink much in senescence
                                    End If
                                    For DOY = DOYdev To DOYefc - 1
                                        DailyKcb(DOY) = DailyKcb(DOY) + 0 + (DOY - (DOYdev - 1)) * (DailyKcboAdjMid - 0) / (DOYefc - (DOYdev - 1)) 'We do doydev-1 because the day before it is initial stage with 0 adjustment, on doyefc it has mid adjustment
                                    Next
                                    For DOY = DOYefc To DOYlate - 1
                                        DailyKcb(DOY) = DailyKcb(DOY) + DailyKcboAdjMid 'Adds adjustment for climate(see EQ. 10-14 of ASCE MOP70 or FAO56 Ch7.
                                    Next
                                    For DOY = DOYlate To DOYterm
                                        DailyKcb(DOY) = DailyKcb(DOY) + DailyKcboAdjMid + (DOY - (DOYlate - 1)) * (DailyKcboAdjEnd - DailyKcboAdjMid) / (DOYterm - (DOYlate - 1)) 'Adds adjustment for climate (see EQ. 10-14 of ASCE MOP70 or FAO56 Ch7), this is a linear scaling which is what you get with FAO procedures, but applied to our non-linear Kcbs
                                    Next
                                End If

                                If DOYMaxBM < 0 Then DOYMaxBM = DOYterm
                                '*** end added by JBB
                                '*****     End Kcb Section
                                '***********************************************************************************

                                '*****     Start WB Section
                                'For DoY = DateIni To DateEnd 'loop through the growing season JBB This needs to be more dynamic 'commented out by JBB
                                For DoY = DOYStartWB To DOYEndWB 'added by JBB
                                    If DoY = 202 Then
                                        DoY = DoY 'added for debugging by JBB
                                    End If
                                    'Dim IWeather As Integer = WeatherPoint.RecordDate.IndexOf(DateSerial(MultispectralDate(0).Year - 1, 12, 31).AddDays(DoY)) 'IWeather will be the index of the current doy in the weatherpoint
                                    Dim IWeather As Integer = WeatherPoint.RecordDate.IndexOf(DateSerial(YearStartWB - 1, 12, 31).AddDays(DoY)) 'IWeather will be the index of the current doy in the weatherpoint
                                    If IWeather > -1 Then
                                        Dim RecordDate As DateTime = BaseDate.AddDays(DoY) 'Finds the date from the DOY by adding it to 12/31 of the previous year JBB
                                        'IWeather += 1'original code commented out by JBB' This appears to have been to keep data in line with images, which were one day out of sinc anyway JBB
                                        Dim ReferenceET As Single = WeatherPoint.ETDailyReference(IWeather) 'Grabs Daily ETo from the weather point table JBB
                                        Dim XEvaporation As Single = 1 'Default Value is 1
                                        If WeatherPointIndex.xEvaporationScalingFactor > -1 Then XEvaporation = WeatherPoint.xEvaporationScalingFactor(IWeather) 'Grabs the input evaporation term scaling factor,added by JBB, code following other code in this program
                                        Dim ETrefMnth As Single = 10 ' Default is high enough to not produce an adjustment regardless of reference type, added by JBB
                                        If WeatherPointIndex.xMonthlyEto > -1 Then ETrefMnth = WeatherPoint.xMonthlyETo(IWeather) 'Grabs monthly ETo if provided JBB, added by JBB following other code in SETMI
                                        Dim Fb As Single = 0.5 ' Default value (ASCE MOP 70, 2ed, 2016, p. 243) added by JBB
                                        If WeatherPointIndex.xFractionWetToday > -1 Then Fb = Limit(WeatherPoint.xFractionWetToday(IWeather), 0, 1) 'Added by JBB following other code in SETMI, ASCE MOP 70, 2ed, 2016 Eq. 9-28
                                        'Dim Irrigation As single = WeatherPoint.Irrigation(IWeather) 'Grabs irrigation from the weather point table JBB'Commented out by JBB
                                        '***********Added by JBB for irrigation and precipitation images ***********************
                                        Dim Irrigation As Single = 0 'Default is rainfed
                                        Dim ImageFlag As Boolean = False
                                        For m = 0 To DaysIrrigation.Count - 1 'Check for irrigation image
                                            If DaysIrrigation(m) = DoY Then
                                                Irrigation = Irrigations(m)
                                                ImageFlag = True
                                                Exit For
                                            End If
                                        Next
                                        If ImageFlag = False Then 'if there is not an irrigation image
                                            Irrigation = WeatherPoint.Irrigation(IWeather) 'Grabs irrigation from the weather point table JBB
                                        End If
                                        Dim Igross As Single = Irrigation 'Save gross irrigation for output
                                        Irrigation = Irrigation * AppEfficiency / 100 'Apply irrigation efficiency added by JBB
                                        SumInetOut = SumInetOut + Irrigation 'sums net irrigation for output added by JBB

                                        'For EQ. 9-28 of ASCE MOP70 ed2, 2016, aded by JBB copying code for irrigation and Precip 
                                        Dim IrrigationTomorrow As Single = 0 'zero if not provided or doy is end of wb
                                        ImageFlag = False
                                        If DoY < DOYEndWB Then
                                            For m = 0 To DaysIrrigation.Count - 1 'Check for next day irrigation image
                                                If DaysIrrigation(m) = DoY + 1 Then
                                                    IrrigationTomorrow = Irrigations(m)
                                                    ImageFlag = True
                                                    Exit For
                                                End If
                                            Next
                                            If ImageFlag = False Then 'if there is not an irrigation image for the next day
                                                IrrigationTomorrow = WeatherPoint.Irrigation(IWeather + 1) 'Grabs irrigation from the weather point table JBB
                                            End If
                                        End If
                                        IrrigationTomorrow = IrrigationTomorrow * AppEfficiency / 100 'Apply irrigation efficiency added by JBB

                                        Dim Precipitation As Single
                                        ImageFlag = False 'Reset image flag
                                        For m = 0 To DaysPrecipitation.Count - 1 'Check for precipitation image
                                            If DaysPrecipitation(m) = DoY Then
                                                Precipitation = Precipitations(m)
                                                ImageFlag = True
                                                Exit For
                                            End If
                                        Next
                                        If ImageFlag = False Then 'If there is not a preciptitation image
                                            Precipitation = WeatherPoint.Precipitation(IWeather) 'Grabs precipitation from the weather point table JBB
                                        End If

                                        'For EQ. 9-28 of ASCE MOP70 ed2, 2016, aded by JBB copying code for  Precip 
                                        Dim PrecipitationTomorrow As Single = 0
                                        If DoY < DOYEndWB Then
                                            ImageFlag = False 'Reset image flag
                                            For m = 0 To DaysPrecipitation.Count - 1 'Check for precipitation image
                                                If DaysPrecipitation(m) = DoY + 1 Then
                                                    PrecipitationTomorrow = Precipitations(m)
                                                    ImageFlag = True
                                                    Exit For
                                                End If
                                            Next
                                            If ImageFlag = False Then 'If there is not a preciptitation image
                                                PrecipitationTomorrow = WeatherPoint.Precipitation(IWeather + 1) 'Grabs precipitation from the weather point table JBB
                                            End If
                                        End If

                                        Dim ActualDepletion As Single
                                        Dim DepletionFlag As Boolean = False
                                        For m = 0 To DaysDepletion.Count - 1 'Check for depletion image
                                            If DaysDepletion(m) = DoY Then
                                                ActualDepletion = Depletions(m)
                                                DepletionFlag = True
                                                Exit For
                                            End If
                                        Next
                                        Dim ActualThetaV As Single
                                        Dim ThetaVFlag As Boolean = False
                                        For m = 0 To DaysThetaV.Count - 1 'Check for thetav image
                                            If DaysThetaV(m) = DoY Then
                                                ActualThetaV = ThetaVs(m) / 1000 'Converts from mm/m to m/m
                                                ThetaVFlag = True
                                                Exit For
                                            End If
                                        Next
                                        Dim ActualEvaporatedDepth As Single
                                        Dim EvaporatedDepthFlag As Boolean = False
                                        For m = 0 To DaysEvaporatedDepth.Count - 1 'Check for thetav image
                                            If DaysEvaporatedDepth(m) = DoY Then
                                                ActualEvaporatedDepth = EvaporatedDepths(m)
                                                EvaporatedDepthFlag = True
                                                Exit For
                                            End If
                                        Next

                                        Dim ActualSkinEvapDepth As Single 'Added by JBB copying code above
                                        Dim SkinEvapDepthFlag As Boolean = False
                                        For m = 0 To DaysSkinEvapDepth.Count - 1 'Check for skin evap image
                                            If DaysSkinEvapDepth(m) = DoY Then
                                                ActualSkinEvapDepth = SkinEvapDepths(m)
                                                SkinEvapDepthFlag = True
                                                Exit For
                                            End If
                                        Next
                                        '****************End added by JBB for irrigation and precipitation images *************

                                        '***********JBB added code for effective precipitation calculation *****************************
                                        Dim EffectivePrecipitation As Single
                                        Dim EffectivePptTomorrow As Single
                                        Dim CN As Single = CurveNumber
                                        Dim S_SCS As Single = 0
                                        Dim InitialAbstractions As Single = 0
                                        If EffectivePrecipitationType = EffectivePrecipType.SCS_Curve_Number Then ' Use Curve Number to Find Runoff JBB
                                            'Uses SCS Curve Number Method to calculate effective precipitation
                                            'Dim CurveNumber As single = 80 'Hard Coded in, But should be an option on the GUI JBB Made CN an input JBB
                                            Dim PptInch As Single = Precipitation / 25.4 'precip in inches
                                            Dim PptInchTomorrow As Single = PrecipitationTomorrow / 25.4 'tomorrow's precip in inches
                                            'If EffectivePrecipitationType = EffectivePrecipType.SCS_Curve_Number_Varying Then
                                            '    'Follows Eqs. 5.77-5.83 of Allen  et al. Ch. 5 in IA's Irrigation 6th Ed., 2011
                                            '    Dim CN1 As Single = CurveNumber / (2.281 - 0.01281 * CurveNumber)
                                            '    Dim CN3 As Single = CurveNumber / (0.427 - 0.00573 * CurveNumber)
                                            '    If De <= 0.5 * REW Then
                                            '        CN = CN1
                                            '    ElseIf De >= 0.7 * REW + 0.3 * TEW
                                            '        CN = CN3
                                            '    Else
                                            '        CN = ((De - 0.5 * REW) * CN1 + (0.7 * REW + 0.3 * TEW - De) * CN3) / (0.2 * REW + 0.3 * TEW)
                                            '    End If
                                            'End If
                                            S_SCS = 1000 / CN - 10
                                            InitialAbstractions = 0.2 * S_SCS
                                            Dim ExcessPrecipitation As Single = IIf(PptInch <= InitialAbstractions, 0, ((PptInch - InitialAbstractions) ^ 2) / (PptInch - InitialAbstractions + S_SCS))
                                            Dim ExcessPptTomorrow As Single = IIf(PptInchTomorrow <= InitialAbstractions, 0, ((PptInchTomorrow - InitialAbstractions) ^ 2) / (PptInchTomorrow - InitialAbstractions + S_SCS))
                                            ExcessPrecipitation = ExcessPrecipitation * 25.4 'convert Pexcess to mm
                                            ExcessPptTomorrow = ExcessPptTomorrow * 25.4 'Covert back to mm
                                            EffectivePrecipitation = Precipitation - ExcessPrecipitation
                                            EffectivePptTomorrow = PrecipitationTomorrow - ExcessPptTomorrow
                                            '***The logic below is commented out per ASCE MOP 70 ed2, 2016 p. 271
                                            'If ETReference = ETReferenceType.Short_Grass Then
                                            '    EffectivePrecipitation = IIf(Precipitation <= 0.2 * ReferenceET, 0, Precipitation - ExcessPrecipitation) 'The 0.2*ETref is from FAO56 p170 last Paragraph.
                                            'Else
                                            '    EffectivePrecipitation = IIf(Precipitation <= 0.15 * ReferenceET, 0, Precipitation - ExcessPrecipitation) 'The 0.15*ETref is 0.2*ETo/1.3 --> (Using ETr/ETo = 1.3, which is reasonable and the 1987 - 2014 May - Sep average for Mead Turf AWDN JBB
                                            'End If
                                        ElseIf EffectivePrecipitationType = EffectivePrecipType.Percent_Effective Then 'Uses Percent Precip JBB
                                            EffectivePrecipitation = PctPeff / 100 * Precipitation
                                            EffectivePptTomorrow = PctPeff / 100 * PrecipitationTomorrow
                                            '***The logic below is commented out per ASCE MOP 70 ed2, 2016 p. 271
                                            'If ETReference = ETReferenceType.Short_Grass Then
                                            '    ' EffectivePrecipitation = IIf(WeatherPoint.Precipitation(IWeather) <= 0.2 * ReferenceET, 0, TextEffectivePrecipPercent.Text / 100 * WeatherPoint.Precipitation(IWeather)) 'The 0.2*ETref is from FAO56 p170 last Paragraph. 
                                            '    EffectivePrecipitation = IIf(Precipitation <= 0.2 * ReferenceET, 0, PctPeff / 100 * Precipitation) 'The 0.2*ETref is from FAO56 p170 last Paragraph. PctPeff is now a tabular input JBB
                                            'Else
                                            '    ' EffectivePrecipitation = IIf(WeatherPoint.Precipitation(IWeather) <= 0.15 * ReferenceET, 0, TextEffectivePrecipPercent.Text / 100 * WeatherPoint.Precipitation(IWeather)) 'The 0.15*ETref is 0.2*ETo/1.3 --> (Using ETr/ETo = 1.3, which is reasonable and the 1987 - 2014 May - Sep average for Mead Turf AWDN JBB
                                            '    EffectivePrecipitation = IIf(Precipitation <= 0.15 * ReferenceET, 0, PctPeff / 100 * Precipitation) 'The 0.15*ETref is 0.2*ETo/1.3 --> (Using ETr/ETo = 1.3, which is reasonable and the 1987 - 2014 May - Sep average for Mead Turf AWDN JBB PctPeff is now a tabular input JBB
                                            'End If
                                        Else 'Warning message is earlier in the code
                                            EffectivePrecipitation = 0
                                            EffectivePptTomorrow = 0
                                        End If
                                        ROOut = Precipitation - EffectivePrecipitation 'Sums Runoff for output added by JBB
                                        SumROOut = SumROOut + ROOut
                                        SumPcpOut = SumPcpOut + Precipitation 'Sums Precip for output added by JBB
                                        '*************End JBB Added Code *********************************
                                        'Dim EffectivePrecipitation As single = IIf(WeatherPoint.Precipitation(IWeather) <= 0.2 * ReferenceET, 0, 0.8 * WeatherPoint.Precipitation(IWeather)) 'The 0.2*ETref is from FAO56 p170 last Paragraph. JBB'Commented out by JBB
                                        'Dim U2Limited As single = Limit(calcU2(WeatherPoint.WindSpeed(IWeather), WeatherPoint.AnemometerReferenceHeight(IWeather), 1), 1, 6) ' 'FAO 56 Eq 62 & p 123 JBB the -1 in calcU2 for cases when using gridded data ' **** There is a manual toogle here!!! if the third input is = -1 then no adjustment is made, otherwise the correction is done, 'ASCE STD MANUAL EQ 33 (ASSUMES GROUND COVER IS 10 CM TALL) JBB'Commented out by JBB because the input data for wind in the weather table is needed to be instantaneous for the energy balance
                                        'Dim RHminLimited As single = Limit(WeatherPoint.RelativeHumidity(IWeather), 20, 80) 'FAO 56 Eq 62 & p 123 JBB'Commented out by JBB because the input data for wind in the weather table is needed to be instantaneous for the energy balance


                                        'Dim HcTab = calcSeasonInterpolation(DoY, MinimumCoverHeight, MaximumCoverHeight, MaximumCoverHeight, DateIni, DateDev, DateMid, DateLate, DateEnd) 'Crop height from tabular heights calculated like a kcb JBB commented out by JBB and moved a few lines down
                                        'Dim HcStage = calcSeasonInterpolation(DoY, DaysCropHeight.ToArray, CropHeights) 'Added by JBB for Kcmax for short reference JBB 'commented out by JBB Kc Max is now a manual input!

                                        Dim Kcb As Single = KcbOffSeason 'added by JBB
                                        If KcbMethod = KcbType.VI_Interpolation Then 'logic added by JBB
                                            'Kcb = calcKcbInterpolation(DoY, Days.ToArray, SAVIs, KcbOffSeason, Cover, ETReference, KcbVI, KcbrfSlope, KcbrfIntercept, MaxKcbrfInput, MinKcbrfInput) 'Iterpolates the Kcbs JBB
                                            Kcb = DailyKcb(DoY)
                                        ElseIf KcbMethod = KcbType.Fitted_Curve Then 'logic added by JBB
                                            Kcb = DailyKcb(DoY) 'calcKcbPolynomial(Cover, DoY, DOYini, DOYefc, DOYterm, ETReference, KcbOffSeason) 'added by JBB
                                        ElseIf KcbMethod = KcbType.VI_Log_Regression Then  'GDD Log Interp Method
                                            Kcb = DailyKcb(DoY)
                                        End If 'logic added by JBB

                                        '**********************Hard Coded Paramters Added by JBB ***************THESE HAVE BEEN ADDED AS INPUT IN THE WEATHER FILE
                                        Dim U2Limited As Single = 2 'Hard Coded by JBB to make no adjustment
                                        Dim RHminLimited As Single = 45 'Hard Coded by JBB to make no adjustment
                                        Dim HcPeriod As Single = 3 'Will result in unity by default
                                        If ETReference = ETReferenceType.Short_Grass Then
                                            If DoY >= DOYlate And DoY <= DOYterm Then 'this logic is written in backward order because of how the logic is written in computed RHmin and U2
                                                U2Limited = U2LimitedLate
                                                RHminLimited = RHminLimitedLate
                                                HcPeriod = HcEFC 'used for both mid and late assuming crop doesn't shrink much
                                            ElseIf DoY >= DOYefc And DoY < DOYlate Then
                                                U2Limited = U2LimitedMid
                                                RHminLimited = RHminLimitedMid
                                                HcPeriod = HcEFC
                                            ElseIf DoY >= DOYdev And DoY < DOYefc Then
                                                U2Limited = U2LimitedDev
                                                RHminLimited = RHminLimitedDev
                                                HcPeriod = HcDev
                                            ElseIf DoY >= DOYini And DoY < DOYdev Then
                                                U2Limited = U2LimitedIni
                                                RHminLimited = RHminLimitedIni
                                                HcPeriod = HcIni
                                            End If
                                        End If

                                        '****Moved the KcMax stuff up here to flow better
                                        'These Kcmin/max equations are also used in computing the development doy in the Kcb portion of the code
                                        Dim KcMax As Single
                                        Dim KcMin As Single

                                        If ETReference = ETReferenceType.Short_Grass Then 'Allows Selection of Short or Tall Reference ET, This logic was added by JBB
                                            'KcMax = Math.Max(1.2 + (0.04 * (U2Limited - 2) - 0.004 * (RHminLimited - 45)) * (HcTab / 3) ^ (0.3), Kcb + 0.05) 'dimensionless ' FAO 56 EQ 72 JBB'Doesn't work correctly because RHmin and U2 should be stage averages !!!!!!! JBB ' Commented out by JBB because HcTab should not be used in the equation below h is actually average of crop stage
                                            'KcMax = Math.Max(1.2 + (0.04 * (U2Limited - 2) - 0.004 * (RHminLimited - 45)) * (HcStage / 3) ^ (0.3), Kcb + 0.05) 'Added by JBB with correct average stage crop height commented out by JBB
                                            'KcMax = Math.Max(KcMaxInput, Kcb + 0.05) 'Added by JBB 
                                            KcMax = Math.Max(KcMaxInput + (0.04 * (U2Limited - 2) - 0.004 * (RHminLimited - 45)) * (HcPeriod / 3) ^ (0.3), Kcb + 0.05) 'dimensionless ' FAO 56 EQ 72 JBB'This now works
                                            'KcMin = Limit(KcbIni, 0.15, 0.2)
                                            KcMin = KcbOffSeason 'Limit(KcbOffSeason, 0.15, 0.2)
                                            'If ETrefMnth < 5 Then
                                            '    TEW = TEWo * Math.Sqrt(ETrefMnth / 5) 'ASCE MOP 70, 2ed, 2016, Eq. 9-20b'added by JBB if ETrefMonth <=0, there will be problems, but I left it unbounded so that it would not give a false sense of being okay added by JBB
                                            '    REW = Math.Min(REWo, TEW * 0.999) 'REW cannot exceed TEW, following ASCE MOP 70, ed.2, 2016 p. 235 
                                            'End If
                                        Else
                                            'KcMax = Math.Max(1, Kcb + 0.05) 'added for tall referece based on ASABE Design and Op of Farm Irr. Sys p 246 JBB, Need to make reference type option on GUI 'Commented out by JBB
                                            KcMax = Math.Max(KcMaxInput, Kcb + 0.05) 'added by JBB 
                                            'KcMin = Limit(KcbIni, 0.12, 0.15) 'Added for tall reference by JBB, note this will need to be modified when Kcb method is moved from FAO to another method JBB
                                            KcMin = KcbOffSeason 'Limit(KcbOffSeason, 0.12, 0.15) 'Added for tall reference by JBB, 
                                            'If ETrefMnth < 6 Then
                                            '    TEW = TEWo * Math.Sqrt(ETrefMnth / 6) 'Adapted from ASCE MOP 70, 2ed, 2016, Eq. 9-20b'added by JBB assuming that ETr = 1.2ETo if ETrefMonth <=0, there will be problems, but I left it unbounded so that it would not give a false sense of being okay
                                            '    REW = Math.Min(REWo, TEW * 0.999) 'REW cannot exceed TEW, following ASCE MOP 70, ed.2, 2016 p. 235 
                                            'End If
                                        End If
                                        If ETrefMnth < 5 Then 'Has to be ETo
                                            TEW = TEWo * Math.Sqrt(ETrefMnth / 5) 'ASCE MOP 70, 2ed, 2016, Eq. 9-20b'added by JBB if ETrefMonth <=0, there will be problems, but I left it unbounded so that it would not give a false sense of being okay added by JBB
                                        Else
                                            TEW = TEWo
                                        End If
                                        REW = Math.Min(REWo, TEW * 0.999) 'REW cannot exceed TEW, following ASCE MOP 70, ed.2, 2016 p. 235 


                                        '*******************End Added by JBB *******************************************
                                        '****Note: the HcTab will be equal to MinimumCoverHeight before crop inititation, this won't effect Fc calcs, but should be noted JBB
                                        'Dim HcTab = Math.Min(MinimumCoverHeight + (MaximumCoverHeight - MinimumCoverHeight) * (Math.Max(Kcb, KcbIni) - KcbIni) / (KcbMid - KcbIni), MaximumCoverHeight) 'Crop height from tabular heights calculated as a linear relationship with Kcb JBB
                                        Dim HcTab As Single = Math.Max(Math.Min(MinimumCoverHeight + (MaximumCoverHeight - MinimumCoverHeight) * (Kcb - KcbOffSeason) / (MaxKcb - KcbOffSeason), MaximumCoverHeight), MinimumCoverHeight) 'Crop height from tabular heights calculated as a linear relationship with Kcb JBB
                                        If HcTab < HcTabPrevious And DoY <= DOYterm Then HcTab = HcTabPrevious 'keeps crop from shrinking until end of season added by JBB
                                        HcTabPrevious = HcTab 'added by JBB
                                        If ETReference = ETReferenceType.Short_Grass Then HcTab = DailyHcETo(DoY) 'Overwright with Hc from unadjusted Kcb, this helps otherwise there are two Hc's going on, the one to help compute Kcb and the one based on Kcb
                                        '***Moved Fc Up here to flow better, Fc uses crop height computed daily
                                        Dim Fc As Single = 0 'added by JBB
                                        'This code is also used in the Kcb section to compute DoYdev, but doesn't include climate adjustments to KcMax (because those adjustments are dependent on that calcuation).
                                        If FcMethodWB = FcType.Vegetation_Index And (KcbMethod = KcbType.VI_Log_Regression Or KcbMethod = KcbType.VI_Interpolation) Then 'Added by JBB as a better Fc Methodology
                                            'Fc = 1 - ((SAVImaxFc - Limit(DailySAVI(DoY), SAVImin, SAVImaxFc)) / (SAVImaxFc - SAVImin)) ^ SAVIFcExponent 'See Choudhury 1994 Eq. 15
                                            'Fc = Limit(Fc, 0, 0.99) 'Limits match FAO 56 Eq. 76
                                            Fc = calcFcWB(Cover, DailySAVI(DoY), DailySAVI(DoY), KcbVI, MaxVIInput, MinVIInput, ExpVIInput, 0.99, 0) 'We use DailySAVI for both NDVI and SAVI, because that is where the NDVI will be saved if that index is selected, sorry for the confusion, JBB
                                        Else 'Follow FAO 56, this seems to underestimate Fc appreciably
                                            ' If Kcb >= KcMin Then Fc = Math.Max(Math.Min((Math.Max(Kcb - KcMin, 0.01) / Math.Max(KcMax - KcMin, 0.01)) ^ (1 + 0.5 * HcTab), 0.99), 0) 'added by JBB to prevent undefined # if numerator goes negative this matches related text in FAO56 Eq 76
                                            Fc = Math.Max(Math.Min((Math.Max(Kcb - KcMin, 0.01) / Math.Max(KcMax - KcMin, 0.01)) ^ (1 + 0.5 * HcTab), 0.99), 0) 'added by JBB to prevent undefined # if numerator goes negative this matches related text in FAO56 Eq 76
                                        End If

                                        'If Kcb > 0.45 And DoY >= DateDev Then Kcb += (0.04 * (U2Limited - 2) - 0.004 * (RHminLimited - 45)) * (HcTab / 3) ^ (0.3) 'FAO 56 Eq 62 ' And FAO 56 EQ 65 JBB Commented out by JBB because it was not correctly applied, Htab is fine, but U and RH should be the average for the Kcmid and KCend periods, and should be applied before interpolation, U as input in the weather data is also used in the TSEB so it should not be used here as it isn't the correct U for this equation JBB
                                        'Dim Zri = Math.Min(Zrmin + (Zrmax - Zrmin) * (Kcb - KcbIni) / (KcbMid - KcbIni), Zrmax) 'Root depth changes linearly with Kcb-devel. and reaches a mazimum with Kcbmid. JBB commented out by JBB
                                        'Dim Zri = Math.Min(Zrmin + (Zrmax - Zrmin) * (Math.Max(Kcb, KcbIni) - KcbIni) / (KcbMid - KcbIni), Zrmax) 'Root depth changes linearly with Kcb-devel. and reaches a mazimum with Kcbmid. added by JBB
                                        'Dim Zri = Math.Max(Math.Min(Zrmin + (Zrmax - Zrmin) * (Kcb - KcbOffSeason) / (MaxKcb - KcbOffSeason), Zrmax), Zrmin) 'Root depth changes linearly with Kcb-devel. and reaches a mazimum with Kcbmid. added by JBB

                                        '*****Grow the root zone first!
                                        Dim Zri As Single = Zrmin
                                        Dim ZL As Single
                                        If DoY <= DOYterm Then 'compute Zri This methodolgy is kinda like Eq 5.65 in Irrigation 6th Ed. IA. 2011.
                                            Zri = Limit(Zrmin + (ZrmaxIn - Zrmin) * (DoY - DOYini) / (DOYefc - DOYini), Zrmin, Zrmax) 'Root depth changes linearly with Kcb-devel. and reaches a mazimum with Kcbmid. added by JBB
                                        End If
                                        If Zri < ZriPrevious And DoY <= DOYterm Then Zri = ZriPrevious 'added by JBB so rootzone cannot decrease until the end of the season, this will be a problem in multi cropping systems. JBB
                                        'If Zri > DpthMax Then Zri = DpthMax  'Added by JBB 2/9/2018, limits root growth to no greater than the total depth of soil input into the model'<-- Already done above now
                                        '*****Added by JBB Compute Rootzone Weighted Avg. Parameters
                                        Dim IntM As Integer = 1 'Layer that root zone bottom is in
                                        Dim IntK As Integer = 1 'Layer that root zone bottom was in in previous time stamp
                                        '***Compute layered soil variables
                                        ArrWP(0) = 0
                                        ArrFC(0) = 0
                                        VWCsat(0) = 0
                                        Ksat(0) = 0
                                        For iii = 1 To LyrCnt
                                            ArrDelZr(iii) = Limit(Zri - ArrDpth(iii - 1), 0, ArrThick(iii))
                                            ArrDelZL(iii) = Limit(ArrDpth(iii) - Zri, 0, ArrThick(iii))
                                            ArrFC(0) = ArrFC(0) + ArrFC(iii) * ArrDelZr(iii)
                                            ArrSWSFC(iii) = ArrFC(iii) * ArrDelZL(iii) * 1000
                                            ARRSWSsat(iii) = VWCsat(iii) * ArrDelZL(iii) * 1000
                                            ArrWP(0) = ArrWP(0) + ArrWP(iii) * ArrDelZr(iii)
                                            Ksat(0) = Ksat(0) + ArrDelZr(iii) / Ksat(iii) 'Eq. 3.10 of Radcliffe, D. E. and J. Simunek. 2010. Soil Physics with HYDRUS. CRC Press, Taylor and Francis Group, New York.  p. 93.
                                            VWCsat(0) = VWCsat(0) + VWCsat(iii) * ArrDelZr(iii)
                                            If Zri >= ArrDpth(iii - 1) And Zri <= ArrDpth(iii) Then IntM = iii
                                            If ZriPrevious >= ArrDpth(iii - 1) And ZriPrevious <= ArrDpth(iii) Then IntK = iii
                                        Next
                                        ArrFC(0) = ArrFC(0) / Zri
                                        ArrWP(0) = ArrWP(0) / Zri
                                        ArrSWSFC(0) = ArrFC(0) * Zri * 1000 'could be a litte more efficient, but this makes sense
                                        Ksat(0) = Zri / Ksat(0) 'Eq. 3.10 of Radcliffe, D. E. and J. Simunek. 2010. Soil Physics with HYDRUS. CRC Press, Taylor and Francis Group, New York.  p. 93.
                                        VWCsat(0) = VWCsat(0) / Zri
                                        ArrTau(0) = Limit(0.0866 * Ksat(0) ^ 0.35, 0, 1)  'Raes et al 2017. Aqua Crop 6.0
                                        ARRSWSsat(0) = VWCsat(0) * Zri * 1000

                                        '*****Initiate the water balance, also moved this code up several lines to be above termination
                                        If DoY = DOYStartWB Then 'for the first day in the WB added by JBB
                                            ' DrLast = (θFC - θvLini) * Zri * 1000 'mm ' This is Dr,i-1 , see FAO 56 EQ. 85 added by JBB
                                            'θvLPrevious = (θvLini) 'soil moisture in the lower soil layer added by JBB
                                            ProfileAvgThtavPrev = 0
                                            ArrThtaIni(0) = 0
                                            For iii = 1 To 3
                                                ArrThtaIni(0) = ArrThtaIni(0) + ArrThtaIni(iii) * ArrDelZr(iii)
                                                ArrThta(iii) = ArrThtaIni(iii)
                                                ArrThtaPrev(iii) = ArrThta(iii)
                                                ArrDelZLprev(iii) = ArrDelZL(iii)
                                                ArrSWS(iii) = ArrThtaIni(iii) * 1000 * ArrDelZL(iii)
                                                ProfileAvgThtavPrev = ProfileAvgThtavPrev + ArrThtaIni(iii) * ArrThick(iii) / Zrmax
                                            Next
                                            ProfileThetaIni = ProfileAvgThtavPrev
                                            ArrThtaIni(0) = ArrThtaIni(0) / Zri
                                            DrLast = (ArrFC(0) - ArrThtaIni(0)) * Zri * 1000 'mm ' This is Dr,i-1 , see FAO 56 EQ. 85 added by JBB We don't add depletion from root growth because they haven't grown yet
                                            DrIni = DrLast
                                            ArrSWS(0) = Zri * ArrThtaIni(0) * 1000
                                            ' De = Limit((θFC - θvLini) * Ze * 1000, 0, TEW) 'Initial depth evaporated assumes that it is same as rest of soil profile thetav added by JBB
                                            De = Limit((ArrFC(1) - ArrThtaIni(1)) * Ze * 1000, 0, TEW) 'Initial depth evaporated assumes that it is same as rest of soil profile thetav added by JBB The evaporative layer is all in the topmost layer
                                            DREW = Limit((ArrFC(1) - ArrThtaIni(1)) * Ze * 1000, 0, REW) 'Initial depth evaporated assumes that it is same as rest of soil profile thetav added by JBB The evaporative layer is all in the topmost layer
                                            ZriPrevious = Zri 'added by JBB
                                            FwPrevious = 1 'added by JBB assumes initial soil wetting was from rain
                                        End If
                                        '*******************End Initiation

                                        '***Termination and redistribution of water into respective layers added by JBB
                                        If DoY = DOYterm + 1 Then 'added by JBB to adjust upper and lower soil layers at end of growing season
                                            If Zri < ZriPrevious Then
                                                'Dim DrLowerLast = (θFC - θvLPrevious) * (Zrmax - ZriPrevious) * 1000 'intermediate lower layer depletion calc
                                                Dim ArrDrLwrLast(3) As Single
                                                Dim ArrDrLast(3) As Single
                                                'Dim DrLInt0 As Single
                                                Dim DrLastInt As Single = DrLast
                                                Dim ArrDelZRprev(3) As Single
                                                Dim DrtLast As Single = DrLast - De 'This is DrLast excluding evaporation, this is available for redistribution
                                                Dim DrLastDenom As Single = 0
                                                Dim DrLastNumer As Single = 0

                                                ' DrLInt0 = DrLast * Zri * (ArrFC(1) - ArrWP(1)) / (ZriPrevious * (FCrPrev - WPrPrev))
                                                'ArrDrLast(0) = Math.Max(De, DrLInt0)
                                                'Dim DrLLast As Single = DrLast - ArrDrLast(0) 'subtracts out the portion of DrLast in the upper layer including difference in De and DrLast for the new root zone
                                                'Recompute previous FC and WP to account for subtracting out Ze from root zone this will help with weighted averages
                                                'ArrDelZRprev(1) = Limit(ZriPrevious - ArrDpth(0) - Zri, 0, ArrThick(1) - Zri)
                                                'ArrDelZRprev(1) = Limit(ZriPrevious - ArrDpth(0) - Zri, 0, ArrThick(1))
                                                ArrDelZRprev(1) = Limit(ZriPrevious - ArrDpth(0), 0, ArrThick(1))
                                                'ArrDelZLprev(1) = Limit(ArrDpth(1) - ZriPrevious, 0, ArrThick(1) - Zri) 'This is how much of the top layer could have been in the lower layer in the previous time step, since Zri is now Zrmin, this should work out
                                                ArrDelZLprev(1) = Limit(ArrDpth(1) - ZriPrevious, 0, ArrThick(1)) 'This is how much of the top layer could have been in the lower layer in the previous time step, since Zri is now Zrmin, this should work out
                                                DrLastNumer = (ArrDelZr(1) - Zrlim) * (ArrFC(1) - ArrWP(1))
                                                DrLastDenom = (ArrDelZRprev(1) - Zrlim) * (ArrFC(1) - ArrWP(1))
                                                For iii = 2 To LyrCnt
                                                    ArrDelZLprev(iii) = Limit(ArrDpth(iii) - ZriPrevious, 0, ArrThick(iii)) 'this was redundent
                                                    ArrDelZRprev(iii) = Limit(ZriPrevious - ArrDpth(iii - 1), 0, ArrThick(iii))
                                                    DrLastNumer = DrLastNumer + ArrDelZr(iii) * (ArrFC(iii) - ArrWP(iii))
                                                    DrLastDenom = DrLastDenom + ArrDelZRprev(iii) * (ArrFC(iii) - ArrWP(iii))
                                                Next
                                                If DrLastDenom = 0 Then
                                                    ArrDrLast(0) = DrLast
                                                Else
                                                    ArrDrLast(0) = De + DrtLast * DrLastNumer / DrLastDenom 'New DrLast
                                                End If

                                                Dim DrLLast As Single = DrLast - ArrDrLast(0)
                                                FCrPrev = 0
                                                WPrPrev = 0

                                                For iii = 1 To LyrCnt
                                                    FCrPrev = FCrPrev + ArrFC(iii) * (ArrDelZRprev(iii) - ArrDelZr(iii))
                                                    WPrPrev = WPrPrev + ArrWP(iii) * (ArrDelZRprev(iii) - ArrDelZr(iii))
                                                Next
                                                FCrPrev = FCrPrev / (ZriPrevious - Zri)
                                                WPrPrev = WPrPrev / (ZriPrevious - Zri)

                                                For iii = 1 To LyrCnt 'Find that part of each layer that was below the root zone before, and compute it's Dr (should be pretty uniform)
                                                    ArrDrLwrLast(iii) = (ArrFC(iii) - ArrThta(iii)) * ArrDelZLprev(iii) * 1000 'intermediate lower layer depletion calc
                                                    ArrDrLast(iii) = DrLLast * (ArrDelZRprev(iii) - ArrDelZr(iii)) * (ArrFC(iii) - ArrWP(iii)) / ((ZriPrevious - Zri) * (FCrPrev - WPrPrev))
                                                    ArrDrLwrLast(iii) = ArrDrLast(iii) + ArrDrLwrLast(iii)
                                                    If ArrDelZL(iii) > 0 Then
                                                        ArrThta(iii) = ArrFC(iii) - ArrDrLwrLast(iii) / 1000 / ArrDelZL(iii) 'Now we divide by the current values of thickness in the lower level
                                                    Else
                                                        ArrThta(iii) = ArrFC(iii)
                                                    End If
                                                    ArrThtaPrev(iii) = ArrThta(iii)
                                                    ArrSWS(iii) = ArrThta(iii) * 1000 * ArrDelZL(iii)
                                                Next
                                                DrLast = ArrDrLast(0) 'Root zone
                                                ArrSWS(0) = ArrFC(0) * Zri * 1000 - DrLast
                                                '******Note
                                                '******It may be possible if De > 0 and DrLast <0 for some layers to exceed saturation at termination with this process,
                                                '******It seems the chance is small and not worth the logic to avoid. The water balance should drain off that excess water rather quickly should it every occur
                                                '*****Commented out how it used to work
                                                'DrLowerLast = DrLowerLast + DrLast ' * (ZriPrevious - Zri) / ZriPrevious 'depletion in lower layer + depletion in rootzone that is now going to be the lower layer
                                                ' DrLast = DrLowerLast * Zri / Zrmax 'Intermediate calc, evenly partitions depletion from ending of growing season to root zone and then carries on the depth for the upper level JBB, there may be a slight error with conflict with De
                                                ' DrLast = Math.Max(DrLast, De) 'Keeps the upper water balance consistent
                                                ' If Math.Abs(Zri - Zrmax) < 0.0001 Then
                                                '     θvLPrevious = θFC 'In this case the lower layer has a zero thickness so the value doesn't matter we just need to avoid a divide by zero
                                                ' Else
                                                'θvLPrevious = θFC - (DrLowerLast - DrLast) / (Zrmax - Zri) / 1000 'water content in new lower layer evenly partitioned unless DrLast = De, in which case the lower level is a little wetter to conserve mass
                                                'End If
                                                FCrPrev = ArrFC(0)
                                                WPrPrev = ArrWP(0)
                                                ZriPrevious = Zri
                                                IntK = IntM 'sets index of previous layer of root zone bottom at current layer of root zone bottom
                                                DelDrlastFromDelZr = DrLast - DrLastInt
                                                'End If
                                                '************************End Termination
                                            End If
                                        Else 'Prevents DrLast and DrPrevLast from being overwritten on the day after termination
                                            '*************Update DrLast for Root Growth*******************************

                                            DelDrlastFromDelZr = 0
                                            For iii = IntK To IntM
                                                'ArrSWS(0) = ArrSWS(0) + 1000 * ArrThta(iii) * (ArrDelZLprev(iii) - ArrDelZL(iii)) '38.4 of my notes
                                                DelDrlastFromDelZr = DelDrlastFromDelZr + (ArrFC(iii) - ArrThta(iii)) * (ArrDelZLprev(iii) - ArrDelZL(iii)) * 1000 'in terms of DrLast instead
                                            Next
                                            DrLast = DrLast + DelDrlastFromDelZr
                                        End If
                                        '***The line below could be changed to be based off an EFC date
                                        'If DoY >= DateMid Then Zri = Zrmax 'After mid stage is reached the root zone depth is also at a maximum JBB commented out by JBB
                                        'Dim KcMax As single = Math.Max(1.2 + (0.04 * (U2Limited - 2) - 0.004 * (RHminLimited - 45)) * (HcTab / 3) ^ (0.3), Kcb + 0.05) 'dimensionless ' FAO 56 EQ 72 JBB'Doesn't work correctly because RHmin and U2 should be stage averages !!!!!!! JBB
                                        'Dim KcMin As single = Limit(KcbIni, 0.15, 0.2)
                                        'Dim Fc As single = ((Kcb - KcMin) / (KcMax - KcMin)) ^ (1 + 0.5 * HcTab) 'FAO 56 EQ 76, But don't we have lots of ways to get Fc already in this program? JBB commented out by JBB

                                        If DoY = 312 Then
                                            DoY = DoY
                                        End If


                                        'Compute Ke first so we can estimate ETc for p for RAW, this differs from the order in ASCE MOP 70 ed.2, 2016 (p. 272), but seems necessary
                                        Dim Fwi As Single = 1 'FAO 56 Table 20 'Precip, Sprinkle, Border, Basin JBB
                                        Select Case IrrigationMethod ' Averages from FAO50 Table 20 JBB
                                            Case Functions.IrrigationMethod.Drip : Fwi = 0.35
                                            Case Functions.IrrigationMethod.Drip_Shaded : Fwi = 0.35 * (1 - (2 / 3) * Fc) 'added by JBB per ASCE MOP70, ed.2, 2016, p. 239
                                            Case Functions.IrrigationMethod.Furrow_Narrow_Bed : Fwi = 0.8
                                            Case Functions.IrrigationMethod.Furrow_Narrow_Bed_Alternating : Fwi = 0.45
                                            Case Functions.IrrigationMethod.Furrow_Wide_Bed : Fwi = 0.5
                                            Case Functions.IrrigationMethod.Furrow_Wide_Bed_Alternating : Fwi = 0.35
                                        End Select
                                        Dim Fw As Single = Fwi
                                        If IrrigationMethod = IrrigationMethod.Subsurface Then 'If subsurface irrigation, then only precip wets the surface
                                            Fw = 1 'A weighted average Few, for precip only JBB
                                        Else 'All other forms of irrigation
                                            Fw = IIf(Irrigation + EffectivePrecipitation = 0, FwPrevious, Irrigation / (Irrigation + EffectivePrecipitation) * Fw + EffectivePrecipitation / (Irrigation + EffectivePrecipitation)) 'A weighted average Few JBB
                                        End If
                                        FwPrevious = Fw 'added by JBB now grabs previous Fw
                                        Dim Few As Single = Limit(Math.Min(1 - Fc, Fw), 0.01, 1) 'FAO56 Eq. 75 & related text JBB
                                        Dim KeScale As Single = Limit(XEvaporation, 0, 1) 'Limit input evap scaling factor to 0-1 JBB
                                        Dim Fstage1 As Single = 0 'Default is not to include skin evap JBB
                                        If WBSkinEvapMethod = WBSkinMethod.Include_Skin_Evap Then Fstage1 = Limit((REW - DREW) / (KcMax * ReferenceET), 0, 1) 'ASCE MOP 70, 2ed., 2016, eq. 9-22, but using KcMax not KeMax per p. 235
                                        'Dim Kr As Single = (TEW - Math.Max(De, REW)) / (TEW - REW) 'FAO56 Eq. 74 JBB
                                        Dim Kr As Single = Fstage1 + (1 - Fstage1) * Limit((TEW - De) / (TEW - REW), 0, 1) 'ASCE MOP 70, 2ed, 2016, eq. 9-21
                                        Dim Ke As Single = Math.Min(Kr * (KcMax - Kcb), Few * KcMax) 'FAO56 Eq. 71 JBB, 
                                        Ke = KeScale * Ke 'added by JBB to allow for dampening of evaporation term
                                        Dim EvapNS As Single = Ke * ReferenceET 'ASCE MOP 70 ed.2, 2016 p272, for non-stressed conditions on that day
                                        Dim Ke_ns As Single = Ke 'This is the non-stressed Ke
                                        Dim Kc As Single = Kcb + Ke 'FAO56 Eq 57 JBB
                                        Dim ETc As Single = Kc * ReferenceET 'mm 'FAO56 EQ 69 JBB
                                        'Dim P As Single = Limit((Ptab + 0.04 * (5 - ETc)) * (1.15 - 1.25 * (θFC - θWP)), 0.1, 0.8) '(P adjustment curve fit to data in FAO 56 Table 19), <-- This comment makes no sense! Ptab + 0.04 * (5 - ETc)) and the 0.1 to 0.8 limits are from FAO 56 p 162
                                        Dim P As Single = Limit((Ptab + 0.04 * (5 - ETc)), 0.1, 0.8) '(P adjustment curve fit to data in FAO 56 Table 19) matches FAO 56 also ASCE MOP 70 ed.2, p. 483
                                        'Dim TAW = 1000 * (θFC - θWP) * Zri 'mm ' FAO56 Eq 73 JBB
                                        Dim TAW As Single = 1000 * (ArrFC(0) - ArrWP(0)) * Zri 'mm ' FAO56 Eq 73 JBB
                                        Dim RAW As Single = P * TAW ' FAO 56 EQ 83 JBB
                                        'DrLast is upded above now
                                        'DrLast = DrLast + (θFC - θvLPrevious) * Math.Max(Zri - ZriPrevious, 0) 'Adds depletion from root growth into DrLast
                                        Dim Ks As Single = (TAW - Math.Max(DrLast, RAW)) / (TAW - RAW) 'FAO56 EQ 84, The max logic limits it to <= 1, DrLast is appropriate following FAO 56 Examp 37 JBB
                                        Ke = Math.Min(Kr * (KcMax - Kcb * Ks), Few * KcMax) 'ASCE MOP 70, ed.2, 2016, Eq. 9-19
                                        Ke = KeScale * Ke
                                        Dim Ke_strss As Single = Ke 'saves stressed Ke
                                        Dim Tmax As Single = WeatherPoint.DailyMaximumTemperature(IWeather) 'Grabs Daily Tmax from the weather point table JBB
                                        Dim Tmin As Single = WeatherPoint.DailyMinimumTemperature(IWeather) 'Grabs Daily Tmin from the weather point table JBB
                                        Dim Tavg As Single = (Tmax + Tmin) / 2
                                        Dim Srx As Single = 1 'Copying a spreadsheet supplied by Ivo Goncalves Zution, 2/16/2018
                                        Dim Srn As Single = 0.001 'Copying a spreadsheet supplied by Ivo Goncalves Zution, 2/16/2018
                                        Dim r As Single = -2 * Math.Log((2 * Srn * Srx - Srn) / (Srx - Srn)) 'See AquaCrop v6 Manual, Raes et al. 2017, FAO, Rome, Eq. 3.2c
                                        Dim Srel As Single = Limit((TstressMax - Tavg) / (TstressMax - TstressBase), 0, 1) 'See Campos et al. 2018 J. Ag. Forest. Meteor.
                                        Dim Kst As Single = Srn * Srx / (Srn + (Srx - Srn) * Math.Exp(-r * (1 - Srel))) 'Currently not used, but this is temperature stress following Campos et al. 2018
                                        Kst = (Kst - Srn) / (1 - 2 * Srn) 'Scales from 1 to 0 since srn <> 0
                                        Dim Te As Single = 0
                                        Dim Tdrew As Single = 0
                                        If WBTeMethod = WBTeMethod.Include_Transpiration Or WBTeMethod = WBTeMethod.Include_Transp_frm_Skn_Lyr_Also Then
                                            Kt = 0
                                            Dim KtNum As Single = Limit(1 - De / TEW, 0.001, 1) 'ASCE MOP 70, 2ed, 2016, EQ. 9-31 with upper limit because I don't think the equation was expecting drLast to be able to be <0, whcih it will in this version of SETMI the other limit follows MOP70 p. 245
                                            Dim KtDen As Single = Limit(1 - DrLast / TAW, 0.001, 1) 'ASCE MOP 70, 2ed, 2016, EQ. 9-31 with upper limit because I don't think the equation was expecting drLast to be able to be <0, whcih it will in this version of SETMI the other limit follows MOP70 p. 245
                                            Kt = Math.Min(1, KtNum / KtDen * ((Ze / Zri) ^ 0.6)) 'ASCE MOP 70, 2ed, 2016 EQ. 9-31 p. 245
                                            Te = Kt * Ks * Kcb * ReferenceET 'ASCE MOP 70 2ed, 2016, Eq. 9-30
                                            If WBTeMethod = WBTeMethod.Include_Transp_frm_Skn_Lyr_Also Then 'added by JBB, this allows transpiration from the skin layer to help conserve mass in the case of a small root zone and a kcb>0
                                                If De < TEW Then
                                                    Tdrew = Math.Min((1 - DREW / REW) / (1 - De / TEW) * Te, Te)
                                                Else
                                                    Tdrew = 0
                                                End If
                                            End If
                                        End If
                                        Dim Kcadj As Single = Ks * Kcb + Ke 'ASCE MOP 70 ed.2, 2016 Eq. 10-2
                                        'Dim ETcAdjusted As Single = (Ks * Kcb + Ke) * ReferenceET ' FAO56 Eq 80 JBB
                                        Dim ETcAdjusted As Single = Kcadj * ReferenceET 'ASCE MOP 70 ed.2, 2016 Eq. 8-4
                                        Dim ETcAdjOut As Single = ETcAdjusted 'added by JBB for output
                                        Dim ZriOut As Single = Zri 'added by JBB for output
                                        Dim AvailWaterAboveMAD As Single 'added by JBB as a metric for irrigtion scheduling JBB
                                        Dim EffectiveIrrigation As Single = Irrigation / Fwi 'FAO 56 EQ 77 JBB
                                        Dim EffectiveIrrTomorrow As Single = IrrigationTomorrow / Fwi 'FAO 56 EQ 77 JBB added by JBB for ASCE MOP 70 ed.2, 2016 eq 9-28 etc
                                        Dim EffectiveEvaporation As Single = Math.Min(Ke * ReferenceET / Few, TEW) 'FAO 56 EQ 77 JBB, the TEW limit is unnecessary here.
                                        Dim DePrev As Single = De
                                        Dim DrewPrev As Single = DREW

                                        Dim OldDe As Single = 0 'added by JBB for output
                                        Dim OldDREW As Single = 0 'added by JBB for output
                                        Dim Evap As Single = Ke * ReferenceET 'ASCE MOP 70 ed.2, 2016 p272
                                        If IrrigationMethod = IrrigationMethod.Subsurface Then 'Surface isn't wetted from irrigation
                                            'De = Limit(De - EffectivePrecipitation + EffectiveEvaporation, 0, TEW) ' FAO 56 EQ 77 including Eq 78, w/out irrigation term Doesn't Include Deep Perc. Term! JBB
                                            De = Limit(De - ((1 - Fb) * (EffectivePrecipitation) + Fb * (EffectivePptTomorrow)) + Evap / Few + Te, 0, TEW) 'ASCE MOP 70 ed.2, 2016 eq. 9-28 w/out irrigation term JBB
                                            DREW = Limit(DREW - ((1 - Fb) * (EffectivePrecipitation) + Fb * (EffectivePptTomorrow)) + Evap / Few + Tdrew, 0, REW) 'ASCE MOP 70 ed.2, 2016 eq. 9-28 w/out irrigation term JBB Eq. 9-29 Also Irrigation 6th Ed., IA, 2011 Eq. 5.56
                                        Else 'All other forms of irrigation
                                            ' De = Limit(De - EffectivePrecipitation - EffectiveIrrigation + EffectiveEvaporation, 0, TEW) ' FAO 56 EQ 77 including Eq 78, Doesn't Include Deep Perc. Term! JBB
                                            De = Limit(De - ((1 - Fb) * (EffectivePrecipitation + EffectiveIrrigation) + Fb * (EffectivePptTomorrow + EffectiveIrrTomorrow)) + Evap / Few + Te, 0, TEW) 'ASCE MOP 70 ed.2, 2016 eq. 9-28
                                            DREW = Limit(DREW - ((1 - Fb) * (EffectivePrecipitation + EffectiveIrrigation) + Fb * (EffectivePptTomorrow + EffectiveIrrTomorrow)) + Evap / Few + Tdrew, 0, REW) 'ASCE MOP 70 ed.2, 2016 eq. 9-29 Also Irrigation 6th Ed., IA, 2011 Eq. 5.56
                                        End If
                                        If EvaporatedDepthFlag = True And AssimilationMethod <> DataAssimilation.No_Assimilation Then
                                            OldDe = De 'added by JBB to preserve old value
                                            De = Limit(calcETAssimilation(OldDe, ActualEvaporatedDepth, AssimilationMethod, EvaporatedDepthWeight), 0, TEW) 'added by JBB, The limit was included for consistency, but really the soil should be allowed to have negative depletion for 3 days or something. JBB
                                        End If
                                        If SkinEvapDepthFlag = True And AssimilationMethod <> DataAssimilation.No_Assimilation Then
                                            OldDREW = DREW 'added by JBB to preserve old value
                                            DREW = Limit(calcETAssimilation(OldDREW, ActualSkinEvapDepth, AssimilationMethod, SkinEvapWeight), 0, REW) 'added by JBB, The limit was included for consistency, but really the soil should be allowed to have negative depletion for 3 days or something. JBB
                                        End If
                                        ' Dr = Limit(DrLast - EffectivePrecipitation - EffectiveIrrigation - CR + ETcAdjusted, 0, TAW) ' FAO 56 EQ 86, No Deep Perc. Term, CR = 0 commented out by JBB
                                        'ArrSWS(0) = ArrSWS(0) + EffectivePrecipitation + Irrigation - ETcAdjusted + CR
                                        'ArrThta(0) = ArrSWS(0) / Zri / 1000
                                        'ArrDPLimit(0) = Limit(ArrTau(0) * (Math.Exp(ArrThta(0) - ArrFC(0)) - 1) / (Math.Exp(VWCsat(0) - ArrFC(0)) - 1), 0, ArrTau(0))

                                        Dim Fnet As Single = EffectivePrecipitation + Irrigation + CR - ETcAdjusted 'Net infiltration (plus caplilary rise)
                                        Dim Fexc As Single = 0 'default is no excess if no limit is applied to drainage it will all drain
                                        Dim FexcOut As Single = 0
                                        Dim AssimExc As Single = 0 'This is excess water from assimilation
                                        Dim AssimExcL As Single = 0 'This is excess water from lower layer theta assimilation
                                        Dim DPlmt As Single = 0 'Deep Perc Limit
                                        '**************************************************************************************
                                        '***  'Finish Infiltration <--- kinda like AquaCrop (Raes et al. Ch. 3 AquaCrop v.6 Manual, 2017, FAO, Rome)
                                        '***  and Drainage Stuff following AquaCrop for decaying method and otherwise based on ASCE MOP 70 2ed, 2016
                                        If DPMethod = DPLimit.Decaying_Limit Then
                                            ArrSWS(0) = 1000 * ArrFC(0) * Zri - DrLast 'Drains yesterday's water
                                            ArrThta(0) = ArrSWS(0) / Zri / 1000
                                            ArrDrain(0) = Limit(ArrTau(0) * (VWCsat(0) - ArrFC(0)) * (Math.Exp(ArrThta(0) - ArrFC(0)) - 1) / (Math.Exp(VWCsat(0) - ArrFC(0)) - 1), 0, ArrTau(0) * (VWCsat(0) - ArrFC(0)))
                                            DPlmt = 1000 * ArrDrain(0) * Zri 'Mynotes 41
                                        ElseIf DPMethod = DPLimit.Constant_Limit Then
                                            DPlmt = DPupper
                                        Else
                                            DPlmt = 999999999
                                        End If
                                        ArrDP(0) = Limit(Fnet - DrLast, 0, DPlmt) 'Jensen and Allen 2016 Eq. 10-12b
                                        DPtot = ArrDP(0) 'My notes 42.1
                                        If Zri < Zrmax Then
                                            '***************Simple bucket in the bottom
                                            '*****Sequentially fill soil layers to FC
                                            For iii = IntM To LyrCnt
                                                ArrSWS(iii) = ArrThta(iii) * ArrDelZL(iii) * 1000
                                                ArrSWS(iii) = ArrSWS(iii) + DPtot
                                                DPtot = Math.Max(ArrSWS(iii) - ArrSWSFC(iii), 0)
                                                ArrSWS(iii) = Math.Min(ArrSWS(iii), ArrSWSFC(iii))
                                            Next
                                        End If
                                        '*** End Drainage by JBB
                                        '*****************************************************************************************
                                        If DPMethod = DPLimit.Decaying_Limit Or DPMethod = DPLimit.Constant_Limit Then
                                            'fexc = Math.Min(DrInt + (VWCsat(0) - (EvapSat + De)), 0) 'Infiltration beyond saturation 
                                            DrLimit = Math.Min(0, ArrSWSFC(0) - ARRSWSsat(0) + (EvapSat + De))
                                        Else
                                            DrLimit = 0
                                        End If

                                        Dim DrInt As Single = Math.Min(DrLast - Fnet + ArrDP(0), TAW) 'cannot drop below WP, hopefully it would not anyway
                                        Dim DrInt2 As Single = Math.Max(DrInt, DrLimit)
                                        '***  'Finish Infiltration <--- kinda like AquaCrop (Raes et al. Ch. 3 AquaCrop v.6 Manual, 2017, FAO, Rome)
                                        Fexc = Math.Max(DrInt2 - DrInt, 0) 'Infiltration beyond field capacity, just the FAO model if no limit is used
                                        FexcOut = Fexc
                                        'Finish Infiltration <--- kinda like AquaCrop (Raes et al. Ch. 3 AquaCrop v.6 Manual, 2017, FAO, Rome)
                                        If Zri < Zrmax Then 'Pass water down
                                            'Fills next layer up to max of saturation (actually field capacity for the time being)
                                            'Dim Fint As Single = Math.Min(ARRSWSsat(IntM) - ArrSWS(IntM), Fexc)'<--if we drain through the lower profile
                                            Dim Fint As Single = Math.Min(ArrSWSFC(IntM) - ArrSWS(IntM), Fexc)
                                            Fexc = Fexc - Fint
                                            ArrSWS(IntM) = ArrSWS(IntM) + Fint
                                            If Fexc > 0 Then
                                                'If still more water, pass it down
                                                For iii = IntM + 1 To LyrCnt
                                                    'Fint = Math.Min(ARRSWSsat(iii) - ArrSWS(iii), Fexc)'<-- If we drain through the lower profile
                                                    Fint = Math.Min(ArrSWSFC(iii) - ArrSWS(iii), Fexc)
                                                    Fexc = Fexc - Fint
                                                    ArrSWS(iii) = ArrSWS(iii) + Fint
                                                Next
                                            End If
                                        End If
                                        'Dr = Limit(DrLast - EffectivePrecipitation - Irrigation - CR + ETcAdjusted + (θFC - θvLPrevious) * Math.Max(Zri - ZriPrevious, 0) * 1000, 0, TAW) ' Depletion from lower layer now in DrLast, FAO 56 EQ 86, No Deep Perc. Term, CR = 0 added by JBB accounts for little bit of soil water deficit in new rootzone during growth
                                        Dr = DrInt2

                                        Dim EnergyBalanceET As Single = 0
                                        Dim ETcAdjustedAssimilated As Single = -1
                                        Dim OldDr As Single = 0 'added by JBB for output
                                        Dim OldKs As Single = 0 'added by JBB for output
                                        Dim OldThtvL As Single = 0 'added by JBB for output
                                        Dim OldDrLast As Single = DrLast
                                        Dim ETWBClr As Single = ETcAdjusted
                                        'WeatherGridIndex.ETDailyActual = GetSameDateImageIndex(RecordDate, WeatherGrid.ETDailyActual) 'Original Code commented out by JBB because it was trying to match a date to a datetime JBB
                                        WeatherGridIndex.ETDailyActual = GetSameDateOnlyImageIndex(RecordDate, WeatherGrid.ETDailyActual) 'Adjusted Code by JBB
                                        If WeatherGridIndex.ETDailyActual > -1 Then EnergyBalanceET = WeatherETDailyActualPixels(WeatherGridIndex.ETDailyActual).GetValue(Col, Row)
                                        AssimExc = 0 'This is excess water from assimilation
                                        If AssimilationMethod <> DataAssimilation.No_Assimilation Then
                                            If WeatherGridIndex.ETDailyActual > -1 Then
                                                'We don't update Ke here even though it is dependent on Ks because it would tend to negate the assimilation under some if not all conditions JBB
                                                If EnergyBalanceET >= ETcAdjusted And Ks = 1 Then 'logic added by JBB
                                                    'No adjustment is made because no water stress is modeled and so adjusting will actually bring it closer to water stress by setting dr = TAW-RAW, this logic added by JBB
                                                    ETcAdjustedAssimilated = calcETAssimilation(ETcAdjusted, EnergyBalanceET, AssimilationMethod, AssimilationWeight) 'We still calculate the assimilated ET though JBB
                                                    ETcAdjOut = ETcAdjustedAssimilated 'overwrites ETcAdjusted for output added by JBB
                                                    'Recalculates Dr based on the new ETc, even though no adjustements were made to the water balance. This deviates from how we ran it in the field experiment in 2016!!!
                                                    OldDr = Dr 'added by JBB to keep track of original DR
                                                    'Dr = Limit(DrLast - EffectivePrecipitation - Irrigation - CR + ETcAdjustedAssimilated + (θFC - θvLPrevious) * Math.Max(Zri - ZriPrevious, 0) * 1000, 0, TAW) 'added byJBB
                                                    'Redo Drainage and infiltration

                                                    Fnet = EffectivePrecipitation + Irrigation + CR - ETcAdjustedAssimilated 'Net infiltration (plus caplilary rise)
                                                    Fexc = 0 'default is no excess if no limit is applied to drainage it will all drain
                                                    DPlmt = 0 'Deep Perc Limit
                                                    '***  'Finish Infiltration <--- kinda like AquaCrop (Raes et al. Ch. 3 AquaCrop v.6 Manual, 2017, FAO, Rome)
                                                    '**************************************************************************************
                                                    '***  'Finish Infiltration <--- kinda like AquaCrop (Raes et al. Ch. 3 AquaCrop v.6 Manual, 2017, FAO, Rome)
                                                    '***  and Drainage Stuff following AquaCrop for decaying method and otherwise based on ASCE MOP 70 2ed, 2016
                                                    If DPMethod = DPLimit.Decaying_Limit Then
                                                        ArrSWS(0) = 1000 * ArrFC(0) * Zri - DrLast 'Drains yesterday's water
                                                        ArrThta(0) = ArrSWS(0) / Zri / 1000
                                                        ArrDrain(0) = Limit(ArrTau(0) * (VWCsat(0) - ArrFC(0)) * (Math.Exp(ArrThta(0) - ArrFC(0)) - 1) / (Math.Exp(VWCsat(0) - ArrFC(0)) - 1), 0, ArrTau(0) * (VWCsat(0) - ArrFC(0)))
                                                        DPlmt = 1000 * ArrDrain(0) * Zri 'Mynotes 41
                                                    ElseIf DPMethod = DPLimit.Constant_Limit Then
                                                        DPlmt = DPupper
                                                    Else
                                                        DPlmt = 999999999
                                                    End If
                                                    ArrDP(0) = Limit(Fnet - DrLast, 0, DPlmt) 'Jensen and Allen 2016 Eq. 10-12b
                                                    DPtot = ArrDP(0) 'My notes 42.1
                                                    If Zri < Zrmax Then
                                                        '***************Simple bucket in the bottom
                                                        '*****Sequentially fill soil layers to FC
                                                        For iii = IntM To LyrCnt
                                                            ArrSWS(iii) = ArrSWS(iii) + DPtot
                                                            DPtot = Math.Max(ArrSWS(iii) - ArrSWSFC(iii), 0)
                                                            ArrSWS(iii) = Math.Min(ArrSWS(iii), ArrSWSFC(iii))
                                                        Next
                                                    End If
                                                    '*** End Drainage by JBB
                                                    '*****************************************************************************************
                                                    If DPMethod = DPLimit.Decaying_Limit Or DPMethod = DPLimit.Constant_Limit Then
                                                        'fexc = Math.Min(DrInt + (VWCsat(0) - (EvapSat + De)), 0) 'Infiltration beyond saturation 
                                                        DrLimit = Math.Min(0, ArrSWSFC(0) - ARRSWSsat(0) + (EvapSat + De))
                                                    Else
                                                        DrLimit = 0
                                                    End If

                                                    DrInt = Math.Min(DrLast - Fnet + ArrDP(0), TAW) 'cannot drop below WP, hopefully it would not anyway
                                                    DrInt2 = Math.Max(DrInt, DrLimit)
                                                    '***  'Finish Infiltration <--- kinda like AquaCrop (Raes et al. Ch. 3 AquaCrop v.6 Manual, 2017, FAO, Rome)
                                                    Fexc = Math.Max(DrInt2 - DrInt, 0) 'Infiltration beyond field capacity, just the FAO model if no limit is used
                                                    FexcOut = Fexc
                                                    'Finish Infiltration <--- kinda like AquaCrop (Raes et al. Ch. 3 AquaCrop v.6 Manual, 2017, FAO, Rome)
                                                    If Zri < Zrmax Then 'Pass water down
                                                        'Fills next layer up to max of saturation (actually field capacity for the time being)
                                                        'Dim Fint As Single = Math.Min(ARRSWSsat(IntM) - ArrSWS(IntM), Fexc)'<--if we drain through the lower profile
                                                        Dim Fint As Single = Math.Min(ArrSWSFC(IntM) - ArrSWS(IntM), Fexc)
                                                        Fexc = Fexc - Fint
                                                        ArrSWS(IntM) = ArrSWS(IntM) + Fint
                                                        If Fexc > 0 Then
                                                            'If still more water, pass it down
                                                            For iii = IntM + 1 To LyrCnt
                                                                'Fint = Math.Min(ARRSWSsat(iii) - ArrSWS(iii), Fexc)'<-- If we drain through the lower profile
                                                                Fint = Math.Min(ArrSWSFC(iii) - ArrSWS(iii), Fexc)
                                                                Fexc = Fexc - Fint
                                                                ArrSWS(iii) = ArrSWS(iii) + Fint
                                                            Next
                                                        End If
                                                    End If
                                                    'Dr = Limit(DrLast - EffectivePrecipitation - Irrigation - CR + ETcAdjusted + (θFC - θvLPrevious) * Math.Max(Zri - ZriPrevious, 0) * 1000, 0, TAW) ' Depletion from lower layer now in DrLast, FAO 56 EQ 86, No Deep Perc. Term, CR = 0 added by JBB accounts for little bit of soil water deficit in new rootzone during growth
                                                    Dr = DrInt2
                                                    'AssimExc = AssimExc + (OldDr - Dr) 'commented out no added water, just a difference in ET
                                                Else
                                                    ETcAdjustedAssimilated = calcETAssimilation(ETcAdjusted, EnergyBalanceET, AssimilationMethod, AssimilationWeight)
                                                    ETcAdjOut = ETcAdjustedAssimilated 'overwrites ETcAdjusted for output added by JBB
                                                    OldKs = Ks 'added by JBB to keep track of original Ks
                                                    Ks = Limit(((ETcAdjustedAssimilated / ReferenceET) - Ke) / Kcb, 0, 1) ' Assumes that errors are in Ks, so back calcs for Ks
                                                    OldDrLast = DrLast
                                                    DrLast = Math.Max(TAW - Ks * (TAW - RAW), 0) 'Back Calcs for Dr,i-1, for "corrected" Ks
                                                    OldDr = Dr 'added by JBB to keep track of original DR
                                                    Kcadj = Ks * Kcb + Ke
                                                    Fnet = EffectivePrecipitation + Irrigation + CR - ETcAdjustedAssimilated 'Net infiltration (plus caplilary rise)
                                                    Fexc = 0 'default is no excess if no limit is applied to drainage it will all drain

                                                    DPlmt = 0 'Deep Perc Limit
                                                    '***  'Finish Infiltration <--- kinda like AquaCrop (Raes et al. Ch. 3 AquaCrop v.6 Manual, 2017, FAO, Rome)
                                                    '**************************************************************************************
                                                    '***  'Finish Infiltration <--- kinda like AquaCrop (Raes et al. Ch. 3 AquaCrop v.6 Manual, 2017, FAO, Rome)
                                                    '***  and Drainage Stuff following AquaCrop for decaying method and otherwise based on ASCE MOP 70 2ed, 2016
                                                    If DPMethod = DPLimit.Decaying_Limit Then
                                                        ArrSWS(0) = 1000 * ArrFC(0) * Zri - DrLast 'Drains yesterday's water
                                                        ArrThta(0) = ArrSWS(0) / Zri / 1000
                                                        ArrDrain(0) = 0
                                                        DPlmt = 0 'Mynotes 41
                                                    ElseIf DPMethod = DPLimit.Constant_Limit Then
                                                        DPlmt = DPupper
                                                    Else
                                                        DPlmt = 999999999
                                                    End If
                                                    ArrDP(0) = 0 'Upon assimilation there will not be any drainage
                                                    DPtot = 0 'My notes 42.1
                                                    '*** End Drainage by JBB
                                                    '*****************************************************************************************
                                                    If DPMethod = DPLimit.Decaying_Limit Or DPMethod = DPLimit.Constant_Limit Then
                                                        'fexc = Math.Min(DrInt + (VWCsat(0) - (EvapSat + De)), 0) 'Infiltration beyond saturation 
                                                        DrLimit = Math.Min(0, ArrSWSFC(0) - ARRSWSsat(0) + (EvapSat + De))
                                                    Else
                                                        DrLimit = 0
                                                    End If

                                                    DrInt = Math.Min(DrLast - Fnet + ArrDP(0), TAW) 'cannot drop below WP, hopefully it would not anyway
                                                    DrInt2 = Math.Max(DrInt, DrLimit)
                                                    '***  'Finish Infiltration <--- kinda like AquaCrop (Raes et al. Ch. 3 AquaCrop v.6 Manual, 2017, FAO, Rome)
                                                    Fexc = Math.Max(DrInt2 - DrInt, 0) 'Infiltration beyond field capacity, just the FAO model if no limit is used
                                                    FexcOut = Fexc
                                                    'Finish Infiltration <--- kinda like AquaCrop (Raes et al. Ch. 3 AquaCrop v.6 Manual, 2017, FAO, Rome)
                                                    If Zri < Zrmax Then 'Pass water down
                                                        'Fills next layer up to max of saturation (actually field capacity for the time being)
                                                        'Dim Fint As Single = Math.Min(ARRSWSsat(IntM) - ArrSWS(IntM), Fexc)'<--if we drain through the lower profile
                                                        Dim Fint As Single = Math.Min(ArrSWSFC(IntM) - ArrSWS(IntM), Fexc)
                                                        Fexc = Fexc - Fint
                                                        ArrSWS(IntM) = ArrSWS(IntM) + Fint
                                                        If Fexc > 0 Then
                                                            'If still more water, pass it down
                                                            For iii = IntM + 1 To LyrCnt
                                                                'Fint = Math.Min(ARRSWSsat(iii) - ArrSWS(iii), Fexc)'<-- If we drain through the lower profile
                                                                Fint = Math.Min(ArrSWSFC(iii) - ArrSWS(iii), Fexc)
                                                                Fexc = Fexc - Fint
                                                                ArrSWS(iii) = ArrSWS(iii) + Fint
                                                            Next
                                                        End If
                                                    End If
                                                    'Dr = Limit(DrLast - EffectivePrecipitation - Irrigation - CR + ETcAdjusted + (θFC - θvLPrevious) * Math.Max(Zri - ZriPrevious, 0) * 1000, 0, TAW) ' Depletion from lower layer now in DrLast, FAO 56 EQ 86, No Deep Perc. Term, CR = 0 added by JBB accounts for little bit of soil water deficit in new rootzone during growth
                                                    Dr = DrInt2

                                                    ''Dr = Limit(DrLast - EffectivePrecipitation - EffectiveIrrigation - CR + ETcAdjustedAssimilated, 0, TAW) 'Recalcualtes Today's Dr,i, Still no Deep Perc <-- But deep perc is built in, CR = 0 commented out by JBB
                                                    'Dr = Limit(DrLast - EffectivePrecipitation - Irrigation - CR + ETcAdjustedAssimilated + (θFC - θvLPrevious) * Math.Max(Zri - ZriPrevious, 0) * 1000, 0, TAW) 'added byJBB
                                                    AssimExc = AssimExc + (OldDrLast - DrLast)
                                                End If
                                                ETWBClr = ETcAdjustedAssimilated
                                                '**************RE Calc Kt and TE and Ke if Te is included, and reassimilate De

                                                If WBTeMethod = WBTeMethod.No_Transpiration Then
                                                    'Does nothing, but using the other two options with an Or statement didn't work for some reason
                                                Else
                                                    Kt = 0
                                                    De = DePrev 'Goes back to previous timestep De
                                                    DREW = DrewPrev 'goes back to previous timestep Drew
                                                    Tdrew = 0
                                                    Dim KtNum As Single = Limit(1 - De / TEW, 0.001, 1) 'ASCE MOP 70, 2ed, 2016, EQ. 9-31 with upper limit because I don't think the equation was expecting drLast to be able to be <0, whcih it will in this version of SETMI the other limit follows MOP70 p. 245
                                                    Dim KtDen As Single = Limit(1 - DrLast / TAW, 0.001, 1) 'ASCE MOP 70, 2ed, 2016, EQ. 9-31 with upper limit because I don't think the equation was expecting drLast to be able to be <0, whcih it will in this version of SETMI the other limit follows MOP70 p. 245
                                                    Kt = Math.Min(1, KtNum / KtDen * ((Ze / Zri) ^ 0.6)) 'ASCE MOP 70, 2ed, 2016 EQ. 9-31 p. 245
                                                    Te = Kt * Ks * Kcb * ReferenceET 'ASCE MOP 70 2ed, 2016, Eq. 9-30
                                                    If WBTeMethod = WBTeMethod.Include_Transp_frm_Skn_Lyr_Also Then 'added by JBB, this allows transpiration from the skin layer to help conserve mass in the case of a small root zone and a kcb>0
                                                        If De < TEW Then
                                                            Tdrew = Math.Min((1 - DREW / REW) / (1 - De / TEW) * Te, Te)
                                                        Else
                                                            Tdrew = 0
                                                        End If
                                                    End If
                                                    If IrrigationMethod = IrrigationMethod.Subsurface Then 'Surface isn't wetted from irrigation
                                                        De = Limit(De - ((1 - Fb) * (EffectivePrecipitation) + Fb * (EffectivePptTomorrow)) + Evap / Few + Te, 0, TEW) 'ASCE MOP 70 ed.2, 2016 eq. 9-28 w/out irrigation term JBB
                                                        DREW = Limit(DREW - ((1 - Fb) * (EffectivePrecipitation) + Fb * (EffectivePptTomorrow)) + Evap / Few + Tdrew, 0, REW) 'ASCE MOP 70 ed.2, 2016 eq. 9-28 w/out irrigation term JBB Eq. 9-29 Also Irrigation 6th Ed., IA, 2011 Eq. 5.56
                                                    Else 'All other forms of irrigation
                                                        De = Limit(De - ((1 - Fb) * (EffectivePrecipitation + EffectiveIrrigation) + Fb * (EffectivePptTomorrow + EffectiveIrrTomorrow)) + Evap / Few + Te, 0, TEW) 'ASCE MOP 70 ed.2, 2016 eq. 9-28
                                                        DREW = Limit(DREW - ((1 - Fb) * (EffectivePrecipitation + EffectiveIrrigation) + Fb * (EffectivePptTomorrow + EffectiveIrrTomorrow)) + Evap / Few + Tdrew, 0, REW) 'ASCE MOP 70 ed.2, 2016 eq. 9-29 Also Irrigation 6th Ed., IA, 2011 Eq. 5.56
                                                    End If
                                                    If EvaporatedDepthFlag = True And AssimilationMethod <> DataAssimilation.No_Assimilation Then
                                                        OldDe = De 'added by JBB to preserve old value
                                                        De = Limit(calcETAssimilation(OldDe, ActualEvaporatedDepth, AssimilationMethod, EvaporatedDepthWeight), 0, TEW) 'added by JBB, The limit was included for consistency, but really the soil should be allowed to have negative depletion for 3 days or something. JBB
                                                    End If
                                                    If SkinEvapDepthFlag = True And AssimilationMethod <> DataAssimilation.No_Assimilation Then
                                                        OldDREW = DREW 'added by JBB to preserve old value
                                                        DREW = Limit(calcETAssimilation(OldDREW, ActualSkinEvapDepth, AssimilationMethod, SkinEvapWeight), 0, REW) 'added by JBB, The limit was included for consistency, but really the soil should be allowed to have negative depletion for 3 days or something. JBB
                                                    End If

                                                End If
                                            End If
                                        End If
                                        SumETcAdjOut = SumETcAdjOut + ETcAdjOut 'sum ETc for output added by JBB
                                        If DoY >= DOYini And DoY <= DOYMaxBM Then SumKcOut = SumKcOut + Kcb * Ks * Kst 'Campos et al. 2018
                                        '***Added by JBB to incorporate actual depletion input***
                                        'Dim ActualDepletion As single = 0
                                        'WeatherGridIndex.Depletion = GetSameDateOnlyImageIndex(RecordDate, WeatherGrid.Depletion) 'Adjusted Code by JBB
                                        'If WeatherGridIndex.Depletion > -1 Then ActualDepletion = WeatherDepletionDailyPixels(WeatherGridIndex.Depletion).GetValue(Col, Row)
                                        'If WeatherGridIndex.Depletion > -1 Then
                                        If DPMethod = DPLimit.Decaying_Limit Or DPMethod = DPLimit.Constant_Limit Then
                                            'fexc = Math.Min(DrInt + (VWCsat(0) - (EvapSat + De)), 0) 'Infiltration beyond saturation 
                                            DrLimit = Math.Min(0, ArrSWSFC(0) - ARRSWSsat(0) + (EvapSat + De))
                                        Else
                                            DrLimit = 0
                                        End If
                                        If DepletionFlag = True And AssimilationMethod <> DataAssimilation.No_Assimilation Then
                                            OldDr = Dr 'added by JBB to preserve old value
                                            Dr = Limit(calcETAssimilation(OldDr, ActualDepletion, AssimilationMethod, DepletionWeight), DrLimit, TAW) 'added by JBB, The limit was included for consistency, but really the soil should be allowed to have negative depletion for 3 days or something. JBB
                                            AssimExc = AssimExc + (OldDr - Dr)
                                        End If
                                        'If ETcAdjustedAssimilated >= 0 Then 'including ETcassim here also deviates from how we ran it in the 2016 field experiment.
                                        '    DP = Dr - Limit(DrLast - EffectivePrecipitation - Irrigation - CR + ETcAdjustedAssimilated + (θFC - θvLPrevious) * Math.Max(Zri - ZriPrevious, 0) * 1000, -9999, TAW) 'ADDED by JBB as estimate of deep percolation as water above field capacity that was chopped off by DR calc JBB
                                        'Else
                                        '    DP = Dr - Limit(DrLast - EffectivePrecipitation - Irrigation - CR + ETcAdjusted + (θFC - θvLPrevious) * Math.Max(Zri - ZriPrevious, 0) * 1000, -9999, TAW) 'ADDED by JBB as estimate of deep percolation as water above field capacity that was chopped off by DR calc JBB
                                        'End If
                                        DP = ArrDP(0)
                                        DPL = DPtot
                                        WBClosureRZ = OldDrLast - Dr + ETWBClr + ROOut + DP - Precipitation - Irrigation + FexcOut - AssimExc




                                        DrLastOut = DrLast
                                        DrLast = Dr 'Dr,i-1 = Dr,i
                                        TAWLower = 0
                                        ThtvL = 0
                                        ZL = Zrmax - Zri
                                        If Math.Abs(Zri - Zrmax) < 0.00001 Then
                                            For iii = IntM To LyrCnt
                                                ArrThta(iii) = ArrFC(iii)
                                            Next
                                            ThtvL = ThtvL + ArrFC(LyrCnt)
                                            'θvL = θFC 'in this case there is a zero thickness of the lower layer, the value doesn't matter, but we need to avoid a divide by zero
                                        Else
                                            For iii = IntM To LyrCnt
                                                If ArrDelZL(iii) > 0 Then
                                                    ArrThta(iii) = Math.Max(ArrSWS(iii) / ArrDelZL(iii) / 1000, ArrWP(iii)) 'Doesn't include provision for CR anymore
                                                Else
                                                    ArrThta(iii) = ArrFC(iii)
                                                End If
                                                ThtvL = ThtvL + ArrThta(iii) * ArrDelZL(iii) / ZL
                                                TAWLower = TAWLower + (ArrFC(iii) - ArrWP(iii)) * ArrDelZL(iii) * 1000 'TAW in lower layer
                                            Next
                                            'θvL = Math.Max(((DP + CRL - CR) + θvLPrevious * (Zrmax - Zri) * 1000) / (Zrmax - Zri) / 1000, θWP) 'soil moisture content of lower soil layer added by JBB the lower limit is a formality, the only way to extract water from the lower layer is for the upper layer to grow into at which point is then the upper layer
                                        End If


                                        If ThetaVFlag = True And AssimilationMethod <> DataAssimilation.No_Assimilation Then 'Thetav is assimilated as one image (me being lazy) but weighted out by AWC and soil water storage
                                            'OldθvL = θvL 'added by JBB to preserve old value
                                            OldThtvL = ThtvL
                                            WPL = 0
                                            If ZL >= 0.00001 Then
                                                For iii = IntM To LyrCnt
                                                    ArrOldThta(iii) = ArrThta(iii)
                                                    WPL = WPL + ArrWP(iii) * ArrDelZL(iii) / ZL 'WP in lower layer
                                                Next
                                                For iii = IntM To LyrCnt
                                                    If ArrDelZL(iii) > 0 Then
                                                        ArrThtaWt(iii) = ((ArrFC(iii) - ArrWP(iii)) * ArrDelZL(iii) * 1000 / TAWLower) * (ActualThetaV - WPL) * (ZL / ArrDelZL(iii)) + ArrWP(iii)
                                                        ArrThta(iii) = Limit(calcETAssimilation(ArrOldThta(iii), ArrThtaWt(iii), AssimilationMethod, ThetaVWeight), ArrWP(iii), ArrFC(iii)) 'added by JBB limited to FC because we instantly drain the lower layer
                                                    Else
                                                        ArrThta(iii) = ArrFC(iii)
                                                    End If
                                                Next
                                                'θvL = Math.Max(calcETAssimilation(OldθvL, ActualThetaV, AssimilationMethod, ThetaVWeight), θWP) 'added by JBB
                                                ThtvL = 0
                                                For iii = IntM To LyrCnt
                                                    ThtvL = ThtvL + ArrThta(iii) * ArrDelZL(iii) / (ZL)
                                                Next
                                                AssimExcL = (ThtvL - OldThtvL) * ZL * 1000
                                            End If
                                        End If

                                        'DPL = Math.Max((θvL - θFC) * (Zrmax - Zri) * 1000, 0) 'deep percolation from the lower layer added by JBB
                                        'DPL = DPtot
                                        'θvL = Math.Min(θvL, θFC) 'now that DPL has been calculated limit θvL to θFC added by JBB
                                        ' θvLPrevious = θvL 'added by JBB
                                        ' ProfileAvgθv = (θvL * (Zrmax - Zri) / Zrmax) + (θFC - Dr / 1000 / Zri) * Zri / Zrmax 'm/m added by JBB this is useful as an initial state variable for wrapping a water balance across years
                                        ProfileAvgθv = (ArrFC(0) - Dr / 1000 / Zri) * Zri / Zrmax 'm/m added by JBB this is useful as an initial state variable for wrapping a water balance across years
                                        For iii = IntM To LyrCnt
                                            ProfileAvgθv = ProfileAvgθv + ArrThta(iii) * ArrDelZL(iii) / Zrmax
                                        Next
                                        WBClosureFull = (ProfileAvgThtavPrev - ProfileAvgθv) * Zrmax * 1000 - ETWBClr - ROOut - DPL + Precipitation + Irrigation - Fexc + AssimExc + AssimExcL
                                        ProfileAvgThtavPrev = ProfileAvgθv
                                        AvailWaterAboveMAD = MAD / 100 * TAW - Dr 'Available water above MAD, will be negative if dr>MAD added by JBB
                                        RequiredGrossIrrigation = Math.Min(Math.Max(TargetDepthAboveMAD - AvailWaterAboveMAD, 0), Dr) / (AppEfficiency / 100) 'Produce spatial maps of irrigation requirement added by JBB
                                        For M = 0 To MultispectralDate.Count - 1
                                            'If MultispectralDate(M) = RecordDate Then 'Original Code Commented out by JBB
                                            Dim MultispectralDateOnly As DateTime = MultispectralDate(M).Date '.AddDays(1) 'For some reason Record Date was one day ahead of where it should be, I changed that, but if the original code is used a day must be added here; Gets the year, month, day only part of the date added by JBB
                                            If MultispectralDateOnly = RecordDate Then 'Added by JBB because the date & time will never equal the date only JBB
                                                'OutputImages(M) = New WaterBalance_ImageOverpassOutput(RecordDate, Ke, Ks, calcSeasonInterpolation(DoY, KcbIni, KcbMid, KcbEnd, DateIni, DateDev, DateMid, DateLate, DateEnd), Kcb, Zri, θFC, TAW - Dr, ReferenceET, ETcAdjusted, EnergyBalanceET, ETcAdjustedAssimilated)'commented out by JBB
                                                OutputImages(M) = New WaterBalance_ImageOverpassOutput(RecordDate, Ke, OldKs, Ks, Kcb, Kcbs(M), Zri, ArrFC(0), ArrWP(0), ArrThtaIni(0), TAW - Dr, OldDr, Dr, ReferenceET, ETcAdjusted, EnergyBalanceET, ETcAdjustedAssimilated) 'added by JBB to include different output
                                            End If
                                        Next
                                        '*****Added by JBB to output depletion map assimilation and others
                                        For M = 0 To DepletionCount - 1
                                            Dim DepletionDateOnly As DateTime = DepletionDate(M).Date '.AddDays(1) 'For some reason Record Date was one day ahead of where it should be, I changed that, but if the original code is used a day must be added here; Gets the year, month, day only part of the date added by JBB
                                            If DepletionDateOnly = RecordDate Then 'Added by JBB because the date & time will never equal the date only JBB
                                                OutputDepletions(M) = New WaterBalance_InputDepletionOutput(RecordDate, OldDr, ActualDepletion, Dr) 'added by JBB to include different output
                                            End If
                                        Next
                                        For M = 0 To ThetaVCount - 1
                                            Dim ThetaVDateOnly As DateTime = ThetaVDate(M).Date '.AddDays(1) 'For some reason Record Date was one day ahead of where it should be, I changed that, but if the original code is used a day must be added here; Gets the year, month, day only part of the date added by JBB
                                            If ThetaVDateOnly = RecordDate Then 'Added by JBB because the date & time will never equal the date only JBB
                                                OutputThetaVs(M) = New WaterBalance_InputThetaVOutput(RecordDate, OldThtvL, ActualThetaV, ThtvL) 'added by JBB to include different output
                                            End If
                                        Next
                                        For M = 0 To EvaporatedDepthCount - 1
                                            Dim EvaporatedDepthDateOnly As DateTime = EvaporatedDepthDate(M).Date '.AddDays(1) 'For some reason Record Date was one day ahead of where it should be, I changed that, but if the original code is used a day must be added here; Gets the year, month, day only part of the date added by JBB
                                            If EvaporatedDepthDateOnly = RecordDate Then 'Added by JBB because the date & time will never equal the date only JBB
                                                OutputEvaporatedDepths(M) = New WaterBalance_InputEvaporatedDepthOutput(RecordDate, OldDe, ActualEvaporatedDepth, De) 'added by JBB to include different output
                                            End If
                                        Next

                                        For M = 0 To SkinEvapDepthCount - 1 'Added by JBB Copying Code from above 2/9/2018
                                            Dim SkinEvapDepthDateOnly As DateTime = SkinEvapDepthDate(M).Date '.AddDays(1) 'For some reason Record Date was one day ahead of where it should be, I changed that, but if the original code is used a day must be added here; Gets the year, month, day only part of the date added by JBB
                                            If SkinEvapDepthDateOnly = RecordDate Then 'Added by JBB because the date & time will never equal the date only JBB
                                                OutputSkinEvapDepths(M) = New WaterBalance_InputSkinEvapDepthOutput(RecordDate, OldDREW, ActualSkinEvapDepth, DREW) 'added by JBB to include different output modified 3/23/2018
                                            End If
                                        Next
                                        If FlgOutLyr = False Then
                                            For m = 1 To LyrCnt
                                                OutputLayerInfo(m - 1) = New WaterBalance_LayerOutput(m, ArrThick(m), ArrFC(m), ArrWP(m), ArrThtaIni(m), VWCsat(m), Ksat(m))
                                            Next
                                            FlgOutLyr = True
                                        End If

                                        '*****Added to remember the previous time stamp
                                        ZriPrevious = Zri 'added by JBB
                                        For iii = 1 To 3
                                            ArrThtaPrevOut(iii) = ArrThtaPrev(iii)
                                            ArrDelZLprev(iii) = ArrDelZL(iii)
                                            ArrThtaPrev(iii) = ArrThta(iii)
                                        Next
                                        '****Added for termination
                                        FCrPrev = ArrFC(0)
                                        WPrPrev = ArrWP(0)
                                        '****end added by JBB for termination
                                        '****end added by JBB
                                        SumDPOut = SumDPOut + DP
                                        SumDPLOut = SumDPLOut + DPL
                                        SumFexcOut = SumFexcOut + FexcOut
                                        SumFexc = SumFexc + Fexc
                                        SumAssimExcOut = SumAssimExcOut + AssimExc
                                        SumAssimExcLOut = SumAssimExcLOut + AssimExcL
                                        SumETrOut = SumETrOut + ReferenceET
                                        SumAddFromDelZr = SumAddFromDelZr + DelDrlastFromDelZr

                                        CumWBClosureRZ = DrIni - Dr + SumAddFromDelZr + SumETcAdjOut + SumROOut + SumDPOut - SumPcpOut - SumInetOut + SumFexcOut - SumAssimExcOut
                                        CumWBClosureFull = (ProfileThetaIni - ProfileAvgθv) * 1000 * Zrmax - SumETcAdjOut - SumROOut - SumDPLOut + SumPcpOut + SumInetOut - SumFexc + SumAssimExcOut + SumAssimExcLOut

                                        Dim StrSeason As String = "Offseason"
                                        If DoY >= DOYini And DoY < DOYdev Then
                                            StrSeason = "Initial"
                                        ElseIf DoY >= DOYdev And DoY < DOYefc Then
                                            StrSeason = "Development"
                                        ElseIf DoY >= DOYefc And DoY < DOYlate Then
                                            StrSeason = "Full Cover"
                                        ElseIf DoY >= DOYlate And DoY <= DOYterm Then
                                            StrSeason = "Late"
                                        End If


                                        'If AssimilationMethod <> DataAssimilation.No_Assimilation Then If WeatherGridIndex.ETDailyActual > -1 Then ETcAdjusted = ETcAdjustedAssimilated 'This allows unassimilated ETc to be output for the ET image outputs in the text file, but then assimilated ETc later on JBB moved up here by JBB
                                        '******************START JBB ADDED CODE******** Much was added, copying code already present in SETMI
                                        For M = 0 To WBOutDateCount - 1 'Loop Through Selected Output Dates
                                            If WBOutDatesARR(M) = RecordDate Then 'Checks if it is an output date JBB
                                                Dim CountOut As Integer = 0
                                                If OutputImagesBoxWater.GetItemChecked(0) Then OutPixelsDbg(CountOut, M).SetValue(Clean(AvailWaterAboveMAD), {Col, Row}) : CountOut += 1
                                                If OutputImagesBoxWater.GetItemChecked(1) Then OutPixelsDbg(CountOut, M).SetValue(Clean(Kcb), {Col, Row}) : CountOut += 1

                                                If OutputImagesBoxWater.GetItemChecked(2) Then OutPixelsDbg(CountOut, M).SetValue(Clean(De), {Col, Row}) : CountOut += 1
                                                If OutputImagesBoxWater.GetItemChecked(3) Then OutPixelsDbg(CountOut, M).SetValue(Clean(DREW), {Col, Row}) : CountOut += 1

                                                If OutputImagesBoxWater.GetItemChecked(4) Then OutPixelsDbg(CountOut, M).SetValue(Clean(DOYefc), {Col, Row}) : CountOut += 1
                                                If OutputImagesBoxWater.GetItemChecked(5) Then OutPixelsDbg(CountOut, M).SetValue(Clean(Ke), {Col, Row}) : CountOut += 1

                                                If OutputImagesBoxWater.GetItemChecked(6) Then OutPixelsDbg(CountOut, M).SetValue(Clean(Fexc), {Col, Row}) : CountOut += 1

                                                If OutputImagesBoxWater.GetItemChecked(7) Then OutPixelsDbg(CountOut, M).SetValue(Clean(Fc), {Col, Row}) : CountOut += 1
                                                If OutputImagesBoxWater.GetItemChecked(8) Then OutPixelsDbg(CountOut, M).SetValue(Clean(DOYini), {Col, Row}) : CountOut += 1

                                                If OutputImagesBoxWater.GetItemChecked(9) Then OutPixelsDbg(CountOut, M).SetValue(Clean(SumKcOut), {Col, Row}) : CountOut += 1
                                                If OutputImagesBoxWater.GetItemChecked(10) Then OutPixelsDbg(CountOut, M).SetValue(Clean(Irrigation), {Col, Row}) : CountOut += 1

                                                If OutputImagesBoxWater.GetItemChecked(11) Then OutPixelsDbg(CountOut, M).SetValue(Clean(RequiredGrossIrrigation), {Col, Row}) : CountOut += 1
                                                If OutputImagesBoxWater.GetItemChecked(12) Then OutPixelsDbg(CountOut, M).SetValue(Clean(DP), {Col, Row}) : CountOut += 1

                                                If OutputImagesBoxWater.GetItemChecked(13) Then OutPixelsDbg(CountOut, M).SetValue(Clean(ZriOut), {Col, Row}) : CountOut += 1
                                                If OutputImagesBoxWater.GetItemChecked(14) Then OutPixelsDbg(CountOut, M).SetValue(Clean(ROOut), {Col, Row}) : CountOut += 1

                                                If OutputImagesBoxWater.GetItemChecked(15) Then OutPixelsDbg(CountOut, M).SetValue(Clean(Dr), {Col, Row}) : CountOut += 1

                                                If OutputImagesBoxWater.GetItemChecked(16) Then OutPixelsDbg(CountOut, M).SetValue(Clean(DPL), {Col, Row}) : CountOut += 1

                                                If OutputImagesBoxWater.GetItemChecked(17) Then OutPixelsDbg(CountOut, M).SetValue(Clean(DOYterm), {Col, Row}) : CountOut += 1
                                                If OutputImagesBoxWater.GetItemChecked(18) Then OutPixelsDbg(CountOut, M).SetValue(Clean(TAW), {Col, Row}) : CountOut += 1

                                                If OutputImagesBoxWater.GetItemChecked(19) Then OutPixelsDbg(CountOut, M).SetValue(Clean(SumFexcOut), {Col, Row}) : CountOut += 1
                                                If OutputImagesBoxWater.GetItemChecked(20) Then OutPixelsDbg(CountOut, M).SetValue(Clean(SumInetOut), {Col, Row}) : CountOut += 1
                                                If OutputImagesBoxWater.GetItemChecked(21) Then OutPixelsDbg(CountOut, M).SetValue(Clean(SumPcpOut), {Col, Row}) : CountOut += 1

                                                If OutputImagesBoxWater.GetItemChecked(22) Then OutPixelsDbg(CountOut, M).SetValue(Clean(SumDPOut), {Col, Row}) : CountOut += 1

                                                If OutputImagesBoxWater.GetItemChecked(23) Then OutPixelsDbg(CountOut, M).SetValue(Clean(SumROOut), {Col, Row}) : CountOut += 1

                                                If OutputImagesBoxWater.GetItemChecked(24) Then OutPixelsDbg(CountOut, M).SetValue(Clean(SumDPLOut), {Col, Row}) : CountOut += 1

                                                If OutputImagesBoxWater.GetItemChecked(25) Then OutPixelsDbg(CountOut, M).SetValue(Clean(SumETcAdjOut), {Col, Row}) : CountOut += 1

                                                If OutputImagesBoxWater.GetItemChecked(26) Then OutPixelsDbg(CountOut, M).SetValue(Clean(ETc), {Col, Row}) : CountOut += 1

                                                If OutputImagesBoxWater.GetItemChecked(27) Then OutPixelsDbg(CountOut, M).SetValue(Clean(ThtvL * 1000), {Col, Row}) : CountOut += 1
                                                If OutputImagesBoxWater.GetItemChecked(28) Then OutPixelsDbg(CountOut, M).SetValue(Clean(ProfileAvgθv * 1000), {Col, Row}) : CountOut += 1
                                                If OutputImagesBoxWater.GetItemChecked(29) Then OutPixelsDbg(CountOut, M).SetValue(Clean(Ks), {Col, Row}) : CountOut += 1
                                                If OutputImagesBoxWater.GetItemChecked(30) Then OutPixelsDbg(CountOut, M).SetValue(Clean(ETcAdjOut), {Col, Row}) : CountOut += 1
                                            End If
                                        Next
                                        '******************End JBB ADDED CODE********

                                        ' If AssimilationMethod <> DataAssimilation.No_Assimilation Then If WeatherGridIndex.ETDailyActual > -1 Then ETcAdjusted = ETcAdjustedAssimilated 'This allows unassimilated ETc to be output for the ET image outputs in the text file, but then assimilated ETc later on JBB commented out and moved above a few lines by JBB

                                        ' OutputSeason(RecordDate.DayOfYear - 1) = New WaterBalance_SeasonOutput(RecordDate, ReferenceET, ETcAdjusted, Dr, Kcb, Ke, Ks, Zri) Original Code
                                        'OutputSeason(RecordDate.DayOfYear - 1) = New WaterBalance_SeasonOutput(RecordDate, ReferenceET, Fc, Kcb, Ke, Ks, ETcAdjusted, EffectivePrecipitation, Irrigation, DP, CR, TAW, RAW, TEW, REW, Zri, Dr, MAD, RequiredGrossIrrigation, AvailWaterAboveMAD, Zrmax, θvL, DPL, CRL, ProfileAvgθv, P, Kr, De, Fw, Few, GDD(DoY)) 'Adjusted Code by JBB, commented out by JBB to run for two years
                                        'added by JBB to run for two years
                                        Dim DOYOut As Integer
                                        DOYOut = RecordDate.DayOfYear
                                        If RecordDate.Year = YearStartWB + 1 Then
                                            Dim EndFirstYr = DateSerial(YearStartWB, 12, 31)
                                            DOYOut = DOYOut + EndFirstYr.DayOfYear
                                        End If
                                        'OutputSeason(DOYOut - 1) = New WaterBalance_SeasonOutput(RecordDate, ReferenceET, Fc, Kcb, Ke, Ks, ETcAdjusted, EffectivePrecipitation, Irrigation, DP, CR, TAW, RAW, TEW, REW, Zri, Dr, MAD, RequiredGrossIrrigation, AvailWaterAboveMAD, Zrmax, θvL, DPL, CRL, ProfileAvgθv, P, Kr, De, Fw, Few, GDD(DoY)) 'Adjusted Code by JBB
                                        'OutputSeason(DOYOut - 1) = New WaterBalance_SeasonOutput(RecordDate, ReferenceET, Fc, Kcb, Ke, Ks, ETcAdjusted, EffectivePrecipitation, Irrigation, DP, CR, TAW, RAW, TEW, REW, Zri, Dr, MAD, RequiredGrossIrrigation, AvailWaterAboveMAD, Zrmax, ArrFC(0), DPL, -999, ProfileAvgθv, P, Kr, De, Fw, Few, GDD(DoY)) 'Adjusted Code by JBB
                                        OutputSeason(DOYOut - 1) = New WaterBalance_SeasonOutput(RecordDate, CoverName, YearStartWB, DOYOut, ReferenceET, Precipitation, Igross, Tavg, GDD(DOYOut), DailySAVI(DOYOut), DailyKcb(DOYOut), StrSeason, HcPeriod, U2Limited, RHminLimited, HcTab, Fc, Zri, ArrDelZr(1), ArrDelZr(2), ArrDelZr(3), ArrThtaIni(0), ArrWP(0), ArrFC(0), VWCsat(0), Ksat(0), TAW, Irrigation, CN, S_SCS * 25.4, InitialAbstractions * 25.4, ROOut, EffectivePrecipitation, KcMax, Fw, Few, TEW, REW, Fstage1, Kt, Te, Tdrew, Kr, De, DREW, Ke_ns, Ke_strss, P, RAW, DelDrlastFromDelZr, DrLastOut, CR, ArrThta(0), ArrTau(0), DP, Fexc, DrLimit, Dr, Ks, Kcadj, ETc, ETcAdjOut, EvapNS, Evap, ETc - Evap, ETcAdjOut - Evap, ArrDelZL(1), ArrDelZL(2), ArrDelZL(3), ZL, TAWLower, WPrPrev, FCrPrev, ArrThtaPrevOut(1), ArrThtaPrevOut(2), ArrThtaPrevOut(3), DPtot, ThtvL, ArrThta(1), ArrThta(2), ArrThta(3), ProfileAvgθv, AssimExc, AssimExcL, WBClosureFull, WBClosureRZ, MAD / 100 * TAW, RequiredGrossIrrigation, AvailWaterAboveMAD, Srel, Kst, SumKcOut, SumETrOut, SumPcpOut, SumROOut, SumInetOut, SumETcAdjOut, SumDPOut, SumDPLOut, SumFexcOut, SumAddFromDelZr, SumAssimExcOut, SumAssimExcLOut, CumWBClosureFull, CumWBClosureRZ)
                                        'OutputSeason(DOYOut - 1) = New WaterBalance_SeasonOutput(RecordDate, CoverName, YearStartWB, DOYOut, ReferenceET, Precipitation, Igross, Tavg, GDD(DOYOut), SAVIs(DOYOut), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
                                        'OutputSeason(DOYOut - 1) = New WaterBalance_SeasonOutput(RecordDate, "C", 0, 0, 0, 0, 0, 0, 0, 0, 0, "S", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
                                        'end added by JBB
                                        '***Start JBB Code If the weather data is input properly this added code may not be necessary!!!!!!!!!*****
                                    Else 'Without this added code OutputSeason has a bunch of nothing lines that crash the print out code JBB
                                        Dim RecordDate = BaseDate.AddDays(DoY) 'Without this added code OutputSeason has a bunch of nothing lines that crash the print out code JBB
                                        'OutputSeason(DoY - 1) = New WaterBalance_SeasonOutput(RecordDate, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0) 'Without this added code OutputSeason has a bunch of nothing lines that crash the print out code JBB
                                        OutputSeason(DoY - 1) = New WaterBalance_SeasonOutput(RecordDate, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0) 'Added by JBB 3/12/2018 copying code above

                                        '**** End JBB Code
                                    End If '
                                Next 'End of loop through the growing season JBB

                                If PixelIndex > -1 Then 'This if statement was moved by JBB to allow for single pixel and full image output JBB
                                    Using CSV As New IO.StreamWriter(DebugPath & "\Pixel Lat-" & PixelLatitude & " Long-" & PixelLongitude & ".csv") 'Prints out water balance time series in a .csv file. balance for image dates prints on the next day JBB
                                        CSV.WriteLine("X,Y,Cover,Start WB Date,End WB Date,Inititiation Date,Devel. Date,Eff. Full Cover Date,Late Date,Termination Date")
                                        'CSV.WriteLine(PixelLongitude & "," & PixelLatitude & "," & CoverName & "," & DateSerial(MultispectralDate(0).Year - 1, 12, 31).AddDays(DOYini) & "," & DateSerial(MultispectralDate(0).Year - 1, 12, 31).AddDays(DOYefc) & "," & DateSerial(MultispectralDate(0).Year - 1, 12, 31).AddDays(DOYterm))
                                        CSV.WriteLine(PixelLongitude & "," & PixelLatitude & "," & CoverName & "," & DateSerial(YearStartWB - 1, 12, 31).AddDays(DOYStartWB) & "," & DateSerial(YearStartWB - 1, 12, 31).AddDays(DOYEndWB) & "," & DateSerial(YearStartWB - 1, 12, 31).AddDays(DOYini) & ", " & DateSerial(YearStartWB - 1, 12, 31).AddDays(DOYdev) & "," & DateSerial(YearStartWB - 1, 12, 31).AddDays(DOYefc) & ", " & DateSerial(YearStartWB - 1, 12, 31).AddDays(DOYlate) & "," & DateSerial(YearStartWB - 1, 12, 31).AddDays(DOYterm))
                                        CSV.WriteLine()
                                        'CSV.WriteLine("Date, Ke, Ks, Kcb(FAO), Kcb(VI), Root Depth, Field Capacity, Soil Moisture, ETo(mm), ET(Kcb)(mm), ET(EB)(mm), ET(As)(mm)")commented out by JBB
                                        CSV.WriteLine("ImageTimestamp, Date, Ke, Ks_Old, Ks, Kcb, Kcb_VI, Root Depth (m), Field Capacity (m/m), Wilting Point (m/m), Initial Thetav (m/m), Soil Moisture, Dr(Old) (mm),Dr(As) (mm), ETo (mm), ET(Kc) (mm), ET(EB) (mm), ET(As) (mm)") 'added by JBB for more output

                                        For Each Line In OutputImages '
                                            CSV.WriteLine(Line.WriteDelimited)
                                        Next
                                        CSV.WriteLine()
                                        '***added by JBB
                                        CSV.WriteLine("Date,CGDD,Kcb_VI_Forecasted,Stage 2 Max VI") 'Copied from other code around here
                                        For Each Line In OutputForecasts 'Copied from other code around here
                                            CSV.WriteLine(Line.WriteDelimited) 'Copied from other code around here
                                        Next 'Copied from other code around here
                                        CSV.WriteLine()
                                        'added by JBB copying other code 3/21/2108
                                        CSV.WriteLine("Layer No.,Thickness (m),FC (m/m),WP(m/m),ThtaIni(m/m),ThtaSat(m/m),Ksat(mm/d)")
                                        For Each line In OutputLayerInfo
                                            CSV.WriteLine(line.WriteDelimited)
                                        Next
                                        CSV.WriteLine()
                                        'end added by JBB
                                        CSV.WriteLine("Date,Dr(Old) (mm),Dr(Act) (mm),Dr(As) (mm)")
                                        For Each Line In OutputDepletions
                                            CSV.WriteLine(Line.WriteDelimited)
                                        Next
                                        CSV.WriteLine()
                                        CSV.WriteLine("Date,ThetaVL(Old) (m/m),ThetaVL(Act) (m/m),ThetaVL(As) (m/m)")
                                        For Each Line In OutputThetaVs
                                            CSV.WriteLine(Line.WriteDelimited)
                                        Next
                                        CSV.WriteLine()
                                        CSV.WriteLine("Date,De(Old) (mm),De(Act) (mm),De(As) (mm)")
                                        For Each Line In OutputEvaporatedDepths
                                            CSV.WriteLine(Line.WriteDelimited)
                                        Next
                                        CSV.WriteLine()
                                        CSV.WriteLine("Date,Drew(Old) (mm),Drew(Act) (mm),Drew(As) (mm)") 'Added by JBB copying code from above 2/9/2018
                                        For Each Line In OutputSkinEvapDepths
                                            CSV.WriteLine(Line.WriteDelimited)
                                        Next
                                        CSV.WriteLine()
                                        '***end added by JBB
                                        'CSV.WriteLine("Date,ETo,ETa,Depletion,Kcb,Ke,Ks,Root Depth")'commented out by JBB
                                        'CSV.WriteLine("Date,ETo (mm),Fc,Kcb,Ke,Ks,Kc,ETc (mm),ETc-Adj (mm),Eff. Pcp. (mm),Net Irr. (mm),DP To Lwr Lyr (mm),CR frm Lwr Lyr (mm),TAW (mm),RAW (mm),TEW (mm),REW (mm),Root Depth (m),Root Zone Depletion (mm),MAD (%),Req. Gross Irr. (mm),AW Abv. MAD (mm),Soil Bottom Depth (m),Thetav In Lower Layer (m/m),DP frm Lwr Lyr (mm),CR To Lwr Lyr (mm),Profile Avg Thetav (m/m),P,Kr,Depth Evapd (mm),Fw,Few,CGDD") 'added by JBB for increased output
                                        CSV.WriteLine("Date,Cover,WB Year,DOY,ETref (mm/d),Ppt (mm/d),Igrss (mm/d),Tavg (C),CGDD (C-d),SAVI (-),Kcb (-),Season,Season Hc (m),Season U2m (m/s),Season RHmin (%),Hc (m),Fc (-),Zr (m),delZr1 (m),delZr2 (m),delZr3 (m),Thtarini (m3/m3),WPr (m3/m3),FCr (m3/m3),thetavsatr (m3/m3),Ksatr (mm/d),TAW (mm),In (mm/d),CN,S (mm),Ia (mm),RO (mm/d),Peff (mm/d),Kcmax (-),Fw (-),Few (-),TEW (mm),REW (mm),Fstg1 (-),Kt (-),Te (mm/d),TDREW (mm/d),Kr (-),De (mm),DREW (mm),Ke_ns (-),ke_strss (-),p (-),RAW (mm),+DrStart from delZr (mm/d),Drstart (mm),CR (mm/d),thetar (m3/m3),Taur (-),DP (mm/d),Excs Wtr (mm),Dr Limit (mm),Dr (mm),Ks (-),Kc (-),Etc_ns (mm),Etc_strss (mm),E_ns (mm), E_strss (mm),Tc_ns (mm),Tc_strss (mm),delZL1 (m),delZL2 (m),delZL2 (m),ZL (m),Lwr TAW (mm),WPrPrev (m3/m3),FCrPrev (m3/m3),thtavLast1 (m3/m3),thtavLast2 (m3/m3),thtavLast3 (m3/m3),DPL (mm/d),thtav Lower (m3/m3),thtav1 (m3/m3),thtav2 (m3/m3),thtav3 (m3/m3),Prf. Avg. thtav (m3/m3),Assim. Excess Dr (mm),Assim. Excess ThtvL (mm),WB Clsr (mm),RZ WB Clsr (mm),MAD (mm),Grs Irr Req (mm),Avail Wtr Abv MAD (mm),Sr (-),Ktmp (-),CKcbKsKtmp (-),CETref (mm),CPpt (mm),CRO (mm),CIn (mm),CETc_strss (mm),CDP (mm),CDPL (mm),CExcs (mm),Cadd from delZr (mm),Sum Assim. Excess Dr (mm),Sum Assim. Excess ThtvL (mm),CWB Clsr (mm),CRZWB Clsr (mm)")
                                        For Each Line In OutputSeason
                                            CSV.WriteLine(Line.WriteDelimited)
                                        Next
                                    End Using
                                End If
                                'End If 'End of if the selected pixel has been found JBB'<----Uncomment for single pixel 

                            End If 'End of if the cover has been selected for calculation JBB
                        End If 'End of if checking if there is a cover pixle JBB
                    Else
                        '***************Start JBB Added Code, Previously there was nothing in this else statement
                        'for multiple debug images
                        For N = 0 To OutputRasterNames.Count - 1
                            For M = 0 To WBOutDateCount - 1 'MultispectralImages.Count - 1
                                OutPixelsDbg(N, M).SetValue(Single.MinValue, {Col, Row}) 'sets no data value for non intersecting pixels
                            Next
                        Next
                        '****************End JBB Added Code
                    End If 'End of if checking for intersection pixels JBB
                Next 'End of loop through columns JBB
            Next 'End of loop through rows JBB

            '****Start JBB Added Code
            'for multiple dbg images
            For I = 0 To OutputRasterNames.Count - 1
                For R = 0 To WBOutDateCount - 1 'MultispectralImages.Count - 1
                    OutRasterPixelBlockDbg(I, R).PixelData(0) = OutPixelsDbg(I, R)
                    OutRasterEditDbg(I, R).Write(CType(OutRasterCursorDbg(I, R).TopLeft, ESRI.ArcGIS.Geodatabase.IPnt), OutRasterPixelBlockDbg(I, R))
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(OutRasterPixelBlockDbg(I, R))
                Next
            Next

            IntersectRasterCursor.Next()

            'for multipe debug images

            For I = 0 To OutputRasterNames.Count - 1
                For R = 0 To WBOutDateCount - 1 'MultispectralImages.Count - 1
                    OutRasterCursorDbg(I, R).Next() 'for one dbg image
                Next
            Next
            '******************End JBB Added Code

            For I = 1 To InputRasters.Count - 1
                InRasterCursor(I).Next()
            Next

            'System.Runtime.InteropServices.Marshal.ReleaseComObject(MultispectralPixelBlock)
            ProgressPartWater.PerformStep()
            If Abort = True Then Exit Sub 'Exits is the abort button is clicked on the form JBB
            Windows.Forms.Application.DoEvents()
        Loop While InRasterCursor(0).Next = True
        ProgressAllWater.PerformStep()

        'For i = 0 To OutputRasters.Count - 1 'This Block Commented out by JBB
        '    Dim Geoprocessor As New ESRI.ArcGIS.Geoprocessor.Geoprocessor
        '    Geoprocessor.Execute(New ESRI.ArcGIS.DataManagementTools.CalculateStatistics(OutputRasters(i)), Nothing)
        '    System.Runtime.InteropServices.Marshal.ReleaseComObject(OutputRasters(i))
        '    System.Runtime.InteropServices.Marshal.ReleaseComObject(OutRaster2(i))
        '    System.Runtime.InteropServices.Marshal.ReleaseComObject(OutRasterBand(i))
        '    System.Runtime.InteropServices.Marshal.ReleaseComObject(OutRasterCursor(i))
        '    System.Runtime.InteropServices.Marshal.ReleaseComObject(OutRasterEdit(i))
        'Next

        '***************Added by JBB********************
        'for multiple debug images
        Dim Geoprocessor1 As New ESRI.ArcGIS.Geoprocessor.Geoprocessor
        CalculationTextWater.AppendText(vbNewLine & "   Calculating statistics For debug images...") : Windows.Forms.Application.DoEvents() 'Updates the progress text JBB
        For I = 0 To OutputRasterNames.Count - 1
            For R = 0 To WBOutDateCount - 1 'MultispectralImages.Count - 1
                Geoprocessor1.AddOutputsToMap = False
                Dim SetRasterProperties1 As New ESRI.ArcGIS.DataManagementTools.SetRasterProperties
                SetRasterProperties1.in_raster = OutputRastersDbg(I, R)
                SetRasterProperties1.nodata = "1 -3.40282346638529E+38"
                Geoprocessor1.Execute(SetRasterProperties1, Nothing) 'I believe that this is taking the no data number and changing it to null or taking a null value and changing ti to a really large negative JBB

                Geoprocessor1.AddOutputsToMap = True 'This helps following output images be added to the current map JBB
                Dim CalculateStatistics1 As New ESRI.ArcGIS.DataManagementTools.CalculateStatistics()
                CalculateStatistics1.in_raster_dataset = OutputRastersDbg(I, R)
                CalculateStatistics1.ignore_values = Single.MinValue
                Geoprocessor1.Execute(CalculateStatistics1, Nothing) 'I think this helps the image not display the null pixels JBB
                System.Runtime.InteropServices.Marshal.ReleaseComObject(OutputRastersDbg(I, R))
                System.Runtime.InteropServices.Marshal.ReleaseComObject(OutRaster2Dbg(I, R))
                System.Runtime.InteropServices.Marshal.ReleaseComObject(OutRasterBandDbg(I, R))
                System.Runtime.InteropServices.Marshal.ReleaseComObject(OutRasterCursorDbg(I, R))
                System.Runtime.InteropServices.Marshal.ReleaseComObject(OutRasterEditDbg(I, R))
            Next
        Next
        '***********End Added by JBB*********************

        For i = 0 To InputRasters.Count - 1
            Dim Path As String = InputRasters(i).CompleteName
            DeleteArcGISFile(Path)
        Next

        For i = 0 To InputRasters.Count - 1
            System.Runtime.InteropServices.Marshal.ReleaseComObject(InputRasters(i))
        Next
        System.Runtime.InteropServices.Marshal.ReleaseComObject(IntersectRaster)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(IntersectRasterValue)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(CoverTable)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(WeatherTable)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(WBOutDateTbl) 'Added by JBB to Close OUt WBOut Date Table JBB
        System.Runtime.InteropServices.Marshal.ReleaseComObject(Workspace)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(WorkspaceFactory)
        Timer.Stop() : ProgressAllWater.PerformStep()
        CalculationTextWater.AppendText(vbNewLine & "Succeeded at " & Now & vbNewLine & "Elasped time (" & Timer.Elapsed.ToString & ").")

        'Some log-type warnings added by JBB
        If FlgFalseSAVI = True Then
            MsgBox("Tabular False Peak And/Or End SAVI value(s) > 0 And were used In calculations")
        End If
        If FlgImageCntLate = True Then
            MsgBox("Too Few Late Images For SAVI Log Interpolation, Program only used False peak And End SAVI If > 0")
        End If
        If FlgImageCntEarly = True Then
            MsgBox("Too Few Early Images For SAVI Log Interpolation, Program did Not run properly")
        End If
        If FlgImageCntEarly2 = True Then
            MsgBox("Too Few Early Images For SAVI Log Interpolation With Peak SAVI, Program did Not run properly")
        End If
        If NegFlag = True Then
            MsgBox("A negative vegetation index was computed, check output For invalid pixels, e.g. NDVI, SAVI≤ 0") 'moved this out of the calcs to be less obtrusive
        End If
        If CntForceFakeSAVI > 0 Then
            Dim StrOut As String = "The False Peak SAVI DOY was forced To input value " & CntForceFakeSAVI & " times. Use caution."
            MsgBox(StrOut)
        End If
        If MaxSAVIsEarly = True Then
            MsgBox("At least one late high SAVI pixel was forced To be In the early season")
        End If
        If FlgNoInputYr = True Then
            MsgBox("No Water Balance Start Year was input, the year Of the first multispectral image was used, results may be incorrect")
        End If
        If FlgDepth1 = True Then
            MsgBox("The depth Of the top layer was extended For at least one pixel based On Zrmin (which includes the TEW limit).")
        End If
    End Sub

#End Region

#End Region

    Private Sub HelpfulButton_Click(sender As Object, e As EventArgs) Handles HelpfulButton.Click

        MsgBox("Exhaustive help information has Not been added at this time. However, help messages may be accessed by clicking On labels throughout the SETMI Interface. This code was developed at Utah State University And modified at the University Of Nebraska-Lincoln Using educational licenses Of Visual Studio, professional licenses Of VS2012 And VS2015 must be purchased If this software Is comercialized. This version includes code modifications based On PyTSEB by Hector Nieto As Shared by Bill Kustas, these portions are copywrighted undr the GNU Public License, see https://github.com/hectornieto/pyTSEB") 'added by JBB
        MsgBox("Energy Balance may work best if run in blank .mxd, Water Balance should be run in an .mxd with a map layer within the study area so that points can be properly selected")
        MsgBox("Since conversion to ArcGIS v 10.4, SETMI saves some temporary files in the C:\Users\myname\AppData\Local\Temp\  path, which are not automatically deleted. The user should periodically check this path and carefully remove any g_g###.aux.xml files with corresponding g_g### subfolders, and any tmp###.img, tmp###.img.aux.xml and tmp###.tmp files. These latter files are related to the intersect rasters. The processing steps generating the g_g files is unknown at present.")
        MsgBox("Currently the WB assimilated ET is used to recalculate Dr even if no adjustment was made to Ks (i.e. if Ks was already 1)")
    End Sub
    'Below are alist of helps that are accessed when clicking a lable in the SETMI GUI 'Added by JBB >=1/25/2018
    Private Sub Label7_Click(sender As Object, e As EventArgs) Handles Label7.Click
        MsgBox("Only the two-source model (Norman et al. 1995 https://doi.org/10.1016/0168-1923(95)02265-Y, etc.) has been extensively tested, the one-layer model (Chavez et al. 2005, https://doi.org/10.1175/JHM467.1) is in pretty good shape.")
    End Sub
    Private Sub Label21_Click(sender As Object, e As EventArgs) Handles Label21.Click
        MsgBox("How instantaneous LE is scaled to daily ET see Chavez et al. (2008) http://doi.org/10.1007/s00271-008-0122-3. 'Reference ET' uses the ratio of instantaneous reference ET to daily reference ET. 'Evaporative Fraction' uses the ratio of modeled (Rn - G) to daily (Rn-G).")
    End Sub
    Private Sub Label47_Click(sender As Object, e As EventArgs) Handles Label47.Click
        MsgBox("This is for the TSEB canopy latent heat flux estimation, see Colaizzi et al. (2012) http://dx.doi.org/10.1016/j.advwatres.2012.06.004 or Colaizzi et al. (2014) http://doi.org/10.13031/trans.57.10423")
    End Sub
    Private Sub Label48_Click(sender As Object, e As EventArgs) Handles Label48.Click
        MsgBox("How soil heat flux is computed from soil net radiation in the TSEB (Norman et al. 1995). Only 'Contant Ratio' works, following Norman et al. (1995) https://doi.org/10.1016/0168-1923(95)02265-Y, or Brutseart 1982 Evaporation Into the Atmosphere. D. Reidel Publishing Company, Dordrecht, Holland. Currenlty the computation of G is independent of selection here, but this list was added for possible future functionality.")
    End Sub
    Private Sub Label40_Click(sender As Object, e As EventArgs) Handles Label40.Click
        MsgBox("How to adjust wind speed as if it were measured over the modeled surface. 'No Adjustment' takes the measured as if over the modeled surface. 'Without Stability Terms' follows Allen & Wright  (1997) https://doi.org/10.1061/(ASCE)1084-0699(1997)2:1(26), 'With Stability Terms' is non-functional at this time")
    End Sub
    Private Sub Label49_Click(sender As Object, e As EventArgs) Handles Label49.Click
        MsgBox("How the fraction of green LAI is computed for the TSEB (e.g., Norman et al. 1995 https://doi.org/10.1016/0168-1923(95)02265-Y). 'Tabular Only' uses the input tabular value. 'Using Past Peak LAI' is recommended and uses divides the current day modeled - or input (I think) - LAI by the maximum past or present input LAI for a pixel to get Fg, kind of like Houborg et al. (2009) https://doi.org/10.1016/j.agrformet.2009.06.014")
    End Sub
    Private Sub Label50_Click(sender As Object, e As EventArgs) Handles Label50.Click
        MsgBox("How canopy height is computed for the TSEB. 'Current Hc Only' uses the current modeled (or input I think) value. 'Using Past Hc' is recommended and prevents the crop height from shrinking below the maximum value for the pixel out of input and current values")
    End Sub
    Private Sub Label52_Click(sender As Object, e As EventArgs) Handles Label52.Click
        MsgBox("How fraction of cover is computed for the TSEB. 'Current Fc Only' is recommended and uses the current modeled (or input I think) value. 'Using Past Fc' prevents the fraction of cover from nadir from shrinking below the maximum value for the pixel out of input and current values")
    End Sub
    'Private Sub Label55_Click(sender As Object, e As EventArgs)
    '    MsgBox("Which vegetation index is used to compute fraction of cover for the TSEB (Choudhury et al. 1994, https://doi.org/10.1016/0034-4257(94)90090-6).")
    'End Sub
    Private Sub Label54_Click(sender As Object, e As EventArgs) Handles Label54.Click
        MsgBox("How the canopy height/width ratio is computed for TSEB clumping (Kustas and Norman, 2000, https://doi.org/10.2134/agronj2000.925847x)")
    End Sub
    Private Sub Label20_Click(sender As Object, e As EventArgs) Handles Label20.Click
        MsgBox("This allows the user to select point and spatial water balance outputs. The difference is the view on the 'Calculation' tab. It is advisable to generate point output even if spatial is desired also. Select 'Crop Coefficient Grid' here, select output options in the 'Calculation' tab, then come back here and select 'Crop Coefficient Point', select points in the 'Calculation' tab and execute.")
    End Sub
    Private Sub Label28_Click(sender As Object, e As EventArgs) Handles Label28.Click
        MsgBox("How Actual ET, Actual Depletion, Actual Evaporated Depth, and Actual Lower Layer Water Content images are incorporated into the model' Only 'No Assimilation' (they aren't incorporated) and 'Single Weight' (see Neale et al. 2012 http://dx.doi.org/10.1016/j.advwatres.2012.10.008) function")
    End Sub
    Private Sub Label44_Click(sender As Object, e As EventArgs) Handles Label44.Click
        MsgBox("Which type of reference ET was input? The type matters, currently short reference works for corn and possibly wheat (unverified) and the tall reference works for corn and soybean, vegetation index - to - Kcb relationships programmed in SETMI are the important thing here")
    End Sub
    Private Sub Label45_Click(sender As Object, e As EventArgs) Handles Label45.Click
        MsgBox("How is runoff computed? 'SCS Curve Number' uses the SCS Runoff Equation, NRCS 2004 https://www.wcc.nrcs.usda.gov/ftpref/wntsc/H&H/NEHhydrology/ch10.pdf. 'Perecent Effective' computes precipiation as a set percentage of input precipitaion")
    End Sub
    Private Sub Label42_Click(sender As Object, e As EventArgs) Handles Label42.Click
        MsgBox("How the basal crop coefficient is interpolated / extrapolated between and beyond image dates. 'VI Interpolation' just connects the dots, so to say, between images. 'SAVI Log Interpolation' follows Campos et al. 2017 http://dx.doi.org/10.1016/j.agwat.2017.03.022. 'Fitted Curve' minimizes the sum of squares between image date Kcbrfs and hard programed time-based Kcb functions - currenlty there is one for corn based on Allen and Wright 2002 https://www.kimberly.uidaho.edu/water/asceewri/Conversion_of_Wright_Kcs_2c.pdf, I am pretty sure")
    End Sub
    Private Sub Label53_Click(sender As Object, e As EventArgs) Handles Label53.Click
        MsgBox("How fraction of cover is computed for the Water Balance. 'FAO 56 follows the methodology of FAO-56 http://www.fao.org/docrep/X0490E/X0490E00.htm. 'Vegetation Index' follows Choudhury et al. (1994) https://doi.org/10.1016/0034-4257(94)90090-6, and will used the vegetation index and parameters input for the TSEB 'Fraction of Cover Veg. Index'.")
    End Sub
    Private Sub Label8_Click(sender As Object, e As EventArgs) Handles Label8.Click
        MsgBox("Some relationships in SETMI are image source specific. Only 'Landsat' and 'Airborne' function. 'Airborne' refers to USU's airborne system. 'Unmanned Aircraft' was added in anticipation of using UAV imagery at UNL.")
    End Sub
    Private Sub Label17_Click(sender As Object, e As EventArgs) Handles Label17.Click
        MsgBox("Time zone is important in the TSEB model. All input images should have a time stamp in the file name matching this time zone.")
    End Sub
    Private Sub Label57_Click(sender As Object, e As EventArgs) Handles Label57.Click
        MsgBox("Method of forecasting peak SAVI and/or end-of-season SAVI when using the 'SAVI Log Interpolation' 'Kcb Progression Method' above. The forecasted latest date for peak SAVI or actual date for end-of-season SAVI are determined based on 'Day of Year' or cumulative 'Growing Degree Days' from (and including) the water balance start day of year (see 'Cover' tab). GDD could be added or subtracted as desired to represent a different datum date. At this time the SAVI progression itself is only growing degree based, to make this be day based, simply input daily maximum and minimum air temperatures such that each day computes the same number of growing degree days (best would be 1 GDD/day), thus accomplishing the same daily increase as days would. If this is done, then selecting 'Growing Degree Days' here is the same as using days after intitiation, if the input forecast values are appropriately assigned. ***** Note that the forecasting method is determined on the 'Cover Property Table' inputs in the 'Cover' tab. *****")
    End Sub
    Private Sub Label56_Click(sender As Object, e As EventArgs) Handles Label56.Click
        MsgBox("Method for limiting the deep percolation rate. 'No Limit' is like FAO-56 http://www.fao.org/docrep/X0490E/X0490E00.htm. 'Constant Limit' applies the same upper limit to deep percolation every day. 'Decaying Limit' uses the drainage function of AquaCrop v.6.0 Raes, D., P. Steduto, T. C. Hsaio, E. Fereres. March 2017. Ch.3 Calculation procedures, AquaCrop Version 6.0 Reference Manual, Food and Agriculture Organization, United Nations, Rome, Italy. Deep percolation limits are applied as in Jensen and Allen. 2016. Evaporation, Evapotranspiration, and Irrigation Water Requirements 2nd Ed. ASCE Manuals and Reports on Engineering Practice No. 70. ASCE, Reston, VA. Eq. 10-12b.")
    End Sub
    Private Sub Label12_Click(sender As Object, e As EventArgs) Handles Label12.Click
        MsgBox("Match the input land cover with SETMI programmed land cover. Most variables here are used for the TSEB, though some are for the water balance. Arundo, Conifer, Cottonwood, Mesophytes, Mesquite, Watermelon, Pineapple, Non Irrigated Arable Land should all be the same as 'Default Cover'. If multiple 'Default Cover' crops are desired, please select one of these for each unique cover as SETMI doesn't seem to like multiple input covers using the same SETMI defined cover in a given model run. Zero or negative values toggle some of these inputs on and off (e.g. Fc parameters, for which default values are used if any of the inputs are 0 or negative). For Kcb, only selecting 'Hardcoded' as the 'Kcb VI' fully toggles this to default, so beware. In the case of Wc and Clumping D, if the respective paramter is to be used (see 'Project Tab' > 'Clumping Method'), if a zero or negative is input, the results are intended for water and maybe bare surfaces, results may be erroneous in other cases, particularly for some of the trees, etc. because of how roughness terms are computed and because all clumping values are set = 1.")
    End Sub
    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        MsgBox("Input a land cover image with land cover names as strings as the value for each pixel. Coordinate system and pixel grid would be best if matched to imagery. Must have a date i.e. _MM-DD-YYYY hh-mm.tif, etc. SETMI should use the landcover image closest in time for a given multispectral image. However multiple land cover images have not been tested and would only function effectively for the energy balance models.")
    End Sub
    Private Sub Label36_Click(sender As Object, e As EventArgs) Handles Label36.Click
        MsgBox("This will be a .xlsx formatted file. There is logic to guess the input column name with SETMI programmed variables, however the user should double check to make sure each pairing is correct. These are water balance variables. Not all variables are needed depending on modeling options, particularly those relating to crop coefficients")
    End Sub
    Private Sub Label18_Click(sender As Object, e As EventArgs) Handles Label18.Click
        MsgBox("This will be a .xlsx formatted file with only one column. The column must still be matched with the SETMI programmed variable. This is a list for which the user wants SETMI to generate water balance output imagery. SETMI will compute the full period water balance, but will only generate output on these dates. Outputting many images may take time.")
    End Sub
    Private Sub Label13_Click(sender As Object, e As EventArgs) Handles Label13.Click
        MsgBox("This will be a .xlsx formatted file. Weather data should be supplied for each cover included in the energy balance. It is apparent that the waterbalance uses only the first set (first cover) of data (data should be sorted by cover and then by date). There is logic to guess the input column name with SETMI programmed variables, however the user should double check to make sure each pairing is correct. Instantaneous variables are for energy balance computations, daily are for water balance computations. Of the input variables, only daily reference ET is used in both energy balance and water balance computations. Reference ET should match the selection on the 'Project' tab if using the water balance. For the energy balance instananeous and daily reference ET should use the same reference. Variables relating to measurement heights, fetch, etc. are for the energy balance. Not all variables are needed depending on modeling options. Typical RHmin and wind speed are not used at this time.")
    End Sub
    Private Sub Label19_Click(sender As Object, e As EventArgs) Handles Label19.Click
        MsgBox("Used in energy balance computations. If provided, takes presidence over tabular data.")
    End Sub
    Private Sub Label24_Click(sender As Object, e As EventArgs) Handles Label24.Click
        MsgBox("Used in energy balance computations. If provided, takes presidence over tabular vapor pressure data.")
    End Sub
    Private Sub Label35_Click(sender As Object, e As EventArgs) Handles Label35.Click
        MsgBox("Used in energy balance computations. If provided, takes presidence over tabular data.")
    End Sub
    Private Sub Label26_Click(sender As Object, e As EventArgs) Handles Label26.Click
        MsgBox("Used in energy balance computations. If provided, takes presidence over tabular data.")
    End Sub
    Private Sub Label27_Click(sender As Object, e As EventArgs) Handles Label27.Click
        MsgBox("Used in energy balance computations. If provided, takes presidence over tabular data.")
    End Sub
    Private Sub Label34_Click(sender As Object, e As EventArgs) Handles Label34.Click
        MsgBox("Used in energy balance computations. If provided, takes presidence over tabular data.")
    End Sub
    Private Sub Label33_Click(sender As Object, e As EventArgs) Handles Label33.Click
        MsgBox("Used in ENERGY BALANCE computations ONLY. If provided, takes presidence over tabular data.")
    End Sub
    Private Sub Label29_Click(sender As Object, e As EventArgs) Handles Label29.Click
        MsgBox("Used in water balance computations. This is energy balance modeled ET. If provided, and 'Single Weight' is selected as the 'Data Assimilation' method in the 'Project' tab, then this is how the hybrid functionality of the model is implemented.")
    End Sub
    Private Sub Label37_Click(sender As Object, e As EventArgs) Handles Label37.Click
        MsgBox("Used in water balance computations. If provided, takes presidence over tabular data.")
    End Sub
    Private Sub Label38_Click(sender As Object, e As EventArgs) Handles Label38.Click
        MsgBox("Used in water balance computations. If provided, takes presidence over tabular data. This is the only way to get multiple areas of the same crop to be irrigated at different times")
    End Sub
    Private Sub Label39_Click(sender As Object, e As EventArgs) Handles Label39.Click
        MsgBox("Used in water balance computations. This is measured or otherwise modeled root zone depletion and is 'assimilated' similarly to the 'Actual Evapotranspiration'. If input brings water balance to permanent wilting point, the evaporative layer will not automatically be forced to zero evaporation, this would need to be accomplished by inputting an 'Actual Evaporated Depth' image of sufficient magnitude to do this.")
    End Sub
    Private Sub Label43_Click(sender As Object, e As EventArgs) Handles Label43.Click
        MsgBox("Used in water balance computations. This is measured or otherwise modeled water content in the lower soil layer and is 'assimilated' similarly to the 'Actual Evapotranspiration'.")
    End Sub
    Private Sub Label46_Click(sender As Object, e As EventArgs) Handles Label46.Click
        MsgBox("Used in water balance computations. This is measured or otherwise modeled evaporated depth and is 'assimilated' similarly to the 'Actual Evapotranspiration'.")
    End Sub
    Private Sub Label9_Click(sender As Object, e As EventArgs) Handles Label9.Click
        MsgBox("Stacked shortwave reflectance (0 to 1), including at least red and near infrared reflectance. Used in both energy balance and water balance computations. As with all other raster inputs, it should have a time stamp _MM-DD-YYYY hh-mm.file extension at the end of the file.")
    End Sub
    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter
        MsgBox("Match the input imagery band order here, leave blank for bands not included in the imagery")
    End Sub
    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click
        MsgBox("Corrected surface temperature imagery (C). Used in energy balance computations. Not needed for water balance only runs.")
    End Sub
    Private Sub Label30_Click(sender As Object, e As EventArgs) Handles Label30.Click
        MsgBox("Measured or otherwise derived vegetation height (m). Used in energy balance computations. Takes precidence over modeled values. Also, when canopy height imagery from previous days in the season is input, it provides a means of preventing the modeled crop height from decreasing below values computed earlier in the season. See 'Canopy Height Method' in the 'Project' tab.")
    End Sub
    Private Sub Label31_Click(sender As Object, e As EventArgs) Handles Label31.Click
        MsgBox("Measured or otherwise derived leaf area index (-). Used in energy balance computations. Takes precidence over modeled values. Also, when LAI imagery from previous days in the season is input, it provides a means of preventing the modeled LAI from decreasing below values computed earlier in the season. See 'Fg Method' in the 'Project' tab.")
    End Sub
    Private Sub Label32_Click(sender As Object, e As EventArgs) Handles Label32.Click
        MsgBox("For MODIS imagery. This is for the two-source energy balance (Norman et al. 1995 https://doi.org/10.1016/0168-1923(95)02265-Y, etc.). Not all modeled cover properties have programmed relationships for MODIS")
    End Sub
    Private Sub Label51_Click(sender As Object, e As EventArgs) Handles Label51.Click
        MsgBox("Measured or otherwise derived nadir fraction of vegetation cover. Used in energy balance computations. Takes precidence over modeled values. Also, when fraction of cover imagery from previous days in the season is input, it provides a means of preventing the modeled fraction of cover from decreasing below values computed earlier in the season. See 'Fraction of Cover Method' in the 'Project' tab.")
    End Sub
    Private Sub Label22_Click(sender As Object, e As EventArgs) Handles Label22.Click
        MsgBox("Input vertically uniform field capacity for the entire root zone. Used in water balance computations.")
    End Sub
    Private Sub Label14_Click(sender As Object, e As EventArgs) Handles Label14.Click
        MsgBox("Input vertically uniform permanent wilting point for the entire root zone. Used in water balance computations.")
    End Sub
    Private Sub Label41_Click(sender As Object, e As EventArgs) Handles Label41.Click
        MsgBox("Input vertically uniform initial water content for the entire root zone. Used in water balance computations. Will be applied on the 'DOY Start Water Balance', see the 'Cover' tab.")
    End Sub
    Private Sub Label15_Click(sender As Object, e As EventArgs) Handles Label15.Click
        MsgBox("Energy balance output imagery will be saved to this location. Note that it is best practice to save and close the ArcGIS .mxd file (even if just temporarily) upon model execution completion in order to properly generate the output files.")
    End Sub
    Private Sub Label16_Click(sender As Object, e As EventArgs) Handles Label16.Click
        MsgBox("Select the desired output raster datasets. Some datasets are energy balance model specific (e.g., Rah, and To, and albedo are for the 'One Layer' model, see Chavez et al. (2008) http://doi.org/10.1007/s00271-008-0122-3; Rc, aPT, LEc, LEs, Hc, Hs, Lo, Fg, etc. are for TSEB, see Norman et al. 1995 https://doi.org/10.1016/0168-1923(95)02265-Y, etc.). Also, SETMI may have difficulty properly overwriting old output files. The user may choose to clear or move old files from this directory before running the model.")
    End Sub
    Private Sub Label25_Click(sender As Object, e As EventArgs) Handles Label25.Click
        MsgBox("Water balance output imagery will be saved to this location. Note that it is best practice to save and close the ArcGIS .mxd file (even if just temporarily) upon model execution completion in order to properly generate the output files. Also, SETMI may have difficulty properly overwriting old output files. The user may choose to clear or move old files from this directory before running the model.")
    End Sub
    Private Sub Label23_Click(sender As Object, e As EventArgs) Handles Label23.Click
        MsgBox("Select the desired output raster datasets. Season total datasets, e.g. 'xTotal Net Irrigation (SumInet)', are totals from 'DOY Start Water Balance' (see 'Cover' tab) to the output date (see 'Cover' tab).")
    End Sub
    Private Sub Label58_Click(sender As Object, e As EventArgs) Handles Label58.Click
        MsgBox("If running in 'Point' mode, this selects whether points are selected in the 'Calculation' tab,  provided as tabular input in the 'Cover' tab, or both.")
    End Sub
    Private Sub Label55_Click(sender As Object, e As EventArgs) Handles Label55.Click
        MsgBox("If running in 'Point' mode (see 'Project' tab), this is a table of x and y coordinates of points to produce time-series text file outputs from the water balance. The coordinates should be in WGS84 and should match the coordinate system of the input files. A good practice would be to provide all spatial inputs - and this table - in WGS84 UTM coordinates.")
    End Sub
    Private Sub Label59_Click(sender As Object, e As EventArgs) Handles Label59.Click
        MsgBox("Topmost layer, if only one layer, this is the one to use. There must always be input into this layer for Water Balance Input. The depth of this layer must be at least as large as the evaporative layer.")
    End Sub
    Private Sub Label60_Click(sender As Object, e As EventArgs) Handles Label60.Click
        MsgBox("Layer imediately below the topmost layer. Middle layer for a three-layer system; bottom layer for a two-layer system. Ommit this input if a one-layer system.")
    End Sub
    Private Sub Label61_Click(sender As Object, e As EventArgs) Handles Label61.Click
        MsgBox("Bottommost layer if a three-layer system. Ommit this input if a one- or two-layer system.")
    End Sub
    Private Sub Label71_Click(sender As Object, e As EventArgs) Handles Label71.Click
        MsgBox("The thickness of each layer. The bottom most image will restrict root growth if not sufficiently deep. Regardless of input, in a one- or two-layer model, the bottom most layer will restrict root growth beyond the layer's bottom depth.")
    End Sub
    Private Sub Label74_Click(sender As Object, e As EventArgs) Handles Label74.Click
        MsgBox("Saturated hydraulic conductivity is used if deep percolation is to be limited (see 'Project' tab). The total hydraulic conductivity is computed as a weighted average following Eq. 3.10 of Radcliffe, D. E. and J. Simunek. 2010. Soil Physics with HYDRUS. CRC Press, Taylor and Francis Group, New York.  p. 93.")
    End Sub
    Private Sub Label64_Click(sender As Object, e As EventArgs) Handles Label64.Click
        MsgBox("Topmost layer, if only one layer, this is the one to use. There must always be input into this layer for Water Balance Input. The depth of this layer must be at least as large as the evaporative layer.")
    End Sub
    Private Sub Label63_Click(sender As Object, e As EventArgs) Handles Label63.Click
        MsgBox("Layer imediately below the topmost layer. Middle layer for a three-layer system; bottom layer for a two-layer system. Ommit this input if a one-layer system.")
    End Sub
    Private Sub Label62_Click(sender As Object, e As EventArgs) Handles Label62.Click
        MsgBox("Bottommost layer if a three-layer system. Ommit this input if a one- or two-layer system.")
    End Sub
    Private Sub Label67_Click(sender As Object, e As EventArgs) Handles Label67.Click
        MsgBox("Topmost layer, if only one layer, this is the one to use. There must always be input into this layer for Water Balance Input. The depth of this layer must be at least as large as the evaporative layer.")
    End Sub
    Private Sub Label66_Click(sender As Object, e As EventArgs) Handles Label66.Click
        MsgBox("Layer imediately below the topmost layer. Middle layer for a three-layer system; bottom layer for a two-layer system. Ommit this input if a one-layer system.")
    End Sub
    Private Sub Label65_Click(sender As Object, e As EventArgs) Handles Label65.Click
        MsgBox("Bottommost layer if a three-layer system. Ommit this input if a one- or two-layer system.")
    End Sub
    Private Sub Label70_Click(sender As Object, e As EventArgs) Handles Label70.Click
        MsgBox("Topmost layer, if only one layer, this is the one to use. There must always be input into this layer for Water Balance Input. The depth of this layer must be at least as large as the evaporative layer.")
    End Sub
    Private Sub Label69_Click(sender As Object, e As EventArgs) Handles Label69.Click
        MsgBox("Layer imediately below the topmost layer. Middle layer for a three-layer system; bottom layer for a two-layer system. Ommit this input if a one-layer system.")
    End Sub
    Private Sub Label73_Click(sender As Object, e As EventArgs) Handles Label73.Click
        MsgBox("Topmost layer, if only one layer, this is the one to use. There must always be input into this layer for Water Balance Input. The depth of this layer must be at least as large as the evaporative layer.")
    End Sub
    Private Sub Label72_Click(sender As Object, e As EventArgs) Handles Label72.Click
        MsgBox("Layer imediately below the topmost layer. Middle layer for a three-layer system; bottom layer for a two-layer system. Ommit this input if a one-layer system.")
    End Sub
    Private Sub Label68_Click(sender As Object, e As EventArgs) Handles Label68.Click
        MsgBox("Bottommost layer if a three-layer system. Ommit this input if a one- or two-layer system.")
    End Sub
    Private Sub Label78_Click(sender As Object, e As EventArgs) Handles Label78.Click
        MsgBox("Volumetric water content at saturation. Used in limiting the deep percolation, see Raes, D., P. Steduto, T. C. Hsaio, E. Fereres. March 2017. Ch.3 Calculation procedures, AquaCrop Version 6.0 Reference Manual, Food and Agriculture Organization, United Nations, Rome, Italy.")
    End Sub
    Private Sub Label77_Click(sender As Object, e As EventArgs) Handles Label77.Click
        MsgBox("Topmost layer, if only one layer, this is the one to use. There must always be input into this layer for Water Balance Input. The depth of this layer must be at least as large as the evaporative layer.")
    End Sub
    Private Sub Label76_Click(sender As Object, e As EventArgs) Handles Label76.Click
        MsgBox("Layer imediately below the topmost layer. Middle layer for a three-layer system; bottom layer for a two-layer system. Ommit this input if a one-layer system.")
    End Sub
    Private Sub Label75_Click(sender As Object, e As EventArgs) Handles Label75.Click
        MsgBox("Bottommost layer if a three-layer system. Ommit this input if a one- or two-layer system.")
    End Sub
    Private Sub Label79_Click(sender As Object, e As EventArgs) Handles Label79.Click
        MsgBox("Actual 'skin layer' evaporation. See Jensen and Allen. 2016. Evaporation, Evapotranspiration, and Irrigation Water Requirements 2nd Ed. ASCE Manuals and Reports on Engineering Practice No. 70. ASCE, Reston, VA.")
    End Sub
    Private Sub Label81_Click(sender As Object, e As EventArgs) Handles Label81.Click
        MsgBox("Readily evaporable water image (REW; FAO-56 http://www.fao.org/docrep/X0490E/X0490E00.htm). This input overwrites modeled values of REW based on the upper layer soil properties. However, if this is greater than the estimated total evaporable water (TEW; FAO-56 http://www.fao.org/docrep/X0490E/X0490E00.htm), then it will be reduced to TEW, see Jensen and Allen. 2016. Evaporation, Evapotranspiration, and Irrigation Water Requirements 2nd Ed. ASCE Manuals and Reports on Engineering Practice No. 70. ASCE, Reston, VA. Eq.")
    End Sub
    Private Sub Label80_Click(sender As Object, e As EventArgs) Handles Label80.Click
        MsgBox("Bottommost layer if a three-layer system. Ommit this input if a one- or two-layer system.")
    End Sub
    Private Sub Label82_Click(sender As Object, e As EventArgs) Handles Label82.Click
        MsgBox("Include transpiration from the evaporative layer or not. If so, follows Eqs. 9-30 and 9-31 of Jensen and Allen. 2016. Evaporation, Evapotranspiration, and Irrigation Water Requirements 2nd Ed. ASCE Manuals and Reports on Engineering Practice No. 70. ASCE, Reston, VA. ")
    End Sub
    Private Sub Label83_Click(sender As Object, e As EventArgs) Handles Label83.Click
        MsgBox("Constant limit of deep percolation if such is selected as the 'Deep Percolation Limit Method' in the 'Project' tab. See, Jensen and Allen. 2016. Evaporation, Evapotranspiration, and Irrigation Water Requirements 2nd Ed. ASCE Manuals and Reports on Engineering Practice No. 70. ASCE, Reston, VA.")
    End Sub
    Private Sub Label84_Click(sender As Object, e As EventArgs) Handles Label84.Click
        MsgBox("Select whether skin evaporation is included or not, See, See, Jensen and Allen. 2016. Evaporation, Evapotranspiration, and Irrigation Water Requirements 2nd Ed. ASCE Manuals and Reports on Engineering Practice No. 70. ASCE, Reston, VA.")
    End Sub
End Class