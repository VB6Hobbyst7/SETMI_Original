Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI

<ComClass(SETMItool_Original.ClassId, SETMItool_Original.InterfaceId, SETMItool_Original.EventsId), System.Runtime.InteropServices.ProgId("SETMI.SETMItool")>
Public NotInheritable Class SETMItool_Original
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "f98008a3-ea61-4d99-8f3d-dea58d79da05"
    Public Const InterfaceId As String = "4f0faf6e-0e4c-4d70-8317-55b6929f152c"
    Public Const EventsId As String = "365a2d70-dafa-4ac1-a414-2594c4139755"
#End Region

#Region "COM Registration Function(s)"
    <System.Runtime.InteropServices.ComRegisterFunction(), System.Runtime.InteropServices.ComVisibleAttribute(False)>
    Public Shared Sub RegisterFunction(ByVal registerType As Type)
        ' Required for ArcGIS Component Category Registrar support
        ArcGISCategoryRegistration(registerType)

        'Add any COM registration code after the ArcGISCategoryRegistration() call

    End Sub

    <System.Runtime.InteropServices.ComUnregisterFunction(), System.Runtime.InteropServices.ComVisibleAttribute(False)>
    Public Shared Sub UnregisterFunction(ByVal registerType As Type)
        ' Required for ArcGIS Component Category Registrar support
        ArcGISCategoryUnregistration(registerType)

        'Add any COM unregistration code after the ArcGISCategoryUnregistration() call

    End Sub

#Region "ArcGIS Component Category Registrar generated code"
    Private Shared Sub ArcGISCategoryRegistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxCommands.Register(regKey)

    End Sub
    Private Shared Sub ArcGISCategoryUnregistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxCommands.Unregister(regKey)

    End Sub

#End Region
#End Region

    Public Shared MainForm As Main
    Public Shared ArcApplication As IApplication
    Public Shared ArcMapDocument As IMxDocument

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()

        ' TODO: Define values for the public properties
        ' TODO: Define values for the public properties
        MyBase.m_category = "Utah State University"  'localizable text 
        MyBase.m_caption = "SETMI"   'localizable text 
        MyBase.m_message = "Spatial ET Modeling Interface"   'localizable text 
        MyBase.m_toolTip = "Spatial ET Modeling Interface" 'localizable text 
        MyBase.m_name = "SETMI_Original"  'unique id, non-localizable (e.g. "MyCategory_ArcMapTool")

        Try
            'TODO: change resource name if necessary
            MyBase.m_bitmap = My.Resources.SETMI.ToBitmap
        Catch ex As Exception
            System.Diagnostics.Trace.WriteLine(ex.Message, "Invalid Bitmap")
        End Try
    End Sub

    Public Overrides Sub OnCreate(ByVal hook As Object)
        If Not hook Is Nothing Then
            ArcApplication = CType(hook, IApplication)
            If TypeOf hook Is IMxApplication Then
                MyBase.m_enabled = True
            Else
                MyBase.m_enabled = False
            End If
            ArcMapDocument = CType(ArcApplication.Document, IMxDocument)
        End If
    End Sub

    Public Overrides Sub OnClick()
        MainForm = New Main()
        MainForm.Show()
    End Sub

End Class

