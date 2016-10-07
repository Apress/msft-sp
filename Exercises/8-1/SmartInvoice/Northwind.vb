Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.Office.Interop.SmartTag
Imports Word = Microsoft.Office.Interop.Word

Public Class Northwind
    Implements ISmartDocument

    'Variables
    Public Const URI As String = "urn:schemas-microsoft-com.DataLan.SmartInvoice"
    Public Const PRODUCTID As String = Uri & "#ProductID"
    Public Const PRODUCTNAME As String = Uri & "#ProductName"
    Public Const QUANTITYPERUNIT As String = Uri & "#QuantityPerUnit"
    Public Const UNITPRICE As String = Uri & "#UnitPrice"
    Dim intSelectedIndex As Integer = 0
    Dim objDataSet As DataSet

    Public ReadOnly Property ControlCaptionFromID(ByVal ControlID As Integer, _
    ByVal ApplicationName As String, ByVal LocaleID As Integer, _
    ByVal Text As String, ByVal Xml As String, ByVal Target As Object) _
    As String Implements Microsoft.Office.Interop.SmartTag.ISmartDocument.ControlCaptionFromID
        Get
            Select Case ControlID
                Case 1
                    Return "Product ID"
                Case 2, 102, 202, 302
                    Return "Insert"
                Case 101
                    Return "Product Name"
                Case 201
                    Return "Quantity Per Unit"
                Case 301
                    Return "Unit Price"
            End Select
        End Get
    End Property

    Public ReadOnly Property ControlCount(ByVal XMLTypeName As String) As Integer _
    Implements Microsoft.Office.Interop.SmartTag.ISmartDocument.ControlCount
        Get
            Select Case XMLTypeName
                Case PRODUCTID
                    Return 2
                Case PRODUCTNAME
                    Return 2
                Case QUANTITYPERUNIT
                    Return 2
                Case UNITPRICE
                    Return 2
            End Select
        End Get
    End Property

    Public ReadOnly Property ControlID(ByVal XMLTypeName As String, _
    ByVal ControlIndex As Integer) _
    As Integer Implements Microsoft.Office.Interop.SmartTag.ISmartDocument.ControlID
        'ControlIndex is just the index for each set 1,2,3...
        'Therefore, we add an arbitrary number to guarantee uniqueness
        Get
            Select Case XMLTypeName
                Case PRODUCTID
                    Return ControlIndex
                Case PRODUCTNAME
                    Return ControlIndex + 100
                Case QUANTITYPERUNIT
                    Return ControlIndex + 200
                Case UNITPRICE
                    Return ControlIndex + 300
            End Select
        End Get
    End Property

    Public ReadOnly Property ControlNameFromID(ByVal ControlID As Integer) As String _
    Implements Microsoft.Office.Interop.SmartTag.ISmartDocument.ControlNameFromID
        Get
            Return URI & ControlID.ToString
        End Get
    End Property

    Public ReadOnly Property ControlTypeFromID(ByVal ControlID As Integer, _
    ByVal ApplicationName As String, ByVal LocaleID As Integer) _
    As Microsoft.Office.Interop.SmartTag.C_TYPE Implements _
    Microsoft.Office.Interop.SmartTag.ISmartDocument.ControlTypeFromID
        Get
            Select Case ControlID
                Case 1
                    Return C_TYPE.C_TYPE_COMBO
                Case 2, 102, 202, 302
                    Return C_TYPE.C_TYPE_BUTTON
                Case 101, 201, 301
                    Return C_TYPE.C_TYPE_TEXTBOX
            End Select
        End Get
    End Property

    Public Sub ImageClick(ByVal ControlID As Integer, ByVal ApplicationName As String, ByVal Target As Object, ByVal Text As String, ByVal Xml As String, ByVal LocaleID As Integer, ByVal XCoordinate As Integer, ByVal YCoordinate As Integer) Implements Microsoft.Office.Interop.SmartTag.ISmartDocument.ImageClick

    End Sub

    Public Sub OnCheckboxChange(ByVal ControlID As Integer, ByVal Target As Object, ByVal Checked As Boolean) Implements Microsoft.Office.Interop.SmartTag.ISmartDocument.OnCheckboxChange

    End Sub

    Public Sub OnPaneUpdateComplete(ByVal Document As Object) Implements Microsoft.Office.Interop.SmartTag.ISmartDocument.OnPaneUpdateComplete

    End Sub

    Public Sub OnRadioGroupSelectChange(ByVal ControlID As Integer, ByVal Target As Object, ByVal Selected As Integer, ByVal Value As String) Implements Microsoft.Office.Interop.SmartTag.ISmartDocument.OnRadioGroupSelectChange

    End Sub

    Public Sub OnTextboxContentChange(ByVal ControlID As Integer, ByVal Target As Object, ByVal Value As String) Implements Microsoft.Office.Interop.SmartTag.ISmartDocument.OnTextboxContentChange

    End Sub

    Public Sub PopulateActiveXProps(ByVal ControlID As Integer, ByVal ApplicationName As String, ByVal LocaleID As Integer, ByVal Text As String, ByVal Xml As String, ByVal Target As Object, ByVal Props As Microsoft.Office.Interop.SmartTag.ISmartDocProperties, ByVal ActiveXPropBag As Microsoft.Office.Interop.SmartTag.ISmartDocProperties) Implements Microsoft.Office.Interop.SmartTag.ISmartDocument.PopulateActiveXProps

    End Sub

    Public Sub PopulateCheckbox(ByVal ControlID As Integer, ByVal ApplicationName As String, ByVal LocaleID As Integer, ByVal Text As String, ByVal Xml As String, ByVal Target As Object, ByVal Props As Microsoft.Office.Interop.SmartTag.ISmartDocProperties, ByRef Checked As Boolean) Implements Microsoft.Office.Interop.SmartTag.ISmartDocument.PopulateCheckbox

    End Sub

    Public Sub PopulateDocumentFragment(ByVal ControlID As Integer, ByVal ApplicationName As String, ByVal LocaleID As Integer, ByVal Text As String, ByVal Xml As String, ByVal Target As Object, ByVal Props As Microsoft.Office.Interop.SmartTag.ISmartDocProperties, ByRef DocumentFragment As String) Implements Microsoft.Office.Interop.SmartTag.ISmartDocument.PopulateDocumentFragment

    End Sub

    Public Sub PopulateHelpContent(ByVal ControlID As Integer, ByVal ApplicationName As String, ByVal LocaleID As Integer, ByVal Text As String, ByVal Xml As String, ByVal Target As Object, ByVal Props As Microsoft.Office.Interop.SmartTag.ISmartDocProperties, ByRef Content As String) Implements Microsoft.Office.Interop.SmartTag.ISmartDocument.PopulateHelpContent

    End Sub

    Public Sub PopulateImage(ByVal ControlID As Integer, ByVal ApplicationName As String, ByVal LocaleID As Integer, ByVal Text As String, ByVal Xml As String, ByVal Target As Object, ByVal Props As Microsoft.Office.Interop.SmartTag.ISmartDocProperties, ByRef ImageSrc As String) Implements Microsoft.Office.Interop.SmartTag.ISmartDocument.PopulateImage

    End Sub

    Public Sub PopulateListOrComboContent(ByVal ControlID As Integer, _
    ByVal ApplicationName As String, ByVal LocaleID As Integer, _
    ByVal Text As String, ByVal Xml As String, ByVal Target As Object, _
    ByVal Props As Microsoft.Office.Interop.SmartTag.ISmartDocProperties, _
    ByRef List As System.Array, ByRef Count As Integer, ByRef InitialSelected As Integer) _
    Implements Microsoft.Office.Interop.SmartTag.ISmartDocument.PopulateListOrComboContent

        Select Case ControlID
            Case 1
                'Set up connection string from custom properties
                Dim strPassword As String = ""
                Dim strUserName As String = "sa"
                Dim strDatabase As String = "Northwind"
                Dim strSQLServer = "SPSPortal"

                Dim strConnection As String = "Password=" & strPassword
                strConnection += ";Persist Security Info=True;User ID=" + strUserName
                strConnection += ";Initial Catalog=" + strDatabase
                strConnection += ";Data Source=" + strSQLServer

                'Create SQL String
                Dim strSQL As String = "SELECT ProductID, ProductName, " & _
    "QuantityPerUnit, UnitPrice FROM Products"

                'Try to run the query
                With New SqlDataAdapter
                    objDataSet = New DataSet("root")
                    .SelectCommand = New SqlCommand(strSQL, New SqlConnection(strConnection))
                    .Fill(objDataSet, "Products")
                End With

                'Fill List
                Dim index As Integer = 0
                Dim objTable As DataTable = objDataSet.Tables("Products")
                Dim objRows As DataRowCollection = objTable.Rows
                Dim objRow As DataRow

                'Set the number of items in the list
                Count = objTable.Rows.Count
                For Each objRow In objRows
                    index += 1
                    List(index) = objRow.Item("ProductID")
                Next
                'Select the first item
                InitialSelected = 1
        End Select

    End Sub

    Public Sub PopulateOther(ByVal ControlID As Integer, ByVal ApplicationName As String, _
    ByVal LocaleID As Integer, ByVal Text As String, ByVal Xml As String, ByVal Target As _
    Object, ByVal Props As Microsoft.Office.Interop.SmartTag.ISmartDocProperties) Implements _
    Microsoft.Office.Interop.SmartTag.ISmartDocument.PopulateOther
        'Set control values
        Select Case ControlID
            Case 2, 102, 202, 302
                Text = "Insert"
        End Select
    End Sub

    Public Sub PopulateRadioGroup(ByVal ControlID As Integer, ByVal ApplicationName As String, ByVal LocaleID As Integer, ByVal Text As String, ByVal Xml As String, ByVal Target As Object, ByVal Props As Microsoft.Office.Interop.SmartTag.ISmartDocProperties, ByRef List As System.Array, ByRef Count As Integer, ByRef InitialSelected As Integer) Implements Microsoft.Office.Interop.SmartTag.ISmartDocument.PopulateRadioGroup

    End Sub

    Public Sub PopulateTextboxContent(ByVal ControlID As Integer, _
    ByVal ApplicationName As String, ByVal LocaleID As Integer, ByVal Text As String, _
    ByVal Xml As String, ByVal Target As Object, _
    ByVal Props As Microsoft.Office.Interop.SmartTag.ISmartDocProperties, ByRef Value As _
    String) Implements Microsoft.Office.Interop.SmartTag.ISmartDocument.PopulateTextboxContent

        'Variables for insert text
        Dim txtProductName As String
        Dim txtQuantityPerUnit As String
        Dim txtUnitPrice As String

        'Get values based on Product ID selection
        If objDataSet.Tables.Count > 0 Then
            txtProductName = _
   objDataSet.Tables("Products").Rows.Item(intSelectedIndex).Item("ProductName").ToString
            txtQuantityPerUnit = _
    objDataSet.Tables("Products").Rows.Item(intSelectedIndex).Item("QuantityPerUnit").ToString
            txtUnitPrice = _
    objDataSet.Tables("Products").Rows.Item(intSelectedIndex).Item("UnitPrice").ToString
        End If

        'Set control values
        Select Case ControlID
            Case 101
                Value = txtProductName
            Case 201
                Value = txtQuantityPerUnit
            Case 301
                Value = txtUnitPrice
        End Select

    End Sub

    Public Sub SmartDocInitialize(ByVal ApplicationName As String, ByVal Document As Object, ByVal SolutionPath As String, ByVal SolutionRegKeyRoot As String) Implements Microsoft.Office.Interop.SmartTag.ISmartDocument.SmartDocInitialize

    End Sub

    Public ReadOnly Property SmartDocXmlTypeCaption(ByVal XMLTypeID As Integer, _
    ByVal LocaleID As Integer) As String Implements Microsoft.Office.Interop.SmartTag.ISmartDocument.SmartDocXmlTypeCaption
        'Order must be the same as in Step 2
        Get
            Select Case XMLTypeID
                Case 1
                    Return "Product ID"
                Case 2
                    Return "Product Name"
                Case 3
                    Return "Quantity per Unit"
                Case 4
                    Return "Unit Price"
            End Select
        End Get
    End Property

    Public ReadOnly Property SmartDocXmlTypeCount() As Integer Implements _
    Microsoft.Office.Interop.SmartTag.ISmartDocument.SmartDocXmlTypeCount
        Get
            Return 4
        End Get
    End Property

    Public ReadOnly Property SmartDocXmlTypeName(ByVal XMLTypeID As Integer) As String _
    Implements Microsoft.Office.Interop.SmartTag.ISmartDocument.SmartDocXmlTypeName
        'Returns the the element name
        'Order is not important
        Get
            Select Case XMLTypeID
                Case 1
                    Return PRODUCTID
                Case 2
                    Return PRODUCTNAME
                Case 3
                    Return QUANTITYPERUNIT
                Case 4
                    Return UNITPRICE
            End Select
        End Get
    End Property

    'List
    Public Sub OnListOrComboSelectChange(ByVal ControlID As Integer, ByVal Target As Object, _
    ByVal Selected As Integer, ByVal Value As String) Implements _
    Microsoft.Office.Interop.SmartTag.ISmartDocument.OnListOrComboSelectChange
        intSelectedIndex = Selected - 1
    End Sub

    'Buttons
    Public Sub InvokeControl(ByVal ControlID As Integer, ByVal ApplicationName As String, _
    ByVal Target As Object, ByVal Text As String, ByVal Xml As String, ByVal LocaleID As _
    Integer) Implements Microsoft.Office.Interop.SmartTag.ISmartDocument.InvokeControl

        Dim objRange As Word.Range
        objRange = CType(Target, Word.Range)

        'Create insert text from a text control based on which button is pushed
        Select Case ControlID
            Case 2
                Dim intIndex As Integer = _
    objRange.XMLNodes(1).SmartTag.SmartTagActions(1).ListSelection
                objRange.XMLNodes(1).Text = "Product " & _
    objDataSet.Tables("Products").Rows.Item(intSelectedIndex).Item("ProductID")
            Case 102, 202, 302
                objRange.XMLNodes(1).Text = _
    objRange.XMLNodes(1).SmartTag.SmartTagActions(1).TextboxText
            Case 203

        End Select

    End Sub


End Class

