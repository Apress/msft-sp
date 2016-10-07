Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports Word = Microsoft.Office.Interop.Word

Public Class ThisDocument

    'Controls
    Private WithEvents lstProductID As New ListBox
    Private txtName As New TextBox
    Private txtQuantity As New TextBox
    Private txtPrice As New TextBox
    Private WithEvents btnInsertID As New Button
    Private WithEvents btnInsertName As New Button
    Private WithEvents btnInsertQuantity As New Button
    Private WithEvents btnInsertPrice As New Button

    'Cached DataSet
    <Cached()> Private objProductSet As DataSet

    'Current Selection
    Private objSelection As Word.Selection

    'INITIALIZATION
    Private Sub ThisDocument_Startup(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Startup

        'Add Controls to Actions Pane
        With ActionsPane.Controls
            .Add(lstProductID)
            btnInsertID.Text = "Insert Number"
            btnInsertID.Enabled = False
            .Add(btnInsertID)
            .Add(txtName)
            btnInsertName.Text = "Insert Name"
            btnInsertName.Enabled = False
            .Add(btnInsertName)
            .Add(txtQuantity)
            btnInsertQuantity.Text = "Insert Quantity"
            btnInsertQuantity.Enabled = False
            .Add(btnInsertQuantity)
            .Add(txtPrice)
            btnInsertPrice.Text = "Insert Price"
            btnInsertPrice.Enabled = False
            .Add(btnInsertPrice)
        End With

        'Get Product Information if not already cached
        If objProductSet Is Nothing Then

            Try

                'Connection string
                'Dim strConnection As String = "Password=;User ID=sa;Initial Catalog=Northwind;Data Source=VSTO\SQLExpress;"
                Dim strConnection As String = "Integrated Security=SSPI;Initial Catalog=Northwind;Data Source=VSTO\SQLExpress;"

                'SQL string
                Dim strSQL As String = "Select ProductID, ProductName, QuantityPerUnit, UnitPrice FROM Products"

                'Run query
                With New SqlDataAdapter
                    objProductSet = New DataSet("root")
                    .SelectCommand = New SqlCommand(strSQL, New SqlConnection(strConnection))
                    .Fill(objProductSet, "Products")
                End With

            Catch x As Exception
                MsgBox(x.Message)
            End Try

        End If

        'Fill List with Product IDs
        With lstProductID
            .DataSource = objProductSet.Tables("Products")
            .DisplayMember = "ProductName"
            .ValueMember = "ProductID"
        End With

    End Sub

    'SHOW DATA IN ACTIONS PANE
    Private Sub lstProductID_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstProductID.SelectedIndexChanged

        'Update the text boxes
        txtName.Text = DirectCast(lstProductID.SelectedItem, DataRowView).Row("ProductName").ToString
        txtQuantity.Text = DirectCast(lstProductID.SelectedItem, DataRowView).Row("QuantityPerUnit").ToString
        txtPrice.Text = DirectCast(lstProductID.SelectedItem, DataRowView).Row("UnitPrice").ToString

    End Sub

    'ENABLE BUTTONS
    Private Sub ProductIDNode_ContextEnter(ByVal sender As Object, ByVal e As Microsoft.Office.Tools.Word.ContextChangeEventArgs) Handles ProductIDNode.ContextEnter
        btnInsertID.Enabled = True
    End Sub

    Private Sub ProductIDNode_ContextLeave(ByVal sender As Object, ByVal e As Microsoft.Office.Tools.Word.ContextChangeEventArgs) Handles ProductIDNode.ContextLeave
        btnInsertID.Enabled = False
    End Sub

    Private Sub ProductNameNode_ContextEnter(ByVal sender As Object, ByVal e As Microsoft.Office.Tools.Word.ContextChangeEventArgs) Handles ProductNameNode.ContextEnter
        btnInsertName.Enabled = True
    End Sub

    Private Sub ProductNameNode_ContextLeave(ByVal sender As Object, ByVal e As Microsoft.Office.Tools.Word.ContextChangeEventArgs) Handles ProductNameNode.ContextLeave
        btnInsertName.Enabled = False
    End Sub

    Private Sub QuantityPerUnitNode_ContextEnter(ByVal sender As Object, ByVal e As Microsoft.Office.Tools.Word.ContextChangeEventArgs) Handles QuantityPerUnitNode.ContextEnter
        btnInsertQuantity.Enabled = True
    End Sub

    Private Sub QuantityPerUnitNode_ContextLeave(ByVal sender As Object, ByVal e As Microsoft.Office.Tools.Word.ContextChangeEventArgs) Handles QuantityPerUnitNode.ContextLeave
        btnInsertQuantity.Enabled = False
    End Sub

    Private Sub UnitPriceNode_ContextEnter(ByVal sender As Object, ByVal e As Microsoft.Office.Tools.Word.ContextChangeEventArgs) Handles UnitPriceNode.ContextEnter
        btnInsertPrice.Enabled = True
    End Sub

    Private Sub UnitPriceNode_ContextLeave(ByVal sender As Object, ByVal e As Microsoft.Office.Tools.Word.ContextChangeEventArgs) Handles UnitPriceNode.ContextLeave
        btnInsertPrice.Enabled = False
    End Sub

    'INSERT TEXT
    Private Sub ThisDocument_SelectionChange(ByVal sender As Object, ByVal e As Microsoft.Office.Tools.Word.SelectionEventArgs) Handles Me.SelectionChange
        objSelection = e.Selection
    End Sub

    Private Sub btnInsertID_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnInsertID.Click
        objSelection.Text = DirectCast(lstProductID.SelectedItem, DataRowView).Row("ProductID").ToString
    End Sub

    Private Sub btnInsertName_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnInsertName.Click
        objSelection.Text = DirectCast(lstProductID.SelectedItem, DataRowView).Row("ProductName").ToString
    End Sub

    Private Sub btnInsertQuantity_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnInsertQuantity.Click
        objSelection.Text = DirectCast(lstProductID.SelectedItem, DataRowView).Row("QuantityPerUnit").ToString
    End Sub

    Private Sub btnInsertPrice_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnInsertPrice.Click
        objSelection.Text = DirectCast(lstProductID.SelectedItem, DataRowView).Row("UnitPrice").ToString
    End Sub

End Class
