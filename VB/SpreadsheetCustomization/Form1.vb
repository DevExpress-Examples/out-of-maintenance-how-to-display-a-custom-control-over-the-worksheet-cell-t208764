Imports DevExpress.Spreadsheet
Imports System
Imports System.Drawing
Imports System.Windows.Forms

Namespace SpreadsheetCustomization
    Partial Public Class Form1
        Inherits DevExpress.XtraBars.Ribbon.RibbonForm

        Public Shared categories() As String = { "Meat/Poultry", "Condiments", "Seafood", "Dairy Products", "Grains/Cereals", "Beverages", "Confections" }
        Private workbook As IWorkbook
        Private worksheet As Worksheet
        Private categoryColumn As Range
        Private activeCell As Cell
        Private comboBox As ComboBox

        Public Sub New()
            InitializeComponent()
            workbook = spreadsheetControl1.Document
            workbook.LoadDocument("Documents\Document.xlsx", DocumentFormat.Xlsx)

            worksheet = workbook.Worksheets("Sales report")
            categoryColumn = worksheet("Table[Category]")

            ' Create a ComboBox and specify its settings.
            comboBox = CreateComboBox()

            ' Specify the SpreadsheetControl's options.
            spreadsheetControl1.Options.Behavior.Selection.AllowExtendSelection = False
            spreadsheetControl1.Options.VerticalScrollbar.Visibility = DevExpress.XtraSpreadsheet.SpreadsheetScrollbarVisibility.Hidden
            spreadsheetControl1.Options.HorizontalScrollbar.Visibility = DevExpress.XtraSpreadsheet.SpreadsheetScrollbarVisibility.Hidden

            AddHandler spreadsheetControl1.SelectionChanged, AddressOf spreadsheetControl1_SelectionChanged
            AddHandler spreadsheetControl1.CellBeginEdit, AddressOf spreadsheetControl1_CellBeginEdit
            AddHandler spreadsheetControl1.MouseWheel, AddressOf spreadsheetControl1_MouseWheel
        End Sub

        Private Sub spreadsheetControl1_MouseWheel(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
            UpdateComboBox()
        End Sub

        #Region "#displaycombobox"
        Private Sub spreadsheetControl1_SelectionChanged(ByVal sender As Object, ByVal e As EventArgs)
            UpdateComboBox()
        End Sub

        Private Sub UpdateComboBox()
            ' Access the active cell.
            activeCell = spreadsheetControl1.ActiveCell

            ' If the currently selected cell is not in the "Category" column of the worksheet table, return. 
            If Not CheckCondition() Then
                comboBox.Visible = False
                Return
            End If

            ' Otherwise, obtain the bounds of the active cell and display the ComboBox control over it. 
            Dim cellRect As Rectangle = spreadsheetControl1.GetCellBounds(activeCell.RowIndex, activeCell.ColumnIndex)
            If cellRect.IsEmpty Then
                comboBox.Visible = False
            Else
                comboBox.Bounds = cellRect
                comboBox.Visible = True
                comboBox.SelectedItem = activeCell.Value.TextValue

                UpdateComboBoxAppearance(activeCell)
            End If
        End Sub

        Private Function CheckCondition() As Boolean
            Return If(worksheet Is workbook.Worksheets.ActiveWorksheet, categoryColumn.IsIntersecting(activeCell), False)
        End Function

        #End Region ' #displaycombobox

        Private Sub UpdateComboBoxAppearance(ByVal source As Cell)
            comboBox.BackColor = source.Fill.BackgroundColor
            Dim font As SpreadsheetFont = source.Font
            comboBox.ForeColor = font.Color
            comboBox.Font = New Font(font.Name, CSng(font.Size), GetFontStyle(font))
        End Sub

        Private Function GetFontStyle(ByVal font As SpreadsheetFont) As FontStyle
            Dim result As FontStyle = FontStyle.Regular
            If font.Bold Then
                result = result Or FontStyle.Bold
            End If
            If font.Italic Then
                result = result Or FontStyle.Italic
            End If
            Return result
        End Function

        Private Sub spreadsheetControl1_CellBeginEdit(ByVal sender As Object, ByVal e As DevExpress.XtraSpreadsheet.SpreadsheetCellCancelEventArgs)
            ' Disable editing of the "Category" column's cells to prevent end-users from entering wrong values. 
            If CheckCondition() Then
                e.Cancel = True
            End If
        End Sub

        Private Function CreateComboBox() As ComboBox
            Dim cmbBox As New ComboBox()
            cmbBox.DropDownStyle = ComboBoxStyle.DropDownList
            cmbBox.Items.AddRange(categories)
            cmbBox.Visible = False

            cmbBox.Parent = spreadsheetControl1
            spreadsheetControl1.Controls.Add(cmbBox)

            AddHandler cmbBox.SelectedValueChanged, AddressOf comboBox_SelectedValueChanged
            Return cmbBox
        End Function

        Private Sub comboBox_SelectedValueChanged(ByVal sender As Object, ByVal e As EventArgs)
            Dim selectedItem As Object = comboBox.SelectedItem
            If selectedItem Is Nothing Then
                Return
            End If

            ' Assign a value of the currently selected item in the ComboBox to the active cell.  
            activeCell.Value = selectedItem.ToString()
        End Sub
    End Class
End Namespace
