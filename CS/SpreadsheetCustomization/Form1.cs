using DevExpress.Spreadsheet;
using System;
using System.Drawing;
using System.Windows.Forms;

namespace SpreadsheetCustomization
{
    public partial class Form1 : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public static string[] categories = { "Meat/Poultry", "Condiments", "Seafood", "Dairy Products", "Grains/Cereals", "Beverages", "Confections" };
        IWorkbook workbook;
        Worksheet worksheet;
        CellRange categoryColumn;
        Cell activeCell;
        ComboBox comboBox;

        public Form1()
        {
            InitializeComponent();
            workbook = spreadsheetControl1.Document;
            workbook.LoadDocument("Documents\\Document.xlsx", DocumentFormat.Xlsx);

            worksheet = workbook.Worksheets["Sales report"];
            categoryColumn = worksheet["Table[Category]"];

            // Create a ComboBox and specify its settings.
            comboBox = CreateComboBox();

            // Specify the SpreadsheetControl's options.
            spreadsheetControl1.Options.Behavior.Selection.AllowExtendSelection = false;
            spreadsheetControl1.Options.VerticalScrollbar.Visibility = DevExpress.XtraSpreadsheet.SpreadsheetScrollbarVisibility.Hidden;
            spreadsheetControl1.Options.HorizontalScrollbar.Visibility = DevExpress.XtraSpreadsheet.SpreadsheetScrollbarVisibility.Hidden;

            spreadsheetControl1.SelectionChanged += spreadsheetControl1_SelectionChanged;
            spreadsheetControl1.CellBeginEdit +=spreadsheetControl1_CellBeginEdit;
            spreadsheetControl1.MouseWheel += spreadsheetControl1_MouseWheel;
        }

        void spreadsheetControl1_MouseWheel(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            UpdateComboBox();
        }

        #region #displaycombobox
        void spreadsheetControl1_SelectionChanged(object sender, EventArgs e)
        {
            UpdateComboBox();
        }

        void UpdateComboBox()
        {
            // Access the active cell.
            activeCell = spreadsheetControl1.ActiveCell;

            // If the currently selected cell is not in the "Category" column of the worksheet table, return. 
            if (!CheckCondition())
            {
                comboBox.Visible = false;
                return;
            }
            
            // Otherwise, obtain the bounds of the active cell and display the ComboBox control over it. 
            Rectangle cellRect = spreadsheetControl1.GetCellBounds(activeCell.RowIndex, activeCell.ColumnIndex);
            if (cellRect.IsEmpty)
                comboBox.Visible = false;
            else
            {
                comboBox.Bounds = cellRect;
                comboBox.Visible = true;
                comboBox.SelectedItem = activeCell.Value.TextValue;

                UpdateComboBoxAppearance(activeCell);
            }
        }

        bool CheckCondition()
        {
            return worksheet == workbook.Worksheets.ActiveWorksheet ? categoryColumn.IsIntersecting(activeCell) : false;
        }

        #endregion #displaycombobox

        void UpdateComboBoxAppearance(Cell source)
        {
            comboBox.BackColor = source.Fill.BackgroundColor;
            SpreadsheetFont font = source.Font;
            comboBox.ForeColor = font.Color;
            comboBox.Font = new Font(font.Name, (float)font.Size, GetFontStyle(font));
        }

        FontStyle GetFontStyle(SpreadsheetFont font)
        {
            FontStyle result = FontStyle.Regular;
            if (font.Bold)
                result |= FontStyle.Bold;
            if (font.Italic)
                result |= FontStyle.Italic;
            return result;
        }

        void spreadsheetControl1_CellBeginEdit(object sender, DevExpress.XtraSpreadsheet.SpreadsheetCellCancelEventArgs e)
        {
            // Disable editing of the "Category" column's cells to prevent end-users from entering wrong values. 
            if (CheckCondition())
            {
                e.Cancel = true;
            }
        }

        ComboBox CreateComboBox()
        {
            ComboBox cmbBox = new ComboBox();
            cmbBox.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbBox.Items.AddRange(categories);
            cmbBox.Visible = false;

            cmbBox.Parent = spreadsheetControl1;
            spreadsheetControl1.Controls.Add(cmbBox);

            cmbBox.SelectedValueChanged += comboBox_SelectedValueChanged;
            return cmbBox;
        }

        private void comboBox_SelectedValueChanged(object sender, EventArgs e)
        {
            object selectedItem = comboBox.SelectedItem;
            if (selectedItem == null)
                return;
            
            // Assign a value of the currently selected item in the ComboBox to the active cell.  
            activeCell.Value = selectedItem.ToString();
        }
    }
}
