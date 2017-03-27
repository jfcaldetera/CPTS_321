/*Jhenna Foronda, 11423409
 * March 23, 2017
 * Spreadsheet Application -- Save and Load
 */

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Xml;
using System.Xml.Linq;
using SpreadsheetEngine;


namespace Spreadsheet_JForonda
{
    public partial class SpreadsheetForm : Form
    {
        private Spreadsheet testSpreadsheet = new Spreadsheet(50, 26);
        public UndoRedoClass UndoRedo = new UndoRedoClass();

        public SpreadsheetForm()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // subscribe to events
            testSpreadsheet.CellPropertyChanged += OnCellPropertyChanged;
            dataGridView1.CellBeginEdit += dataGridView1_CellBeginEdit;
            dataGridView1.CellEndEdit += dataGridView1_CellEndEdit;

            // clear columns
            dataGridView1.Columns.Clear();

            // create columns A-Z
            for (char c = 'A'; c <= 'Z'; c++)
            {
                string name = Convert.ToString(c);
                dataGridView1.Columns.Add(name, name);
            }

            // create rows
            dataGridView1.Rows.Add(50);

            int rowVal = 1;

            // for evach row, set the row value
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                row.HeaderCell.Value = Convert.ToString(rowVal++);
            }

            for (int k = 0; k < testSpreadsheet.RowCount; k++)
            {
                for (int j = 0; j < testSpreadsheet.ColCount; j++)
                {
                    testSpreadsheet.GetCell(k, j).PropertyChanged += OnCellPropertyChanged;
                }
            }
            // automatically resize 
            dataGridView1.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);

        }

        void dataGridView1_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            int row = e.RowIndex;
            int col = e.ColumnIndex;

            Cell CurrentCell = testSpreadsheet.GetCell(row, col);

            // set the value of the cell to the text
            dataGridView1.Rows[row].Cells[col].Value = CurrentCell.Text;
        }

        void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            int row = e.RowIndex;
            int col = e.ColumnIndex;
            string NewText;
            UndoRedoCwd[] undoStack = new UndoRedoCwd[1];

            Cell CorrespondingCell = testSpreadsheet.GetCell(row, col);

            try
            {
                NewText = dataGridView1.Rows[row].Cells[col].Value.ToString();
            }
            catch (NullReferenceException)
            {
                NewText = "";
            }
            
            // restore the spreadsheet
            undoStack[0] = new RestoreText(CorrespondingCell.Text, CorrespondingCell.Name);
            // reset the text in the cell
            CorrespondingCell.Text = NewText;
            // add undo command with cell text change description
            UndoRedo.addUndo(new UndoRedoCollection(undoStack, "cell Text Change"));
            dataGridView1.Rows[row].Cells[col].Value = CorrespondingCell.Value;
            //update drop down menu 
            UpdateMenuText();

        }

        private void OnCellPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            Cell CurrentCell = sender as Cell;

            if (e.PropertyName == "Value" || CurrentCell != null)
            {
                dataGridView1.Rows[CurrentCell.rowIndex].Cells[CurrentCell.colIndex].Value = CurrentCell.Value;

            }

            if (e.PropertyName == "BackColor" && CurrentCell != null)
            {
                dataGridView1.Rows[CurrentCell.rowIndex].Cells[CurrentCell.colIndex].Style.BackColor = Color.FromArgb(CurrentCell.BackGround);
            }
        }


        private void Button1_Click(object sender, EventArgs e)
        {
            Random r = new Random();

            // randomly set 50 cells' text value to "Hello World"
            for (int i = 0; i < 50; i++)
            {
                int row = r.Next(0, 49);
                int col = r.Next(2, 25);
                
                
                testSpreadsheet.GetCell(row, col).Text = "Hello World!";
            }

            for (int j = 0; j < 50; j++)
            {

                testSpreadsheet.GetCell(j, 1).Text = ("This is cell B" + (j + 1).ToString());
            }

            for (int k = 0; k < 50; k++)
            {

                testSpreadsheet.GetCell(k, 0).Text = "This is cell B" + (k + 1).ToString();
            }

        }

        private void UpdateMenuText()
        {
            ToolStripMenuItem menuItems = menuStrip1.Items[1] as ToolStripMenuItem;

            foreach (ToolStripItem item in menuItems.DropDownItems)     
            {
                if (item.Text.Substring(0, 4) == "Undo")
                {
                    item.Enabled = UndoRedo.checkUndo;                   
                    item.Text = "Undo " + UndoRedo.UndoDescrptn;           
                }
                else if (item.Text.Substring(0, 4) == "Redo")
                {
                    item.Enabled = UndoRedo.checkRedo;                  
                    item.Text = "Redo " + UndoRedo.RedoDescrptn;            
                }
            }
        }


        private void undoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            UndoRedo.Undo(testSpreadsheet);
            UpdateMenuText();
        }

        private void redoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            UndoRedo.Redo(testSpreadsheet);
            UpdateMenuText();
        }

        private void cellToolStripMenuItem1_Click(object sender, EventArgs e)
        {

        }

        private void changeBackgrounColorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int colorSelected = 0;
            List<UndoRedoCwd> undos = new List<UndoRedoCwd>();
            ColorDialog colorDialog = new ColorDialog();

            // prompt user for colorchange
            if (colorDialog.ShowDialog() == DialogResult.OK)
            {
                colorSelected = colorDialog.Color.ToArgb();

                // change selected cells to selected color
                // add restore commands
                foreach (DataGridViewCell cell in dataGridView1.SelectedCells)
                {
                    Cell CurrentCell = testSpreadsheet.GetCell(cell.RowIndex, cell.ColumnIndex);
                    undos.Add(new RestoreColor(CurrentCell.BackGround, CurrentCell.Name));
                    CurrentCell.BackGround = colorSelected;
                }
                // add restore actions with cell background change description
                UndoRedo.addUndo(new UndoRedoCollection(undos, "cell background color change"));
                UpdateMenuText();
            }
        }

        // clears all cells edited in the spreadsheet
        public void Clear()
        {
            int mRowCount = testSpreadsheet.RowCount;
            int mColCount = testSpreadsheet.ColCount;

            for (int i = 0; i < mRowCount; i++)
            {
                for (int j = 0; j < mColCount; j++)
                {
                    // have the cells been edited? --if yes, clear contents of cell
                    if (testSpreadsheet.cells[i,j].Text != "" || testSpreadsheet.cells[i,j].Value != "" || testSpreadsheet.cells[i,j].BackGround != -1)
                    {
                        testSpreadsheet.cells[i, j].Clear();
                    }
                }
            }
        }
        
        private void loadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var openFileDialog = new OpenFileDialog();

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                // clear the spreadsheet
                Clear();
                // open stream to read
                Stream infilefStream = new FileStream(openFileDialog.FileName, FileMode.Open, FileAccess.Read);
                // load to spreadsheet
                testSpreadsheet.Load(infilefStream);
                // trash stream
                infilefStream.Dispose();
                // clear undo and redo stacks
                UndoRedo.Clear();
            }
            UpdateMenuText();
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var saveFileDialog = new SaveFileDialog();

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                // open stream to write
                Stream outfilefStream = new FileStream(saveFileDialog.FileName, FileMode.Create, FileAccess.Write);
                // save to spreadsheet
                testSpreadsheet.Save(outfilefStream);
                // trash stream
                outfilefStream.Dispose();
            }
        }
    }
}
