/*Jhenna Foronda, 11423409
 * March 23, 2017
 * Spreadsheet Application -- Save and Load
 */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Xml;
using System.Xml.Linq;
using System.IO;


namespace SpreadsheetEngine 
{
    public class Spreadsheet
    {
        public Cell[,] cells;
        public event PropertyChangedEventHandler CellPropertyChanged;
        private Dictionary<string, HashSet<string>> depDict = new Dictionary<string, HashSet<string>>();

        private class InstanciateCell : Cell
        {
            // create a cell that inherits row and col value
            public InstanciateCell(int row, int col) : base (row, col)
            {

            }

            // value setter
            public void SetVal(string value)
            {
                mVal = value;
            }
        }

        // spreadsheet constructor
        public Spreadsheet(int rows, int cols)
        {
            cells = new Cell[rows, cols];
            

            // loop through and create instantiation of cell 
            for (int i = 0; i < rows; i ++)
            {
                for (int j = 0; j < cols; j++)
                {
                    cells[i, j] = new InstanciateCell(i, j);
                    cells[i,j].PropertyChanged += OnPropertyChanged;
                }
            }
        }

        // cell getter
        public Cell GetCell(int row, int col)

        {
            return cells[row, col];
        }

        public Cell GetCell(string str)
        {
            char letter = str[0];
            Int16 num;
            Cell result;

            if (!Char.IsLetter(letter))
            {
                return null;
            }
            if (!Int16.TryParse(str.Substring(1), out num))
            {
                return null;
            }
            try
            {
                result = GetCell(num - 1, letter - 'A');
            }
            catch (Exception C)
            {
                return null;
            }
            return result;
        }

        // row count getter
        public int RowCount
        {
            get { return cells.GetLength(0); }
        }

        // col count getter
        public int ColCount
        {
            get { return cells.GetLength(1); }
        }

        public void OnPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            // if property name is equal to text
            if (e.PropertyName == "Text")
            {
                InstanciateCell tCell = sender as InstanciateCell;
                DeleteDeps(tCell.Name);
                

                // update cell value to text
                //tCell.SetVal(tCell.Text);

                // if it starts with an equal sign
                if (tCell.Text != "" && tCell.Text[0] == '=' && tCell.Text.Length > 1)
                {
                    ExpTree exp = new ExpTree(tCell.Text.Substring(1));
                    MakeDeps(tCell.Name, exp.getVars());
                }
                EvalCell(sender as Cell);
            }
            else if (e.PropertyName == "BackColor")
            {
                CellPropertyChanged(sender, new PropertyChangedEventArgs("BackColor"));
            }
            
            
        }

        private void EvalCell(Cell cell)
        {
            InstanciateCell mCell = cell as InstanciateCell;
            
            // if the string is empty
            if (string.IsNullOrEmpty(mCell.Text))
            {
                mCell.SetVal("");
                CellPropertyChanged(cell, new PropertyChangedEventArgs("Value"));
            }
            // if it an equation
            else if (mCell.Text[0] == '=' && mCell.Text.Length > 1)
            {
                // remove = 
                string exp = mCell.Text.Substring(1);
                SpreadsheetEngine.ExpTree mExp = new SpreadsheetEngine.ExpTree(exp);
                string[] arrayVars = mExp.getVars();

                foreach (string v in arrayVars)
                {
                    if (GetCell(v) == null)
                    {
                        mCell.SetVal("Bad Ref");
                        CellPropertyChanged(cell, new PropertyChangedEventArgs("Value"));

                        break;
                    }
                    // attempt at setting variable to value
                    SetExpVar(mExp, v);
                }
                mCell.SetVal(mExp.Eval().ToString());
                CellPropertyChanged(cell, new PropertyChangedEventArgs("Value"));
            }
            // not an expression
            else
            {
                mCell.SetVal(mCell.Text);
                CellPropertyChanged(cell, new PropertyChangedEventArgs("Value"));
            }

            // evaluate dependencies
            if (depDict.ContainsKey(mCell.Name))
            {
                foreach (string name in depDict[mCell.Name])
                {
                    Eval(name);
                }
            }
            
        }

        private void Eval(string str)
        {
            EvalCell(GetCell(str));
        }

        private void MakeDeps(string cellName, string[] varsUsed)
        {
            foreach (string vName in varsUsed)
            {
                if (!depDict.ContainsKey(vName))
                {
                    // create dictionary for this variable name
                    depDict[vName] = new HashSet<string>();
                }
                // add cell name to dependencies 
                depDict[vName].Add(cellName);
            }
        }

        private void DeleteDeps(string cellName)
        {
            List<string> dep = new List<string>();

            // add to the list, at every cell name conating the key
            foreach (string s in depDict.Keys)
            {
                if (depDict[s].Contains(cellName))
                {
                    dep.Add(s);
                }
            }

            foreach(string s in dep)
            {
                // remove matching name with key
                HashSet<string> set = depDict[s];
                if (set.Contains(cellName))
                {
                    set.Remove(cellName);
                   
                }
            }
        }

        private void SetExpVar(ExpTree exp, string vName)
        {
            Cell vCell = GetCell(vName);
            double val;

            // case: empty string, set to 0
            if (string.IsNullOrEmpty(vCell.Value))
            {
                exp.SetVar(vCell.Name, 0);
            }
            // case: not value, setto 0
            else if (!double.TryParse(vCell.Value, out val))
            {
                exp.SetVar(vName, 0);
            }
            else
            {
                exp.SetVar(vName, val);
            }

        }

        public void Load(Stream infile)
        {
            XDocument readXML = XDocument.Load(infile);

            foreach (XElement tag in readXML.Root.Elements("cell"))
            {
                // get the cell with the same name value
                Cell mCell = GetCell(tag.Element("name").Value);

                // if text isnt default
                if (tag.Element("text") != null)
                {
                    mCell.Text = tag.Element("text").Value.ToString();
                }

                // if background color isnt default
                if (tag.Element("backgroundcolor") != null)
                {
                    mCell.BackGround = int.Parse(tag.Element("backgroundcolor").Value.ToString());
                }
            }
        }

        public void Save(Stream outfile)
        {
            XmlWriter writeXML = XmlWriter.Create(outfile);

            // starting element for spreadsheet
            writeXML.WriteStartElement("Spreadsheet");

            foreach(Cell mCell in cells)
            {
                // are the cells at default? 
                if (mCell.Text != "" || mCell.Value != "" || mCell.BackGround != -1)
                {
                    // starting element for cell
                    writeXML.WriteStartElement("cell");

                    // element strings
                    writeXML.WriteElementString("name", mCell.Name.ToString());
                    writeXML.WriteElementString("backgroundcolor", mCell.BackGround.ToString());
                    writeXML.WriteElementString("text", mCell.Text.ToString());

                    // ending element for cell
                    writeXML.WriteEndElement();
                }
            }
            // ending element for spreadsheet
            writeXML.WriteEndElement();
            // close XML writer
            writeXML.Close();
        }
    }

}
