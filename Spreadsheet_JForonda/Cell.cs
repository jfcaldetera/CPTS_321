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

namespace SpreadsheetEngine
{
    public abstract class Cell : INotifyPropertyChanged
    {
        // cell values
        private readonly int mRow = 0;
        private readonly int mCol = 0;

        protected string mText = "";
        protected string mVal = "";
        protected string mName = "";

        protected int BGColor = -1;

        public List<Cell> references = new List<Cell>();
        public List<Cell> referencedBy = new List<Cell>();          

        public event PropertyChangedEventHandler PropertyChanged;

        // cell constructor
        public Cell()
        {
        }

        // cell constructor that takes in row and col number
        public Cell(int row, int col)
        {
            mRow = row;
            mCol = col;
            mName += Convert.ToChar('A' + col);
            mName += (row + 1).ToString();
        }

   
        // row getter
        public int rowIndex
        {
            get { return mRow; }
        }

        // col getter
        public int colIndex
        {
            get { return mCol; }
        }

        // text setter and getter
        public string Text
        {
            get { return mText; }

            set
            {
                // if the value is not the same 
                // invoke property change
                if (value != mText)
                {
                    mText = value;
                    PropertyChanged(this, new PropertyChangedEventArgs("Text"));
                }
                
            }
        }

        // value getter
        public string Value
        {
            get { return mVal; }
        }

        public string Name
        {
            get { return mName; }
        }

        public int BackGround
        {
            get { return BGColor; }
            set
            {
                if (value != BGColor)
                {
                    BGColor = value;
                    PropertyChanged(this, new PropertyChangedEventArgs("BackColor"));
                }
            }
        }

        public void Clear()
        {
            mText = "";
            BGColor = -1;
        }
    }


    
}
