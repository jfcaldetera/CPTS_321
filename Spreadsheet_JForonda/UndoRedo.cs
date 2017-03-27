/*Jhenna Foronda, 11423409
 * March 23, 2017
 * Spreadsheet Application -- Save and Load
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace SpreadsheetEngine
{
    // interface
    public interface UndoRedoCwd
    {
        UndoRedoCwd Restore(Spreadsheet sheet);
    }

    // collection of commands
    public class UndoRedoCollection
    {
        public string mDescription;
        private UndoRedoCwd[] mCwds;
        
        public UndoRedoCollection()
        {

        }

        public UndoRedoCollection(UndoRedoCwd[] cwds, string description)
        {
            mCwds = cwds;
            mDescription = description;
        }

        public UndoRedoCollection(List<UndoRedoCwd> cwds, string description)
        {
            mCwds = cwds.ToArray();
            mDescription = description;
        }

        public UndoRedoCollection Restore(Spreadsheet sheet)
        {
            List<UndoRedoCwd> cwdList = new List<UndoRedoCwd>();

            foreach (UndoRedoCwd cwd in mCwds)
            { cwdList.Add(cwd.Restore(sheet)); }

            return new UndoRedoCollection(cwdList.ToArray(), this.mDescription);
        }
    }

    public class UndoRedoClass
    {
        private Stack<UndoRedoCollection> mUndos = new Stack<UndoRedoCollection>();
        private Stack<UndoRedoCollection> mRedos = new Stack<UndoRedoCollection>();

        // check if undo stack is empty
        public bool checkUndo
        {
            get { return mUndos.Count != 0; }
        }

        // check if redo stack is empty
        public bool checkRedo
        {
            get { return mRedos.Count != 0; }
        }

        // if there's an undp
        // get its description
        public string UndoDescrptn
        {
            get
            {
                if (checkUndo)
                {  return mUndos.Peek().mDescription; }

                return "";
            }
        }

        // if there's a redo
        // get its description
        public string RedoDescrptn
        {
            get
            {
                if (checkRedo)
                { return mRedos.Peek().mDescription; }

                return "";
            }
        }

        // add undo action to undo stack
        public void addUndo(UndoRedoCollection undos)
        {
            mUndos.Push(undos);
            mRedos.Clear();
        }
        
        // pops undo action and does it 
        public void Undo(Spreadsheet sheet)
        {
            UndoRedoCollection action = mUndos.Pop();
            mRedos.Push(action.Restore(sheet));
        }

        // ops  redo action and does it 
        public void Redo(Spreadsheet sheet)
        {
            UndoRedoCollection action = mRedos.Pop();
            mUndos.Push(action.Restore(sheet));
        }

        public void Clear()
        {
            mUndos.Clear();
            mRedos.Clear();
        }
    }

    public class RestoreText : UndoRedoCwd
    {
        private string mText;
        private string mCellName;

        // text getter
        public RestoreText(string text, string name)
        {
            mText = text;
            mCellName = name;
        }

        // restore cell with OG text
        public UndoRedoCwd Restore(Spreadsheet sheet)
        {
            Cell cell = sheet.GetCell(mCellName);
            string OGtext = cell.Text;

            cell.Text = mText;

            return new RestoreText(OGtext, mCellName);
        }
    }

    public class RestoreColor: UndoRedoCwd
    {
        private int mBGcolor;
        private string mCellName;

        // color getter
        public RestoreColor(int color, string name)
        {
            mBGcolor = color;
            mCellName = name;
        }

        // restore cell to OG background color
        public UndoRedoCwd Restore (Spreadsheet sheet)
        {
            Cell cell = sheet.GetCell(mCellName);
            int OGcolor = cell.BackGround;

            cell.BackGround = mBGcolor;
            
            return new RestoreColor(OGcolor, mCellName);
        }
    }
}
