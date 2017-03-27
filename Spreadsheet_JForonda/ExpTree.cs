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
    class ExpTree
    {
        static string mExp = "A1+B1+C1";
        static ExpTree tree;
        static Node mRoot;
        public static Dictionary<string, double> mVars = new Dictionary<string, double>();
        public readonly static char[] mOperators = { '+', '-', '*', '/' };
        static bool invalidInput = false;



        public abstract class Node
        {
            //public Node Left;
            //public Node Right;
            public abstract double Eval();
        }

        public class ConstNode : Node
        {
            public double mConst;

            public ConstNode(double constant)
            {
                mConst = constant;
            }
            public override double Eval()
            {
                return mConst;
            }
        }

        public class VarNode : Node
        {
            public string mVar;

            public VarNode(string name)
            {
                this.mVar = name;
            }
            public override double Eval()
            {

                if (mVars.ContainsKey(mVar))
                {
                    return mVars[mVar];
                }

                else
                {
                    mVars.Add(mVar, 0);
                    return mVars[mVar];
                }

                return 0.0;
            }

        }

        public class OpNode : Node
        {
            public char mOp;
            private Node mLeft;
            private Node mRight;

            public OpNode(char op, Node lChild, Node rChild)
            {
                mOp = op;
                mLeft = lChild;
                mRight = rChild;
            }
            public override double Eval()
            {
                double left = mLeft.Eval();
                double right = mRight.Eval();

                if (mOp == null)
                {
                    return -1;
                }
                if (mOp != null)
                {
                    switch (mOp)
                    {
                        case '+':
                            return left + right;
                        case '-':
                            return left - right;
                        case '*':
                            return left * right;
                        case '/':
                            return left / right;

                    }
                }
                return 0.0;
            }

        }

        public ExpTree()
        {
            mVars = new Dictionary<string, double>();
        }

        public ExpTree(string expression)
        {
            mExp = expression.Replace(" ", "");
            mRoot = Compile(mExp);
        }

        private static Node Compile(string exp)
        {
            int pCount = 0;
            char[] ops = ExpTree.mOperators;

            // if incoming string is empty
            if ((exp == "") || (exp == null))
            { return null; }

            // if we get find left parethesis, we need to search for the right
            if (exp[0] == '(')
            {
                // loop through the string
                for (int i = 0; i < exp.Length; i++)
                {
                    // if we encounter left parenthesis
                    // increment parenthesis counter 
                    if (exp[i] == '(')
                    { pCount++; }

                    // if we hit associated right parenthesis
                    // decrement counter
                    else if (exp[i] == ')')
                    {
                        pCount--;

                        // check if counter is = 0
                        // if not = 0, parenthesis is unmatched
                        if (pCount == 0)
                        {
                            // when we're done, break out 
                            if(exp.Length - 1 != i)
                            { break; }

                            // else recursively evaluate substring
                            else
                            { return Compile(exp.Substring(1, exp.Length - 2)); }
                        }
                    }
                }
            }

            // evaluate by precedence
            foreach (char op in ops)
            {
                Node tNode = Compile(exp, op);
                {
                    if (invalidInput == true)
                    { break; }

                    if (tNode != null)
                    {  return tNode; }
                }
            }

            
            double constDouble;

            if (double.TryParse(exp, out constDouble))
            { return new ConstNode(constDouble); }

            // else variable is found in dictionary
            // its varNode
            else
            {
                mVars[exp] = 0;
                return new VarNode(exp);
            }

            int index = FindOp(exp), constDoubleValue;

            // at this pointer, operator isnt found
            // try it as a constNode
            if (index == -1)
            {
                if (int.TryParse(exp, out constDoubleValue))
                { return new ConstNode(constDoubleValue); }

                // if neither, its a varNode
                else
                { return new VarNode(exp); }
            }

            // Left and right subtrees
            Node Left = Compile(exp.Substring(0, index));
            Node Right = Compile(exp.Substring(index + 1));

            return new OpNode(exp[index], Left, Right);
            
        }

        static Node Compile(string exp, char op)
        {
            int pCount = 0;
            bool terminate = false;
            bool rightParenth = false;

            // default to the left mort
            int i = exp.Length - 1;

            while (!terminate)
            {
                // if left parenthesis
                if (exp[i] == '(')
                {
                    // if right associated, decrement count
                    if (rightParenth)
                    {  pCount--; }

                    // if left associated, increment count
                    else
                    { pCount++; }
                }

                // if right parenthesis
                else if (exp[i] == ')')
                {
                    // if right associated, increment count
                    if (rightParenth)
                    { pCount++; }

                    // if left, decrement count
                    else
                    { pCount--; }
                }

                // the case wehre its not in parenthesis
                // reset root to current op and evaluate
                if (pCount == 0 && exp[i] == op)
                { return new OpNode(op, Compile(exp.Substring(0, i)), Compile(exp.Substring(i + 1)));  }

                // if right associated
                if (rightParenth)
                {
                    // we have reached the end
                    if (i == exp.Length - 1)
                    {  terminate = true; }

                    i++;
                }

                // else left associated 
                else
                {
                    // reached end
                    if (i == 0)
                    { terminate = true; }

                    i--;
                }
            }

            // if pCount != 0 at this point, 
            // parenthsis did not match up 
            if (pCount != 0)
            {
                Console.WriteLine("Error - Invalid Expression Unbalanced Expression");
                invalidInput = true;
            }
            return null;
        }

        // find next op in expression
        static int FindOp(string exp)
        {
            for (int i = 0; i < exp.Length; i++)
            {
                // reverse order of operations
                if ((exp[i] == '*') || (exp[i] == '/') || (exp[i] == '+') || (exp[i] == '-'))
                {
                    return i;
                }
                
            }
            return -1;
        }

        public double Eval()
        {
            if (mRoot != null)
            {
                return mRoot.Eval();
            }
            else
            {
                return double.NaN;
            }
        }
        
        public void SetVar(string varName, double varVariable)
        { 
            mVars[varName] = varVariable;
        }
        
        public string[] getVars()
        {
            return mVars.Keys.ToArray();
        }

        static public void Menu()
        {
            Console.WriteLine("= = = = = = = = = = = = = = = = = = = = =");
            Console.WriteLine(" Menu (current expression: " + mExp + ")");
            Console.WriteLine("    1 - Enter a new expression");
            Console.WriteLine("    2 - Set a variable value");
            Console.WriteLine("    3 - Evaluate tree");
            Console.WriteLine("    4 - Quit");
        }
    }
}
