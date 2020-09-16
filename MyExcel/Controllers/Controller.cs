using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using MyExcel.Models;

namespace MyExcel.Controllers
{
    class Controller
    {
        public List<ViewDataSource> viewDatas = new List<ViewDataSource>();
        public List<ProgDataSource> progDatas = new List<ProgDataSource>();
        public Dictionary<string, string> links = new Dictionary<string, string>();
        public List<KeyValuePair<string, string>> links1 = new List<KeyValuePair<string, string>>();
        public bool IsUpdate { get; set; }

        public Controller()
        {
            for (int i = 0; i < 23; i++)
            {
                viewDatas.Add(new ViewDataSource("", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""));
                progDatas.Add(new ProgDataSource("", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""));
            }
            IsUpdate = false;
        }
        public void Evaluate(string expression, int row, int col, string colStr)
        {
            expression.Replace(".", ",");
            int resI;
            bool isInt = Int32.TryParse(expression, out resI);
            double resD;
            bool isDouble = Double.TryParse(expression, out resD);
            if (expression == "" || isInt == true || isDouble == true)
            {
                viewDatas[row][col, row + 1] = expression;
                progDatas[row][col] = expression;
                foreach (KeyValuePair<string, string> kvp in links1)
                {
                    if (kvp.Key == colStr + (row + 1).ToString())
                    {
                        string[] addressTemp = kvp.Value.Split();
                        int rowTemp = Convert.ToInt32(addressTemp[0]);
                        int colTemp = Convert.ToInt32(addressTemp[1]);
                        Evaluate(progDatas[rowTemp][colTemp], rowTemp, colTemp, "");
                        IsUpdate = true;
                    }
                }
            }
            else if (expression[0] == '\'')
            {
                viewDatas[row][col, row + 1] = expression.Remove(0, 1);
                progDatas[row][col] = expression;
            }
            else if (expression[0] == '=')
            {
                expression = expression.Remove(0, 1);
                expression = expression.Replace(" ", "");
                try
                {
                    int errorCounter = Regex.Matches(expression, @"[a-zA-Z]").Count;
                    if (errorCounter == 0)
                    {
                        using (DataTable eval = new DataTable())
                        {
                            viewDatas[row][col, row + 1] = eval.Compute(expression, null).ToString();
                            progDatas[row][col] = "=" + expression;
                        }
                    }
                    else
                    {
                        string[] st = expression.Split(new char[] { '+', '-', '*', '/' });
                        string x;
                        int y;
                        string expression1 = expression;
                        int perev;
                        for (int i = 0; i < st.Length; i++)
                        {
                            if (Int32.TryParse(st[i], out perev) == true)
                            {
                                continue;
                            }
                            if (colStr != "")
                                x = GetAddressItem(st[i], row.ToString() + " " + col.ToString());
                            else
                                x = GetAddressItem(st[i]);
                            if ((y = Convert.ToInt32(x)) == -1)
                                throw new Exception(st[i]);
                            expression = expression.Replace(st[i], y.ToString());
                        }
                        using (DataTable eval = new DataTable())
                        {
                            viewDatas[row][col, row + 1] = eval.Compute(expression, null).ToString();
                            progDatas[row][col] = "=" + expression1;
                        }
                    }
                }
                catch (Exception e)
                {
                    viewDatas[row][col, row + 1] = $"#ошибка: {e.Message}";
                    progDatas[row][col] = e.Message;
                }


                //using (DataTable eval = new DataTable())
                //{
                //    return eval.Compute(expression, null).ToString();
                //}
            }
            //return "#oшибка";
        }

        public string GetProgItem(int row, int col)
        {
            return progDatas[row][col];
        }

        public string GetAddressItem(string str, string AddAddressToLinks = "")
        {
            int resI;
            string strtemp;
            strtemp = str;
            bool isInt = Int32.TryParse(str.Remove(0, 1), out resI);
            try
            {
                str = viewDatas[resI - 1][str[0]];
                isInt = Int32.TryParse(str, out resI);
                if (isInt != true)
                    throw new Exception("Указана неверная строка");
            }
            catch (Exception e)
            {
                str = e.Message;
                return e.Message;
            }
            if (AddAddressToLinks != "")
            {
                //links.Add(strtemp/*str[0] + resI.ToString()*/, AddAddressToLinks);
                links1.Add(new KeyValuePair<string, string>(strtemp, AddAddressToLinks));
            }
            return str;
        }

        public void Upload()
        {
            links1.Clear();
            viewDatas[0][0, 1] = "12";
            progDatas[0][0] = "12";

            viewDatas[1][0, 2] = "9.6";
            progDatas[1][0] = "=A1+B1*C1/5";
            links1.Add(new KeyValuePair<string, string>("A1", "1 0"));
            links1.Add(new KeyValuePair<string, string>("B1", "1 0"));
            links1.Add(new KeyValuePair<string, string>("C1", "1 0"));

            viewDatas[2][0, 3] = "Test";
            progDatas[2][0] = "'Test";


            viewDatas[0][1, 1] = "-4";
            progDatas[0][1] = "=C2";
            links1.Add(new KeyValuePair<string, string>("C2", "0 1"));

            viewDatas[1][1, 2] = "-38.4";
            progDatas[1][1] = "=A2*B1";
            links1.Add(new KeyValuePair<string, string>("A2", "1 1"));
            links1.Add(new KeyValuePair<string, string>("B1", "1 1"));

            viewDatas[2][1, 3] = "1";
            progDatas[2][1] = "=4-3";


            viewDatas[0][2, 1] = "3";
            progDatas[0][2] = "3";

            viewDatas[1][2, 2] = "-4";
            progDatas[1][2] = "=B3-C3";
            links1.Add(new KeyValuePair<string, string>("B3", "1 2"));
            links1.Add(new KeyValuePair<string, string>("C3", "1 2"));

            viewDatas[2][2, 3] = "5";
            progDatas[2][2] = "5";


            viewDatas[0][3, 1] = "Sample";
            progDatas[0][3] = "'Sample";

            viewDatas[1][3, 2] = "Spread";
            progDatas[1][3] = "'Spread";

            viewDatas[2][3, 3] = "Sheet";
            progDatas[2][3] = "'Sheet";
        }
    }
}
