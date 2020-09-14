using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MyExcel.Controllers;

namespace MyExcel.Models
{
    class ViewDataSource
    {
        public ViewDataSource(string a, string b, string c, string d, string e, string f, string g, string h, string i, string j, string k, string l, string m, string n, string o, string p, string q)
        {
            A = a;
            B = b;
            C = c;
            D = d;
            E = e;
            F = f;
            G = g;
            H = h;
            I = i;
            J = j;
            K = k;
            L = l;
            M = m;
            N = n;
            O = o;
            P = p;
            Q = q;
        }

        public string A { get; set; }
        public string B { get; set; }
        public string C { get; set; }
        public string D { get; set; }
        public string E { get; set; }
        public string F { get; set; }
        public string G { get; set; }
        public string H { get; set; }
        public string I { get; set; }
        public string J { get; set; }
        public string K { get; set; }
        public string L { get; set; }
        public string M { get; set; }
        public string N { get; set; }
        public string O { get; set; }
        public string P { get; set; }
        public string Q { get; set; }

        public string this[int index, int indexTemp]
        {
            get
            {
                switch (index)
                {
                    case 0: return A;
                    case 1: return B;
                    case 2: return C;
                    case 3: return D;
                    case 4: return E;
                    case 5: return F;
                    case 6: return G;
                    case 7: return H;
                    case 8: return I;
                    case 9: return J;
                    case 10: return K;
                    case 11: return L;
                    case 12: return M;
                    case 13: return N;
                    case 14: return O;
                    case 15: return P;
                    case 16: return Q;
                }
                return "-1";
            }

            set
            {
                switch (index)
                {
                    case 0: A = value;
                        DBViewConnect.UpdateToViewDB(indexTemp.ToString(), "A", value);
                        break;
                    case 1: B = value;
                        DBViewConnect.UpdateToViewDB(indexTemp.ToString(), "B", value);
                        break;
                    case 2: C = value;
                        DBViewConnect.UpdateToViewDB(indexTemp.ToString(), "C", value);
                        break;
                    case 3: D = value;
                        DBViewConnect.UpdateToViewDB(indexTemp.ToString(), "D", value);
                        break;
                    case 4: E = value;
                        DBViewConnect.UpdateToViewDB(indexTemp.ToString(), "E", value);
                        break;
                    case 5: F = value;
                        DBViewConnect.UpdateToViewDB(indexTemp.ToString(), "F", value);
                        break;
                    case 6: G = value;
                        DBViewConnect.UpdateToViewDB(indexTemp.ToString(), "G", value);
                        break;
                    case 7: H = value;
                        DBViewConnect.UpdateToViewDB(indexTemp.ToString(), "H", value);
                        break; 
                    case 8: I = value;
                        DBViewConnect.UpdateToViewDB(indexTemp.ToString(), "I", value);
                        break;
                    case 9: J = value;
                        DBViewConnect.UpdateToViewDB(indexTemp.ToString(), "J", value);
                        break;
                    case 10: K = value;
                        DBViewConnect.UpdateToViewDB(indexTemp.ToString(), "K", value);
                        break;
                    case 11: L = value;
                        DBViewConnect.UpdateToViewDB(indexTemp.ToString(), "L", value);
                        break;
                    case 12: M = value;
                        DBViewConnect.UpdateToViewDB(indexTemp.ToString(), "M", value);
                        break;
                    case 13: N = value;
                        DBViewConnect.UpdateToViewDB(indexTemp.ToString(), "N", value);
                        break;
                    case 14: O = value;
                        DBViewConnect.UpdateToViewDB(indexTemp.ToString(), "O", value);
                        break;
                    case 15: P = value;
                        DBViewConnect.UpdateToViewDB(indexTemp.ToString(), "P", value);
                        break;
                    case 16: Q = value;
                        DBViewConnect.UpdateToViewDB(indexTemp.ToString(), "Q", value);
                        break;
                }
            }
        }

        public string this[char address]
        {
            get
            {
                switch (address)
                {
                    case 'A': return A;
                    case 'B': return B;
                    case 'C': return C;
                    case 'D': return D;
                    case 'E': return E;
                    case 'F': return F;
                    case 'G': return G;
                    case 'H': return H;
                    case 'I': return I;
                    case 'J': return J;
                    case 'K': return K;
                    case 'L': return L;
                    case 'M': return M;
                    case 'N': return N;
                    case 'O': return O;
                    case 'P': return P;
                    case 'Q': return Q;
                }
                return "-1";
            }

            /*set
            {
                switch (address)
                {
                    case 'A': A = value; break;
                    case 'B': B = value; break;
                    case 'C': C = value; break;
                    case 'D': D = value; break;
                    case 'E': E = value; break;
                    case 'F': F = value; break;
                    case 'G': G = value; break;
                    case 'H': H = value; break;
                    case 'I': I = value; break;
                    case 'J': J = value; break;
                    case 'K': K = value; break;
                    case 'L': L = value; break;
                    case 'M': M = value; break;
                    case 'N': N = value; break;
                    case 'O': O = value; break;
                    case 'P': P = value; break;
                    case 'Q': Q = value; break;
                }
            }*/
        }
    }
}
