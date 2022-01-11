using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Создатель_расчетов
{
    class Errors_checker
    {
        public string textBox_checker(string textbox)
        {
            char[] temp = textbox.ToCharArray();
            for (int index = 0; index < temp.Length; index++)
            {
                if (temp[index] == '.')
                    temp[index] = ',';
            }
            string output = new string(temp);
            return output;
        }

    }
}
