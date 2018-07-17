using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DynamicsCRMCustomizationToolForExcel.Model
{
    public class ExcelMatrix
    {
        private string[][] excelMatrix;
        private int _numberofElements;

        public int numberofElements
        {
            get { return _numberofElements; }
        }

        public int rows
        {
            get { return excelMatrix.Length; }
        }


        public ExcelMatrix(int rows, int columns)
        {
            excelMatrix = new string[rows][];
            for (int i = 0; i < rows; i++)
            {
                excelMatrix[i] = new string[columns];
            }
            _numberofElements = 0;
        }

        public void AddRow(int row, string [] value)
        {
            excelMatrix[row]= value;
            _numberofElements++;
        }

        public string getElement(int row, int column)
        {
            return excelMatrix[row][column] ;
        }

        public string[] getRow(int row)
        {
            return excelMatrix[row] ;
        }
    }
}
