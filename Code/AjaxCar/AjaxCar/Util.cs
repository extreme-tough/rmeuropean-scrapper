using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Text;
using System.IO;
using System.Windows.Forms;

namespace FeedGen
{
    public static class Util
    {
        public static DataTable csvToDataTable(string file, bool isRowOneHeader)
        {

            DataTable csvDataTable = new DataTable();

            //no try/catch - add these in yourselfs or let exception happen
            String[] csvData = File.ReadAllLines(file);

            //if no data in file ‘manually’ throw an exception
            if (csvData.Length == 0)
            {
                throw new Exception("CSV File Appears to be Empty");
            }

            String[] headings = csvData[0].Split(',');
            int index = 0; //will be zero or one depending on isRowOneHeader

            csvDataTable.Columns.Add("SK", typeof(int));

            if(isRowOneHeader) //if first record lists headers
            {
                index = 1; //so we won’t take headings as data

                //for each heading
                for(int i = 0; i < headings.Length; i++)
                {
                    //replace spaces with underscores for column names
                    headings[i] = headings[i].Replace(" ", "_");

                    //add a column for each heading
                    csvDataTable.Columns.Add(headings[i], typeof(string));
                } 
            }
            else //if no headers just go for col1, col2 etc.
            {
                for (int i = 0; i < headings.Length; i++)
                {
                   //create arbitary column names
                   csvDataTable.Columns.Add("col"+(i+1).ToString(), typeof(string));
                }
            }

            //populate the DataTable
            for (int i = index; i < csvData.Length; i++)
            {
                //create new rows
                DataRow row = csvDataTable.NewRow();

                for (int j = -1; j < headings.Length; j++)
                {
                    if (j == 0)
                    {
                        if (csvData[i].Split(',')[0] == "")
                        {
                            //Check if master category is empty
                            MessageBox.Show("Master_Category empty in the map file. This is not valid. Please fix and try again");
                            return null;
                        }
                    }
                     //fill them
                    if (j == -1)
                        row[j+1] = i  + csvData.Length + 1;
                    else
                        row[j+1] = csvData[i].Split(',')[j];
                }

                //add rows to over DataTable
                csvDataTable.Rows.Add(row);
            }

            //return the CSV DataTable
            return csvDataTable;

        } 
    }
}
