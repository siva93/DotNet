
namespace MyExcelValidation.Helper
{
    using System.Collections.Generic;
    using System.Linq;
    using ClosedXML.Excel;
    using System.Data;
    using WebGrease.Css.Extensions;

    public static class Helper
    {
        public static DataTable ConvertToDataTable(XLWorkbook workBook, string sheetName)
        {
            var dataTable = new DataTable("DataValidation");
            var validationTable = new DataTable("Validation");

            IXLWorksheet workSheet;
            workBook.TryGetWorksheet(sheetName, out workSheet);
            if (workSheet != null)
            {
                var firstRowUsed = workSheet.FirstRowUsed();
                var firstPossibleAddress = workSheet.Row(firstRowUsed.RowNumber()).FirstCell().Address;
                var lastPossibleAddress = workSheet.LastCellUsed().Address;

                // Get a range with the remainder of the worksheet data (the range used)
                var range = workSheet.Range(firstPossibleAddress, lastPossibleAddress).RangeUsed();
                var table = range.AsTable();
                var rows = table.DataRange.RowsUsed();
                var col = workSheet.Columns().ToList();

                workSheet.Columns().ForEach(a => dataTable.Columns.Add($"{a.ColumnLetter()}{a.ColumnNumber()}"));

                DataColumn isContainErrorColumn = new DataColumn("IsContainError", typeof(bool));
                DataColumn errorColumn = new DataColumn("Error", typeof(string));

                dataTable.Columns.Add(isContainErrorColumn);
                dataTable.Columns.Add(errorColumn);

                validationTable = dataTable.Clone();

                List<int> uniqueColumnsIndex = new List<int>();

                int primaryKeyColumnIndex = 0;

                uniqueColumnsIndex.Add(0);
                uniqueColumnsIndex.Add(1);
                uniqueColumnsIndex.Add(2);


                AddPrimaryKeyConstraint(dataTable, primaryKeyColumnIndex);

                AddUniqueConstraints(dataTable, uniqueColumnsIndex);
                dataTable.RowChanged += table_RowChanged;
                rows.ForEach(a =>
                {
                    var rowValue = a.Cells().Select(b => b.Value).ToList();
                    string error = string.Empty;
                    int errorColumnIndex = 0;
                    int errorMessageColumnIndex = 0;
                    try
                    {
                        if (rowValue[3] != rowValue[4])
                        {
                            error = " Column values are not equal";
                        }

                        if (!rowValue[3].ToString().Contains(rowValue[4].ToString()))
                        {
                            error += " Column values are not present";
                        }

                        if (!string.IsNullOrEmpty(error))
                        {                             
                            rowValue.Add(true);
                            errorColumnIndex = rowValue.Count() - 1;
                            rowValue.Add(error);
                            errorMessageColumnIndex = rowValue.Count() - 1;
                        }
                        dataTable.Rows.Add(rowValue.ToArray());
                        validationTable.Rows.Add(rowValue.ToArray());
                    }
                    catch (System.Exception ex)
                    {
                        if (errorColumnIndex==0)
                        {
                            rowValue.Add(true);
                            rowValue.Add(error);
                        }
                        else
                        {
                            rowValue[errorColumnIndex] = true;
                            rowValue[errorMessageColumnIndex] += ex.Message;
                        }
                        validationTable.Rows.Add(rowValue.ToArray());
                    }
                });
            }
            return validationTable;
        }
        static void table_RowChanged(object sender, DataRowChangeEventArgs e)
        {
            //if(e.Row[3] != e.Row[4])
            //{
            //    e.Row.RowError += " Column values are not equal";
            //}

            //if (!e.Row[3].ToString().Contains(e.Row[4].ToString()))
            //{
            //    e.Row.RowError += " Column values are not present";
            //}

            if (e.Row.HasErrors)
            {

                System.Diagnostics.Debug.WriteLine($"Row Error :::: {e.Row.RowError}");
                //e.Row.SetField("Error", e.Row.RowError);
                //e.Row.SetField("IsContainError", true);
            }
        }
        public static List<DataRow> ValidatRules(DataTable dataTable)
        {
            List<DataRow> invalidRows = new List<DataRow>();
            Dictionary<string, string> conditions = new Dictionary<string, string>();
            return invalidRows;
        }

        private static void AddPrimaryKeyConstraint(DataTable dataTable, int index)
        {
            var primaryKeyColumn = dataTable.Columns[index];
            UniqueConstraint primaryKey = new UniqueConstraint(primaryKeyColumn, true);
            dataTable.Constraints.Add(primaryKey);
        }

        private static void AddUniqueConstraints(DataTable dataTable, List<int> uniqueColumnsIndex)
        {
            if (uniqueColumnsIndex != null && uniqueColumnsIndex.Any())
            {
                List<DataColumn> dataColumns = new List<DataColumn>();
                foreach (var columnIndex in uniqueColumnsIndex)
                {
                    dataColumns.Add(dataTable.Columns[columnIndex]);
                }

                UniqueConstraint unique = new UniqueConstraint("groupByColumnConstraint", dataColumns.ToArray());
                dataTable.Constraints.Add(unique);
            }
        }
    }
}