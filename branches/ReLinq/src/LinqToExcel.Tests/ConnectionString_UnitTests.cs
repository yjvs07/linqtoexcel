using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MbUnit.Framework;
using System.Data.OleDb;

namespace LinqToExcel.Tests
{
    [Author("Paul Yoder", "paulyoder@gmail.com")]
    [FixtureCategory("Unit")]
    [TestFixture]
    public class ConnectionString_UnitTests : SQLLogStatements_Helper
    {
        [TestFixtureSetUp]
        public void fs()
        {
            InstantiateLogger();
        }

        [SetUp]
        public void s()
        {
            ClearLogEvents();
        }

        [Test]
        public void xls_connection_string()
        {
            var companies = from c in ExcelQueryFactory.Worksheet("spreadsheet.xls")
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            var expected = string.Format(
                @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=""Excel 8.0;HDR=YES;""",
                "spreadsheet.xls");
            Assert.AreEqual(expected, GetConnectionString());
        }

        [Test]
        public void unknown_file_type_connection_string()
        {
            var companies = from c in ExcelQueryFactory.Worksheet("spreadsheet.dlo")
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            var expected = string.Format(
                @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=""Excel 8.0;HDR=YES;""",
                "spreadsheet.dlo");
            Assert.AreEqual(expected, GetConnectionString());
        }

        [Test]
        public void csv_connection_string()
        {
            var companies = from c in ExcelQueryFactory.Worksheet(@"C:\Desktop\spreadsheet.csv")
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            var expected = string.Format(
                @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=""text;HDR=Yes;FMT=Delimited;""",
                @"C:\Desktop");
            Assert.AreEqual(expected, GetConnectionString());
        }

        [Test]
        public void xlsx_connection_string()
        {
            var companies = from c in ExcelQueryFactory.Worksheet("spreadsheet.xlsx")
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            var expected = string.Format(
                @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0 Xml;HDR=YES""",
                "spreadsheet.xlsx");
            Assert.AreEqual(expected, GetConnectionString());
        }

        [Test]
        public void xlsm_connection_string()
        {
            var companies = from c in ExcelQueryFactory.Worksheet("spreadsheet.xlsm")
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            var expected = string.Format(
                @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0 Xml;HDR=YES""",
                "spreadsheet.xlsm");
            Assert.AreEqual(expected, GetConnectionString());
        }

        [Test]
        public void xlsb_connection_string()
        {
            var companies = from c in ExcelQueryFactory.Worksheet("spreadsheet.xlsb")
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            var expected = string.Format(
                @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0;HDR=YES""",
                "spreadsheet.xlsb");
            Assert.AreEqual(expected, GetConnectionString());
        }

        //[Test]
        //public void FileType_is_set_to_ExcelVersion_PreExcel2007_for_files_with_xls_extensions()
        //{
        //    var repo = new ExcelRepository<Company>("spreadsheet.xls");
        //    Assert.AreEqual(ExcelVersion.PreExcel2007, repo.FileType);
        //}

        //[Test]
        //public void FileType_is_set_to_ExcelVersion_PreExcel2007_for_files_with_XLS_extensions()
        //{
        //    var repo = new ExcelRepository<Company>("spreadsheet.XLS");
        //    Assert.AreEqual(ExcelVersion.PreExcel2007, repo.FileType);
        //}

        //[Test]
        //public void FileType_is_set_to_ExcelVersion_Csv_for_files_with_csv_extensions()
        //{
        //    var repo = new ExcelRepository<Company>("spreadsheet.csv");
        //    Assert.AreEqual(ExcelVersion.Csv, repo.FileType);
        //}

        //[Test]
        //public void FileType_is_set_to_ExcelVersion_Csv_for_files_with_CSV_extensions()
        //{
        //    var repo = new ExcelRepository<Company>("spreadsheet.CSV");
        //    Assert.AreEqual(ExcelVersion.Csv, repo.FileType);
        //}

        //[Test]
        //public void FileType_is_set_to_ExcelVersion_PreExcel2007_for_files_with_unrecognized_extensions()
        //{
        //    var repo = new ExcelRepository<Company>("spreadsheet.tdl");
        //    Assert.AreEqual(ExcelVersion.PreExcel2007, repo.FileType);
        //}
    }
}
