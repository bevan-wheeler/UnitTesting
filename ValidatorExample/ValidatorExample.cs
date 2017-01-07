using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Reflection;

namespace ValidatorExample
{
    [TestClass]
    public class ValidatorExample
    {
        private static Excel.Application xlApp;
        private bool bActual;

        private string sProcedureName;
        private string sStartDate;
        private string sEndDate;

        #region Class Settings
        [ClassInitialize]
        public static void classInitialize(TestContext tc)
        {
            xlApp = new Excel.Application();
            xlApp = (Excel.Application)Marshal.GetActiveObject("excel.application");
            xlApp.Visible = true;
            xlApp.DisplayAlerts = true;
        }

        [ClassCleanup]
        public static void classClean()
        {
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xlApp);
        }
        #endregion

        #region Test Settings
        [TestInitialize]
        public void testInitialize()
        {
            bActual = false;
            sProcedureName = "";
        }

        [TestCleanup]
        public void testClean()
        {
            bActual = false;
            sProcedureName = null;
        }
        #endregion

        [TestMethod]
        public void ValidatorTester()
        {
            //Arrange
            sProcedureName = "callValidator";
            string sSectionName = "Line Section";

            string[] sColumnNames = new string[] { "Supplier" };
            string[] sDepColNames = new string[] { "Document Line" };
            string[] sFNDValues = new string[] { "1", "2", "3", "4", "5", "6", "7", "8", "9" };
            string[] sFNDKeys = new string[] { "LINE_NUM", "SUPPLIER", "SITE", "ITEM", "LINE_TYPE", "CATEGORY", "ITEM_DESCRIPTION", "UNIT_MEAS_LOOKUP_CODE", "UNIT_PRICE" };
            string[] sDOMValues = new string[] { "1", "2", "3", "4", "5", "6", "7", "8", "9" };
            string[] sDOMKeys = new string[] { "Document Line", "Supplier", "Site", "Item", "Line Type", "Category", "Description", "Unit of Measure", "Unit Price" };
            string[] sDOMIDs = new string[] { "1", "", "", "", "", "", "", "", "" };

            //Act
            bActual = xlApp.Run(sProcedureName, sColumnNames, sDepColNames, sSectionName, sFNDValues, sFNDKeys, sDOMValues, sDOMKeys, sDOMIDs);

            //Assert
            Assert.IsTrue(bActual);
        }

        [TestMethod]
        public void validatorDates()
        {
            //Arrange
            sProcedureName = "callInnerDatelogic";
            sStartDate = "05/12/2015";
            sEndDate = "12/05/2012";

            //Act
            bActual = xlApp.Run(sProcedureName, sStartDate, sEndDate);

            //Assert
            Assert.IsTrue(bActual);

        }

    }


}
