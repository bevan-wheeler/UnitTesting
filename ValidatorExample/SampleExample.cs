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
    /// <summary>
    /// Macro Test Example against VBA Functions
    /// </summary>
    [TestClass]
    public class SampleExample
    {
        private static Excel.Application xlApp;
        private string sProcedureName;
        private int iExpected;
        private int iActual;
        private int iValueA;
        private int iValueB;

        #region Class Setup
        /// <summary>
        /// Creating the Excel Objects. Based on active workbook.
        /// </summary>
        [ClassInitialize()]
        public static void MyClassInitialize(TestContext tc)
        {
            xlApp = new Excel.Application();
            xlApp = (Excel.Application)Marshal.GetActiveObject("excel.application");
            xlApp.Visible = true;
            xlApp.DisplayAlerts = true;
        }

        /// <summary>
        /// Releasing COM Excel Objects
        /// </summary>
        [ClassCleanup()]
        public static void MyClassCleanup()
        {
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xlApp);
        }
        #endregion

        #region Test Setup
        [TestInitialize()]
        public void MyTestInitialize()
        {
            iValueA = 2;
            iValueB = 5;
        }

        [TestCleanup()]
        public void MyTestCleanup()
        {
            sProcedureName = null;
            iValueA = 0;
            iValueB = 0;
            iExpected = 0;
            iActual = 0;
        }
        #endregion

        [TestMethod]
        public void Add()
        {
            //Arrange
            sProcedureName = "func_Add";
            iExpected = iValueA + iValueB;

            //Act
            iActual = xlApp.Application.Run(sProcedureName, iValueA, iValueB);

            //Assert
            Assert.AreEqual(iExpected, iActual);
        }

        [TestMethod]
        public void Subtract()
        {
            //Arrange
            sProcedureName = "func_Subtract";
            iExpected = iValueA - iValueB;

            //Act
            iActual = xlApp.Application.Run(sProcedureName, iValueA, iValueB);

            //Assert
            Assert.AreEqual(iExpected, iActual);
        }
    }
}
