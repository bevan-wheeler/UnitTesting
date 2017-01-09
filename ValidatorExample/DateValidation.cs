using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Vbe.Interop;
using System.Runtime.InteropServices;
using System.Reflection;


namespace UnitTest
{
    /// <summary>
    /// Macro Testing ValidatorPriceListDates
    /// </summary>
    [TestClass]
    public class DateRangeValidatorUnitTesting
    {
        private static Excel.Application xlApp;
        private static object validator;
        private string sProcedureName;
        private string sDateStart;
        private string sDateEnd;
        private string sParentStart;
        private string sParentEnd;
        private string sDate;
        private bool bActual;

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

            // Bevan - Test
            validator = getVBAClass("ValidatorPriceListDates");
        }

        /// <summary>
        /// Bevan Test Add Component method
        /// We can't use reflection to call the classes by name
        /// </summary>
        public static object getVBAClass(string sClassName) 
        {
            // Grab the Class component being passed in from name
            VBProject xlProj = xlApp.ActiveWorkbook.VBProject;
            VBComponent compVal = xlProj.VBComponents.Item(sClassName);

            // Function name to run
            string sFunctionName = "UNIT_TEST" + sClassName;

            //Add a new module/function
            VBComponent compModule = xlProj.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
            compModule.CodeModule.InsertLines(compModule.CodeModule.CountOfLines + 1,"Public Function " + sFunctionName + "() As " + sClassName + "\r\n Set " + sFunctionName + " = New " + sClassName + "\r\n End Function");

            // Run the function
            object validator = xlApp._Run2(sFunctionName);

            // Remove the function
            xlProj.VBComponents.Remove(compModule);

            // Clean the COM references
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(compModule);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(compVal);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xlProj);

            // Return the object
            return validator;
        }

        /// <summary>
        /// Releasing COM Excel Objects
        /// </summary>
        [ClassCleanup()]
        public static void MyClassCleanup()
        {
            // Holy shit clean up everything or you'll get a memory leak
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(validator);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xlApp);
        }
        #endregion

        #region Test Setup

        [TestInitialize()]
        public void MyTestInitialize()
        {
            sProcedureName = null;
            sDateStart = null;
            sDateEnd = null;
            sParentStart = null;
            sParentEnd = null;
            sDate = null;
            bActual = false;
        }

        [TestCleanup()]
        public void MyTestCleanup()
        {
            sProcedureName = null;
            sDateStart = null;
            sDateEnd = null;
            sParentStart = null;
            sParentEnd = null;
            sDate = null;
            bActual = false;
        }
        #endregion

        /// Unit Tests against InnerDateLogic Function
        /// Date string uses the format of dd/mm/yyyy
        #region InnerDateLogic

        /// <summary>
        /// Start Date is larger than End Date
        /// </summary>
        [TestMethod]
        public void InnerDateLogic_A()
        {
            //Arrange
            sProcedureName = "callInnerDatelogic";
            sDateStart = "12/12/2100";
            sDateEnd = "01/01/2000";

            //Act
            object ret = validator.GetType().GetMethod(sProcedureName).Invoke(validator, new object[]{sDateStart, sDateEnd}); ;
            // bActual = xlApp.Run(sProcedureName, sDateStart, sDateEnd);

            //Assert
            Assert.IsFalse(bActual);
        }

        /// <summary>
        /// Start Date is smaller than End Date
        /// </summary>
        [TestMethod]
        public void InnerDateLogic_B()
        {
            //Arrange
            sProcedureName = "callInnerDatelogic";
            sDateStart = "01/01/2000";
            sDateEnd = "12/12/2100";

            //Act
            bActual = xlApp.Run(sProcedureName, sDateStart, sDateEnd);

            //Assert
            Assert.IsTrue(bActual);
        }

        /// <summary>
        /// Start Date is Blank, End Date contains Value
        /// </summary>
        [TestMethod]
        public void InnerDateLogic_C()
        {
            //Arrange
            sProcedureName = "callInnerDatelogic";
            sDateStart = "";
            sDateEnd = "01/01/2000";

            //Act
            bActual = xlApp.Run(sProcedureName, sDateStart, sDateEnd);

            //Assert
            Assert.IsTrue(bActual);
        }

        /// <summary>
        /// Start Date contains Value, End Date is Blank
        /// </summary>
        [TestMethod]
        public void InnerDateLogic_D()
        {
            //Arrange
            sProcedureName = "callInnerDatelogic";
            sDateStart = "01/01/2000";
            sDateEnd = "";

            //Act
            bActual = xlApp.Run(sProcedureName, sDateStart, sDateEnd);

            //Assert
            Assert.IsTrue(bActual);
        }

        /// <summary>
        /// Start Date is Blank, End Date is Blank
        /// </summary>
        [TestMethod]
        public void InnerDateLogic_E()
        {
            //Arrange
            sProcedureName = "callInnerDatelogic";
            sDateStart = "";
            sDateEnd = "";

            //Act
            bActual = xlApp.Run(sProcedureName, sDateStart, sDateEnd);

            //Assert
            Assert.IsTrue(bActual);
        }

        /// <summary>
        /// Start Date the same as End Date
        /// </summary>
        [TestMethod]
        public void InnerDateLogic_F()
        {
            //Arrange
            sProcedureName = "callInnerDatelogic";
            sDateStart = "12/12/2012";
            sDateEnd = "12/12/2012";

            //Act
            bActual = xlApp.Run(sProcedureName, sDateStart, sDateEnd);

            //Assert
            Assert.IsTrue(bActual);
        }

        #endregion

        // Marco Tests against InnerParentDateLogic Function
        // Date String uses the formate of yyyymmdd
        #region InnerParentDateLogic

        /// <summary>
        /// Date is smaller than Parent Start Date
        /// </summary>
        [TestMethod]
        public void InnerParentDateLogic_A()
        {
            //Arrange
            sProcedureName = "callInnerParentDateLogic";
            sDate = "19001212";
            sParentStart = "20000101";
            sParentEnd = "21001212";

            //Act
            bActual = xlApp.Run(sProcedureName, sDate, sParentStart, sParentEnd);

            //Assert
            Assert.IsFalse(bActual);
        }

        /// <summary>
        /// Date is between Parent Start and Parent End Dates
        /// </summary>
        [TestMethod]
        public void InnerParentDateLogic_B()
        {
            //Arrange
            sProcedureName = "callInnerParentDateLogic";
            sDate = "20501212";
            sParentStart = "20000101";
            sParentEnd = "21001212";

            //Act
            bActual = xlApp.Run(sProcedureName, sDate, sParentStart, sParentEnd);

            //Assert
            Assert.IsTrue(bActual);
        }

        /// <summary>
        /// Date is larger than Parent End Date
        /// </summary>
        [TestMethod]
        public void InnerParentDateLogic_C()
        {
            //Arrange
            sProcedureName = "callInnerParentDateLogic";
            sDate = "21501212";
            sParentStart = "20000101";
            sParentEnd = "21001212";

            //Act
            bActual = xlApp.Run(sProcedureName, sDate, sParentStart, sParentEnd);

            //Assert
            Assert.IsFalse(bActual);
        }

        /// <summary>
        /// Date is Blank, Parent Start and End contain values
        /// </summary>
        [TestMethod]
        public void InnerParentDateLogic_D()
        {
            //Arrange
            sProcedureName = "callInnerParentDateLogic";
            sDate = "";
            sParentStart = "20000101";
            sParentEnd = "210001212";

            //Act
            bActual = xlApp.Run(sProcedureName, sDate, sParentStart, sParentEnd);

            //Assert
            Assert.IsTrue(bActual);
        }

        /// <summary>
        /// Parent Start is Blank, Parent End and Date contain values
        /// Date Value is larger than Parent End
        /// </summary>
        [TestMethod]
        public void InnerParentDateLogic_E()
        {
            //Arrange
            sProcedureName = "callInnerParentDateLogic";
            sDate = "23001212";
            sParentStart = "";
            sParentEnd = "21001212";

            //Act
            bActual = xlApp.Run(sProcedureName, sDate, sParentStart, sParentEnd);

            //Assert
            Assert.IsFalse(bActual);
        }

        /// <summary>
        /// Parent Start is Blank, Parent End and Date contain values
        /// Date Value is smaller than Parent End
        /// </summary>
        [TestMethod]
        public void InnerParentDateLogic_F()
        {
            //Arrange
            sProcedureName = "callInnerParentDateLogic";
            sDate = "20001212";
            sParentStart = "";
            sParentEnd = "21001212";

            //Act
            bActual = xlApp.Run(sProcedureName, sDate, sParentStart, sParentEnd);

            //Assert
            Assert.IsTrue(bActual);
        }

        /// <summary>
        /// Parent End is Blank, Parent Start and Date contain Values
        /// Date is larger than Parent Start
        /// </summary>
        [TestMethod]
        public void InnerParentDateLogic_G()
        {
            //Arrange
            sProcedureName = "callInnerParentDateLogic";
            sDate = "22001212";
            sParentStart = "21001212";
            sParentEnd = "";

            //Act
            bActual = xlApp.Run(sProcedureName, sDate, sParentStart, sParentEnd);

            //Assert
            Assert.IsTrue(bActual);
        }

        /// <summary>
        /// Parent End is Blank, Parent Start and Date contain Values
        /// Date Value is smaller than Parent Start
        /// </summary>
        [TestMethod]
        public void InnerParentDateLogic_H()
        {
            //Arrange
            sProcedureName = "callInnerParentDateLogic";
            sDate = "19001212";
            sParentStart = "21001212";
            sParentEnd = "";

            //Act
            bActual = xlApp.Run(sProcedureName, sDate, sParentStart, sParentEnd);

            //Assert
            Assert.IsFalse(bActual);
        }

        /// <summary>
        /// Parent End is Blank, Parent Start is Blank, Date contains a value
        /// </summary>
        [TestMethod]
        public void InnerParentDateLogic_I()
        {
            //Arrange
            sProcedureName = "callInnerParentDateLogic";
            sDate = "19001212";
            sParentStart = "";
            sParentEnd = "";

            //Act
            bActual = xlApp.Run(sProcedureName, sDate, sParentStart, sParentEnd);

            //Assert
            Assert.IsTrue(bActual);
        }

        /// <summary>
        /// Date is the same as Parent Start
        /// </summary>
        [TestMethod]
        public void InnerParentDateLogic_J()
        {
            //Arrange
            sProcedureName = "callInnerParentDateLogic";
            sDate = "20001212";
            sParentStart = "20001212";
            sParentEnd = "21001212";

            //Act
            bActual = xlApp.Run(sProcedureName, sDate, sParentStart, sParentEnd);

            //Assert
            Assert.IsTrue(bActual);
        }

        /// <summary>
        /// Date is the same as Parent End
        /// </summary>
        [TestMethod]
        public void InnerParentDateLogic_K()
        {
            //Arrange
            sProcedureName = "callInnerParentDateLogic";
            sDate = "21001212";
            sParentStart = "20001212";
            sParentEnd = "21001212";

            //Act
            bActual = xlApp.Run(sProcedureName, sDate, sParentStart, sParentEnd);

            //Assert
            Assert.IsTrue(bActual);
        }

        /// <summary>
        /// Date is Blank, Parent Start is Blank, Parent End is Blank
        /// </summary>
        [TestMethod]
        public void InnerParentDateLogic_L()
        {
            //Arrange
            sProcedureName = "callInnerParentDateLogic";
            sDate = "";
            sParentStart = "";
            sParentEnd = "";

            //Act
            bActual = xlApp.Run(sProcedureName, sDate, sParentStart, sParentEnd);

            //Assert
            Assert.IsTrue(bActual);
        }

        /// <summary>
        /// Date is Negative
        /// A Negative Date will never be passed into the InnerParentDateLogic function
        /// Object Logic checks fail negative numbers
        /// </summary>
        [TestMethod]
        public void InnerParentDateLogic_M()
        {
            //Arrange
            sProcedureName = "callInnerParentDateLogic";
            sDate = "-20501212";
            sParentStart = "20000101";
            sParentEnd = "21001212";

            //Act
            bActual = xlApp.Run(sProcedureName, sDate, sParentStart, sParentEnd);

            //Assert
            Assert.IsFalse(bActual);
        }

        /// <summary>
        /// Parent Start and Parent End are the same
        /// Date is smaller than Parent Start/End
        /// </summary>
        [TestMethod]
        public void InnerParentDateLogic_N()
        {
            //Arrange
            sProcedureName = "callInnerParentDateLogic";
            sDate = "20001212";
            sParentStart = "21001212";
            sParentEnd = "21001212";

            //Act
            bActual = xlApp.Run(sProcedureName, sDate, sParentStart, sParentEnd);

            //Assert
            Assert.IsFalse(bActual);
        }

        /// <summary>
        /// Parent Start and Parent End are the same
        /// Date is larger than Parent Start/End
        /// </summary>
        [TestMethod]
        public void InnerParentDateLogic_O()
        {
            //Arrange
            sProcedureName = "callInnerParentDateLogic";
            sDate = "22001212";
            sParentStart = "21001212";
            sParentEnd = "21001212";

            //Act
            bActual = xlApp.Run(sProcedureName, sDate, sParentStart, sParentEnd);

            //Assert
            Assert.IsFalse(bActual);
        }

        /// <summary>
        /// Date, Parent Start and Parent End are all the same
        /// </summary>
        [TestMethod]
        public void InnerParentDateLogic_P()
        {
            //Arrange
            sProcedureName = "callInnerParentDateLogic";
            sDate = "20000101";
            sParentStart = "20000101";
            sParentEnd = "20000101";

            //Act
            bActual = xlApp.Run(sProcedureName, sDate, sParentStart, sParentEnd);

            //Assert
            Assert.IsTrue(bActual);
        }

        /// <summary>
        /// Parent End date is before Parent Start Date
        /// Date is smaller than Parent End Date
        /// </summary>
        [TestMethod]
        public void InnerParentDateLogic_Q()
        {
            //Arrange
            sProcedureName = "callInnerParentDateLogic";
            sDate = "19000101";
            sParentStart = "21000101";
            sParentEnd = "20000101";

            //Act
            bActual = xlApp.Run(sProcedureName, sDate, sParentStart, sParentEnd);

            //Assert
            Assert.IsFalse(bActual);
        }

        /// <summary>
        /// Parent End date is before Parent Start Date
        /// Date is larger than Parent Start Date
        /// </summary>
        [TestMethod]
        public void InnerParentDateLogic_R()
        {
            //Arrange
            sProcedureName = "callInnerParentDateLogic";
            sDate = "22000101";
            sParentStart = "21000101";
            sParentEnd = "20000101";

            //Act
            bActual = xlApp.Run(sProcedureName, sDate, sParentStart, sParentEnd);

            //Assert
            Assert.IsFalse(bActual);
        }

        /// <summary>
        /// Parent End date is before Parent Start Date
        /// Date is larger than Parent End Date and smaller than Parent Start Date
        /// </summary>
        [TestMethod]
        public void InnerParentDateLogic_S()
        {
            //Arrange
            sProcedureName = "callInnerParentDateLogic";
            sDate = "21000101";
            sParentStart = "22000101";
            sParentEnd = "20000101";

            //Act
            bActual = xlApp.Run(sProcedureName, sDate, sParentStart, sParentEnd);

            //Assert
            Assert.IsFalse(bActual);
        }

        #endregion

    }
}
