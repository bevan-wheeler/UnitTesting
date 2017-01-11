namespace UnitTest
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// SetupExample class
    /// </summary>
    [TestClass]
    public class ValidatorDates
    {
        private static ExcelTest xlTest; // Testing object

        private string sProcedureName;
        private string sDate;
        private string sDateStart;
        private string sDateEnd;
        private string sParentEnd;
        private string sParentStart;

        private bool bResult;        // Result boolean

        /// <summary>
        /// Use ClassInitialize to run code before running the first test in the class
        /// </summary>
        /// <param name="testContext">Current Testing Context</param>
        [ClassInitialize]
        public static void MyClassInitialize(TestContext testContext)
        {
            // Generate our new testing object
            xlTest = new ExcelTest();

            // Create a new instance of our class
            // This is only needed for class objects
            xlTest.SetClass("ValidatorPriceListDates");
        }

        /// <summary>
        /// Use ClassCleanup to run code after all tests in a class have run
        /// </summary>
        [ClassCleanup]
        public static void MyClassCleanup()
        {
            xlTest.Dispose();
        }

        /// <summary>
        /// Use TestInitialize to run code before running each test
        /// </summary>
        [TestInitialize]
        public void MyTestInitialize()
        {
            // Add Initialise Test code here
            this.bResult = false;
        }

        /// <summary>
        /// Use TestCleanup to run code after each test has run
        /// </summary>
        [TestCleanup]
        public void MyTestCleanup()
        {
            // Add Termination Test code here
            this.bResult = false;
        }

        /// <summary>
        /// Start Date is larger than End Date
        /// </summary>
        [TestMethod]
        public void InnerDateLogic_A()
        {
            // Arrange
            this.sProcedureName = "callInnerDatelogic";
            this.sDateStart = "12/12/2100";
            this.sDateEnd = "01/01/2000";

            // Act
            // object ret = validator.GetType().GetMethod(this.sProcedureName).Invoke(validator, new object[] { this.sDateStart, this.sDateEnd });
            this.bResult = xlTest.ExcelApp.Run(this.sProcedureName, this.sDateStart, this.sDateEnd);

            // Assert
            Assert.IsFalse(this.bResult);
        }

        /// <summary>
        /// Date is smaller than Parent Start Date
        /// </summary>
        [TestMethod]
        public void InnerParentDateLogic_A()
        {
            // Arrange
            this.sProcedureName = "callInnerParentDateLogic";
            this.sDate = "19001212";
            this.sParentStart = "20000101";
            this.sParentEnd = "21001212";

            // Act
            this.bResult = xlTest.ExcelApp.Run(this.sProcedureName, this.sDate, this.sParentStart, this.sParentEnd);

            // object objResult = xlTest.RunClass(this.sProcedureName, new object[] { this.sDate, this.sParentStart, this.sParentEnd });

            // Assert
            Assert.IsFalse(this.bResult);
        }
    }
}
