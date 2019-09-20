using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Office2010.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExtractMergeFields.Tests
{
    [TestClass]
    public class DocumentServiceTests
    {
        private IDocumentService _service;
        [TestInitialize]
        public void Setup()
        {
            _service = new DocumentService();
        }

        [TestMethod]
        public void FixIncorrectTextValue_ShouldUpdateValue()
        {
            // Arrange
            var paragraph = GetMockParagraph();
            var text = new Text("hello world");
            Run runUnderTest = new Run();
            runUnderTest.Append(new FieldCode());
            runUnderTest.Append(new Text("hello world"));
            paragraph.Append(runUnderTest);

            // Act
            _service.FixIncorrectTextValue(paragraph.Descendants<Run>(), "hello world", "HELLO WORLD");

            // Assert
            Assert.AreEqual(paragraph.ChildElements[5].InnerText, "HELLO WORLD");
        }

        [TestMethod]
        public void ChangeSingleMergefield_ShouldDoNothing_IfLengthIncorrect()
        {
            // Arrange
            var paragraph = GetMockParagraph();
            paragraph.RemoveChild((Run)paragraph.ChildElements[0]);
            var expected = "original";
            var fields = paragraph.Descendants<FieldCode>();

            // Act
            _service.ChangeSingleMergefield(fields, "test");
            var actual = paragraph.ChildElements[3].InnerText;

            // Assert
            Assert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void ChangeSingleMergefield_ShouldChangeText_IfLengthCorrect()
        {
            // Arrange
            var paragraph = GetMockParagraph();
            var fields = paragraph.Descendants<FieldCode>();
            var expected = "test";

            // Act
            _service.ChangeSingleMergefield(fields, "test");
            var actual = paragraph.ChildElements[3].InnerText;

            // Assert
            Assert.AreEqual(expected, actual);
        }

        public Paragraph GetMockParagraph()
        {
            var paragraph = new Paragraph();
            var runs = new Run[5] { new Run(), new Run(), new Run(), new Run(), new Run() };
            foreach (var run in runs)
            {
                run.Append(new FieldCode());
                run.Append(new Text("original"));
                paragraph.Append(run);
            }
            return paragraph;
        }
    }
}
