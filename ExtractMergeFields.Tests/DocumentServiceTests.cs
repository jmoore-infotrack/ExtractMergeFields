using System;
using System.Collections.Generic;
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
            paragraph.ChildElements[3].Append(text);

            // Act
            _service.FixIncorrectTextValue(paragraph.Descendants<Run>(), "hello world", "HELLO WORLD");

            // Assert
            Assert.AreEqual(text.InnerText, "HELLO WORLD");
        }

        public Paragraph GetMockParagraph()
        {
            var paragraph = new Paragraph();
            var runs = new Run[5] { new Run(), new Run(), new Run(), new Run(), new Run() };
            foreach (var run in runs)
            {
                paragraph.Append(run);
            }
            return paragraph;
        }
    }
}
