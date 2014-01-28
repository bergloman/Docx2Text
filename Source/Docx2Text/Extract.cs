using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Docx2Text {

    public class DocxExtractor : IDisposable {

        private bool disposed = false;

        private WordprocessingDocument package = null;

        private string file_name = string.Empty;

        public DocxExtractor(string input_file) {
            this.file_name = input_file;
            if (string.IsNullOrEmpty(input_file) || !File.Exists(input_file)) {
                throw new Exception("The file is invalid. Please select an existing file again");
            }
            this.package = WordprocessingDocument.Open(input_file, true);
        }

        public string ReadWordDocument() {
            var sb = new StringBuilder();
            var element = package.MainDocumentPart.Document.Body;
            if (element == null) {
                return string.Empty;
            }
            sb.Append(GetPlainText(element));
            return sb.ToString();
        }

        private string GetPlainText(OpenXmlElement element) {
            var PlainTextInWord = new StringBuilder();
            foreach (var section in element.Elements()) {
                switch (section.LocalName) {
                    // Text
                    case "t":
                        PlainTextInWord.Append(section.InnerText);
                        break;

                    case "cr":                          // Carriage return
                    case "br":                          // Page break
                        PlainTextInWord.Append(Environment.NewLine);
                        break;

                    // Tab
                    case "tab":
                        PlainTextInWord.Append("\t");
                        break;

                    // Paragraph
                    case "p":
                        PlainTextInWord.Append(GetPlainText(section));
                        PlainTextInWord.AppendLine(Environment.NewLine);
                        break;

                    default:
                        PlainTextInWord.Append(GetPlainText(section));
                        break;
                }
            }

            return PlainTextInWord.ToString();
        }

        #region IDisposable interface

        public void Dispose() {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing) {
            if (disposed) {
                return;
            }
            if (disposing) {
                if (this.package != null) {
                    this.package.Dispose();
                }
            }
            disposed = true;
        }
        #endregion
    }
}
