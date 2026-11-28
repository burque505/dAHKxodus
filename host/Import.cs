// Opening older formats: the file becomes .docx bytes here, in the host, and the page opens those as it opens any
// .docx. The window then holds a new document named like the file with .docx, which Save puts beside the old file
// (dAHKxodus writes .docx, never the old format: the old file stays as it was).
//   .doc  Word 97-2003: DocSharp's converter (lib\docsharp, built by build-docsharp.ps1 from the fork commit in
//         docsharp.lock), then Docxodus's pass that puts elements in the order the standard gives
// What the file is comes from its first bytes, not its name: a ".doc" is often really an RTF file (WordPad and many
// programs saved RTF under .doc) or a .docx renamed.
using System.Xml.Linq;
using Docxodus;
using DocumentFormat.OpenXml.Packaging;

public static class Import {
    // The extensions File > Open, dropping and Recent take besides Word's own (.docx, .docm, .dotx, .dotm).
    public static readonly string[] Extensions = [".doc"];

    public static bool Handles(string path) => Extensions.Contains(Path.GetExtension(path).ToLowerInvariant());

    public enum Kind { Doc, Docx, Rtf, Unknown }

    public static Kind KindOf(byte[] bytes) {
        if (bytes.Length >= 8 && bytes[0] == 0xD0 && bytes[1] == 0xCF && bytes[2] == 0x11 && bytes[3] == 0xE0
            && bytes[4] == 0xA1 && bytes[5] == 0xB1 && bytes[6] == 0x1A && bytes[7] == 0xE1) return Kind.Doc;
        if (bytes.Length >= 4 && bytes[0] == (byte)'P' && bytes[1] == (byte)'K' && bytes[2] == 3 && bytes[3] == 4) return Kind.Docx;
        if (bytes.Length >= 5 && System.Text.Encoding.ASCII.GetString(bytes, 0, 5) == @"{\rtf") return Kind.Rtf;
        return Kind.Unknown;
    }

    /// Why a file couldn't be opened, in words for the user (a Lang key and its argument).
    public sealed class ImportException(string key, string detail = "") : Exception(key) {
        public string Key { get; } = key;
        public string Detail { get; } = detail;
    }

    /// The file's content as .docx bytes. Throws ImportException when it can't be converted.
    public static byte[] ToDocx(byte[] bytes) {
        switch (KindOf(bytes)) {
            case Kind.Docx: return bytes;    // (a .docx under another name)
            case Kind.Doc: return Order(FromDoc(bytes));
            case Kind.Rtf: throw new ImportException("import.rtfNotYet");
            default: throw new ImportException("import.unknown");
        }
    }

    static byte[] FromDoc(byte[] bytes) {
        try {
            using var input = new MemoryStream(bytes, writable: false);
            using var reader = new DocSharp.Binary.StructuredStorage.Reader.StructuredStorageReader(input);
            var doc = new DocSharp.Binary.DocFileFormat.WordDocument(reader);
            // (DocSharp doesn't decrypt: an encrypted file's text would come out as noise)
            if (doc.FIB.fEncrypted) throw new ImportException("import.encrypted");
            using var output = new MemoryStream();
            using (var docx = DocSharp.Binary.OpenXmlLib.WordprocessingML.WordprocessingDocument.Create(output, DocSharp.Binary.OpenXmlLib.WordprocessingDocumentType.Document))
                DocSharp.Binary.WordprocessingMLMapping.Converter.Convert(doc, docx);
            return output.ToArray();
        } catch (ImportException) {
            throw;
        } catch (DocSharp.Binary.DocFileFormat.ByteParseException ex) when (ex.Message.Contains("Word 95")) {
            throw new ImportException("import.tooOld");
        } catch (Exception ex) {
            throw new ImportException("import.failed", ex.Message);
        }
    }

    // Docxodus's pass putting each part's elements in the order the standard gives (DocSharp writes some properties
    // out of order; Word and the engine read them anyway, but a file saved from them should be one Word accepts as is).
    static byte[] Order(byte[] docx) {
        using var ms = new MemoryStream();
        ms.Write(docx);
        using (var doc = WordprocessingDocument.Open(ms, true)) {
            var parts = new List<OpenXmlPart>();
            void Walk(OpenXmlPartContainer container) {
                foreach (var part in container.Parts.Select(p => p.OpenXmlPart))
                    if (!parts.Contains(part)) { parts.Add(part); Walk(part); }
            }
            Walk(doc);
            foreach (var part in parts.Where(p => p.ContentType.Contains("wordprocessingml") && p.ContentType.EndsWith("+xml"))) {
                XDocument xml;
                using (var s = part.GetStream()) xml = XDocument.Load(s);
                if (xml.Root == null) continue;
                var root = (XElement)WordprocessingMLUtil.WmlOrderElementsPerStandard(xml.Root);
                using var w = part.GetStream(FileMode.Create, FileAccess.Write);
                new XDocument(xml.Declaration, root).Save(w);
            }
        }
        return ms.ToArray();
    }
}
