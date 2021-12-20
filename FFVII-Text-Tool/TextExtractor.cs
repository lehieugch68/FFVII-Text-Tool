using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace FFVII_Text_Tool
{
    public static class TextExtractor
    {
        #region Structure
        private struct Header
        {
            public short Unk;
            public int LangLen;
            public byte[] Language;
            public int TextCount;
            public byte[] HeaderData;
            public BlockText[] Blocks;
            public int Magic;
            public string FileName;
        }
        private struct BlockText
        {
            public int IdLen;
            public string Id;
            public int StrLen;
            public string Str;
            public int InfoCount;
            public TextInfo[] TextInfos;
        }
        private struct TextInfo
        {
            public int Type;
            public int Unk;
            public int InfoLen;
            public string Info;
        }
        #endregion
        private static Header ReadHeader(ref BinaryReader reader)
        {
            reader.BaseStream.Seek(0, SeekOrigin.Begin);
            Header header = new Header();
            header.Unk = reader.ReadInt16();
            header.LangLen = reader.ReadInt32();
            header.Language = reader.ReadBytes(header.LangLen);
            reader.BaseStream.Position += 4;
            header.TextCount = reader.ReadInt32();
            long headerLen = reader.BaseStream.Position;
            reader.BaseStream.Seek(0, SeekOrigin.Begin);
            header.HeaderData = reader.ReadBytes((int)headerLen);
            header.Blocks = ReadEntries(ref reader, header);
            header.Magic = reader.ReadInt32();
            return header;
        }
        private static BlockText[] ReadEntries(ref BinaryReader reader, Header header)
        {
            BlockText[] result = new BlockText[header.TextCount];
            for (int i = 0; i < result.Length; i++)
            {
                result[i].IdLen = reader.ReadInt32();
                result[i].Id = Encoding.ASCII.GetString(reader.ReadBytes(result[i].IdLen - 1));
                reader.BaseStream.Position++;
                result[i].StrLen = reader.ReadInt32();
                if (result[i].StrLen < 0)
                {
                    result[i].StrLen = (int)(result[i].StrLen ^ 0xFFFFFFFF) * 2;
                    result[i].Str = Encoding.Unicode.GetString(reader.ReadBytes(result[i].StrLen));
                    reader.BaseStream.Position += 2;
                }
                else if (result[i].StrLen > 0)
                {
                    result[i].Str = Encoding.UTF8.GetString(reader.ReadBytes(result[i].StrLen - 1));
                    reader.BaseStream.Position++;
                }
                result[i].InfoCount = reader.ReadInt32();
                ReadTextInfos(ref reader, ref result[i]);
            }
            return result;
        }
        private static void ReadTextInfos(ref BinaryReader reader, ref BlockText block)
        {
            block.TextInfos = new TextInfo[block.InfoCount];
            for (int i = 0; i < block.InfoCount; i++)
            {
                block.TextInfos[i].Type = reader.ReadInt32();
                block.TextInfos[i].Unk = reader.ReadInt32();
                block.TextInfos[i].InfoLen = reader.ReadInt32();
                if (block.TextInfos[i].InfoLen < 0)
                {
                    block.TextInfos[i].InfoLen = (int)(block.TextInfos[i].InfoLen ^ 0xFFFFFFFF) * 2;
                    block.TextInfos[i].Info = Encoding.Unicode.GetString(reader.ReadBytes(block.TextInfos[i].InfoLen));
                    reader.BaseStream.Position += 2;
                }
                else if (block.TextInfos[i].InfoLen > 0)
                {
                    block.TextInfos[i].Info = Encoding.UTF8.GetString(reader.ReadBytes(block.TextInfos[i].InfoLen - 1));
                    reader.BaseStream.Position++;
                }
            }
        }
        private static Header[] GetHeaders(string dir)
        {
            string[] files = Directory.GetFiles(dir, "*.uexp", SearchOption.TopDirectoryOnly);
            List<Header> headers = new List<Header>();
            for (int i = 0; i < files.Length; i++)
            {
                using (FileStream stream = File.Open(files[i], FileMode.Open, FileAccess.Read))
                {
                    BinaryReader reader = new BinaryReader(stream);
                    Header header = new Header();
                    header = ReadHeader(ref reader);
                    header.FileName = Path.GetFileNameWithoutExtension(files[i]);
                    if (header.TextCount > 0) headers.Add(header);
                }
            }
            return headers.ToArray();
        }
        public static void ExportXLSX(string dir, string output)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage excel = new ExcelPackage();

            Header[] headers = GetHeaders(dir);
            foreach (var header in headers)
            {
                var workSheet = excel.Workbook.Worksheets.Add(header.FileName);
                workSheet.Cells.Style.WrapText = true;
                workSheet.Cells[1, 1].Value = "ID";
                workSheet.Cells[1, 2].Value = "Original";
                workSheet.Cells[1, 3].Value = "TextInfo";
                workSheet.Cells[1, 4].Value = "Translation";
                workSheet.Cells[1, 5].Value = "Translation (TextInfo)";
                int recordIndex = 2;
                foreach (var block in header.Blocks)
                {
                    workSheet.Cells[recordIndex, 1].Value = block.Id;
                    workSheet.Cells[recordIndex, 2].Value = block.Str;
                    workSheet.Cells[recordIndex, 3].Value = string.Join("\n", block.TextInfos.Select(t => t.Info).ToArray());
                    recordIndex++;
                }
                workSheet.Cells["A1:E1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells.Style.Font.Name = "Consolas";
                workSheet.Cells.Style.Font.Size = 12;
                workSheet.Column(1).Width = 30;
                workSheet.Column(2).Width = 80;
                workSheet.Column(3).Width = 30;
                workSheet.Column(4).Width = 80;
                workSheet.Column(5).Width = 30;
            }
            FileInfo xlsx = new FileInfo(output);
            excel.SaveAs(xlsx);
        }

        public static void ImportXLSX(string xlsx, string originalDir, string outputDir)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            FileInfo xlsxInfo = new FileInfo(xlsx);
            ExcelPackage excel = new ExcelPackage(xlsxInfo);
            Header[] headers = GetHeaders(originalDir);
            if (!Directory.Exists(outputDir)) Directory.CreateDirectory(outputDir);
            foreach (var worksheet in excel.Workbook.Worksheets)
            {
                if (!headers.Any(h => h.FileName == worksheet.Name)) continue;
                Header header = Array.Find(headers, h => h.FileName == worksheet.Name);
                string uexpPath = Path.Combine(outputDir, $"{header.FileName}.uexp");
                //int rows = worksheet.Dimension.Rows;
                int rowIndex = 2;
                using (MemoryStream stream = new MemoryStream())
                {
                    using (BinaryWriter writer = new BinaryWriter(stream))
                    {
                        writer.Write(header.HeaderData);
                        foreach (var block in header.Blocks)
                        {
                            writer.Write(block.IdLen);
                            writer.Write(Encoding.UTF8.GetBytes(block.Id));
                            writer.BaseStream.Position++;
                            string str = block.Str;
                            if (worksheet.Cells[rowIndex, 4].Value != null && worksheet.Cells[rowIndex, 4].Value.ToString().Length > 0)
                            {
                                str = worksheet.Cells[rowIndex, 4].Value.ToString();
                            }
                            if (str != null && str.Length > 0)
                            {
                                int strLen = (int)(str.Length ^ 0xFFFFFFFF);
                                writer.Write(strLen);
                                writer.Write(Encoding.Unicode.GetBytes(str));
                                writer.BaseStream.Position += 2;
                            }
                            else writer.BaseStream.Position += 4;
                            writer.Write(block.InfoCount);
                            if (block.InfoCount > 0)
                            {
                                List<string> infoStr = new List<string>();
                                if (worksheet.Cells[rowIndex, 5].Value != null && worksheet.Cells[rowIndex, 5].Value.ToString().Length > 0)
                                {
                                    string[] infos = worksheet.Cells[rowIndex, 5].Value.ToString().Split((char)10);
                                    foreach (var info in infos) infoStr.Add(info);
                                }
                                for (int i = 0; i < block.TextInfos.Length; i++)
                                {
                                    writer.Write(block.TextInfos[i].Type);
                                    writer.Write(block.TextInfos[i].Unk);
                                    string info = block.TextInfos[i].Info;
                                    if (info != null && info.Length > 0)
                                    {
                                        if (i < infoStr.Count) info = infoStr[i];
                                        int infoLen = (int)(info.Length ^ 0xFFFFFFFF);
                                        writer.Write(infoLen);
                                        writer.Write(Encoding.Unicode.GetBytes(info));
                                        writer.BaseStream.Position += 2;
                                    }
                                    else writer.BaseStream.Position += 4;
                                }
                            }
                            rowIndex++;
                        }
                        writer.Write(header.Magic);
                        string uassetPath = Path.Combine(outputDir, $"{header.FileName}.uasset");
                        File.Copy(Path.Combine(originalDir, $"{header.FileName}.uasset"), uassetPath, true);
                        using (BinaryWriter uassetWr = new BinaryWriter(File.Open(uassetPath, FileMode.Open, FileAccess.Write)))
                        {
                            uassetWr.BaseStream.Seek(uassetWr.BaseStream.Length - 92, SeekOrigin.Begin);
                            uassetWr.Write((int)writer.BaseStream.Length - 4);
                        }
                    }
                    File.WriteAllBytes(uexpPath, stream.ToArray());
                    Console.WriteLine("Repacked: {0}", header.FileName);
                }
            }
        }
    }
}
