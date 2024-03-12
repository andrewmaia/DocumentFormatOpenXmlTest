using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office.CustomUI;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;


namespace ConsoleApp1
{
    public class EscritorExcel: IDisposable
    {
        private readonly MemoryStream _novoArquivo;
        private readonly SpreadsheetDocument _document; 
        private WorksheetPart _worksheetPart;

        public EscritorExcel(string enderecoTemplate,string nomePlanilha="")
        {
            _novoArquivo = new();
            CopiarTemplate(enderecoTemplate, _novoArquivo);
            _document = SpreadsheetDocument.Open(_novoArquivo, true); 
            _worksheetPart =  ObterWorkSheet(nomePlanilha);
        }
        private void CopiarTemplate(string enderecoTemplate, MemoryStream ms)
        {
            using SpreadsheetDocument document = SpreadsheetDocument.Open(enderecoTemplate, false);
            document.Clone(ms);
            document.Close();
        }

        private WorksheetPart ObterWorkSheet( string nomePlanilha)
        {
            IEnumerable<Sheet>? sheets = _document.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == nomePlanilha || string.IsNullOrEmpty(nomePlanilha));

            if (!sheets.Any())
                throw new Exception($"Não existe a planilha de nome {nomePlanilha}");
 
            string idPlanilha = sheets.First().Id.Value;
            return  (WorksheetPart) _document.WorkbookPart.GetPartById(idPlanilha);
        }
        public void AlterarWorksheet(string nomePlanilha)
        {
            _worksheetPart =  ObterWorkSheet(nomePlanilha);            
        }

        #region Interação com Celulas

        public void AlterarValorCelula(string nomeCelula,object novoValor){
            Cell? celula = _worksheetPart.Worksheet.Descendants<Cell>()?.Where(c => c.CellReference == nomeCelula).FirstOrDefault();
            celula.DataType = new EnumValue<CellValues>(Helper.ResolverTipo(novoValor.GetType())); 
            string valor = Helper.ResolverValor(novoValor);
            celula.CellValue= new CellValue(valor);    
        }
        #endregion
        
        #region Interação com Linhas

        public Linha ObterLinha(string nomeCelulaReferencia){
            Cell? celulaReferencia = _worksheetPart.Worksheet.Descendants<Cell>()?.Where(c => c.CellReference == nomeCelulaReferencia).FirstOrDefault();
            Row linhaReferencia= (Row)celulaReferencia.Parent;  
            return new Linha(linhaReferencia);
        }
        public Linha CriarLinha(uint index){
            Row row= new(){
                RowIndex = index
            };
            return new Linha(row);   
        }
        public void AdicionarLinha(Linha linha){
            SheetData sheetData = _worksheetPart.Worksheet.GetFirstChild<SheetData>();
            sheetData.InsertAt(linha.WorkRow,int.Parse(linha.WorkRow.RowIndex));
            //sheetData.Append(linha.WorkRow);
        }            
        #endregion
        public void GerarArquivo(string endereco)
        {
            _document.Save();
            _document.Dispose();
            File.WriteAllBytes(endereco,_novoArquivo.ToArray());
        }
        public void Dispose()
        {
            _document.Dispose();
            _novoArquivo.Dispose();
        }

        public Linha ClonarLinha(Linha linhaOrigem, uint indexNovaLinha){
            var sharedStringTablePart1 =  _document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();

            Row copia=new();
            var celulas = linhaOrigem.WorkRow.ChildElements.Where(x=>x is Cell && x is not null).ToList();
            var celulasTipadas = celulas.Select(c => (Cell)c).ToArray().ToList();

            Dictionary<int,string> valoresCopia= new();
            int i=0;
            foreach(var celula in celulasTipadas){
                string valor="";

                if (celula.DataType!=null)
                    if(celula.DataType == CellValues.SharedString)
                        valor=sharedStringTablePart1.SharedStringTable.ElementAt(int.Parse(celula.InnerText)).InnerText;;

                valoresCopia.Add(i,valor);           
                i++;

                Cell cell = new()
                {
                    DataType = CellValues.InlineString,
                    StyleIndex=celula.StyleIndex
                };
                copia.Append(cell);
            }

            Linha l=  new Linha(copia);
            foreach(int key in valoresCopia.Keys){
                l.AlterarCelula(valoresCopia[key],key);
            }

            l.Index=indexNovaLinha;
            AdicionarLinha(l);

            return l;
        }
    }
}