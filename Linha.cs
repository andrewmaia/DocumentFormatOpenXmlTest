using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ConsoleApp1
{
    public class Linha
    {
        private readonly Row _workRow;
        public  Linha(Row workRow){
            _workRow = workRow;
        }

        public Row WorkRow {
            get{return _workRow;}
        }        

        public uint  Index {
            get{ return _workRow.RowIndex; }
            set{  _workRow.RowIndex=value; }            
        }

        public void AdicionarCelula(object valor)
        {
            Cell cell = new()
            {
                CellValue = new CellValue(Helper.ResolverValor(valor)),
                DataType = Helper.ResolverTipo(valor.GetType())
            };
            _workRow.Append(cell);
        }

        public void AlterarCelula(object valor,int index,uint? estilo=null ){
            Cell cell= (Cell)_workRow.ChildElements[index];
            cell.CellValue = new CellValue(Helper.ResolverValor(valor));
            cell.DataType = Helper.ResolverTipo(valor.GetType());
            if(estilo.HasValue)
                cell.StyleIndex=estilo;
        }

    }
}