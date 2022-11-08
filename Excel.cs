using System.Data;
using System.Data.OleDb;

namespace ImportarExcel
{
    public static class Excel
    {
        public static DataTable ObterDados()
        {
            var arquivo = @"C:\CursoCSharp\Exportar_Excel_Para_SQL_Server\base.xlsx";
            var planilha = "SELECT * FROM [plan_base$] ORDER BY [DATA] DESC";

            var strCon = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + arquivo + "; Extended Properties=\"Excel 12.0; HDR=Yes; IMEX=0\"";

            DataTable dt = new DataTable();

            using (OleDbConnection con = new OleDbConnection(strCon))
            {
                using (OleDbDataAdapter da = new OleDbDataAdapter(planilha, con))
                {
                    da.Fill(dt);
                }
            }

            return dt;
        }


    }
}
