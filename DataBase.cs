using System;
using System.Data.SqlClient;

namespace ExcelBancoDeDadosSQLServer
{
    public class DataBase
    {
        private static string server = @"(localdb)\MSSQLLocalDB";
        private static string dataBase = "Financas";

        public static string MsgErro { get; private set; }

        public static string StrCon
        {
            get { return $"Data Source = {server};Initial Catalog = {dataBase}; Integrated Security = True; Connect Timeout = 30; Encrypt=False;TrustServerCertificate=False;ApplicationIntent=ReadWrite;MultiSubnetFailover=False"; }
        }

        public static bool AdicionarDados(Dados dados)
        {
            try
            {
                using(SqlConnection cn = new SqlConnection(StrCon))
                {
                    cn.Open();

                    var sql = "INSERT INTO BD(Classe, Conta, Data, Valor, Descricao) VALUES(@Classe, @Conta, @Data, @Valor, @Descricao)";

                    using(SqlCommand cmd = new SqlCommand(sql, cn))
                    {
                        cmd.Parameters.AddWithValue("@Classe", dados.Classe);
                        cmd.Parameters.AddWithValue("@Conta", dados.Conta);
                        cmd.Parameters.AddWithValue("@Data", dados.Data);
                        cmd.Parameters.AddWithValue("@Valor", dados.Valor.ToString().Replace(",", "."));
                        cmd.Parameters.AddWithValue("@Descricao", dados.Descricao);

                        cmd.ExecuteNonQuery();
                    }
                }

                MsgErro = "";
                return true;
            }
            catch (Exception ex)
            {
                MsgErro = ex.Message;
                return false;
            }
        }
    }
}
