using System;
using System.Collections.Generic;
using System.Data;

namespace ImportarExcel
{
    public class Dados
    {
        public int Classe { get; set; }
        public int Conta { get; set; }
        public string Data { get; set; }
        public decimal Valor { get; set; }
        public string Descricao { get; set; }

        public Dados() { }

        public Dados(int classe, int conta, string data, decimal valor, string descricao)
        {
            Classe = classe;
            Conta = conta;
            Data = data;
            Valor = valor;
            Descricao = descricao;
        }

        public List<Dados> ObterDados()
        {
            var listaDados = new List<Dados>();
            var dt = Excel.ObterDados();

            foreach (DataRow item in dt.Rows)
            {
                listaDados.Add(new Dados(int.Parse(item["classe"].ToString()), int.Parse(item["conta"].ToString()), item["data"].ToString(), Convert.ToDecimal(item["valor"]), item["descricao"].ToString()));
            }

            return listaDados;
        }

        public bool AdicionarDados()
        {
            return DataBase.AdicionarDados(this);
        }
    }
}
