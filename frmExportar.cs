using System;
using System.Windows.Forms;

namespace ImportarExcel
{
    public partial class frmExportar : Form
    {
        public frmExportar()
        {
            InitializeComponent();
        }

        private void btnSalvar_Click(object sender, EventArgs e)
        {
            string txt;
            var dados = new Dados();
            var listaDados = dados.ObterDados();

            foreach (var item in listaDados)
            {
                txt = item.Classe.ToString() + "\t" + item.Conta.ToString() + "\t" + item.Data.ToString() + "\t" + item.Valor.ToString() + "\t" + item.Descricao.ToString();

                Console.WriteLine(txt);
                
                //dados = item;
                //if (!dados.AdicionarDados())
                //    MessageBox.Show(DataBase.MsgErro);
            }

            MessageBox.Show("Processo realizado com sucesso!");

        }
    }
}
