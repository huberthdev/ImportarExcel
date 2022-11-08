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
            var dados = new ;
            var listaDados = dados.ObterDados();

            foreach (var item in listaDados)
            {
                dados = item;
                if (!dados.AdicionarDados())
                    MessageBox.Show(DataBase.MsgErro);
            }

            MessageBox.Show("Processo realizado com sucesso!");

        }
    }
}
