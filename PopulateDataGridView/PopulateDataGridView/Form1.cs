using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PopulateDataGridView
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            List<DTOPokemon> pokemones = new List<DTOPokemon>();
            DTOPokemon pokemon = new DTOPokemon();
            pokemon.Id = 1;
            pokemon.Nombre = "Pikachu";
            pokemon.Clase = "Eletrico";
            pokemones.Add(pokemon);

            pokemon = new DTOPokemon();
            pokemon.Id = 2;
            pokemon.Nombre = "Charmander";
            pokemon.Clase = "Fuego";
            pokemones.Add(pokemon);

            dataGridView1.DataSource = pokemones;
            dataGridView1.Refresh();

        }
    }
}
