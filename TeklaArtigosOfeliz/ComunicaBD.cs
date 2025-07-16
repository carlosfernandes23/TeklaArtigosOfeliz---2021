﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace TeklaArtigosOfeliz
{
    class ComunicaBDtekla
    {
        SqlConnection MiConexion = new SqlConnection("Data Source=GALILEU\\PREPARACAO;Initial Catalog=ArtigoTekla;Persist Security Info=True;User ID=SA;Password=preparacao");
        public void ConectarBD()
        {
            MiConexion.Open();
        }
        public void DesonectarBD()
        {
            MiConexion.Close();
        }
        public SqlConnection GetConnection()
        {
            return MiConexion;
        }

        public List<string> Procurarbd(string Query)
        {
            SqlCommand MiComando = new SqlCommand(Query, MiConexion);
            List<string> Result = new List<string>();

            using (SqlDataReader reader = MiComando.ExecuteReader())
            {
                while (reader.Read())
                {
                    for (int i = 0; i < reader.VisibleFieldCount; i++)
                    {
                        Result.Add(reader[i].ToString());
                    }
                }
            }
            return Result;
        }
    }

    class ComunicaBDprimavera
    {
        SqlConnection MiConexion = new SqlConnection("Data Source=TESLA\\PRIMAVERA;Initial Catalog=PRIOFELIZ;Persist Security Info=True;User ID=CM;Password=OF€l1z201");

        public void ConectarBD()
        {
            MiConexion.Open();
        }
        public void DesonectarBD()
        {
            MiConexion.Close();
        }
        public List<string> Procurarbd(string Query)
        {
            SqlCommand MiComando = new SqlCommand(Query, MiConexion);
            List<string> Result = new List<string>();

            using (SqlDataReader reader = MiComando.ExecuteReader())
            {
                while (reader.Read())
                {
                    for (int i = 0; i < reader.VisibleFieldCount; i++)
                    {
                        Result.Add(reader[i].ToString());
                    }
                }
            }
            return Result;
        }

      public bool dbOperations(List<string> dbOperations)
        {
            bool SUCESSFULL = true;

            using ( MiConexion = new SqlConnection("Data Source=TESLA\\PRIMAVERA;Initial Catalog=PRIOFELIZ;Persist Security Info=True;User ID=CM;Password=OF€l1z201"))
            {
                MiConexion.Open();
                SqlTransaction transaction = MiConexion.BeginTransaction();

                foreach (string commandString in dbOperations)
                {
                    SqlCommand cmd = new SqlCommand(commandString, MiConexion, transaction);
                    cmd.ExecuteNonQuery();
                }
                try
                {
                    transaction.Commit();
                }
                catch (Exception EX)
                {
                    transaction.Rollback();
                    SUCESSFULL = false;
                    MessageBox.Show("ERRO A SALVAR DADOS" + Environment.NewLine+EX.Message,"ERRO",MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                MiConexion.Close();
            }
            return SUCESSFULL;
        }

    }

    class ComunicaBaseDados
    {

        SqlConnection MiConexionArtigo = new SqlConnection("Data Source=GALILEU\\PREPARACAO;Initial Catalog=TempoPreparacao;Persist Security Info=True;User ID=SA;Password=preparacao");

        public void ConectarBDArtigo()
        {
            if (MiConexionArtigo.State == ConnectionState.Closed)
            {
                MiConexionArtigo.Open();
            }
        }

        public void DesonectarBDArtigo()
        {
            if (MiConexionArtigo.State == ConnectionState.Open)
            {
                MiConexionArtigo.Close();
            }
        }
        public SqlConnection GetConnection()
        {
            return MiConexionArtigo;
        }

        public DataTable ProcurarbdArtigo(string Query)
        {
            SqlCommand MiComando = new SqlCommand(Query, MiConexionArtigo);
            SqlDataAdapter dataAdapter = new SqlDataAdapter(MiComando);
            DataTable dataTable = new DataTable();

            dataAdapter.Fill(dataTable);
            return dataTable;
        }

        public List<string> ProcurarbdlistArtigo(string Query)
        {
            SqlCommand MiComando = new SqlCommand(Query, MiConexionArtigo);
            List<string> Result = new List<string>();

            using (SqlDataReader reader = MiComando.ExecuteReader())
            {
                while (reader.Read())
                {
                    for (int i = 0; i < reader.VisibleFieldCount; i++)
                    {
                        Result.Add(reader[i].ToString());
                    }
                }
            }
            return Result;
        }

        public DataTable BuscarRegistros(SqlCommand command)
        {
            SqlDataAdapter dataAdapter = new SqlDataAdapter(command);
            DataTable dataTable = new DataTable();
            dataAdapter.Fill(dataTable);
            return dataTable;
        }

    }

}
