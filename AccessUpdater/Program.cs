using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.IO;
using System.Threading;

namespace AccessUpdater
{
    class Program
    {
        public static readonly string BASE = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source={0}; Jet OLEDB:Database Password={1}";
        static void Main(string[] args)
        {
            Console.WriteLine("Access Updater performs SQL-Commands on a database from a .sql-File");
            Console.WriteLine("Please provide a valid access database path:\n");
            Console.Write("> ");

            string dbPath = Console.ReadLine();

            Console.WriteLine("\nProvide a password if neccessary:\n");
            Console.Write("> ");

            string pw = Console.ReadLine();

            string fullConString = string.Format(BASE, @dbPath, @pw);
            if (pw.Equals(""))
                fullConString = string.Join(";", fullConString.Split(';'), 0, 2) + "; Persist Security Info=False;";

            Console.WriteLine("\nThe connection string seems to be: {0}", @fullConString);
            Console.WriteLine("Trying to establish a connection to the database...");

            OleDbConnection dbCon = null;
            try
            {
                dbCon = new OleDbConnection(@fullConString);
                dbCon.Open();
            }
            catch (OleDbException ode)
            {
                Console.WriteLine("An error appeared while connecting to the database: " + ode.Message);
                Console.ReadLine();
                return;
            }

            Console.WriteLine("Connection successfully established!\n");
            Console.WriteLine("Beginning main loop\n");
            Thread.Sleep(1000);

            do
            {
                Console.Clear();
                Console.WriteLine("Please provide a path to a .sql-file which should be executed on the database:\n");
                Console.Write("> ");

                string sqlFilePath = Console.ReadLine();

                if (!File.Exists(@sqlFilePath))
                {
                    Console.Write("\nThe provided file doesn't exist. Do you want to provide another file (y/n)? > ");
                    continue;
                }

                string fullFile = File.ReadAllText(@sqlFilePath, Encoding.UTF8).Trim();
                string[] commands = fullFile.Split(';');

                Console.WriteLine();
                foreach (string command in commands)
                {
                    if (command.Equals(""))
                        continue;

                    if (command.ToLower().Contains("insert into"))
                    {
                        string insertBase = command.Split(new string[] { "VALUES", "values", "Values" }, StringSplitOptions.RemoveEmptyEntries)[0] + "VALUES";
                        string[] values = command.Split(new string[] { "VALUES", "values", "Values" }, StringSplitOptions.RemoveEmptyEntries)[1].Split('(');

                        foreach(string value in values)
                        {
                            if (value.Trim().Equals(""))
                                continue;
                            try
                            {
                                OleDbCommand cmd = new OleDbCommand(insertBase + " (" + value.Split(')')[0] + ");", dbCon);
                                cmd.ExecuteNonQuery();
                            }
                            catch (OleDbException ode)
                            {
                                Console.WriteLine("An error appeared while executing the command: {0}; please check the syntax and try again: {1}", command, ode.Message);
                            }
                        }
                        Console.WriteLine("Command successfully executed, {0} Row(s) affected\n", values.Length - 1);
                    }
                    else
                    {
                        try
                        {
                            OleDbCommand cmd = new OleDbCommand(command, dbCon);
                            int cmdResult = cmd.ExecuteNonQuery();
                            if (cmdResult == -1)
                                Console.WriteLine("Command successful executed\n");
                            else
                                Console.WriteLine("Command successfully executed, {0} Row(s) affected\n", cmdResult);
                        }
                        catch (OleDbException ode)
                        {
                            Console.WriteLine("An error appeared while executing the command: {0}; please check the syntax and try again: {1}", command, ode.Message);
                        }
                    }

                }
                Console.Write("\nAll commands executed. Do you want to provide another file (y/n)? > ");
            } while (Console.ReadLine().ToLower().Equals("y"));

            dbCon.Close();
        }
    }
}
