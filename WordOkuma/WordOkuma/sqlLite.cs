using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SQLite;
using System.Data;
using Microsoft.Office.Core;
using word = Microsoft.Office.Interop.Word;

namespace WordOkuma
{
    class sqlLite
    {
        SQLiteConnection con;
        SQLiteCommand cmd;

        public string getPath()
        {
            string fileName = "veriler.";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            return path;
        }

        public void olusturDataBase()
        {
            if (!System.IO.File.Exists(getPath()))
            {
                SQLiteConnection.CreateFile(getPath());

                using (con = new SQLiteConnection("Data Source="+getPath()+";Version=3;"))
                {
                    con.Open();
                    string sql = "create table words (kelime varchar(30), font varchar(30), size varchar(20), bold varchar(20), italic varchar(20), color varchar(20))";
                    cmd = new SQLiteCommand(sql, con);
                    cmd.ExecuteNonQuery();
                    con.Close();
                }
            }
        }

        public void veriEkle(word.Document doc, int count, System.Windows.Forms.ProgressBar pB)
        {
            pB.Maximum = count;
            pB.Step = 1;
            pB.Value = 0;
            for (int i = 1; i < 3; i++)
            {
                System.GC.Collect();
                System.GC.WaitForPendingFinalizers();

                if (doc.Words[i].Text != null && doc.Words[i].Text != string.Empty)
                {
                    cmd = new SQLiteCommand();
                    con = new SQLiteConnection("Data Source=" + getPath() + ";Version=3;");
                    con.Open();
                    cmd.Connection = con;
                    cmd.CommandText = "insert into words(kelime, font, size, bold, italic, color) values (@kelime, @font, @size, @bold, @italic, @color)";
                    cmd.Parameters.AddWithValue("@kelime", doc.Words[i].Text);
                    cmd.Parameters.AddWithValue("@font", doc.Words[i].Font.Name);
                    cmd.Parameters.AddWithValue("@size", doc.Words[i].Font.Size);
                    cmd.Parameters.AddWithValue("@bold", doc.Words[i].Font.Bold);
                    cmd.Parameters.AddWithValue("@italic", doc.Words[i].Font.Italic);
                    cmd.Parameters.AddWithValue("@color", doc.Words[i].Font.Color);
                    cmd.ExecuteNonQuery();
                    con.Close();
                }
                pB.PerformStep();
            } 
        }
    }
}
