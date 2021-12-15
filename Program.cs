using System;
using System.Diagnostics;
using ClosedXML.Excel;

namespace Desafio_06
{
    class Program
    {
        public class Aluno
        {
            public string Nome;
            public int Idade;
            public double Nota;
        }
        static void Main(string[] args)
        {
            List<Aluno> listaluno = new List<Aluno>();
            double somalist = 0;

            for (int i = 0; i < 3; i++)
            {
                Aluno aluno = new Aluno();

                Console.WriteLine($"Informe o nome do aluno:");
                aluno.Nome = Console.ReadLine();
                Console.WriteLine();

                Console.WriteLine($"Informe a idade do aluno:");
                aluno.Idade = int.Parse(Console.ReadLine());
                Console.WriteLine();

                Console.WriteLine($"Informe a nota do aluno:");
                aluno.Nota = double.Parse(Console.ReadLine());
                Console.WriteLine();

                listaluno.Add(aluno);
                somalist = somalist + listaluno[i].Nota;

            }

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Planilha1");
                worksheet.Cell("A1").Value = "NOMES";
                worksheet.Cell("A2").Value = listaluno[0].Nome;
                worksheet.Cell("A3").Value = listaluno[1].Nome;
                worksheet.Cell("A4").Value = listaluno[2].Nome;

                worksheet.Cell("B1").Value = "IDADE";
                worksheet.Cell("B2").Value = listaluno[0].Idade;
                worksheet.Cell("B3").Value = listaluno[1].Idade;
                worksheet.Cell("B4").Value = listaluno[2].Idade;

                worksheet.Cell("C1").Value = "NOTA";
                worksheet.Cell("C2").Value = listaluno[0].Nota;
                worksheet.Cell("C3").Value = listaluno[1].Nota;
                worksheet.Cell("C4").Value = listaluno[2].Nota;

                worksheet.Cell("D1").Value = "Soma das Notas";
                worksheet.Cell("D2").Value = somalist;

                workbook.SaveAs(@"d:\ubuntu\testeexcel.xlsx");
            }

            Process.Start(new ProcessStartInfo(@"d:\ubuntu\testeexcel.xlsx") { UseShellExecute = true });
        }
    }
}