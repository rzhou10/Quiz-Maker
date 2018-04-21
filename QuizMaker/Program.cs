using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace QuizMaker
{
    class Program
    {
        static void Main(string[] args)
        {
            int numCorrectA = 0;
            Random rand = new Random();
            int numQuestions = rand.Next(8, 12);

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook quizWorkbook = xlApp.Workbooks.Open(@"C:\Users\Class2018\Documents\Visual_Studio_2015\Projects\QuizMaker\questions_and_answers.xlsx");
            Excel._Worksheet xlWorksheet = quizWorkbook.Sheets[1];

            Excel.Range firstCol = xlWorksheet.Columns[1];
            Excel.Range seccondCol = xlWorksheet.Columns[2];

            //read columns and add them to array. I discovered that it was easier to get a random
            //element from array rather than getting a random cell.
            System.Array listVals = (System.Array)firstCol.Cells.Value;
            string[] allQuestions = listVals.OfType<object>().Select(o => o.ToString()).ToArray();
            System.Array listVals1 = (System.Array)seccondCol.Cells.Value;
            string[] allAnswers = listVals1.OfType<object>().Select(o => o.ToString()).ToArray();

            //create actual quiz
            string[] possibleQuestions = new string[numQuestions];
            string[] possibleAnswers = new string[numQuestions];
            for (int i = 0; i < numQuestions; i++)
            {
                int r = rand.Next(allQuestions.Length);
                possibleQuestions[i] = (string)allQuestions[r];
                possibleAnswers[i] = (string)allAnswers[r];
            }

            //remove duplicates
            string[] quizQuestions = possibleQuestions.Distinct().ToArray();
            string[] quizAnswers = possibleAnswers.Distinct().ToArray();

            string[] userAnswers = new string[quizQuestions.Length];
            //print out questions and user answers
            for (int i = 0; i < quizQuestions.Length; i++) {
                Console.WriteLine(quizQuestions[i]);
                userAnswers[i] = Console.ReadLine();
                Console.WriteLine();
            }
            
            //check answers by comparing them
            for (int i = 0; i < userAnswers.Length; i++)
            {
                //possible ways of inputting answers, will try to accomodate for some
                if (userAnswers[i] == quizAnswers[i] || userAnswers[i].Equals(quizAnswers[i], StringComparison.OrdinalIgnoreCase))
                {
                    numCorrectA++;
                }
            }

            double score = ((double)numCorrectA / (double)quizQuestions.Length) * 100;

            //output results
            for (int i = 0; i < userAnswers.Length; i++)
            {
                Console.WriteLine("Question: {0}", quizQuestions[i]);
                Console.WriteLine("Your Answer: {0}", userAnswers[i]);
                Console.WriteLine("Correct Answer: {0}", quizAnswers[i]);
                Console.WriteLine();
            }

            Console.WriteLine("You got {0} out of {1} correct. Your score is {2}%", numCorrectA.ToString(), quizQuestions.Length.ToString(), Math.Round(score, 2).ToString());
            
            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //kill app
            Marshal.ReleaseComObject(firstCol);
            Marshal.ReleaseComObject(seccondCol);
            Marshal.ReleaseComObject(xlWorksheet);

            quizWorkbook.Close();
            Marshal.ReleaseComObject(quizWorkbook);

            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }
    }
}