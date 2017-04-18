using Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Animation;

namespace OlympicScores
{

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private Timer _timer;
        private const string _filePath = "Olympics.xlsx";
        private int _readFileInterval = 30000;
        private int _marqueeInterval = 15;

        public MainWindow()
        {
            InitializeComponent();
            TextBlockEventTitle.Text = GetEventTitle();

            Loaded += new RoutedEventHandler(MainWindow_Loaded);
            RefreshContestants();

            _timer = new Timer(timer_Elapsed);
            _timer.Change(_readFileInterval, _readFileInterval);
        }

        private void timer_Elapsed(object state)
        {
            Dispatcher.Invoke(() =>
            {
                RefreshContestants();
            });
        }

        private void RefreshContestants()
        {
            var contestantString = string.Empty;

            var contestants = GetContestantsFromFile();
            contestantString = BuildContestantString(contestants);
            TextBlockMarquee.Text = contestantString;
        }

        private string BuildContestantString(List<Contestant> contestants)
        {
            var contestantString = string.Empty;

            if (contestants != null && contestants.Count > 1)
            {
                contestants = contestants.OrderByDescending(c => c.Score).ToList();

                int previousScore = 0;
                int rank = 0;
                foreach (var thisContestant in contestants)
                {
                    var thisScore = thisContestant.Score;

                    if (contestantString != string.Empty)
                        contestantString += Environment.NewLine;

                    if (previousScore > 0 && thisScore == previousScore)
                        contestantString += rank + "(t)";
                    else
                    {
                        rank++;
                        contestantString += rank.ToString();
                    }
                    contestantString += " - " + thisScore.ToString() + " " + thisContestant.Name;
                    previousScore = thisScore;
                }

            }
            return contestantString;
        }

        private List<Contestant> GetContestantsFromFile()
        {
            var contestants = new List<Contestant>();

            using (FileStream stream = new FileStream(_filePath,
                                  FileMode.Open,
                                  FileAccess.Read,
                                  FileShare.ReadWrite))
            {
                IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                excelReader.IsFirstRowAsColumnNames = false;

                DataSet result = excelReader.AsDataSet();
                if (result != null && result.Tables.Count > 0)
                {
                    DataTable dataTable = result.Tables[0];

                    var isContestantRow = false;

                    foreach (DataRow thisRow in dataTable.Rows)
                    {
                        string contestant = thisRow[0].ToString();

                        if (isContestantRow && contestant != string.Empty)
                        {
                            string scoreString = thisRow[6].ToString();

                            var newContestant = new Contestant();
                            newContestant.Name = contestant;
                            newContestant.Score = int.Parse(scoreString);

                            contestants.Add(newContestant);
                        }

                        if (contestant == "Contestant")
                            isContestantRow = true;
                    }
                }
            }
            return contestants;
        }


        // Assume the text in cell A1 is the event title
        private string GetEventTitle()
        {
            string title = string.Empty;

            using (FileStream stream = new FileStream(_filePath,
                                  FileMode.Open,
                                  FileAccess.Read,
                                  FileShare.ReadWrite))
            {
                IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                excelReader.IsFirstRowAsColumnNames = false;

                DataSet result = excelReader.AsDataSet();
                if (result != null && result.Tables.Count > 0)
                {
                    DataTable dataTable = result.Tables[0];

                    title = dataTable.Rows[0][0].ToString();


                }
            }
            return title;
        }


        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            DoubleAnimation doubleAnimation = new DoubleAnimation();
            doubleAnimation.From = -TextBlockMarquee.ActualHeight - 300;
            // doubleAnimation.From = -CanvasMarquee.ActualHeight - 100;
            doubleAnimation.To = CanvasMarquee.ActualHeight;
            doubleAnimation.RepeatBehavior = RepeatBehavior.Forever;
            doubleAnimation.Duration = new Duration(TimeSpan.FromSeconds(_marqueeInterval));
            TextBlockMarquee.BeginAnimation(Canvas.BottomProperty, doubleAnimation);
        }

    }
}
