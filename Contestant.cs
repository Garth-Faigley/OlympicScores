using System.ComponentModel;

namespace OlympicScores
{
    public class Contestant : INotifyPropertyChanged
    {
        private string _name;
        private int _score;

        public string Name
        {
            get { return _name; }
            set { _name = value; OnPropertyChanged("Name"); }
        }

        public int Score
        {
            get { return _score; }
            set { _score = value; OnPropertyChanged("Score"); }
        }


        // boilerplate code supporting PropertyChanged events
        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(string name)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null)
            {
                handler(this, new PropertyChangedEventArgs(name));
            }
        }
    }
}
