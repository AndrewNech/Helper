using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;

namespace Helper.ViewModel
{
    public class MainViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        public void OnPropertyChanged(string i="")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(i));
        }
        private int count = 0;
        public int Count { get { OnPropertyChanged(); return count++; } }

    }
}
