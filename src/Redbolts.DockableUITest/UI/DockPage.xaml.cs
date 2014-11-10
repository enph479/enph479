﻿using System;
using System.ComponentModel;
using System.Diagnostics;

namespace Redbolts.DockableUITest.UI
{
    /// <summary>
    /// Interaction logic for DockPage.xaml
    /// </summary>
    public partial class DockPage 
    {
        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(propertyName));
        }

        public string Counter
        {
            get { return _counter; }
            set
            {
                if (value == _counter) return;
                OnPropertyChanged("Counter");
                _counter = value;
            }
        }
        private string _counter;
        public DockPage()
        {
            InitializeComponent();
            Counter = "0";
            MyGrid.DataContext = this;
        }

        private void Button_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            Counter = Globals.CounterString;
            Debug.WriteLine(Globals.CounterString);
        }
    }
}