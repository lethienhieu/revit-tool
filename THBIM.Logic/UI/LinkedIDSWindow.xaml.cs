using System;
using System.Collections.Generic;
using System.Windows;

namespace THBIM
{
    public partial class LinkedIDSWindow : Window
    {
        public LinkedIDSWindow(List<string> idResults)
        {
            InitializeComponent();
            myTextBox.Text = string.Join(Environment.NewLine, idResults);
        }
    }
}
