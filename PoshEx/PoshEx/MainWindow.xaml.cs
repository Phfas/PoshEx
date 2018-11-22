﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Management.Automation;
using System.Collections.ObjectModel;

namespace PoshEx {
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window {
        PowerShell poshInstance = PowerShell.Create();
        public MainWindow() {
            InitializeComponent();

        }

        private void Button_Click(object sender, RoutedEventArgs e) {
            ExecuteSynchronous(textboxScript.Text);
        }

        private void SetOutput(string output, bool setPrompt) {
            textboxOutput.AppendText($"{output}\r\n");
            textboxOutput.ScrollToEnd();

            if (setPrompt) {
                textboxOutput.Text += $"PS C:\\> ";
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e) {
            poshInstance.Dispose();

        }

        private void ExecuteSynchronous(string script) {
             
            poshInstance.AddScript($"{script} | Out-String -Stream");
            Collection<PSObject> psOutput = poshInstance.Invoke();
            SetOutput(script, false);

            foreach (PSObject item in psOutput) {
                SetOutput(item.ToString(), false);
            }
            if (poshInstance.Streams.Error.Count > 0) {
                SetOutput(poshInstance.Streams.Error.ToString(), false);
            }

            SetOutput("", true);
        }

        private void ExecuteAsynchronous(string script) {
            poshInstance.Streams.ClearStreams();
            poshInstance.Commands.Clear();
            poshInstance.AddScript($"{script} | Out-String -Stream");

            PSDataCollection<PSObject> outputCollection = new PSDataCollection<PSObject>();
            outputCollection.DataAdded += outputCollection_DataAdded;

            // the streams (error,debug,progress, etc) are available on the poshInstance
            // we can review them during or after execution
            // we can also be notified when a new item is written to the stream (like this):
            poshInstance.Streams.Error.DataAdded += Error_DataAdded;
            poshInstance.Streams.Progress.DataAdded += Progress_DataAdded;

            //begin invoke execution on the pipeline
            // use this overload to specify an output stream buffer
            IAsyncResult asyncResult = poshInstance.BeginInvoke<PSObject, PSObject>(null,outputCollection);

            SetOutput(poshInstance.Commands.Commands[0].ToString(), false);


        }

        private void Progress_DataAdded(object sender, DataAddedEventArgs e) {
            this.Dispatcher.Invoke(() => {
                var records = (PSDataCollection<ProgressRecord>)sender;
                labelProgress.Content = $"{records[e.Index].Activity.ToString()}: {records[e.Index].StatusDescription.ToString()} - {records[e.Index].PercentComplete.ToString()}% ";
            });
        }

        private void Error_DataAdded(object sender, DataAddedEventArgs e) {
            this.Dispatcher.Invoke(() => {
                var records = (PSDataCollection<ErrorRecord>)sender;
                SetOutput(records[e.Index].ToString(), false);
            });
        }

        private void outputCollection_DataAdded(object sender, DataAddedEventArgs e) {
            this.Dispatcher.Invoke(() =>
            {
                var records = (PSDataCollection<PSObject>)sender;
                SetOutput(records[e.Index].ToString(), false);
            });
        }

        private void Button_Click_1(object sender, RoutedEventArgs e) {
            ExecuteAsynchronous(textboxScript.Text);

        }

        private void Button_Click_2(object sender, RoutedEventArgs e) {
            textboxOutput.Clear();
            SetOutput("", true);
        }
    }
}