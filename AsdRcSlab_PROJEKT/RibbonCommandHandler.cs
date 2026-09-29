using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.Windows;
using System;
using System.Windows.Input;

namespace AsdRcSlab
{
    public class RibbonCommandHandler : ICommand
    {
        private readonly string _command;

        public RibbonCommandHandler(string command)
        {
            _command = command;
        }

        public bool CanExecute(object parameter) => true;

        public void Execute(object parameter)
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            doc.SendStringToExecute(_command + " ", true, false, true);
        }

        public event EventHandler CanExecuteChanged
        {
            add { }
            remove { }
        }
    }
}
