using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace PFUIEventServerSample
{
    /// <summary>
    /// This Class us used for Binding Commands From Model to View
    /// </summary>
    public class BaseCommand : ICommand
    {
        public BaseCommand(Action<object> action)
            : this(null, action)

        { }

        public BaseCommand(Predicate<object> canExecute, Action<object> action)
        {
            _canExecute = canExecute;
            _executeAction = action;
        }
        Predicate<object> _canExecute;
        Action<object> _executeAction;

        public bool CanExecute(object parameter)
        {
            return _canExecute == null ? true : _canExecute(parameter);
        }
        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }



        public void Execute(object parameter)
        {
            if (_executeAction != null)
                _executeAction(parameter);
        }

    }
}
