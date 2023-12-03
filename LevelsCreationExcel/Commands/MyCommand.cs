using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace LevelsCreationExcel.Commands
{
    public class MyCommand : ICommand
    {
        public event EventHandler CanExecuteChanged;

        private readonly Action<Object> _Excute;
        private readonly Predicate<Object> _CanExcute;

        //Main Constructor
        public MyCommand(Action<Object> excute, Predicate<Object> CanExcute)
        {
            _Excute = excute;

            _CanExcute = CanExcute;
        }

        public bool CanExecute(object parameter)
        {
            return _CanExcute(parameter);
        }

        public void Execute(object parameter)
        {
            _Excute(parameter);
        }
    }
}
