using CommunityToolkit.Mvvm.Input;
using DocToPdf.Services;

namespace DocToPdf.ViewModel
{
    public class DocToConvViewModelBase : BindableBase
    {

        private IAsyncRelayCommand<object>? _UserControlLoadedCommand { get; set; }
        public IAsyncRelayCommand<object>? UserControlLoadedCommand
        {
            get { return _UserControlLoadedCommand; }
            set
            {
                if (_UserControlLoadedCommand == value) return;
                _UserControlLoadedCommand = value;
            }
        }
    }
}
