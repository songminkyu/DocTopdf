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
        private IRelayCommand<object>? _TargetPathCommand { get; set; }
        public IRelayCommand<object>? TargetPathCommand
        {
            get { return _TargetPathCommand; }
            set
            {
                if (_TargetPathCommand == value) return;
                _TargetPathCommand = value;
            }
        }
        private IRelayCommand<object>? _SavedPathCommand { get; set; }
        public IRelayCommand<object>? SavedPathCommand
        {
            get { return _SavedPathCommand; }
            set
            {
                if (_SavedPathCommand == value) return;
                _SavedPathCommand = value;
            }
        }


        private string? _targetPath { get; set; }
        public string? targetPath
        {
            get { return _targetPath; }
            set
            {
                if (_targetPath == value) return;
                _targetPath = value;

                OnPropertyChanged(nameof(targetPath));
            }
        }

        private string? _savedPath { get; set; }
        public string? savedPath
        {
            get { return _savedPath; }
            set
            {
                if (_savedPath == value) return;
                _savedPath = value;

                OnPropertyChanged(nameof(savedPath));
            }
        }
    }
}
