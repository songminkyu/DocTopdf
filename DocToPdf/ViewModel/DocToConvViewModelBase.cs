using CommunityToolkit.Mvvm.Input;
using DocToPdf.Model;
using DocToPdf.Services;
using System.Collections.ObjectModel;

namespace DocToPdf.ViewModel
{
    public class DocToConvViewModelBase : BindableBase
    {

        private IRelayCommand<object>? _UserControlLoadedCommand { get; set; }
        public IRelayCommand<object>? UserControlLoadedCommand
        {
            get { return _UserControlLoadedCommand; }
            set
            {
                if (_UserControlLoadedCommand == value) return;
                _UserControlLoadedCommand = value;
            }
        }
        private IRelayCommand<object>? _targetPathCommand { get; set; }
        public IRelayCommand<object>? targetPathCommand
        {
            get { return _targetPathCommand; }
            set
            {
                if (_targetPathCommand == value) return;
                _targetPathCommand = value;
            }
        }
        private IRelayCommand<object>? _savedPathCommand { get; set; }
        public IRelayCommand<object>? savedPathCommand
        {
            get { return _savedPathCommand; }
            set
            {
                if (_savedPathCommand == value) return;
                _savedPathCommand = value;
            }
        }
        private IAsyncRelayCommand<object>? _runCommand { get; set; }
        public IAsyncRelayCommand<object>? runCommand
        {
            get { return _runCommand; }
            set
            {
                if (_runCommand == value) return;
                _runCommand = value;
            }
        }
        private IRelayCommand<object>? _cancelCommand { get; set; }
        public IRelayCommand<object>? cancelCommand
        {
            get { return _cancelCommand; }
            set
            {
                if (_cancelCommand == value) return;
                _cancelCommand = value;
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
        private ObservableCollection<DocConverter>? _convLogs { get; set; }
        public ObservableCollection<DocConverter>? convLogs
        {
            get { return _convLogs; }
            set
            {
                if (_convLogs == value) return;
                _convLogs = value;

                OnPropertyChanged(nameof(convLogs));
            }
        }
    }
}
