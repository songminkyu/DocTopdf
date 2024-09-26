using DocToPdf.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocToPdf.ViewModel
{
    public class MainViewModel : MainViewModelBase
    {
        public MainViewModel()
        {
            LoggingService.LoggingInit();
        }
    }
}
