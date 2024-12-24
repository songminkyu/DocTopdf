using DocToPdf.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocToPdf.Model
{
    public class DocConverter : BindableBase
    {
        private int _index = 0;
        public int index
        {
            get => _index;
            set => SetProperty(ref _index, value);
        }
        private string _description = "";
        public string description
        {
            get => _description;
            set => SetProperty(ref _description, value);
        }

    }
}
