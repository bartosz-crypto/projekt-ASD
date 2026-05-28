using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace AsdRcSlab
{
    public class PileData : INotifyPropertyChanged
    {
        public string PileId         { get; set; } = "";
        public double UtilPct        { get; set; }
        public string LocationType   { get; set; } = "";  // INT / EDGE / CORNER / REENTRANT
        public string PunchingAction { get; set; } = "";  // ADD H12@200 / ADD H16@200 / ADD H16@100 / NO ACTION

        private string _phAction = "";
        public string PhAction
        {
            get => _phAction;
            set
            {
                if (_phAction == value) return;
                _phAction = value;
                OnPropertyChanged();
            }
        }

        public List<string> ApplicablePileIds { get; set; } = new List<string>();
        public string DetailTitle    { get; set; } = "";

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
