using System.ComponentModel;
using System.Windows.Media.Imaging;

namespace TeampptAddin
{
    public class HeaderAsset : INotifyPropertyChanged
    {
        public int Index { get; set; }
        public string Name { get; set; }
        public string PptxPath { get; set; }
        public string ThumbnailPath { get; set; }

        private BitmapImage _thumbnail;
        public BitmapImage Thumbnail
        {
            get => _thumbnail;
            set
            {
                _thumbnail = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Thumbnail)));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
    }
}
