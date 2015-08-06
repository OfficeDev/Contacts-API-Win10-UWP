using Microsoft.Office365.OutlookServices;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using Windows.UI.Xaml.Media.Imaging;

namespace ContactsWin10App.Models
{
    public class ContactModel : INotifyPropertyChanged
    {
        string _displayName = null;
        public string DisplayName
        {
            get { return _displayName; }
            set { this.SetValue<string>(ref _displayName, value); }
        }

        string _email = null;
        public string Email
        {
            get { return _email; }
            set { this.SetValue<string>(ref _email, value); }
        }

        BitmapImage _thumbnail = new BitmapImage(new Uri("ms-appx:///assets/nopic.png"));
        public BitmapImage Thumbnail
        {
            get { return _thumbnail; }
            set { this.SetValue<BitmapImage>(ref _thumbnail, value); }
        }

        private void SetValue<T>(ref T field, T newValue, [CallerMemberName] string propertyName="")
        {
            if(object.Equals(field, newValue) == false)
            {
                field = newValue;
                if (PropertyChanged != null)
                    PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
    }

    public static class Extensions
    {
        public static ObservableCollection<ContactModel> ToContactModelList(this List<IContact> contacts)
        {
            return new ObservableCollection<ContactModel>(contacts.Select(i => new ContactModel()
            {
                DisplayName = i.DisplayName,
                Email = i.EmailAddresses[0].Address
            }).ToList());
        }
    }
}
