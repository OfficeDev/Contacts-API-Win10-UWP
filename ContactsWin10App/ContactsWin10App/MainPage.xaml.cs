using Microsoft.Office365.OutlookServices;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Threading.Tasks;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.Security.Authentication.Web.Core;
using Windows.Security.Credentials;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Controls.Primitives;
using Windows.UI.Xaml.Data;
using Windows.UI.Xaml.Input;
using Windows.UI.Xaml.Media;
using Windows.UI.Xaml.Navigation;
using ContactsWin10App.Models;
using Windows.UI.Xaml.Media.Imaging;
using Windows.Storage.Streams;

// The Blank Page item template is documented at http://go.microsoft.com/fwlink/?LinkId=402352&clcid=0x409

namespace ContactsWin10App
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainPage : Page
    {
        public MainPage()
        {
            this.InitializeComponent();
            this.Loaded += MainPage_Loaded;
        }

        private async void MainPage_Loaded(object sender, RoutedEventArgs e)
        {
            var task = await Controllers.MyContactsController.GetContacts();
            var contacts = task.ToContactModelList();
            this.DataContext = contacts;
            wait.IsActive = false;

            //FUTURE: in the future we will be able to get pictures within org...can only get ME today
            //LoadImages(contacts);
        }

        private async Task LoadImages(IList<ContactModel> contacts)
        {
            foreach (var contact in contacts)
            {
                var bytes = await Controllers.MyContactsController.GetImage(contact.Email);
                if (bytes != null && bytes.Length > 0)
                {
                    using (var stream = new InMemoryRandomAccessStream())
                    {
                        await stream.WriteAsync(bytes.AsBuffer());
                        var image = new BitmapImage();
                        stream.Seek(0);
                        image.SetSource(stream);
                        contact.Thumbnail = image;
                    } 
                }
            }
        }
    }
}
