using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Net;
using System.Threading;
using System.Threading.Tasks;
using Windows.UI.Core;
using Windows.UI.Popups;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util;

// The Blank Page item template is documented at http://go.microsoft.com/fwlink/?LinkId=402352&clcid=0x409

namespace LargeEmailsGmail
{
    public class EmailMessage
    {
        public string Id { get; set; }
        public bool IsSelected { get; set; }
        public string Snippet { get; set; }
        public string SizeEstimate { get; set; }
        public string From { get; set; }
        public string Date { get; set; }
        public string Subject { get; set; }
    }

    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainPage : Page
    {
        private UserCredential _credential;
        const string AppName = "LargeMail";

        public MainPage()
        {
            this.InitializeComponent();
            GetCredential();
        }

        public async Task<UserCredential> GetCredential()
        {
            var scopes = new[] { GmailService.Scope.GmailModify };
            var uri = new Uri("ms-appx:///client_id.json");
            _credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                uri, scopes, "user", CancellationToken.None);
            return _credential;
        }

        private async void GetMessagesClick(object sender, RoutedEventArgs e)
        {
            var service = new GmailService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = _credential,
                ApplicationName = AppName,
            });
            var sizeEstimate = 0L;
            IList<Message> messages = null;
            var emailMessages = new List<EmailMessage>();
            OperationText.Text = "downloading messages";
            BusyBorder.Visibility = Visibility.Visible;
            await Task.Run(async () =>
            {
                UsersResource.MessagesResource.ListRequest request =
                service.Users.Messages.List("me");
                request.Q = "larger:5M";
                request.MaxResults = 1000;
                messages = request.Execute().Messages;
                await Dispatcher.RunAsync(CoreDispatcherPriority.Normal, () =>
                    ProgressBar.Maximum = messages.Count);

                for (int index = 0; index < messages.Count; index++)
                {
                    var message = messages[index];
                    var getRequest = service.Users.Messages.Get("me", message.Id);
                    getRequest.Format =
                        UsersResource.MessagesResource.GetRequest.FormatEnum.Metadata;
                    getRequest.MetadataHeaders = new Repeatable<string>(
                        new[] { "Subject", "Date", "From" });
                    messages[index] = getRequest.Execute();
                    sizeEstimate += messages[index].SizeEstimate ?? 0;
                    emailMessages.Add(new EmailMessage()
                    {
                        Id = messages[index].Id,
                        Snippet = WebUtility.HtmlDecode(messages[index].Snippet),
                        SizeEstimate = $"{messages[index].SizeEstimate:n0}",
                        From = messages[index].Payload.Headers.FirstOrDefault(h => 
                            h.Name == "From").Value,
                        Subject = messages[index].Payload.Headers.FirstOrDefault(h => 
                            h.Name == "Subject").Value,
                        Date = messages[index].Payload.Headers.FirstOrDefault(h => 
                            h.Name == "Date").Value,
                    });
                    var index1 = index+1;
                    await Dispatcher.RunAsync(CoreDispatcherPriority.Normal, () =>
                    {
                        ProgressBar.Value = index1;
                        DownloadText.Text = $"{index1} of {messages.Count}";
                    });
                }
            });
            BusyBorder.Visibility = Visibility.Collapsed;
            MessagesList.ItemsSource = new ObservableCollection<EmailMessage>(
                emailMessages.OrderByDescending(m => m.SizeEstimate));
            CountText.Text = $"{messages.Count} messages. Estimated size: {sizeEstimate:n0}";
        }

        private async void DeleteMessagesClick(object sender, RoutedEventArgs e)
        {
            var messages = (ObservableCollection<EmailMessage>) MessagesList.ItemsSource;
            var messagesToDelete = messages.Where(m => m.IsSelected).ToList();
            if (!messagesToDelete.Any())
            {
                await (new MessageDialog("There are no selected messages to delete")).ShowAsync();
                return;
            }
            var service = new GmailService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = _credential,
                ApplicationName = AppName,
            });
            OperationText.Text = "deleting messages";
            ProgressBar.Maximum = messagesToDelete.Count;
            DownloadText.Text = "";
            BusyBorder.Visibility = Visibility.Visible;
            var sizeEstimate = messages.Sum(m => Convert.ToInt64(m.SizeEstimate));
            await Task.Run(async () =>
            {
                for (int index = 0; index < messagesToDelete.Count; index++)
                {
                    var message = messagesToDelete[index];
                    var response = service.Users.Messages.Trash("me", message.Id);
                    response.Execute();
                    var index1 = index+1;
                    await Dispatcher.RunAsync(CoreDispatcherPriority.Normal, () =>
                    {
                        ProgressBar.Value = index1;
                        DownloadText.Text = $"{index1} of {messagesToDelete.Count}";
                        messages.Remove(message);
                        sizeEstimate -= Convert.ToInt64(message.SizeEstimate);
                        CountText.Text = $"{messages.Count} messages. Estimated size: {sizeEstimate:n0}";
                    });
                }
            });
            BusyBorder.Visibility = Visibility.Collapsed;
        }
    }
}
