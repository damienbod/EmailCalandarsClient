// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using EmailCalendarsClient.MailSender;
using Microsoft.Identity.Client;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
// The following using statements were added for this sample.
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace TodoListClient
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        AadNativeClient _aadNativeClient = new AadNativeClient();
        // Button content
        const string SignInString = "Sign In";
        const string ClearCacheString = "Clear Cache";

        public MainWindow()
        {
            InitializeComponent();
            _aadNativeClient.InitClient();
            GetGraphData();
        }

        private async void SignIn(object sender = null, RoutedEventArgs args = null)
        {
            var accounts = await _aadNativeClient.GetAccountsAsync();

            // If there is already a token in the cache, clear the cache and update the label on the button.
            if (SignInButton.Content.ToString() == ClearCacheString)
            {
                TodoList.ItemsSource = string.Empty;

                await _aadNativeClient.RemoveAccountsAsync();

                SignInButton.Content = SignInString;
                UserName.Content = Properties.Resources.UserNotSignedIn;
                return;
            }

            try
            {
                var account = await _aadNativeClient.SignIn();

                Dispatcher.Invoke(() =>
                {
                    SignInButton.Content = ClearCacheString;
                    SetUserName(account);
                    GetGraphData();
                });
            }
            catch (MsalException ex)
            {
                if (ex.ErrorCode == "access_denied")
                {
                    // The user canceled sign in, take no action.
                }
                else
                {
                    // An unexpected error occurred.
                    string message = ex.Message;
                    if (ex.InnerException != null)
                    {
                        message += "Error Code: " + ex.ErrorCode + "Inner Exception : " + ex.InnerException.Message;
                    }

                    MessageBox.Show(message);
                }

                Dispatcher.Invoke(() =>
                {
                    UserName.Content = Properties.Resources.UserNotSignedIn;
                });
            }
        }

        private void GetGraphData()
        {
            GetGraphDataAsync(SignInButton.Content.ToString() != ClearCacheString).ConfigureAwait(false);
        }

        private async Task GetGraphDataAsync(bool isAppStarting)
        {
            var accounts = await _aadNativeClient.GetAccountsAsync();
            if (!accounts.Any())
            {
                SignInButton.Content = SignInString;
                return;
            }

            // Get an access token to call the To Do service.
            AuthenticationResult result = null;
            try
            {
                result = await _aadNativeClient.AcquireTokenSilent();

                Dispatcher.Invoke(
                    () =>
                    {
                        SignInButton.Content = ClearCacheString;
                        SetUserName(result.Account);
                    });
            }
            // There is no access token in the cache, so prompt the user to sign-in.
            catch (MsalUiRequiredException)
            {
                if (!isAppStarting)
                {
                    MessageBox.Show("Please sign in to view your To-Do list");
                    SignInButton.Content = SignInString;
                }
            }
            catch (MsalException ex)
            {
                // An unexpected error occurred.
                string message = ex.Message;
                if (ex.InnerException != null)
                {
                    message += "Error Code: " + ex.ErrorCode + "Inner Exception : " + ex.InnerException.Message;
                }
                MessageBox.Show(message);

                UserName.Content = Properties.Resources.UserNotSignedIn;
                return;
            }

            // Once the token has been returned by MSAL, add it to the http authorization header, before making the call to access the To Do list service.
            //_httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", result.AccessToken);

            // Call the To Do list service.
            //HttpResponseMessage response = await _httpClient.GetAsync(TodoListApiAddress);

            //if (response.IsSuccessStatusCode)
            //{
            //    // Read the response and data-bind to the GridView to display To Do items.
            //    string s = await response.Content.ReadAsStringAsync();
            //    List<TodoItem> toDoArray = JsonConvert.DeserializeObject<List<TodoItem>>(s);

            //    Dispatcher.Invoke(() =>
            //    {
            //        TodoList.ItemsSource = toDoArray.Select(t => new { t.Title });
            //    });
            //}
            //else
            //{
            //    await DisplayErrorMessage(response);
            //}
        }

        private static async Task DisplayErrorMessage(HttpResponseMessage httpResponse)
        {
            string failureDescription = await httpResponse.Content.ReadAsStringAsync();
            if (failureDescription.StartsWith("<!DOCTYPE html>"))
            {
                string path = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
                string errorFilePath = Path.Combine(path, "error.html");
                File.WriteAllText(errorFilePath, failureDescription);
                Process.Start(errorFilePath);
            }
            else
            {
                MessageBox.Show($"{httpResponse.ReasonPhrase}\n {failureDescription}", "An error occurred while getting /api/todolist", MessageBoxButton.OK);
            }
        }

        private async void AddTodoItem(object sender, RoutedEventArgs e)
        {
            var accounts = await _aadNativeClient.GetAccountsAsync();

            if (!accounts.Any())
            {
                MessageBox.Show("Please sign in first");
                return;
            }
            if (string.IsNullOrEmpty(TodoText.Text))
            {
                MessageBox.Show("Please enter a value for the To Do item name");
                return;
            }

            // Get an access token to call the To Do service.
            AuthenticationResult result = null;
            try
            {
                result = await _aadNativeClient.AcquireTokenSilent();

                Dispatcher.Invoke(() =>
                {
                    SetUserName(result.Account);
                    UserName.Content = Properties.Resources.UserNotSignedIn;
                });
            }
            // There is no access token in the cache, so prompt the user to sign-in.
            catch (MsalUiRequiredException)
            {
                MessageBox.Show("Please re-sign");
                SignInButton.Content = SignInString;
            }
            catch (MsalException ex)
            {
                // An unexpected error occurred.
                string message = ex.Message;
                if (ex.InnerException != null)
                {
                    message += "Error Code: " + ex.ErrorCode + "Inner Exception : " + ex.InnerException.Message;
                }

                Dispatcher.Invoke(() =>
               {
                   UserName.Content = Properties.Resources.UserNotSignedIn;
                   MessageBox.Show("Unexpected error: " + message);
               });

               return;
            }

            //
            // Call the To Do service.
            //

            // Once the token has been returned by MSAL, add it to the http authorization header, before making the call to access the To Do service.
            //_httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", result.AccessToken);

            // Forms encode Todo item, to POST to the todo list web api.
            //TodoItem todoItem = new TodoItem() { Title = TodoText.Text };
            //string json = JsonConvert.SerializeObject(todoItem);
            //StringContent content = new StringContent(json, Encoding.UTF8, "application/json");

            //// Call the To Do list service.

            //HttpResponseMessage response = await _httpClient.PostAsync(TodoListApiAddress, content);

            //if (response.IsSuccessStatusCode)
            //{
            //    TodoText.Text = "";
            //    GetTodoList();
            //}
            //else
            //{
            //        await DisplayErrorMessage(response);
            //}
        }

        // Set user name to text box
        private void SetUserName(IAccount userInfo)
        {
            string userName = null;

            if (userInfo != null)
            {
                userName = userInfo.Username;
            }

            if (userName == null)
                userName = Properties.Resources.UserNotIdentified;

            UserName.Content = userName;
        }
    }
}
