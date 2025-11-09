using Google.Apis.Auth.OAuth2;
using Google.Apis.Auth.OAuth2.Flows;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Oauth2.v2;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace CalendarSyncEngine
{
    // This class handles all the Google OAuth 2.0 logic
    public static class GoogleAuth
    {
        // This scope allows us to read, write, and delete calendar events.
        static string[] Scopes = { CalendarService.Scope.Calendar };
        static string ApplicationName = "Outlook to Google Calendar Sync";

        public static async Task<(UserCredential credential, CalendarList calendars, string email)> AuthorizeAsync()
        {
            UserCredential credential;

            // 1. Load the client_secret.json file
            using (var stream = new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                // 2. Define where to store the token.json (the user's login)
                string credPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "token.json", ".credentials/calendar-sync.json");

                // 3. This is the magic call. It will:
                //    - Look for a stored token.json file.
                //    - If not found, it will open a browser for the user to log in.
                //    - After login, it saves the token (with the Refresh Token) to the credPath.
                credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)
                );
            }

            // 4. Now that we are logged in, create the Calendar Service
            var service = new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // 5. Get the list of all calendars the user owns
            var calendarList = await service.CalendarList.List().ExecuteAsync();

            // 6. Get the users email address
            string email = "Unknown User";
            try
            {
                var oauthService = new Oauth2Service(new BaseClientService.Initializer
                {
                    HttpClientInitializer = credential
                });
                var userInfo = await oauthService.Userinfo.Get().ExecuteAsync();
                email = userInfo.Email;
            }
            catch (Exception ex)
            {
                Logger.Warning($"Could not get user email: {ex.Message}");
            }

            return (credential, calendarList, email);
        }

        public static async Task<CalendarService> GetCalendarServiceAsync(string refreshToken)
        {
            UserCredential credential;

            using (var stream = new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "token.json", ".credentials/calendar-sync.json");

                var clientSecrets = GoogleClientSecrets.FromStream(stream).Secrets;
                var token = new Google.Apis.Auth.OAuth2.Responses.TokenResponse { RefreshToken = refreshToken };

                credential = new UserCredential(new GoogleAuthorizationCodeFlow(
                    new GoogleAuthorizationCodeFlow.Initializer
                    {
                        ClientSecrets = clientSecrets,
                        Scopes = Scopes,
                        DataStore = new FileDataStore(credPath, true)
                    }),
                    "user", // This "user" ID must match the one in AuthorizeAsync
                    token);

                // This will use the refresh token to get a new access token if needed
                await credential.RefreshTokenAsync(CancellationToken.None);
            }

            return new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
        }
    }
}

