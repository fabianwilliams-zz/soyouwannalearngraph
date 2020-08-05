using Microsoft.Graph;
using Microsoft.Graph.Auth;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JustGetMail
{
    class Program
    {
        static async Task Main(string[] args)
        {

            IPublicClientApplication app = PublicClientApplicationBuilder
                                    .Create("b42bded7-1ea9-466e-8457-0c0685e578f8")
                                    .WithRedirectUri("https://login.microsoftonline.com/common/oauth2/nativeclient")
                                    .Build();

            var authProvider = new InteractiveAuthenticationProvider(app, new[] { "User.Read", "Mail.Read" });
            //var authProvider = new InteractiveAuthenticationProvider(app, new[] { "User.Read"});

            var client = new GraphServiceClient(authProvider);

            //var currUser = await client.Me.Request().GetAsync();
            //Console.WriteLine($"Current User is: { currUser.DisplayName}");

            var clientMessages = await client.Me.Messages.Request().Select("sender, subject").GetAsync();        
            foreach ( Message m in clientMessages.CurrentPage)
                {
                Console.WriteLine($"Sender is: {m.Sender} and the subject is {m.Subject}");
                }
            Console.ReadLine();

        }
    }
}
