using Microsoft.Graph;
using graph_tutorial.Helpers;
using graph_tutorial.Models;
using System;
using System.Threading.Tasks;
using System.Web.Mvc;
using System.Collections.Generic;

namespace graph_tutorial.Controllers
{
    public class RequestsController : BaseController
    {
        [Authorize]
        public async Task<ActionResult> Index()
        {
            try
            {

                GraphServiceClient graphClient = GraphHelper.GetAuthenticatedClient();

                //TODO: Retrieve Conversations to be displayed
                //  https://graph.microsoft.com/v1.0/groups?filter=displayname eq 'maintenance'
                var maintGroupId = await GroupHelper.GetGroupIdAsync("Visitor Intake");

                //https://graph.microsoft.com/v1.0/groups/{id}/conversations
                IList<Conversation> maintConversations = await GroupHelper.GetGroupConversationsAsync(maintGroupId);

                ViewBag.GroupId = maintGroupId;
                return View(maintConversations);

            }
            catch (ServiceException se)
            {
                return RedirectToAction("Index", "Error", new { message = Request.RawUrl + ": " + se.Error.Message });
            }
        }

        [Authorize]
        public async Task<ActionResult> Conversation(string groupId, string conversationId)
        {
            try
            {
                GraphServiceClient graphClient = GraphHelper.GetAuthenticatedClient();

                Conversation conversation = await GroupHelper.GetGroupConversation(groupId, conversationId);

                ViewBag.GroupId = groupId;
                return View(conversation);
            }
            catch (ServiceException se)
            {
                return RedirectToAction("Index", "Error", new { message = Request.RawUrl + ": " + se.Error.Message });
            }
        }

        [Authorize]
        public async Task<ActionResult> CreateProject(string groupId, string conversationId)
        {
            try
            {
                GraphServiceClient graphClient = GraphHelper.GetAuthenticatedClient();
                string taskTitle = await GroupHelper.CreateProject(graphClient, groupId, conversationId);

                ViewBag.taskTitle = taskTitle;
                return View("VisitAuthorized");
            }
            catch (ServiceException se)
            {
                return RedirectToAction("Index", "Error", new { message = Request.RawUrl + ": " + se.Error.Message });
            }

        }

    }
}