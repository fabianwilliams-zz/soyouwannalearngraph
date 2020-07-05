using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web;

namespace graph_tutorial.Helpers
{
    public static class GroupHelper
    {
        public static async Task<string> GetGroupIdAsync(string groupName)
        {
            var graphClient = GraphHelper.GetAuthenticatedClient();

            var groups = await graphClient.Groups
                .Request()
                .Filter($"displayname eq '{groupName}'")
                .GetAsync();

            if (groups.Count == 0)
            {
                throw new ServiceException(new Error
                {
                    Code = GraphErrorCode.ItemNotFound.ToString(),
                    Message = $"Group named: '{groupName}' is not found"
                });
            }

            return groups.CurrentPage[0].Id;

        }

        public static async Task<IList<Conversation>> GetGroupConversationsAsync(string groupID)
        {
            var graphClient = GraphHelper.GetAuthenticatedClient();

            //https://graph.microsoft.com/v1.0/groups/{id}/conversations
            var conversations = await graphClient.Groups[groupID]
                                                 .Conversations
                                                 .Request()
                                                 .GetAsync();

            return conversations.CurrentPage;
        }

        public static async Task<Conversation> GetGroupConversation(string groupId, string conversationId)
        {
            var graphClient = GraphHelper.GetAuthenticatedClient();

            Conversation conversation = await graphClient.Groups[groupId]
                                                .Conversations[conversationId]
                                                .Request()
                                                .Expand("threads")
                                                .GetAsync();

            return conversation;
        }

        public static async Task<string> CreateProject(GraphServiceClient graphClient, string groupId, string conversationId)
        {
            User me = await graphClient.Me.Request().GetAsync();

            Conversation conversation = await GetGroupConversation(groupId, conversationId);

            PlannerTask task = await CreatePlanner(graphClient, groupId, conversation, me);

            DriveItem copiedFile = await DuplicateCostSpreadsheet(graphClient, groupId, $"{task.Title}.xlsx");

            await UpdateTaskDetails(graphClient, groupId, task, copiedFile);

            await ReplyToCustomer(graphClient, groupId, conversation, me);

            return task.Title;
        }

        private static async Task<PlannerTask> CreatePlanner(GraphServiceClient graphClient, string groupId, Conversation conversation, User me)
        {
            string taskTitle = $"{DateTime.Now.ToString("yyyyMMddHHmmss")}-{conversation.Topic}";

            PlannerPlan plan = await GetPlannerPlanAsync(graphClient, groupId, "maintenance plan");

            string bucketId = await GetBucketIdAsync(graphClient, plan, "Requests");

            var newTask = new PlannerTask()
            {
                PlanId = plan.Id,
                ConversationThreadId = conversation.Id,
                BucketId = bucketId,
                Title = taskTitle
            };

            newTask.Assignments = new PlannerAssignments();
            newTask.Assignments.AddAssignee(me.Id); //Assign the current logged in user

            var createdTask = await graphClient.Planner.Tasks.Request().AddAsync(newTask);

            return createdTask;
        }

        private static async Task<PlannerPlan> GetPlannerPlanAsync(GraphServiceClient graphClient, string groupId, string planName)
        {
            var plans = await graphClient.Groups[groupId]
                .Planner
                .Plans
                .Request()
                .Filter($"title eq '{planName}'")
                .GetAsync();

            if (plans.Count == 0)
            {
                throw new ServiceException(new Error
                {
                    Code = GraphErrorCode.ItemNotFound.ToString(),
                    Message = $"Plan: {planName} was not found"
                });
            }

            return plans[0];
        }

        private static async Task<string> GetBucketIdAsync(GraphServiceClient graphClient, PlannerPlan plan, string bucketName)
        {
            //Note: Buckets does not support $filter, so need to use linq
            var planBuckets = await graphClient.Planner
                .Plans[plan.Id]
                .Buckets
                .Request()
                .GetAsync();
            return planBuckets.FirstOrDefault(bucket => bucket.Name == bucketName).Id;
        }

        private static async Task UpdateTaskDetails(GraphServiceClient graphClient, string groupId, PlannerTask task, DriveItem attachFile)
        {
            PlannerTaskDetails taskDetails = await GetPlannerTaskDetailsAsync(graphClient, task);

            taskDetails.References = new PlannerExternalReferences();
            taskDetails.References.AddReference(attachFile.WebUrl, attachFile.Name);

            taskDetails.Checklist = new PlannerChecklistItems();
            taskDetails.Checklist.AddChecklistItem("Schedule Time with client");

            var updatedTask = await graphClient.Planner.Tasks[task.Id].Details.Request()
                .Header("If-Match", taskDetails.GetEtag())
                .Header("Prefer", "return=represendation")
                .UpdateAsync(taskDetails);
        }

        private static async Task<PlannerTaskDetails> GetPlannerTaskDetailsAsync(GraphServiceClient graphClient, PlannerTask task)
        {
            int cnt = 0;
            PlannerTaskDetails taskDetails = null;

            while (taskDetails == null)
            {
                //Sometimes takes a little time to create the task, so wait until the item is created
                cnt++;
                try
                {
                    taskDetails = await graphClient.Planner.Tasks[task.Id].Details.Request().GetAsync();
                }
                catch (ServiceException se)
                {

                }

            }

            return taskDetails;
        }

        private static async Task<DriveItem> DuplicateCostSpreadsheet(GraphServiceClient graphClient, string groupId, string newName)
        {
            string projectFolderId = await GroupHelper.GetDriveFolderIdAsync(graphClient, groupId, "Projects");

            string budgetTemplateName = "FRM_Budget_Base.xlsx";
            DriveItem budgetTemplate = await GroupHelper.GetDriveFileAsync(graphClient, groupId, projectFolderId, budgetTemplateName);

            await graphClient.Groups[groupId]
                .Drive
                .Items[budgetTemplate.Id]
                .Copy(newName)
                .Request()
                .PostAsync();

            //Note: While the PostAsync says that it returns a DriveItem, it is currently returning null
            //      So to retrieve the copied file requires querying for it

            DriveItem copiedItem = await GetDriveFileAsync(graphClient, groupId, projectFolderId, newName);

            return copiedItem;
        }

        private static async Task<string> GetDriveFolderIdAsync(GraphServiceClient graphClient, string groupId, string folderName)
        {
            var options = new[]
            {
                new QueryOption("$filter", $"name eq '{folderName}'")
            };
            var groupDrive = await graphClient.Groups[groupId].Drive.Root.Children.Request(options).GetAsync();

            if (groupDrive.Count == 0)
            {
                throw new ServiceException(new Error
                {
                    Code = GraphErrorCode.ItemNotFound.ToString(),
                    Message = "Projects Drive folder was not found"
                });
            }

            return groupDrive.CurrentPage[0].Id;
        }

        private static async Task<DriveItem> GetDriveFileAsync(GraphServiceClient graphClient, string groupId, string projectFolderId, string budgetTemplateName)
        {
            var options = new[]
            {
                new QueryOption("$filter", $"name eq '{budgetTemplateName}'")
            };

            var driveItems = await graphClient.Groups[groupId].Drive.Items[projectFolderId].Children.Request(options).GetAsync();
            if (driveItems.Count == 0)
            {
                throw new ServiceException(new Error
                {
                    Code = GraphErrorCode.ItemNotFound.ToString(),
                    Message = $"Budget Template: {budgetTemplateName} was not found"
                });
            }

            return driveItems[0];
        }

        private static async Task ReplyToCustomer(GraphServiceClient graphClient, string groupId, Conversation conversation, User me)
        {
            var firstSender = conversation.UniqueSenders.FirstOrDefault();
            string replyMessage = $@"
            Hello {firstSender},

            Thank you for reaching out to the Maintenance Department. 

            We wanted to let you know that {me.DisplayName} has been assigned to your request, 
            and should be contacting you shortly to schedule a time to stop by.

            Thank you,
            Maintenance Team
                            ";

            var post = new Post() { Body = new ItemBody() { Content = replyMessage, ContentType = BodyType.Text } }; //TODO Improve body text 

            var thread = conversation.Threads.FirstOrDefault();
            await graphClient.Groups[groupId].Threads[thread.Id].Reply(post).Request().PostAsync();
        }
    }
}