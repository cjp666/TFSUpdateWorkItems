using System;
using System.Collections.Generic;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;

namespace TFSUpdateWorkItems
{
	public class UpdateWorkItems
	{
		private const string OldSharePointSite = "http://oldsharepoint/sites/";
		private const string NewSharePointSite = "http://newsharepoint/sites/";

		/// <summary>
		///		When set to true the code just counts the number of
		///		external links that point to the old SharePoint site
		/// </summary>
		private const bool TestingOnly = false;

		private int _linksToUpdateCount = 0;

		/// <summary>
		///		Loop through the Product Backlog Items for a project
		///		and update the work item links from the old SharePoint
		///		site to the new one
		/// </summary>
		public void DoUpdate()
		{
			Console.WriteLine("");
			Console.WriteLine("Updating...");

			var collectionUri = new Uri("http://fshtfs:8080/tfs/FSH");

			// Connect to the server and the store.
			using (var teamProjectCollection = new TfsTeamProjectCollection(collectionUri))
			{
				var workItemStore = teamProjectCollection.GetService<WorkItemStore>();

				RegisteredLinkType storyboardType = null;
				foreach (RegisteredLinkType registeredLinkType in workItemStore.RegisteredLinkTypes)
				{
					if (String.Compare(registeredLinkType.Name, "Storyboard", StringComparison.CurrentCultureIgnoreCase) == 0)
					{
						storyboardType = registeredLinkType;
						break;
					}
				}

				// create a couple of test external links to be able to update them
				// CreateTestLinks(workItemStore, storyboardType, 1495);

				var q = "SELECT [System.Id] FROM WorkItems WHERE [System.TeamProject] = 'SiclopsLIVE' ORDER BY [System.Id]";
				var workItems = workItemStore.Query(q);
				foreach (WorkItem workItem in workItems)
				{
					if (!TestingOnly && workItem.Id != 1495)
					{
						continue;
					}

					if (String.Compare(workItem.Type.Name, "Product Backlog Item", StringComparison.CurrentCultureIgnoreCase) == 0)
					{
						var e = workItem.ExternalLinkCount;
						var c = workItem.Attachments.Count;
						Console.WriteLine("{0} - {1} - {2} - {3} - {4}", workItem.Id, workItem.Type.Name, workItem.Title, c, e);
						if (c > 0)
						{
							foreach (Attachment attachment in workItem.Attachments)
							{
								if (!attachment.Uri.ToString().StartsWith("http://fshtfs", StringComparison.InvariantCultureIgnoreCase))
								{
									// TODO: still need to do something about these, not important for now
									Console.WriteLine("\t{0}", attachment.Uri);
								}
							}
						}

						if (e > 0)
						{
							UpdateWorkItemLinks(workItemStore, storyboardType, workItem.Id);
						}
					}
				}
			}

			Console.WriteLine("Count of links to update {0}", _linksToUpdateCount);
		}

		/// <summary>
		///		Create a couple of external links on the work item that
		///		point to the old SharePoint site to test them being updated
		/// </summary>
		private void CreateTestLinks(WorkItemStore workItemStore, RegisteredLinkType storyboardType, int workItemId)
		{
			var workItem = workItemStore.GetWorkItem(workItemId);

			var newUri = "vstfs:///Requirements/Storyboard/" + Uri.EscapeDataString(String.Format("{0}Shared%20Documents/Storyboards/1418%20Calls%20-%20Call%20Monitor%20Search%20Entity.docx", OldSharePointSite));
			var externalLink = new ExternalLink(storyboardType, newUri);
			workItem.Links.Add(externalLink);

			newUri = "vstfs:///Requirements/Storyboard/" + Uri.EscapeDataString(String.Format("{0}Shared%20Documents/Storyboards/Define%2520Global%2520Search%2520Queries.pptx", OldSharePointSite));
			externalLink = new ExternalLink(storyboardType, newUri);
			workItem.Links.Add(externalLink);

			workItem.Save();
		}

		/// <summary>
		///		Loop through any links for 'gateway' and change them to the new sharepoint link
		/// </summary>
		private void UpdateWorkItemLinks(WorkItemStore workItemStore, RegisteredLinkType storyboardType, int workItemId)
		{
			var newLinks = new List<string>();

			WorkItem workItem;
			bool linkFound;

			do
			{
				// loop through the external links:
				// * keeping a track of the storyboard documents on the old SharePoint site
				// * remove the old link
				// * add a new link pointing to the same document on the new SharePoint site
				linkFound = false;
				workItem = workItemStore.GetWorkItem(workItemId);
				foreach (var externalLink in workItem.Links)
				{
					var el = externalLink as ExternalLink;
					if (el != null)
					{
						var uri = el.LinkedArtifactUri;
						if (uri.StartsWith("vstfs:///Requirements/Storyboard/http", StringComparison.CurrentCultureIgnoreCase))
						{
							var u = Uri.UnescapeDataString(uri);
							u = u.Replace("%20", " ");
							Console.WriteLine("\t{0}", u);
							if (u.Contains(OldSharePointSite))
							{
								u = u.Substring(33);
								var newUri = "vstfs:///Requirements/Storyboard/" + Uri.EscapeDataString(u.Replace(OldSharePointSite, NewSharePointSite));
								newLinks.Add(newUri);

								_linksToUpdateCount++;

								if (!TestingOnly)
								{
									workItem.Links.Remove(el);
									workItem.Save();

									linkFound = true;
									break;
								}
							}
						}
					}
				}

				if (TestingOnly)
				{
					break;
				}
			} while (linkFound);

			if (TestingOnly || newLinks.Count == 0)
			{
				return;
			}

			// now add the links that point to the new SharePoint
			workItem = workItemStore.GetWorkItem(workItemId);
			foreach (var newLink in newLinks)
			{
				var externalLink = new ExternalLink(storyboardType, newLink);
				workItem.Links.Add(externalLink);
			}
			workItem.Save();
		}
	}
}