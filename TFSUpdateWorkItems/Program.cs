using System;

namespace TFSUpdateWorkItems
{
	public class Program
	{
		public static void Main(string[] args)
		{
			var t = new UpdateWorkItems();
			t.DoUpdate();

			Console.WriteLine("");
			Console.WriteLine("Press ENTER to terminate...");
			Console.ReadKey();
		}
	}
}
