// Guids.cs
// MUST match guids.h

using System;

namespace Community.VSBrowserSwitcher
{
    static class GuidList
    {
		public const string guidCommunity_VSBrowserSwitcherPkgString = "31F6453E-5BC0-412A-B6F4-10B55AB41DEA";
        //public const string guidCommunity_VSBrowserSwitcherCmdSetString = "73578247-4bd1-4d12-bc66-162165a43a22";
    	public const string guidIDEToolbarCmdSet = "8570C04B-5EC8-453B-BB29-E7E384B840BD";

		public static readonly Guid guidCommunity_VSBrowserSwitcherCmdSet = new Guid(guidIDEToolbarCmdSet);
    };
}