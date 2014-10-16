MDX-Runner
==========

Tools for building and running MDX &amp; XMLA queries

*MDX Runner - Copyright (C) 2014 Andrew Prendergast & VizDynamics.*

Licensed under the GNU GPL v3.
Author contact: ap@vizdynamics.com or http://blog.andrewprendergast.com/
More info and latest version at mdxrunner.org


Example Invocation:
-------------------
 
    MDXRunner oMDXRunner = new MDXRunner();
    IEnumerable<Dictionary<string,object>> listGrains = 
        oMDXRunner
            .setCube("[PerformanceCube]")
            .addWith("MEMBER [total] AS [Measures].[Clicks] + [Measures].[Actions]")
            .addDimension("[Advertising Dim].[Campaign Name].[Campaign Name]")
            .addDimension("[Calendar Time Dim].[Month].[Month]")
            .addMetric("[Measures].[Clicks]")
            .addMetric("[Measures].[Actions]")
            .addMetric("[Measures].[total]")
            .addFilter(new string[] { "[Advertising Dim].[Advertiser Name].&[937]", "[Advertising Dim].[Advertiser Name].&[1075]" })
            .executeMDX<object>("http://ssasserver.somecompany.com/ssas/msmdpump.dll", "Corp DW")
            ;
    Debug.WriteLine(oMDXRunner.getMDX());
    Debug.WriteLine(listGrains);


Example output from getMDX():
-----------------------------
 
    WITH MEMBER [total] AS [Measures].[Clicks] + [Measures].[Actions]
    SELECT
    NON EMPTY ([Advertising Dim].[Campaign Name].[Campaign Name],[Calendar Time Dim].[Month].[Month]) ON ROWS,
    NON EMPTY {[Measures].[Clicks],[Measures].[Actions],[Measures].[total]} ON COLUMNS
    FROM (SELECT {[Advertising Dim].[Advertiser Name].&[937],[Advertising Dim].[Advertiser Name].&[1075]} ON 0 FROM (SELECT FROM [PerformanceCube]))

Example output from executeMDX():
---------------------------------

    {[CampaignName, Affiliate];[_CampaignName, [Advertising Dim].[Campaign Name].&[Affiliate]];[Month, March 2009];[_Month, [Calendar Time Dim].[Month].&[2009]&[1]&[3]];[Clicks, 5];[_Clicks, [Measures].[Clicks]];[Actions, 10];[_Actions, [Measures].[Actions]];[total, 15];[_total, [Measures].[total]]}
    {[CampaignName, Affiliate];[_CampaignName, [Advertising Dim].[Campaign Name].&[Affiliate]];[Month, April 2009];[_Month, [Calendar Time Dim].[Month].&[2009]&[2]&[4]];[Clicks, 1318];[_Clicks, [Measures].[Clicks]];[Actions, 32];[_Actions, [Measures].[Actions]];[total, 1350];[_total, [Measures].[total]]}
    {[CampaignName, Affiliate];[_CampaignName, [Advertising Dim].[Campaign Name].&[Affiliate]];[Month, May 2009];[_Month, [Calendar Time Dim].[Month].&[2009]&[2]&[5]];[Clicks, 2011];[_Clicks, [Measures].[Clicks]];[Actions, 54];[_Actions, [Measures].[Actions]];[total, 2065];[_total, [Measures].[total]]}
    {[CampaignName, Affiliate];[_CampaignName, [Advertising Dim].[Campaign Name].&[Affiliate]];[Month, June 2009];[_Month, [Calendar Time Dim].[Month].&[2009]&[2]&[6]];[Clicks, 730];[_Clicks, [Measures].[Clicks]];[Actions, 43];[_Actions, [Measures].[Actions]];[total, 773];[_total, [Measures].[total]]}
    {[CampaignName, Affiliate];[_CampaignName, [Advertising Dim].[Campaign Name].&[Affiliate]];[Month, July 2009];[_Month, [Calendar Time Dim].[Month].&[2009]&[3]&[7]];[Clicks, 6366];[_Clicks, [Measures].[Clicks]];[Actions, 40];[_Actions, [Measures].[Actions]];[total, 6406];[_total, [Measures].[total]]}
    {[CampaignName, BAU Campaigns];[_CampaignName, [Advertising Dim].[Campaign Name].&[BAU Campaigns]];[Month, April 2009];[_Month, [Calendar Time Dim].[Month].&[2009]&[2]&[4]];[Clicks, 44337];[_Clicks, [Measures].[Clicks]];[Actions, 111];[_Actions, [Measures].[Actions]];[total, 44448];[_total, [Measures].[total]]}
    {[CampaignName, BAU Campaigns];[_CampaignName, [Advertising Dim].[Campaign Name].&[BAU Campaigns]];[Month, May 2009];[_Month, [Calendar Time Dim].[Month].&[2009]&[2]&[5]];[Clicks, 84654];[_Clicks, [Measures].[Clicks]];[Actions, 172];[_Actions, [Measures].[Actions]];[total, 84826];[_total, [Measures].[total]]}
    {[CampaignName, BAU Campaigns];[_CampaignName, [Advertising Dim].[Campaign Name].&[BAU Campaigns]];[Month, June 2009];[_Month, [Calendar Time Dim].[Month].&[2009]&[2]&[6]];[Clicks, 73386];[_Clicks, [Measures].[Clicks]];[Actions, 188];[_Actions, [Measures].[Actions]];[total, 73574];[_total, [Measures].[total]]}
    {[CampaignName, BAU Campaigns];[_CampaignName, [Advertising Dim].[Campaign Name].&[BAU Campaigns]];[Month, July 2009];[_Month, [Calendar Time Dim].[Month].&[2009]&[3]&[7]];[Clicks, 59243];[_Clicks, [Measures].[Clicks]];[Actions, 145];[_Actions, [Measures].[Actions]];[total, 59388];[_total, [Measures].[total]]}
    {[CampaignName, Non-Campaign];[_CampaignName, [Advertising Dim].[Campaign Name].&[Non-Campaign]];[Month, May 2009];[_Month, [Calendar Time Dim].[Month].&[2009]&[2]&[5]];[Clicks, 0];[_Clicks, [Measures].[Clicks]];[Actions, 293];[_Actions, [Measures].[Actions]];[total, 293];[_total, [Measures].[total]]}
    {[CampaignName, Non-Campaign];[_CampaignName, [Advertising Dim].[Campaign Name].&[Non-Campaign]];[Month, June 2009];[_Month, [Calendar Time Dim].[Month].&[2009]&[2]&[6]];[Clicks, 0];[_Clicks, [Measures].[Clicks]];[Actions, 705];[_Actions, [Measures].[Actions]];[total, 705];[_total, [Measures].[total]]}
    {[CampaignName, Non-Campaign];[_CampaignName, [Advertising Dim].[Campaign Name].&[Non-Campaign]];[Month, July 2009];[_Month, [Calendar Time Dim].[Month].&[2009]&[3]&[7]];[Clicks, 0];[_Clicks, [Measures].[Clicks]];[Actions, 151];[_Actions, [Measures].[Actions]];[total, 151];[_total, [Measures].[total]]}

