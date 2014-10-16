/*
MDX Runner - Copyright (C) 2014 Andrew Prendergast & VizDynamics
Author contact: ap@vizdynamics.com, http://blog.andrewprendergast.com/
More info and latest version at mdxrunner.org

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <http://www.gnu.org/licenses/>.
*/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web.UI;
using Microsoft.AnalysisServices.AdomdClient;
using System.Text.RegularExpressions;
using System.Data;
using System.Diagnostics;
using System.Collections;
using System.Security.Cryptography;
using System.Data.OleDb;

/**

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

**/

namespace com.andrewprendergast.XMLA
{
    public class MDXFilter
    {
        public List<List<string>> listMemberSets = new List<List<string>>();
        public List<List<string>> listExcludedFilters = new List<List<string>>();
    }

    public class MDXRunner
    {
        private string sInstanceName = null;
        private List<string> listWith = new List<string>();
        private List<string> listDimensions = new List<string>();
        private List<string> listMetrics = new List<string>();
        private LinkedList<MDXFilter> listFilters = new LinkedList<MDXFilter>();
        private bool bDisableExecuteExceptions = false;
        private bool bNonEmptyRows = true;
        private bool bNonEmptyColumns = true;
        static Regex oRegexCompiled = new Regex("(\\[(([^\\]]|\\]\\])*)\\]\\.?)*(\\&\\[(([^\\]]|\\]\\])*)\\])*(\\.?(.*))", RegexOptions.Compiled);

        public MDXRunner()
        {
        }

        public MDXRunner clone()
        {
            return new MDXRunner()
            {
                sInstanceName = sInstanceName,
                listWith = new List<string>(listWith),
                listDimensions = new List<string>(listDimensions),
                listMetrics = new List<string>(listMetrics),
                listFilters = new LinkedList<MDXFilter>(listFilters),
                bDisableExecuteExceptions = bDisableExecuteExceptions,
                bNonEmptyRows = bNonEmptyRows,
                bNonEmptyColumns = bNonEmptyColumns
            };
        }

        public MDXRunner setCube(string sInstanceName)
        {
            this.sInstanceName = sInstanceName;
            return this;
        }

        public MDXRunner setNonEmptyRows(bool bNonEmptyBehaviour)
        {
            this.bNonEmptyRows = bNonEmptyBehaviour;
            return this;
        }

        public MDXRunner setNonEmptyColumns(bool bNonEmptyBehaviour)
        {
            this.bNonEmptyColumns = bNonEmptyBehaviour;
            return this;
        }

        public MDXRunner setNonEmptyBehaviour(bool bNonEmptyBehaviour)
        {
            this.bNonEmptyColumns = bNonEmptyBehaviour;
            this.bNonEmptyRows = bNonEmptyBehaviour;
            return this;
        }

        public MDXRunner addWith(string sWith)
        {
            this.listWith.Add(sWith);
            return this;
        }

        public MDXRunner addFilter(MDXFilter oMDXFilter)
        {
            this.listFilters.AddLast(oMDXFilter);
            return this;
        }

        public MDXRunner addFilter(string sMember)
        {
            if (sMember == null)
                return this;
            MDXFilter oFilter = new MDXFilter();
            oFilter.listMemberSets.Add(new List<string>(new string[] { sMember }));
            this.listFilters.AddLast(oFilter);
            return this;
        }

        public MDXRunner addFilter(IEnumerable<string> listMembers)
        {
            if (listMembers == null || listMembers.Count() == 0)
                return this;
            MDXFilter oFilter = new MDXFilter();
            oFilter.listMemberSets.Add(new List<string>(listMembers));
            this.listFilters.AddLast(oFilter);
            return this;
        }

        public MDXRunner addFilters(IEnumerable<IEnumerable<string>> listlistMembers)
        {
            foreach (IEnumerable<string> listMembers in listlistMembers)
                addFilter(listMembers);
            return this;
        }

        public MDXRunner addFilter(params IEnumerable<string>[] acollMembers)
        {
            if ( acollMembers == null || acollMembers.Count() == 0 )
                return this;
            MDXFilter oFilter = new MDXFilter();
            foreach ( IEnumerable<string> asMembers in acollMembers )
                oFilter.listMemberSets.Add(new List<string>(asMembers));
            this.listFilters.AddLast(oFilter);
            return this;
        }

        public MDXRunner excludeFromLastFilter(IEnumerable<string> listMembers)
        {
            if ( this.listFilters.Count > 0 )
                this.listFilters.Last.Value.listExcludedFilters.Add(listMembers.ToList<string>());
            return this;
        }

        public MDXRunner addDimension(string sAttribute)
        {
            this.listDimensions.Add(sAttribute);
            return this;
        }

        public MDXRunner addDimensions(IEnumerable<string> listAttributes)
        {
            this.listDimensions.AddRange(listAttributes);
            return this;
        }

        public MDXRunner addMetric(string sAttribute)
        {
            this.listMetrics.Add(sAttribute);
            return this;
        }

        public string getMDX()
        {
            StringBuilder oSB = new StringBuilder();

            if (this.listDimensions.Count == 0 && this.listMetrics.Count == 0 )
                throw new Exception("dimensions and/or metrics are required");

            if (this.listWith.Count > 0)
            {
                oSB.Append("WITH ");
                foreach (string sWith in this.listWith)
                    oSB.Append(string.Format("{0} ", sWith));
            }

            oSB.Append("SELECT ");

            if (this.listDimensions.Count >0 && this.listMetrics.Count > 0 )
            {
                oSB.AppendFormat("{0} (", bNonEmptyRows ? "NON EMPTY" : "");
                for (int i = 0; i < this.listDimensions.Count; i++)
                    oSB.Append(string.Format("{0}{1}", (i > 0 ? "," : ""), this.listDimensions[i]));
                oSB.Append(") ON ROWS, ");
            }

            oSB.AppendFormat("{0} {{", bNonEmptyColumns ? "NON EMPTY" : "");
            if ( this.listMetrics.Count > 0 )
                for (int i = 0; i < this.listMetrics.Count; i++)
                    oSB.Append(string.Format("{0}{1}", (i > 0 ? "," : ""), this.listMetrics[i]));
            else
                for (int i = 0; i < this.listDimensions.Count; i++)
                    oSB.Append(string.Format("{0}{1}", (i > 0 ? "," : ""), this.listDimensions[i]));
            oSB.Append("} ON COLUMNS ");

            oSB.Append("FROM ");
            oSB.Append(convertFiltersToCubeName_r(this.listFilters));
            return oSB.ToString();
        }

        public string getCubeName()
        {
            return convertFiltersToCubeName_r(this.listFilters);
        }

        private string convertFiltersToCubeName_r(LinkedList<MDXFilter> listFilters)
        {
            StringBuilder oSB = new StringBuilder();

            if (listFilters.Count == 0)
            {
                if (string.IsNullOrEmpty(this.sInstanceName))
                    throw new ArgumentException("cube name is required");
                oSB.Append(string.Format("(SELECT FROM {0})", this.sInstanceName));
            }
            else
            {
                MDXFilter oFilter = listFilters.Last.Value;
                if (oFilter.listMemberSets.Count == 0)
                    oSB.Append(convertFiltersToCubeName_r(new LinkedList<MDXFilter>(listFilters.Take(listFilters.Count - 1))));
                else
                {
                    oSB.Append("(");
                    if (oFilter.listMemberSets.Count == 1)
                        oSB.Append(string.Format("SELECT {{{0}}} ON 0 FROM ", Strings.implode(oFilter.listMemberSets[0])));
                    else
                        oSB.Append(string.Format("SELECT {{{0}}} ON 0 FROM ", generateFilters(oFilter)));
                    oSB.Append(convertFiltersToCubeName_r(new LinkedList<MDXFilter>(listFilters.Take(listFilters.Count - 1))));
                    oSB.Append(")");
                }
            }
            return oSB.ToString();
        }

        private string generateFilters(MDXFilter oFilter)
        {
            List<string> listFilters = generateFilter_r(oFilter.listMemberSets);
            foreach (List<string> listExcludedFilter in oFilter.listExcludedFilters)
                listFilters.Remove(Strings.implode(listExcludedFilter));
            return Strings.implode(listFilters, ",", "(", ")");
        }

        private List<string> generateFilter_r(IEnumerable<List<string>> listMemberSets)
        {
            List<string> listFilters, listChildFilters;
            if (listMemberSets.Count() == 1)
                return listMemberSets.First();
            listChildFilters = generateFilter_r(listMemberSets.Skip(1));
            listFilters = new List<string>();
            foreach (string sMember in listMemberSets.First())
                foreach (string sChildFilter in listChildFilters)
                    listFilters.Add(string.Format("{0},{1}", sMember, sChildFilter));
            return listFilters;
        }

        public MDXRunner disableExecuteExceptions(bool bDisableExecuteExceptions)
        {
            this.bDisableExecuteExceptions = bDisableExecuteExceptions;
            return this;
        }

        public IEnumerable<Dictionary<string, T>> executeMDX<T>(AdomdConnection oAdomdConnection)
        {
            return MDXRunner.executeMDX<T>(this.getMDX(), oAdomdConnection);
        }

        public IEnumerable<Dictionary<string, T>> executeMDX<T>(string sConnectionString)
        {
            using (AdomdConnection oAdomdConnection = new AdomdConnection(sConnectionString))
            {
                oAdomdConnection.Open();
                return MDXRunner.executeMDX<T>(this.getMDX(), oAdomdConnection);
            }
        }

        public IEnumerable<Dictionary<string, T>> executeMDX<T>(string sEndpointURL, string sInitialCatalog)
        {
            if (sEndpointURL == null)
                throw new ArgumentException("endpoint URL is required");

            OleDbConnectionStringBuilder oCSBuilder = new OleDbConnectionStringBuilder();
            oCSBuilder.DataSource = sEndpointURL;
            oCSBuilder.Provider = "MSOLAP";
            oCSBuilder.PersistSecurityInfo = true;
            oCSBuilder.Add("SspropInitAppName", "AP's MDXBuilder");
            if (!String.IsNullOrWhiteSpace(sInitialCatalog))
                oCSBuilder.Add("Initial Catalog", sInitialCatalog);

            using (AdomdConnection oAdomdConnection = new AdomdConnection(oCSBuilder.ToString()))
            {
                oAdomdConnection.Open();
                return MDXRunner.executeMDX<T>(this.getMDX(), oAdomdConnection);
            }
        }


        public static IEnumerable<Dictionary<string, T>> executeMDX<T>(string sMDX, AdomdConnection oAdomdConnection)
        {
            AdomdCommand oAdoMDCommand;
            List<Dictionary<string, object>> listGrains;
            CellSet oCellSets;
            TupleCollection tuplesOnColumns;
            TupleCollection tuplesOnRows;

            /*
            * run query & convert into IEnumerable
            */
            oAdoMDCommand = oAdomdConnection.CreateCommand();
            oAdoMDCommand.CommandText = sMDX;

            DateTime oNow = DateTime.Now;
            oCellSets = oAdoMDCommand.ExecuteCellSet();

            Debug.WriteLine(string.Format("MDX query: took {0}", Strings.formatTimeSpan(DateTime.Now - oNow)));

            tuplesOnColumns = oCellSets.Axes[0].Set.Tuples;
            tuplesOnRows = oCellSets.Axes.Count > 1 ? oCellSets.Axes[1].Set.Tuples : null;

            listGrains = new List<Dictionary<string, object>>();

            if (tuplesOnRows != null)
            {
                for (int iRow = 0; iRow < tuplesOnRows.Count; iRow++)
                {
                    Dictionary<string, object> dict = new Dictionary<string, object>();

                    //Loop through all the dimensions
                    foreach (Member oMember in tuplesOnRows[iRow].Members)
                    {
                        string sPropertyName = generateUniqueDictionaryKey(MDXRunner.convertMemberNameToPropertyName(oMember.LevelName), dict);
                        dict[sPropertyName] = oMember.Name.EndsWith(".UNKNOWNMEMBER", StringComparison.InvariantCulture) ? null : oMember.Caption;
                        dict[string.Format("_{0}", sPropertyName)] = oMember.Name;
                    }

                    //Loop through all the measures/metrics
                    for (int iCol = 0; iCol < tuplesOnColumns.Count; iCol++)
                    {
                        string sPropertyName = generateUniqueDictionaryKey(MDXRunner.convertMemberNameToPropertyName(tuplesOnColumns[iCol].Members[0].Name), dict);
                        dict[sPropertyName] = Convert.ToString(oCellSets.Cells[iCol, iRow].Value);
                        dict[string.Format("_{0}", sPropertyName)] = tuplesOnColumns[iCol].Members[0].Name;
                    }

                    listGrains.Add(dict);
                }
            }
            else
            {
                //Loop through all the measures/metrics
                for (int iCol = 0; iCol < tuplesOnColumns.Count; iCol++)
                {
                    Boolean bHasMetrics = false;
                    Dictionary<string, object> dict = new Dictionary<string, object>();

                    // add dimensions & metrics from column, checksumming dimensions
                    StringBuilder oSB = new StringBuilder();
                    foreach (Member oMember in tuplesOnColumns[iCol].Members)
                    {
                        string sPropertyName;
                        object oPropertyValue;

                        sPropertyName = MDXRunner.convertMemberNameToPropertyName(oMember.LevelName);
                        if (sPropertyName == "MeasuresLevel")
                        {
                            sPropertyName = MDXRunner.convertMemberNameToPropertyName(oMember.Name);
                            oPropertyValue = oCellSets.Cells[iCol].Value;
                            bHasMetrics = true;
                        }
                        else
                        {
                            oPropertyValue = oMember.Caption;
                            oSB.Append(string.Format("{0}:{1}:{2};", sPropertyName, oMember.Name, oPropertyValue));
                        }

                        sPropertyName = generateUniqueDictionaryKey(sPropertyName, dict);
                        dict[sPropertyName] = oPropertyValue;
                        dict[string.Format("_{0}", sPropertyName)] = oMember.Name;
                    }
                    dict["__dimension_checksum"] = Strings.calcSHA256Hash(oSB.ToString());

                    // move metrics into previous dict if checksums match
                    if (bHasMetrics && listGrains.Count > 0 && Strings.Equals(listGrains.Last()["__dimension_checksum"], dict["__dimension_checksum"]))
                    {
                        dict = listGrains.Last();
                        foreach (Member oMember in tuplesOnColumns[iCol].Members)
                        {
                            string sPropertyName;
                            object oPropertyValue;
                            sPropertyName = MDXRunner.convertMemberNameToPropertyName(oMember.LevelName);
                            if (sPropertyName == "MeasuresLevel")
                            {
                                sPropertyName = generateUniqueDictionaryKey(MDXRunner.convertMemberNameToPropertyName(oMember.Name), dict);
                                oPropertyValue = oCellSets.Cells[iCol].Value;
                                dict[sPropertyName] = oPropertyValue;
                                dict[string.Format("_{0}", sPropertyName)] = oMember.Name;
                            }
                        }
                        continue;
                    }

                    // add metric when it is implicity added (ie no metrics in query)
                    if (!bHasMetrics && MDXRunner.convertMemberNameToPropertyName(oCellSets.FilterAxis.Positions[0].Members[0].LevelName) == "MeasuresLevel")
                        dict[generateUniqueDictionaryKey(MDXRunner.convertMemberNameToPropertyName(oCellSets.FilterAxis.Positions[0].Members[0].Name), dict)] = Convert.ToString(oCellSets.Cells[iCol].Value);

                    listGrains.Add(dict);
                }
            }

            // remove checksums & convert dictionary to desired time
            Type typeParameterType = typeof(T);
            List<Dictionary<string, T>> listFinishedGrains = new List<Dictionary<string, T>>();
            foreach (Dictionary<string, object> dictGrain in listGrains)
            {
                Dictionary<string, T> dictFinishedGrain = new Dictionary<string, T>();
                dictGrain.Remove("__dimension_checksum");
                foreach (string sKey in dictGrain.Keys)
                    if (sKey[0] != '_' || typeParameterType == typeof(string) || typeParameterType == typeof(object))
                        dictFinishedGrain[sKey] = (T)Convert.ChangeType(dictGrain[sKey], typeParameterType);
                listFinishedGrains.Add(dictFinishedGrain);
            }

            return listFinishedGrains.ToArray();
        }

        private static string generateUniqueDictionaryKey(string sDesiredKey, Dictionary<string, object> dict)
        {
            if (dict.ContainsKey(sDesiredKey))
            {
                string sAlternateKey;
                int i = 1;
                do
                {
                    sAlternateKey = string.Format("{0}{1}", sDesiredKey, i++);
                }
                while (dict.ContainsKey(sAlternateKey));
                sDesiredKey = sAlternateKey;
            }
            return sDesiredKey;
        }

        public static string escapeForMDXMember(string s)
        {
            return string.IsNullOrEmpty(s) ? null : s.Replace("]", "]]");
        }

        public static string convertMemberNameToPropertyName(string sMemberName)
        {
            string sPropertyName, sFinalPropertyName;

            if ( string.IsNullOrEmpty(sMemberName) )
                return null;

            if ((sPropertyName = getNthStringOnSquareBraces(sMemberName, 0)) == "(All)")
                sPropertyName = getNthStringOnSquareBraces(sMemberName, 1);

            if (sPropertyName == null)
                return null;

            sFinalPropertyName = "";
            foreach (char c in sPropertyName)
                if ((c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || c >= '0' && c <= '9')
                    sFinalPropertyName += c;

            return sFinalPropertyName;
        }

        /**
         * Extract components from a member, eg:
         * getNthStringOnSquareBraces("[Advertising Dim].[Advertiser Name].[Advertiser Name].&[XXX].&[YYY]", 0) = "YYY"
         * getNthStringOnSquareBraces("[Advertising Dim].[Advertiser Name].[Advertiser Name].&[XXX].&[YYY]", 1) = "XXX"
         * getNthStringOnSquareBraces("[Advertising Dim].[Advertiser Name].[Advertiser Name].&[XXX].&[YYY]", 2) = "Advertiser Name"
         * getNthStringOnSquareBraces("[Advertising Dim].[Advertiser Name].[Advertiser Name]", 0) = "Advertiser Name"
         * getNthStringOnSquareBraces("[Advertising Dim].[Advertiser Name].[Advertiser Name]", 2) = "Advertising Dim"
         * getNthStringOnSquareBraces("[Advertising Dim].[Advertiser Name].[Advertiser Name]", 3) = null
         * @return null if sMember is not something that looks like a member or hierarchy or if n is out of bounds
         */
        private static string getNthStringOnSquareBraces(string sMember, int n)
        {
            if (string.IsNullOrEmpty(sMember))
                return null;
            string[] aString = sMember.Split('.');
            if (aString.Length <= n)
                return null;
            string memberName = aString[aString.Length - 1 - n];

            return memberName.Substring(1, memberName.Length - 2);
        }

        public static string getDimensionName(string sMember)
        {
            if (string.IsNullOrEmpty(sMember))
                return null;
            Triple<string[], string[], string> oTripleMemberElements = splitMemberIntoHierarchyAndMemberValues(sMember);
            return String.Format("[{0}]", oTripleMemberElements.getFirst().ElementAt(0));
        }

        /**
         * Extract attribute name from member/hierarchy/whatever, eg:
         * getHierarchyName("[Advertising Dim].[Advertiser Name].[Advertiser Name].&[XXX].&[YYY]") = "[Advertising Dim].[Advertiser Name]"
         * getHierarchyName("[Advertising Dim].[Advertiser Name].[Advertiser Name]") = "[Advertising Dim].[Advertiser Name]"
         * getHierarchyName("[Measures].[Clicks]") = "[Measures].[Clicks]"
         * @return null if sMember is not something that looks like a member or hierarchy
         */
        public static string getAttributeName(string sMember)
        {
            if (string.IsNullOrEmpty(sMember))
                return null;
            Triple<string[], string[], string> oTripleMemberElements = splitMemberIntoHierarchyAndMemberValues(sMember);
            return oTripleMemberElements == null ? null : String.Format("[{0}].[{1}]", oTripleMemberElements.getFirst().ElementAt(0), oTripleMemberElements.getFirst().ElementAt(1));
        }

        /**
         * Extract hierarchy compontent from member, eg:
         * getHierarchyName("[Advertising Dim].[Advertiser Name].[Advertiser Name].&[XXX].&[YYY]") = "[Advertising Dim].[Advertiser Name].[Advertiser Name]"
         * getHierarchyName("[Advertising Dim].[Advertiser Name].[Advertiser Name]") = "[Advertising Dim].[Advertiser Name].[Advertiser Name]"
         * getHierarchyName("[Measures].[Clicks]") = "[Measures].[Clicks]"
         * @return null if sMember is not something that looks like a member or hierarchy
         */
        public static string getHierarchyName(string sMember)
        {
            if (string.IsNullOrEmpty(sMember))
                return null;
            Triple<string[], string[], string> oTripleMemberElements = splitMemberIntoHierarchyAndMemberValues(sMember);
            return oTripleMemberElements == null ? null : Strings.implode(oTripleMemberElements.getFirst(), ".", "[", "]");
        }

        /**
         * get member components
         * getMemberValues("[Advertising Dim].[Advertiser Name].[Advertiser Name].&[XXX]") = {"XXX" }
         * getMemberValues("[Advertising Dim].[Advertiser Name].[Advertiser Name].&[XXX]&[YYY]") = {"XXX", "YYY"}
         * getMemberValues("[Advertising Dim].[Advertiser Name].[Advertiser Name]") = null
         * @return null if sMember is not a member (ie, does not contain &)
         */
        public static string[] getMemberValues(string sMember)
        {
            Triple<string[], string[], string> oTripleMemberElements = splitMemberIntoHierarchyAndMemberValues(sMember);
            return oTripleMemberElements == null ? null : oTripleMemberElements.getSecond();
        }

        /**
         * get member sufix
         * getMemberSuffix("[Advertising Dim].[Advertiser Name].[Advertiser Name].&[XXX].&[YYY].Hello") = "Hello"
         * getMemberSuffix("[Advertising Dim].[Advertiser Name].[Advertiser Name].&[XXX]&[YYY].Hello") = "Hello"
         * getMemberSuffix("[Advertising Dim].[Advertiser Name].[Advertiser Name]") = null
         * getMemberSuffix("[Advertising Dim].[Advertiser Name].[Advertiser Name].Hello") = "Hello"
         * @return null if sMember is not a member or doesn't contain a suffix
         */
        public static string getMemberSuffix(string sMember)
        {
            Triple<string[], string[], string> oTripleMemberElements = splitMemberIntoHierarchyAndMemberValues(sMember);
            return oTripleMemberElements == null ? null : oTripleMemberElements.getThird();
        }

        /**
         * @return pair of hierarchy elements & member values
         */
        public static Triple<string[], string[], string> splitMemberIntoHierarchyAndMemberValues(string sMember)
        {
            List<string> listHierarchyElements, listMemberValues;
            string sSuffix;
            Match oMatch;
            
            oMatch = oRegexCompiled.Match(sMember);
            if (!oMatch.Success || oMatch.Groups.Count < 6 )
                return null;
            
            listHierarchyElements = new List<string>();
            listMemberValues = new List<string>();
            
            foreach (Capture c in oMatch.Groups[2].Captures)
                listHierarchyElements.Add(c.Value);
            foreach (Capture c in oMatch.Groups[5].Captures)
                listMemberValues.Add(c.Value);

            //when listMemberValues is empty, then we expect that this is UNKNOWNMEMBER 
            if (listMemberValues.Count == 0 &&
                oMatch.Groups[8].Captures[0].Value == "UNKNOWNMEMBER")            
                sSuffix = null;            
            else
            {
                sSuffix = oMatch.Groups[8].Captures.Count == 0 ? null : oMatch.Groups[8].Captures[0].Value; 
                sSuffix = string.IsNullOrEmpty(sSuffix) ? null : sSuffix;
            }
            return new Triple<string[], string[], string>(listHierarchyElements.ToArray(), listMemberValues.ToArray(), sSuffix);
        }

        /*
         * 0: preamble
         * 1: FROM clause
         * 2: WHERE clause
         */
        public static string[] splitMDXQuery(string sMDX)
        {
            string sPreamble = null, sFROM = null, sWHERE = null;
            int iFROMPos, iWHEREPos;

            if (sMDX != null)
            {
                iFROMPos = sMDX.ToUpper().IndexOf("FROM");
                iWHEREPos = sMDX.ToUpper().IndexOf("WHERE");

                if (iFROMPos < 0)
                    sPreamble = sMDX;
                else if (iWHEREPos < 0)
                {
                    sPreamble = sMDX.Substring(0, iFROMPos);
                    sFROM = sMDX.Substring(iFROMPos);
                }
                else
                {
                    sPreamble = sMDX.Substring(0, iFROMPos);
                    sFROM = sMDX.Substring(iFROMPos, iWHEREPos - iFROMPos);
                    sWHERE = sMDX.Substring(iWHEREPos);
                }
            }

            return new string[] { sPreamble, sFROM, sWHERE };
        }
    }

    public class Strings
    {
        public static string calcSHA256Hash(string sData)
        {
            HashAlgorithm oHashAlg = new SHA256CryptoServiceProvider();
            byte[] ayData = System.Text.Encoding.UTF8.GetBytes(sData);
            byte[] ayHash = oHashAlg.ComputeHash(ayData);
            string base64 = Convert.ToBase64String(ayHash);
            return base64;
        }

        public static string formatTimeSpan(TimeSpan span)
        {
            StringBuilder oSB = new StringBuilder();
            int i = 0, iMax = 2;

            int iDays = span.Days;

            if (iDays > 365 && i < iMax)
            {
                int iYears = (int)Math.Floor((double)iDays / 365.0);
                oSB.AppendFormat("{0:0} year{1} ", iYears, iYears > 1 ? "s" : "", i++);
                iDays -= iYears * 365;
            }

            if (iDays > 7 && i < iMax)
            {
                int iWeeks = (int)Math.Floor((double)iDays / 7.0);
                oSB.AppendFormat("{0:0} week{1} ", iWeeks, iWeeks > 1 ? "s" : "", i++);
                iDays -= iWeeks * 7;
            }

            if (iDays > 0 && i < iMax)
                oSB.AppendFormat("{0:0} day{1} ", iDays, iDays > 1 ? "s" : "", i++);
            if (span.Hours > 0 && i < iMax)
                oSB.AppendFormat("{0:0} hrs ", span.Hours, i++);
            if (span.Minutes > 0 && i < iMax)
                oSB.AppendFormat("{0:0} mins ", span.Minutes, i++);
            if (span.Seconds > 0 && i < iMax)
                oSB.AppendFormat("{0:0} secs ", span.Seconds, i++);
            if (span.Milliseconds > 0 && i < iMax)
                oSB.AppendFormat("{0:0} ms", span.Milliseconds, i++);

            return oSB.ToString().Trim(new char[] { ',', ' ' });
        }

        public static string implode(IEnumerable enumerable)
        {
            return implode(enumerable, ",", "", "");
        }

        public static string implode(IDictionary enumerable)
        {
            return implode(enumerable, ";", "", "");
        }

        public static string implode(IEnumerable enumerable, string separator)
        {
            return implode(enumerable, separator, "", "");
        }

        public static string implode(IDictionary enumerable, string separator)
        {
            return implode(enumerable, separator, "", "");
        }

        public static string implode(IEnumerable enumerable, string separator, string prefix, string suffix)
        {
            return enumerable == null ? "" : join(separator, Strings.ToString(enumerable).ToArray(), prefix, suffix);
        }

        public static string implode(IDictionary enumerable, string separator, string prefix, string suffix)
        {
            return enumerable == null ? "" : join(separator, Strings.ToString(enumerable).ToArray(), prefix, suffix);
        }

        private static string join(string separator, string[] value, string prefix, string suffix)
        {
            StringBuilder oSB = new StringBuilder();
            foreach (string s in value)
                oSB.Append(string.Format(
                    "{0}{1}{2}{3}",
                    (oSB.Length == 0 ? "" : separator),
                    (prefix == null ? "" : prefix), s, (suffix == null ? "" : suffix)
                    ));
            return oSB.ToString();
        }

        public static List<string> ToString(IEnumerable enumerable)
        {
            List<string> list = new List<string>();

            if (enumerable != null)
                foreach (Object o in enumerable)
                    if (o == null)
                        list.Add(null);
                    else if (o is IDictionary && !(o is string))
                        list.Add(String.Format("{{{0}}}", Strings.implode((IDictionary)o)));
                    else if (o is IEnumerable && !(o is string))
                        list.Add(String.Format("{{{0}}}", Strings.implode((IEnumerable)o)));
                    else
                        list.Add(Convert.ToString(o));
            return list;
        }
    }

    public class Triple<X, Y, Z>
    {
        private X oX;
        private Y oY;
        private Z oZ;

        public Triple(X oX, Y oY, Z oZ)
        {
            this.oX = oX;
            this.oY = oY;
            this.oZ = oZ;
        }

        public X getFirst()
        {
            return this.oX;
        }

        public Y getSecond()
        {
            return this.oY;
        }

        public Z getThird()
        {
            return this.oZ;
        }

        public override string ToString()
        {
            return string.Format("[Triple:{0}:{1}:{2}]", oX, oY, oZ);
        }
    }
}
