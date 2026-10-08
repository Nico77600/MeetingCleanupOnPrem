// Meeting Cleanup On-Prem - compiled helpers, loaded once by MeetingCleanupOnPrem.psm1 (Add-Type).
//
// A PowerShell function call costs tens of microseconds; the steps below run once per calendar item, copy or
// report cell - hundreds of thousands of times on a large organization. They are compiled here (as in Meeting
// Cleanup 1.2.1) so that a search of thousands of copies spends its time in Exchange, not in PowerShell. They
// behave exactly like the PowerShell code they replace:
//
//   Fast   property reads (no exception when a property is missing), the calendar items of the EWS answers,
//          the recurrence in words, dates, the totals, the report rows, the CSV cells (formula injection
//          neutralised) and the JSON of the report
//
// Author : Nicolas Fabert
// Version: 1.1.0

using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Management.Automation;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml;

namespace MeetingCleanupOnPremNative
{
    public static class Fast
    {
        public const string Version = "1.1.0";
        static readonly CultureInfo Inv = CultureInfo.InvariantCulture;
        static readonly string[] CopyRoles = { "Organizer", "Attendee", "Room" };
        const string TypesNs = "http://schemas.microsoft.com/exchange/services/2006/types";
        static readonly Regex RecurrenceNs = new Regex("^<t:Recurrence[^>]*xmlns:t=", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        static readonly Regex RecurrenceTag = new Regex("^<t:Recurrence", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        static readonly Regex RecurrenceSuffix = new Regex("Recurrence$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        static readonly HashSet<string> RangeNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "NoEndRecurrence", "EndDateRecurrence", "NumberedRecurrence" };

        // ---- properties ------------------------------------------------------------------------------

        static object Base(object o)
        {
            var p = o as PSObject;
            if (p != null && !(p.BaseObject is PSCustomObject)) { return p.BaseObject; }
            return o;
        }

        /// <summary>A property of an object or a key of a dictionary, or null (Get-McoProperty).</summary>
        public static object Prop(object o, string name)
        {
            if (o == null) { return null; }
            var b = Base(o);
            var d = b as IDictionary;
            if (d != null) { return d.Contains(name) ? d[name] : null; }
            var p = PSObject.AsPSObject(o).Properties[name];
            return p == null ? null : p.Value;
        }

        /// <summary>A property along a path, or null.</summary>
        public static object Path(object o, params string[] names)
        {
            foreach (var n in names) { o = Prop(o, n); if (o == null) { return null; } }
            return o;
        }

        /// <summary>A property as text ([string] of PowerShell); "" when missing.</summary>
        public static string Text(object o, params string[] names)
        {
            return ToText(Path(o, names));
        }

        /// <summary>A property as a boolean (PowerShell truth: [bool]).</summary>
        public static bool Flag(object o, string name)
        {
            return LanguagePrimitives.IsTrue(Prop(o, name));
        }

        public static string ToText(object v)
        {
            if (v == null) { return ""; }
            v = Base(v);
            var s = v as string;
            if (s != null) { return s; }
            return (string)LanguagePrimitives.ConvertTo(v, typeof(string), Inv) ?? "";
        }

        static IEnumerable Items(object o)
        {
            if (o == null) { yield break; }
            var b = Base(o);
            if (b is string || b is IDictionary) { yield return o; yield break; }
            var e = b as IEnumerable;
            if (e == null) { yield return o; yield break; }
            foreach (var x in e) { if (x != null) { yield return x; } }
        }

        static int ToInt(object v)
        {
            if (v == null) { return 0; }
            try { return (int)LanguagePrimitives.ConvertTo(v, typeof(int), Inv); } catch (Exception) { return 0; }
        }

        static PSObject Row(params object[] pairs)
        {
            var o = new PSObject();
            for (int i = 0; i < pairs.Length; i += 2) { o.Properties.Add(new PSNoteProperty((string)pairs[i], pairs[i + 1])); }
            return o;
        }

        static bool Is(string a, string b)
        {
            return string.Equals(a, b, StringComparison.OrdinalIgnoreCase);
        }

        // ---- EWS calendar items ----------------------------------------------------------------------

        static string NodeText(XmlNode node, string xpath, XmlNamespaceManager ns)
        {
            if (node == null) { return null; }
            var hit = node.SelectSingleNode(xpath, ns);
            return hit == null ? null : hit.InnerText;
        }

        /// <summary>An EWS date (UTC) as a UTC DateTime, or null (ConvertFrom-McoEwsDate).</summary>
        public static object EwsDate(string text)
        {
            if (string.IsNullOrEmpty(text)) { return null; }
            return DateTime.Parse(text, Inv, DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal);
        }

        static string Address(XmlNode mailbox, XmlNamespaceManager ns)
        {
            if (mailbox == null) { return ""; }
            var a = mailbox.SelectSingleNode("t:EmailAddress", ns);
            return a == null ? "" : (a.InnerText ?? "").ToLowerInvariant();
        }

        static object[] Addresses(XmlNode node, string xpath, XmlNamespaceManager ns)
        {
            var list = new List<object>();
            foreach (XmlNode m in node.SelectNodes(xpath, ns))
            {
                var a = Address(m, ns);
                if (a.Length > 0) { list.Add(a); }
            }
            return list.ToArray();
        }

        /// <summary>
        /// A calendar item of an EWS answer (CalendarView, GetItem, Recoverable Items) as the object of the tool
        /// (ConvertFrom-McoCalendarNode): the same properties, in the same order.
        /// </summary>
        public static PSObject CalendarNode(XmlNode node, XmlNamespaceManager ns, string mailbox)
        {
            var id = node.SelectSingleNode("t:ItemId", ns) as XmlElement;
            var organizer = node.SelectSingleNode("t:Organizer/t:Mailbox", ns);
            var cancelledText = NodeText(node, "t:IsCancelled", ns) ?? "";
            bool cancelled = false;
            if (cancelledText.Length > 0) { bool.TryParse(cancelledText, out cancelled); }
            var recurrence = node.SelectSingleNode("t:Recurrence", ns);
            var zone = node.SelectSingleNode("t:StartTimeZone", ns) as XmlElement;
            var start = EwsDate(NodeText(node, "t:Start", ns));
            var end = EwsDate(NodeText(node, "t:End", ns));
            string recurrenceXml = "";
            if (recurrence != null)
            {
                recurrenceXml = recurrence.OuterXml;
                if (!RecurrenceNs.IsMatch(recurrenceXml)) { recurrenceXml = RecurrenceTag.Replace(recurrenceXml, "<t:Recurrence xmlns:t=\"" + TypesNs + "\"", 1); }
            }
            return Row(
                "EventId", id != null ? id.GetAttribute("Id") : "",
                "ChangeKey", id != null ? id.GetAttribute("ChangeKey") : "",
                "Mailbox", (mailbox ?? "").ToLowerInvariant(),
                "Subject", NodeText(node, "t:Subject", ns) ?? "",
                "ItemClass", NodeText(node, "t:ItemClass", ns),
                "MeetingId", (NodeText(node, "t:UID", ns) ?? "").ToUpperInvariant(),
                "Organizer", Address(organizer, ns),
                "OrganizerName", NodeText(organizer, "t:Name", ns) ?? "",
                "Start", start ?? DateTime.MinValue,
                "End", end ?? DateTime.MinValue,
                "AppointmentType", NodeText(node, "t:CalendarItemType", ns) ?? "",
                "Location", NodeText(node, "t:Location", ns) ?? "",
                "Response", NodeText(node, "t:MyResponseType", ns) ?? "",
                "IsCancelled", cancelled,
                "RequiredAttendees", Addresses(node, "t:RequiredAttendees/t:Attendee/t:Mailbox", ns),
                "OptionalAttendees", Addresses(node, "t:OptionalAttendees/t:Attendee/t:Mailbox", ns),
                "Resources", Addresses(node, "t:Resources/t:Attendee/t:Mailbox", ns),
                "Body", NodeText(node, "t:Body", ns),
                "DateTimeCreated", NodeText(node, "t:DateTimeCreated", ns),
                "LastModifiedUtc", EwsDate(NodeText(node, "t:LastModifiedTime", ns)),
                "RecurrenceXml", recurrenceXml,
                "TimeZoneId", zone != null ? zone.GetAttribute("Id") : "",
                "SeriesId", "",
                "Occurrences", new object[0]);
        }

        /// <summary>The calendar items of an answer that have an item ID (CalendarView, FindItem).</summary>
        public static List<PSObject> CalendarNodes(XmlNode root, string xpath, XmlNamespaceManager ns, string mailbox)
        {
            var list = new List<PSObject>();
            if (root == null) { return list; }
            foreach (XmlNode n in root.SelectNodes(xpath, ns))
            {
                var item = CalendarNode(n, ns, mailbox);
                if (Text(item, "EventId").Length > 0) { list.Add(item); }
            }
            return list;
        }

        static string ChildText(XmlNode parent, string localName, int max)
        {
            foreach (XmlNode c in parent.ChildNodes)
            {
                if (c.NodeType == XmlNodeType.Element && c.LocalName == localName)
                {
                    var t = c.InnerText ?? "";
                    return max > 0 && t.Length > max ? t.Substring(0, max) : t;
                }
            }
            return "";
        }

        /// <summary>An EWS Recurrence element in words: Weekly (Monday), 4 occurrences from 2026-10-12 (Format-McoEwsRecurrence).</summary>
        public static string Recurrence(string xml)
        {
            if (string.IsNullOrEmpty(xml)) { return ""; }
            var doc = new XmlDocument();
            try { doc.LoadXml(xml); } catch (Exception) { return ""; }
            XmlNode pattern = null, range = null;
            foreach (XmlNode c in doc.DocumentElement.ChildNodes)
            {
                if (c.NodeType != XmlNodeType.Element) { continue; }
                if (RangeNames.Contains(c.LocalName)) { if (range == null) { range = c; } }
                else if (pattern == null) { pattern = c; }
            }
            var text = "";
            if (pattern != null)
            {
                var interval = ChildText(pattern, "Interval", 0);
                var every = interval.Length > 0 && interval != "1" ? "every " + interval + " " : "";
                switch (pattern.LocalName)
                {
                    case "DailyRecurrence": text = every.Length > 0 ? every + "days" : "Daily"; break;
                    case "WeeklyRecurrence": text = (every.Length > 0 ? every + "weeks" : "Weekly") + " (" + ChildText(pattern, "DaysOfWeek", 0) + ")"; break;
                    case "AbsoluteMonthlyRecurrence": text = (every.Length > 0 ? every + "months" : "Monthly") + " (day " + ChildText(pattern, "DayOfMonth", 0) + ")"; break;
                    case "RelativeMonthlyRecurrence": text = (every.Length > 0 ? every + "months" : "Monthly") + " (" + ChildText(pattern, "DayOfWeekIndex", 0) + " " + ChildText(pattern, "DaysOfWeek", 0) + ")"; break;
                    default: text = RecurrenceSuffix.Replace(pattern.LocalName, ""); break;
                }
            }
            if (range != null)
            {
                var from = ChildText(range, "StartDate", 10);
                switch (range.LocalName)
                {
                    case "NumberedRecurrence": text += ", " + ChildText(range, "NumberOfOccurrences", 0) + " occurrences from " + from; break;
                    case "EndDateRecurrence": text += ", from " + from + " until " + ChildText(range, "EndDate", 10); break;
                    default: text += ", from " + from + ", no end"; break;
                }
            }
            return text;
        }

        // ---- dates -----------------------------------------------------------------------------------

        /// <summary>A UTC date shown in a time zone: yyyy-MM-dd HH:mm or yyyy-MM-dd (Format-McoDate).</summary>
        public static string FormatDate(object utc, TimeZoneInfo zone, bool dateOnly, bool periodEnd)
        {
            if (utc == null) { return ""; }
            utc = Base(utc);
            DateTime d;
            if (utc is DateTime) { d = (DateTime)utc; }
            else
            {
                var text = ToText(utc);
                if (text.Length == 0) { return ""; }
                d = DateTime.Parse(text, Inv, DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal);
            }
            d = DateTime.SpecifyKind(d.ToUniversalTime(), DateTimeKind.Utc);
            var local = TimeZoneInfo.ConvertTimeFromUtc(d, zone ?? TimeZoneInfo.Local);
            if (periodEnd && local.TimeOfDay == TimeSpan.Zero) { local = local.AddDays(-1); dateOnly = true; }
            return local.ToString(dateOnly ? "yyyy-MM-dd" : "yyyy-MM-dd HH:mm", Inv);
        }

        // ---- totals ----------------------------------------------------------------------------------

        /// <summary>The totals of a result (Update-McoResultCounts), in one pass over the copies.</summary>
        public static PSObject Counts(object meetings)
        {
            int count = 0, series = 0, selected = 0, transferred = 0, copies = 0, roomCopies = 0, organizerCopies = 0, attendeeCopies = 0, occurrenceCopies = 0;
            int notProcessed = 0, removed = 0, cancelled = 0, alreadyGone = 0, kept = 0, restored = 0, alreadyPresent = 0, notFound = 0, notRestorable = 0, failed = 0;
            var organizers = new HashSet<string>(StringComparer.Ordinal);
            var mailboxes = new HashSet<string>(StringComparer.Ordinal);
            foreach (var m in Items(meetings))
            {
                count++;
                if (Is(Text(m, "Kind"), "Series")) { series++; }
                if (Flag(m, "Selected")) { selected++; }
                if (Is(Text(m, "Status"), "Transferred")) { transferred++; }
                var organizer = Text(m, "Organizer");
                if (organizer.Length > 0) { organizers.Add(organizer); }
                foreach (var c in Items(Prop(m, "Copies")))
                {
                    var result = Text(c, "Result");
                    if (result.Length > 0)
                    {
                        if (Is(result, "Not processed")) { notProcessed++; }
                        else if (Is(result, "Removed")) { removed++; }
                        else if (Is(result, "Cancelled")) { cancelled++; }
                        else if (Is(result, "Already gone")) { alreadyGone++; }
                        else if (Is(result, "Kept")) { kept++; }
                        else if (Is(result, "Restored")) { restored++; }
                        else if (Is(result, "Already present")) { alreadyPresent++; }
                        else if (Is(result, "Not restorable")) { notRestorable++; }
                        else if (Is(result, "Failed") || Is(result, "Not done")) { failed++; }
                        else if (Is(result, "Not found") && Is(Text(c, "Action"), "Restore")) { notFound++; }
                    }
                    var role = Text(c, "Role");
                    if (Text(c, "EventId").Length == 0) { continue; }
                    if (Is(role, "Room")) { roomCopies++; }
                    else if (Is(role, "Organizer")) { organizerCopies++; }
                    else if (Is(role, "Attendee")) { attendeeCopies++; }
                    else { continue; }
                    copies++;
                    mailboxes.Add(Text(c, "Mailbox"));
                    if (Text(c, "Occurrence").Length > 0) { occurrenceCopies++; }
                }
            }
            return Row(
                "Meetings", count, "Series", series, "Selected", selected, "Organizers", organizers.Count,
                "Copies", copies, "Mailboxes", mailboxes.Count, "RoomCopies", roomCopies, "OrganizerCopies", organizerCopies,
                "AttendeeCopies", attendeeCopies, "OccurrenceCopies", occurrenceCopies, "Rooms", roomCopies, "Attendees", attendeeCopies,
                "NotProcessed", notProcessed, "Removed", removed, "Cancelled", cancelled, "AlreadyGone", alreadyGone, "Kept", kept,
                "Restored", restored, "AlreadyPresent", alreadyPresent, "NotFound", notFound, "NotRestorable", notRestorable,
                "Transferred", transferred, "Failed", failed);
        }

        /// <summary>The copies of a meeting found in a calendar, one per mailbox (an occurrence copy counts once for its mailbox).</summary>
        public static List<object> RealCopies(object meeting)
        {
            var list = new List<object>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var c in Items(Prop(meeting, "Copies")))
            {
                if (Text(c, "EventId").Length == 0 || !IsCopyRole(Text(c, "Role"))) { continue; }
                if (seen.Add(Text(c, "Mailbox"))) { list.Add(c); }
            }
            return list;
        }

        static bool IsCopyRole(string role)
        {
            foreach (var r in CopyRoles) { if (Is(r, role)) { return true; } }
            return false;
        }

        static int CountRole(List<object> copies, string role)
        {
            int n = 0;
            foreach (var c in copies) { if (Is(Text(c, "Role"), role)) { n++; } }
            return n;
        }

        // ---- report ----------------------------------------------------------------------------------

        /// <summary>One row per meeting, for the CSV and the HTML (Get-McoMeetingRows).</summary>
        public static Table MeetingTable(object meetings)
        {
            var t = new Table("MeetingId", "Subject", "Organizer", "OrganizerName", "Kind", "Scope", "Occurrences", "OccurrencesSkipped", "NewOrganizer", "NewMeetingId", "TransferMethod",
                "StartText", "EndText", "NextInPeriod", "Recurrence", "Location", "OrganizerCopy", "Copies", "RoomCopies", "AttendeeCopies", "NotProcessed",
                "Cancelled", "Selected", "Status", "Notes");
            foreach (var m in Items(meetings))
            {
                var copies = RealCopies(m);
                int notProcessed = 0;
                foreach (var c in Items(Prop(m, "Copies"))) { if (Is(Text(c, "Result"), "Not processed")) { notProcessed++; } }
                var notes = new List<string>();
                foreach (var n in Items(Prop(m, "Notes"))) { notes.Add(ToText(n)); }
                t.Rows.Add(new object[] {
                    Text(m, "MeetingId"), Text(m, "Subject"), Text(m, "Organizer"), Text(m, "OrganizerName"), Text(m, "Kind"), Text(m, "Scope"), ToInt(Prop(m, "Occurrences")), SkippedCount(m),
                    Text(m, "NewOrganizer"), Text(m, "NewMeetingId"), Text(m, "TransferMethod"),
                    Text(m, "StartText"), Text(m, "EndText"), Text(m, "NextInPeriod"), Text(m, "Recurrence"), Text(m, "Location"), Text(m, "OrganizerCopy"),
                    copies.Count, CountRole(copies, "Room"), CountRole(copies, "Attendee"), notProcessed,
                    Flag(m, "Cancelled"), Flag(m, "Selected"), Text(m, "Status"), notes.ToArray() });
            }
            return t;
        }

        /// <summary>How many occurrences of a series are left out (SkippedOccurrences of a reviewed report).</summary>
        public static int SkippedCount(object meeting)
        {
            var keys = new HashSet<string>(StringComparer.Ordinal);
            foreach (var k in Items(Prop(meeting, "SkippedOccurrences"))) { var s = KeyText(k); if (s.Length > 0) { keys.Add(s); } }
            return keys.Count;
        }

        /// <summary>
        /// An occurrence key as text: its start, UTC, round-trip format ("o"). A report read again with ConvertFrom-Json
        /// gives a date: it is written back the same way.
        /// </summary>
        public static string KeyText(object v)
        {
            var b = v == null ? null : Base(v);
            if (b is DateTime) { return ((DateTime)b).ToUniversalTime().ToString("o", Inv); }
            if (b is DateTimeOffset) { return ((DateTimeOffset)b).UtcDateTime.ToString("o", Inv); }
            return ToText(v);
        }

        /// <summary>The slot of an occurrence copy (OccurrenceKey; Occurrence for a copy of an older report).</summary>
        public static string OccurrenceKeyOf(object copy)
        {
            var k = KeyText(Prop(copy, "OccurrenceKey"));
            return k.Length > 0 ? k : KeyText(Prop(copy, "Occurrence"));
        }

        /// <summary>The Kind column of the console: Single, Series, or the occurrences of a series acted on (2 occ., 1/3 occ.).</summary>
        public static string KindText(object m)
        {
            if (Text(m, "Scope") != "Occurrences") { return Text(m, "Kind"); }
            int all = ToInt(Prop(m, "Occurrences"));
            int skipped = SkippedCount(m);
            return skipped > 0 ? string.Format(Inv, "{0}/{1} occ.", Math.Max(0, all - skipped), all) : string.Format(Inv, "{0} occ.", all);
        }

        /// <summary>One row per copy (Get-McoCopyRows).</summary>
        public static Table CopyTable(object meetings)
        {
            var t = new Table("MeetingId", "MeetingSubject", "Organizer", "Mailbox", "Role", "Via", "Occurrence", "Response", "ShowAs", "Cancelled", "Action", "Result",
                "HttpStatus", "Verified", "ActionUtc", "Detail", "EventId");
            foreach (var m in Items(meetings))
            {
                var id = Text(m, "MeetingId"); var subject = Text(m, "Subject"); var organizer = Text(m, "Organizer");
                foreach (var c in Items(Prop(m, "Copies")))
                {
                    var status = Prop(c, "HttpStatus");
                    t.Rows.Add(new object[] {
                        id, subject, organizer, Text(c, "Mailbox"), Text(c, "Role"), Text(c, "Via"), Text(c, "Occurrence"), Text(c, "Response"), Text(c, "ShowAs"),
                        Flag(c, "Cancelled"), Text(c, "Action"), Text(c, "Result"), LanguagePrimitives.IsTrue(status) ? (object)ToInt(status) : "",
                        Text(c, "Verified"), Text(c, "ActionUtc"), Text(c, "Detail"), Text(c, "EventId") });
                }
            }
            return t;
        }

        /// <summary>One row per organizer of the run, with what was found and done for its meetings (Get-McoOrganizerRows).</summary>
        public static Table OrganizerTable(object organizers, object meetings)
        {
            var t = new Table("Input", "DisplayName", "PrimaryAddress", "State", "Detail", "Meetings", "Series", "Copies", "Removed", "Cancelled", "Restored", "Transferred", "Failed");
            var byKey = new Dictionary<string, List<object>>(StringComparer.OrdinalIgnoreCase);
            foreach (var m in Items(meetings))
            {
                var k = Text(m, "OrganizerKey");
                if (k.Length == 0) { k = Text(m, "Organizer"); }
                List<object> list;
                if (!byKey.TryGetValue(k, out list)) { list = new List<object>(); byKey[k] = list; }
                list.Add(m);
            }
            foreach (var o in Items(organizers))
            {
                var mine = new List<object>();
                var counted = new HashSet<object>(ReferenceEqualityComparer.Instance);
                var keys = new List<string>();
                foreach (var a in Items(Prop(o, "Addresses"))) { keys.Add(ToText(a)); }
                keys.Add(Text(o, "PrimaryAddress"));
                foreach (var k in keys)
                {
                    List<object> list;
                    if (k.Length > 0 && byKey.TryGetValue(k, out list)) { foreach (var m in list) { if (counted.Add(m)) { mine.Add(m); } } }
                }
                int series = 0, copies = 0, removed = 0, cancelled = 0, restored = 0, transferred = 0, failed = 0;
                foreach (var m in mine)
                {
                    if (Is(Text(m, "Kind"), "Series")) { series++; }
                    if (Is(Text(m, "Status"), "Transferred")) { transferred++; }
                    foreach (var c in Items(Prop(m, "Copies")))
                    {
                        if (Text(c, "EventId").Length > 0) { copies++; }
                        var r = Text(c, "Result");
                        if (Is(r, "Removed")) { removed++; }
                        else if (Is(r, "Cancelled")) { cancelled++; }
                        else if (Is(r, "Restored")) { restored++; }
                        else if (Is(r, "Failed")) { failed++; }
                    }
                }
                t.Rows.Add(new object[] { Text(o, "Input"), Text(o, "DisplayName"), Text(o, "PrimaryAddress"), Text(o, "State"), Text(o, "Detail"), mine.Count, series, copies, removed, cancelled, restored, transferred, failed });
            }
            return t;
        }

        /// <summary>
        /// One row per meeting of a transfer (Get-McoTransferRows): the old organizer and its state, the new one, the
        /// method (Exchange Server: always re-created), the new meeting and its invitation, what became of the old
        /// organizer's copy and of the old copies.
        /// </summary>
        public static Table TransferTable(object organizers, object meetings)
        {
            var t = new Table("MeetingId", "Subject", "StartText", "Kind", "Recurrence", "OldOrganizer", "OldOrganizerName", "OldOrganizerState", "OldOrganizerDetail",
                "NewOrganizer", "Method", "Status", "NewMeetingId", "NewMeeting", "NewMeetingDetail", "Invited", "Rooms", "OldOrganizerCopy",
                "OldCopiesRemoved", "OldCopiesFailed", "OldCopiesLeft", "Selected", "Notes");
            // The state of each organizer of the run (short, and the detail of the search), by every address it has.
            var state = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase);
            foreach (var o in Items(organizers))
            {
                var st = Text(o, "State");
                var s = Is(st, "Mailbox") ? "Mailbox present" : Is(st, "NoMailbox") ? "No mailbox" : Is(st, "NotInDirectory") ? "Not in the directory" : "Not checked";
                var pair = new[] { s, Text(o, "Detail") };
                foreach (var a in Items(Prop(o, "Addresses"))) { var k = ToText(a); if (k.Length > 0) { state[k] = pair; } }
                var p = Text(o, "PrimaryAddress"); if (p.Length > 0) { state[p] = pair; }
                var i = Text(o, "Input"); if (i.Length > 0 && !state.ContainsKey(i)) { state[i] = pair; }
            }
            foreach (var m in Items(meetings))
            {
                // The meetings of the transfer only: an unticked meeting was not part of it.
                if (!Flag(m, "Selected")) { continue; }
                string newMeeting = "", newDetail = "", oldOrganizerCopy = "";
                int invited = 0, rooms = 0, removed = 0, failed = 0, left = 0;
                foreach (var c in Items(Prop(m, "Copies")))
                {
                    var role = Text(c, "Role"); var result = Text(c, "Result");
                    if (Is(role, "New organizer")) { newMeeting = result; newDetail = Text(c, "Detail"); continue; }
                    if (Text(c, "EventId").Length == 0) { continue; }
                    if (Is(role, "Organizer")) { if (oldOrganizerCopy.Length == 0) { oldOrganizerCopy = result.Length > 0 ? result : "Present"; } }
                    else if (Is(role, "Attendee") || Is(role, "Room"))
                    {
                        if (Is(role, "Room")) { rooms++; } else { invited++; }
                        if (Is(result, "Removed") || Is(result, "Already gone")) { removed++; }
                        else if (Is(result, "Failed") || Is(result, "Not done")) { failed++; }
                        else { left++; }
                    }
                }
                if (oldOrganizerCopy.Length == 0) { oldOrganizerCopy = Text(m, "OrganizerCopy"); }
                var method = Text(m, "TransferMethod");
                var methodText = Is(method, "Recreate") ? "Re-created" : method.Length > 0 ? method : "";
                var key = Text(m, "OrganizerKey"); if (key.Length == 0) { key = Text(m, "Organizer"); }
                string[] orgState;
                if (!state.TryGetValue(key, out orgState) && !state.TryGetValue(Text(m, "Organizer"), out orgState)) { orgState = new[] { "", "" }; }
                var notes = new List<string>();
                foreach (var n in Items(Prop(m, "Notes"))) { notes.Add(ToText(n)); }
                t.Rows.Add(new object[] {
                    Text(m, "MeetingId"), Text(m, "Subject"), Text(m, "StartText"), Text(m, "Kind"), Text(m, "Recurrence"),
                    Text(m, "Organizer"), Text(m, "OrganizerName"), orgState[0], orgState[1],
                    Text(m, "NewOrganizer"), methodText, Text(m, "Status"), Text(m, "NewMeetingId"), newMeeting, newDetail,
                    invited, rooms, oldOrganizerCopy, removed, failed, left, Flag(m, "Selected"), notes.ToArray() });
            }
            return t;
        }

        /// <summary>A CSV file of a table: UTF-8 with BOM, the columns given (in that order), one line per row.</summary>
        public static void WriteTableCsv(Table table, string[] columns, string path, string delimiter)
        {
            var index = new int[columns.Length];
            for (int i = 0; i < columns.Length; i++) { index[i] = Array.IndexOf(table.Columns, columns[i]); }
            var sb = new StringBuilder();
            var cells = new string[columns.Length];
            for (int i = 0; i < columns.Length; i++) { cells[i] = CsvCell(columns[i], delimiter); }
            sb.AppendLine(string.Join(delimiter, cells));
            foreach (var r in table.Rows)
            {
                for (int i = 0; i < columns.Length; i++) { cells[i] = index[i] < 0 ? "" : CsvCell(r[index[i]], delimiter); }
                sb.AppendLine(string.Join(delimiter, cells));
            }
            File.WriteAllText(path, sb.ToString(), new UTF8Encoding(true));
        }

        /// <summary>The rows of a table as JSON objects (one property per column), HTML-safe (no &lt; &gt; &amp;).</summary>
        public static string TableJson(Table table)
        {
            var options = new JsonWriterOptions { Encoder = JavaScriptEncoder.Default };
            using (var stream = new MemoryStream())
            {
                using (var w = new Utf8JsonWriter(stream, options))
                {
                    w.WriteStartArray();
                    foreach (var r in table.Rows)
                    {
                        w.WriteStartObject();
                        for (int i = 0; i < table.Columns.Length; i++) { w.WritePropertyName(table.Columns[i]); WriteJson(w, r[i], 1); }
                        w.WriteEndObject();
                    }
                    w.WriteEndArray();
                }
                return Encoding.UTF8.GetString(stream.ToArray());
            }
        }

        /// <summary>A CSV cell: text starting with = + - @ (or tab, CR) prefixed with an apostrophe; quoted when needed (Format-McoCsvCell).</summary>
        public static string CsvCell(object value, string delimiter)
        {
            if (value == null) { return ""; }
            var v = Base(value);
            string text;
            if (v is bool) { text = (bool)v ? "True" : "False"; }
            else if (v is string)
            {
                text = (string)v;
                if (text.Length > 0 && "=+-@\t\r".IndexOf(text[0]) >= 0) { text = "'" + text; }
            }
            else if (v is IEnumerable && !(v is IDictionary))
            {
                var parts = new List<string>();
                foreach (var x in (IEnumerable)v) { parts.Add(ToText(x)); }
                text = string.Join(" | ", parts);
            }
            else { text = ToText(v); }
            if (text.Contains(delimiter) || text.Contains("\"") || text.IndexOf('\r') >= 0 || text.IndexOf('\n') >= 0) { text = "\"" + text.Replace("\"", "\"\"") + "\""; }
            return text;
        }

        /// <summary>A CSV file: UTF-8 with BOM, the columns given, one line per row (Write-McoCsv).</summary>
        public static void WriteCsv(object rows, string[] columns, string path, string delimiter)
        {
            var sb = new StringBuilder();
            var cells = new string[columns.Length];
            for (int i = 0; i < columns.Length; i++) { cells[i] = CsvCell(columns[i], delimiter); }
            sb.AppendLine(string.Join(delimiter, cells));
            foreach (var r in Items(rows))
            {
                for (int i = 0; i < columns.Length; i++) { cells[i] = CsvCell(Prop(r, columns[i]), delimiter); }
                sb.AppendLine(string.Join(delimiter, cells));
            }
            File.WriteAllText(path, sb.ToString(), new UTF8Encoding(true));
        }

        static void WriteJson(Utf8JsonWriter w, object value, int depth)
        {
            if (value == null || depth > 12) { w.WriteNullValue(); return; }
            var v = Base(value);
            if (v is string) { w.WriteStringValue((string)v); return; }
            if (v is bool) { w.WriteBooleanValue((bool)v); return; }
            if (v is int || v is long || v is short || v is byte) { w.WriteNumberValue(Convert.ToInt64(v, Inv)); return; }
            if (v is double || v is float || v is decimal) { w.WriteNumberValue(Convert.ToDouble(v, Inv)); return; }
            if (v is DateTime) { w.WriteStringValue(((DateTime)v).ToString("o", Inv)); return; }
            var d = v as IDictionary;
            if (d != null)
            {
                w.WriteStartObject();
                foreach (DictionaryEntry e in d) { w.WritePropertyName(ToText(e.Key)); WriteJson(w, e.Value, depth + 1); }
                w.WriteEndObject();
                return;
            }
            if (v is PSCustomObject || (value is PSObject && ((PSObject)value).BaseObject is PSCustomObject))
            {
                w.WriteStartObject();
                foreach (var p in PSObject.AsPSObject(value).Properties) { w.WritePropertyName(p.Name); WriteJson(w, p.Value, depth + 1); }
                w.WriteEndObject();
                return;
            }
            var e2 = v as IEnumerable;
            if (e2 != null)
            {
                w.WriteStartArray();
                foreach (var x in e2) { WriteJson(w, x, depth + 1); }
                w.WriteEndArray();
                return;
            }
            w.WriteStringValue(ToText(v));
        }
    }

    /// <summary>Rows of the report: one array of values per row, in the order of the columns.</summary>
    public sealed class Table
    {
        public Table(params string[] columns) { Columns = columns; Rows = new List<object[]>(); }
        public string[] Columns { get; private set; }
        public List<object[]> Rows { get; private set; }
        public int Count { get { return Rows.Count; } }

        /// <summary>The rows as PowerShell objects, one property per column.</summary>
        public List<PSObject> ToObjects()
        {
            var list = new List<PSObject>();
            foreach (var r in Rows)
            {
                var o = new PSObject();
                for (int i = 0; i < Columns.Length; i++) { o.Properties.Add(new PSNoteProperty(Columns[i], r[i])); }
                list.Add(o);
            }
            return list;
        }
    }
}
