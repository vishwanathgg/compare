using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ConsoleApp1
{
    public class LogStore
    { 
        private object _lockObject = new object();
        private List<Log> _logObjects= new List<Log>();

        public List<Log> GetLogs()
        {
            lock (_logObjects)
            {
                return _logObjects;
            }
        }

        public List<Log> SetLogs(string[] logContent)
        {
            lock (_logObjects)
            {
                return _logObjects = GetLogObjects(logContent);
            }
        }

        private List<Log> GetLogObjects(string[] logContent)
        {
            List<Log> logs = new List<Log>();
            int number = 1;
            for (int i = 0; i < logContent.Length; i++)
            {
                string logMessage;
                var re = GetLogAttributes(logContent[i]);
                if (re.isMatchedLogFormat)
                {
                    logMessage = re.logMessage;
                    int nextLineIndex = i + 1;
                    while (true)
                    {
                        if (nextLineIndex > logContent.Length - 1)
                            break;
                        var reNextLine = GetLogAttributes(logContent[nextLineIndex]);
                        if (!reNextLine.isMatchedLogFormat)
                        {
                            logMessage = logMessage + Environment.NewLine + reNextLine.logMessage;
                            nextLineIndex++;
                            i++;
                        }
                        else
                        {
                            break;
                        }
                    }
                    DateTime? d = DateTime.TryParse(re.timeStamp, out DateTime dt) ? dt : null;
                    var r = new Log()
                    {
                        Number = number++,
                        LogTime = d,
                        LogType = re.logType,
                        Thread = re.thread,
                        LogMessage = logMessage
                    };
                    logs.Add(r);
                }

            }
            return logs;
        }

        private (bool isMatchedLogFormat, string logType, string timeStamp, string thread, string logMessage) GetLogAttributes(string log)
        {
            Regex pattern = new Regex(@"(\S+) " +
                            @"(\d{4}-\d{1,2}-\d{2} \d{2}:\d{2}:\d{2}.\d{4}) " +
                            @"\[(\d{1,2})\] " +
                            @"(.*)");

            Match match = pattern.Match(log);
            if (match.Success)
            {
                var res = (true, match.Groups[1].Value, match.Groups[2].Value, match.Groups[3].Value, match.Groups[4].Value);
                return res;
            }
            else
            {
                return (false, "", "", "", log);
            }
        }
    }

    public enum LogServiceType
    {
        ReportTriggerServiceLog,
        GeneratationServiceLog,
        DistributionServiceLog,
        ArchiveServiceLog
    }
        

    public class LogService
    {
        private LogStore _reportTriggerServiceLogStore = new LogStore();
        private LogStore _generateServiceLogStore = new LogStore();
        private LogStore _distributionServiceLogStore = new LogStore();
        private LogStore _archiveServiceLogStore = new LogStore();

        List<Log> logObjects = new List<Log>();
        LogAnalyzer logAnalyzer = new LogAnalyzer();

        public void PopulateLogStore(LogServiceType logServiceType, string[] logContent)
        {
            switch (logServiceType)
            {
                case LogServiceType.ReportTriggerServiceLog:
                    _reportTriggerServiceLogStore.SetLogs(logContent);
                    break;
                case LogServiceType.GeneratationServiceLog:
                    _generateServiceLogStore.SetLogs(logContent);
                    break;
                case LogServiceType.DistributionServiceLog:
                    _distributionServiceLogStore.SetLogs(logContent);
                    break;
                case LogServiceType.ArchiveServiceLog:
                    _archiveServiceLogStore.SetLogs(logContent);
                    break;
            }
        }

        public List<Log> GetSpecificLogs(LogServiceType logServiceType, string key)
        {
            var res = logAnalyzer.GetSpecificLogs( GetLogsFromStore(logServiceType), key);
            return res;
        }

        public List<Log> GetRelatedLogs(LogServiceType logServiceType, string key, string groupKey = "")
        {
            var res = logAnalyzer.GetRelatedLogs(GetLogsFromStore(logServiceType), key, groupKey);
            return res;
        }

        public List<Log> GetRelatedLogsForLogLine(LogServiceType logServiceType, int lineNumber)
        {
            var res = logAnalyzer.GetRelatedLogsForLogLine(GetLogsFromStore(logServiceType), lineNumber);
            return res;
        }

        public List<Log> GetAllLogsInSnapshot(LogServiceType logServiceType, string key)
        {
            var res = logAnalyzer.GetAllLogsInSnapshot(GetLogsFromStore(logServiceType), key);
            return res;
        }

        private IEnumerable<Log> GetLogsFromStore(LogServiceType logServiceType)
        {
            switch (logServiceType)
            {
                case LogServiceType.ReportTriggerServiceLog:
                    return _reportTriggerServiceLogStore.GetLogs();
                case LogServiceType.GeneratationServiceLog:
                    return _generateServiceLogStore.GetLogs();
                case LogServiceType.DistributionServiceLog:
                    return _distributionServiceLogStore.GetLogs();
                case LogServiceType.ArchiveServiceLog:
                    return _archiveServiceLogStore.GetLogs();
                default:
                    return new List<Log>();
            }
        }
    }

    public class LogAnalyzer
    {
        public List<Log> GetSpecificLogs(IEnumerable<Log> logs, string key)
        {
            var res = logs.Where(x => x.LogMessage.Contains(key)).ToList();
            return res;
        }

        public List<Log> GetRelatedLogs(IEnumerable<Log> logs, string key, string groupKey = "")
        {
            List<Log> logResult = new List<Log>();
            Log firstLog = logs.FirstOrDefault(x => x.LogMessage.Contains(key));
            if (firstLog != null)
            {
                logResult = logs.Where(x => x.Thread == firstLog.Thread).ToList();
                logResult = FilterBasedOnGroup(logResult, key, groupKey);
            }
            return logResult;
        }

        public List<Log> GetRelatedLogsForLogLine(IEnumerable<Log> logs, int lineNumber)
        {
            List<Log> logResult = new List<Log>();
            Log lg = logs.FirstOrDefault(x => x.Number == lineNumber);

            if (lg != null)
            {
                var filteredByThreadLogs = logs.Where(x => x.Thread == lg.Thread).ToList();
                string group = GetImmediateGroupForLine(filteredByThreadLogs, lineNumber);
                if (group != null)
                {
                    logResult = GetRelatedLogs(filteredByThreadLogs, group, group);
                }
            }
            return logResult;
        }

        public List<Log> GetAllLogsInSnapshot(IEnumerable<Log> logs, string key)
        {
            List<Log> logResult = new List<Log>();
            var inputLogs = logs.ToList();
            int startLineIndex = inputLogs.FindIndex(x => x.LogMessage.Contains(key));
            if (startLineIndex != -1)
            {
                int endIndex = (startLineIndex + 500) > logs.Count() ? logs.Count() - 1 : (startLineIndex + 500);

                for (int i = startLineIndex; i <= endIndex; i++)
                {
                    logResult.Add(inputLogs[i]);
                }
            }
            return logResult;
        }

        private string GetImmediateGroupForLine(List<Log> logs, int lineNumber)
        {
            int startLineIndex = logs.FindIndex(x => x.Number == lineNumber);

            for (int i = startLineIndex; i >= 0; i--)
            {
                string lg = logs[i].LogMessage;
                var group = CheckForGroupKey(lg);
                if (group != null)
                    return group;
            }
            return null;
        }

        private List<Log> FilterBasedOnGroup(List<Log> logs, string key, string groupKey = "")
        {
            List<Log> filteredLogs = new List<Log>();
            int startLineIndex = logs.FindIndex(x => x.LogMessage.Contains(key));

            for (int i = startLineIndex; i >= 0; i--)
            {

                string lg = logs[i].LogMessage;

                var grp = CheckForGroupKey(lg);
                if (grp != null && grp != groupKey)
                    break;
                else
                    filteredLogs.Insert(0, logs[i]);
            }

            for (int i = startLineIndex + 1; i < logs.Count(); i++)
            {
                string lg = logs[i].LogMessage;

                var grp = CheckForGroupKey(lg);
                if (grp != null && grp != groupKey)
                    break;
                else
                    filteredLogs.Add(logs[i]);
            }

            return filteredLogs;
        }

        private string CheckForGroupKey(string logMessage)
        {
            Regex pattern = new Regex(@"""GroupId"":""(.*)""");
            Match match = pattern.Match(logMessage);
            if (match.Success)
            {
                string grpId = match.Groups[1].Value;
                return grpId;
            }
            else
                return null;
        }
    }
}
