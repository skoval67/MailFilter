// ***************************************************************
// Настройки
// ***************************************************************

namespace Microsoft.Exchange.Agents.MailFilterAgent
{
  using System;
  using System.Diagnostics;
  using System.IO;
  using System.Xml;
  using System.Net;
  using System.DirectoryServices;
  using System.Collections.Generic;

  // This class stores settings that will be used by the MailFilter agent and factory.
  public class MailFilterSettings
  {
#if (DEBUG)
    private string _LOG_FILE_PATH = @"C:\mailfilteragent.log";
#endif
    List<IPAddress> _hosts = new List<IPAddress>();

    private string _Type4ArmsGroupDN;
    private string _AllowSend2InternetGroupDN;
    private string _SecurityAuditEmail;

    private bool _correct;

    private LogEvents logger;

    // DataPath the path to an XML file that contains the settings
    public MailFilterSettings(string DataPath)
    {
#if (DEBUG)
      _LOG_FILE_PATH = Path.Combine(DataPath, "MailFilterAgent.log");
#endif

      logger = new LogEvents();

      // Read nondefault settings from file.
      _correct = this.ReadXMLConfig(Path.Combine(DataPath, "MailFilterConfig.xml"));
    }

    #region Parameters

#if (DEBUG)
    public string LOG_FILE_PATH
    {
      get { return this._LOG_FILE_PATH; }
      set { this._LOG_FILE_PATH = value; }
    }
#endif

    public string Type4ArmsGroupDN
    {
      get { return this._Type4ArmsGroupDN; }
      set { this._Type4ArmsGroupDN = value; }
    }

    public string AllowSend2InternetGroupDN
    {
      get { return this._AllowSend2InternetGroupDN; }
      set { this._AllowSend2InternetGroupDN = value; }
    }

    public string SecurityAuditEmail
    {
      get { return this._SecurityAuditEmail; }
      set { this._SecurityAuditEmail = value; }
    }

    public bool correct
    {
      get { return this._correct; }
      set { this._correct = value; }
    }

    public List<IPAddress> hosts
    {
      get { return this._hosts; }
    }

    #endregion Parameters

    // Reads in configuration options from an XML file and sets the instance variables to the corresponding values that are read in if they
    // are valid. If an invalid value is found, or a value is not set in the XML file the function return false.
    public bool ReadXMLConfig(string path)
    {
      IPAddress dummy;
      bool retval = false;

      try
      {
        // Load the file into the XML reader.
        XmlDocument xmlDoc = new XmlDocument();
        xmlDoc.Load(path);

        XmlNode xmlRoot = xmlDoc.SelectSingleNode("BlackListConfig/group[@type='Type4Arms']");
        _Type4ArmsGroupDN = xmlRoot.Attributes["name"].Value;
        CheckExist(_Type4ArmsGroupDN);
 
        xmlRoot = xmlDoc.SelectSingleNode("BlackListConfig/group[@type='AllowSend2Internet']");
        _AllowSend2InternetGroupDN = xmlRoot.Attributes["name"].Value;
        CheckExist(_AllowSend2InternetGroupDN);

        xmlRoot = xmlDoc.SelectSingleNode("BlackListConfig/group[@type='SecurityAudit']");
        _SecurityAuditEmail = xmlRoot.Attributes["email"].Value;

        foreach (XmlNode host in xmlDoc.DocumentElement.SelectNodes("/BlackListConfig/arm"))
        {
          if (IPAddress.TryParse(host.Attributes["ipaddress"].Value, out dummy))
            _hosts.Add(dummy);
          else
            throw new InvalidOperationException(String.Format("Invalid IP-address: {0}, name: {1}", host.Attributes["ipaddress"].Value, host.Attributes["name"].Value));
        }

        retval = true;
      }
      catch (Exception ex)
      {
        logger.LogMessage(new EventInstance(LogEvents.MSG_EXCEPTION, 0, EventLogEntryType.Error), new string[] { "ReadXMLConfig", ex.GetType().FullName, ex.Message });
      }

      return retval;
    }

    private void CheckExist(string dname)
    {
      if (!DirectoryEntry.Exists(String.Format("LDAP://{0}", dname)))
        throw new InvalidOperationException(String.Concat("No such group in the domain - ", dname));
    }

#if (DEBUG)
    public void LogMessage(string eventName, string message)
    {
      TextWriter tw = File.AppendText(LOG_FILE_PATH);
      tw.WriteLine(DateTime.Now + "\t" + eventName + " - " + message);
      tw.Close();
    }
#endif
  }
}
