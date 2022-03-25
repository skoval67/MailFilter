// ***************************************************************
// Транспортный адаптер для MS Exchange 2010
// реализует требование ГУБиЗИ о запрете отправки почты в Интернет
// с компьютеров, на которых ведется обработка ДСП информации (АРМ
// 4 типа)
// ***************************************************************
#define PRODUCTION

namespace Microsoft.Exchange.Agents.MailFilterAgent
{
  using System;
  using System.Net;
  using System.Diagnostics;
  using System.Collections.Generic;
  using System.Text.RegularExpressions;
  using System.DirectoryServices;
  using System.DirectoryServices.AccountManagement;

  using Microsoft.Exchange.Data.Transport;
  using Microsoft.Exchange.Data.Transport.Smtp;
  using Microsoft.Exchange.Data.Mime;
  using Microsoft.Exchange.Data.Transport.Routing;
  using Microsoft.Exchange.Data.Transport.Email;

  // Agent class
  public class MailFilterAgent : RoutingAgent
  {
    // The error message that will be sent to the client if you want to temporarily reject the message.
    private static readonly SmtpResponse expandSmtpResponse = new SmtpResponse("250", "2.1.5", "Replaced all recipients with a single recipient");

    // A reference to the server object.
    private SmtpServer server;

    // A reference to a MailFilter settings object.
    private MailFilterSettings settings;

    // записыватель сообщений в журнал приложений Windows
    private LogEvents mylog = new LogEvents();

    // The constructor registers all event handlers. It should only be called
    // from the MailFilterAgentFactory class.
    public MailFilterAgent(MailFilterSettings settings, SmtpServer server)
    {
      // Initialize instance variables.
      this.settings = settings;
      this.server = server;

      // Set up the hooks to have your functions called when certain events occur.
      this.OnResolvedMessage += new ResolvedMessageEventHandler(OnResolvedMessageHandler);
    }

    public void OnResolvedMessageHandler(ResolvedMessageEventSource source, QueuedMessageEventArgs e)
    {
      if (!settings.correct)
        return;

      try
      {
        Header[] headers = e.MailItem.Message.RootPart.Headers.FindAll("X-Originating-IP");
        if (headers == null || headers.Length == 0)
        {
          throw new InvalidOperationException("Unknown host.");                   // нет заголовка X-Originating-IP
        }
        if (headers.Length > 1)
          throw new InvalidOperationException("Only one sender host allowed.");   // несколько заголовков X-Originating-IP

        String ipaddrstr = headers[0].Value.Replace("[", "").Replace("]", "");

        // Если почта предназначена для отправки в Интернет, отправитель имеет разрешение на переписку с Интернет
        // и АРМ с которого отправлена почта 4 типа, то ...
        if (Mail2Internet(e.MailItem.Recipients) && InetSendValid(e.MailItem.FromAddress) && ArmIsType4(IPAddress.Parse(ipaddrstr)))
        {
          // пишем сообщение в журнал приложений и шлем привет нарушителю
          mylog.LogMessage(new EventInstance(LogEvents.MSG_SEND_PROHIBITED, 0, EventLogEntryType.Warning), new string[] { ipaddrstr, e.MailItem.FromAddress.ToString() });

          EmailMessage origMsg = e.MailItem.Message;

          // Replace all but the last recipient with nothing and replace the last recipient with a sender.
          List<RecipientExpansionInfo> expansionInfoList = new List<RecipientExpansionInfo>();
          for (int i = 0; i < e.MailItem.Recipients.Count - 1; i++)
          {
            EnvelopeRecipient envelopeRecipient = e.MailItem.Recipients[i];
            expansionInfoList.Add(new RecipientExpansionInfo(envelopeRecipient, new RoutingAddress[] { }, expandSmtpResponse));
          }

          RoutingAddress expandRecipient = new RoutingAddress(origMsg.Sender.SmtpAddress);
          expansionInfoList.Add(new RecipientExpansionInfo(e.MailItem.Recipients[e.MailItem.Recipients.Count - 1],
            new RoutingAddress[] { expandRecipient, new RoutingAddress(settings.SecurityAuditEmail) },
            expandSmtpResponse));

          origMsg.From = new EmailRecipient("Service Account", "SYSTEM");
          origMsg.Subject = String.Format("Попытка отправки почты в Интернет с АРМ четвертого типа: {0}", ipaddrstr);

          source.ExpandRecipients(expansionInfoList);
          source.Defer(TimeSpan.Zero);
        }
      }
      catch (Exception ex)
      {
        mylog.LogMessage(new EventInstance(LogEvents.MSG_EXCEPTION, 0, EventLogEntryType.Error), new string[] { "OnResolvedMessage", ex.GetType().FullName, ex.Message});
      }
    }
   
    private bool Mail2Internet(EnvelopeRecipientCollection recipients)
    {
      // Возвращает true, если адрес хотя бы одного получателя не соответствует маске
      bool return_value = false;
#if !(PRODUCTION)
      Regex rx = new Regex(@"^(\w+\.)?example.com$");     // соответствует example.com или *.example.com
#else
      Regex rx = new Regex(@"^(\w+\.)?cbr.ru$");          // соответствует cbr.ru или *.cbr.ru
#endif
      try
      {
        foreach (EnvelopeRecipient currentRecipient in recipients)
        {
          if (!rx.IsMatch(currentRecipient.Address.DomainPart.ToLower()))
          {
            return_value = true;
            break;
          }
        }
      }
      catch (Exception ex)
      {
        mylog.LogMessage(new EventInstance(LogEvents.MSG_EXCEPTION, 0, EventLogEntryType.Error), new string[] { "Mail2Internet", ex.GetType().FullName, ex.Message });
      }

      return return_value;
    }

    private bool InetSendValid(RoutingAddress Sender)
    {
      // Возвращает true если smtp-адрес хотя бы одного члена группы AllowSend2InternetGroupDN совпадает с адресом Sender
      bool found = false;

      try
      {
        DirectoryEntry group = new DirectoryEntry(String.Format("LDAP://{0}", settings.AllowSend2InternetGroupDN));
        object members = group.Invoke("Members", null);
        foreach (object member in (System.Collections.IEnumerable) members)
        {
          DirectoryEntry x = new DirectoryEntry(member);
          found = x.Properties["mail"].Value.ToString().Equals(Sender.ToString());
          if (found) break;
        }
      }
      catch (Exception ex)
      {
        mylog.LogMessage(new EventInstance(LogEvents.MSG_EXCEPTION, 0, EventLogEntryType.Error), new string[] { "InetSendValid", ex.GetType().FullName, ex.Message });
      }

      return found;
    }

    private bool ArmIsType4(IPAddress ipaddr)
    {
      // Возвращает true если компьютер отправителя указан в конфигурационном файле или входит в группу Type4ArmsGroupDN
      bool found = false;

      try
      {
        foreach (IPAddress ip in settings.hosts)
        {
          if (ipaddr.Equals(ip))
            return true;
        }

        String host = Dns.GetHostEntry(ipaddr).HostName;
        
        foreach (IPAddress ip in Dns.GetHostAddresses(host))
        {
          found = ipaddr.Equals(ip);
          if (found) break;
        }

        if (!found)
          throw new InvalidOperationException(String.Format("Записи в прямой и обратной зонах DNS не совпадают для hostname: {0}, ip: {1}", host, ipaddr.ToString()));

        int dotpos = host.IndexOf(".");
        if (dotpos > 0 && dotpos < host.Length)
          host = host.Substring(0, dotpos);

        ComputerPrincipal computer = ComputerPrincipal.FindByIdentity(new PrincipalContext(ContextType.Domain), host);
        if (computer == null)
        {
          found = false;
          throw new InvalidOperationException(String.Concat("Нет такого компьютера в домене - ", host));
        }

        foreach (Principal result in computer.GetGroups())
        {
          found = result.DistinguishedName == settings.Type4ArmsGroupDN;
          if (found) break;
        }
      }
      catch (Exception ex)
      {
        mylog.LogMessage(new EventInstance(LogEvents.MSG_EXCEPTION, 0, EventLogEntryType.Error), new string[] { "ArmIsType4", ex.GetType().FullName, ex.Message });
      }

      return found;
    }
  }
}