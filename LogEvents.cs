// ***************************************************************
// Класс LogEvents реализует запись сообщений в журнал
// приложений Windows
// ***************************************************************

namespace Microsoft.Exchange.Agents.MailFilterAgent
{
  using System;
  using System.Diagnostics;

  public class LogEvents
  {
    // идентификаторы сообщений из файла messages.h
    public const Int64 MSG_EXCEPTION = 0xC0010002L;
    public const Int64 MSG_SEND_PROHIBITED = 0x80010001L;

    private const string sourceName = "MailFilter";

    public void LogMessage(EventInstance event_instance, params Object[] Strings)
    {
      EventLog myEventLog = new EventLog(EventLog.LogNameFromSourceName(sourceName, "."), ".", sourceName);
      myEventLog.WriteEvent(event_instance, Strings);
    } 
  }
}
