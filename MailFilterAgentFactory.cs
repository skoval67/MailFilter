// ***************************************************************
// A factory for the MailFilter agent.
// ***************************************************************

namespace Microsoft.Exchange.Agents.MailFilterAgent
{
  using System.IO;
  using System.Reflection;

  using Microsoft.Exchange.Data.Transport;
  using Microsoft.Exchange.Data.Transport.Routing;

  // The agent factory
  public class MailFilterAgentFactory : RoutingAgentFactory
  {
    // An object that contains settings to be used by agents.
    private MailFilterSettings MailFilterSettings;

    // Factory constructor.
    public MailFilterAgentFactory()
    {
      Assembly currAssembly = Assembly.GetAssembly(this.GetType());

      // Read the XML configuration file and apply its settings.
      this.MailFilterSettings = new MailFilterSettings(Path.GetDirectoryName(currAssembly.Location));
    }

    // Create a new MailFilter Agent.
    public override RoutingAgent CreateAgent(SmtpServer server)
    {
      return new MailFilterAgent(this.MailFilterSettings, server);
    }
  }
}