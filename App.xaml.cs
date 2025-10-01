using log4net;
using System.Configuration;
using System.Data;
using System.Text;
using System.Windows;

namespace SIM_TIA_DeviceAlarmgenerator
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        /// <summary>
        /// <para>
        /// Der Logger:
        /// </para>
        /// <para>
        /// Wurde hier so implementiert: <seealso href="App.cs"/> App.config, App.xml.cs, Assembly.cs und dann in den entsprechenden Klassen.
        /// Siehe z.B. <seealso cref="MainWindow"/>
        /// </para>
        /// <<para>
        /// Beispiele wie man log4net einrichtet:
        /// https://michaelhorstmann.de/blog/logging-mit-log4net-unter-net
        /// https://logging.apache.org/log4net/release/config-examples.html
        /// </para>
        /// </summary>
        private static readonly ILog log = LogManager.GetLogger(typeof(App));

        public App()
        {
            //Sollte alles abfangen, was durch das UI geht
            this.DispatcherUnhandledException += App_DispatcherUnhandledException;
            //Sollte auch den Rest abfangen
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
        }

        private void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            log.Fatal("Unhandled Exception");
            if (e.ExceptionObject != null)
            {
                log.Fatal(e.ExceptionObject);
                var fatalMessage = $"Exception raised by: {e.ExceptionObject}\r\nApplication will be terminated: {e.IsTerminating}";
                log.Fatal(fatalMessage);
                _ = MessageBox.Show(fatalMessage, "Unbehandelter Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            else
            {
                _ = MessageBox.Show($"Fehlerursache: Unbekannt\r\nVerursacht: Unbekannt\r\nAnwendung wird beendet: {e.IsTerminating}", "Unbehandelter Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            //e.Handled = true;
        }

        /// <summary>
        /// Gibt den Logger der Anwendung zurück.
        /// </summary>
        /// <returns></returns>
        public static ILog GetLogger()
        {
            return log;
        }

        private void App_DispatcherUnhandledException(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            log.Fatal("Unhandled Exception");
            log.Fatal(e.Exception);

            //Unbehandelter Fehler
            //Build a message and take a look inside the innerExceptions
            var stringBuilder = new StringBuilder();

            _ = stringBuilder.Append("Sender:\n").Append(sender.ToString()).Append("\n\n");
            _ = stringBuilder.Append("Message:\n").Append(e.Exception.Message).Append("\n\n");

            if (e.Exception.InnerException != null)
            {
                _ = stringBuilder.Append("InnerException:\n").Append(e.Exception.InnerException.Message).Append("\n\n").Append("Stacktrace:").Append('\n').Append(e.Exception.InnerException.StackTrace);
            }
            else
            {
                _ = stringBuilder.Append("Stacktrace:\n").Append(e.Exception.StackTrace);
            }

            _ = MessageBox.Show(stringBuilder.ToString(), "Unbehandelter Fehler", MessageBoxButton.OK, MessageBoxImage.Error);

            //e.Handled = true;
        }

        private void Application_Startup(object sender, StartupEventArgs e)
        {
            var logger = GetLogger();
            logger.Debug("####################################################################################");
            logger.Debug("================================== Anwendung gestartet =============================");
            logger.Debug("####################################################################################");
        }

        private void Application_Exit(object sender, ExitEventArgs e)
        {
            var logger = GetLogger();
            logger.Debug("##################################################################################");
            logger.Debug($"======================= Anwendung beendet - Exit Code ({e.ApplicationExitCode}) =======================");
            logger.Debug("##################################################################################");
        }
    }
}
