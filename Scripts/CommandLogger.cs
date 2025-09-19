using System;
using System.Globalization;
using System.IO;
using System.Text;
using Autodesk.Revit.UI;

namespace WallRvt.Scripts
{
    /// <summary>
    /// Логгер для команд Revit, который записывает диагностику в файл и при необходимости
    /// показывает сообщения в <see cref="TaskDialog"/>.
    /// </summary>
    /// <remarks>
    /// Экземпляр создаётся на время выполнения команды. Если файл открыть невозможно,
    /// логгер продолжит работу в «беззвучном» режиме, чтобы диагностика не мешала основной логике.
    /// </remarks>
    internal sealed class CommandLogger : IDisposable
    {
        private readonly object _syncRoot = new object();
        private readonly StreamWriter _writer;
        private readonly bool _mirrorToDialog;
        private bool _disposed;

        private CommandLogger(StreamWriter writer, string logFilePath, bool mirrorToDialog, string sessionId)
        {
            _writer = writer;
            _mirrorToDialog = mirrorToDialog;
            LogFilePath = logFilePath;
            SessionId = sessionId;
        }

        /// <summary>
        /// Полный путь к файлу журнала. Может быть <c>null</c>, если файл создать не удалось.
        /// </summary>
        public string LogFilePath { get; }

        /// <summary>
        /// Идентификатор текущей сессии логирования.
        /// </summary>
        public string SessionId { get; }

        /// <summary>
        /// Создаёт логгер для конкретной команды.
        /// </summary>
        /// <param name="commandName">Имя команды (используется в заголовках и имени файла).</param>
        /// <param name="mirrorToDialog">Если true, все сообщения дополнительно показываются в TaskDialog.</param>
        /// <returns>Экземпляр <see cref="CommandLogger"/>.</returns>
        public static CommandLogger CreateForCommand(string commandName, bool mirrorToDialog = false)
        {
            string sessionId = Guid.NewGuid().ToString("N", CultureInfo.InvariantCulture);
            StreamWriter writer = null;
            string logFilePath = null;

            try
            {
                string safeCommandName = string.IsNullOrWhiteSpace(commandName) ? "Command" : commandName.Trim();
                string fileName = $"{safeCommandName}_{DateTime.Now:yyyyMMdd_HHmmssfff}.log";
                string tempDirectory = Path.GetTempPath();
                logFilePath = Path.Combine(tempDirectory, fileName);

                FileStream stream = new FileStream(logFilePath, FileMode.Create, FileAccess.Write, FileShare.Read);
                writer = new StreamWriter(stream, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false))
                {
                    AutoFlush = true
                };

                string header = $"{FormatTimestamp()} [{sessionId}] Старт сессии команды {safeCommandName}.";
                writer.WriteLine(header);
            }
            catch (Exception)
            {
                writer?.Dispose();
                writer = null;
                logFilePath = null;
            }

            return new CommandLogger(writer, logFilePath, mirrorToDialog, sessionId);
        }

        /// <summary>
        /// Записывает строку в журнал.
        /// </summary>
        /// <param name="message">Текст сообщения.</param>
        /// <param name="showDialog">Показывать ли сообщение в TaskDialog независимо от настройки логгера.</param>
        public void Log(string message, bool showDialog = false)
        {
            if (_disposed || string.IsNullOrWhiteSpace(message))
            {
                return;
            }

            string formatted = $"{FormatTimestamp()} [{SessionId}] {message}";
            WriteLineSafe(formatted);

            if (_mirrorToDialog || showDialog)
            {
                ShowDialogSafe(message);
            }
        }

        /// <summary>
        /// Форматирует строку по шаблону и записывает её в журнал.
        /// </summary>
        public void LogFormat(string format, params object[] args)
        {
            if (string.IsNullOrEmpty(format))
            {
                return;
            }

            string message = args == null || args.Length == 0
                ? format
                : string.Format(CultureInfo.InvariantCulture, format, args);

            Log(message);
        }

        /// <summary>
        /// Записывает сообщение и стек исключения.
        /// </summary>
        /// <param name="context">Контекст или сопроводительный текст.</param>
        /// <param name="exception">Исключение.</param>
        public void LogException(string context, Exception exception)
        {
            if (_disposed)
            {
                return;
            }

            StringBuilder builder = new StringBuilder();

            if (!string.IsNullOrWhiteSpace(context))
            {
                builder.AppendLine(context);
            }

            if (exception != null)
            {
                builder.AppendLine($"Тип исключения: {exception.GetType().FullName}");
                builder.AppendLine($"Сообщение: {exception.Message}");
                builder.AppendLine("Стек вызовов:");
                builder.AppendLine(exception.ToString());
            }

            Log(builder.ToString().TrimEnd());
        }

        /// <inheritdoc />
        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;

            try
            {
                WriteLineSafe($"{FormatTimestamp()} [{SessionId}] Завершение сессии логирования.");
            }
            finally
            {
                _writer?.Dispose();
            }
        }

        private static string FormatTimestamp()
        {
            return DateTime.Now.ToString("O", CultureInfo.InvariantCulture);
        }

        private void WriteLineSafe(string line)
        {
            if (_writer == null)
            {
                return;
            }

            try
            {
                lock (_syncRoot)
                {
                    _writer.WriteLine(line);
                }
            }
            catch (Exception)
            {
                // Любые ошибки ввода-вывода игнорируются, чтобы не прерывать основную команду.
            }
        }

        private static void ShowDialogSafe(string message)
        {
            try
            {
                TaskDialog.Show("WallLayerSplitter", message);
            }
            catch (Exception)
            {
                // Если TaskDialog недоступен (например, в контексте тестов), просто пропускаем показ.
            }
        }
    }
}

