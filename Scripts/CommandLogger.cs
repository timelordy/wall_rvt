using System;
using System.IO;
using System.Text;
using Autodesk.Revit.UI;

namespace WallRvt.Scripts
{
    /// <summary>
    /// Простой логгер для записи диагностических сообщений команды.
    /// </summary>
    /// <remarks>
    /// Лог создаётся в каталоге временных файлов операционной системы и
    /// предназначен для записи только в рамках одной сессии выполнения команды.
    /// Все операции ввода-вывода безопасно обёрнуты в блоки try/catch, чтобы
    /// логирование не влияло на основную бизнес-логику.
    /// </remarks>
    internal class CommandLogger
    {
        private readonly object _syncRoot = new object();
        private readonly string _logFilePath;
        private readonly bool _mirrorToDialog;

        private CommandLogger(string logFilePath, bool mirrorToDialog)
        {
            _logFilePath = logFilePath;
            _mirrorToDialog = mirrorToDialog;
        }

        /// <summary>
        /// Создаёт новый экземпляр логгера и подготавливает файл для записи.
        /// </summary>
        /// <param name="logFileName">Имя файла журнала. Если не указано, создаётся имя с отметкой времени.</param>
        /// <param name="mirrorToDialog">Если true, сообщения дублируются в окно TaskDialog.</param>
        /// <returns>Экземпляр <see cref="CommandLogger"/>.</returns>
        public static CommandLogger Create(string logFileName = null, bool mirrorToDialog = false)
        {
            string filePath = null;

            try
            {
                string tempDirectory = Path.GetTempPath();
                string fileName = string.IsNullOrWhiteSpace(logFileName)
                    ? $"WallLayerSplitter_{DateTime.Now:yyyyMMdd_HHmmssfff}.log"
                    : logFileName;

                filePath = Path.Combine(tempDirectory, fileName);

                using (FileStream stream = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.Read))
                using (StreamWriter writer = new StreamWriter(stream, Encoding.UTF8))
                {
                    writer.WriteLine($"[{DateTime.Now:O}] Старт новой сессии логирования.");
                }
            }
            catch (Exception)
            {
                // Игнорируем любые проблемы с доступом к файловой системе.
                filePath = null;
            }

            return new CommandLogger(filePath, mirrorToDialog);
        }

        /// <summary>
        /// Записывает строку в журнал.
        /// </summary>
        /// <param name="message">Сообщение.</param>
        /// <param name="showDialog">Если true, сообщение дополнительно показывается в TaskDialog.</param>
        public void Log(string message, bool showDialog = false)
        {
            if (string.IsNullOrWhiteSpace(message))
            {
                return;
            }

            string formattedMessage = $"[{DateTime.Now:O}] {message}";

            if (!string.IsNullOrWhiteSpace(_logFilePath))
            {
                try
                {
                    lock (_syncRoot)
                    {
                        File.AppendAllText(_logFilePath, formattedMessage + Environment.NewLine, Encoding.UTF8);
                    }
                }
                catch (Exception)
                {
                    // Ошибки логирования игнорируются, чтобы не прерывать основной сценарий.
                }
            }

            if (_mirrorToDialog || showDialog)
            {
                try
                {
                    TaskDialog.Show("WallLayerSplitter", message);
                }
                catch (Exception)
                {
                    // Игнорируем ошибки при показе диалога (например, если API недоступно).
                }
            }
        }

        /// <summary>
        /// Записывает информацию об исключении в журнал.
        /// </summary>
        /// <param name="message">Сопроводительный текст.</param>
        /// <param name="exception">Исключение.</param>
        public void LogException(string message, Exception exception)
        {
            StringBuilder builder = new StringBuilder();
            builder.AppendLine(message);

            if (exception != null)
            {
                builder.AppendLine($"Тип исключения: {exception.GetType().FullName}");
                builder.AppendLine($"Сообщение: {exception.Message}");
                builder.AppendLine("Стек вызовов:");
                builder.AppendLine(exception.ToString());
            }

            Log(builder.ToString());
        }

        /// <summary>
        /// Полный путь к файлу журнала или null, если журнал не используется.
        /// </summary>
        public string LogFilePath => _logFilePath;
    }
}

