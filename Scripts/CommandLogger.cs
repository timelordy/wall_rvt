using System;
using System.Diagnostics;
using System.Globalization;
using System.Text;

namespace WallRvt.Scripts
{
    /// <summary>
    /// Простейший логгер, записывающий диагностические сообщения команды в отладочный вывод.
    /// </summary>
    internal sealed class CommandLogger : IDisposable
    {
        private readonly string _commandName;
        private readonly StringBuilder _buffer = new StringBuilder();
        private bool _disposed;

        private CommandLogger(string commandName)
        {
            _commandName = string.IsNullOrWhiteSpace(commandName) ? "UnknownCommand" : commandName.Trim();
            Log("Логгер создан.");
        }

        /// <summary>
        /// Создаёт экземпляр логгера для указанной команды.
        /// </summary>
        public static CommandLogger CreateForCommand(string commandName)
        {
            try
            {
                return new CommandLogger(commandName);
            }
            catch
            {
                // В случае непредвиденной ошибки не блокируем выполнение команды.
                return null;
            }
        }

        /// <summary>
        /// Записывает текстовое сообщение.
        /// </summary>
        public void Log(string message)
        {
            if (_disposed || string.IsNullOrWhiteSpace(message))
            {
                return;
            }

            string entry = FormatEntry(message);
            _buffer.AppendLine(entry);
            Debug.WriteLine(entry);
        }

        /// <summary>
        /// Записывает сообщение, используя формат с параметрами.
        /// </summary>
        public void LogFormat(string format, params object[] args)
        {
            if (_disposed || string.IsNullOrWhiteSpace(format))
            {
                return;
            }

            string formatted = args == null || args.Length == 0
                ? format
                : string.Format(CultureInfo.InvariantCulture, format, args);
            Log(formatted);
        }

        /// <summary>
        /// Записывает сообщение об исключении.
        /// </summary>
        public void LogException(string message, Exception exception)
        {
            if (_disposed)
            {
                return;
            }

            if (!string.IsNullOrWhiteSpace(message))
            {
                Log(message);
            }

            if (exception != null)
            {
                Log($"Исключение: {exception.GetType().FullName}: {exception.Message}");
                if (!string.IsNullOrWhiteSpace(exception.StackTrace))
                {
                    Debug.WriteLine(exception.StackTrace);
                }
            }
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            if (_buffer.Length > 0)
            {
                Debug.WriteLine($"[{_commandName}] Завершение логирования. Всего сообщений: {_buffer.ToString().Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries).Length}.");
            }
        }

        private string FormatEntry(string message)
        {
            string timestamp = DateTime.Now.ToString("HH:mm:ss", CultureInfo.InvariantCulture);
            return $"[{timestamp}] [{_commandName}] {message.Trim()}";
        }
    }
}
