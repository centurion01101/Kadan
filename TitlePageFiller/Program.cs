using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace TitlePageFiller
{
    /// <summary>
    /// Класс для автоматического заполнения титульного листа технологического процесса по шаблону Word (.docx).
    /// </summary>
    public static class TitlePageFiller
    {
        /// <summary>
        /// Заполняет шаблон документа значениями из словаря.
        /// </summary>
        /// <param name="templatePath">Путь к файлу шаблона (.docx).</param>
        /// <param name="outputPath">Путь для сохранения заполненного документа.</param>
        /// <param name="data">Словарь с данными для подстановки (ключ — имя поля без скобок, значение — текст).</param>
        /// <exception cref="FileNotFoundException">Выбрасывается, если файл шаблона не найден.</exception>
        /// <exception cref="UnauthorizedAccessException">Выбрасывается при ошибке доступа к файлам.</exception>
        /// <exception cref="Exception">Выбрасывается при других ошибках обработки документа.</exception>
        public static void FillTitlePage(string templatePath, string outputPath, Dictionary<string, string> data)
        {
            if (!File.Exists(templatePath))
            {
                throw new FileNotFoundException($"Файл шаблона не найден: {templatePath}");
            }

            // Копируем шаблон во временный файл для работы, чтобы не изменять оригинал
            string tempPath = Path.GetTempFileName();
            try
            {
                File.Copy(templatePath, tempPath, true);

                using (WordprocessingDocument doc = WordprocessingDocument.Open(tempPath, true))
                {
                    // Обрабатываем основное тело документа
                    var body = doc.MainDocumentPart?.Document?.Body;
                    if (body != null)
                    {
                        ReplacePlaceholdersInElement(body, data);
                    }

                    // Обрабатываем колонтитулы (header и footer)
                    if (doc.MainDocumentPart != null)
                    {
                        foreach (var headerPart in doc.MainDocumentPart.HeaderParts)
                        {
                            var header = headerPart.Header;
                            if (header != null)
                            {
                                ReplacePlaceholdersInElement(header, data);
                            }
                        }

                        foreach (var footerPart in doc.MainDocumentPart.FooterParts)
                        {
                            var footer = footerPart.Footer;
                            if (footer != null)
                            {
                                ReplacePlaceholdersInElement(footer, data);
                            }
                        }
                    }

                    // Сохраняем изменения
                    doc.Save();
                }

                // Копируем результат в выходной файл
                // Если файл уже существует, удаляем его
                if (File.Exists(outputPath))
                {
                    File.Delete(outputPath);
                }
                File.Copy(tempPath, outputPath, true);
            }
            catch (UnauthorizedAccessException ex)
            {
                throw new UnauthorizedAccessException($"Ошибка доступа к файлу: {ex.Message}", ex);
            }
            catch (IOException ex)
            {
                throw new IOException($"Ошибка ввода-вывода: {ex.Message}", ex);
            }
            catch (Exception ex)
            {
                throw new Exception($"Ошибка при обработке документа: {ex.Message}", ex);
            }
            finally
            {
                // Удаляем временный файл
                if (File.Exists(tempPath))
                {
                    try
                    {
                        File.Delete(tempPath);
                    }
                    catch
                    {
                        // Игнорируем ошибку удаления временного файла
                    }
                }
            }
        }

        /// <summary>
        /// Выполняет замену плейсхолдеров в элементе документа (тело, колонтитул и т.д.).
        /// Учитывает ситуацию, когда текст плейсхолдера разбит на несколько Run элементов.
        /// </summary>
        private static void ReplacePlaceholdersInElement(OpenXmlElement element, Dictionary<string, string> data)
        {
            // Регулярное выражение для поиска плейсхолдеров {{ИмяПоля}}
            var placeholderRegex = new Regex(@"\{\{(.+?)\}\}", RegexOptions.Compiled);

            // Получаем все параграфы в элементе
            var paragraphs = element.Descendants<Paragraph>().ToList();

            foreach (var paragraph in paragraphs)
            {
                ProcessParagraph(paragraph, data, placeholderRegex);
            }

            // Также обрабатываем таблицы (ячейки могут содержать текст напрямую)
            var tables = element.Descendants<Table>().ToList();
            foreach (var table in tables)
            {
                var tableParagraphs = table.Descendants<Paragraph>().ToList();
                foreach (var paragraph in tableParagraphs)
                {
                    ProcessParagraph(paragraph, data, placeholderRegex);
                }
            }
        }

        /// <summary>
        /// Обрабатывает отдельный параграф, объединяя разбитые плейсхолдеры и выполняя замену.
        /// </summary>
        private static void ProcessParagraph(Paragraph paragraph, Dictionary<string, string> data, Regex placeholderRegex)
        {
            // Получаем все элементы Run в параграфе
            var runs = paragraph.Elements<Run>().ToList();
            if (runs.Count == 0)
            {
                return;
            }

            // Объединяем текст из всех Run для поиска плейсхолдеров
            // При этом сохраняем информацию о позициях для корректной замены
            string fullText = string.Concat(runs.Select(r => r.InnerText));

            // Находим все плейсхолдеры в тексте
            var matches = placeholderRegex.Matches(fullText);
            if (matches.Count == 0)
            {
                return;
            }

            // Для каждого найденного плейсхолдера выполняем замену
            // Обработка идёт в обратном порядке, чтобы не сбить позиции при модификации
            var matchesList = matches.Cast<Match>().OrderByDescending(m => m.Index).ToList();

            foreach (var match in matchesList)
            {
                string placeholderName = match.Groups[1].Value;
                
                if (data.TryGetValue(placeholderName, out string? replacementValue))
                {
                    // Плейсхолдер найден в словаре — выполняем замену
                    ReplacePlaceholderInRuns(runs, match.Index, match.Length, replacementValue ?? string.Empty);
                }
                // Если значения нет в словаре, оставляем плейсхолдер без изменений
            }
        }

        /// <summary>
        /// Заменяет текст плейсхолдера в последовательности Run элементов.
        /// </summary>
        /// <param name="runs">Список Run элементов параграфа.</param>
        /// <param name="startIndex">Начальная позиция плейсхолдера в объединённом тексте.</param>
        /// <param name="length">Длина плейсхолдера.</param>
        /// <param name="replacementValue">Значение для подстановки.</param>
        private static void ReplacePlaceholderInRuns(List<Run> runs, int startIndex, int length, string replacementValue)
        {
            // Определяем, какие Run участвуют в плейсхолдере
            int currentPosition = 0;
            int runStartIndex = -1;
            int runEndIndex = -1;
            int startOffsetInRun = 0;
            int endOffsetInRun = 0;

            for (int i = 0; i < runs.Count; i++)
            {
                var run = runs[i];
                var textElement = run.Elements<Text>().FirstOrDefault();
                if (textElement == null)
                {
                    continue;
                }

                string runText = textElement.Text ?? string.Empty;
                int runLength = runText.Length;
                int runEndPosition = currentPosition + runLength;

                // Проверяем, попадает ли начало плейсхолдера в этот Run
                if (runStartIndex == -1 && startIndex >= currentPosition && startIndex < runEndPosition)
                {
                    runStartIndex = i;
                    startOffsetInRun = startIndex - currentPosition;
                }

                // Проверяем, попадает ли конец плейсхолдера в этот Run
                if (startIndex + length <= runEndPosition && runStartIndex != -1)
                {
                    runEndIndex = i;
                    endOffsetInRun = startIndex + length - currentPosition;
                    break;
                }

                currentPosition = runEndPosition;
            }

            if (runStartIndex == -1 || runEndIndex == -1)
            {
                // Не удалось определить диапазон Run для замены
                return;
            }

            // Выполняем замену
            if (runStartIndex == runEndIndex)
            {
                // Плейсхолдер полностью находится в одном Run
                var run = runs[runStartIndex];
                var textElement = run.Elements<Text>().FirstOrDefault();
                if (textElement != null)
                {
                    string originalText = textElement.Text ?? string.Empty;
                    string newText = originalText.Substring(0, startOffsetInRun) + 
                                     replacementValue + 
                                     originalText.Substring(endOffsetInRun);
                    textElement.Text = newText;
                }
            }
            else
            {
                // Плейсхолдер разбит на несколько Run
                // Модифицируем первый Run: оставляем текст до плейсхолдера + заменяющее значение
                var firstRun = runs[runStartIndex];
                var firstTextElement = firstRun.Elements<Text>().FirstOrDefault();
                if (firstTextElement != null)
                {
                    string originalText = firstTextElement.Text ?? string.Empty;
                    firstTextElement.Text = originalText.Substring(0, startOffsetInRun) + replacementValue;
                }

                // Удаляем промежуточные Run (между start и end)
                for (int i = runStartIndex + 1; i < runEndIndex; i++)
                {
                    runs[i].Remove();
                }

                // Модифицируем последний Run: оставляем текст после плейсхолдера
                if (runEndIndex > runStartIndex)
                {
                    // После удаления промежуточных, индекс end смещается
                    int newEndIndex = runStartIndex + 1;
                    if (newEndIndex < runs.Count)
                    {
                        var lastRun = runs[newEndIndex];
                        var lastTextElement = lastRun.Elements<Text>().FirstOrDefault();
                        if (lastTextElement != null)
                        {
                            string originalText = lastTextElement.Text ?? string.Empty;
                            if (endOffsetInRun < originalText.Length)
                            {
                                lastTextElement.Text = originalText.Substring(endOffsetInRun);
                            }
                            else
                            {
                                // Весь текст в этом Run был частью плейсхолдера — удаляем Run
                                lastRun.Remove();
                            }
                        }
                    }
                }
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Пример использования
            var data = new Dictionary<string, string>
            {
                { "InvNumber", "ИНВ-001" },
                { "Date", "20.05.2024" },
                { "ChangeType", "Изменение №1" },
                { "ChangeSheets", "2" },
                { "ChangeNumber", "ИЗМ-001" },
                { "Developer", "Иванов И.И." },
                { "Checker", "Петров П.П." },
                { "OTK", "Сидоров С.С." },
                { "NormControl", "Кузнецов К.К." },
                { "ChiefTB", "Смирнов А.А." },
                { "ChiefTech", "Главный Технолог В.В." },
                { "Organization", "ООО «ТехноПром»" },
                { "ProductCode", "АБВГ.123456.001" },
                { "DocNumber", "ТП.123456.001" },
                { "CurrentSheet", "1" },
                { "ProductName", "Изделие опытное" },
                { "Applicability", "Для всех модификаций" },
                { "TotalSheets", "5" },
                { "MachineName", "Станок токарный ЧПУ" },
                { "ProcessName", "Процесс механической обработки" }
            };

            try
            {
                // Проверка наличия шаблона
                string templatePath = "TL_Template.docx";
                string outputPath = "Title_Filled.docx";

                if (!File.Exists(templatePath))
                {
                    Console.WriteLine($"Шаблон '{templatePath}' не найден.");
                    Console.WriteLine("Создайте файл шаблона с плейсхолдерами формата {{ИмяПоля}}");
                    Console.WriteLine();
                    Console.WriteLine("Список поддерживаемых плейсхолдеров:");
                    Console.WriteLine("  {{InvNumber}}, {{Date}}, {{ChangeType}}, {{ChangeSheets}},");
                    Console.WriteLine("  {{ChangeNumber}}, {{Developer}}, {{Checker}}, {{OTK}},");
                    Console.WriteLine("  {{NormControl}}, {{ChiefTB}}, {{ChiefTech}}, {{Organization}},");
                    Console.WriteLine("  {{ProductCode}}, {{DocNumber}}, {{CurrentSheet}}, {{ProductName}},");
                    Console.WriteLine("  {{Applicability}}, {{TotalSheets}}, {{MachineName}}, {{ProcessName}}");
                    return;
                }

                TitlePageFiller.FillTitlePage(templatePath, outputPath, data);
                Console.WriteLine($"Документ успешно заполнен и сохранён как '{outputPath}'");
            }
            catch (FileNotFoundException ex)
            {
                Console.WriteLine($"Ошибка: {ex.Message}");
            }
            catch (UnauthorizedAccessException ex)
            {
                Console.WriteLine($"Ошибка доступа: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex.Message}");
            }
        }
    }
}
