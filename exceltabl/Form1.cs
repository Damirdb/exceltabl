using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;
using System.IO;

namespace exceltabl
{
    public partial class Form1 : Form
    {
        private List<Discipline> disciplines = new List<Discipline>();
        private Solution solution = null;
        private Dictionary<int, double> targetSums = new Dictionary<int, double>
        {
            { 1, 60.5 }, // Для 1+2 семестра
            { 3, 59.5 }, // Для 3+4 семестра
            { 5, 60.0 }, // Для 5+6 семестра
            { 7, 60.0 }  // Для 7+8 семестра
        };

        private readonly HashSet<string> excludedDisciplines = new HashSet<string>
        {
            "Онтологическое моделирование",
            "Проектирование пользовательского интерфейса"
        };
        private DateTime optimizationStartTime;
        private TimeSpan optimizationTimeLimit = TimeSpan.FromMinutes(5);
        private const double StepSize = 1;
        private const double DifferenceThreshold = 1.0; // Порог разницы для оптимизации
        // Допуски для приближённых решений
        const double BlockTolerance = 0.01; // допустимое отклонение по блоку
        const double GlobalTolerance = 2.0; // допустимое отклонение по общей сумме

        // В классе Form1 добавим поле для хранения найденных вариантов
        private List<Dictionary<string, double>> foundVariants = new List<Dictionary<string, double>>();

        private List<double> objectiveValues = new List<double>();

        public Form1()
        {
            InitializeComponent();
            InitializeDataGridView();
        }

        private void InitializeDataGridView()
        {
            dgvResult.AutoGenerateColumns = false;
            dgvResult.Columns.Clear();

            dgvResult.Columns.Add(new DataGridViewTextBoxColumn()
            {
                DataPropertyName = "Название_дисциплины",
                HeaderText = "Название дисциплины",
                Width = 200
            });

            dgvResult.Columns.Add(new DataGridViewTextBoxColumn()
            {
                DataPropertyName = "min_трудоемкость",
                HeaderText = "min трудоемкость",
                Width = 80
            });

            dgvResult.Columns.Add(new DataGridViewTextBoxColumn()
            {
                DataPropertyName = "max_трудоемкость",
                HeaderText = "max трудоемкость",
                Width = 80
            });

            dgvResult.Columns.Add(new DataGridViewTextBoxColumn()
            {
                DataPropertyName = "коэф_значимости",
                HeaderText = "коэф. значимости",
                Width = 100
            });

            dgvResult.Columns.Add(new DataGridViewTextBoxColumn()
            {
                DataPropertyName = "№_семестра",
                HeaderText = "№ семестра",
                Width = 150
            });


        }

        private void btnOpenFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog openDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx;*.xls"
            };

            if (openDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    string filePath = openDialog.FileName;
                    lblFilePath.Text = filePath;
                    disciplines = LoadDataFromExcel(filePath);
                    CheckForDisciplineDuplicates();

                    dgvResult.DataSource = disciplines.Select(d => new
                    {
                        Название_дисциплины = d.Name,
                        min_трудоемкость = d.MinValue,
                        max_трудоемкость = d.MaxValue,
                        коэф_значимости = d.Coefficient,
                        N_семестра = string.Join(", ", d.Semesters)
                    }).ToList();

                    MessageBox.Show($"Загружено {disciplines.Count} дисциплин", "Успех",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка загрузки файла: {ex.Message}", "Ошибка",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private List<Discipline> LoadDataFromExcel(string filePath)
        {
            var result = new List<Discipline>();

            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheets.First();
                int row = 2; // Первая строка — заголовки

                while (!string.IsNullOrWhiteSpace(worksheet.Cell(row, 1).GetString()))
                {
                    try
                    {
                        var name = worksheet.Cell(row, 1).GetString();
                        double min = TryGetDouble(worksheet.Cell(row, 2));
                        double max = TryGetDouble(worksheet.Cell(row, 3));
                        double coeff = TryGetDouble(worksheet.Cell(row, 4));
                        var semestersText = worksheet.Cell(row, 5).GetString();
                        var semesters = semestersText.Split(',')
                            .Select(s => int.Parse(s.Trim()))
                            .ToArray();

                        result.Add(new Discipline
                        {
                            Id = Guid.NewGuid().ToString(),
                            Name = name,
                            MinValue = min,
                            MaxValue = max,
                            Coefficient = coeff,
                            Semesters = semesters
                        });
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка в строке {row}: {ex.Message}", "Ошибка данных",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }

                    row++;
                }
            }

            return result;
        }

        private double TryGetDouble(IXLCell cell)
        {
            if (cell.DataType == XLDataType.Number)
                return cell.GetDouble();
            if (cell.DataType == XLDataType.Text)
            {
                if (double.TryParse(cell.GetString().Replace(',', '.'), System.Globalization.NumberStyles.Any,
                                    System.Globalization.CultureInfo.InvariantCulture, out double result))
                    return result;
            }
            throw new InvalidCastException("Значение ячейки не является числом");
        }

        private async void btnRun_Click(object sender, EventArgs e)
        {
            if (disciplines == null || disciplines.Count == 0)
            {
                MessageBox.Show("Сначала загрузите файл с данными", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                ToggleUIState(false, "Выполняется перебор");
                progressBar.Style = ProgressBarStyle.Marquee;
                progressBar.MarqueeAnimationSpeed = 30;

                var blockKeys = new[] { 1, 3, 5, 7 };
                var blockResultsList = new List<List<Dictionary<string, double>>>();
                bool blockFail = false;
                bool memoryError = false;

                await Task.Run(() =>
                {
                    try
                    {
                        foreach (int semKey in blockKeys)
                        {
                            double blockTarget = targetSums[semKey]; // 60.5, 59.5, 60, 60
                            var blockDisciplines = disciplines
                                .Where(d => d.Semesters.Any(s => s == semKey || s == semKey + 1))
                                .ToList();

                            // Исключённые дисциплины сразу добавляем с MinValue
                            var path = new Dictionary<string, double>();
                            foreach (var d in blockDisciplines.Where(d => excludedDisciplines.Contains(d.Name)))
                                path[d.Id] = d.MinValue;

                            // Оставшиеся дисциплины для подбора блока
                            var blockDisciplinesForBlock = blockDisciplines.Where(d => !excludedDisciplines.Contains(d.Name)).ToList();

                            // Сортируем дисциплины по разнице между max и min значениями
                            var sortedDisciplines = blockDisciplinesForBlock
                                .OrderByDescending(d => d.MaxValue - d.MinValue)
                                .ToList();

                            // Берем половину дисциплин для гибкой группы
                            int flexibleCount = sortedDisciplines.Count / 2;
                            var flexible = sortedDisciplines.Take(flexibleCount).ToList();
                            var fixedDiscs = sortedDisciplines.Skip(flexibleCount).ToList();

                            foreach (var d in fixedDiscs)
                                path[d.Id] = d.MinValue;
                            // fixedSum только по не-исключённым дисциплинам!
                            double fixedSum = fixedDiscs.Sum(d => d.MinValue);

                            var blockResult = new List<Dictionary<string, double>>();
                            GenerateCombinationsApproximate(flexible, blockTarget, fixedSum, 0, path, blockResult);

                            if (blockResult.Count == 0)
                            {
                                blockFail = true;
                                break;
                            }

                            blockResultsList.Add(blockResult);
                        }
                    }
                    catch (OutOfMemoryException)
                    {
                        memoryError = true;
                    }
                });

                if (memoryError)
                {
                    MessageBox.Show($"Перебор остановлен из-за нехватки памяти. Найдено {foundVariants?.Count ?? 0} вариантов. Для сохранения используйте кнопку 'Сохранить'.", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    progressBar.Style = ProgressBarStyle.Blocks;
                    ToggleUIState(true, "Готово");
                    return;
                }

                if (blockFail)
                {
                    MessageBox.Show("Не найдено решений хотя бы для одного из блоков.", "Результат", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    foundVariants = new List<Dictionary<string, double>>();
                    return;
                }

                // Генерируем все возможные комбинации
                int blocks = blockResultsList.Count;
                int[] counts = blockResultsList.Select(l => l.Count).ToArray();
                const int MaxFinalVariants = 1000; // Ограничение на количество итоговых вариантов

                foundVariants = new List<Dictionary<string, double>>();
                try
                {
                    void GenerateAndProcess(int[] indices, int depth)
                    {
                        if (foundVariants.Count >= MaxFinalVariants)
                            return;
                        if (depth == blocks)
                        {
                            var finalVariant = new Dictionary<string, double>();
                            for (int b = 0; b < blocks; b++)
                            {
                                var block = blockResultsList[b][indices[b]];
                                foreach (var kv in block)
                                {
                                    if (!finalVariant.ContainsKey(kv.Key))
                                        finalVariant[kv.Key] = 0;
                                    finalVariant[kv.Key] += kv.Value;
                                }
                            }
                            if (!foundVariants.Any(existing => AreDictionariesEqual(existing, finalVariant)))
                                foundVariants.Add(finalVariant);
                            return;
                        }
                        for (int i = 0; i < counts[depth]; i++)
                        {
                            indices[depth] = i;
                            GenerateAndProcess(indices, depth + 1);
                            if (foundVariants.Count >= MaxFinalVariants)
                                break;
                        }
                    }

                    GenerateAndProcess(new int[blocks], 0);

                    // Вычисляем целевую функцию для каждого варианта
                    objectiveValues = new List<double>();
                    foreach (var variant in foundVariants)
                    {
                        double product = 1.0;
                        for (int sem = 1; sem <= 8; sem++)
                        {
                            double sum = 0.0;
                            foreach (var kv in variant)
                            {
                                var disc = disciplines.First(d => d.Id == kv.Key);
                                if (disc.Semesters.Contains(sem))
                                {
                                    sum += disc.Coefficient * kv.Value;
                                }
                            }
                            product *= sum;
                        }
                        objectiveValues.Add(product);
                    }

                    // Сортируем варианты по значению целевой функции
                    var sortedVariants = foundVariants
                        .Zip(objectiveValues, (v, f) => new { Variant = v, Objective = f })
                        .OrderByDescending(x => x.Objective)
                        .ToList();

                    // Обновляем foundVariants отсортированными вариантами
                    foundVariants = sortedVariants.Select(x => x.Variant).ToList();

                    var allDisciplineIds = disciplines.Select(d => d.Id).ToList();
                    var table = allDisciplineIds.Select(id =>
                    {
                        var disc = disciplines.First(d => d.Id == id);
                        var row = new Dictionary<string, object>
                        {
                            ["Название_дисциплины"] = disc.Name,
                            ["min_трудоемкость"] = disc.MinValue,
                            ["max_трудоемкость"] = disc.MaxValue,
                            ["коэф_значимости"] = disc.Coefficient,
                            ["N_семестра"] = string.Join(", ", disc.Semesters)
                        };
                        for (int i = 0; i < foundVariants.Count; i++)
                        {
                            double value = excludedDisciplines.Contains(disc.Name)
                                ? disc.MinValue
                                : foundVariants[i].TryGetValue(id, out var val) ? val : disc.MinValue;
                            row[$"Fц={sortedVariants[i].Objective:F2}"] = value;
                        }
                        return row;
                    }).ToList();

                    // Добавляем строку с суммой
                    var sumRow = new Dictionary<string, object>
                    {
                        ["Название_дисциплины"] = "Сумма",
                        ["min_трудоемкость"] = 0.0,
                        ["max_трудоемкость"] = 0.0,
                        ["коэф_значимости"] = 0.0,
                        ["N_семестра"] = ""
                    };

                    // Суммы по семестрам
                    for (int sem = 1; sem <= 8; sem++)
                    {
                        sumRow[$"Семестр_{sem}"] = 0.0;
                    }

                    // Итоговые суммы по вариантам (по всем дисциплинам)
                    for (int i = 0; i < foundVariants.Count; i++)
                    {
                        double sum = disciplines
                            .Sum(d => foundVariants[i].TryGetValue(d.Id, out var val) ? val : d.MinValue);
                        sumRow[$"Fц={sortedVariants[i].Objective:F2}"] = sum;
                    }
                    table.Add(sumRow);

                    dgvResult.DataSource = table.Select(x => x.ToDictionary(
                        kv => kv.Key,
                        kv => kv.Value is double d ? d.ToString("F2") : kv.Value.ToString()
                    )).ToList();

                    // После формирования foundVariants и перед выводом/сохранением
                    // Проверим сумму MinValue по excludedDisciplines и выведем предупреждение, если итог не 244
                    double excludedSumCheck = disciplines
                        .Where(d => excludedDisciplines.Contains(d.Name))
                        .Sum(d => d.MinValue);
                    double blockTargetSum = targetSums.Values.Sum(); // 240
                    double expectedTotal = blockTargetSum + excludedSumCheck;
                    if (Math.Abs(expectedTotal - 244.0) > 0.01)
                    {
                        MessageBox.Show($"Внимание! Сумма по блокам: {blockTargetSum}, сумма по исключённым дисциплинам: {excludedSumCheck}, итоговая сумма: {expectedTotal}.\nОжидается 244. Проверьте targetSums и MinValue исключённых дисциплин.", "Контроль суммы", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }

                    MessageBox.Show($"Найдено {foundVariants.Count} уникальных вариантов. Для сохранения используйте кнопку 'Сохранить'.", "Готово", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    // Добавляем исключённые дисциплины с MinValue в каждый вариант
                    foreach (var variant in foundVariants)
                    {
                        foreach (var d in disciplines.Where(d => excludedDisciplines.Contains(d.Name)))
                        {
                            variant[d.Id] = d.MinValue;
                        }
                    }

                    if (foundVariants.Count > 0)
                    {
                        var variant = foundVariants[0];
                        double block1 = disciplines.Where(d => !excludedDisciplines.Contains(d.Name) && (d.Semesters.Contains(1) || d.Semesters.Contains(2)))
                            .Sum(d => variant.TryGetValue(d.Id, out var val) ? val : d.MinValue);
                        double block2 = disciplines.Where(d => !excludedDisciplines.Contains(d.Name) && (d.Semesters.Contains(3) || d.Semesters.Contains(4)))
                            .Sum(d => variant.TryGetValue(d.Id, out var val) ? val : d.MinValue);
                        double block3 = disciplines.Where(d => !excludedDisciplines.Contains(d.Name) && (d.Semesters.Contains(5) || d.Semesters.Contains(6)))
                            .Sum(d => variant.TryGetValue(d.Id, out var val) ? val : d.MinValue);
                        double block4 = disciplines.Where(d => !excludedDisciplines.Contains(d.Name) && (d.Semesters.Contains(7) || d.Semesters.Contains(8)))
                            .Sum(d => variant.TryGetValue(d.Id, out var val) ? val : d.MinValue);
                        double excludedSum = disciplines.Where(d => excludedDisciplines.Contains(d.Name))
                            .Sum(d => variant.TryGetValue(d.Id, out var val) ? val : d.MinValue);
                        double total = block1 + block2 + block3 + block4 + excludedSum;
                        MessageBox.Show($"Диагностика суммы для первого варианта:\n" +
                            $"Блок 1+2: {block1}\nБлок 3+4: {block2}\nБлок 5+6: {block3}\nБлок 7+8: {block4}\n" +
                            $"Excluded: {excludedSum}\nИтого: {total}", "Диагностика суммы");
                    }
                }
                catch (OutOfMemoryException)
                {
                    MessageBox.Show($"Перебор остановлен из-за нехватки памяти на этапе генерации комбинаций. Найдено {foundVariants?.Count ?? 0} вариантов. Для сохранения используйте кнопку 'Сохранить'.", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    progressBar.Style = ProgressBarStyle.Blocks;
                    ToggleUIState(true, "Готово");
                    return;
                }
            }
            finally
            {
                progressBar.Style = ProgressBarStyle.Blocks;
                ToggleUIState(true, "Готово");
            }
        }

        private void ToggleUIState(bool enable, string statusText)
        {
            progressBar.Visible = !enable;
            btnRun.Enabled = enable;
            lblStatus.Text = statusText;
            Application.DoEvents();
        }

        private void ShowOptimizationSummary(Solution solution)
        {
            if (!solution.Combinations.Any())
                return;

            var differences = solution.Combinations[0].Differences;
            var diffMessage = string.Join("\n", differences.Select(d =>
                $"Семестры {d.Key}-{d.Key + 1}: разница {d.Value:F2} ч."));

            MessageBox.Show(
                $"Оптимизация завершена!\n" +
                $"Целевая функция: {solution.Value:F2}\n\n{diffMessage}",
                "Результат", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void CheckForDisciplineDuplicates()
        {
            var duplicateGroups = disciplines
                .GroupBy(d => d.Name)
                .Where(g => g.Count() > 1)
                .ToList();
            foreach (var group in duplicateGroups)
            {
                // Проверяем, есть ли среди "дубликатов" различие по семестрам или другим параметрам
                var uniqueVariants = group.Select(d => $"{string.Join(",", d.Semesters)}|{d.MinValue}|{d.MaxValue}|{d.Coefficient}").Distinct().Count();
                if (uniqueVariants > 1)
                {
                    // Это не дубликаты, а разные дисциплины с одинаковым названием
                    // Можно вывести предупреждение в лог или для разработчика
                    System.Diagnostics.Debug.WriteLine($"Внимание: дисциплина '{group.Key}' встречается несколько раз с разными параметрами. Это ОК!");
                }
            }
        }

        // Чистый рекурсивный перебор по всем дисциплинам от min до max с шагом 0.5


        private void StrictRecursiveSearch(List<Discipline> disciplines, int index, Dictionary<string, double> current, List<Dictionary<string, double>> results)
        {
            if (index == disciplines.Count)
            {
                if (IsValidCombination(current))
                    results.Add(new Dictionary<string, double>(current));
                return;
            }
            var disc = disciplines[index];
            for (double val = disc.MinValue; val <= disc.MaxValue + 0.0001; val += 0.25)
            {
                current[disc.Id] = val;
                StrictRecursiveSearch(disciplines, index + 1, current, results);
                current.Remove(disc.Id);
            }
        }

        // Проверка полной комбинации
        private bool IsValidCombination(Dictionary<string, double> combination)
        {
            // 1. Проверка суммы по парам семестров (без исключённых дисциплин)
            var targetSums = new Dictionary<int, double> { { 1, 60.5 }, { 3, 59.5 }, { 5, 60.0 }, { 7, 60.0 } };
            foreach (var target in targetSums)
            {
                double sum = 0;
                foreach (var kv in combination)
                {
                    var disc = disciplines.First(d => d.Id == kv.Key);
                    if (excludedDisciplines.Contains(disc.Name)) continue;
                    if (disc.Semesters.Contains(target.Key) || disc.Semesters.Contains(target.Key + 1))
                        sum += kv.Value;
                }
                if (Math.Abs(sum - target.Value) > 0.01) return false;
            }
            // 2. Проверка разницы по семестрам (все дисциплины)
            var semesterLoad = new Dictionary<int, double>();
            foreach (var kv in combination)
            {
                var disc = disciplines.First(d => d.Id == kv.Key);
                foreach (var sem in disc.Semesters)
                {
                    if (!semesterLoad.ContainsKey(sem)) semesterLoad[sem] = 0;
                    semesterLoad[sem] += kv.Value / disc.Semesters.Length;
                }
            }
            var semesters = semesterLoad.Keys.OrderBy(x => x).ToList();
            for (int i = 0; i < semesters.Count - 1; i += 2)
            {
                int s1 = semesters[i];
                int s2 = semesters[i + 1];
                if (Math.Abs(semesterLoad[s1] - semesterLoad[s2]) > 6.0) return false;
            }
            // 3. Общая сумма (только по дисциплинам без исключений)
            double total = disciplines
                .Where(d => !excludedDisciplines.Contains(d.Name))
                .Sum(d => combination.TryGetValue(d.Id, out var val) ? val : d.MinValue);
            if (Math.Abs(total - 240.0) > 0.01) return false;
            return true;
        }
        // Сохранение одной комбинации в Excel
        private void SaveCombinationToExcel(Dictionary<string, double> combination, string fileName, int variantIndex = 1, int totalVariants = 1)
        {
            string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string path = Path.Combine(desktop, fileName);
            using (var workbook = new ClosedXML.Excel.XLWorkbook())
            {
                var sheet = workbook.Worksheets.Add("Вариант");
                sheet.Cell(1, 1).Value = "Название дисциплины";
                sheet.Cell(1, 2).Value = "min трудоемкость";
                sheet.Cell(1, 3).Value = "max трудоемкость";
                sheet.Cell(1, 4).Value = "коэф. значимости";
                sheet.Cell(1, 5).Value = "№ семестра";
                // Если вариантов больше одного, делаем столбцы Трудоемкость 1, 2, ...
                for (int i = 0; i < totalVariants; i++)
                {
                    sheet.Cell(1, 6 + i).Value = $"Трудоемкость {i + 1}";
                }
                int row = 2;
                double sum = 0;
                foreach (var kv in combination)
                {
                    var disc = disciplines.First(d => d.Id == kv.Key);
                    sheet.Cell(row, 1).Value = disc.Name;
                    sheet.Cell(row, 2).Value = disc.MinValue;
                    sheet.Cell(row, 3).Value = disc.MaxValue;
                    sheet.Cell(row, 4).Value = disc.Coefficient;
                    sheet.Cell(row, 5).Value = string.Join(", ", disc.Semesters);
                    // Заполняем только нужный столбец для этого варианта
                    double value = excludedDisciplines.Contains(disc.Name)
                        ? disc.MinValue
                        : (kv.Value);
                    sheet.Cell(row, 6 + variantIndex - 1).Value = value;
                    sum += value;
                    row++;
                }
                // Добавляем строку с суммой
                sheet.Cell(row, 1).Value = "Сумма";
                sheet.Cell(row, 6 + variantIndex - 1).Value = sum;
                sheet.Columns().AdjustToContents();
                workbook.SaveAs(path);
            }
        }

        // Генерация приближённых комбинаций для блока
        private void GenerateCombinationsApproximate(
            List<Discipline> arr,
            double target_sum,
            double curr_sum,
            int index,
            Dictionary<string, double> path,
            List<Dictionary<string, double>> result)
        {
            if (curr_sum > target_sum + BlockTolerance)
                return;

            if (index == arr.Count)
            {
                if (Math.Abs(curr_sum - target_sum) <= BlockTolerance)
                {
                    result.Add(new Dictionary<string, double>(path));
                }
                return;
            }

            var disc = arr[index];
            var valuesToTry = new List<double> { disc.MinValue, disc.MaxValue };
            for (double val = disc.MinValue + 1.0; val < disc.MaxValue; val += 1.0)
            {
                if (!valuesToTry.Contains(val))
                    valuesToTry.Add(val);
            }
            valuesToTry.Sort();

            foreach (double val in valuesToTry)
            {
                if (curr_sum + val > target_sum + BlockTolerance)
                    continue;

                path[disc.Id] = val;
                GenerateCombinationsApproximate(arr, target_sum, curr_sum + val, index + 1, path, result);
                path.Remove(disc.Id);
            }
        }

        // Итоговая проверка с допуском
        private bool IsApproximateCombination(Dictionary<string, double> combination)
        {
            // Все проверки убраны — всегда возвращаем true
            return true;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (foundVariants == null || foundVariants.Count == 0)
            {
                MessageBox.Show("Сначала выполните поиск комбинаций", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            SaveFileDialog saveDialog = new SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                Title = "Сохранить результаты как Excel файл"
            };

            if (saveDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    SaveAllCombinationsToExcel(foundVariants, objectiveValues, saveDialog.FileName);
                    MessageBox.Show("Результаты успешно сохранены", "Успех",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка сохранения файла: {ex.Message}", "Ошибка",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // Инициализация формы при загрузке
        }

        private void SaveAllCombinationsToExcel(List<Dictionary<string, double>> variants, List<double> objectiveValues, string fileName)
        {
            string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string path = Path.Combine(desktop, fileName);
            using (var workbook = new ClosedXML.Excel.XLWorkbook())
            {
                var sheet = workbook.Worksheets.Add("Варианты");
                sheet.Cell(1, 1).Value = "Название дисциплины";
                sheet.Cell(1, 2).Value = "min трудоемкость";
                sheet.Cell(1, 3).Value = "max трудоемкость";
                sheet.Cell(1, 4).Value = "коэф. значимости";
                sheet.Cell(1, 5).Value = "№ семестра";
                for (int i = 0; i < variants.Count; i++)
                {
                    sheet.Cell(1, 6 + i).Value = $"Fц={objectiveValues[i]:F2}";
                }
                int row = 2;
                // Сначала обычные дисциплины
                foreach (var disc in disciplines)
                {
                    sheet.Cell(row, 1).Value = disc.Name;
                    sheet.Cell(row, 2).Value = disc.MinValue;
                    sheet.Cell(row, 3).Value = disc.MaxValue;
                    sheet.Cell(row, 4).Value = disc.Coefficient;
                    sheet.Cell(row, 5).Value = string.Join(", ", disc.Semesters);
                    for (int i = 0; i < variants.Count; i++)
                    {
                        double value = excludedDisciplines.Contains(disc.Name)
                            ? disc.MinValue
                            : variants[i].TryGetValue(disc.Id, out var val) ? val : disc.MinValue;
                        sheet.Cell(row, 6 + i).Value = value;
                    }
                    row++;
                }
                // Строка с суммой по всем дисциплинам
                sheet.Cell(row, 1).Value = "Сумма";
                for (int i = 0; i < variants.Count; i++)
                {
                    double sum = disciplines
                        .Sum(d => variants[i].TryGetValue(d.Id, out var val) ? val : d.MinValue);
                    sheet.Cell(row, 6 + i).Value = sum;
                }
                row++;
                sheet.Columns().AdjustToContents();
                workbook.SaveAs(path);
            }
        }

        private bool AreDictionariesEqual(Dictionary<string, double> a, Dictionary<string, double> b)
        {
            if (a.Count != b.Count) return false;
            foreach (var kv in a)
            {
                if (!b.TryGetValue(kv.Key, out var val) || Math.Abs(val - kv.Value) > 0.0001)
                    return false;
            }
            return true;
        }

        private void button_max_Click(object sender, EventArgs e)
        {
            if (foundVariants == null || foundVariants.Count == 0)
            {
                MessageBox.Show("Сначала выполните поиск комбинаций", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Запросить у пользователя файл для выделения
            OpenFileDialog openDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                Title = "Выберите файл с вариантами для выделения максимального"
            };

            if (openDialog.ShowDialog() != DialogResult.OK)
                return;

            // Открываем файл и выделяем нужный столбец
            using (var workbook = new ClosedXML.Excel.XLWorkbook(openDialog.FileName))
            {
                var sheet = workbook.Worksheet(1);
                int maxCol = -1;
                double maxF = double.MinValue;
                int lastCol = sheet.LastColumnUsed().ColumnNumber();
                // Ищем столбец с максимальным значением Fц в заголовке
                for (int col = 6; col <= lastCol; col++)
                {
                    var header = sheet.Cell(1, col).GetString();
                    if (header.StartsWith("Fц="))
                    {
                        if (double.TryParse(header.Substring(3), out double f) && f > maxF)
                        {
                            maxF = f;
                            maxCol = col;
                        }
                    }
                }
                if (maxCol > 0)
                {
                    var headerCell = sheet.Cell(1, maxCol);
                    headerCell.Style.Fill.BackgroundColor = XLColor.Yellow;
                    workbook.Save();
                    MessageBox.Show($"Максимальное значение Fц = {maxF:F2} (столбец {headerCell.Address.ColumnLetter}). Заголовок выделен цветом.", "Результат", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Не удалось найти столбец с максимальным значением Fц.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }
    }

    public class Discipline
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public double MinValue { get; set; }
        public double MaxValue { get; set; }
        public double Coefficient { get; set; }
        public int[] Semesters { get; set; }
    }

    public class WorkloadCombination
    {
        public Dictionary<string, double> Values { get; set; } = new Dictionary<string, double>();
        public double TotalValue { get; set; }
        public Dictionary<int, double> SemesterTotals { get; set; } = new Dictionary<int, double>();
        public Dictionary<int, double> Differences { get; set; } = new Dictionary<int, double>();

        public WorkloadCombination Clone()
        {
            return new WorkloadCombination
            {
                Values = new Dictionary<string, double>(Values),
                TotalValue = TotalValue,
                SemesterTotals = new Dictionary<int, double>(SemesterTotals),
                Differences = new Dictionary<int, double>(Differences)
            };
        }
    }

    public class Solution
    {
        public List<WorkloadCombination> Combinations { get; set; } = new List<WorkloadCombination>();
        public double Value { get; set; }

        public Solution Clone()
        {
            return new Solution
            {
                Combinations = Combinations.Select(c => c.Clone()).ToList(),
                Value = Value
            };
        }
    }
} /// в выводе 1 максимальный варианта, вывод всех не удалять, закоментировать
