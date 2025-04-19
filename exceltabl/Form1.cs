using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;

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
        private const double StepSize = 0.5;
        private const double DifferenceThreshold = 1.0; // Порог разницы для оптимизации

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
                DataPropertyName = "Name",
                HeaderText = "Дисциплина",
                Width = 200
            });

            dgvResult.Columns.Add(new DataGridViewTextBoxColumn()
            {
                DataPropertyName = "Hours",
                HeaderText = "Часы",
                Width = 80
            });

            dgvResult.Columns.Add(new DataGridViewTextBoxColumn()
            {
                DataPropertyName = "Coefficient",
                HeaderText = "Коэффициент",
                Width = 100
            });

            dgvResult.Columns.Add(new DataGridViewTextBoxColumn()
            {
                DataPropertyName = "Semesters",
                HeaderText = "Семестры",
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

                    dgvResult.DataSource = disciplines.Select(d => new
                    {
                        d.Name,
                        Hours = $"{d.MinValue}-{d.MaxValue}",
                        Coefficient = d.Coefficient.ToString("F3"),
                        Semesters = string.Join(", ", d.Semesters)
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
                int row = 2;

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
                ToggleUIState(false, "Выполняется проверка и оптимизация...");

                var optimizedSolution = await Task.Run(OptimizeSolution);

                if (optimizedSolution?.Values?.Count > 0)
                {
                    solution = optimizedSolution;

                    dgvResult.DataSource = solution.Values.Select(kv =>
                    {
                        var disc = disciplines.First(d => d.Name == kv.Key);
                        return new
                        {
                            Дисциплина = kv.Key,
                            Часы = kv.Value.ToString("F2"),
                            Коэффициент = disc.Coefficient.ToString("F3"),
                            Семестры = string.Join(", ", disc.Semesters)
                        };
                    }).ToList();

                    ShowOptimizationSummary(optimizedSolution);
                }
                else
                {
                    MessageBox.Show("Не удалось найти допустимое решение", "Результат",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка оптимизации: {ex.Message}", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
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
            var differences = CalculateDifferences(solution);
            var diffMessage = string.Join("\n", differences.Select(d =>
                $"Семестры {d.Key}-{d.Key + 1}: разница {d.Value:F2} ч."));

            MessageBox.Show(
                $"Оптимизация завершена!\n" +
                $"Целевая функция: {solution.Value:F2}\n\n{diffMessage}",
                "Результат", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private Solution OptimizeSolution()
        {
            optimizationStartTime = DateTime.Now;

            var filteredDisciplines = disciplines
                .Where(d => !excludedDisciplines.Contains(d.Name))
                .ToList();

            // Создаем начальное решение с минимальными значениями
            var initialSolution = new Solution();
            foreach (var disc in filteredDisciplines)
            {
                initialSolution.Values[disc.Name] = disc.MinValue;
            }

            // Проверяем разницу между текущими и целевыми значениями
            var differences = CalculateDifferences(initialSolution);

            // Оптимизируем только если есть превышения (разница > 0)
            bool needsOptimization =
                !CheckSemesterConstraintsInt(initialSolution, 1) ||
                !CheckSemesterConstraintsDouble(initialSolution, DifferenceThreshold);

            if (!needsOptimization)
            {
                initialSolution.Value = CalculateObjectiveFunction(initialSolution);
                return initialSolution;
            }

            // Если есть превышения — пытаемся оптимизировать
            var greedySolution = TryGreedyApproach(filteredDisciplines);
            if (greedySolution != null && CheckSemesterConstraints(greedySolution))
            {
                greedySolution.Value = CalculateObjectiveFunction(greedySolution);
                return greedySolution;
            }

            return RunSimulatedAnnealing(filteredDisciplines);
        }

        private Dictionary<int, double> CalculateDifferences(Solution solution)
        {
            var semesterLoad = new Dictionary<int, double>();

            // Проходим по всем дисциплинам и распределяем часы по семестрам
            foreach (var kvp in solution.Values)
            {
                var discipline = disciplines.FirstOrDefault(d => d.Name == kvp.Key);
                if (discipline == null) continue;

                // Для каждой дисциплины распределяем часы по семестрам
                foreach (int semester in discipline.Semesters)
                {
                    if (!semesterLoad.ContainsKey(semester))
                        semesterLoad[semester] = 0;

                    semesterLoad[semester] += kvp.Value / discipline.Semesters.Length;
                }
            }

            var differences = new Dictionary<int, double>();
            var semesters = semesterLoad.Keys.OrderBy(k => k).ToList();

            // Считаем разницу между нагрузками на соседние семестры
            for (int i = 0; i < semesters.Count - 1; i++)
            {
                int s1 = semesters[i];
                int s2 = semesters[i + 1];
                differences[s1] = Math.Abs(semesterLoad[s1] - semesterLoad[s2]);
            }

            return differences;
        }



        private Dictionary<int, double> CalculateSemesterTotals(Solution solution)
        {
            var semesterGroups = solution.Values
                .SelectMany(kv => disciplines.First(d => d.Name == kv.Key).Semesters
                    .Select(sem => new { Semester = sem, Hours = kv.Value }))
                .GroupBy(x => (x.Semester - 1) / 2 * 2 + 1)
                .ToDictionary(g => g.Key, g => g.Sum(x => x.Hours));

            return semesterGroups;
        }

        private Solution TryGreedyApproach(List<Discipline> filteredDisciplines)
        {
            var solution = new Solution();
            var random = new Random();

            // Начинаем с минимальных значений
            foreach (var disc in filteredDisciplines)
            {
                solution.Values[disc.Name] = disc.MinValue;
            }

            // Распределяем оставшиеся часы
            double remainingHours = CalculateRemainingHours(solution);

            while (remainingHours > 0)
            {
                var available = filteredDisciplines
                    .Where(d => solution.Values[d.Name] < d.MaxValue)
                    .OrderByDescending(d => d.Coefficient)
                    .ToList();

                if (!available.Any()) break;

                double totalCoeff = available.Sum(d => d.Coefficient);
                foreach (var disc in available)
                {
                    double share = disc.Coefficient / totalCoeff;
                    double toAdd = Math.Min(remainingHours * share, disc.MaxValue - solution.Values[disc.Name]);
                    solution.Values[disc.Name] += toAdd;
                    remainingHours -= toAdd;

                    if (remainingHours <= 0) break;
                }
            }

            return CheckSemesterConstraints(solution) ? solution : null;
        }

        private double CalculateRemainingHours(Solution solution)
        {
            double totalAssigned = solution.Values.Sum(kv => kv.Value);
            double totalRequired = targetSums.Sum(ts => ts.Value);
            return totalRequired - totalAssigned;
        }

        private Solution RunSimulatedAnnealing(List<Discipline> filteredDisciplines)
        {
            var current = new Solution();
            foreach (var disc in filteredDisciplines)
            {
                current.Values[disc.Name] = disc.MinValue;
            }
            current.Value = CalculateObjectiveFunction(current);

            var best = current.Clone();
            double temperature = 1000;
            double coolingRate = 0.003;
            var random = new Random();

            while (temperature > 1 && DateTime.Now - optimizationStartTime < optimizationTimeLimit)
            {
                var neighbor = current.Clone();

                int changes = random.Next(1, 5);
                for (int i = 0; i < changes; i++)
                {
                    var disc = filteredDisciplines[random.Next(filteredDisciplines.Count)];
                    double newValue = Math.Max(disc.MinValue,
                        Math.Min(disc.MaxValue,
                            neighbor.Values[disc.Name] + (random.NextDouble() - 0.5) * StepSize * 2));
                    neighbor.Values[disc.Name] = newValue;
                }

                if (CheckSemesterConstraints(neighbor))
                {
                    neighbor.Value = CalculateObjectiveFunction(neighbor);
                    double delta = neighbor.Value - current.Value;

                    if (delta > 0 || random.NextDouble() < Math.Exp(delta / temperature))
                    {
                        current = neighbor;

                        if (current.Value > best.Value)
                        {
                            best = current.Clone();
                        }
                    }
                }

                temperature *= 1 - coolingRate;
            }

            return best.Value > 0 ? best : null;
        }

        private bool CheckSemesterConstraints(Solution solution)
        {
            var semesterTotals = CalculateSemesterTotals(solution);

            foreach (var target in targetSums)
            {
                if (!semesterTotals.ContainsKey(target.Key)) continue;

                double currentSum = semesterTotals[target.Key];
                if (Math.Abs(currentSum - target.Value) > DifferenceThreshold * 2) // Более строгая проверка при оптимизации
                {
                    return false;
                }
            }

            foreach (var kv in solution.Values)
            {
                var disc = disciplines.First(d => d.Name == kv.Key);
                if (kv.Value < disc.MinValue || kv.Value > disc.MaxValue)
                {
                    return false;
                }
            }

            return true;
        }

        private double CalculateObjectiveFunction(Solution solution)
        {
            double totalValue = 0;
            foreach (var item in solution.Values)
            {
                var discipline = disciplines.First(d => d.Name == item.Key);
                totalValue += item.Value * discipline.Coefficient;
            }
            return totalValue;
        }

        private void SaveResultsToExcel(string filePath)
        {
            using (var workbook = new XLWorkbook())
            {
                var sheet = workbook.Worksheets.Add("Результаты");

                // Заголовки
                sheet.Cell(1, 1).Value = "Дисциплина";
                sheet.Cell(1, 2).Value = "Часы";
                sheet.Cell(1, 3).Value = "Коэффициент";
                sheet.Cell(1, 4).Value = "Семестры";
                sheet.Cell(1, 5).Value = "Разница по семестрам";

                var differences = CalculateDifferences(solution);
                int row = 2;

                foreach (var item in solution.Values)
                {
                    var discipline = disciplines.First(d => d.Name == item.Key);
                    double assigned = item.Value;

                    sheet.Cell(row, 1).Value = item.Key;
                    sheet.Cell(row, 2).Value = assigned;
                    sheet.Cell(row, 3).Value = discipline.Coefficient;
                    sheet.Cell(row, 4).Value = string.Join(", ", discipline.Semesters);

                    // Разница по семестрам (внешняя)
                    string diffInfo = string.Join("; ", discipline.Semesters
                        .Select(s => (s - 1) / 2 * 2 + 1)
                        .Distinct()
                        .Select(s =>
                        {
                            differences.TryGetValue(s, out double diff);
                            return $"Сем.{s}-{s + 1}: {diff:F2}";
                        }));

                    sheet.Cell(row, 5).Value = diffInfo;

                    // Форматирование
                    sheet.Cell(row, 2).Style.NumberFormat.Format = "0.00";
                    sheet.Cell(row, 3).Style.NumberFormat.Format = "0.000";

                    // Выделяем зелёным, если есть разница между Min и Max (даже если assigned == MinValue)
                    if (discipline.MaxValue > discipline.MinValue)
                    {
                        var range = sheet.Range(row, 1, row, 5);
                        range.Style.Fill.BackgroundColor = XLColor.LightGreen;
                    }

                    row++;
                }

                sheet.Columns().AdjustToContents();
                workbook.SaveAs(filePath);
            }
        }
        private bool CheckSemesterConstraintsInt(Solution solution, int tolerance = 1)
        {
            var semesterTotals = CalculateSemesterTotals(solution);
            foreach (var target in targetSums)
            {
                if (!semesterTotals.ContainsKey(target.Key)) continue;
                double currentSum = semesterTotals[target.Key];
                if (Math.Abs(currentSum - target.Value) > tolerance)
                    return false;
            }
            return true;
        }

        private bool CheckSemesterConstraintsDouble(Solution solution, double tolerance = 1.0)
        {
            var semesterTotals = CalculateSemesterTotals(solution);
            foreach (var target in targetSums)
            {
                if (!semesterTotals.ContainsKey(target.Key)) continue;
                double currentSum = semesterTotals[target.Key];
                if (Math.Abs(currentSum - target.Value) > tolerance)
                    return false;
            }
            return true;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (solution == null || solution.Values == null || solution.Values.Count == 0)
            {
                MessageBox.Show("Сначала выполните оптимизацию и дождитесь её завершения", "Ошибка",
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
                    SaveResultsToExcel(saveDialog.FileName);
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
    }

    public class Discipline
    {
        public string Name { get; set; }
        public double MinValue { get; set; }
        public double MaxValue { get; set; }
        public double Coefficient { get; set; }
        public int[] Semesters { get; set; }
    }

    public class Solution
    {
        public Dictionary<string, double> Values { get; set; } = new Dictionary<string, double>();
        public double Value { get; set; }

        public Solution Clone()
        {
            return new Solution
            {
                Values = new Dictionary<string, double>(Values),
                Value = Value
            };
        }
    }
}