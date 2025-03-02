using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using System;
using System.IO;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Word.Application;

namespace DocReplacer
{
    public partial class Form1 : Form
    {
        // 初始化窗体组件
        public Form1()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 调用Windows目录选择器获取文件夹路径
        /// </summary>
        /// <returns>用户选择的文件夹路径</returns>
        private string SelectFolder()
        {
            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                // 设置初始目录为桌面
                folderDialog.SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                folderDialog.Description = "请选择目录";

                // 返回用户选择的路径或空字符串
                return folderDialog.ShowDialog() == DialogResult.OK ? folderDialog.SelectedPath : string.Empty;
            }
        }

        // 线程安全的进度更新方法
        private void UpdateProgress(int processed, int total, string currentFile = "")
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => UpdateProgress(processed, total, currentFile)));
                return;
            }

            // 更新进度条
            progressBar1.Maximum = total;
            progressBar1.Value = processed;

            // 可选：更新状态标签
            // labelStatus.Text = $"正在处理：{currentFile} ({processed}/{total})";

            // 强制刷新界面
            //Application.DoEvents();
        }

        #region 按钮点击事件
        // 源目录选择按钮
        private void button1_Click(object sender, EventArgs e)
        {
            textBox1.Text = SelectFolder(); // 将选择的路径显示到源目录文本框
        }

        // 目标目录选择按钮
        private void button2_Click(object sender, EventArgs e)
        {
            textBox2.Text = SelectFolder(); // 将选择的路径显示到目标目录文本框
        }

        // 执行替换按钮

        // 查找按钮点击事件
        private async void button3_Click_1(object sender, EventArgs e)
        {
            try
            {
                button3.Enabled = false;
                listView1.Items.Clear(); // 清空旧数据
                progressBar1.Value = 0; // 重置进度条
                // 验证输入
                if (!ValidateSearchInputs()) return;

                // 获取所有查找规则
                var searchKeywords = textBox3.Lines;

                // 获取源目录所有Word文件
                // 修改后（支持doc和docx）
                var sourceFiles = Directory.GetFiles(textBox1.Text, "*.*")
                    .Where(f => f.EndsWith(".doc", StringComparison.OrdinalIgnoreCase) ||
                                f.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
                    .ToArray();

                // 设置进度条最大值
                UpdateProgress(0, sourceFiles.Length, "准备开始");

                // 异步执行查找
                await System.Threading.Tasks.Task.Run(() => SearchFiles(sourceFiles, searchKeywords));

                MessageBox.Show("查找完成！");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"查找失败：{ex.Message}");
            }
            finally
            {
                button3.Enabled = true;
            }

        }

        private async void button4_Click(object sender, EventArgs e)
        {
            //dynamic wordApp = Activator.CreateInstance(Type.GetTypeFromProgID("Word.Application"));
            try
            {
                button4.Enabled = false; // 禁用按钮防止重复点击
                progressBar1.Value = 0; // 重置进度条


                // 验证输入有效性
                if (!ValidateInputs()) return;

                // 获取所有替换规则（支持多行）
                var replaceRules = ParseReplaceRules();

                // 获取源目录所有Word文件
                var sourceFiles = Directory.GetFiles(textBox1.Text, "*.*")
                    .Where(f => f.EndsWith(".doc", StringComparison.OrdinalIgnoreCase) ||
                    f.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
                    .ToArray();

                // 设置进度条最大值
                UpdateProgress(0, sourceFiles.Length, "准备开始");

                // 异步执行替换操作
                await System.Threading.Tasks.Task.Run(() => ProcessFiles(sourceFiles, textBox2.Text, replaceRules));

                MessageBox.Show("文件替换完成！");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"操作失败：{ex.Message}");
            }
            finally
            {
                button4.Enabled = true; // 重新启用按钮
            }
        }
        #endregion

        /// <summary>
        /// 批量处理文件的核心方法
        /// </summary>
        /// <param name="sourceFiles">源文件路径数组</param>
        /// <param name="targetFolder">目标保存目录</param>
        /// <param name="replaceRules">替换规则集合</param>
        private void ProcessFiles(string[] sourceFiles, string targetFolder, Tuple<string, string>[] replaceRules)
        {
            int totalFiles = sourceFiles.Length;
            int processedCount = 0;

            Application wordApp = null;
            try
            {
                // 创建Word应用实例
                wordApp = new Application
                {
                    Visible = false,
                    ScreenUpdating = false,
                    DisplayAlerts = WdAlertLevel.wdAlertsNone
                };

                foreach (var filePath in sourceFiles)
                {

                    // 生成目标路径
                    var targetPath = Path.Combine(targetFolder, Path.GetFileName(filePath));

                    // 更新进度
                    Interlocked.Increment(ref processedCount);
                    UpdateProgress(processedCount, totalFiles, Path.GetFileName(filePath));

                    Document doc = null;
                    try
                    {

                        // 打开文档
                        doc = wordApp.Documents.Open(filePath, ReadOnly: false, Visible: false);

                        // 应用所有替换规则
                        foreach (var rule in replaceRules)
                        {
                            ReplaceText(doc, rule.Item1, rule.Item2);
                        }

                        // 另存到新位置
                        // 修改后的保存逻辑
                        doc.SaveAs2(
                            FileName: targetPath,
                            FileFormat: Path.GetExtension(targetPath).ToLower() == ".doc" ?
                                WdSaveFormat.wdFormatDocument : WdSaveFormat.wdFormatDocumentDefault
                        );
                        doc.Close(SaveChanges: false);
                    }
                    catch (COMException ex) // 新增的异常捕获
                    {
                        this.Invoke((MethodInvoker)delegate
                        {
                            MessageBox.Show($"无法打开文件 {Path.GetFileName(filePath)}：{ex.Message}");
                        });
                        continue; // 跳过当前文件
                    }
                    finally
                    {
                        // 释放COM对象
                        if (doc != null) Marshal.ReleaseComObject(doc);
                    }
                }
            }
            finally
            {
                // 清理Word进程
                if (wordApp != null)
                {
                    wordApp.Quit(SaveChanges: false);
                    Marshal.ReleaseComObject(wordApp);
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
        }

        /// <summary>
        /// 在单个文档中查找关键词并返回上下文
        /// </summary>
        private List<string> FindKeywordInDocument(Document doc, string keyword)
        {
            var results = new List<string>();
            Microsoft.Office.Interop.Word.Range range = doc.Content;
            Find find = range.Find;

            try
            {
                find.ClearFormatting();
                find.Text = keyword;
                find.MatchCase = false;
                find.MatchWholeWord = false;

                while (find.Execute())
                {
                    if (range.Find.Found)
                    {
                        // 获取匹配段落全文作为上下文
                        string context = range.Paragraphs[1].Range.Text.Trim();
                        results.Add(context);
                    }
                }
            }
            finally
            {
                Marshal.ReleaseComObject(find);
                Marshal.ReleaseComObject(range);
            }

            return results;
        }
        /// <summary>
        /// 批量查找文档内容
        /// </summary>
        private void SearchFiles(string[] sourceFiles, string[] keywords)
        {
            int totalFiles = sourceFiles.Length;
            int processedCount = 0;

            Application wordApp = null;
            try
            {
                wordApp = new Application
                {
                    Visible = false,
                    ScreenUpdating = false,
                    DisplayAlerts = WdAlertLevel.wdAlertsNone
                };

                foreach (var filePath in sourceFiles)
                {
                    // 更新进度
                    Interlocked.Increment(ref processedCount);
                    UpdateProgress(processedCount, totalFiles, Path.GetFileName(filePath));

                    Document doc = null;
                    try
                    {
                        doc = wordApp.Documents.Open(filePath, ReadOnly: true, Visible: false);

                        foreach (var keyword in keywords)
                        {
                            if (string.IsNullOrWhiteSpace(keyword)) continue;

                            var foundItems = FindKeywordInDocument(doc, keyword.Trim());
                            if (foundItems.Count > 0)
                            {
                                // 跨线程更新UI
                                this.Invoke((MethodInvoker)delegate
                                {
                                    foreach (var item in foundItems)
                                    {
                                        var listItem = new ListViewItem(Path.GetFileName(filePath));
                                        listItem.SubItems.Add(keyword);
                                        listItem.SubItems.Add(item);
                                        listView1.Items.Add(listItem);
                                    }
                                });
                            }
                        }
                        doc.Close(SaveChanges: false);
                    }
                    catch (COMException ex) // 新增的异常捕获
                    {
                        this.Invoke((MethodInvoker)delegate
                        {
                            MessageBox.Show($"无法打开文件 {Path.GetFileName(filePath)}：{ex.Message}");
                        });
                        continue; // 跳过当前文件
                    }
                    finally
                    {
                        if (doc != null) Marshal.ReleaseComObject(doc);
                    }
                }
            }
            finally
            {
                if (wordApp != null)
                {
                    wordApp.Quit(SaveChanges: false);
                    Marshal.ReleaseComObject(wordApp);
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
        }
        /// <summary>
        /// 在文档中执行文本替换
        /// </summary>
        private void ReplaceText(Document doc, string oldText, string newText)
        {
            var range = doc.Content;
            var find = range.Find;

            try
            {
                // 配置查找参数
                find.ClearFormatting();
                find.Replacement.ClearFormatting();

                // 执行全部替换
                find.Execute(
                    FindText: oldText,
                    ReplaceWith: newText,
                    Replace: WdReplace.wdReplaceAll,
                    MatchCase: false,
                    MatchWholeWord: false
                );
            }
            finally
            {
                // 释放COM对象
                Marshal.ReleaseComObject(find);
                Marshal.ReleaseComObject(range);
            }
        }

        /// <summary>
        /// 解析替换规则（支持多行）
        /// </summary>
        private Tuple<string, string>[] ParseReplaceRules()
        {
            // 获取两个文本框的内容
            var findLines = textBox3.Lines;
            var replaceLines = textBox4.Lines;

            // 验证规则数量匹配
            if (findLines.Length != replaceLines.Length)
                throw new ArgumentException("查找文本与被替换文本行数不匹配");

            // 创建规则数组
            var rules = new Tuple<string, string>[findLines.Length];
            for (int i = 0; i < findLines.Length; i++)
            {
                rules[i] = Tuple.Create(findLines[i].Trim(), replaceLines[i].Trim());
            }

            return rules;
        }

        /// <summary>
        /// 验证用户输入有效性
        /// </summary>
        private bool ValidateInputs()
        {
            // 验证目录路径
            if (string.IsNullOrWhiteSpace(textBox1.Text) ||
                string.IsNullOrWhiteSpace(textBox2.Text))
            {
                MessageBox.Show("请先选择源目录和目标目录");
                return false;
            }

            // 验证目录存在性
            if (!Directory.Exists(textBox1.Text))
            {
                MessageBox.Show("源目录不存在");
                return false;
            }

            // 验证规则输入
            if (textBox3.Lines.Length == 0 ||
                textBox4.Lines.Length == 0)
            {
                MessageBox.Show("请输入替换的文本");
                return false;
            }

            return true;
        }

        /// <summary>
        /// 验证查找操作的输入
        /// </summary>
        private bool ValidateSearchInputs()
        {
            if (string.IsNullOrWhiteSpace(textBox1.Text))
            {
                MessageBox.Show("请先选择源目录");
                return false;
            }

            if (!Directory.Exists(textBox1.Text))
            {
                MessageBox.Show("源目录不存在");
                return false;
            }

            if (textBox3.Lines.Length == 0 || string.IsNullOrWhiteSpace(textBox3.Text))
            {
                MessageBox.Show("请输入要查找的文本");
                return false;
            }

            return true;
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            MessageBox.Show("1.在使用工具查找或替换前关闭所有word进程！\n2.使用替换和查找期间不要使用word进行操作！\n3.查找和替换执行期间不要中途中断任务，如果中断任务需要自行清理内存中的word进程！");
        }
    }
}