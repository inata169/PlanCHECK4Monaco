using System;
using System.Collections.Generic;
using System.Configuration;
using System.Windows.Forms;
using System.Threading;
using System.Linq;
using System.Text;
using System.IO;
using System.Drawing;
using System.Text.RegularExpressions;
using Elekta.MonacoScripting.API;
using Elekta.MonacoScripting.API.General;
using Elekta.MonacoScripting.API.Planning;
using Elekta.MonacoScripting.API.DICOMImport;
using Elekta.MonacoScripting.API.Beams;
using Elekta.MonacoScripting.DataType;
using Elekta.MonacoScripting.Log;

namespace EnhancedMonacoPlanCheck
{
    public class PlanCheckResult
    {
        public string Category { get; set; }
        public string Item { get; set; }
        public bool Pass { get; set; }
        public string ActualValue { get; set; }
        public string ExpectedValue { get; set; }
        public string Severity { get; set; }
    }

    public class ResultForm : Form
    {
        private readonly DataGridView grid;
        private readonly Label summaryLabel;

        public ResultForm()
        {
            this.Text = "プランチェック結果";
            this.Size = new Size(1400, 800);
            this.StartPosition = FormStartPosition.CenterScreen;

            grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AllowUserToAddRows = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                RowHeadersVisible = false,
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            };

            summaryLabel = new Label
            {
                Dock = DockStyle.Bottom,
                Height = 30,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font(Font.FontFamily, 10, FontStyle.Bold)
            };

            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(10)
            };
            panel.Controls.Add(grid);

            this.Controls.AddRange(new Control[] { panel, summaryLabel });
        }

        public void DisplayResults(List<PlanCheckResult> results)
        {
            grid.Columns.Clear();
            grid.Columns.AddRange(
                new DataGridViewTextBoxColumn { Name = "Category", HeaderText = "カテゴリ", Width = 100 },
                new DataGridViewTextBoxColumn { Name = "Item", HeaderText = "項目", Width = 200 },
                new DataGridViewCheckBoxColumn { Name = "Pass", HeaderText = "合格", Width = 50 },
                new DataGridViewTextBoxColumn { Name = "ActualValue", HeaderText = "実際値", Width = 200 },
                new DataGridViewTextBoxColumn { Name = "ExpectedValue", HeaderText = "期待値", Width = 200 },
                new DataGridViewTextBoxColumn { Name = "Severity", HeaderText = "重要度", Width = 80 }
            );

            foreach (var result in results.OrderBy(r => r.Severity == "エラー" ? 0 : r.Severity == "警告" ? 1 : 2)
                                        .ThenBy(r => r.Category == "処方" ? 0 : 1)
                                        .ThenBy(r => r.Category)
                                        .ThenBy(r => r.Item))
            {
                var rowIndex = grid.Rows.Add(
                    result.Category,
                    result.Item,
                    result.Pass,
                    result.ActualValue,
                    result.ExpectedValue,
                    result.Severity
                );

                var row = grid.Rows[rowIndex];
                switch (result.Severity)
                {
                    case "エラー":
                        row.DefaultCellStyle.BackColor = Color.LightPink;
                        break;
                    case "警告":
                        row.DefaultCellStyle.BackColor = Color.LightYellow;
                        break;
                    default:
                        row.DefaultCellStyle.BackColor = Color.White;
                        break;
                }
            }

            int errorCount = results.Count(r => r.Severity == "エラー");
            int warningCount = results.Count(r => r.Severity == "警告");
            int infoCount = results.Count(r => r.Severity == "情報");
            summaryLabel.Text = $" 概要: {errorCount} エラー, {warningCount} 警告, {infoCount} 情報項目";
        }
    }

    internal class Program
    {
        private const double MIN_MU = 10.0;
        private const double MIN_DOSE = 1.0;
        private const double MAX_DOSE_RATIO = 1.17;
        private const double MIN_COLLIMATOR_SIZE = 3.0;

        [STAThread]
        static void Main()
        {
            try
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);

                Logger.Instance.DeactivateExitErrorHandler();
                Logger.Instance.ActivateScreenShot();

                MonacoApplication app = MonacoApplication.Instance ??
                    throw new Exception("Monacoアプリケーションの初期化に失敗しました。");
                var results = new List<PlanCheckResult>();

                var pat = app.GetCurrentPatient() ??
                    throw new Exception("現在の患者情報の取得に失敗しました。");
                var plan = app.GetActivePlan() ??
                    throw new Exception("アクティブな治療計画の取得に失敗しました。");

                //
                //What kind of CHECK's ?
                //
                CheckBasicPlanInfo(pat, plan, results);
                CheckPrescriptionAndDose(app, results);
                CheckBeamProperties(app.GetBeamsSpreadsheet(), results);
                CheckDVHStatistics(app, results);
                CheckCalculationSettings(app, results);

                //

                var form = new ResultForm();
                form.DisplayResults(results);
                Application.Run(form);

                SaveResultsToFile(pat, plan, results);

                Logger.Instance.Info("プランチェックが正常に完了しました。");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"プランチェック中にエラーが発生しました: {ex.Message}", "エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                Logger.Instance.Error($"プランチェック中にエラーが発生しました: {ex.Message}");
                if (ex.InnerException != null)
                {
                    Logger.Instance.Error($"内部例外: {ex.InnerException.Message}");
                }
                Logger.Instance.Error($"スタックトレース: {ex.StackTrace}");
            }
        }

        private static double GetMaxPlanDose(MonacoApplication app)
        {
            if (!app.VerifyMaxDose())
                return 0;

            var dvhStats = app.GetDVHStatisticsSpreadsheet();
            if (dvhStats == null) return 0;

            var statsInfo = dvhStats.GetStatisticsInfo();
            if (statsInfo?.StatisticsOfStructureList == null) return 0;

            // 全構造の中から最大線量を探す
            double maxDose = 0;
            foreach (var structure in statsInfo.StatisticsOfStructureList)
            {
                if (structure.MaxDose > maxDose)
                {
                    maxDose = structure.MaxDose;
                }
            }

            Logger.Instance.Info($"最大線量チェック: 治療計画最大線量 / 処方線量 <= {MAX_DOSE_RATIO}");
            return maxDose;
        }

        private static void CheckPrescriptionAndDose(MonacoApplication app, List<PlanCheckResult> results)
        {
            try
            {
                var prescription = app.GetPrescription();
                if (prescription == null)
                {
                    Logger.Instance.Warn("処方情報を取得できませんでした。");
                    return;
                }

                var prescriptionInfo = prescription.getAllPrescriptionInfo();
                if (prescriptionInfo?.PrescriptionUIpropertyList?.Any() != true)
                {
                    Logger.Instance.Warn("処方UIプロパティが見つかりませんでした。");
                    return;
                }

                var rxProperty = prescriptionInfo.PrescriptionUIpropertyList[0];
                double rxDose = rxProperty.RxDose;
                double maxDose = GetMaxPlanDose(app);

                results.Add(new PlanCheckResult
                {
                    Category = "処方",
                    Item = "処方線量",
                    Pass = rxDose >= MIN_DOSE,
                    ActualValue = $"{rxDose:F2} cGy",
                    ExpectedValue = $">= {MIN_DOSE} cGy",
                    Severity = rxDose >= MIN_DOSE ? "情報" : "エラー"
                });

                results.Add(new PlanCheckResult
                {
                    Category = "処方",
                    Item = "分割線量",
                    Pass = rxProperty.FractionalDose >= MIN_DOSE,
                    ActualValue = $"{rxProperty.FractionalDose:F2} cGy",
                    ExpectedValue = $">= {MIN_DOSE} cGy",
                    Severity = rxProperty.FractionalDose >= MIN_DOSE ? "情報" : "エラー"
                });

                if (maxDose > 0 && rxDose > 0)
                {
                    var doseRatio = maxDose / rxDose;
                    results.Add(new PlanCheckResult
                    {
                        Category = "処方",
                        Item = "最大線量比",
                        Pass = doseRatio <= MAX_DOSE_RATIO,
                        ActualValue = $"{doseRatio:F3} ({maxDose:F1} cGy / {rxDose:F1} cGy)",
                        ExpectedValue = $"<= {MAX_DOSE_RATIO:F2}",
                        Severity = doseRatio <= MAX_DOSE_RATIO ? "情報" : "警告"
                    });
                }
            }
            catch (Exception ex)
            {
                Logger.Instance.Error($"処方と線量のチェック中にエラーが発生しました: {ex.Message}");
            }
        }

        private static void CheckBasicPlanInfo(PatientDemographic pat, string planId, List<PlanCheckResult> results)
        {
            try
            {
                results.Add(new PlanCheckResult
                {
                    Category = "患者",
                    Item = "患者ID",
                    Pass = true,
                    ActualValue = pat?.Id ?? "N/A",
                    Severity = "情報"
                });

                results.Add(new PlanCheckResult
                {
                    Category = "患者",
                    Item = "患者名",
                    Pass = true,
                    ActualValue = pat?.Name ?? "N/A",
                    Severity = "情報"
                });

                results.Add(new PlanCheckResult
                {
                    Category = "患者",
                    Item = "クリニック",
                    Pass = true,
                    ActualValue = pat?.Clinic ?? "N/A",
                    Severity = "情報"
                });

                // プランID形式のチェック - 3D, VMAT, DCATに続く3桁の数字
                var planIdPattern = @"^(3D|VMAT|DCAT)\d{3}$";
                var planIdValid = !string.IsNullOrEmpty(planId) && Regex.IsMatch(planId, planIdPattern);
                results.Add(new PlanCheckResult
                {
                    Category = "プラン",
                    Item = "プランID形式",
                    Pass = planIdValid,
                    ActualValue = planId,
                    ExpectedValue = "(3D|VMAT|DCAT) + 3桁の数字",
                    Severity = planIdValid ? "情報" : "エラー"
                });
            }
            catch (Exception ex)
            {
                Logger.Instance.Error($"基本プラン情報のチェック中にエラーが発生しました: {ex.Message}");
            }
        }

        private static void CheckBeamProperties(BeamsSpreadsheet sheet, List<PlanCheckResult> results)
        {
            if (sheet == null)
            {
                Logger.Instance.Warn("ビームスプレッドシートを取得できませんでした。");
                return;
            }

            try
            {
                var generalProps = sheet.GetBeamGeneralProperties();
                var geomProps = sheet.GetBeamGeometryProperties();
                var treatmentAids = sheet.GetTreatmentAidsProperties();

                if (generalProps == null || geomProps == null || treatmentAids == null)
                {
                    Logger.Instance.Warn("ビームプロパティの一部または全部を取得できませんでした。");
                    return;
                }

                CheckIsocenterConsistency(generalProps, results);

                foreach (var prop in generalProps)
                {
                    if (prop != null)
                    {
                        AddBeamGeneralChecks(prop, results);
                    }
                }

                foreach (var geom in geomProps)
                {
                    if (geom != null)
                    {
                        AddGeometryChecks(geom, results);
                        // コリメータの詳細 (Width1, Width2, Length1, Length2) をチェック
                        AddCollimatorChecks(geom, results);
                    }
                }

                foreach (var aid in treatmentAids)
                {
                    if (aid != null)
                    {
                        AddTreatmentAidsChecks(aid, results);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Instance.Error($"ビームプロパティのチェック中にエラーが発生しました: {ex.Message}");
            }
        }

        private static void CheckIsocenterConsistency(List<BeamGeneralProperty> generalProps, List<PlanCheckResult> results)
        {
            bool allSameIsocenter = generalProps.Select(bp => bp.BeamIso?.IsocenterLocation).Distinct().Count() == 1;
            results.Add(new PlanCheckResult
            {
                Category = "ビーム",
                Item = "アイソセンター一貫性",
                Pass = allSameIsocenter,
                ActualValue = allSameIsocenter ? "全ビームが同じアイソセンターを持つ" : "アイソセンターが一致しない",
                ExpectedValue = "全ビームが同じアイソセンターを持つべき",
                Severity = allSameIsocenter ? "情報" : "警告"
            });
        }

        private static void AddBeamGeneralChecks(BeamGeneralProperty prop, List<PlanCheckResult> results)
        {
            // フィールドID形式のチェック - 数字4桁が必要
            var fieldIdPattern = @"^\d{4}$";
            bool fieldIdValid = !string.IsNullOrEmpty(prop.FieldID) && Regex.IsMatch(prop.FieldID, fieldIdPattern);

            results.Add(new PlanCheckResult
            {
                Category = "ビーム",
                Item = $"フィールドID形式 (ビーム {prop.BeamID})",
                Pass = fieldIdValid,
                ActualValue = prop.FieldID ?? "N/A",
                ExpectedValue = "数字4桁",
                Severity = fieldIdValid ? "情報" : "エラー"
            });

            results.Add(new PlanCheckResult
            {
                Category = "ビーム",
                Item = $"最小MU (ビーム {prop.BeamID})",
                Pass = prop.MUFx >= MIN_MU,
                ActualValue = prop.MUFx.ToString("F2"),
                ExpectedValue = $">= {MIN_MU}",
                Severity = prop.MUFx >= MIN_MU ? "情報" : "エラー"
            });

            results.Add(new PlanCheckResult
            {
                Category = "ビーム",
                Item = $"計算アルゴリズム (ビーム {prop.BeamID})",
                Pass = true,
                ActualValue = prop.Algorithm.ToString(),
                Severity = "情報"
            });

            if (!string.IsNullOrEmpty(prop.Energy))
            {
                results.Add(new PlanCheckResult
                {
                    Category = "ビーム",
                    Item = $"エネルギー (ビーム {prop.BeamID})",
                    Pass = true,
                    ActualValue = prop.Energy,
                    Severity = "情報"
                });
            }

            if (prop.BeamIso != null)
            {
                results.Add(new PlanCheckResult
                {
                    Category = "ビーム",
                    Item = $"アイソセンター位置 (ビーム {prop.BeamID})",
                    Pass = true,
                    ActualValue = $"({prop.BeamIso.IsoX:F2}, {prop.BeamIso.IsoY:F2}, {prop.BeamIso.IsoZ:F2})",
                    ExpectedValue = "PTVの中心",
                    Severity = "情報"
                });
            }
        }

        private static void AddGeometryChecks(BeamGeometryProperty geom, List<PlanCheckResult> results)
        {
            // 角度情報をチェックし、結果を追加
            results.Add(new PlanCheckResult
            {
                Category = "ジオメトリ",
                Item = $"角度 (ビーム {geom.BeamID})",  // どのビームの角度かを示す
                Pass = true,
                ActualValue = $"ガントリ: {geom.Gantry:F1}, コリメータ: {geom.Collimator:F1}, カウチ: {geom.Couch:F1}",   // 各角度を表示
                Severity = "情報"
            });

            // アーク治療の場合、アーク方向やアーク長をチェック
            if (geom.Dir == Direction.CW || geom.Dir == Direction.CCW)
            {
                results.Add(new PlanCheckResult
                {
                    Category = "ジオメトリ",
                    Item = $"アーク方向 (ビーム {geom.BeamID})",  // どのビームのアーク方向かを示す
                    Pass = true,
                    ActualValue = geom.Dir.ToString(),// 実際のアーク方向を表示
                    Severity = "情報"
                });

                results.Add(new PlanCheckResult
                {
                    Category = "ジオメトリ",
                    Item = $"アーク長 (ビーム {geom.BeamID})", // どのアーク長かを示す
                    Pass = true,
                    ActualValue = $"{geom.Arc:F1}°",  // 実際のアーク長を表示
                    Severity = "情報"
                });
            }
        }
        // コリメータの情報をチェックし、結果を追加するメソッド
        private static void AddCollimatorChecks(BeamGeometryProperty geom, List<PlanCheckResult> results)
        {
            // コリメータの情報をチェックし、結果を追加
            bool collimatorSizeCheck = (geom.Width1 + geom.Width2 > MIN_COLLIMATOR_SIZE) && (geom.Length1 + geom.Length2 > MIN_COLLIMATOR_SIZE);
            results.Add(new PlanCheckResult
            {
                Category = "コリメータ",
                Item = $"コリメータ詳細 (ビーム {geom.BeamID})",  // どのビームのコリメータかを示す
                Pass = collimatorSizeCheck,
                ActualValue = $"Width1: {geom.Width1:F1} cm, Width2: {geom.Width2:F1} cm, Length1: {geom.Length1:F1} cm, Length2: {geom.Length2:F1} cm",// コリメータの各サイズを表示
                ExpectedValue = $"Width1 + Width2 > {MIN_COLLIMATOR_SIZE} cm, Length1 + Length2 > {MIN_COLLIMATOR_SIZE} cm. 半開の場合、Lower:Length1=Y2=0.0cm、Upper:Length2=y1=0.0cm",
                Severity = collimatorSizeCheck ? "情報" : "警告"
            });
        }


        private static void AddTreatmentAidsChecks(BeamTreatmentAidsProperty aid, List<PlanCheckResult> results)
        {
            // カウチのチェック結果を追加
            results.Add(new PlanCheckResult
            {
                Category = "治療補助具",
                Item = $"カウチ有効化 (ビーム {aid.BeamID})", // どのビームのカウチ有効化かを示す
                Pass = aid.Couch, // カウチが有効かどうか
                ActualValue = aid.Couch.ToString(),  // 実際のカウチの設定を表示
                ExpectedValue = "True",
                Severity = aid.Couch ? "情報" : "警告" // 有効なら情報、無効なら警告
            });

            // ボーラスの情報を追加
            results.Add(new PlanCheckResult
            {
                Category = "治療補助具",
                Item = $"ボーラス (ビーム {aid.BeamID})",  // どのビームのボーラスかを示す
                Pass = true,
                ActualValue = string.IsNullOrEmpty(aid.Bolus) ? "なし" : aid.Bolus,  // ボーラスがある場合は名前を表示、ない場合は「なし」を表示
                Severity = "情報"
            });

            // ウェッジIDの情報を追加
            if (!string.IsNullOrEmpty(aid.WedgeID))
            {
                results.Add(new PlanCheckResult
                {
                    Category = "治療補助具",
                    Item = $"ウェッジID (ビーム {aid.BeamID})",    // どのビームのウェッジIDかを示す
                    Pass = true,
                    ActualValue = aid.WedgeID,  // ウェッジIDを表示
                    Severity = "情報"
                });
            }
        }

        // DVH統計情報をチェックするメソッド
        private static void CheckDVHStatistics(MonacoApplication app, List<PlanCheckResult> results)
        {
            try
            {
                var dvhStats = app.GetDVHStatisticsSpreadsheet(); // DVH統計のスプレッドシートを取得
                Thread.Sleep(500);  // DVH統計の計算を待つ
                var statinfo = dvhStats?.GetStatisticsInfo();   // 統計情報を取得

                // 統計情報がある場合
                if (statinfo?.StatisticsOfStructureList != null)
                {
                    // 各構造について
                    foreach (var statProp in statinfo.StatisticsOfStructureList)
                    {
                        if (statProp.StructureName.Contains("PTV"))  // PTVを含む構造の場合
                        {
                            // 適合度指数についてのチェック結果を追加
                            results.Add(new PlanCheckResult
                            {
                                Category = "DVH統計",
                                Item = $"CI ({statProp.StructureName})",    // どの構造の適合度指数かを示す
                                Pass = true,
                                ActualValue = statProp.ConformityIndex.ToString("F2"), // 実際の適合度指数を表示
                                ExpectedValue = "Conformity Index (CI)",
                                Severity = "情報"
                            });

                            // 不均一性指数のチェック結果を追加
                            results.Add(new PlanCheckResult
                            {
                                Category = "DVH統計",
                                Item = $"HI ({statProp.StructureName})",  // どの構造の不均一性指数かを示す
                                Pass = true,
                                ActualValue = statProp.HeterogeneityIndex.ToString("F2"),   // 実際の不均一性指数を表示
                                ExpectedValue = "Heterogeneity Index",
                                Severity = "情報"
                            });
                        }
                    }
                }
                else
                {
                    // DVH統計が取得できない場合、警告ログを出力
                    Logger.Instance.Warn("DVH統計を取得できませんでした");
                }
            }
            catch (Exception ex)
            {
                // エラーが発生した場合、ログを出力
                Logger.Instance.Error($"DVH統計のチェック中にエラーが発生しました: {ex.Message}");
            }
        }

        // 計算設定に関するチェックを行うメソッド
        private static void CheckCalculationSettings(MonacoApplication app, List<PlanCheckResult> results)
        {
            try
            {
                // 計算プロパティ設定を取得
                var calcSettings = app.GetCalculationPropertiesSettings();
                if (calcSettings == null)
                {
                    Logger.Instance.Warn("計算プロパティを取得できませんでした。");
                    return;
                }
                else
                {
                    // 計算プロパティを取得
                    var calcProp = calcSettings.GetCalculationProperty();
                    if (calcProp != null)
                    {
                        // 線量計算アルゴリズムのチェック
                        results.Add(new PlanCheckResult
                        {
                            Category = "計算設定",
                            Item = "線量計算アルゴリズム",
                            Pass = true,
                            ActualValue = calcProp.DoseDeposition ?? "N/A",
                            ExpectedValue = "Calc DoseDeposition to",
                            Severity = "情報"
                        });
                        // 最終計算アルゴリズムのチェック
                        results.Add(new PlanCheckResult
                        {
                            Category = "計算設定",
                            Item = "最終計算アルゴリズム",
                            Pass = true,
                            ActualValue = calcProp.FinalCalculationAlg ?? "N/A",
                            Severity = "情報"
                        });
                        // グリッド間隔のチェック
                        results.Add(new PlanCheckResult
                        {
                            Category = "計算設定",
                            Item = "グリッド間隔",
                            Pass = true,
                            ActualValue = calcProp.GridSpacing.ToString("F2"),
                            ExpectedValue = "0.1 ~ 0.8",
                            Severity = "情報"
                        });
                        // ビームあたりの最大粒子数のチェック
                        results.Add(new PlanCheckResult
                        {
                            Category = "計算設定",
                            Item = "ビームあたりの最大粒子数",
                            Pass = true,
                            ActualValue = calcProp.MaxParticlesPerBeam.ToString("F0"),
                            ExpectedValue = "設定済み",
                            Severity = "情報"
                        });
                        // スポットごとの粒子の不確かさのチェック
                        //results.Add(new PlanCheckResult
                        //{
                        //    Category = "計算設定",
                        //    Item = "1スロットあたりの粒子の不確かさ(%)",
                        //    Pass = true,
                        //    ActualValue = calcProp.UncentaintyPerSpot.ToString("F2") + "%",
                        //    ExpectedValue = "0%または設定値",
                        //    Severity = "情報"
                        //});
                    }
                }
            }
            catch (Exception ex)
            {
                // エラーが発生した場合、ログを出力
                Logger.Instance.Error($"計算設定のチェック中にエラーが発生しました: {ex.Message}");
            }
        }

        private static void SaveResultsToFile(PatientDemographic pat, string planId, List<PlanCheckResult> results)
        {
            try
            {
                // 現在時刻をタイムスタンプとしてフォーマット
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                // ファイル名を作成
                string fileName = $"PlanCheck_{pat.Id}_{pat.Name}_{planId}_{timestamp}.txt";
                // ファイルパスを作成（デスクトップ上の MonocoPlanCheck フォルダ内）
                string filePath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                    "MonacoPlanCheck",
                    fileName);

                // フォルダが存在しなければ作成
                Directory.CreateDirectory(Path.GetDirectoryName(filePath));

                // 結果をテーブル形式に変換
                string tableResult = ConvertToTable(results);
                // ファイルに結果を書き込み
                File.WriteAllText(filePath, tableResult);

                // メッセージボックスで結果保存の成功を通知
                MessageBox.Show(
                    $"プランチェックが完了しました。結果は以下に保存されました:\n{filePath}",
                    "Monaco プランチェック",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                // 保存したファイルをエクスプローラーで開く
                //System.Diagnostics.Process.Start(filePath);
            }
            catch (Exception ex)
            {
                // エラーが発生した場合、ログを出力
                Logger.Instance.Error($"結果をファイルに保存中にエラーが発生しました: {ex.Message}");
                // エラーが発生した場合、ユーザーにエラーメッセージを表示
                MessageBox.Show(
                    $"結果をファイルに保存中にエラーが発生しました: {ex.Message}",
                    "エラー",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        // リストされた結果をテーブル形式の文字列に変換する - 見栄えを改善
        private static string ConvertToTable(List<PlanCheckResult> results)
        {
            StringBuilder table = new StringBuilder();

            // タイトルとヘッダー
            table.AppendLine("==================================================================================");
            table.AppendLine("                               Monaco プランチェック結果                          ");
            table.AppendLine("==================================================================================");
            table.AppendLine();

            // 統計情報
            int errorCount = results.Count(r => r.Severity == "エラー");
            int warningCount = results.Count(r => r.Severity == "警告");
            int infoCount = results.Count(r => r.Severity == "情報");
            table.AppendLine($"概要: {errorCount} エラー, {warningCount} 警告, {infoCount} 情報項目");
            table.AppendLine();

            // カテゴリごとに結果を整理
            var categorizedResults = results
                .OrderBy(r => r.Severity == "エラー" ? 0 : r.Severity == "警告" ? 1 : 2)
                .ThenBy(r => r.Category)
                .ThenBy(r => r.Item)
                .GroupBy(r => r.Category);

            foreach (var category in categorizedResults)
            {
                // カテゴリヘッダー
                table.AppendLine($"【{category.Key}】");
                table.AppendLine(new string('-', 80));
                table.AppendLine(String.Format("{0,-30} {1,-7} {2,-20} {3,-20}", "項目", "結果", "実際値", "期待値"));
                table.AppendLine(new string('-', 80));

                // カテゴリ内の各項目
                foreach (var item in category)
                {
                    string resultMark = item.Pass ? "✓" : "✗";
                    string severityMark = "";

                    if (item.Severity == "エラー")
                        severityMark = "[E]";
                    else if (item.Severity == "警告")
                        severityMark = "[W]";

                    table.AppendLine(String.Format("{0,-30} {1,-7} {2,-20} {3,-20}",
                        item.Item,
                        resultMark + severityMark,
                        item.ActualValue ?? "",
                        item.ExpectedValue ?? ""));
                }

                table.AppendLine();
            }

            table.AppendLine("==================================================================================");
            table.AppendLine($"生成日時: {DateTime.Now}");

            return table.ToString();
        }

        // 旧メソッド - 不要になったため削除
        //private static int[] GetColumnWidths(List<PlanCheckResult> results)
        //{
        //    // 削除
        //}

        // 旧メソッド - 不要になったため削除
        //private static string CreateTableRow(int[] columnWidths, params string[] values)
        //{
        //    // 削除
        //}

        // 旧メソッド - 不要になったため削除
        //private static string CreateSeparatorLine(int[] columnWidths)
        //{
        //    // 削除
        //}
    }
}