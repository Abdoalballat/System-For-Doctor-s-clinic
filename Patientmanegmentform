using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Data.SQLite;

namespace ClinicManagementSystem
{
    public partial class PatientManagementForm : Form
    {
        private DataTable patientTable;
        private string connectionString;

        public PatientManagementForm()
        {
            InitializeComponent();
            connectionString = DatabaseManager.GetConnectionString();

            InitializeDatabase();
            SetupDataGridViewStyle();
            AttachEvents();
            LoadData();
            UpdateGenderSelection();
            txtDebt.Text = "0";
        }

        private bool ColumnExists(SQLiteConnection connection, string tableName, string columnName)
        {
            using (var cmd = new SQLiteCommand($"PRAGMA table_info({tableName})", connection))
            using (var reader = cmd.ExecuteReader())
            {
                while (reader.Read())
                {
                    if (reader["name"].ToString().Equals(columnName, StringComparison.OrdinalIgnoreCase))
                        return true;
                }
            }
            return false;
        }

        private void InitializeDatabase()
        {
            using (var connection = new SQLiteConnection(connectionString))
            {
                connection.Open();

                string sqlPatients = @"CREATE TABLE IF NOT EXISTS Patients (
                                ID INTEGER PRIMARY KEY AUTOINCREMENT,
                                Name TEXT NOT NULL,
                                Phone TEXT,
                                Age TEXT,
                                Gender TEXT,
                                BloodType TEXT,
                                Chronic TEXT,
                                Meds TEXT,
                                Debt INTEGER DEFAULT 0)";
                using (var command = new SQLiteCommand(sqlPatients, connection)) command.ExecuteNonQuery();

                if (!ColumnExists(connection, "Patients", "Debt"))
                {
                    using (var cmd = new SQLiteCommand("ALTER TABLE Patients ADD COLUMN Debt INTEGER DEFAULT 0", connection)) cmd.ExecuteNonQuery();
                }
            }
        }

        private void SetupDataGridViewStyle()
        {
            patientTable = new DataTable();
            patientTable.Columns.Add("الكود");
            patientTable.Columns.Add("الاسم");
            patientTable.Columns.Add("الهاتف");
            patientTable.Columns.Add("العمر");
            patientTable.Columns.Add("النوع");
            patientTable.Columns.Add("الدم");
            patientTable.Columns.Add("أمراض مزمنة");
            patientTable.Columns.Add("الأدوية");
            patientTable.Columns.Add("الديون");
            dgvPatients.DataSource = patientTable;

            dgvPatients.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(52, 152, 219);
            dgvPatients.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvPatients.ColumnHeadersDefaultCellStyle.Font = new Font(FontFamily.GenericSansSerif, 10, FontStyle.Bold);
            dgvPatients.EnableHeadersVisualStyles = false;
        }

        private void LoadData()
        {
            patientTable.Rows.Clear();
            using (var connection = new SQLiteConnection(connectionString))
            {
                connection.Open();
                string sql = "SELECT ID, Name, Phone, Age, Gender, BloodType, Chronic, Meds, IFNULL(Debt, 0) AS Debt FROM Patients";
                using (var command = new SQLiteCommand(sql, connection))
                {
                    using (var reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            patientTable.Rows.Add(
                                reader["ID"].ToString(),
                                reader["Name"].ToString(),
                                reader["Phone"].ToString(),
                                reader["Age"].ToString(),
                                reader["Gender"].ToString(),
                                reader["BloodType"].ToString(),
                                reader["Chronic"].ToString(),
                                reader["Meds"].ToString(),
                                reader["Debt"].ToString()
                            );
                        }
                    }
                }
            }
        }

        private void AttachEvents()
        {
            txtPhone.KeyPress += (s, e) => { if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar)) e.Handled = true; };
            txtDebt.KeyPress += (s, e) => { if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar)) e.Handled = true; };

            txtName.TextChanged += (s, e) => { lblNameError.Visible = string.IsNullOrWhiteSpace(txtName.Text); };
            txtAge.TextChanged += (s, e) => { lblAgeError.Visible = string.IsNullOrWhiteSpace(txtAge.Text); };

            txtName.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; txtPhone.Focus(); } };
            txtPhone.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; txtAge.Focus(); } };
            txtAge.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; cmbBloodType.Focus(); } };
            cmbBloodType.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; txtChronic.Focus(); } };
            txtChronic.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; txtMeds.Focus(); } };
            txtMeds.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; txtDebt.Focus(); } };
            txtDebt.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; btnAdd.PerformClick(); } };

            btnAdd.Click += (s, e) => AddPatient();
            btnClear.Click += (s, e) => ClearInputs();
            btnUpdate.Click += (s, e) => UpdatePatient();
            btnDelete.Click += (s, e) => DeletePatient();

            dgvPatients.CellClick += (s, e) => {
                if (dgvPatients.CurrentRow != null)
                {
                    var row = dgvPatients.CurrentRow;
                    txtID.Text = row.Cells[0].Value?.ToString();
                    txtName.Text = row.Cells[1].Value?.ToString();
                    txtPhone.Text = row.Cells[2].Value?.ToString();
                    txtAge.Text = row.Cells[3].Value?.ToString();
                    if (row.Cells[4].Value?.ToString() == "ذكر") rbMale.Checked = true; else rbFemale.Checked = true;
                    cmbBloodType.Text = row.Cells[5].Value?.ToString();
                    txtChronic.Text = row.Cells[6].Value?.ToString();
                    txtMeds.Text = row.Cells[7].Value?.ToString();
                    txtDebt.Text = row.Cells[8].Value?.ToString();
                }
            };

            txtSearch.Enter += (s, e) => { if (txtSearch.Text == "الاسم أو رقم الهاتف...") { txtSearch.Text = ""; txtSearch.ForeColor = Color.Black; } };
            txtSearch.Leave += (s, e) => { if (string.IsNullOrWhiteSpace(txtSearch.Text)) { txtSearch.Text = "الاسم أو رقم الهاتف..."; txtSearch.ForeColor = Color.Gray; patientTable.DefaultView.RowFilter = string.Empty; } };
            txtSearch.TextChanged += (s, e) => {
                if (txtSearch.Text != "الاسم أو رقم الهاتف..." && !string.IsNullOrWhiteSpace(txtSearch.Text))
                {
                    string query = txtSearch.Text.Replace("'", "''");
                    patientTable.DefaultView.RowFilter = $"الاسم LIKE '%{query}%' OR الهاتف LIKE '%{query}%'";
                }
                else if (txtSearch.Text == "") patientTable.DefaultView.RowFilter = string.Empty;
            };

            rbMale.CheckedChanged += (s, e) => UpdateGenderSelection();
            rbFemale.CheckedChanged += (s, e) => UpdateGenderSelection();
            lblMaleText.Click += (s, e) => rbMale.Checked = true;
            lblFemaleText.Click += (s, e) => rbFemale.Checked = true;
        }

        private void AddPatient()
        {
            if (ValidateInputs())
            {
                string gender = rbMale.Checked ? "ذكر" : "أنثى";
                int debtAmount = string.IsNullOrWhiteSpace(txtDebt.Text) ? 0 : Convert.ToInt32(txtDebt.Text);

                try
                {
                    using (var connection = new SQLiteConnection(connectionString))
                    {
                        connection.Open();
                        string sql = "INSERT INTO Patients (Name, Phone, Age, Gender, BloodType, Chronic, Meds, Debt) VALUES (@Name, @Phone, @Age, @Gender, @BloodType, @Chronic, @Meds, @Debt)";
                        using (var command = new SQLiteCommand(sql, connection))
                        {
                            command.Parameters.AddWithValue("@Name", txtName.Text);
                            command.Parameters.AddWithValue("@Phone", txtPhone.Text);
                            command.Parameters.AddWithValue("@Age", txtAge.Text);
                            command.Parameters.AddWithValue("@Gender", gender);
                            command.Parameters.AddWithValue("@BloodType", cmbBloodType.Text);
                            command.Parameters.AddWithValue("@Chronic", txtChronic.Text);
                            command.Parameters.AddWithValue("@Meds", txtMeds.Text);
                            command.Parameters.AddWithValue("@Debt", debtAmount);
                            command.ExecuteNonQuery();
                        }
                    }

                    Logger.LogAction("إدارة المرضى", $"قام بإنشاء ملف جديد للمريض ({txtName.Text}) برقم هاتف ({txtPhone.Text})");

                    MessageBox.Show("تم إضافة بيانات المريض إلى السجل بنجاح!", "نجاح", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    LoadData();
                    ClearInputs();
                    txtName.Focus();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("خطأ أثناء إضافة المريض: " + ex.Message, "خطأ");
                }
            }
        }

        private int GetOldDebt(int id)
        {
            using (var conn = new SQLiteConnection(connectionString))
            {
                conn.Open();
                string sql = "SELECT Debt FROM Patients WHERE ID=@id";
                using (var cmd = new SQLiteCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@id", id);
                    object res = cmd.ExecuteScalar();
                    if (res != null && res != DBNull.Value)
                        return Convert.ToInt32(res);
                    return 0;
                }
            }
        }

        private void UpdatePatient()
        {
            if (dgvPatients.CurrentRow == null || string.IsNullOrEmpty(txtID.Text)) return;
            if (ValidateInputs())
            {
                string gender = rbMale.Checked ? "ذكر" : "أنثى";
                int patientId = Convert.ToInt32(txtID.Text);
                int newDebt = string.IsNullOrWhiteSpace(txtDebt.Text) ? 0 : Convert.ToInt32(txtDebt.Text);

                try
                {
                    int oldDebt = GetOldDebt(patientId);
                    int paidAmount = oldDebt - newDebt;

                    using (var connection = new SQLiteConnection(connectionString))
                    {
                        connection.Open();

                        string sql = "UPDATE Patients SET Name=@Name, Phone=@Phone, Age=@Age, Gender=@Gender, BloodType=@BloodType, Chronic=@Chronic, Meds=@Meds, Debt=@Debt WHERE ID=@ID";
                        using (var command = new SQLiteCommand(sql, connection))
                        {
                            command.Parameters.AddWithValue("@ID", patientId);
                            command.Parameters.AddWithValue("@Name", txtName.Text);
                            command.Parameters.AddWithValue("@Phone", txtPhone.Text);
                            command.Parameters.AddWithValue("@Age", txtAge.Text);
                            command.Parameters.AddWithValue("@Gender", gender);
                            command.Parameters.AddWithValue("@BloodType", cmbBloodType.Text);
                            command.Parameters.AddWithValue("@Chronic", txtChronic.Text);
                            command.Parameters.AddWithValue("@Meds", txtMeds.Text);
                            command.Parameters.AddWithValue("@Debt", newDebt);
                            command.ExecuteNonQuery();
                        }

                        if (paidAmount > 0)
                        {
                            string today = DateTime.Today.ToString("yyyy-MM-dd");
                            string nowTime = DateTime.Now.ToString("hh:mm tt");

                            string insertApp = "INSERT INTO Appointments (PatientID, AppDate, AppTime, VisitType, Status, Fees, Paid) VALUES (@pId, @date, @time, @type, @status, @fees, @paid)";
                            using (var cmdApp = new SQLiteCommand(insertApp, connection))
                            {
                                cmdApp.Parameters.AddWithValue("@pId", patientId);
                                cmdApp.Parameters.AddWithValue("@date", today);
                                cmdApp.Parameters.AddWithValue("@time", nowTime);
                                cmdApp.Parameters.AddWithValue("@type", "سداد ديون سابقة");
                                cmdApp.Parameters.AddWithValue("@status", "تم الكشف");
                                cmdApp.Parameters.AddWithValue("@fees", 0);
                                cmdApp.Parameters.AddWithValue("@paid", paidAmount);
                                cmdApp.ExecuteNonQuery();
                            }
                        }
                    }

                    Logger.LogAction("إدارة المرضى", $"قام بتعديل بيانات المريض رقم {patientId} (الاسم: {txtName.Text})");

                    if (paidAmount > 0)
                        MessageBox.Show($"تم تحديث البيانات وتسجيل سداد مبلغ ({paidAmount} ج.م) في درج الخزنة بنجاح!", "نجاح مالي", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    else
                        MessageBox.Show("تم تحديث بيانات المريض بنجاح!", "نجاح", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    LoadData();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("خطأ أثناء تحديث البيانات: " + ex.Message, "خطأ");
                }
            }
        }

        // =========================================================
        // الدالة المحدثة لحذف المريض (الاحتفاظ بالحجوزات الذكي)
        // =========================================================
        private void DeletePatient()
        {
            if (dgvPatients.CurrentRow != null && !string.IsNullOrEmpty(txtID.Text))
            {
                DialogResult result = MessageBox.Show($"هل أنت متأكد من حذف ملف المريض ({txtName.Text}) نهائياً؟\n\nتنبيه: سجلات المواعيد السابقة المتعلقة به سيتم الاحتفاظ بها تحت اسم (مريض محذوف) لأغراض الحسابات.", "تأكيد الحذف", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                if (result == DialogResult.Yes)
                {
                    try
                    {
                        using (var connection = new SQLiteConnection(connectionString))
                        {
                            connection.Open();

                            // 1. فحص هل يوجد حجوزات ومواعيد قادمة لم يحن وقتها بعد
                            bool hasFutureAppointments = false;
                            string checkSql = "SELECT AppDate, AppTime FROM Appointments WHERE PatientID=@ID AND Status != 'ملغي'";

                            using (var checkCmd = new SQLiteCommand(checkSql, connection))
                            {
                                checkCmd.Parameters.AddWithValue("@ID", txtID.Text);
                                using (var reader = checkCmd.ExecuteReader())
                                {
                                    while (reader.Read())
                                    {
                                        string appDateStr = reader["AppDate"].ToString();
                                        string appTimeStr = reader["AppTime"].ToString();

                                        // محاولة دمج التاريخ والوقت للفحص بدقة
                                        if (DateTime.TryParse($"{appDateStr} {appTimeStr}", out DateTime appDateTime))
                                        {
                                            if (appDateTime > DateTime.Now)
                                            {
                                                hasFutureAppointments = true;
                                                break;
                                            }
                                        }
                                        else if (DateTime.TryParse(appDateStr, out DateTime justDate))
                                        {
                                            // فحص احتياطي في حال كان تنسيق الوقت غير مقروء
                                            if (justDate.Date > DateTime.Today)
                                            {
                                                hasFutureAppointments = true;
                                                break;
                                            }
                                        }
                                    }
                                }
                            }

                            // إذا كان هناك حجز قادم نمنع الحذف تماماً
                            if (hasFutureAppointments)
                            {
                                MessageBox.Show("لا يمكن حذف هذا المريض الآن!\n\nلديه حجوزات قادمة لم يحن موعدها بعد.\nيجب انتظار مرور وقت الحجز أو قيام السكرتير بإلغائه من شاشة المواعيد أولاً.", "تنبيه النظام", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                                return; // خروج وإلغاء عملية الحذف
                            }

                            // 2. إذا كانت الحجوزات كلها في الماضي، نربطها بمريض محذوف لكي لا تتضرر الحسابات
                            string sqlUpdateAppointments = "UPDATE Appointments SET PatientID = NULL, PatientName = 'مريض محذوف' WHERE PatientID=@ID";
                            using (var commandApp = new SQLiteCommand(sqlUpdateAppointments, connection))
                            {
                                commandApp.Parameters.AddWithValue("@ID", txtID.Text);
                                commandApp.ExecuteNonQuery();
                            }

                            // 3. مسح ملف المريض الأساسي من قاعدة البيانات
                            string sqlDeletePatient = "DELETE FROM Patients WHERE ID=@ID";
                            using (var commandPatient = new SQLiteCommand(sqlDeletePatient, connection))
                            {
                                commandPatient.Parameters.AddWithValue("@ID", txtID.Text);
                                commandPatient.ExecuteNonQuery();
                            }
                        }

                        Logger.LogAction("إدارة المرضى", $"قام بمسح المريض ({txtName.Text}) وتحويل حجوزاته السابقة إلى (مريض محذوف)");

                        MessageBox.Show("تم حذف ملف المريض بنجاح، وتم الاحتفاظ بحجوزاته السابقة في الخزنة كـ (مريض محذوف).", "تم الحذف", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        LoadData();
                        ClearInputs();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("حدث خطأ أثناء محاولة مسح البيانات: " + ex.Message, "خطأ تقني", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private bool ValidateInputs()
        {
            if (string.IsNullOrWhiteSpace(txtName.Text) || string.IsNullOrWhiteSpace(txtAge.Text))
            {
                if (string.IsNullOrWhiteSpace(txtName.Text)) lblNameError.Visible = true;
                if (string.IsNullOrWhiteSpace(txtAge.Text)) lblAgeError.Visible = true;
                if (string.IsNullOrWhiteSpace(txtName.Text)) txtName.Focus(); else txtAge.Focus();
                return false;
            }
            if (!string.IsNullOrWhiteSpace(txtPhone.Text) && txtPhone.Text.Length != 11)
            {
                MessageBox.Show("يجب أن يتكون رقم الهاتف من 11 رقم بالضبط!", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtPhone.Focus();
                return false;
            }
            return true;
        }

        private void ClearInputs()
        {
            txtID.Clear(); txtName.Clear(); txtPhone.Clear(); txtAge.Clear();
            txtChronic.Clear(); txtMeds.Clear();
            txtDebt.Text = "0";
            cmbBloodType.SelectedIndex = 0;
            rbMale.Checked = true;
            lblNameError.Visible = false;
            lblAgeError.Visible = false;
        }

        private void UpdateGenderSelection()
        {
            if (rbMale.Checked)
            {
                lblMaleText.Font = new Font(FontFamily.GenericSansSerif, 11, FontStyle.Bold); lblMaleText.ForeColor = Color.FromArgb(0, 120, 215);
                lblFemaleText.Font = new Font(FontFamily.GenericSansSerif, 11, FontStyle.Regular); lblFemaleText.ForeColor = Color.Black;
            }
            else
            {
                lblFemaleText.Font = new Font(FontFamily.GenericSansSerif, 11, FontStyle.Bold); lblFemaleText.ForeColor = Color.FromArgb(0, 120, 215);
                lblMaleText.Font = new Font(FontFamily.GenericSansSerif, 11, FontStyle.Regular); lblMaleText.ForeColor = Color.Black;
            }
        }

        private void dgvPatients_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
        }
    }
}
