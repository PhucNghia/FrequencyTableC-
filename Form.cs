using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CS_DL_DPT
{
    public partial class Form : System.Windows.Forms.Form
    {
        public Form()
        {
            InitializeComponent();
        }

        private void FormCSDLDPT_Load(object sender, EventArgs e)
        {
            setRowColumHeader(dgvIn, 6, 6);
            setRowColumHeader(dgvOut1, 6, 6);
            setRowColumHeader(dgvOut2, 6, 6);
            setValueToCells(dgvIn, input());
        }

        // Fake data
        private double[][] input()
        {
            double[][] array = new double[5][];
            for (int i = 0; i < 5; i++)
                array[i] = new double[6];

            array[0][0] = 615;  array[0][1] = 390;  array[0][2] = 10;
            array[0][3] = 10;   array[0][4] = 18;   array[0][5] = 65;

            array[1][0] = 15;   array[1][1] = 4;    array[1][2] = 76;
            array[1][3] = 217;  array[1][4] = 91;   array[1][5] = 816;

            array[2][0] = 2;    array[2][1] = 8;    array[2][2] = 815;
            array[2][3] = 142;  array[2][4] = 765;  array[2][5] = 1;

            array[3][0] = 312;  array[3][1] = 511;  array[3][2] = 677;
            array[3][3] = 11;   array[3][4] = 711;  array[3][5] = 2;

            array[4][0] = 45;   array[4][1] = 33;   array[4][2] = 516;
            array[4][3] = 64;   array[4][4] = 491;  array[4][5] = 59;

            return array;
        }
        
        //============================ KHỞI TẠO BẢNG ===============================

        // Button input row and column
        private void btnInRowCol_Click(object sender, EventArgs e)
        {
            int row = 0, col = 0;
            int check = 1;
            try
            {
                row = int.Parse(txtRow.Text);
                col = int.Parse(txtCol.Text);
                if(row <= 0 || col <= 0)
                    check = -1;
                if (row > 5000 || col > 1000)
                    check = 0;
            }
            catch
            {
                check = -1;
            }
            if (check == 1)
            {
                resetDgv(dgvIn);
                resetDgv(dgvOut1);
                resetDgv(dgvOut2);
                setRowColumHeader(dgvIn, row, col);
                setRowColumHeader(dgvOut1, row, col);
                setRowColumHeader(dgvOut2, row, col);
            }
            else if(check == -1)
                MessageBox.Show("Hàng và cột phải là số nguyên dương > 0");
            else
                MessageBox.Show("Giới hạn hàng <= 5000 và cột <= 1000");
        }

        // Format Row and Column header
        private void setRowColumHeader(DataGridView dgv, int row, int col)
        {
            // Add column header
            try { dgv.ColumnCount = col; }
            catch
            {
                MessageBox.Show((dgv.ColumnCount - 1).ToString() + "  row: " + (dgv.RowCount - 1).ToString());
            }
            for (int i = 0; i < col; i++)
            {
                dgv.Columns[i].Width = 75;
                dgv.Columns[i].Name = "D" + (i + 1).ToString();
            }

            // Add row header
            dgv.RowCount = row;
            for (int i = 0; i < row; i++)
            {
                dgv.Rows[i].Height = 30;
                dgv.Rows[i].HeaderCell.Value = "T" + (i + 1).ToString();
            }
            dgv.AllowUserToAddRows = false;
        }
        
        //===================== XỬ LÝ GET - SET DL TRONG BẢNG ======================

        // Set value to Cells
        private void setValueToCells(DataGridView dgv, double[][] array)
        {
            int row = dgv.RowCount;
            int column = dgv.ColumnCount;
            for (int i = 0; i < row; i++)
                for (int j = 0; j < column; j++)
                    dgv.Rows[i].Cells[j].Value = array[i][j].ToString();
        }

        // Khởi tạo những dòng và cột bị lỗi
        int[] rowError;
        int[] colError;
        int indexError;

        // Get value for Cells
        private double[][] getValueForCells(DataGridView dgv, out int flagError)
        {
            bool check = true;
            int row, column;
            row = dgv.RowCount;
            column = dgv.ColumnCount;
            double[][] kq = new double[row][];
            for (int i = 0; i < row; i++)
                kq[i] = new double[column];

            if (row <= 0 || column <= 0)
            {
                flagError = -1;
                return null;
            }
            
            rowError = new int[row * column];
            colError = new int[column * column];
            indexError = 0;
            for (int i = 0; i < row; i++)
                for (int j = 0; j < column; j++)
                {
                    try
                    {
                        dgvIn.Rows[i].Cells[j].Style.ForeColor = Color.Black;
                        dgvIn.Rows[i].Cells[j].Style.BackColor = Color.Ivory;
                        string s = dgv.Rows[i].Cells[j].Value.ToString();
                        kq[i][j] = double.Parse(s);
                        if (kq[i][j] < 0)
                        {
                            check = false;
                            rowError[indexError] = i; colError[indexError] = j;
                            indexError++;
                        }
                    }
                    catch
                    {
                        check = false;
                        rowError[indexError] = i; colError[indexError] = j;
                        indexError++;
                    }
                }
            if (!check)
            {
                flagError = 1;
                return null;
            }
            flagError = 0;
            return kq;
        }

        //===================== XỬ LÝ BUTTON ======================

        // Button submit
        private void btnSubmit_Click(object sender, EventArgs e)
        {
            
            int row = dgvIn.RowCount;
            int column = dgvIn.ColumnCount;
            int flagError;
            double[][] dlVao = getValueForCells(dgvIn, out flagError);

            if (dlVao != null)
            {
                double[][] kqChuanHoa = chuanHoa(dlVao);
                double[][] vectorTrongSo = timVetorTrongSo(kqChuanHoa);

                setValueToCells(dgvOut1, kqChuanHoa);
                setValueToCells(dgvOut2, vectorTrongSo);
            }
            else
            {
                if (flagError == 1)
                {
                    for (int i = 0; i < indexError; i++)
                    {
                        dgvIn.Rows[rowError[i]].Cells[colError[i]].Style.BackColor = Color.IndianRed;
                        dgvIn.Rows[rowError[i]].Cells[colError[i]].Style.ForeColor = Color.White;
                    }
                    MessageBox.Show("Bạn phải nhập số dương vào các ô sai");
                }
                if(flagError == -1)
                {
                    MessageBox.Show("Số dòng và số cột phải lớn hơn 0");
                }
            }
        }

        // Button reset
        private void btnReset_Click(object sender, EventArgs e)
        {
            resetDgv(dgvIn);
        }

        // reset dgv
        private void resetDgv(DataGridView dgv)
        {
            int row = dgv.RowCount;
            int column = dgv.ColumnCount;
            for (int i = 0; i < row; i++)
                for (int j = 0; j < column; j++)
                {
                    dgv.Rows[i].Cells[j].Value = "";
                    dgv.Rows[i].Cells[j].Style.BackColor = Color.Ivory;
                }
        }

        //================= IINSERT, REMOVE - COLUMN, ROW ====================

        // Insert Column
        private void insertColumn(DataGridView dgv, int position)
        {
            int col = dgvIn.CurrentCell.ColumnIndex;    // Phải lấy col ở bảng dgvIn
            DataGridViewTextBoxColumn clm = new DataGridViewTextBoxColumn();
            clm.HeaderText = "D" + dgv.ColumnCount;
            clm.Width = 75;
            clm.SortMode = DataGridViewColumnSortMode.Programmatic;
            if (position == -1)     // Insert Column to left
                dgv.Columns.Insert(col, clm);
            else                    // Insert Column to right
                dgv.Columns.Insert(col + 1, clm);
        }

        // Insert Row
        private void insertRow(DataGridView dgv, int position)
        {
            int row = dgvIn.CurrentCell.RowIndex;
            DataGridViewRow rowClone = (DataGridViewRow)dgv.Rows[0].Clone();
            rowClone.HeaderCell.Value = "T" + dgv.RowCount;
            if (position == 1)   // Insert Row to top
                dgv.Rows.Insert(row, rowClone);
            else                // Insert Row to bottom
                dgv.Rows.Insert(row + 1, rowClone);
        }

        // Remove Column
        private void removeColumn(DataGridView dgv)
        {
            int col = dgvIn.CurrentCell.ColumnIndex;
            dgv.Columns.RemoveAt(col);
        }

        // Remove Row
        private void removeRow(DataGridView dgv)
        {
            int row = dgvIn.CurrentCell.RowIndex;
            dgv.Rows.RemoveAt(row);
        }

        // Sự kiện chuột phải
        private void dgvIn_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                int currentRow = dgvIn.HitTest(e.X, e.Y).RowIndex;
                int currentColumn = dgvIn.HitTest(e.X, e.Y).ColumnIndex;

                if (currentRow >= 0 && currentColumn >= 0)
                {
                    dgvIn.CurrentCell = dgvIn.Rows[currentRow].Cells[currentColumn];

                    ContextMenuStrip m = new ContextMenuStrip();
                    m.Items.Add("Thêm cột trước").Name = "insert_col_before";
                    m.Items.Add("Thêm cột sau").Name = "insert_col_after";
                    m.Items.Add("Thêm dòng trước").Name = "insert_row_before";
                    m.Items.Add("Thêm dòng sau").Name = "insert_row_after";
                    m.Items.Add("Xóa cột").Name = "remove_col";
                    m.Items.Add("Xóa dòng").Name = "remove_row";

                    m.Show(dgvIn, new Point(e.X, e.Y));
                    m.ItemClicked += new ToolStripItemClickedEventHandler(menuItemClicked);
                }
            }
        }

        // Xử lý khi nhấn danh sách tùy chọn chuột phải
        private void menuItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            switch (e.ClickedItem.Name)
            {
                case "insert_col_before":
                    insertColumn(dgvIn, -1);
                    insertColumn(dgvOut1, -1);
                    insertColumn(dgvOut2, -1);
                    break;
                case "insert_col_after":
                    insertColumn(dgvIn, 1);
                    insertColumn(dgvOut1, 1);
                    insertColumn(dgvOut2, 1);
                    break;
                case "insert_row_before":
                    insertRow(dgvIn, 1);
                    insertRow(dgvOut1, 1);
                    insertRow(dgvOut2, 1);
                    break;
                case "insert_row_after":
                    insertRow(dgvIn, -1);
                    insertRow(dgvOut1, -1);
                    insertRow(dgvOut2, -1);
                    break;
                case "remove_col":
                    removeColumn(dgvIn);
                    removeColumn(dgvOut1);
                    removeColumn(dgvOut2);
                    break;
                default:
                    removeRow(dgvIn);
                    removeRow(dgvOut1);
                    removeRow(dgvOut2);
                    break;
            }
        }

        //===================== CHUẨN HÓA VÀ TÌM VECTOR TRỌNG SỐ ======================

        // Chuẩn hóa bảng tần số
        private double[][] chuanHoa(double[][] data)
        {
            int row = data.Length;
            int column = data[0].Length;
            double[] tanSo = new double[column];
            double[][] kqChuanHoa = new double[row][];
            for (int i = 0; i < row; i++)
                kqChuanHoa[i] = new double[column];

            for (int i = 0; i < column; i++)
            {
                double idf =0;
                for (int j = 0; j < row; j++)
                    idf += data[j][i]; 
                tanSo[i] = + idf;
            }

            for (int i = 0; i < column; i++)
                for (int j = 0; j < row; j++)
                    kqChuanHoa[j][i] = Math.Round(data[j][i] / tanSo[i], 2);

            return kqChuanHoa;
        }

        // Tìm vector trọng số
        private double[][] timVetorTrongSo(double[][] data)
        {
            int row = data.Length;
            int column = data[0].Length;
            double[] doQuanTrong = new double[row];
            double[][] kq = new double[row][];
            for (int i = 0; i < row; i++)
                kq[i] = new double[column];

            for (int i = 0; i < row; i++)
            {
                int dem = 0;
                for (int j = 0; j < column; j++)
                    if(data[i][j] == 0)
                        dem++;
                doQuanTrong[i] = Math.Round(Math.Log10((float)column / (column - dem)), 2);
            }
            for (int i = 0; i < row; i++)
                for (int j = 0; j < column; j++)
                    kq[i][j] = Math.Round(data[i][j] * doQuanTrong[i], 2);

            return kq;
        }

        private void dgvIn_ColumnAdded(object sender, DataGridViewColumnEventArgs e)
        {
            e.Column.FillWeight = 10;
        }

        private void dgvOut1_ColumnAdded(object sender, DataGridViewColumnEventArgs e)
        {
            e.Column.FillWeight = 10;
        }

        private void dgvOut2_ColumnAdded(object sender, DataGridViewColumnEventArgs e)
        {
            e.Column.FillWeight = 10;
        }
    }
}
