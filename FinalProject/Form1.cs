using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Printing;
using System.Globalization;
using System.Windows.Forms;

namespace FinalProject
{
    public partial class Form1 : Form
    {
        private const int TarifMotor = 2000;
        private const int TarifMobil = 5000;

        private class DataParkir
        {
            public string Petugas, NoTiket, Plat, Tanggal, CheckIn, CheckOut, Jenis;
            public int Total;
        }

        private List<DataParkir> daftarParkir = new List<DataParkir>();
        private DataGridView dgvParkir;

        public Form1()
        {
            InitializeComponent();
            this.Load += Form1_Load;
            btnSimpan.Click += btnSimpan_Click;
            btnBatal.Click += btnBatal_Click;
            btnCetak.Click += btnCetak_Click;
            btnKeluar.Click += btnKeluar_Click;
            chkMotor.CheckedChanged += chkMotor_CheckedChanged;
            chkMobil.CheckedChanged += chkMobil_CheckedChanged;
            txtCheckIn.TextChanged += txtCheckIn_TextChanged;
            txtCheckOut.TextChanged += txtCheckOut_TextChanged;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            BersihkanForm();
            txtTotal.ReadOnly = true;
            txtTanggal.ReadOnly = true;
            BuatTabelExcel();
        }

        private void BuatTabelExcel()
        {
            listBox1.Visible = false;

            dgvParkir = new DataGridView();
            dgvParkir.Location = listBox1.Location;
            dgvParkir.Size = listBox1.Size;
            dgvParkir.Anchor = listBox1.Anchor;
            dgvParkir.BackgroundColor = Color.White;
            dgvParkir.BorderStyle = BorderStyle.FixedSingle;
            dgvParkir.GridColor = Color.LightGray;
            dgvParkir.AllowUserToAddRows = false;
            dgvParkir.AllowUserToDeleteRows = false;
            dgvParkir.AllowUserToResizeRows = false;
            dgvParkir.ReadOnly = true;
            dgvParkir.RowHeadersVisible = false;
            dgvParkir.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgvParkir.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgvParkir.EnableHeadersVisualStyles = false;

            dgvParkir.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(225, 230, 235);
            dgvParkir.ColumnHeadersDefaultCellStyle.ForeColor = Color.Black;
            dgvParkir.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9f, FontStyle.Bold);
            dgvParkir.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvParkir.ColumnHeadersHeight = 30;
            dgvParkir.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;

            dgvParkir.DefaultCellStyle.Font = new Font("Segoe UI", 9f, FontStyle.Regular);
            dgvParkir.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvParkir.DefaultCellStyle.SelectionBackColor = Color.FromArgb(204, 228, 247);
            dgvParkir.DefaultCellStyle.SelectionForeColor = Color.Black;
            dgvParkir.RowTemplate.Height = 26;

            dgvParkir.Columns.Add("No", "No");
            dgvParkir.Columns.Add("Petugas", "Petugas");
            dgvParkir.Columns.Add("NoTiket", "No Tiket");
            dgvParkir.Columns.Add("Plat", "Plat Nomor");
            dgvParkir.Columns.Add("Tanggal", "Tanggal");
            dgvParkir.Columns.Add("In", "In");
            dgvParkir.Columns.Add("Out", "Out");
            dgvParkir.Columns.Add("Jenis", "Jenis");
            dgvParkir.Columns.Add("Total", "Total");

            dgvParkir.Columns["No"].FillWeight = 40;
            dgvParkir.Columns["NoTiket"].FillWeight = 135;
            dgvParkir.Columns["In"].FillWeight = 55;
            dgvParkir.Columns["Out"].FillWeight = 55;
            dgvParkir.Columns["Jenis"].FillWeight = 65;

            this.Controls.Add(dgvParkir);
            dgvParkir.BringToFront();
        }

        private void BersihkanForm()
        {
            txtPetugas.Clear();
            txtNoTiket.Text = "TKT" + DateTime.Now.ToString("yyMMddHHmmss");
            txtPlat.Clear();
            txtCheckIn.Text = DateTime.Now.ToString("HH:mm");
            txtCheckOut.Clear();
            txtTanggal.Text = DateTime.Now.ToString("dd/MM/yyyy");
            chkMotor.Checked = false;
            chkMobil.Checked = false;
            txtTotal.Clear();
            txtPetugas.Focus();
        }

        private void HitungTotal()
        {
            DateTime masuk, keluar;
            var ci = CultureInfo.InvariantCulture;
            var style = DateTimeStyles.None;

            bool okMasuk = DateTime.TryParseExact(txtCheckIn.Text.Trim(), "HH:mm", ci, style, out masuk);
            bool okKeluar = DateTime.TryParseExact(txtCheckOut.Text.Trim(), "HH:mm", ci, style, out keluar);

            if (!okMasuk || !okKeluar || (!chkMotor.Checked && !chkMobil.Checked))
            {
                txtTotal.Clear();
                return;
            }

            TimeSpan durasi = keluar - masuk;
            if (durasi.TotalMinutes < 0)
                durasi = durasi.Add(TimeSpan.FromHours(24));

            int jam = (int)Math.Ceiling(durasi.TotalHours);
            if (jam < 1) jam = 1;

            int tarif = chkMotor.Checked ? TarifMotor : TarifMobil;
            txtTotal.Text = (jam * tarif).ToString("N0");
        }

        private void chkMotor_CheckedChanged(object sender, EventArgs e)
        {
            if (chkMotor.Checked) chkMobil.Checked = false;
            HitungTotal();
        }

        private void chkMobil_CheckedChanged(object sender, EventArgs e)
        {
            if (chkMobil.Checked) chkMotor.Checked = false;
            HitungTotal();
        }

        private void txtCheckIn_TextChanged(object sender, EventArgs e)
        {
            HitungTotal();
        }

        private void txtCheckOut_TextChanged(object sender, EventArgs e)
        {
            HitungTotal();
        }

        private void btnSimpan_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtPetugas.Text) ||
                string.IsNullOrWhiteSpace(txtPlat.Text) ||
                string.IsNullOrWhiteSpace(txtCheckOut.Text))
            {
                MessageBox.Show("Petugas, Plat Nomor, dan Check Out harus diisi!", "Peringatan", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!chkMotor.Checked && !chkMobil.Checked)
            {
                MessageBox.Show("Pilih jenis kendaraan (Motor / Mobil)!", "Peringatan", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrEmpty(txtTotal.Text))
            {
                MessageBox.Show("Format jam harus HH:mm (contoh 08:30).", "Peringatan", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string jenis = chkMotor.Checked ? "Motor" : "Mobil";
            int nominal = int.Parse(txtTotal.Text.Replace(".", "").Replace(",", ""));

            DataParkir d = new DataParkir
            {
                Petugas = txtPetugas.Text.Trim(),
                NoTiket = txtNoTiket.Text.Trim(),
                Plat = txtPlat.Text.Trim().ToUpper(),
                Tanggal = txtTanggal.Text.Trim(),
                CheckIn = txtCheckIn.Text.Trim(),
                CheckOut = txtCheckOut.Text.Trim(),
                Jenis = jenis,
                Total = nominal
            };
            daftarParkir.Add(d);

            dgvParkir.Rows.Add(
                daftarParkir.Count,
                d.Petugas,
                d.NoTiket,
                d.Plat,
                d.Tanggal,
                d.CheckIn,
                d.CheckOut,
                d.Jenis,
                "Rp " + d.Total.ToString("N0")
            );

            MessageBox.Show("Data berhasil disimpan.", "Informasi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            BersihkanForm();
        }

        private void btnBatal_Click(object sender, EventArgs e)
        {
            BersihkanForm();
        }

        private void btnCetak_Click(object sender, EventArgs e)
        {
            if (daftarParkir.Count == 0)
            {
                MessageBox.Show("Belum ada data untuk dicetak.", "Informasi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            PrintDocument pd = new PrintDocument();
            pd.PrintPage += Pd_PrintPage;

            PrintPreviewDialog preview = new PrintPreviewDialog();
            preview.Document = pd;
            preview.Width = 850;
            preview.Height = 650;
            preview.ShowDialog();
        }

        private void Pd_PrintPage(object sender, PrintPageEventArgs e)
        {
            Graphics g = e.Graphics;
            Font fJudul = new Font("Arial", 16, FontStyle.Bold);
            Font fSub = new Font("Arial", 10, FontStyle.Regular);
            Font fHead = new Font("Arial", 9, FontStyle.Bold);
            Font fIsi = new Font("Arial", 9, FontStyle.Regular);

            Pen pen = new Pen(Color.Black, 1);
            Brush bHead = new SolidBrush(Color.FromArgb(225, 230, 235));

            string[] headers = { "No", "Petugas", "No Tiket", "Plat Nomor", "Tanggal", "In", "Out", "Jenis", "Total Biaya" };
            int[] widths = { 35, 85, 125, 95, 80, 50, 50, 65, 110 };
            int h = 26;

            int totalTableWidth = 0;
            for (int i = 0; i < widths.Length; i++) totalTableWidth += widths[i];

            float pageWidth = e.PageBounds.Width;
            float startX = (pageWidth - totalTableWidth) / 2f;
            float y = 50;

            StringFormat sfC = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };

            RectangleF rJudul = new RectangleF(0, y, pageWidth, 28);
            g.DrawString("LAPORAN TRANSAKSI PARKIR", fJudul, Brushes.Black, rJudul, sfC);
            y += 26;

            RectangleF rSub = new RectangleF(0, y, pageWidth, 20);
            g.DrawString("Dicetak pada: " + DateTime.Now.ToString("dd/MM/yyyy HH:mm"), fSub, Brushes.DimGray, rSub, sfC);
            y += 32;

            float curX = startX;
            for (int i = 0; i < headers.Length; i++)
            {
                RectangleF r = new RectangleF(curX, y, widths[i], h);
                g.FillRectangle(bHead, r);
                g.DrawRectangle(pen, curX, y, widths[i], h);
                g.DrawString(headers[i], fHead, Brushes.Black, r, sfC);
                curX += widths[i];
            }
            y += h;

            int no = 1;
            int grandTotal = 0;

            foreach (var item in daftarParkir)
            {
                curX = startX;
                string[] row = { no.ToString(), item.Petugas, item.NoTiket, item.Plat, item.Tanggal, item.CheckIn, item.CheckOut, item.Jenis, "Rp " + item.Total.ToString("N0") };

                for (int i = 0; i < row.Length; i++)
                {
                    g.DrawRectangle(pen, curX, y, widths[i], h);
                    RectangleF tr = new RectangleF(curX + 2, y, widths[i] - 4, h);
                    g.DrawString(row[i], fIsi, Brushes.Black, tr, sfC);
                    curX += widths[i];
                }

                grandTotal += item.Total;
                no++;
                y += h;
            }

            int wTotalLabel = totalTableWidth - widths[widths.Length - 1];

            RectangleF rLbl = new RectangleF(startX, y, wTotalLabel, h);
            g.FillRectangle(bHead, rLbl);
            g.DrawRectangle(pen, startX, y, wTotalLabel, h);
            g.DrawString("TOTAL PENDAPATAN", fHead, Brushes.Black, rLbl, sfC);

            float xNom = startX + wTotalLabel;
            int wNom = widths[widths.Length - 1];
            RectangleF rVal = new RectangleF(xNom, y, wNom, h);
            g.FillRectangle(bHead, rVal);
            g.DrawRectangle(pen, xNom, y, wNom, h);
            g.DrawString("Rp " + grandTotal.ToString("N0"), fHead, Brushes.Black, rVal, sfC);
        }

        private void btnKeluar_Click(object sender, EventArgs e)
        {
            DialogResult hasil = MessageBox.Show("Yakin ingin keluar?", "Konfirmasi", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (hasil == DialogResult.Yes) Application.Exit();
        }
    }
}