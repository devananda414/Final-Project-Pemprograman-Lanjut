using System;
using System.Drawing;
using System.Drawing.Printing;
using System.Globalization;
using System.Windows.Forms;

namespace FinalProject
{
    public partial class Form1 : Form
    {
        // Tarif per jam (silakan ubah sesuai kebutuhan)
        private const int TarifMotor = 2000;
        private const int TarifMobil = 5000;

        public Form1()
        {
            InitializeComponent();

            // Hubungkan event (tidak perlu klik dua kali di desain)
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

            listBox1.Items.Clear();
            listBox1.Items.Add("Petugas | No Tiket | Plat | Tanggal | In | Out | Jenis | Total");
        }

        // ===== Reset isian form =====
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

        // ===== Hitung total biaya parkir =====
        private void HitungTotal()
        {
            DateTime masuk, keluar;

            if (!DateTime.TryParseExact(txtCheckIn.Text.Trim(), "HH:mm",
                    CultureInfo.InvariantCulture, DateTimeStyles.None, out masuk) ||
                !DateTime.TryParseExact(txtCheckOut.Text.Trim(), "HH:mm",
                    CultureInfo.InvariantCulture, DateTimeStyles.None, out keluar) ||
                (!chkMotor.Checked && !chkMobil.Checked))
            {
                txtTotal.Clear();
                return;
            }

            TimeSpan durasi = keluar - masuk;
            if (durasi.TotalMinutes < 0)
                durasi = durasi.Add(TimeSpan.FromHours(24)); // lewat tengah malam

            int jam = (int)Math.Ceiling(durasi.TotalHours);
            if (jam < 1) jam = 1; // minimal 1 jam

            int tarif = chkMotor.Checked ? TarifMotor : TarifMobil;
            txtTotal.Text = (jam * tarif).ToString("N0");
        }

        // ===== Jenis kendaraan: hanya boleh pilih salah satu =====
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

        // ===== Tombol Simpan =====
        private void btnSimpan_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtPetugas.Text) ||
                string.IsNullOrWhiteSpace(txtPlat.Text) ||
                string.IsNullOrWhiteSpace(txtCheckOut.Text))
            {
                MessageBox.Show("Petugas, Plat Nomor, dan Check Out harus diisi!",
                    "Peringatan", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!chkMotor.Checked && !chkMobil.Checked)
            {
                MessageBox.Show("Pilih jenis kendaraan (Motor / Mobil)!",
                    "Peringatan", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrEmpty(txtTotal.Text))
            {
                MessageBox.Show("Format jam harus HH:mm (contoh 08:30).",
                    "Peringatan", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string jenis = chkMotor.Checked ? "Motor" : "Mobil";

            listBox1.Items.Add(
                txtPetugas.Text + " | " +
                txtNoTiket.Text + " | " +
                txtPlat.Text.ToUpper() + " | " +
                txtTanggal.Text + " | " +
                txtCheckIn.Text + " | " +
                txtCheckOut.Text + " | " +
                jenis + " | Rp " +
                txtTotal.Text);

            MessageBox.Show("Data berhasil disimpan.", "Informasi",
                MessageBoxButtons.OK, MessageBoxIcon.Information);

            BersihkanForm();
        }

        // ===== Tombol Batal =====
        private void btnBatal_Click(object sender, EventArgs e)
        {
            BersihkanForm();
        }

        // ===== Tombol Cetak =====
        private void btnCetak_Click(object sender, EventArgs e)
        {
            if (listBox1.Items.Count <= 1)
            {
                MessageBox.Show("Belum ada data untuk dicetak.", "Informasi",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            PrintDocument pd = new PrintDocument();
            pd.PrintPage += Pd_PrintPage;

            PrintPreviewDialog preview = new PrintPreviewDialog();
            preview.Document = pd;
            preview.ShowDialog();
        }

        private void Pd_PrintPage(object sender, PrintPageEventArgs e)
        {
            Font judul = new Font("Arial", 14, FontStyle.Bold);
            Font isi = new Font("Arial", 10);
            float y = 50;

            e.Graphics.DrawString("LAPORAN PARKIR", judul, Brushes.Black, 50, y);
            y += 40;

            foreach (var item in listBox1.Items)
            {
                e.Graphics.DrawString(item.ToString(), isi, Brushes.Black, 50, y);
                y += 25;
            }
        }

        // ===== Tombol Keluar =====
        private void btnKeluar_Click(object sender, EventArgs e)
        {
            DialogResult hasil = MessageBox.Show("Yakin ingin keluar?", "Konfirmasi",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (hasil == DialogResult.Yes)
                Application.Exit();
        }
    }
}