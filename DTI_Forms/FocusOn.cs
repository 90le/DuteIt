using System;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;

namespace DuteIT
{
    public partial class FocusOn : Form
    {
        public FocusOn()
        {
            InitializeComponent();
        }

        private bool m_isMouseDown = false;
        private Point m_mousePos = new Point();
        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            m_mousePos = Cursor.Position;
            m_isMouseDown = true;
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);
            m_isMouseDown = false;
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            if (m_isMouseDown)
            {
                Point tempPos = Cursor.Position;
                this.Location = new Point(Location.X + (tempPos.X - m_mousePos.X), Location.Y + (tempPos.Y - m_mousePos.Y));
                m_mousePos = Cursor.Position;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FocusOn.ActiveForm.Close();
        }

        private void label1_Click(object sender, EventArgs e)
        {
            // 跳转到网页
            string url = "https://90le.cn";  // 将此URL替换为您要跳转的网页
            try
            {
                Process.Start(url);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"无法打开浏览器。\n请手动前往：https://90le.cn", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void OK_FocusOn_Load(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {
            // 跳转到网页
            string url = "https://90le.cn/api/DTI_Tool/index.html";  // 将此URL替换为您要跳转的网页
            try
            {
                Process.Start(url);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"无法打开浏览器。\n请手动前往：https://90le.cn/api/DTI_Tool/index.html", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
