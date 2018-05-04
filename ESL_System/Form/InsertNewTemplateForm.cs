﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using K12.Data;
using DevComponents.DotNetBar;
using FISCA.Data;
using System.Xml;
using System.Xml.Linq;
using System.IO;

namespace ESL_System.Form
{   
    public partial class InsertNewTemplateForm : BaseForm
    {
        public class Item
        {
            public string Name;
            public string Value;
            public Item(string name, string value)
            {
                Name = name; Value = value;
            }
            public override string ToString()
            {
                // Generates the text shown in the combo box
                return Name;
            }
            public string GetDescriptionString()
            {
                // Generates the text shown in the combo box
                return Value;
            }
        }

        public InsertNewTemplateForm()
        {
            InitializeComponent();

            txtTemplateName.Text = "請輸入新ESL 樣板名稱";

            // 2018/05/01 穎驊重要備註， 在table exam_template 欄位 description 不為空代表其為ESL 的樣板
            string query = "select * from exam_template where description !=''";

            QueryHelper qh = new QueryHelper();
            DataTable dt = qh.Select(query);

            if (dt.Rows.Count > 0)
            {
                foreach (DataRow dr in dt.Rows)
                {                                                            
                    cboExistTemplates.Items.Add(new Item("" + dr[1], "" + dr[5])); // dr[5] 為description 內容
                }
            }
                            
            cboExistTemplates.SelectedIndex = 0; //預設選不複製
        }

        private void btnConfirm_Click(object sender, EventArgs e)
        {
            if (txtTemplateName.Text =="")
            {
                MsgBox.Show("請輸入ESL 樣板名稱");

                return;
            }

            string desciption = "";

            if( cboExistTemplates.SelectedIndex != 0)// 不是選第一個 "不複製"
            {
                Item i = (Item) cboExistTemplates.SelectedItem; // 將 Object SelectedItem 轉型成 Item 處理

                desciption = "" + i.GetDescriptionString();
            }
            else
            {
                //取得預設的樣板資料
                XDocument doc = XDocument.Parse(Properties.Resources.Description_Example);

                desciption = doc.ToString();
            }

            UpdateHelper uh = new UpdateHelper();

            //依照所選項目新增 (allow_upload 此項固定為 0 且型別 為 bit)
            string updQuery = "INSERT INTO exam_template (name, allow_upload, description) VALUES('"+ txtTemplateName.Text +"',0::bit,'"+ desciption +"')";

            //執行sql，更新
            uh.Execute(updQuery);

            MsgBox.Show("新增樣板成功");

            DialogResult = DialogResult.OK;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }
    }
}
