using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net;
using System.IO;
using System.Web;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using OfficeOpenXml;
using System.Threading;

namespace HeHe
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            int page = 0;
            string put2 = textBox3.Text;
            string put1 = textBox2.Text;
            ExcelPackage package = new ExcelPackage(new FileInfo(put1));
            ExcelWorksheet sheet = package.Workbook.Worksheets[1];
            ExcelPackage package2 = new ExcelPackage(new FileInfo(put2));
            ExcelWorksheet sheet2 = package2.Workbook.Worksheets[1];
            int kol = sheet2.Dimension.Rows;
            string[] login = new string[kol];
            string[] password = new string[kol];
            int chet=0;
            for (int z = 2; z < kol; z++)
            {
                CookieContainer cookies = new CookieContainer();
                int k, m, porog=0;
                login[z] = (string)(sheet2.Cells[z, 1]).Value;
                password[z] = Convert.ToString((sheet2.Cells[z, 2]).Value);
                HttpWebRequest request1 = (HttpWebRequest)WebRequest.Create("https://nn.hh.ru/");
                request1.CookieContainer = cookies;
                request1.Headers["Upgrade-Insecure-Requests"] = "1";
                request1.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8";
                request1.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/65.0.3325.181 Safari/537.36";
                HttpWebResponse response1 = (HttpWebResponse)request1.GetResponse();
                StreamReader streamReader = new StreamReader(response1.GetResponseStream());
                string result = streamReader.ReadToEnd();
                string xsrf = Regex.Match(result, "(?<==\"_xsrf\" value=\").*?(?=\"><)").Value;
                request1 = (HttpWebRequest)WebRequest.Create("https://nn.hh.ru/account/login?backurl=%2F");
                request1.Method = "POST";
                request1.CookieContainer = cookies;
                request1.Headers["Upgrade-Insecure-Requests"] = "1";
                request1.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8";
                request1.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36";
                request1.ContentType = "application/x-www-form-urlencoded";
                string data = "username=" + login[z] + "&password=" + password[z] + "&backUrl=https%3A%2F%2Fnn.hh.ru%2F&action=%D0%92%D0%BE%D0%B9%D1%82%D0%B8&_xsrf=" + xsrf;
                byte[] bytedata = Encoding.ASCII.GetBytes(data);
                request1.ContentLength = bytedata.Length;
                Stream stream = request1.GetRequestStream();
                stream.Write(bytedata, 0, bytedata.Length);
                response1 = (HttpWebResponse)request1.GetResponse();
                streamReader = new StreamReader(response1.GetResponseStream());
                result = streamReader.ReadToEnd();
                request1 = (HttpWebRequest)WebRequest.Create("https://nn.hh.ru/");
                request1.CookieContainer = cookies;
                request1.Headers["Upgrade-Insecure-Requests"] = "1";
                request1.Referer = "https://nn.hh.ru/";
                request1.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8";
                request1.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/65.0.3325.181 Safari/537.36";
                response1 = (HttpWebResponse)request1.GetResponse();
                streamReader = new StreamReader(response1.GetResponseStream());
                result = streamReader.ReadToEnd();
                string link = textBox1.Text;
                for (; true; page++)
                {
                    k = 2;
                    request1 = (HttpWebRequest)WebRequest.Create(link + page);
                    request1.CookieContainer = cookies;
                    request1.Headers["Upgrade-Insecure-Requests"] = "1";
                    request1.Referer = link;
                    request1.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8";
                    request1.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/65.0.3325.181 Safari/537.36";
                    response1 = (HttpWebResponse)request1.GetResponse();
                    streamReader = new StreamReader(response1.GetResponseStream());
                    result = streamReader.ReadToEnd();
                    Regex regex = new Regex("(?<=href=\"/resume/).*(?=\" target=\"_)");
                    foreach (Match match in regex.Matches(result))
                    {
                        k++;
                    }
                    string[] linkresume = new string[k];
                    string[] vozrast = new string[k];
                    string[] pol = new string[k];
                    string[] gorod = new string[k];
                    string[] dolgnost = new string[k];
                    string[] zarplata = new string[k];
                    string[] opit = new string[k];
                    string[] naviki = new string[k];
                    string[] obomne = new string[k];
                    string[] obrazovanie = new string[k];
                    string[] language = new string[k];
                    string[] fio = new string[k];
                    string[] telefon = new string[k];
                    string[] url = new string[k];
                    m = 2;
                    regex = new Regex("(?<=href=\"/resume/).*(?=\" target=\"_)");
                    foreach (Match match in regex.Matches(result))
                    {
                        linkresume[m] = match.Value;
                        m++;
                    }
                    for (k = 2; k < m; k++)
                    {
                        request1 = (HttpWebRequest)WebRequest.Create("https://nn.hh.ru/resume/" + linkresume[k]);
                        request1.CookieContainer = cookies;
                        request1.Headers["Upgrade-Insecure-Requests"] = "1";
                        request1.Referer = link;
                        request1.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8";
                        request1.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/65.0.3325.181 Safari/537.36";
                        response1 = (HttpWebResponse)request1.GetResponse();
                        streamReader = new StreamReader(response1.GetResponseStream());
                        result = streamReader.ReadToEnd();
                        vozrast[k] = Regex.Match(result, "(?<=resume-personal-age\">).*?(?=</span>, )").Value;
                        pol[k] = Regex.Match(result, "(?<=-gender\">).*?(?=</span>)").Value;
                        gorod[k] = Regex.Match(result, "(?<=resume-personal-address\">).*?(?=</span>, )").Value;
                        dolgnost[k] = Regex.Match(result, "(?<=IE=edge\"><title>).*?(?=</title><)").Value;
                        zarplata[k] = Regex.Match(result, "(?<=resume-block-salary\">).*?(?=.</span></div)").Value;
                        if (pol[k] == "Male" || pol[k] == "Female")
                            opit[k] = Regex.Match(result, "(?<=\">Work experience ).*?(?=</span)").Value;
                        else
                            opit[k] = Regex.Match(result, "(?<=sub\">Опыт работы ).*?(?=</span)").Value;
                        regex = new Regex("(?<=bloko-tag__text\">).*?(?=</span)");
                        foreach (Match match in regex.Matches(result))
                        {
                            naviki[k] = naviki[k] + match.Value + ",";
                        }
                        obomne[k] = Regex.Match(result, "(?<=resume-block-skills\"><script data-name=\"HH/Linkify\" data-params=\"\"></script>).*?(?=</div></div></)").Value;
                        obomne[k] = obomne[k].Replace("<br>", " ");
                        regex = new Regex("(?<=bloko-link bloko-link_list\">).*?(?=</div></div></div></div></)");
                        foreach (Match match in regex.Matches(result))
                        {
                            obrazovanie[k] = obrazovanie[k] + match.Value + ";";
                            obrazovanie[k] = obrazovanie[k].Replace("</a></div><div data-qa=\"resume-block-education-organization\">", " : ");
                        }
                        if (obrazovanie[k] == null)
                        {
                            regex = new Regex("(?<=resume-block-education-name\">).*?(?=</div></div></div></div></)");
                            foreach (Match match in regex.Matches(result))
                            {
                                obrazovanie[k] = obrazovanie[k] + match.Value + ";";
                                obrazovanie[k] = obrazovanie[k].Replace("</div><div data-qa=\"resume-block-education-organization\">", " : ");
                            }
                        }
                        regex = new Regex("(?<=resume-block-language-item\">).*?(?=</p><)");
                        foreach (Match match in regex.Matches(result))
                        {
                            language[k] = language[k] + "," + match.Value;
                        }
                        fio[k] = Regex.Match(result, "(?<=resume-personal-name\">).*?(?=</h1></)").Value;
                        telefon[k] = Regex.Match(result, "(?<=telephone\" data-qa=\"resume-contact-preferred\">)[\\w\\W]*?(?=</span>)").Value;
                        if (telefon[k] == "")
                            telefon[k] = Regex.Match(result, "(?<=telephone\">)[\\w\\W]*?(?=</span>)").Value;
                        telefon[k] = telefon[k].Replace("\n              ", "");
                        Thread.Sleep(5000);
                    }
                    for (int i = 2; i < k; i++)
                    {
                        sheet.Cells[chet + i, 1].Value = vozrast[i];
                        sheet.Cells[chet + i, 2].Value = pol[i];
                        sheet.Cells[chet + i, 3].Value = gorod[i];
                        sheet.Cells[chet + i, 4].Value = dolgnost[i];
                        sheet.Cells[chet + i, 5].Value = zarplata[i];
                        sheet.Cells[chet + i, 6].Value = opit[i];
                        sheet.Cells[chet + i, 7].Value = naviki[i];
                        sheet.Cells[chet + i, 8].Value = obomne[i];
                        sheet.Cells[chet + i, 9].Value = obrazovanie[i];
                        sheet.Cells[chet + i, 10].Value = language[i];
                        sheet.Cells[chet + i, 11].Value = fio[i];
                        sheet.Cells[chet + i, 12].Value = telefon[i];
                        sheet.Cells[chet + i, 13].Value = "https://nn.hh.ru/resume/" + linkresume[i];
                        porog++;
                    }
                    chet=chet+k-2;
                    if (porog >= 4)
                        break;
                }
            }
            package.Save();
            MessageBox.Show("Парсинг завершен!");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "EXCEL|*.xlsx";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string path = dialog.FileName;
                textBox2.Text = path;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "EXCEL|*.xlsx";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string path = dialog.FileName;
                textBox3.Text = path;
            }
        }
    }
}
