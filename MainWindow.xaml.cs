using DevExpress.Mvvm;
using DevExpress.Xpf.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Word = Microsoft.Office.Interop.Word;

namespace Reimbursement
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : ThemedWindow
    {
        public MainWindow()
        {
            //存放所有图片地址的列表
            picSrc1 = new List<string>();
            picSrc2 = new List<string>();
            picSrc3 = new List<string>();

            InitializeComponent();

            topText.Text = "报销文件一共需要三部分内容，一为商品明细，二为支付凭证（微信截图或银行卡消费明细），三为商品发票（哈尔滨工业大学，税号12100000400000456B）。";
        }

        public List<string> picSrc1 { get; set; }
        public List<string> picSrc2 { get; set; }
        public List<string> picSrc3 { get; set; }


        /// <summary>
        /// 选择商品明细图片
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Open1(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "*.*|*.bmp;*.jpg;*.jpeg;*.tiff;*.tiff;*.png";
            ofd.Title = "选择商品明细图片";
            ofd.Multiselect = true;
            if (ofd.ShowDialog() != true)
                return;
            
            //显示预览图片
            BitmapImage bmp = new BitmapImage(new Uri(@ofd.FileNames[0]));
            imageShow.Source = bmp;
            foreach (string str in ofd.FileNames)
            {
                picSrc1.Add(str);
            }

            text1.Text = string.Format("已选择{0}张图片", picSrc1.Count());
        }

        /// <summary>
        /// 选择支付凭证图片
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Open2(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "*.*|*.bmp;*.jpg;*.jpeg;*.tiff;*.tiff;*.png";
            ofd.Title = "选择支付凭证图片";
            ofd.Multiselect = true;
            if (ofd.ShowDialog() != true)
                return;

            //显示预览图片
            BitmapImage bmp = new BitmapImage(new Uri(@ofd.FileNames[0]));
            imageShow.Source = bmp;
            foreach (string str in ofd.FileNames)
            {
                picSrc2.Add(str);
            }
            text2.Text = string.Format("已选择{0}张图片", picSrc2.Count());
        }

        /// <summary>
        /// 选择商品发票图片
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Open3(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "*.*|*.bmp;*.jpg;*.jpeg;*.tiff;*.tiff;*.png";
            ofd.Title = "选择商品发票图片";
            ofd.Multiselect = true;
            if (ofd.ShowDialog() != true)
                return;

            //显示预览图片
            BitmapImage bmp = new BitmapImage(new Uri(@ofd.FileNames[0]));
            imageShow.Source = bmp;
            foreach (string str in ofd.FileNames)
            {
                picSrc3.Add(str);
            }
            text3.Text = string.Format("已选择{0}张图片", picSrc3.Count());
        }

        /// <summary>
        /// 导出word文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ExportWord(object sender, RoutedEventArgs e)
        {
            object oMissing = System.Reflection.Missing.Value;
            object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */

            //Start Word and create a new document.
            Word._Application oWord;
            Word._Document oDoc;
            oWord = new Word.Application();
            oWord.Visible = false;
            oDoc = oWord.Documents.Add(ref oMissing, ref oMissing,
            ref oMissing, ref oMissing);

            //设置页眉
            oDoc.PageSetup.HeaderDistance = 30.0f;//页眉位置
            oWord.ActiveWindow.View.Type = Word.WdViewType.wdOutlineView;//视图样式
            oWord.ActiveWindow.View.SeekView = Word.WdSeekView.wdSeekPrimaryHeader;//进入页眉设置，其中页眉边距在页面设置中已完成
            oWord.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            //插入页眉图片
            Word.InlineShape shape1 = oWord.ActiveWindow.ActivePane.Selection.InlineShapes.AddPicture("C:/Users/李美程/Desktop/报销/Reimbursement/Reimbursement/logo.png", ref oMissing, ref oMissing, ref oMissing);
            shape1.Height = 12;
            shape1.Width = 12;
            oWord.ActiveWindow.ActivePane.Selection.InsertAfter("哈尔滨工业大学报销凭证");
            oWord.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekMainDocument;//退出页眉设置


            //添加页码
            Word.PageNumbers pns = oWord.Selection.Sections[1].Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages].PageNumbers;
            pns.NumberStyle = Word.WdPageNumberStyle.wdPageNumberStyleNumberInDash;
            pns.HeadingLevelForChapter = 0;
            pns.IncludeChapterNumber = false;
            pns.ChapterPageSeparator = Word.WdSeparatorType.wdSeparatorHyphen;
            pns.RestartNumberingAtSection = false;
            pns.StartingNumber = 0;
            object pagenmbetal = Word.WdPageNumberAlignment.wdAlignPageNumberCenter;
            object first = true;
            oWord.Selection.Sections[1].Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages].PageNumbers.Add(ref pagenmbetal, ref first);

            foreach (string picName in picSrc1)
            {
                Word.Paragraph oPara;
                oPara = oDoc.Content.Paragraphs.Add(ref oMissing);
                oPara.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                oPara.Range.Font.Bold = 1;
                oPara.Range.Font.Size = 10;
                oPara.Format.SpaceBefore = 10;
                string title = string.Format("商品明细-{0}\n", picSrc1.IndexOf(picName) + 1);
                oPara.Range.InsertBefore(title);



                //插入图片
                Word.InlineShape pic;
                pic = oWord.ActiveDocument.InlineShapes.AddPicture(@picName, oMissing, oMissing, oPara.Range);
                oWord.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                pic.Range.ParagraphFormat.SpaceAfter = 6;
                oPara.Range.InsertParagraphAfter();

                object oCollapseEnd = Word.WdCollapseDirection.wdCollapseEnd;
                object oPageBreak = Word.WdBreakType.wdPageBreak;
                oPara.Range.Collapse(ref oCollapseEnd);
                oPara.Range.InsertBreak(ref oPageBreak);
                oPara.Range.Collapse(ref oCollapseEnd);
            }

            foreach (string picName in picSrc2)
            {
                Word.Paragraph oPara;
                oPara = oDoc.Content.Paragraphs.Add(ref oMissing);
                oPara.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                oPara.Range.Font.Bold = 1;
                oPara.Range.Font.Size = 10;
                oPara.Format.SpaceBefore = 10;
                string title = string.Format("支付凭证-{0}\n", picSrc2.IndexOf(picName) + 1);
                oPara.Range.InsertBefore(title);

                //插入图片
                Word.InlineShape pic;
                pic = oWord.ActiveDocument.InlineShapes.AddPicture(@picName, oMissing, oMissing, oPara.Range);
                oWord.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                pic.Range.ParagraphFormat.SpaceAfter = 6;
                oPara.Range.InsertParagraphAfter();

                object oCollapseEnd = Word.WdCollapseDirection.wdCollapseEnd;
                object oPageBreak = Word.WdBreakType.wdPageBreak;
                oPara.Range.Collapse(ref oCollapseEnd);
                oPara.Range.InsertBreak(ref oPageBreak);
                oPara.Range.Collapse(ref oCollapseEnd);
            }

            foreach (string picName in picSrc3)
            {
                Word.Paragraph oPara;
                oPara = oDoc.Content.Paragraphs.Add(ref oMissing);
                oPara.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                oPara.Range.Font.Bold = 1;
                oPara.Range.Font.Size = 10;
                oPara.Format.SpaceBefore = 10;
                string title = string.Format("商品发票-{0}\n", picSrc3.IndexOf(picName) + 1);
                oPara.Range.InsertBefore(title);

                //插入图片
                Word.InlineShape pic;
                pic = oWord.ActiveDocument.InlineShapes.AddPicture(@picName, oMissing, oMissing, oPara.Range);
                oWord.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                if (picSrc3.IndexOf(picName) != picSrc3.Count() - 1)  //不是最后一个 要换页
                {
                    pic.Range.ParagraphFormat.SpaceAfter = 6;
                    oPara.Range.InsertParagraphAfter();
                    object oCollapseEnd = Word.WdCollapseDirection.wdCollapseEnd;
                    object oPageBreak = Word.WdBreakType.wdPageBreak;
                    oPara.Range.Collapse(ref oCollapseEnd);
                    oPara.Range.InsertBreak(ref oPageBreak);
                    oPara.Range.Collapse(ref oCollapseEnd);
                }
                
            }

            oDoc.Save();    //保存文件
        }
    }
}
