using System;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace PDF_PropertyChanger_CUI
{
    class PDF_PropertyChanger_CUI
    {

        static void Main(string[] args)
        {
            Console.WriteLine("PDF Property Changer"); // タイトル的なやつ

            // ソースパス取得
            Console.WriteLine("*PDF格納パスを入力してください");
            Console.Write("格納パス(絶対パス):");
            string PdfSrcPath = Console.ReadLine();

            string title, author, subject, keywords, PW; // タイトル・作成者・サブタイトル・キーワード
            string[] files; // ソースパス内のPDFファイルパス取得用

            
            try
            {
                // 拡張子が「pdf」のファイルのみ、絶対パスを取得する
                files = Directory.GetFiles(PdfSrcPath, "*.pdf");

                // PDFファイルが1つもない場合は終了
                if (files.Length == 0)
                {
                    Console.WriteLine("PDFファイルが格納されていません\n処理を中断します");
                    Console.ReadLine();
                    return;
                }

                // 格納PDFファイル名の一覧を表示
                Console.WriteLine("\n--- 格納PDFファイル ---");

                for (int i = 0; i < files.Length; i++)
                {
                    Console.WriteLine(Path.GetFileName(files[i]));
                }


                // パスワード取得
                Console.WriteLine("\n*設定するパスワードを入力してください\nセキュリティ設定が不要な場合は何も入力せずEnter");
                Console.Write("オーナー(編集)パスワード:");
                PW = Console.ReadLine();

                //　プロパティ項目取得
                Console.WriteLine("\n概要項目を設定します\n（変更不要な項目は、何も入力せずEnter）");

                Console.Write("タイトル：");
                title = Console.ReadLine();

                Console.Write("作成者：");
                author = Console.ReadLine();

                Console.Write("サブタイトル：");
                subject = Console.ReadLine();

                Console.Write("キーワード：");
                keywords = Console.ReadLine();

                // 変更開始
                Console.WriteLine("--- start ---");

                for (int i = 0; i < files.Length; i++)
                {
                    // 引数：PDFファイル絶対パス,パスワード,タイトル,作成者,サブタイトル,キーワード
                    DecriptPdfDoc(files[i], PW, title, author, subject, keywords);
                }
                
                Console.Write("終了しました");
                Console.ReadLine();

            }
            catch (Exception e)
            {
                Console.WriteLine("{0}/{1}/{2}", e.Message, e.Source, e.TargetSite);
            }
        }

        // PDFファイルオープン・パスワード,プロパティ変更、PDFファイル置換
        static void DecriptPdfDoc(string PdfSrcPath, string PW, string title, string author, string subject, string keywords)
        {
            PdfReader reader;
            int PageNum; // ページ数取得用
            int i = 0; // カウント用
            bool setB = false; // プロパティ変更用

            string PdfDstPath = Path.GetTempPath() + Path.GetFileName(PdfSrcPath);

            try
            {
                reader = new PdfReader(PdfSrcPath);

                Document doc = new Document();

                // 一時ファイルのフォルダを取得し、置換用PDFファイルを作成
                PdfWriter writer = PdfWriter.GetInstance(doc, new FileStream(PdfDstPath, FileMode.Create));

                Console.WriteLine("\n{0}:{1}/{2}:{3}", "start", Path.GetFileName(PdfSrcPath), "page(s)", reader.NumberOfPages);
                
                // パスワードが入力されている場合はセキュリティ設定をする
                if (!PW.Equals(string.Empty))
                {
                    writer.Open();

                    // セキュリティ設定
                    writer.SetEncryption(
                        PdfWriter.STRENGTH40BITS, //暗号化強度(40bitでAcrobat3.0)
                        "", // ユーザー(閲覧)パスワード
                        PW, // オーナー(編集)パスワード
                            // セキュリティ設定
                        PdfWriter.AllowCopy | // 内容のコピーと抽出
                        PdfWriter.AllowPrinting | // 印刷
                        PdfWriter.AllowModifyContents | // 文書の変更
                                                        /* PdfWriter.AllowModifyAnnotations |  注釈 */
                        PdfWriter.AllowFillIn | // フォームフィールドの入力と署名
                        PdfWriter.AllowScreenReaders | // アクセシビリティを有効にする
                        PdfWriter.AllowAssembly // 文章アセンブリ
                    );
                }
                
                doc.Open();

                PdfContentByte content = writer.DirectContent;
                PageNum = reader.NumberOfPages;
                // 全ページの内容を一時ファイルに設定
                while (i < PageNum)
                {
                    doc.NewPage();
                    PdfImportedPage iPage = writer.GetImportedPage(reader, ++i);
                    content.AddTemplate(iPage, 0, 0);
                }

                // タイトル
                
                if (!reader.Info.ContainsKey("Title"))
                {
                    Console.WriteLine("Title:{0}→{1}", "No key", title);
                    setB = doc.AddTitle(title);

                }
                else if (title.Equals(string.Empty))
                {
                    Console.WriteLine("Title:{0}→{1}", reader.Info["Title"], "変更なし");
                    setB = doc.AddTitle(reader.Info["Title"]);
                }
                else if (!title.Equals(string.Empty))
                {
                    Console.WriteLine("Title:{0}→{1}", reader.Info["Title"], title);
                    setB = doc.AddTitle(title);
                }


                // 作成者
                if (!reader.Info.ContainsKey("Author"))
                {
                    Console.WriteLine("Author:{0}→{1}", "No key", author);
                    setB = doc.AddAuthor(author);
                }
                else if (author.Equals(string.Empty))
                {
                    Console.WriteLine("Author:{0}→{1}", reader.Info["Author"], "変更なし");
                    setB = doc.AddAuthor(reader.Info["Author"]);
                }
                else if (!author.Equals(string.Empty))
                {
                    Console.WriteLine("Author:{0}→{1}", reader.Info["Author"], author);
                    setB = doc.AddAuthor(author);
                }

                // サブタイトル
                if (!reader.Info.ContainsKey("Subject"))
                {
                    Console.WriteLine("Subject:{0}→{1}", "No key", subject);
                    setB = doc.AddSubject(subject);
                }
                else if (subject.Equals(string.Empty))
                {
                    Console.WriteLine("Subject:{0}→{1}", reader.Info["Subject"], "変更なし");
                    setB = doc.AddSubject(reader.Info["Subject"]);
                }
                else if (!subject.Equals(string.Empty))
                {
                    Console.WriteLine("Subject:{0}→{1}", reader.Info["Subject"], subject);
                    setB = doc.AddSubject(subject);
                }

                
                // キーワード
                if (!reader.Info.ContainsKey("Keywords"))
                {
                    Console.WriteLine("Keywords:{0}→{1}", "No key", keywords);
                    setB = doc.AddKeywords(keywords);
                }
                else if (keywords.Equals(string.Empty))
                {
                    Console.WriteLine("Keywords:{0}→{1}", reader.Info["Keywords"], "変更なし");
                    setB = doc.AddKeywords(reader.Info["Keywords"]);
                }
                else if (!keywords.Equals(string.Empty))
                {
                    Console.WriteLine("Keywords:{0}→{1}", reader.Info["Keywords"], keywords);
                    setB = doc.AddKeywords(keywords);
                }


                doc.Close();
                reader.Close();
                writer.Close();

                // PDFファイル入れ替え
                File.Delete(PdfSrcPath); // 元のファイルを削除
                File.Move(PdfDstPath, PdfSrcPath); // 一時ファイルを移動

                Console.WriteLine("{0}:Prop set done.", Path.GetFileNameWithoutExtension(PdfSrcPath));

            }
            catch (Exception e)
            {
                Console.WriteLine("{0}/{1}", "Exception", e.Message);
            }
        }
    }
}
