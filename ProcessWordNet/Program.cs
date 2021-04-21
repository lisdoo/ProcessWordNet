using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using MSWord = Microsoft.Office.Interop.Word;
using System.IO;
using System.Runtime.Serialization;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.Encodings.Web;
using System.Text.Unicode;

namespace ProcessWordNet
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("请指定参数！ 1：输入文件  2：输入文件");
                return;
            }
            Microsoft.Office.Interop.Word.Application wapp = new Microsoft.Office.Interop.Word.Application();
            MSWord.Document wordDoc;
            wapp.Visible = true;
            object filename = args[0];
            object isread = false;
            object isvisible = true;
            object miss = System.Reflection.Missing.Value;
            wordDoc = wapp.Documents.Open(ref filename, ref miss, ref isread, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref isvisible, ref miss, ref miss, ref miss, ref miss);

            WordParagraph map = createObject(wordDoc.Paragraphs);

            Console.WriteLine(map);

            var options = new JsonSerializerOptions
            {
                Encoder = JavaScriptEncoder.Create(UnicodeRanges.All, UnicodeRanges.All),
                WriteIndented = true
            };



            string jsonString = JsonSerializer.Serialize(map, options);
            File.WriteAllText(args[1], jsonString);
        }

        static WordParagraph createObject(MSWord.Paragraphs paras)
        {
            WordParagraph root = new WordParagraph();
            root.outline = 0;
            root.children = new List<WordParagraph>();

            WordParagraph wpTemp = null;
            WordParagraph wp = null;
            StringBuilder sb = null;
            //HashMap<int,WordParagraph> map = new HashMap<int, WordParagraph>();

            foreach (MSWord.Paragraph para in paras)
            {
                String outLineLevelInStr = para.OutlineLevel.ToString();
                // 当为标题时
                if (!outLineLevelInStr.Contains("wdOutlineLevelBodyText"))
                {
                    // 新标题出现时：
                    // new 当前对象
                    wp = new WordParagraph();

                    // 设置成员
                    int outLineLevel = int.Parse(outLineLevelInStr.Substring(outLineLevelInStr.Length - 1));
                    wp.outline = outLineLevel;
                    wp.title = para.Range.Text.ToString().Trim();
                    wp.children = new List<WordParagraph>();



                    // 对比临时对象和当前对象层次，并更新临时对象
                    if (wpTemp != null)
                    {
                        wpTemp.content = sb.ToString();
                        switch (wpTemp.outline - wp.outline)
                        {
                            // 下三级
                            case -3:
                                {
                                    {
                                        WordParagraph wp2 = new WordParagraph();
                                        wp2.outline = wp.outline-1;
                                        wp2.title = "";
                                        wp2.children = new List<WordParagraph>();

                                        wpTemp.children.Add(wp2);

                                        {
                                            WordParagraph wp3 = new WordParagraph();
                                            wp3.outline = wp.outline - 2;
                                            wp3.title = "";
                                            wp3.children = new List<WordParagraph>();

                                            wp2.children.Add(wp3);

                                            wp3.children.Add(wp);
                                        }

                                    }

                                    wp.parent = wpTemp;
                                }
                                break;
                            // 下二级
                            case -2:
                                {
                                    {
                                        WordParagraph wp2 = new WordParagraph();
                                        wp2.outline = wp.outline-1;
                                        wp2.title = "";
                                        wp2.children = new List<WordParagraph>();

                                        wpTemp.children.Add(wp2);

                                        wp2.children.Add(wp);
                                    }

                                    wp.parent = wpTemp;
                                }
                                break;
                            // 下一级
                            case -1:
                                {

                                    wpTemp.children.Add(wp);

                                    wp.parent = wpTemp;
                                }
                                break;
                            // 同一级
                            case 0:
                                {
                                    wpTemp.parent.children.Add(wp);

                                    wp.parent = wpTemp.parent;
                                } 
                                break;
                            // 上1级
                            case 1:
                                {
                                    wpTemp.parent.parent.children.Add(wp);

                                    wp.parent = wpTemp.parent.parent;
                                }
                                break;
                            // 上2级
                            case 2:
                                {
                                    wpTemp.parent.parent.parent.children.Add(wp);

                                    wp.parent = wpTemp.parent.parent.parent;
                                }
                                break;
                            // 上3级
                            case 3:
                                {
                                    wpTemp.parent.parent.parent.parent.children.Add(wp);

                                    wp.parent = wpTemp.parent.parent.parent.parent;
                                }
                                break;
                            // 上4级
                            case 4:
                                {
                                    wpTemp.parent.parent.parent.parent.parent.children.Add(wp);

                                    wp.parent = wpTemp.parent.parent.parent.parent.parent;
                                }
                                break;
                            // 上5级
                            case 5:
                                {
                                    wpTemp.parent.parent.parent.parent.parent.parent.children.Add(wp);

                                    wp.parent = wpTemp.parent.parent.parent.parent.parent.parent;
                                }
                                break;
                            // 上6级
                            case 6:
                                {
                                    wpTemp.parent.parent.parent.parent.parent.parent.parent.children.Add(wp);

                                    wp.parent = wpTemp.parent.parent.parent.parent.parent.parent.parent;
                                }
                                break;
                            // 上7级
                            case 7:
                                {
                                    wpTemp.parent.parent.parent.parent.parent.parent.parent.parent.children.Add(wp);

                                    wp.parent = wpTemp.parent.parent.parent.parent.parent.parent.parent.parent;
                                }
                                break;
                            // 上7级
                            case 8:
                                {
                                    wpTemp.parent.parent.parent.parent.parent.parent.parent.parent.parent.children.Add(wp);

                                    wp.parent = wpTemp.parent.parent.parent.parent.parent.parent.parent.parent.parent;
                                }
                                break;
                            // 上9级
                            case 9:
                                {
                                    wpTemp.parent.parent.parent.parent.parent.parent.parent.parent.parent.parent.children.Add(wp);

                                    wp.parent = wpTemp.parent.parent.parent.parent.parent.parent.parent.parent.parent.parent;
                                }
                                break;
                        }
                        wpTemp = wp;

                    } else
                    {
                        // 初始临时对象
                        wpTemp = wp;
                        root.children.Add(wpTemp);
                        wpTemp.parent = root;
                    }

                    // 父对象设置到当前对象的属性
                    Console.WriteLine(wpTemp.outline);

                    sb = new StringBuilder();

                    // 当为段落时
                } else
                {
                    if (wp == null) continue;
                    sb.Append(para.Range.Text.ToString());
                    sb.Append("\n");
                }
            }

            if (wp != null)
                wp.content = sb.ToString();

            return root;
        }
    }

    class WordParagraph
    {
        public int outline { get; set; }
        public string title { get; set; }
        public string content { get; set; }
        [JsonIgnore]
        public WordParagraph parent { get; set; }
        public List<WordParagraph> children { get; set; }
    }
}
