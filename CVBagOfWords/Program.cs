using System;
using Microsoft.ML;
using System.Collections.Generic;
using System.IO;
using Microsoft.ML.Data;
using System.Linq;
using System.Text;
using Spire.Doc;
using System.Text.RegularExpressions;
using Microsoft.ML.Runtime;
using Microsoft.ML.Transforms.Text;
using System.Collections;

namespace CVBagOfWords
{
    class Program
    {
        private static MLContext mlcontext;
        private static string dir_path = "C:/Users/ben/Documents/attachments";
        private static string out_path= "C:/Users/ben/Documents/CVData1.txt";
        private static string out_path2 = "C:/Users/ben/Documents/attachments/CVData1.txt";
        private static string delimiter =", ";
        private static List<string> titles = new List<string>{ "System and Web Developer","Java Developer","Senior Java Developer", "Developer",
            "Project Manager","Scrum Master Product Owner Agile Coach", "Project Manager", "Project Manager Agile","Senior Project Manager",
            "IT Project Manager", "Business Manager", "Business Intelligence Consultant", "Analyst", "Senior Data Analyst AMl",
            "Head of UX", "Senior Consultant", "Head of UX","Back-End Developer Front End Developer", "Full Stack Developer (.NET/Node.js)",
            "Full stack Developer", "Senior Software Developer"};
        
        private static List<string> supportedFilesDoc = new List<string>();
        private static List<string> supportedFilesPdf = new List<string>();

        static string _testSet = "C:/Users/ben/Documents/attachments/cvs-test.txt";
        static string _trainingSet = "C:/Users/ben/Documents/attachments/cvs-train.txt";
        static void Main(string[] args)
        {
            //getCV(dir_path);
            mlcontext = new MLContext();
            var texts = new List<string>();
            var training = new List<CVData>();
            var test = new List<CVData>();

            for (var i = 0; i < titles.Count(); i++)
            {
                var title = titles[i];
                using (var rd = new StreamReader(out_path2))
                {
                    while (!rd.EndOfStream)
                    {
                        var splits = rd.ReadLine().Split(',');
                        texts.Add(Regex.Replace(splits[0], "[^A-Za-z0-9 ]", ""));
                    }
                }

                Random _random = new Random();

                texts = texts.OrderBy(s => _random.Next()).ToList();
                var trainingTextsCount = (texts.Count / 100) * 80;
                var trainingTexts = texts.GetRange(0, trainingTextsCount);
                training.AddRange(trainingTexts.Select(s => new CVData { CVText = s, Label = title }).ToList());

                var testTexts = texts.GetRange(trainingTextsCount, texts.Count - trainingTextsCount);
                test.AddRange(testTexts.Select(s => new CVData { CVText = s, Label = title }).ToList());
            }

            File.AppendAllLines(_testSet, test.Select(s => $"{s.CVText}\t{s.Label}"));
            File.AppendAllLines(_trainingSet, training.Select(s => $"{s.CVText}\t{s.Label}"));

            var model = Train();
            Evaluate(model);

            var model2 = TrainWordEmbedding();
            Evaluate(model2);







            /**var dictionary = new List<String>();
            using (var rd = new StreamReader(out_path))
            {
                while (!rd.EndOfStream)
                {
                    var splits = rd.ReadLine().Split(',');
                    dictionary.Add(Regex.Replace(splits[0], "[^A-Za-z0-9 ]", ""));
                }
            }

            dictionary = dictionary.Distinct().ToList();

            int c = dictionary.Count;
            Console.WriteLine("The dictionary contains " + c + " words");**/

            //Initialize an ML context
            /**mlcontext = new MLContext();
            var data = new List<Input> 
            { 
                new Input { Text = GetTxtText(out_path)}
            };

            var title = mlcontext.Data.LoadFromEnumerable(titles);

            //Create an empty list class to hold all the inputs from the file
            var emptydata = mlcontext.Data.LoadFromEnumerable(new List<Input>());

            //create a dataview for the text data
            var dataView = mlcontext.Data.LoadFromEnumerable(data);**/




            /**var text_pipeline = mlcontext.Transforms.Text.TokenizeIntoWords("Tokens", "Text", separators: new[] { ' ', '.', ',' })
                .Append(mlcontext.Transforms.Text.RemoveDefaultStopWords("Tokens", "Tokens",
                    Microsoft.ML.Transforms.Text.StopWordsRemovingEstimator.Language.English));

            var bagWordsTransform = text_pipeline.Fit(dataView);
            var bagWordsDataView = bagWordsTransform.Transform(dataView);

            DataViewSchema columns = bagWordsDataView.Schema;
            Console.WriteLine(columns);

            var bagWordsPipeline = mlcontext.Transforms.Text.NormalizeText("Text")
           .Append(mlcontext.Transforms.Text.ProduceWordBags(
                  "BagOfWords",
                  "Text",
                  ngramLength: 1,
                  useAllLengths: false,
                  weighting: Microsoft.ML.Transforms.Text.NgramExtractingEstimator.WeightingCriteria.Tf
              ));
            var bg = bagWordsPipeline.Fit(bagWordsDataView);*/


            /*using (DataViewRowCursor cursor = bagWordsDataView.GetRowCursor(columns))
            {
                // Define variables where extracted values will be stored to
                string normtext1 = default;
                VBuffer<string> normtext2 = default;
                string normtext3 = default;

                // Define delegates for extracting values from columns
                ValueGetter<string> normtext1Delegate = cursor.GetGetter<string>(columns[0]);
                ValueGetter<VBuffer<string>> normtext2Delegate = cursor.GetGetter<VBuffer<string>>(columns[1]);
                ValueGetter<string> normtext3Delegate = cursor.GetGetter<string>(columns[2]);

                // Iterate over each row
                while (cursor.MoveNext())
                {
                    //Get values from respective columns
                    normtext1Delegate.Invoke(ref normtext1);
                    normtext2Delegate.Invoke(ref normtext2);
                    normtext3Delegate.Invoke(ref normtext3);
                }
            }*/


            //Initialize an ML contet
            //MLContext mlcontext = new MLContext();

            //Create an empty list class to hold all the inputs from the file
            //var emptydata = mlcontext.Data.LoadFromEnumerable(new List<Input>());

            //get data from docx file using getTextData() method
            /*var data = new List<Input>
            {
                new Input { Text = getTextData(@"C:/Users/ben/Documents/attachments/CV1.docx")}
            };

            //create a dataview for the text data
            var dataView = mlcontext.Data.LoadFromEnumerable(data);


            var bagWordsPipeline = mlcontext.Transforms.Text.NormalizeText("Text")
            .Append(mlcontext.Transforms.Text.ProduceWordBags(
                   "BagOfWords",
                   "Text",
                   ngramLength: 1,
                   useAllLengths: false,
                   weighting: Microsoft.ML.Transforms.Text.NgramExtractingEstimator.WeightingCriteria.Tf
               ));

            var bagWordsTransform = bagWordsPipeline.Fit(dataView);
            var bagWordsDataView = bagWordsTransform.Transform(dataView);

            var predictionEngine = mlcontext.Model.CreatePredictionEngine<Input, Output>(bagWordsTransform);

            var prediction = predictionEngine.Predict(data[0]);



            VBuffer<ReadOnlyMemory<char>> slotNames = default;
            bagWordsDataView.Schema["BagOfWords"].GetSlotNames(ref slotNames);


            var bagOfWordColumn = bagWordsDataView.GetColumn<VBuffer<float>>(bagWordsDataView.Schema["BagOfWords"]);
            var slots = slotNames.GetValues();

            /*Console.Write("NGrams: ");
            foreach (var featureRow in bagOfWordColumn)
            {
                foreach (var item in featureRow.Items())
                {
                    Console.Write($"{slots[item.Key]}  ");

                }

                Console.WriteLine();
            }

            Console.Write("Word Counts: ");
            for (int i = 0; i < prediction.BagOfWords.Length; i++)
            {

                Console.Write($"{prediction.BagOfWords[i]:F4}  ");

            }

            Console.WriteLine(Environment.NewLine);*/


        }

        private static ITransformer Train()
        {
            var trainData = mlcontext.Data.LoadFromTextFile<CVData>(_trainingSet, hasHeader: false, separatorChar: '\t');

            var pipeline = mlcontext.Transforms.Text.NormalizeText("CVText")                
                .Append(mlcontext.Transforms.Text.TokenizeIntoWords("Tokens", "CVText"))
                .Append(mlcontext.Transforms.Text.RemoveDefaultStopWords("Tokens", "Tokens", language:
                StopWordsRemovingEstimator.Language.English)) 
                .Append(mlcontext.Transforms.Text.FeaturizeText("Features", "Tokens"))
                .Append(mlcontext.Transforms.Conversion.MapValueToKey(nameof(CVData.Label)))
                .Append(mlcontext.MulticlassClassification.Trainers.NaiveBayes());

           
            Console.WriteLine("=============== Create and Train the Model ===============");
            var model = pipeline.Fit(trainData);
            Console.WriteLine("=============== End of training ===============");
            Console.WriteLine();

            return model;
        }

        private static ITransformer TrainWordEmbedding()
        {
            var trainData = mlcontext.Data.LoadFromTextFile<CVData>(_trainingSet, hasHeader: false, separatorChar: '\t');

            var pipeline = mlcontext.Transforms.Text.NormalizeText("CVText")
                .Append(mlcontext.Transforms.Text.TokenizeIntoWords("Tokens", "CVText"))
                .Append(mlcontext.Transforms.Text.RemoveDefaultStopWords("Tokens", "Tokens", language:
                StopWordsRemovingEstimator.Language.English))
                .Append(mlcontext.Transforms.Text.ApplyWordEmbedding("Features", "Tokens",
                WordEmbeddingEstimator.PretrainedModelKind.GloVe100D))
                .Append(mlcontext.Transforms.Conversion.MapValueToKey(nameof(CVData.Label)))
                .Append(mlcontext.MulticlassClassification.Trainers.NaiveBayes());


            Console.WriteLine("=============== Create and Train the Model ===============");
            var model = pipeline.Fit(trainData);
            Console.WriteLine("=============== End of training ===============");
            Console.WriteLine();

            return model;
        }


        public static void Evaluate(ITransformer model)
        {
            var testData = mlcontext.Data.LoadFromTextFile<CVData>(_trainingSet, hasHeader: false, separatorChar: '\t');

            Console.WriteLine("=============== Evaluating Model accuracy with Test data===============");
            IDataView predictions = model.Transform(testData);

            var metrics = mlcontext.MulticlassClassification.Evaluate(predictions);

            Console.WriteLine($"*************************************************************************************************************");
            Console.WriteLine($"*       Metrics for Multi-class Classification model - Test Data     ");
            Console.WriteLine($"*------------------------------------------------------------------------------------------------------------");
            Console.WriteLine($"*       MicroAccuracy:    {metrics.MicroAccuracy:0.###}");
            Console.WriteLine($"*       MacroAccuracy:    {metrics.MacroAccuracy:0.###}");
            Console.WriteLine($"*       LogLoss:          {metrics.LogLoss:#.###}");
            Console.WriteLine($"*       LogLossReduction: {metrics.LogLossReduction:#.###}");
            Console.WriteLine($"*************************************************************************************************************");
        }


        //function to read all CV documents stored in a folder
        public static void getCV(string paths)
        {
         
                if(Directory.Exists(paths))
                {
                    Console.WriteLine("============Starting writing to text file=============");
                    Console.WriteLine();
                    processCVDirectory(paths);
                    Console.WriteLine("============Finished writing to text file=============");
                }
                else
                {
                    Console.WriteLine("{0} is not a valid file or directory.", paths);
                }
            
        }
        public static void processCVDirectory(string cvdirectory)
        {
            string[] fileentries = Directory.GetFiles(cvdirectory);
            
            foreach (string fileName in fileentries)
            {

                //Check the file extensions of each file whether it is a doc file or pdf file
                //The extension of the file will be .pdf or .doc
                FileInfo file_type = new FileInfo(fileName);
                if(file_type.Extension == ".docx" || file_type.Extension == ".doc")
                {
                    Console.WriteLine("Adding: {0}...", Path.GetFileName(fileName));
                    supportedFilesDoc.Add(getTextData(fileName));

                }
                else if (file_type.Extension == ".pdf") 
                {
                    Console.WriteLine("Adding: {0}...", Path.GetFileName(fileName));
                    supportedFilesPdf.Add(getTextData(fileName));
                }
                else 
                {
                    Console.WriteLine("{0} is not a valid file or directory.", fileName);
                }
                
                
                
            }
            foreach (string text in supportedFilesDoc)
            {
                
                processCV(text, out_path);
                
            }          

        }

        public static void processCV(string text,string path)
        {
            Console.WriteLine("Processed file '{0}'.", path);
            try
            {
                using (StreamWriter file = new StreamWriter(path, true))
                {
                                        
                        file.WriteLine(text + delimiter + Environment.NewLine);
                 
                     
                }
            }
            catch(Exception ex)
            {
                throw new ApplicationException("The file write failed: ", ex);
            }       
            
        }

        public static string getTextData(string path)
        {
            Document document1 = new Document();
            //document1.LoadFromFile(@"C:/Users/ben/Documents/attachments/CV1.docx");
            document1.LoadFromFile(path);

            //Initialzie StringBuilder Instance
            StringBuilder sb = new StringBuilder();

            //Extract Text from Word and Save to StringBuilder Instance
            foreach (Spire.Doc.Section section in document1.Sections)
            {
                foreach (Spire.Doc.Documents.Paragraph paragraph in section.Paragraphs)

                {
                   
                    sb.AppendLine(paragraph.Text);
                }
            }
            return sb.ToString();
            
        }
        public static string GetTextPdf(string path)
        {
            Spire.Pdf.PdfDocument pdoc = new Spire.Pdf.PdfDocument();
            pdoc.LoadFromFile(path);

            //Initialzie StringBuilder Instance for pdf
            StringBuilder sbpdf = new StringBuilder();
            //Extract text from all pages
            foreach (Spire.Pdf.PdfPageBase page in pdoc.Pages)
            {
                sbpdf.Append(page.ExtractText());

            }

            return sbpdf.ToString();


        }

        //method to normalize and remove stop words from the text extracted from the word documents
        //create a pipeline to which has normalization and removal of stop words
        public static IEstimator<ITransformer> NormalizeText()
        {
            var text_pipeline = mlcontext.Transforms.Text.NormalizeText("Text")
                .Append(mlcontext.Transforms.Text.RemoveDefaultStopWords("Text"));
            return text_pipeline;
        }

        public static string GetTxtText(string path)
        {
            string text = System.IO.File.ReadAllText(path);

            return text.ToString();
        }


        public class CVData
        {
            [LoadColumn(0)]
            public string CVText;

            [LoadColumn(1), ColumnName(name: "Label")]
            public string Label;
        }
        public class CVPrediction
        {
            [ColumnName("Score")]
            public float[] Score;
        }


    }
}
