
using System.Net.Http.Headers;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Azure.AI.FormRecognizer;
using System.Text;
using System.Web.UI;
using Microsoft.Azure.CognitiveServices.Vision.CustomVision.Prediction;
using Azure.AI.FormRecognizer.Models;
using Azure;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Azure.CognitiveServices.Vision.CustomVision.Prediction.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Aspose.Words;
using Xceed.Words.NET;
using Xceed.Document.NET;
using System.Linq;
using System.Drawing;
using System.Drawing.Drawing2D;
using Image = Xceed.Document.NET.Image;
using System.Runtime.Remoting.Contexts;
using System.Drawing.Imaging;
using System.Net;
using System.Web;
using System.Web.UI.WebControls;
using System.IO.Compression;
using Table = Xceed.Document.NET.Table;

namespace ImageToForm
{
    public partial class _Default : Page
    {
        #region Private Members
        
        string ImagePath = string.Empty;

       

       ImagePrediction ObjectDetectionJSON;
        
        string FormIdentifierJSON = string.Empty;

        
        int sectioncount = 1;

        List<string> listpageControls = new List<string>();

        int Y_Distance = 5;

        PredictionBlock.ReadResult readResult;

        PredictionBlock.PageResult pageResult;

        string wordDocumentPath = string.Empty;

        string xmlFormPath = string.Empty;


        #endregion

        #region Page Methods
        protected void Page_Load(object sender, EventArgs e)
        {

        }

       
        protected async void btnUpload_Click(object sender, EventArgs e)
        {
          
            if (FUP_WatermarkImage.HasFile)
            {
                
                string FileType = Path.GetExtension(FUP_WatermarkImage.PostedFile.FileName).ToLower().Trim();
               
                // Checking the format of the uploaded file.  
                if (FileType != ".jpg" && FileType != ".png" && FileType != ".gif" && FileType != ".bmp")
                {
                    string alert = "alert('File Format Not Supported. Only .jpg, .png, .bmp and .gif file formats are allowed.');";
                    ScriptManager.RegisterStartupScript(this, GetType(), "JScript", alert, true);
                }
                else
                {
                    System.Drawing.Image img = null;
                    string FileDirectory = "C:\\Users\\SYED.HAZIM\\Pictures\\Saved Pictures\\";
                    FUP_WatermarkImage.PostedFile.SaveAs(FileDirectory + FUP_WatermarkImage.PostedFile.FileName);
                    ImagePath = FileDirectory + FUP_WatermarkImage.PostedFile.FileName;
                    string FileName = Path.GetFileNameWithoutExtension(FUP_WatermarkImage.PostedFile.FileName);
                    img = System.Drawing.Image.FromFile(ImagePath);
                    img = resizeImage(ImagePath, 1157,1157);
                   
                   
                    using (MemoryStream memory = new MemoryStream())
                    {
                        using (FileStream fs = new FileStream(FileDirectory + FileName + "1157.jpeg", FileMode.Create, FileAccess.ReadWrite))
                        {
                            img.Save(memory, System.Drawing.Imaging.ImageFormat.Jpeg);
                            byte[] bytes = memory.ToArray();
                            fs.Write(bytes, 0, bytes.Length);
                            fs.Close();
                        }
                    }
                    ImagePath = FileDirectory + FileName + "1157.jpeg";
                    await PrepareXMLDocument();
                   

                }
                File.Delete(ImagePath);
            }
        }
        #endregion

        #region Authentication for Azure Models
        private static FormRecognizerClient AuthenticateClient()
        {
            string endpoint = "https://imagetoformrecognizer.cognitiveservices.azure.com/";
            string apiKey = "1b3db753f4b344268e948b24476f1663";
            var credential = new AzureKeyCredential(apiKey);
            var client = new FormRecognizerClient(new Uri(endpoint), credential);
            return client;
        }
        private CustomVisionPredictionClient AuthenticatePrediction(string endpoint, string predictionKey)
        {

            // Create a prediction endpoint, passing in the obtained prediction key
            CustomVisionPredictionClient predictionApi = new CustomVisionPredictionClient(new ApiKeyServiceClientCredentials(predictionKey))
            {
                Endpoint = endpoint

            };
            return predictionApi;
        }
        #endregion

        #region Helpers for Detection

        private byte[] GetBytesfromImagedata(string imagepath)
        {
            FileStream fs = new FileStream(imagepath, FileMode.Open, FileAccess.Read);
            BinaryReader reader = new BinaryReader(fs);
            return reader.ReadBytes((int)fs.Length);
        }

        private System.Drawing.Image resizeImage(System.Drawing.Image imgToResize, Size size)
        {
            //Get the image current width  
            int sourceWidth = imgToResize.Width;
            //Get the image current height  
            int sourceHeight = imgToResize.Height;
            float nPercent = 0;
            float nPercentW = 0;
            float nPercentH = 0;
            //Calulate  width with new desired size  
            nPercentW = ((float)size.Width / (float)sourceWidth);
            //Calculate height with new desired size  
            nPercentH = ((float)size.Height / (float)sourceHeight);
            if (nPercentH < nPercentW)
                nPercent = nPercentH;
            else
                nPercent = nPercentW;
            //New Width  
            int destWidth = (int)(sourceWidth * nPercent);
            //New Height  
            int destHeight = (int)(sourceHeight * nPercent);
            Bitmap b = new Bitmap(destWidth, destHeight);
            Graphics g = Graphics.FromImage((System.Drawing.Image)b);
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            // Draw image with new width and height  
            g.DrawImage(imgToResize, 0, 0, destWidth, destHeight);
            g.Dispose();
            return (System.Drawing.Image)b;
        }

        private System.Drawing.Image resizeImage(string path,
                         int width, int height)
        {
            System.Drawing.Image image = System.Drawing.Image.FromFile(path);

            System.Drawing.Image thumbnail = new Bitmap(width, height);
            System.Drawing.Graphics graphic =
                         System.Drawing.Graphics.FromImage(thumbnail);

            graphic.InterpolationMode = InterpolationMode.HighQualityBicubic;
            graphic.SmoothingMode = SmoothingMode.HighQuality;
            graphic.PixelOffsetMode = PixelOffsetMode.HighQuality;
            graphic.CompositingQuality = CompositingQuality.HighQuality;

            graphic.DrawImage(image, 0, 0, width, height);

            System.Drawing.Imaging.ImageCodecInfo[] info =
                             ImageCodecInfo.GetImageEncoders();
            EncoderParameters encoderParameters;
            encoderParameters = new EncoderParameters(1);
            encoderParameters.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality,
                             100L);
            return thumbnail;
        }


        #endregion

        #region Calling the Azure Custom Vision Models/API
        private async Task SearchTemplatesAzureAsyncModelCall()
        {
            string endpoint = "https://centralindia.api.cognitive.microsoft.com/";
            string apiKey = "babb63d3a4a547a58b8d8df5182e76c7";
            var credential = new AzureKeyCredential(apiKey);
            string predictionKey = "babb63d3a4a547a58b8d8df5182e76c7";

            var imageFile = Path.GetFullPath(ImagePath);
            Guid project = new Guid("4e8da76c-08d3-43a4-bb94-83f464bbf76c");
            CustomVisionPredictionClient predictionApi = AuthenticatePrediction(endpoint, predictionKey);

            using (var stream = File.OpenRead(imageFile))
            {
                var result = await predictionApi.DetectImageWithHttpMessagesAsync(project, "BetterImages", stream);

                var Json = result.Body;
                ObjectDetectionJSON = result.Body;
                stream.Close();
            }

        }
        private async Task SearchTemplatesAzureAsyncAPICall()
        {
            string endpoint = "https://centralindia.api.cognitive.microsoft.com/customvision/v3.0/Prediction/4e8da76c-08d3-43a4-bb94-83f464bbf76c/detect/iterations/BetterImages/image";
            var imageFile = Path.GetFullPath(ImagePath);
            string predictionKey = "babb63d3a4a547a58b8d8df5182e76c7";
            var client = new HttpClient();
            client.DefaultRequestHeaders.Add("Prediction-Key", predictionKey);

            byte[] Imagedata = GetBytesfromImagedata(imageFile);
            using (var content = new ByteArrayContent(Imagedata))
            {

                content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
                using (HttpResponseMessage response = await client.PostAsync(endpoint, content))
                {
                    if (response.IsSuccessStatusCode)
                    {
                        var ObjectDetectionJSONLocal = response.Content.ReadAsStringAsync();
                    }
                }


            }
        }
        #endregion

        #region Azure Form Identifier
        private async Task CallAzureFormIdentifier()
        {
            var Formrecogniserclient = AuthenticateClient();
            var Response = await Formrecogniserclient.StartRecognizeContentAsync(OpenFile(ImagePath)).WaitForCompletionAsync();
            FormIdentifierJSON = Response.GetRawResponse().Content.ToString();

        }

        #endregion

        #region Initial Trial : Non-ML Based solution
        //private List<Rectangle> SearchTemplates(Image<Bgr, byte> image, Image<Bgr, byte> template1, List<Rectangle> rectangles, Dictionary<string, List<Rectangle>> ControlsAndMatches, string control)
        //{

        //    Image<Bgr, byte> source = image; // Bigger image

        //    Image<Bgr, byte> template = template1; // smaller image
        //    Image<Bgr, byte> imageToShow = source.Copy();


        //    const double minTreshold = 0.7;
        //    double actualTreshold = 1;

        //    while (actualTreshold > minTreshold)
        //    {
        //        Image<Gray, float> result = imageToShow.MatchTemplate(template, TemplateMatchingType.CcoeffNormed);
        //        double[] minValues, maxValues;
        //        Point[] minLocations, maxLocations;

        //        result.MinMax(out minValues, out maxValues, out minLocations, out maxLocations);
        //        actualTreshold = maxValues[0];
        //        if (maxValues[0] > minTreshold)
        //        {
        //            // This is a match. Do something with it, for example draw a rectangle around it.

        //            Rectangle match = new Rectangle(maxLocations[0], template.Size);
        //            imageToShow.Draw(match, new Bgr(Color.Red), 3);
        //            if (!rectangles.Contains(match))
        //            {
        //                rectangles.Add(match);

        //            }
        //            else
        //                break;

        //        }
        //    }
        //    ControlsAndMatches.Add(control, rectangles);
        //    return rectangles;
        //    // StashImage.Source is an System.Windows.Controls.Image object in my app

        //}

        #endregion

        #region Get Prediction

        private void GetPredictionBlock()
        {
            var JSONData = JObject.Parse(FormIdentifierJSON);
            PredictionBlock predictionBlocks = JSONData.ToObject<PredictionBlock>();
            readResult = predictionBlocks.analyzeResult.readResults[0];
        }

        #endregion 

        #region Form Design helpers

        private Dictionary<string,string> GetAttributes()
        {
            Dictionary<string,string> attributes = new Dictionary<string,string>();
            GetPredictionBlock();
             var FormName = readResult.lines[0].text;
            attributes.Add("Name", "FRM" + FormName.Replace(" ", "").Substring(0, 4));
            attributes.Add("ModuleID", "FRM" + FormName.Replace(" ", "").Substring(0, 4));
            attributes.Add("TableName", "Table" + FormName.Replace(" ","").ToLower());
            attributes.Add("Header", FormName);
            return attributes;
        }

        private bool IgnoreControl(string Name)
        {
            DateTime datetime;
            string[] splitarray = Name.Split(' ');
            if (DateTime.TryParse(splitarray[splitarray.Length - 1], out datetime))
                return true;

            return Name.Contains("Submit") || Name.Contains("Ball In Court") || Name.Contains("0.00")
                   || Name.Contains("Save") || Name.Contains("Cancel") || Name.Contains("Publish")
                   || Name.Contains("GENERAL") || Name.Contains("WORKFLOW") || Name.Contains("Clear")
                   || Name.Equals("Add") || Name.Contains("Edit") || Name.Contains("Delete")
                   || Name.Contains("No records") || Name.Contains("Audit Log") || Name.Contains("...")
                   || Name.Contains("Back") || Name.Contains("OTHERS") || Name.Length == 1;

        }

        #endregion

        #region Preparing XML Form
        public async Task PrepareXMLDocument()
        {
            List<string> listpageControls = new List<string>();
            var xml = new StringBuilder();
            await CallAzureFormIdentifier();
            await SearchTemplatesAzureAsyncModelCall();
            xml = PrepareBaseXMlDocument(xml);
            xml = PrepareHiddenSection(xml);
            xml = PrepareVisibleSection(xml);
            xml = FinishXMLDocument(xml);
            Dictionary<string, string> attributes = GetAttributes();
            using (FileStream fs = File.Create("C:\\Users\\SYED.HAZIM\\Pictures\\XML\\"+ attributes["Header"]+".xml"))
            {
                // Add some text to file    
                Byte[] xmlcontent = new UTF8Encoding(true).GetBytes(xml.ToString());
                fs.Write(xmlcontent, 0, xmlcontent.Length);
                fs.Close();
            }
            xmlFormPath = "C:\\Users\\SYED.HAZIM\\Pictures\\XML\\" + attributes["Header"] + ".xml";
            PrepareDataDictionaryWithScreenshot();

        }

        private StringBuilder PrepareBaseXMlDocument(StringBuilder xml)
        {
            Dictionary<string, string> attributes = GetAttributes();
            xml.AppendFormat(CultureInfo.CurrentCulture, "<Form\n");
            xml.AppendFormat(CultureInfo.CurrentCulture, "Name = " + "\"" + attributes["Name"] + "\"\n");
            xml.AppendFormat(CultureInfo.CurrentCulture, "ModuleID = " + "\"" + attributes["ModuleID"] + "\"\n");
            xml.AppendFormat(CultureInfo.CurrentCulture, "TableName = " + "\"" + attributes["TableName"] + "\"\n");
            xml.AppendFormat(CultureInfo.CurrentCulture, "DropIfExists =\"false\" \n");
            xml.AppendFormat(CultureInfo.CurrentCulture, "ParentModuleID = \"\"\n");//To Do - ParentModuleID;
            xml.AppendFormat(CultureInfo.CurrentCulture, "Header = " + "\"" + attributes["Header"] + "\"\n");
            xml.AppendFormat(CultureInfo.CurrentCulture, "PrimaryKeyName=\"ID\" \n");
            xml.AppendFormat(CultureInfo.CurrentCulture, "AllowAttachments=\""+ ContainsAttachment().ToString() + "\" \n>");
            return xml;
        }

        private StringBuilder PrepareHiddenSection(StringBuilder xml)
        {
            xml.AppendFormat(CultureInfo.CurrentCulture, "\n<Section Name=\"Section0\" Attributes=\"display: none\" >\n");
            xml.AppendFormat(CultureInfo.CurrentCulture, "<Control Name=\"ID\" Caption=\"ID\" DBType=\"Int Identity(1, 1)\" PrimaryKey=\"true\" Type=\"Hidden\" Value=\"{{_REQUEST: InstanceID}}\"/> \n");
            xml.AppendFormat(CultureInfo.CurrentCulture, "<Control Name=\"PID\" Caption=\"PID\" DBType=\"int\" Type=\"Hidden\" Value=\"{{_REQUEST: PID}}\" />\n");
            xml.AppendFormat(CultureInfo.CurrentCulture, "<Control Name=\"ParentID\" Caption=\"ParentID\" DBType=\"int\" FilterKey=\"true\" Type=\"Hidden\" Value=\"{{_REQUEST: ParentID}}\" />\n");
            xml.AppendFormat(CultureInfo.CurrentCulture, "<Control Name=\"AUR_ModifiedBy\" Caption=\"Modified By: \" DBType=\"nvarchar(4000)\" Type=\"Hidden\" Value=\"{{CURRENTUSERNAME}}\" ReEvaluate=\"true\" />\n");
            xml.AppendFormat(CultureInfo.CurrentCulture, "<Control Name=\"AUR_ModifiedOn\" Caption=\"Modified On: \" DBType=\"datetime\" Type=\"Hidden\" Value=\"{{CURRENTDATE}}\" ReEvaluate=\"true\" Format=\"Date\" />\n");
            xml.AppendFormat(CultureInfo.CurrentCulture, "</Section>\n");
            return xml;
        }

        private StringBuilder PrepareVisibleSection(StringBuilder xml)
        {
            int MainTabCount = FindTabCount(0);
            for (int i=MainTabCount; i< readResult.lines.Count;)
            {
                var block = readResult.lines[i];
                if (block.text.Contains("ATTACHMENTS"))
                {
                    i += FindUploadDocument(readResult,i);
                    continue;
                }
               
                if (IgnoreControl(block.text))
                {
                    if (!block.text.Equals("GENERAL INFORMATION"))
                    {
                        i++;
                        continue;
                    }

                }
                int nextUpper = FindNextUpper(readResult, i);
                int nextValidControl = FindNextValidControl(readResult, i);
                //List<int> RightControlList = FindRightControls();
                if (IsAllUpper(block.text))
                {
                    int tabsInTheMiddle = FindTabCount(i);
                    if (tabsInTheMiddle > 1)
                    {
                        i += tabsInTheMiddle;
                        continue;
                    }
                    if(block.text.Contains("LINKED OBJECTS"))
                    {
                        xml = PrepareXMLForLinkedObjects(xml);
                        int LinkedObjectContolcount = FindAdd(readResult, i, nextUpper);
                        i += LinkedObjectContolcount;
                        continue;
                    }
                    
                        int DGControlcount = FindAdd(readResult, i,nextUpper);
                        
                        //DGControlcount = IsDynamicGrid(readResult, i, xml, block.text);
                        if (DGControlcount > 0)
                        {
                            xml = GenerateDynamicGrid(xml, block, i, DGControlcount);
                            i += DGControlcount;
                            continue;
                        }
                        else
                        {
                            if ( nextValidControl < readResult.lines.Count)
                            {
                                if (AreControlsOnTheSameLine(readResult.lines[i ], readResult.lines[nextValidControl]))
                                {

                                    xml.AppendFormat(CultureInfo.CurrentCulture, "\n<Section Name=\"Section" + sectioncount.ToString() + "\" Caption=\"" + GetCaption(block.text) + "\" Orientation=\"Horizontal\" >\n");
                                    xml = PrepareLeftSection(xml,  i, nextUpper);
                                    xml = PrepareRightSection(xml, i, nextUpper);
                                    xml.AppendFormat(CultureInfo.CurrentCulture, "</Section>\n");
                                }
                                else
                                {
                                    xml.AppendFormat(CultureInfo.CurrentCulture, "\n<Section Name=\"Section" + sectioncount.ToString() + "\"  >\n");
                                    xml = PrepareSection(xml, block, i, nextUpper);
                                    xml.AppendFormat(CultureInfo.CurrentCulture, "</Section>\n");

                                }
                                i = (nextUpper);
                                continue;
                            }
                           
                        }
 
                   
                }
                int Tabs = FindTabCount(i);
                if (Tabs > 1)
                {
                    i += Tabs;
                    continue;
                }
                if ( nextValidControl < readResult.lines.Count)
                {
                    if (AreControlsOnTheSameLine(readResult.lines[i], readResult.lines[nextValidControl]))
                    {

                        xml.AppendFormat(CultureInfo.CurrentCulture, "\n<Section Name=\"Section" + sectioncount.ToString() + "\" Orientation=\"Horizontal\" >\n");
                        xml = PrepareLeftSection(xml, i, nextUpper);
                        xml = PrepareRightSection(xml, i, nextUpper);
                        xml.AppendFormat(CultureInfo.CurrentCulture, "</Section>\n");
                       
                    }
                    else
                    {
                        xml.AppendFormat(CultureInfo.CurrentCulture, "\n<Section Name=\"Section" + sectioncount.ToString() + "\"  >\n");
                        xml = PrepareSection(xml, block, i, nextUpper);
                        xml.AppendFormat(CultureInfo.CurrentCulture, "</Section>\n");

                    }
                    i = (nextUpper);
                    continue;
                }

            } 
            return xml;
        }

        private StringBuilder FinishXMLDocument(StringBuilder xml)
        {
            for (int i = 0; i < listpageControls.Count; i++)
            {
                listpageControls[i] = GetName(listpageControls[i]);
            }
            xml.AppendFormat(CultureInfo.CurrentCulture, "<ListColumns Columns =\"" + String.Join(",", listpageControls.ToArray()) + "\" Type=\"Visible\" />\n");
            xml.AppendFormat(CultureInfo.CurrentCulture, "</Form>");
            return xml;
        }

        #endregion

        #region Visible Section Design Helpers

        private StringBuilder PrepareLeftSection(StringBuilder xml, int index, int nextUpper)
        {
            sectioncount++;
            xml.AppendFormat(CultureInfo.CurrentCulture, "\n<Section Name=\"Section" + sectioncount.ToString() + "\"  >\n");
            for (int i=index; i< nextUpper;)
            {
                if(IsControlOntheLeftSide(readResult.lines[i]))
                {
                    string type = string.Empty;
                    if (i + 1 < readResult.lines.Count)
                    {
                        type = GetType(readResult.lines[i + 1].text);
                    }
                    int skip = 0;
                    xml = DecideTypeandGenerateXML(readResult.lines[i], xml, type, false, out skip);
                    i += skip;
                    listpageControls.Add(readResult.lines[i].text);
                    continue;
                }
                i++;
            }
            xml.AppendFormat(CultureInfo.CurrentCulture, "</Section>\n");
            return xml;
        }

        private StringBuilder PrepareRightSection(StringBuilder xml, int index, int nextUpper)
        {

            sectioncount++;
            xml.AppendFormat(CultureInfo.CurrentCulture, "\n<Section Name=\"Section" + sectioncount.ToString() + "\"  >\n");
            for (int i = index; i < nextUpper; )
            {
                if (!IsControlOntheLeftSide(readResult.lines[i]))
                {
                    string type = String.Empty;
                    if (i + 1 < readResult.lines.Count)
                    {
                        type = GetType(readResult.lines[i + 1].text);
                    }
                    int skip = 0;
                    xml = DecideTypeandGenerateXML(readResult.lines[i], xml, type, false, out skip);
                    listpageControls.Add(readResult.lines[i].text);
                    i += skip;
                    continue;                   
                }
                i++;
            }
            xml.AppendFormat(CultureInfo.CurrentCulture, "</Section>\n");
            return xml;
        }

        private StringBuilder PrepareSection(StringBuilder xml, PredictionBlock.Line block, int index, int nextUpper)
        {

            sectioncount++;
            for (int i = index; i < nextUpper; )
            {
               
                    string type = String.Empty;
                    if (i + 1 < readResult.lines.Count)
                    {
                        type = GetType(readResult.lines[i + 1].text);
                    }
                    int skip = 0;
                    xml = DecideTypeandGenerateXML(readResult.lines[i], xml, type, false, out skip);
                listpageControls.Add(readResult.lines[i].text);
                i += skip;
               
            }
            return xml;
        }


        #endregion

        #region Dynamic Grid Preparation

        private StringBuilder PrepareDynamicGridTemplate(StringBuilder xml, string DGName)
        {
            xml.AppendFormat(CultureInfo.CurrentCulture, "<DynamicGrid Name=\"" + GetName(DGName) + "\" TableName =\"DB" + GetName(DGName) + "\" PrimaryKeyName=\"ID\" Caption=\"" + GetCaption(DGName) + "\" >\n");
            xml.AppendFormat(CultureInfo.CurrentCulture, "<Column Name=\"ID\" Caption=\"ID\" DBType=\"Int Identity(1, 1)\" PrimaryKey=\"true\" Type=\"Hidden\" Value=\"{{_REQUEST: InstanceID}}\"/> \n");
            xml.AppendFormat(CultureInfo.CurrentCulture, "<Column Name=\"FID\" Caption=\"PID\" DBType=\"int\" ForeignKey=\"true\" Type=\"Hidden\"  />\n");
            xml.AppendFormat(CultureInfo.CurrentCulture, "<Column Name=\"AUR_ModifiedBy\" Caption=\"Modified By: \" DBType=\"nvarchar(4000)\" Type=\"Hidden\" Value=\"{{CURRENTUSERNAME}}\" ReEvaluate=\"true\" />\n");
            xml.AppendFormat(CultureInfo.CurrentCulture, "<Column Name=\"AUR_ModifiedOn\" Caption=\"Modified On: \" DBType=\"datetime\" Type=\"Hidden\" Value=\"{{CURRENTDATE}}\" ReEvaluate=\"true\" Formate=\"Date\" />\n");
            return xml;
        }

        private StringBuilder FinishDynamicGrid(StringBuilder xml)
        {
            xml.AppendFormat(CultureInfo.CurrentCulture, "</DynamicGrid>\n");
            return xml;
        }

        private StringBuilder GenerateDynamicGrid(StringBuilder xml, PredictionBlock.Line block,int index, int controlcount)
        {
            xml = PrepareDynamicGridTemplate(xml, block.text);
            for (int i = index+1; i < index + controlcount; i++ )
            {
                int skip = 0;
                xml = DecideTypeandGenerateXML(block, xml, string.Empty, true, out skip);
               
            }
            xml = FinishDynamicGrid(xml);
            return xml;
        }
        
        private int IsDynamicGrid(PredictionBlock.ReadResult blocks, int index, StringBuilder xml,string DGName)
        {
            int i = index;
            for (; i < readResult.lines.Count; )
            {
                var block = readResult.lines[i];
                var nextblock = readResult.lines[i + 1];
                bool DynamicGridxmlGenerated = false;
                int dynamicGridControlCount = 0;
                bool AreinDynamicGrid = ContainAtleastTwoControls(block, i, out dynamicGridControlCount);
                if (AreinDynamicGrid)
                {
                    if (index - i == 0)
                        xml = PrepareDynamicGridTemplate(xml, DGName);
                    int skip = 0;
                    for(int controlCount = i; controlCount < dynamicGridControlCount;controlCount++ )
                    {
                        xml = DecideTypeandGenerateXML(block, xml, string.Empty, true, out skip);
                    }
                    i += dynamicGridControlCount;
                    DynamicGridxmlGenerated = true;
                }
                if(!AreinDynamicGrid && index-i == 0)
                {
                    break;
                }
                if (DynamicGridxmlGenerated)
                {
                    int skip = 0;
                    xml = DecideTypeandGenerateXML(block, xml, string.Empty, true, out skip) ;
                    xml = FinishDynamicGrid(xml);

                }

            }

            return i - index ;
        }

    

        #endregion

        #region Visual Helpers (Control Grouping, Dynamic Grid and Attachment and Linked Records)


       private bool ContainsAttachment()
        {
           for(int i = readResult.lines.Count-1;i>=0;i--)
                if(readResult.lines[i].text.Contains("Attachments"))
                    return true;
            return false;
        }

        
        private bool AreControlsOnTheSameLine(PredictionBlock.Line currentblock, PredictionBlock.Line nextblock)
        {
            bool areonthesameline = nextblock.boundingBox[1] - currentblock.boundingBox[1] < Y_Distance
                                     && nextblock.boundingBox[3] - currentblock.boundingBox[3] < Y_Distance
                                     && nextblock.boundingBox[5] - currentblock.boundingBox[5] < Y_Distance
                                     && nextblock.boundingBox[7] - currentblock.boundingBox[7] < Y_Distance;

            return areonthesameline;
        }

        private bool IsControlOntheLeftSide(PredictionBlock.Line currentblock)
        {
            bool isControlOntheLeftSide =  currentblock.boundingBox[0] < 600
                                     &&  currentblock.boundingBox[2] < 600
                                     &&  currentblock.boundingBox[4] < 600
                                     &&  currentblock.boundingBox[6] < 600;

            return isControlOntheLeftSide;
        }

        private bool AreControlsVeryClose(PredictionBlock.Line currentblock, PredictionBlock.Line nextblock)
        {
            bool areVeryClose = nextblock.boundingBox[0] - currentblock.boundingBox[0] < 110
                                       && nextblock.boundingBox[2] - currentblock.boundingBox[2] < 110
                                       && nextblock.boundingBox[4] - currentblock.boundingBox[4] < 110
                                       && nextblock.boundingBox[6] - currentblock.boundingBox[6] < 110;

            return areVeryClose;
        }

        private int FindAdd(PredictionBlock.ReadResult blocks, int index, int nextUpper)
        {
            int i = index+1;
            for (; i <nextUpper;)
            {
                if (IsAllUpper(blocks.lines[i].text))
                    return -1;
                if (blocks.lines[i].text.Equals("Add"))
                    break;
                else
                    i++;
            }
            int nextupper = FindNextUpper(blocks, index);

            return i - index > nextupper ? nextupper : i-index;
        }

        private int FindNextUpper(PredictionBlock.ReadResult blocks, int index)
        {
            int i = index + 1;
            for (; i < blocks.lines.Count; i++)
            {
                DateTime dateTime;
                if (IsAllUpper(blocks.lines[i].text) && !DateTime.TryParse(blocks.lines[i].text,out dateTime) && !IgnoreControl(blocks.lines[i].text))
                    break;

            }
            return i;
        }

        private int FindNextValidControl(PredictionBlock.ReadResult blocks, int index)
        {
            int i = index + 1;
            for (; i < blocks.lines.Count; i++)
            {
                if (!IgnoreControl(blocks.lines[i].text))
                    break;

            }
            return i;
        }

        private int  FindUploadDocument(PredictionBlock.ReadResult blocks, int index)
        {
            int i = index + 1;
            for (; i < blocks.lines.Count; i++)
            {
                if ((blocks.lines[i].text.Contains("Upload Document")))
                    break;

            }
            return i;
        }
        
        private List<int> FindRightControls()
        {
            List<int> indices = new List<int>();
            int i =0;
            for (; i < readResult.lines.Count; i++)
            {
                if (!IsControlOntheLeftSide(readResult.lines[i]))
                    indices.Add(i);

            }
            return indices;
        }

        private StringBuilder PrepareXMLForLinkedObjects(StringBuilder xml)
        {
            Dictionary<string, string> attributes = GetAttributes();
            string DGName = attributes["ModuleID"] + "LinkTo";
            xml.AppendFormat(CultureInfo.CurrentCulture, "<LinkToGrid Name=\"" + GetName(DGName) + "\" TableName =\"DB" + GetName(DGName)+"Data" + "\" PrimaryKeyName=\"ID\" Caption=\"Linked Records\" DescriptionColumnName=\"Description\" TypeColumnName=\"Type\" ModuleIdColumnName=\"ModuleId\" InstanceIdColumnName=\"InstanceId\" ModuleNameColumnName=\"SourceForm\" PIDColumnName=\"PID\" ParentIdColumnName=\"ParentId\" Width=\"960px\" ContractNameColumnName=\"ContractName\" ShowProjectFilter=\"false\" ShowContractFilter=\"true\" NotesColumnName=\"Notes\" >\n");
            xml.AppendFormat(CultureInfo.CurrentCulture, "<Column Name=\"ID\" Caption=\"ID\" DBType=\"Int Identity(1, 1)\" PrimaryKey=\"true\" Type=\"Hidden\" Value=\"{{_REQUEST: InstanceID}}\"/> \n");
            xml.AppendFormat(CultureInfo.CurrentCulture, "<Column Name=\"FormInstanceIDLinkTo\" Type=\"Hidden\" DBType=\"Int\" ForeignKey=\"true\" />");
            xml.AppendFormat(CultureInfo.CurrentCulture, "<Column Name=\"AUR_ModifiedBy\" Caption=\"Modified By: \" DBType=\"nvarchar(4000)\" Type=\"Hidden\" Value=\"{{CURRENTUSERNAME}}\" ReEvaluate=\"true\" />\n");
            xml.AppendFormat(CultureInfo.CurrentCulture, "<Column Name=\"AUR_ModifiedOn\" Caption=\"Modified On: \" DBType=\"datetime\" Type=\"Hidden\" Value=\"{{CURRENTDATE}}\" ReEvaluate=\"true\" Formate=\"Date\" />\n");
            xml.AppendFormat(CultureInfo.CurrentCulture, "<Column Name=\"Description\" Caption=\"Record Identifier\" Type=\"Link\" DBType=\"nvarchar(4000)\"/>\n");
            xml.AppendFormat(CultureInfo.CurrentCulture, "<Column Name=\"Type\" Caption=\"Type\" Type=\"Hidden\" DBType=\"nvarchar(50)\" PickerFilterKey=\"true\" />\n");
            xml.AppendFormat(CultureInfo.CurrentCulture, "<Column Name=\"ModuleId\" Caption=\"Module\" Type=\"Hidden\" DBType=\"nvarchar(50)\" PickerFilterKey=\"true\" />\n");
            xml.AppendFormat(CultureInfo.CurrentCulture, "<Column Name=\"InstanceId\" Caption=\"InstanceId\" Type=\"Hidden\" DBType=\"Int\" PickerFilterKey=\"true\" />\n");
            xml.AppendFormat(CultureInfo.CurrentCulture, "<Column Name=\"SourceForm\" Caption=\"Source Form\" Type=\"Display\" DBType=\"nvarchar(250)\" />\n");
            xml.AppendFormat(CultureInfo.CurrentCulture, "<Column Name=\"PID\" Caption=\"PID\" Type=\"Hidden\" DBType=\"Int\" />\n");
            xml.AppendFormat(CultureInfo.CurrentCulture, "<Column Name=\"ParentId\" Caption=\"ParentId\" Type=\"Hidden\" DBType=\"Int\" />\n");
            xml.AppendFormat(CultureInfo.CurrentCulture, " <Column Name=\"ContractName\" Caption=\"Contract Name\" Type=\"Display\" DBType=\"nvarchar(250)\"/>\n");
            xml.AppendFormat(CultureInfo.CurrentCulture, "<Column Name=\"Notes\" Caption=\"Notes\" DBType=\"nvarchar(max)\" Type=\"TextArea\" Attributes=\"min - width: 200px; width: 200px; max - width: 200px; \" />\n");
            return xml;
        }

        

        private int FindTabCount(int index)
        {
            int tabCount = 1;
            for(int i =index;i< readResult.lines.Count;i++)
            {
                if (AreControlsOnTheSameLine(readResult.lines[i], readResult.lines[i + 1]) && AreControlsVeryClose(readResult.lines[i], readResult.lines[i + 1]))
                    tabCount++;
                else
                    break;
            }
            return tabCount;

        }
        private bool ContainAtleastTwoControls(PredictionBlock.Line block, int index, out int controlsonthesameline)
        {
             controlsonthesameline = 1;
                bool areonthesameline = false;
                if(index + 1 < readResult.lines.Count )
                {
                    int k = 1;
                    while(index + k< readResult.lines.Count)
                    {
                        if(AreControlsOnTheSameLine(readResult.lines[index],readResult.lines[index+k]))
                        {
                            controlsonthesameline++;
                            k++;
                            areonthesameline = true;
                        }
                        else
                        {
                            break;
                        }
                    }

                }
               
            
            return areonthesameline;
        }


        #endregion

        #region Defining Rules based On Captions

        private string GetType(string caption)
        {
            DateTime date;
            string[] splitarray = caption.Split(':');
            if (caption.Contains("Date") || caption.Contains(" On :") || DateTime.TryParse(splitarray[splitarray.Length-1],out date))
                return "Date";
            if (caption.Contains("None"))
                return "DateNull";
            if (VisibleDate(caption))
                return "DateWithToday";
            if (caption.Contains("Select") )
                return "DropDownList";
            if (caption.Contains("Notes") || caption.Contains("Description"))
                return "TextArea";
            if (caption.Contains("Active"))
                return "CheckBox";
            if (caption.Contains("0.00") || caption.Contains("(S)") || caption.Contains("$"))
                return "Numeric";
            if (caption.Contains("Auto Generated"))
                return "AutoIncrement";
            if (caption.Contains("Auto Calculated"))
                return "NumericDisplay";
            if (caption.Contains("Auto Populated"))
                return "DisplayTextBox";
            if (caption.Contains("..."))
                return "Picker";
            if (caption.Contains("Clear"))
                return "Picker";
            return "TextBox";
        }
       
        private bool VisibleDate(string caption)
        {
            var splits = caption.Split(' ');
            foreach (var part in splits)
            {
                DateTime dateValue;
                if (!DateTime.TryParse(part, out dateValue))
                    return false;
            }
            return true;
        }

        private string GetName(string name)
        {
            return name.Split(':')[0].Replace(" ",string.Empty).Replace(".",string.Empty).Replace("-",string.Empty);
        }

        private string GetCaption(string name)
        {
            string caption =  name.Split(':')[0];
            if (caption.Split(' ')[0].Length == 1)           
                caption = string.Join(" ", caption.Split(' ').Skip(1));
             return caption;
        }
        
        #endregion

        #region Deciding Control Type based on Caption

        private StringBuilder DecideTypeandGenerateXML(PredictionBlock.Line block, StringBuilder xml,string nexttype,bool IsGrid,out int skip)
        {
            if (IgnoreControl(block.text))
            {
                skip = 1;
                return xml;
            }
            string type = GetType(block.text);
            string Name = GetName(block.text);
            string Caption = GetCaption(block.text);
            skip = 0;
            if (!IsGrid)
            {
                bool isXMLSet = false;
                switch (nexttype)
                {
                    case "DateWithToday":
                        xml.AppendFormat(CultureInfo.CurrentCulture, "<Control Name=\"" + Name + "\" Caption=\" " + Caption + "\" DBType=\"datetime\" Type=\"Date\" Value=\"{{CURRENTDATE}}\" ReEvaluate=\"true\" Format=\"Date\" />\n");
                        isXMLSet = true;
                        skip = 2;
                        break;
                    case "DateNull":
                        xml.AppendFormat(CultureInfo.CurrentCulture, "<Control Name=\"" + Name + "\" Caption=\" " + Caption + "\" DBType=\"datetime\" Type=\"Date\" AllowNull=\"true\" Format=\"Date\" />\n");
                        isXMLSet = true;
                        skip = 2;
                        break;
                    case "DropDownList":
                        isXMLSet = true;
                        xml.AppendFormat(CultureInfo.CurrentCulture, "<Control Name=\"" + Name + "\" Caption=\"" + Caption + "\" DBType=\"int\" Type=\"DropDownList\"  ListItems=\"Select One:_DBNULL_: 0\" />\n");
                        skip = 2;
                        break;
                    case "Picker":
                        isXMLSet = true;
                        xml.AppendFormat(CultureInfo.CurrentCulture, "<Control Name=\"" + Name + "Set" + "\" Caption=\"" + Caption + "\" Type=\"Set\"   >\n");
                        xml.AppendFormat(CultureInfo.CurrentCulture, "<Control Name=\"" + Name  + "\" Caption=\"\" DBType=\"nvarchar(255)\" Type=\"TextBox\"   />\n");
                        xml.AppendFormat(CultureInfo.CurrentCulture, "<Control Name=\"" + Name + "PickerTrigger" + "\" Caption=\"" + Caption + "\" Type=\"PickerTrigger\"   PickerName=\""+ Name +"Picker\" PickerButtonText=\"...\" EnableMultiSelect=\"false\" PickerButtonWidth=\"30px\" />\n</Control>\n");
                        skip = 2;
                        break;
                    case "Numeric":
                        xml.AppendFormat(CultureInfo.CurrentCulture, "<Control Name=\"" + Name + "\" Caption=\"" + Caption + "\" DBType=\"Numeric(14,2)\" Type=\"Numeric\" Format=\"Amount\"  />\n");
                        skip = 2;
                        isXMLSet = true;
                        break;
                }
                if (isXMLSet)
                    return xml;


                switch (type)
                {

                    case "Date":
                        xml.AppendFormat(CultureInfo.CurrentCulture, "<Control Name=\"" + Name + "\" Caption=\" " + Caption + "\" DBType=\"datetime\" Type=\"Date\" AllowNull=\"true\" Format=\"Date\" />\n");
                        skip = 1;
                        break;
                    case "TextArea":
                        xml.AppendFormat(CultureInfo.CurrentCulture, "<Control Name=\"" + Name + "\" Caption=\"" + Caption + "\" DBType=\"nvarchar(4000)\" Type=\"TextArea\"  />\n");
                        skip = 1;
                        break;
                    case "CheckBox":
                        xml.AppendFormat(CultureInfo.CurrentCulture, "<Control Name=\"" + Name + "\" Caption=\"" + Caption + "\" DBType=\"bit\" Type=\"CheckBox\" IsSelected=\"true\"  />\n");
                        skip = 1;
                        break;
                    case "Numeric":
                        xml.AppendFormat(CultureInfo.CurrentCulture, "<Control Name=\"" + Name + "\" Caption=\"" + Caption + "\" DBType=\"Numeric(14,2)\" Type=\"Numeric\" Format=\"Amount\"  />\n");
                        skip = 1;
                        break;
                    case "NumericDisplay":
                        xml.AppendFormat(CultureInfo.CurrentCulture, "<Control Name=\"" + Name + "\" Caption=\"" + Caption + "\" DBType=\"Numeric(14,2)\" Type=\"Display\" Format=\"Amount\"  />\n");
                        skip = 1;
                        break;
                    case "DisplayTextBox":
                        xml.AppendFormat(CultureInfo.CurrentCulture, "<Control Name=\"" + Name + "\" Caption=\"" + Caption + "\" DBType=\"nvarchar(255)\" Type=\"Display\"   />\n");
                        skip = 1;
                        break;
                    case "TextBox":
                        xml.AppendFormat(CultureInfo.CurrentCulture, "<Control Name=\"" + Name + "\" Caption=\"" + Caption + "\" DBType=\"nvarchar(255)\" Type=\"TextBox\"   />\n");
                        skip = 1;
                        break;
                    case "AutoIncrement":
                        xml.AppendFormat(CultureInfo.CurrentCulture, "<AutoIncrement Name=\"" + Name + "\" Caption=\"" + Caption + "\" DBType=\"nvarchar(255)\"  Seed=\"1\" Interval=\"1\"  Type=\"Display\"   />\n");
                        skip = 1;
                        break;
                    case "":
                        skip = 0;
                        break;
                    case "DropDownList":
                        xml.AppendFormat(CultureInfo.CurrentCulture, "<Control Name=\"" + Name + "\" Caption=\"" + Caption + "\" DBType=\"int\" Type=\"DropDownList\"  ListItems=\"Select One:_DBNULL_: 0\" />\n");
                        skip = 1;
                        break;

                }
            }
            else
            {
                bool isXMLSet = false;
                switch (nexttype)
                {
                    case "DateWithToday":
                        xml.AppendFormat(CultureInfo.CurrentCulture, "<Column Name=\"" + Name + "\" Caption=\" " + Caption + "\" DBType=\"datetime\" Type=\"Date\" Value=\"{{CURRENTDATE}}\" ReEvaluate=\"true\" Format=\"Date\" />\n");
                        isXMLSet = true;
                        skip = 2;
                        break;
                    case "DateNull":
                        xml.AppendFormat(CultureInfo.CurrentCulture, "<Column Name=\"" + Name + "\" Caption=\" " + Caption + "\" DBType=\"datetime\" Type=\"Date\" AllowNull=\"true\" Format=\"Date\" />\n");
                        isXMLSet = true;
                        skip = 2;
                        break;
                    case "DropDownList":
                        isXMLSet = true;
                        xml.AppendFormat(CultureInfo.CurrentCulture, "<Column Name=\"" + Name + "\" Caption=\"" + Caption + "\" DBType=\"int\" Type=\"DropDownList\"  ListItems=\"Select One:_DBNULL_: 0\" />\n");
                        skip = 2;
                        break;

                }
                if (isXMLSet)
                    return xml;


                switch (type)
                {

                    case "Date":
                        xml.AppendFormat(CultureInfo.CurrentCulture, "<Column Name=\"" + Name + "\" Caption=\" " + Caption + "\" DBType=\"datetime\" Type=\"Date\" AllowNull=\"true\" Format=\"Date\" />\n");
                        skip = 1;
                        break;
                    case "TextArea":
                        xml.AppendFormat(CultureInfo.CurrentCulture, "<Column Name=\"" + Name + "\" Caption=\"" + Caption + "\" DBType=\"nvarchar(4000)\" Type=\"TextArea\"  />\n");
                        skip = 1;
                        break;
                    case "CheckBox":
                        xml.AppendFormat(CultureInfo.CurrentCulture, "<Column Name=\"" + Name + "\" Caption=\"" + Caption + "\" DBType=\"bit\" Type=\"CheckBox\" IsSelected=\"true\"  />\n");
                        skip = 1;
                        break;
                    case "Numeric":
                        xml.AppendFormat(CultureInfo.CurrentCulture, "<Column Name=\"" + Name + "\" Caption=\"" + Caption + "\" DBType=\"Numeric(14,2)\" Type=\"Numeric\" Format=\"Amount\"  />\n");
                        skip = 1;
                        break;
                    case "TextBox":
                        xml.AppendFormat(CultureInfo.CurrentCulture, "<Column Name=\"" + Name + "\" Caption=\"" + Caption + "\" DBType=\"nvarchar(255)\" Type=\"TextBox\"   />\n");
                        skip = 1;
                        break;
                   

                }
               
            }
            return xml;
        }

        #endregion

        #region Data Dictionary Design Helpers

        bool IsAllUpper(string input)
        {
            for (int i = 0; i < input.Length; i++)
            {
                if (Char.IsLetter(input[i]) && !Char.IsUpper(input[i]))
                    return false;
            }
            return true;
        }

        #endregion

        #region Designing the Data Dictionary

        private void PrepareDataDictionaryWithScreenshot()
        {
            string FileDirectory = "C:\\Users\\SYED.HAZIM\\Pictures\\Saved Pictures\\";
            string fileName = FileDirectory + "Data Dictionary of " + readResult.lines[0].text + ".docx";
            var doc = DocX.Create(fileName);

            Image img = doc.AddImage(ImagePath);
            Picture p = img.CreatePicture();
            Xceed.Document.NET.Paragraph par = doc.InsertParagraph("");
            par.AppendPicture(p);
            doc.InsertSectionPageBreak();



            Table t = doc.AddTable(readResult.lines.Count, 9);

            t.Alignment = Alignment.center;
            t.Design = TableDesign.ColorfulGridAccent2;

            List<string> TableHeaders = new List<string>() { "Field Name", "Type", "Size", "Default", "Description", "Mandatory", "Source", "Validation" };
            int i = 0;
            foreach (string name in TableHeaders)
            {
                t.Rows[0].Cells[i++].Paragraphs.First().Append(name);
            }
            i = 1;
            var blocks = readResult.lines;
            int linecount = readResult.lines.Count;
            for (int index = 1; index < linecount; index++)
            {
                if (IgnoreControl(blocks[index].text))
                {
                    t.RemoveRow();
                    continue;
                }
                if(blocks[index].text.Contains("Attachments"))
                {
                    int j;
                    for (j=index; j< readResult.lines.Count;j++)
                    {
                        t.RemoveRow();
                        if (readResult.lines[j].text.Contains("Upload Document"))
                            break;

                    }
                    index = j;
                    continue;
                }
                else
                {
                    t.Rows[i].Cells[0].Paragraphs.First().Append(GetCaption(blocks[index].text));
                    t.Rows[i].Cells[1].Paragraphs.First().Append(string.Empty);
                    t.Rows[i].Cells[2].Paragraphs.First().Append(string.Empty);
                    t.Rows[i].Cells[3].Paragraphs.First().Append(string.Empty);
                    t.Rows[i].Cells[4].Paragraphs.First().Append(string.Empty);
                    t.Rows[i].Cells[5].Paragraphs.First().Append(string.Empty);
                    t.Rows[i].Cells[6].Paragraphs.First().Append(string.Empty);
                    t.Rows[i].Cells[7].Paragraphs.First().Append(string.Empty);
                    t.Rows[i].Cells[8].Paragraphs.First().Append(string.Empty);
                    i++;
                }

            }

            doc.InsertTable(t);
            doc.Save();
            wordDocumentPath = fileName;
        }

        #endregion
    }
}