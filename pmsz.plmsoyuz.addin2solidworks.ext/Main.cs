using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ProgramSoyuz.PLM.Scripting;
using SolidWorks.Interop.sldworks;
using ProgramSoyuz.AddinSupport;
using pmsz.plmsoyuz.addin2solidworks.Documents;
using ProgramSoyuz.AddinSupport.DataModel;
using System.Diagnostics;

using System.Windows.Forms;
using SolidWorks.Interop.swconst;
using System.Collections;

using pmsz.plmsoyuz.addin2solidworks.ext.UI;

///  Собранный dll необходимо положить в папку аддина на клиенте.
///  обычно c:\Program Files (x86)\Програмсоюз\PLM Framework\Configuration\AddIns\pmsz.plmsoyuz.addin2solidworks\
///  следует помнить что эта папка синхронизируется с сервером, по-этому на время разработки нужно отключать PLMHelper что бы он не стирал файл при каждом подключении к серверу.
///  что бы аддин разошелся по всем клиентам(когда достигнут рабочий результат)  его необходимо положить на сервере в папку автоапдейта: c:\Program Files (x86)\Програмсоюз\PLM Framework\Server\ClientAutoUpdate\Configuration\AddIns\pmsz.plmsoyuz.addin2solidworks\

namespace pmsz.plmsoyuz.addin2solidworks.ext
{
    [PLMConfigurationModule]
    public class SWAddinExtension : ISolidWorksClientModule, ICustomModuleFunction, ISolidWorksClientModule2
    {
        ISldWorks _swApp;
        IPlmAddin _swAddin;

        private class ValElem
        {
            public String link;
            public String val;
        }

        private class PositionTT
        {
            public double x, y, z;
        }

        private class NoteTT
        {
            public String NoteText;
            public Double curheight;
        }

        private List<ValElem> arrLinks = new List<ValElem>();
        private List<String> arrTT = new List<String>();
        private List<NoteTT> notes = new List<NoteTT>();
        const string ttname = @"TechTr";
        //int ToolbarId;
        // Размеры листа
        private double sh_width = 0, sh_height = 0;
        // Размеры основной надписи с учётом рамки
        const double c_width = 0.19, c_height = 0.07, c_topmargin = 0.025;

        #region Implementation of IAddinClientModule

        public void OnUnload()
        {
            AddinUIworks.ClearCommandMgr();
            // int addinID = SwAddin.Instance.SWAddinID;
            // _swApp.RemoveToolbar2(addinID, ToolbarId);
        }

        public string ModuleId
        {
            get { return @"pmsz.plmsoyuz.addin2solidworks.ext"; }
        }

        public void OnAfterUICreated()
        {
            _swApp = addinCallback.SWApp;
            _swAddin = addinCallback.Addin;

            // Здесь можно добавить комманды в солид, подписаться на события солида и т.п.

            AddinUIworks.RefreshUI();

            // Создание пункта меню "Обновить ТТ"
            /* int addinID = SwAddin.Instance.SWAddinID;
            var strResourcePath = AddinUIworks.GetResourcePath() + "Logo16.bmp";
            var ToolbarImgLgPath = AddinUIworks.GetResourcePath() + "ToolbarLarge.bmp";
            var ToolbarImgSmPath = AddinUIworks.GetResourcePath() + "ToolbarSmall.bmp";
            // Название пункта меню - "Союз-PLM " - обязательно с пробелом в конце!!!
            int cmdIndex0 = _swApp.AddMenuItem4((int)swDocumentTypes_e.swDocDRAWING, addinID, "Обновить ТТ@Союз-PLM ", 13, "CallBackFunction(RefreshTT)", "", "", strResourcePath);
            
            // Это не срабатывает!!! Панель пропадает после загрузки чертежа.
            ToolbarId = _swApp.AddToolbar4(addinID, "Союз-PLM2", ToolbarImgSmPath, ToolbarImgLgPath, 0, (int)swDocumentTypes_e.swDocDRAWING); 
            bool added = _swApp.AddToolbarCommand2(addinID, ToolbarId, 11, "CallBackFunction(RefreshTT)", "", "", "Обновить ТТ");  */
        }

/*        public void OnLoad(ISolidWorksCallback addinCallback)
        {
            _swApp = addinCallback.SWApp;
            _swAddin = addinCallback.Addin;

            // Здесь можно добавить комманды в солид, подписаться на события солида и т.п.

            AddinUIworks.RefreshUI();

            // Создание пункта меню "Обновить ТТ"
            /* int addinID = SwAddin.Instance.SWAddinID;
            var strResourcePath = AddinUIworks.GetResourcePath() + "Logo16.bmp";
            var ToolbarImgLgPath = AddinUIworks.GetResourcePath() + "ToolbarLarge.bmp";
            var ToolbarImgSmPath = AddinUIworks.GetResourcePath() + "ToolbarSmall.bmp";
            // Название пункта меню - "Союз-PLM " - обязательно с пробелом в конце!!!
            int cmdIndex0 = _swApp.AddMenuItem4((int)swDocumentTypes_e.swDocDRAWING, addinID, "Обновить ТТ@Союз-PLM ", 13, "CallBackFunction(RefreshTT)", "", "", strResourcePath);
            
            // Это не срабатывает!!! Панель пропадает после загрузки чертежа.
            ToolbarId = _swApp.AddToolbar4(addinID, "Союз-PLM2", ToolbarImgSmPath, ToolbarImgLgPath, 0, (int)swDocumentTypes_e.swDocDRAWING); 
            bool added = _swApp.AddToolbarCommand2(addinID, ToolbarId, 11, "CallBackFunction(RefreshTT)", "", "", "Обновить ТТ");  */
        } 

        public void RefreshTT()
        {
            AddinDocument addinDoc = _swAddin.GetActiveDoc();
            if (addinDoc != null)
            {
                string swDoc = addinDoc.FullFileName;
                InfoObject ioVersion = addinDoc.PlmVersion;
                PrepareTT(swDoc, ioVersion);
            }
        }

        public void CopyLink()
        {
            Clipboard.Clear();
            AddinDocument addinDoc = _swAddin.GetActiveDoc();
            if (addinDoc == null) return;
            string swDoc = addinDoc.FullFileName;
            ModelDoc2 swModelDoc = SWHelper.GetDocByFileName(swDoc); // нужно учитывать что браться на редактирование может документ не открытый в PLM. тогда здесь будет null и нужно либо открывать документ, либо работать через DMDocument
            if (swModelDoc == null) return;

            SelectionMgr mSelectionMgr = swModelDoc.ISelectionManager;
            String sz = "";

            var tp = mSelectionMgr.GetSelectedObjectType3(1, -1);
            if (tp == (int)swSelectType_e.swSelDIMENSIONS)
            {
                IDisplayDimension sel = (IDisplayDimension)mSelectionMgr.GetSelectedObject6(1, 0);
                sz = sel.GetDimension2(0).FullName;
                Clipboard.SetText(sz);
            }

            /*  if (tp == (int)swSelectType_e.swSelNOTES)
              {
                  INote sel = (INote)mSelectionMgr.GetSelectedObject6(1, 0); mSelectionMgr.IGetSelectedObjectsComponent2
                  sel

                  Annotation swAnn = (Annotation)sel.GetAnnotation();
                  swAnn.
                  ModelDocExtension mde = (ModelDocExtension)swModel.Extension;
                  int id = mde.GetObjectId(swAnn);

                  swFeature.GetNameForSelection(out sz);
              }*/
            //if (sz.Length != 0) MessageBox.Show(sz);
        }

        /// <summary>
        /// Вызывается в процессе сохранения документа в PLM
        /// </summary>
        /// <param name="swDoc">документ</param>
        /// <param name="ioVersion">ио версии в плм</param>
        public void OnCheckIn(string swDoc, InfoObject ioVersion)
        {
            //MessageBox.Show("CheckIn");
            SWDocument addinDoc = AddinDocumentCache.ReturnDocument(swDoc, null) as SWDocument; // SWDocument - обёртка документа  для использования аддином
            //addinDoc.GetStatus() - определение plm-статуса документа -не сохранен, сохранён, на редактировании, забблокирован. если документ сохранён в плм, свойства PlmDocument и PlmVersion будут содержать ссылки на ио соответственно документа и версии
            ModelDoc2 Doc = SWHelper.GetDocByFileName(swDoc); // нужно учитывать что сохраняться может документ не открытый в PLM. тогда здесь будет null и нужно либо открывать документ, либо работать через DMDocument
            if (Doc != null)
            {
                // Определяем тип документа. Пока работаем только с чертежами
                int docType = Doc.GetType();
                if (docType == (int)swDocumentTypes_e.swDocDRAWING)
                {
                    // Ищем в примечании ТТ по кэшу сохранённых ссылок на значения чертежа.
                    // Если находим, то ищем в чертеже соответствующий элемент и считываем его значение для передачи в ПЛМ.
                    // Заполненные значения передаём в ПЛМ.
                    var TecchReqGrig = ioVersion.GetAttribute("TecchReqGrig").CollectionElements;
                    String notetxt = GetNoteTTtext(Doc);
                    foreach (var sz in arrLinks)
                    {
                        String szval = "";
                        if (notetxt.Contains(sz.link))
                        {
                            IDimension d1 = (IDimension)Doc.Parameter(sz.link);
                            if (d1 != null)
                                szval = d1.Value.ToString(); // d1.GetUserValueIn(swDoc).ToString();
                        }
                        if (szval != "")
                        {
                            // Ищем ссылку на размер по всем строкам грида ТТ
                            foreach (var element in TecchReqGrig)
                            {
                                InfoObject childio = element.GetValue<InfoObject>("ChildIO");
                                // Поиск в элементе ТТ атрибутов типа "Значения" и сохранение их значения из чертежа
                                var SizeList = childio.GetAttribute("SizeList").CollectionElements;
                                foreach (var isz in SizeList)
                                {
                                    if (sz.link == isz.GetValue<string>("ttElementSolidSizeLink"))
                                    {
                                        isz["ttElementSolidSizeText"] = szval;
                                        childio.Invoke("RecalcVisibleText", null);
                                        element["ttElementValue"] = childio.GetAttribute("VisibleLabel").GetValue();
                                    }
                                }
                            }
                        }
                    }
                }
            }

            Trace.TraceInformation(@"OnCheckIn:  document:" + swDoc + "   ioVersion: " + ioVersion.ToString());
            //Trace.TraceInformation(@"OnCheckIn:  SWDocument:" + addinDoc.ToString());
        }

        /// <summary>
        /// Вызывается в процессе взятия на редактирование
        /// </summary>
        /// <param name="swDoc">документ</param>
        /// <param name="ioVersion">ио версии в плм</param>
        public void OnCheckOut(string swDoc, InfoObject ioVersion)
        {
            //MessageBox.Show("CheckOut");
            //SWDocument addinDoc = AddinDocumentCache.ReturnDocument(swDoc, null) as SWDocument; // SWDocument - обёртка документа  для использования аддином

            PrepareTT(swDoc, ioVersion);

            Trace.TraceInformation(@"OnCheckOut:  document:" + swDoc + "   ioVersion: " + ioVersion.ToString());
        }

        public bool CallBackFunction(string name)
        {
            switch (name)
            {
                case "RefreshTT":
                    RefreshTT();
                    break;
                case "CopyLink":
                    CopyLink();
                    break;                   
                default:
                    return false;
            }
            return true;
        }

        public bool EnableFunction(string name, out int result)
        {
            result = 0;
            switch (name)
            {
                case "RefreshTT":
                    result = 1;
                    break;
                case "CopyLink":
                    result = 2;
                    break;
                default:
                    return false;
            }
            return true;
        }

        public void PrepareTT(string swDoc, InfoObject ioVersion)
        {
            // Чтение атрибутов ИО ПЛМ и запись в кэш
            var TecchReqGrig = ioVersion.GetAttribute("TecchReqGrig").CollectionElements.OrderBy(i => i.GetValue<int>("IndOrder"));
            var count = TecchReqGrig.Count();
            arrTT.Clear();
            arrLinks.Clear();
            notes.Clear();
            foreach (var element in TecchReqGrig)
            {
                //MessageBox.Show(element.GetValue<string>("ttElementLabel"));
                InfoObject childio = element.GetValue<InfoObject>("ChildIO");
                arrTT.Add(childio.GetValue<string>("TextLabel"));
                // Поиск в элементе ТТ атрибутов типа "Значения" и сохранение их в кэш для синхронизации в ПЛМ
                var SizeList = childio.GetAttribute("SizeList").CollectionElements;
                foreach (var sz in SizeList)
                {
                    arrLinks.Add(new ValElem() { link = sz.GetValue<string>("ttElementSolidSizeLink"), val = "" });
                }
            }
            // Получение документа SW
            ModelDoc2 Doc = SWHelper.GetDocByFileName(swDoc); // нужно учитывать что браться на редактирование может документ не открытый в PLM. тогда здесь будет null и нужно либо открывать документ, либо работать через DMDocument
            if (Doc != null)
            {
                // Определяем тип документа. Пока рисуем ТТ только на чертежах
                int docType = Doc.GetType();
                switch (docType)
                {
                    case (int)swDocumentTypes_e.swDocPART:
                        Debug.Print("swPart");
                        // отрисовка ТТ на основе атрибутов
                        //PaintNoteOnPart(Doc, arrTT.ToArray());
                        break;
                    case (int)swDocumentTypes_e.swDocDRAWING:
                        Debug.Print("swDraw");
                        // Расчёт расположения примечаний на основе атрибутов
                        PreCalcNote(Doc, arrTT.ToArray());
                        // отрисовка ТТ 
                        PaintNoteOnDraw(Doc);
                        break;
                    case (int)swDocumentTypes_e.swDocASSEMBLY:
                        Debug.Print("swAssy");
                        break;
                    default: //swDocNONE, swDocSDM
                        Debug.Print("another doc");
                        break;
                }
            }
        }

        public void PaintNoteOnPart(ModelDoc2 swDoc, String[] listt)
        {
            //IPartDoc swDraw = swDoc as IPartDoc;

            bool isfound = false;
            Note swNote = null;
            object[] vAnts = (object[])swDoc.Extension.GetAnnotations();
            for (int i = 0; i < vAnts.Length; i++)
            {
                Annotation myAnnotation = (Annotation)vAnts[i];
                Debug.Print("Value: " + myAnnotation.GetName());
                if (myAnnotation.GetName() == ttname)
                {
                    isfound = true;
                    break;
                }
            }
            if (!isfound)
            {
                swNote = (Note)swDoc.InsertNote("qwerty");
                swNote.SetName(ttname);
                swNote.LockPosition = true;
            }
        }

        public void PreCalcNote(ModelDoc2 swDoc, String[] listt)
        {
            DrawingDoc swDraw = (DrawingDoc)swDoc;
            TextFormat swTextFormat = null;

            // Находим и удаляем старые ТТ
            FindNDelOldNotes(swDoc);
            if (listt.Length == 0) return;

            // Создадим времменный блок. Туда будем вставлять строки и определять высоту блока.
            Note swNoteTemp = (Note)swDoc.InsertNote("");
            Annotation myAnnot = (Annotation)swNoteTemp.GetAnnotation();
            swTextFormat = myAnnot.GetTextFormat(0) as TextFormat;
            swTextFormat.LineLength = c_width - 0.005;
            myAnnot.SetTextFormat(0, false, swTextFormat);
            swNoteTemp.SetTextJustification((int)swTextJustification_e.swTextJustificationLeft);

            // Определим доступное поле для ТТ                         
            Sheet swSheet = (Sheet)swDraw.GetCurrentSheet();
            swSheet.GetSize(ref sh_width, ref sh_height);
            double fheight = sh_height - c_topmargin - c_height - 1 * 0.005;

            notes.Add(new NoteTT());
            NoteTT ntt = notes[0];
            ntt.NoteText = "<PARA indent=5 findent=-5 number=on ntype=1 nformat=$$. nstartNum = 1>";
            ntt.curheight = 0;

            const string notefmt = "<PARA indent=5 findent=-5 number=on ntype=1 nformat=$$.>";
            int curind = 0;
            for (int i = 0; i < listt.Length; i++)
            {
                ntt = notes[curind];
                string onestring = notefmt + listt[i];
                // Вставляем одну ТТ во временный блок
                swNoteTemp.SetText(onestring);
                double tempnoteheight = GetTempNoteHeight(swNoteTemp);
                // Проверяем получающуюся высоту примечания
                if (ntt.curheight + tempnoteheight <= fheight)
                {
                    // Если помещается, дописываем текст и увеличиваем высоту проверки
                    ntt.NoteText += listt[i] + System.Environment.NewLine;
                    ntt.curheight += tempnoteheight;
                }
                else
                {
                    // Иначе создаём новое примечание
                    notes.Add(new NoteTT());
                    curind++;
                    ntt = notes[curind];
                    ntt.NoteText = "<PARA indent=5 findent=-5 number=on ntype=1 nformat=$$. nstartNum = " + (i + 1).ToString() + ">";
                    ntt.curheight = 0;
                }
            }

            swDoc.ClearSelection2(true);
            myAnnot.Select(false);
            swDoc.DeleteSelection(false);
            swDoc.WindowRedraw();
        }

        private void FindNDelOldNotes(ModelDoc2 swDoc)
        {
            Note swNote = null;
            DrawingDoc swDraw = (DrawingDoc)swDoc;
            object[] vNotes;

            IView swView = swDraw.GetFirstView() as IView;
            vNotes = (object[])swView.GetNotes();
            for (int i = 0; i < vNotes.Length; i++)
            {
                swNote = (Note)vNotes[i];
                String snm = swNote.GetName();
                if (snm.Contains(ttname))
                {
                    Annotation myAnnotation = (Annotation)swNote.GetAnnotation();
                    swDoc.ClearSelection2(true);
                    myAnnotation.Select(false);
                    swDoc.DeleteSelection(false);
                }
            }
            swDoc.WindowRedraw();
        }

        public void PaintNoteOnDraw(ModelDoc2 swDoc)
        {
            bool isA4 = (sh_width == 0.210 && sh_height == 0.297);
            double prevy = 0;
            for (int i = 0; i < notes.Count; i++)
            {
                Note swNote = (Note)swDoc.InsertNote("");
                swNote.SetName(ttname + (i + 1).ToString());
                swNote.SetTextJustification((int)swTextJustification_e.swTextJustificationLeft);

                Annotation myAnnot = (Annotation)swNote.GetAnnotation();
                TextFormat swTextFormat = myAnnot.GetTextFormat(0) as TextFormat;
                swTextFormat.LineLength = c_width - 0.005;
                myAnnot.SetTextFormat(0, false, swTextFormat);

                String txt = notes[i].NoteText.Remove(notes[i].NoteText.Length - 2);
                swNote.SetText(txt);
                // Пересчёт позиции
                PositionTT posTT = CalcPositionTT(swNote);
                // Для А4 ТТ прижимаются к основной надписи, для других форматов - к верху чертежа
                if (isA4 == false) posTT.y = sh_height - c_topmargin;
                // Последующие блоки выравниваем по высоте с первым блоком
                if (i == 0) prevy = posTT.y; else posTT.y = prevy;
                myAnnot.SetPosition(posTT.x - c_width * i, posTT.y, posTT.z);
                swNote.LockPosition = true;
            }
            swDoc.WindowRedraw();
        }

        private PositionTT CalcPositionTT(Note myNote)
        {
            double[] ext = (double[])myNote.GetExtent();
            double nheight = Math.Abs(Math.Abs(ext[4]) - Math.Abs(ext[1]));
            PositionTT result = new PositionTT() { x = sh_width - c_width, y = c_height + nheight + 0.005, z = 0 };
            return result;
        }

        public String GetNoteTTtext(ModelDoc2 swDoc)
        {
            DrawingDoc swDraw = (DrawingDoc)swDoc;
            object[] vNotes;
            String result = "";

            IView swView = swDraw.GetFirstView() as IView;
            vNotes = (object[])swView.GetNotes();
            Note swNote = null;
            for (int j = 0; j < vNotes.Length; j++)
            {
                swNote = (Note)vNotes[j];
                String snm = swNote.GetName();
                if (snm.Contains(ttname))
                {
                    result += swNote.PropertyLinkedText + System.Environment.NewLine;
                }
            }
            return result;
        }

        private double GetTempNoteHeight(Note myNote)
        {
            double[] ext = (double[])myNote.GetExtent();
            double nheight = Math.Abs(Math.Abs(ext[4]) - Math.Abs(ext[1]));
            return nheight;
        }

        public void ReadMat(ModelDoc2 swDoc)
        {
            AssemblyDoc swADoc;
            object[] varComp;
            double[] varMatProp;


            swADoc = (AssemblyDoc)swDoc;

            varComp = (object[])swADoc.GetComponents(true);
            int I = 0;
            for (I = 0; I < varComp.Length; I++)
            {
                Component2 swComp = default(Component2);
                swComp = (Component2)varComp[I];

                varMatProp = (double[])swComp.GetModelMaterialPropertyValues(swComp.ReferencedConfiguration);
                if (!((varMatProp == null)))
                {
                    Debug.Print(swComp.Name2 + "(" + I + ")" + "ConfigName : " + swComp.ReferencedConfiguration + "MatProp : ");
                    Debug.Print("Red: " + (varMatProp[0]) * 255.0);
                    Debug.Print("Green: " + (varMatProp[1]) * 255.0);
                    Debug.Print("Blue: " + (varMatProp[2]) * 255.0);
                    Debug.Print("Ambient: " + (varMatProp[3]) * 100.0 + "%");
                    Debug.Print("Diffuse: " + (varMatProp[4]) * 100.0 + "%");
                    Debug.Print("Specularity: " + (varMatProp[5]) * 100.0 + "%");
                    Debug.Print("Shininess: " + (varMatProp[6]) * 100.0 + "%");
                    Debug.Print("Transparency: " + (varMatProp[7]) * 100.0 + "%");
                    Debug.Print("Emission: " + (varMatProp[8]) * 100.0 + "%");
                }

                Debug.Print("");
            }
        }

        #endregion

        #region Implementation of ICustomModuleFunction

        /// <summary>
        /// обработчик комманд из скриптовой оболочки
        /// из скриптов конфиурации(например обработчик команды UI) можно вызвать так:
        /// if ( Service.ModuleEnvironment.Contains("Solidworks"))
        ///{
        ///      Service.InvokeModuleMethod( "pmsz.plmsoyuz.addin2solidworks.ext", "test_method", obj );
        ///}
        /// 
        /// </summary>
        /// <param name="methodName"></param>
        /// <param name="inputParams"></param>
        /// <param name="outputParams"></param>
        /// <returns></returns>
        public bool Invoke(string methodName, object inputParams, out object outputParams)
        {
            outputParams = null;
            switch (methodName)
            {
                case @"test_method":
                    Trace.TraceInformation(@"Invoke:  document:" + inputParams.ToString());
                    return true;
                default:
                    return false;
            }
        }

        #endregion
    }
}
