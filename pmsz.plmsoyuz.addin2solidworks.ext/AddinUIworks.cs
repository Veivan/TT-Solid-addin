using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Reflection;

using SolidWorks.Interop.swconst;
using SolidWorks.Interop.sldworks;
using Microsoft.Win32;
using System.Windows.Forms;

namespace pmsz.plmsoyuz.addin2solidworks.ext.UI
{
    public class AddinUIworks
    {
        static int doctype = (int)swDocumentTypes_e.swDocDRAWING;
        static int ccomcount = 2; // Число комманд, которые будут созданы
        static string Title = @"Союз-PLM";
        static int UserID = 100500;
        static int[] cmdIDs = new int[ccomcount];
        static int[] TextType = new int[ccomcount];

        /// <summary>
        /// Настройка команд ТТ в интерфейсе SolidWorks
        /// </summary>
        public static void RefreshUI()
        {
            CommandTab cmdTab = SwAddin.Instance.UiManager.CmdMgr.GetCommandTab(doctype, Title);
            if (cmdTab == null) return;
            PrepareCmd();
            if (NeedRecreate())
            {
                DeleteOldTab(cmdTab);
                CreateNewTab(cmdTab);
            }
            else RestoreTab(cmdTab);
        }

        /// <summary>
        /// Удаление созданной группы команд (сами команды останутся в памяти SolidWorks)
        /// </summary>
        public static void ClearCommandMgr()
        {
            SwAddin.Instance.UiManager.CmdMgr.RemoveCommandGroup(UserID);
        }

        /// <summary>
        /// Прдготовка команд для их включения в таббокс
        /// </summary>
        private static void PrepareCmd()
        {
            var cmdGroup = SwAddin.Instance.UiManager.CmdMgr.CreateCommandGroup(UserID, "Техтребования", "", "", -1);
            var strResourcePath = GetResourcePath();
            cmdGroup.LargeIconList = strResourcePath + "ToolBarTT.bmp";
            cmdGroup.SmallIconList = strResourcePath + "ToolBarTT.bmp";
            cmdGroup.LargeMainIcon = strResourcePath + "Logo24.bmp";
            cmdGroup.SmallMainIcon = strResourcePath + "Logo16.bmp";
            int cmdIndex0 = cmdGroup.AddCommandItem2("Обновить ТТ", -1, "Обновить ТТ", "Обновить ТТ", 0, "CallBackFunction(RefreshTT)", "", 1, 3);
            int cmdIndex1 = cmdGroup.AddCommandItem2("Скопировать ссылку", -1, "Скопировать ссылку", "Скопировать ссылку", 1, "CallBackFunction(CopyLink)", "", 1, 3);

            cmdGroup.HasMenu = true;
            cmdGroup.HasToolbar = true;
            cmdGroup.Activate();

            cmdIDs[0] = cmdGroup.get_CommandID(cmdIndex0);
            TextType[0] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;
            cmdIDs[1] = cmdGroup.get_CommandID(cmdIndex1);
            TextType[1] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;
        }

        /// <summary>
        /// Проверка необходимости пересоздания таббокса и команд ТТ
        /// </summary>
        private static bool NeedRecreate()
        {
            bool result = false;
            string key = "addinttversion";
            int innervers = 2; // Версия сборки. Изменить, если нужно пересоздать таббокс.
            try
            {
                RegistryKey rk = Registry.CurrentUser.OpenSubKey("Software\\ProgramSoyuz", true);
                if (rk == null)
                    rk = Registry.CurrentUser.CreateSubKey("Software\\ProgramSoyuz");
                object regvers = rk.GetValue(key);
                if (regvers == null || (int)regvers != innervers)
                {
                    result = true;
                    rk.SetValue(key, innervers);
                }
            }
            catch
            {
                result = true;
                MessageBox.Show("Ошибка работы с реестром.");
            }
            return result;
        }

        /// <summary>
        /// Удаление существующих команд и их таббокса
        /// </summary>
        /// <param name="cmdTab">Лента Союз-PLM</param>
        private static void DeleteOldTab(CommandTab cmdTab)
        {
            int trcount = 0;
            // Найдём все CommandTabBox
            int tblen = cmdTab.GetCommandTabBoxCount();
            object[] cmdTabBoxes = (object[])cmdTab.CommandTabBoxes();
            // Перебираем TabBoxes и ищем - есть такой с нужными командами?
            for (int i = tblen - 1; i >= 0; i--)
            {
                object idObject = null;
                object textTypeObject = null;
                // Выбираем команды из таббокса
                int cmdcount = (cmdTabBoxes[i] as CommandTabBox).GetCommands(out idObject, out textTypeObject);
                int[] ids = (int[])idObject;
                trcount = 0;
                // Перебираем команды ТТ и удаляем их из таббокса
                for (int j = 0; j < cmdIDs.Length; j++)
                {
                    for (int k = 0; k < cmdcount; k++)
                    {
                        if (cmdIDs[j] == ids[k])
                        {
                            trcount++;
                            (cmdTabBoxes[i] as CommandTabBox).RemoveCommands(cmdIDs);
                        }
                    }
                }
                // Если у таббокса ccomcount команд, и есть команды ТТ, то это наш таббокс. удалим его.
                if (cmdcount == ccomcount && trcount > 0)
                {
                    cmdTab.RemoveCommandTabBox((CommandTabBox)cmdTabBoxes[i]);
                    break;
                }
            }
        }

        /// <summary>
        /// /Создание нового таббокса
        /// </summary>
        /// <param name="cmdTab">Лента Союз-PLM</param>
        private static void CreateNewTab(CommandTab cmdTab)
        {
            CommandTabBox cmdBox = cmdTab.AddCommandTabBox();
            bool bResult = cmdBox.AddCommands(cmdIDs, TextType);
            Debug.Print(bResult.ToString());
            cmdTab.Active = true;
        }

        /// <summary>
        /// Проверка наличия команд ТТ и их таббокса. При отсутствии - восстановление.
        /// </summary>
        /// <param name="cmdTab">Лента Союз-PLM</param>
        private static void RestoreTab(CommandTab cmdTab)
        {
            bool tabexists = false;
            int trcount = 0;
            // Найдём все CommandTabBox
            int tblen = cmdTab.GetCommandTabBoxCount();
            object[] cmdTabBoxes = (object[])cmdTab.CommandTabBoxes();
            // Перебираем TabBoxes и ищем - есть такой с командами ТТ?
            for (int i = tblen - 1; i >= 0; i--)
            {
                object idObject = null;
                object textTypeObject = null;
                // Выбираем команды из таббокса
                int cmdcount = (cmdTabBoxes[i] as CommandTabBox).GetCommands(out idObject, out textTypeObject);
                int[] ids = (int[])idObject;
                trcount = 0;
                // Ищем команды ТТ
                for (int j = 0; j < cmdIDs.Length; j++)
                    for (int k = 0; k < cmdcount; k++)
                        if (cmdIDs[j] == ids[k])
                            trcount++;
                // Если у таббокса команд не больше ccomcount, и есть команды ТТ, то это наш таббокс.
                if (cmdcount <= ccomcount && trcount > 0)
                {
                    // Проверим число команд. Если не хватает, то пересоздадим
                    if (trcount != ccomcount)
                    {
                        (cmdTabBoxes[i] as CommandTabBox).RemoveCommands(cmdIDs);
                        (cmdTabBoxes[i] as CommandTabBox).AddCommands(cmdIDs, TextType);
                    }
                    tabexists = true;
                    break;
                }
            }
            // Если таббокс ТТ не нашёлся - пересоздадим его
            if (!tabexists) CreateNewTab(cmdTab);
        }

        private static string GetResourcePath()
        {
            Module thisModule;
            String strModulePath;

            thisModule = Assembly.GetExecutingAssembly().GetModules()[0];
            strModulePath = thisModule.FullyQualifiedName;

            String strResourcePath = "";

            try
            {
                strResourcePath = strModulePath.Substring(0, strModulePath.LastIndexOf("\\"));
            }
            catch
            {
                //Msg("Error with Module Path")
            }

            strResourcePath += @"\Resources\";
            return strResourcePath;
        }
    }
}
