using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using System.Data.Sql;
using System.IO;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using System.Windows;

namespace VirtualClipBoard
{
    public partial class VirtualClipBoard : Form
    {
        String VirtualClipBoard_Name = "VirtualClipBoard"; // название программы
        public String VirtualClipBoard_TARGET; // последний значение текстового БО
        public String VirtualClipBoard_DAT; // путь к файлу истории
        IDictionary<int, string> VirtualClipBoard_History = new Dictionary<int, string>(); // История нашего буфера
        IDictionary<int, int> VirtualClipBoard_Index_ListBox; // список индексов в связки с ключами истории буфера

        // Подключение библиотек WIN
        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr SetClipboardViewer(IntPtr hWndNewViewer);
        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        public static extern bool ChangeClipboardChain(IntPtr hWndRemove, IntPtr hWndNewNext);
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern int SendMessage(IntPtr hwnd, int wMsg, IntPtr wParam, IntPtr lParam);

        // Наш Form
        public VirtualClipBoard()
        {
            InitializeComponent();
            load_configs();

            nextClipboardViewer = (IntPtr)SetClipboardViewer((IntPtr)this.Handle);

            reload_tray(); // Обноавляемменю в трее
            reload_list_clipboard(); // Обновляем ListBox

            _notifyIcon.Text = VirtualClipBoard_Name;
            _notifyIcon.MouseDoubleClick += new MouseEventHandler(_notifyIcon_MouseDoubleClick);
        }

        // Перезагрузка элементов в ListBox
        private void reload_list_clipboard()
        {
            VirtualClipBoard_Index_ListBox = new Dictionary<int, int>();
            int list_target_item = 0; // индекс текущего элемента в ListBox
            list_clipboard.Items.Clear(); // Очищаем список
            String string_name_ite;
            int free_slot_to_tray = Properties.Settings.Default.history_size;
            var list = VirtualClipBoard_History.OrderByDescending(x => x.Key);
            foreach (var item in list)
            {
                if (item.Value.Length > 150)
                {
                    string_name_ite = item.Value.Replace("\n", "\t").Replace("\r", "\t").Substring(0, 60);
                }
                else
                {
                    string_name_ite = item.Value.Replace("\n", "\t").Replace("\r", "\t");
                }
                list_clipboard.Items.Add(string_name_ite);
                VirtualClipBoard_Index_ListBox.Add(list_target_item, item.Key);
                if (free_slot_to_tray == 1) { break; } else { free_slot_to_tray--; }
                list_target_item++; // Увеличиваем индекс текущего элемента в ListBox
            }
        }

        // Выбор элемента в ListBox
        private void list_clipboard_SelectedIndexChanged(object sender, EventArgs e)
        {
            Clipboard.SetText(VirtualClipBoard_History[VirtualClipBoard_Index_ListBox[list_clipboard.SelectedIndex]]);
        }

        // Перезагрузка элементов для трей
        private void reload_tray()
        {
            ContextMenuStrip contextMenu = new ContextMenuStrip();
            ToolStripMenuItem menuItem;

            int free_slot_to_tray = Properties.Settings.Default.size_tray;
            var list = VirtualClipBoard_History.OrderByDescending(x => x.Key);
            foreach (var item in list)
            {
                menuItem = new ToolStripMenuItem();
                menuItem.Tag = item.Key;
                if (item.Value.Length > 60)
                {
                    menuItem.Text = item.Value.Replace("\n", "\t").Replace("\r", "\t").Substring(0, 60);
                } else {
                    menuItem.Text = item.Value.Replace("\n", "\t").Replace("\r", "\t");
                }
                
                menuItem.Click += new System.EventHandler(menu_item_click);
                contextMenu.Items.Add(menuItem);
                if (free_slot_to_tray == 1) { break; } else { free_slot_to_tray--; }
            }

            // Разделитель
            contextMenu.Items.Add(new ToolStripSeparator());

            // Свернуть/Развернуть
            menuItem = new ToolStripMenuItem();
            menuItem.Text = "Настройки";
            menuItem.Click += new System.EventHandler(menu_item_config);
            contextMenu.Items.Add(menuItem);

            // Выход из программы
            menuItem = new ToolStripMenuItem();
            menuItem.Text = "Выход";
            menuItem.Click += new System.EventHandler(exit_Click);
            contextMenu.Items.Add(menuItem);

            _notifyIcon.ContextMenuStrip = contextMenu;
        }

        // Вызов окна настроек
        private void menu_item_config(object sender, EventArgs e)
        {
            // ShowInTaskbar = true;
            Show();
            WindowState = FormWindowState.Normal;
        }

        // Событие по клику на элемент контекстного меню в трее
        private void menu_item_click(object sender, EventArgs e)
        {
            // Console.WriteLine((int)(sender as ToolStripMenuItem).Tag);
            Clipboard.SetText(VirtualClipBoard_History[(int)(sender as ToolStripMenuItem).Tag]);
        }

        // событие при клике мышкой по значку в трее
        private void _notifyIcon_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            Console.WriteLine(WindowState);
            if (WindowState == FormWindowState.Normal || WindowState == FormWindowState.Maximized)
            {
                // ShowInTaskbar = false;
                Hide();
                WindowState = FormWindowState.Minimized;
            }
            else
            {
                // ShowInTaskbar = true;
                Show();
                WindowState = FormWindowState.Normal;
            }
        }

        // Установка путей к файлам конфигурации и истории
        private void load_configs()
        {
            VirtualClipBoard_DAT = Application.UserAppDataPath + "\\history.dat";
            Console.WriteLine("Файл истории: " + VirtualClipBoard_DAT);
            history_size.Value = Properties.Settings.Default.history_size;
            Console.WriteLine("Размер истории загружен из настроек: " + Properties.Settings.Default.history_size);
            size_tray.Value = Properties.Settings.Default.size_tray;
            Console.WriteLine("Количество элементов в трее загружено из настроек: " + Properties.Settings.Default.size_tray);
            RegistryKey reg = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Run\", true);
            if (reg.GetValue(VirtualClipBoard_Name) != null){
                autoload.Checked = true;
                Console.WriteLine("Приложение записано в автозагрузку. (В настройках ставим Checked = true)");
            }
            reg.Close();

            // Загружаем историю из файла
            String XMLString = "";
            XMLString += @"<items>";
            if (File.Exists(VirtualClipBoard_DAT))
            {
                StreamReader stream = new StreamReader(VirtualClipBoard_DAT);
                while (stream.Peek() > -1)
                {
                    XMLString += stream.ReadLine() + "\n";
                }
                stream.Close();
                XMLString += @"</items>";
                int index_new_history = 2;
                XDocument doc = XDocument.Parse(XMLString);
                var items = doc.Element("items").Elements("item");
                foreach (XElement item in items)
                {
                    VirtualClipBoard_History.Add(index_new_history, item.Value);
                    index_new_history++; // увеличиваем индекс новому элементу
                }
            }
            // Чистим историю буфера
            if (VirtualClipBoard_History.Count() > Properties.Settings.Default.history_size)
            {
                int clear_items_count = VirtualClipBoard_History.Count() - Properties.Settings.Default.history_size;
                var list = VirtualClipBoard_History.Keys.ToList();
                list.Sort();
                foreach (var key in list)
                {
                    VirtualClipBoard_History.Remove(key);
                    if (clear_items_count == 1) { break; } else { clear_items_count--; }
                }
            }
            // Обновляем файл истории
            StreamWriter writer = new StreamWriter(VirtualClipBoard_DAT, false, System.Text.Encoding.UTF8);
            var new_list = VirtualClipBoard_History.Keys.ToList();
            new_list.Sort();
            foreach (var key in new_list)
            {
                writer.WriteLine(@"<item>" + VirtualClipBoard_History[key].Replace(@"<", @"&lt;").Replace(@">", @"&gt;") + @"</item>");
            }
            writer.Close();
            // Если элементов ноль, добавляем из буфера
            Console.WriteLine(VirtualClipBoard_History.Count());
            if (VirtualClipBoard_History.Count() == 0)
            {
                VirtualClipBoard_TARGET = Clipboard.GetText();
                VirtualClipBoard_History.Add(1, VirtualClipBoard_TARGET);
            }
            VirtualClipBoard_TARGET = VirtualClipBoard_History.Last().Value;
        }

        // Событие изменения статуса флажка автозагрузки
        // Если флажок - прописываем в реестр на автозагрузку
        private void autoload_CheckedChanged(object sender, EventArgs e)
        {
            RegistryKey reg = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Run\", true);
            if (reg.GetValue(VirtualClipBoard_Name) != null)
            {
                try
                {
                    reg.DeleteValue(VirtualClipBoard_Name);
                    Console.WriteLine("Программа " + VirtualClipBoard_Name + " удалена из автозагрузки в реестре");
                }
                catch
                {
                    Console.WriteLine("Ошибка удаления " + VirtualClipBoard_Name + " из автозагрузки в реестре");
                }
            }
            if(autoload.Checked)
            {
                reg.SetValue(VirtualClipBoard_Name, Application.ExecutablePath);
                Console.WriteLine("Программа " + VirtualClipBoard_Name + " записана в автозагрузку через реестр");
            }
            reg.Close();
        }


        // Завершение работы программы по закрытию через кнопку
        private void exit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        // изменение размера истории
        private void history_size_ValueChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.history_size = (int)history_size.Value;
            Properties.Settings.Default.Save();
            Console.WriteLine("Размер истории изменен: " + Properties.Settings.Default.history_size);
            reload_list_clipboard(); // Обновляем ListBox
        }

        // изменение количества записей БО в трее
        private void size_tray_ValueChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.size_tray = (int)size_tray.Value;
            Properties.Settings.Default.Save();
            Console.WriteLine("Количество элементов в трее изменено: " + Properties.Settings.Default.size_tray);
            reload_tray(); // Обновляем Трей
        }

        // Реагируем на обновление буфераобмена
        private void ClipboardChanged()
        {
            if (Clipboard.ContainsText() && Clipboard.GetText().Length > 0 && VirtualClipBoard_TARGET != Clipboard.GetText())
            {
                VirtualClipBoard_TARGET = Clipboard.GetText();

                // Записываем новый элемент в словарь
                VirtualClipBoard_History.Add((VirtualClipBoard_History.Last().Key + 1), VirtualClipBoard_TARGET);

                reload_tray(); // Обноавляемменю в трее
                reload_list_clipboard(); // Обновляем ListBox

                // Отчистка словаря от лишних элементов
                if (VirtualClipBoard_History.Count() > Properties.Settings.Default.history_size)
                {
                    int clear_items_count = VirtualClipBoard_History.Count() - Properties.Settings.Default.history_size;
                    var list = VirtualClipBoard_History.Keys.ToList();
                    list.Sort();
                    foreach (var key in list)
                    {
                        VirtualClipBoard_History.Remove(key);
                        if (clear_items_count == 1) { break; } else { clear_items_count--; }
                    }
                }

                // Записываем новый элемент в файл истории
                StreamWriter writer = new StreamWriter(VirtualClipBoard_DAT, true, System.Text.Encoding.UTF8);
                writer.WriteLine(@"<item>" + VirtualClipBoard_TARGET.Replace(@"<", @"&lt;").Replace(@">", @"&gt;") + @"</item>");
                writer.Close();
                Console.WriteLine("В историю добавлен новый элемент: " + VirtualClipBoard_TARGET);
            }
        }

        // Затираем всю историю
        private void clear_Click(object sender, EventArgs e)
        {
            StreamWriter writer = new StreamWriter(VirtualClipBoard_DAT, false, System.Text.Encoding.Default);
            writer.Write("");
            writer.Close();

            VirtualClipBoard_History = new Dictionary<int, string>();

            VirtualClipBoard_TARGET = Clipboard.GetText();
            VirtualClipBoard_History.Add(1, VirtualClipBoard_TARGET);

            reload_tray(); // Обноавляемменю в трее
            reload_list_clipboard(); // Обновляем ListBox
        }

        // Сворачивать в трей вместо закрытия программы
        protected override void OnClosing(CancelEventArgs e)
        {
            e.Cancel = true;
            //ShowInTaskbar = false;
            Hide();
            WindowState = FormWindowState.Minimized;
        }

        // дескриптор окна
        private IntPtr nextClipboardViewer;

        // Константы
        public const int WM_DRAWCLIPBOARD = 0x308;
        public const int WM_CHANGECBCHAIN = 0x030D;

        // Метод для реагирование на изменение вбуфере обмена и т.д.
        protected override void WndProc(ref Message m)
        {
            // Console.WriteLine("WndProc");
            switch (m.Msg)
            {
                case WM_DRAWCLIPBOARD:
                    {
                        ClipboardChanged();
                        //Console.WriteLine("WM_DRAWCLIPBOARD ClipboardChanged();");
                        SendMessage(nextClipboardViewer, WM_DRAWCLIPBOARD, m.WParam, m.LParam);
                        break;
                    }
                case WM_CHANGECBCHAIN:
                    {
                        if (m.WParam == nextClipboardViewer)
                        {
                            nextClipboardViewer = m.LParam;
                        }
                        else
                        {
                            SendMessage(nextClipboardViewer, WM_CHANGECBCHAIN, m.WParam, m.LParam);
                        }
                        m.Result = IntPtr.Zero;
                        break;
                    }
                default:
                    {
                        base.WndProc(ref m);
                        break;
                    }
            }
        }
    }
}
