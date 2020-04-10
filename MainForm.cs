// инициализируем юниты
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
// добавленные мною
using System.Management; //выполняет запросы WMI
using System.DirectoryServices;
using System.Net.NetworkInformation; // IP ping юнит
using System.IO;
using System.Net;
using System.Threading.Tasks;
using System.Collections.Generic;
// берем date-time штамп
using System.Globalization;
using System.Threading;
// для работы с Excel
using OfficeOpenXml;

namespace VarshotsInventory
{
    public partial class MainForm : Form
    {
        // инициализируем
        public MainForm()
        {
            InitializeComponent();
        }

        // WMI запрос к локальному компьютеру
        private void btn_LocWMIQueryRun_Click(object sender, EventArgs e)
        {
            
        }

        private void btn_NetWMIQueryRun_Click(object sender, EventArgs e) //WMI запрос к сетевому компьютеру
        {
            ///////////////////////////////////////
            // активизируем таймер пинга IP машины
            var period = 10 * 1000;
            string ipAddressARM = tB_ComputerName.Text;
            System.Threading.Timer t2 = new System.Threading.Timer(CallbackARM, ipAddressARM, 0, period);

            

            // сетевой опрос АРМ - ОС
            // WMI запрос к локальному компьютеру - Пользователь
            // инициализируем выбранный запрос через команду
            SelectQuery query_os = new SelectQuery(@"Select * from Win32_OperatingSystem");

            //инициализируем поисковый бот для запроса выбранного для запуска
            using (ManagementObjectSearcher searcher_os = new ManagementObjectSearcher(query_os))
            {
                //исполняемый запрос
                foreach (ManagementObject process_os in searcher_os.Get())
                {
                    //инфо о Пользователе
                    textBox4.AppendText(" || OS info: " + process_os["Name"].ToString()); // что ищем
                }
            }

            // сетевой опрос АРМ - Имя АРМ
            // WMI запрос к локальному компьютеру - Пользователь
            // инициализируем выбранный запрос через команду
            SelectQuery query_name = new SelectQuery(@"Select * from Win32_ComputerSystem");

            //инициализируем поисковый бот для запроса выбранного для запуска
            using (ManagementObjectSearcher searcher_name = new ManagementObjectSearcher(query_name))
            {
                //исполняемый запрос
                foreach (ManagementObject process_name in searcher_name.Get())
                {
                    //инфо о Пользователе
                    textBox5.AppendText(" || PC info: " + process_name["Name"].ToString()); // что ищем
                }
            }


            // сетевой опрос АРМ - Пользователь
            // WMI запрос к локальному компьютеру - Пользователь
            // инициализируем выбранный запрос через команду
            SelectQuery query_user = new SelectQuery(@"Select * from Win32_ComputerSystem");

            //инициализируем поисковый бот для запроса выбранного для запуска
            using (ManagementObjectSearcher searcher_user = new ManagementObjectSearcher(query_user))
            {
                //исполняемый запрос
                foreach (ManagementObject process_user in searcher_user.Get())
                {
                    //инфо о Пользователе
                    textBox6.AppendText(" || User info: " + process_user["Username"].ToString()); // что ищем
                }
            }


            // сетевой опрос АРМ - Общее инфо
            // сетевой опрос АРМ - Имя АРМ + локальный опрос АРМ - Пользователь
            textBox3.AppendText(textBox5.Text + " " + textBox6.Text); // что ищем


            // сетевой опрос АРМ - МАС адрес
            //////////////////////////////////////////////////////
            // сетевой опрос АРМ - МАС адрес
            // WMI запрос к сетевому компьютеру - МАС адрес
            //////////////////////////////////////////////////////
            ManagementScope scope_mac = new ManagementScope("\\\\" + tB_ComputerName.Text + "\\root\\cimv2"); //область поиска. textBox7.Text - ТекстБокс с которого я возьму имя компьютера
            SelectQuery query_mac = new SelectQuery(@"Select * from Win32_NetworkAdapter where Name='" + comboBox1.SelectedItem.ToString() + "'"); // Запрос
            //инициализируем поисковый бот для запроса выбранного для запуска
            using (ManagementObjectSearcher searcher_mac = new ManagementObjectSearcher(scope_mac, query_mac))
            {
                //исполняемый запрос
                foreach (ManagementObject process_mac in searcher_mac.Get())
                {
                    //инфо о МАС адресе
                    //вывожу наименование мас адреса в метку на форме
                    textBox8.Text = (string)process_mac["MACAddress"]; // что ищем
                }
            }

            // сетевой опрос АРМ - IP
            //////////////////////////////////////////////////////
            // сетевой опрос АРМ - IP
            // WMI запрос к сетевому компьютеру - IP
            //////////////////////////////////////////////////////
            ManagementScope scope_ip = new ManagementScope("\\\\" + tB_ComputerName.Text + "\\root\\cimv2"); //область поиска. textBox7.Text - ТекстБокс с которого я возьму имя компьютера
            SelectQuery query_ip = new SelectQuery(@"Select * from Win32_NetworkAdapterConfiguration WHERE IPEnabled = 'TRUE'"); // Запрос
            //инициализируем поисковый бот для запроса выбранного для запуска
            using (ManagementObjectSearcher searcher_ip = new ManagementObjectSearcher(scope_ip, query_ip))
            {
                //исполняемый запрос
                foreach (ManagementObject process_ip in searcher_ip.Get())
                {
                    //инфо о IP
                    //вывожу значение IP в метку на форме
                    string[] arrIPAddress = (string[])(process_ip["IPAddress"]); // что ищем

                    string sIPAddress = arrIPAddress[0];

                    textBox9.AppendText(" || IP: " + sIPAddress.ToString()); // WMI запрос к локальному компьютеру - Материнская плата модель
                }
            }


            // сетевой опрос АРМ - Материнская плата
            //////////////////////////////////////////////////////
            // сетевой опрос АРМ - Материнская плата
            // WMI запрос к сетевому компьютеру - Материнская плата
            //////////////////////////////////////////////////////
            ManagementScope scope_motherboard = new ManagementScope("\\\\" + tB_ComputerName.Text + "\\root\\cimv2"); //область поиска. textBox7.Text - ТекстБокс с которого я возьму имя компьютера
            SelectQuery query_motherboard = new SelectQuery(@"Select * from Win32_BaseBoard"); // Запрос
            //инициализируем поисковый бот для запроса выбранного для запуска
            using (ManagementObjectSearcher searcher_motherboard = new ManagementObjectSearcher(scope_motherboard, query_motherboard))
            {
                //исполняемый запрос
                foreach (ManagementObject process_motherboard in searcher_motherboard.Get())
                {
                    //инфо о motherboard (материнской плате)
                    //вывожу название матплаты в метку на форме
                    textBox10.AppendText(" || Model: " + process_motherboard["Manufacturer"].ToString()); // WMI запрос к локальному компьютеру - Материнская плата модель
                    textBox10.AppendText(" || Product: " + process_motherboard["Product"].ToString()); // WMI запрос к локальному компьютеру - Материнская плата продукт ID
                    textBox10.AppendText(" || SN.: " + process_motherboard["SerialNumber"].ToString()); // WMI запрос к локальному компьютеру - Материнская плата серийный номер
                }
            }

            // сетевой опрос АРМ - BIOS
            //////////////////////////////////////////////////////
            // сетевой опрос АРМ - BIOS
            // WMI запрос к сетевому компьютеру - BIOS
            //////////////////////////////////////////////////////
            ManagementScope scope_bios = new ManagementScope("\\\\" + tB_ComputerName.Text + "\\root\\cimv2"); //область поиска. textBox7.Text - ТекстБокс с которого я возьму имя компьютера
            SelectQuery query_bios = new SelectQuery(@"Select * from Win32_Processor"); // Запрос
            //инициализируем поисковый бот для запроса выбранного для запуска
            using (ManagementObjectSearcher searcher_bios = new ManagementObjectSearcher(scope_bios, query_bios))
            {
                //исполняемый запрос
                foreach (ManagementObject process_bios in searcher_bios.Get())
                {
                    //инфо о Процессорах (CPU)
                    textBox11.AppendText(" || BIOS Info name & ver.: " + process_bios["Name"].ToString()); // что ищем
                }
            }

            // сетевой опрос АРМ - CPU
            //////////////////////////////////////////////////////
            // сетевой опрос АРМ - CPU
            // WMI запрос к сетевому компьютеру - CPU 
            //////////////////////////////////////////////////////
            ManagementScope scope_cpu = new ManagementScope("\\\\" + tB_ComputerName.Text + "\\root\\cimv2"); //область поиска. textBox7.Text - ТекстБокс с которого я возьму имя компьютера
            SelectQuery query_cpu = new SelectQuery(@"Select * from Win32_Processor"); // Запрос
            //инициализируем поисковый бот для запроса выбранного для запуска
            using (ManagementObjectSearcher searcher_cpu = new ManagementObjectSearcher(scope_cpu, query_cpu))
            {
                //исполняемый запрос
                foreach (ManagementObject process_cpu in searcher_cpu.Get())
                {
                    //инфо о Процессорах (CPU)
                    textBox12.AppendText(" || Info: " + process_cpu["Name"].ToString()); // что ищем
                }
            }


            // сетевой опрос АРМ - Video-adapter
            //////////////////////////////////////////////////////
            // сетевой опрос АРМ - Video-adapter
            // WMI запрос к сетевому компьютеру - Video-adapter
            //////////////////////////////////////////////////////
            ManagementScope scope_video = new ManagementScope("\\\\" + tB_ComputerName.Text + "\\root\\cimv2"); //область поиска. textBox7.Text - ТекстБокс с которого я возьму имя компьютера
            SelectQuery query_video = new SelectQuery(@"Select * from Win32_VideoController"); // Запрос
            //инициализируем поисковый бот для запроса выбранного для запуска
            using (ManagementObjectSearcher searcher_video = new ManagementObjectSearcher(scope_video, query_video))
            {
                //исполняемый запрос
                foreach (ManagementObject process_video in searcher_video.Get())
                {
                    //инфо о Видео-адаптервх
                    textBox13.AppendText(" || Info: " + process_video["Description"].ToString()); // что ищем
                }
            }

            // сетевой опрос АРМ - Monitor
            ManagementScope scope_monitor = new ManagementScope("\\\\" + tB_ComputerName.Text + "\\root\\cimv2"); //область поиска. textBox7.Text - ТекстБокс с которого я возьму имя компьютера
            SelectQuery query_monitor = new SelectQuery(@"Select * from Win32_DesktopMonitor"); // Запрос
            //инициализируем поисковый бот для запроса выбранного для запуска
            using (ManagementObjectSearcher searcher_monitor = new ManagementObjectSearcher(scope_monitor, query_monitor))
            {
                //исполняемый запрос
                foreach (ManagementObject process_monitor in searcher_monitor.Get())
                {
                    //инфо о мониторах
                    textBox14.AppendText("Info: " + process_monitor["Name"].ToString()); // что ищем
                    textBox14.AppendText(" || DeviceID: " + process_monitor["DeviceID"].ToString()); // что ищем
                }
            }

            // сетевой опрос АРМ - HDD
            ManagementScope scope_hdd = new ManagementScope("\\\\" + tB_ComputerName.Text + "\\root\\cimv2"); //область поиска. textBox7.Text - ТекстБокс с которого я возьму имя компьютера
            SelectQuery query_hdd = new SelectQuery(@"Select * from Win32_DiskDrive"); // Запрос
            //инициализируем поисковый бот для запроса выбранного для запуска
            using (ManagementObjectSearcher searcher_hdd = new ManagementObjectSearcher(scope_hdd, query_hdd))
            {
                //исполняемый запрос
                foreach (ManagementObject process_hdd in searcher_hdd.Get())
                {
                    //инфо о HDD
                    textBox15.AppendText("Type: " + process_hdd["MediaType"].ToString() + " || Model: " + process_hdd["Model"].ToString() + " || SN: " + process_hdd["SerialNumber"].ToString() + " || Interface: " + process_hdd["InterfaceType"].ToString());
                }
            }

            // сетевой опрос АРМ - DVD/CD
            ManagementScope scope_dvd = new ManagementScope("\\\\" + tB_ComputerName.Text + "\\root\\cimv2"); //область поиска. textBox7.Text - ТекстБокс с которого я возьму имя компьютера
            SelectQuery query_dvd = new SelectQuery(@"Select * from Win32_CDROMDrive"); // Запрос
            //инициализируем поисковый бот для запроса выбранного для запуска
            using (ManagementObjectSearcher searcher_dvd = new ManagementObjectSearcher(scope_dvd, query_dvd))
            {
                //исполняемый запрос
                foreach (ManagementObject process_dvd in searcher_dvd.Get())
                {
                    //инфо о DVD/CD
                    textBox16.AppendText(" || Model: " + process_dvd["Name"].ToString()); // что ищем
                }
            }
        }

        private void btn_FindNetworkComputer_Click(object sender, EventArgs e) 
        {
            //Список компьютеров в доменной сети, запрос к AD
            //try
            //{
            //    DirectoryEntry enTry = new DirectoryEntry("LDAP://OU=Computers,DC=varshots,DC=ru");
            //    DirectorySearcher mySearcher = new DirectorySearcher(enTry);
            //    int UF_ACCOUNTDISABLE = 0x0002; // Исключаем из поиска отключенный компьютеры
            //    String searchFilter = "(&(objectClass=computer)(!(userAccountControl:1.2.840.113556.1.4.803:=" + UF_ACCOUNTDISABLE.ToString() + ")))";
            //    mySearcher.Filter = (searchFilter);
            //    SearchResultCollection resEnt = mySearcher.FindAll();
            //    foreach (SearchResult srItem in resEnt)
            //    {
            //        listBox1.Items.Add(srItem.GetDirectoryEntry().Name.ToString().Substring(3).ToUpper()); //Добавляю список найденных компов в listbox 
            //    }
            //}
            //catch (Exception){
            //
            //}
        }

        private void button1_Click(object sender, EventArgs e)
        {
            

            
            // локальный опрос АРМ - Общее инфо
            // локальный опрос АРМ - Имя АРМ + локальный опрос АРМ - Пользователь

            // локальный опрос АРМ - ОС
            // WMI запрос к локальному компьютеру - информация о ОС
            ManagementObjectSearcher searcher_os = new ManagementObjectSearcher("SELECT Name FROM Win32_OperatingSystem"); //область поиска
            ManagementObjectCollection searcherCollection_os = searcher_os.Get();
            foreach (ManagementObject mo in searcherCollection_os)
            {
                foreach (PropertyData prop in mo.Properties)
                {
                    textBox4.Text = prop.Value.ToString(); //вывожу наименование ОС в метку на форме
                }
            }

            // локальный опрос АРМ - Имя АРМ
            // WMI запрос к локальному компьютеру - Имя АРМ
            ManagementObjectSearcher searcher_hostname = new ManagementObjectSearcher("SELECT Name FROM Win32_ComputerSystem"); //область поиска
            ManagementObjectCollection searcherCollection_hostname = searcher_hostname.Get();
            foreach (ManagementObject mo in searcherCollection_hostname)
            {
                foreach (PropertyData prop in mo.Properties)
                {
                    textBox5.Text = prop.Value.ToString(); //вывожу Имя АРМ в метку на форме
                }
            }

            // локальный опрос АРМ - Пользователь
            // WMI запрос к локальному компьютеру - Пользователь
            ManagementObjectSearcher searcher_username = new ManagementObjectSearcher("SELECT username FROM Win32_ComputerSystem"); //область поиска
            ManagementObjectCollection searcherCollection_username = searcher_username.Get();
            foreach (ManagementObject mo in searcherCollection_username)
            {
                foreach (PropertyData prop in mo.Properties)
                {
                    textBox6.Text = prop.Value.ToString(); //вывожу активного юзера в метку на форме
                }
            }

            // локальный опрос АРМ - Общее инфо
            // WMI запрос к локальному компьютеру - Общее инфо
            textBox3.Text = textBox6.Text + " на " + textBox5.Text;

            // локальный опрос АРМ - Сетевая карта
            // WMI запрос к локальному компьютеру - Сетевая карта

            ///// опрос делаем при инициализации формы и соотв. вызове  - MainForm_OnLoad 

            // локальный опрос АРМ - МАС
            // WMI запрос к локальному компьютеру - МАС
            ManagementObjectSearcher mos = new ManagementObjectSearcher("select * from Win32_NetworkAdapter where Name='" + comboBox1.SelectedItem.ToString() + "'"); //область поиска

            ManagementObjectCollection moc = mos.Get();

            if (moc.Count > 0)

            {

                foreach (ManagementObject mo in moc)

                {

                    textBox8.Text = (string)mo["MACAddress"]; // что ищем

                }

            }

            // локальный опрос АРМ - IP
            // WMI запрос к локальному компьютеру - IP
            textBox9.AppendText(" || ");

            ManagementObjectSearcher searcher_ip = new ManagementObjectSearcher("SELECT IPAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = 'TRUE'"); //область поиска
            ManagementObjectCollection searcherCollection_ip = searcher_ip.Get();
            foreach (ManagementObject mo in searcherCollection_ip)
            {

                {
                    string[] arrIPAddress = (string[])(mo["IPAddress"]); // что ищем

                    string sIPAddress = arrIPAddress[0];

                    textBox9.AppendText(sIPAddress.ToString());
                    textBox9.AppendText(" || ");


                }
            }

            // +
            // локальный опрос АРМ - Материнская плата
            // WMI запрос к локальному компьютеру - Материнская плата модель + продукт ID + SN
            SelectQuery query_motherboard = new SelectQuery(@"Select * from Win32_BaseBoard");

            //инициализируем поисковый бот для запроса выбранного для запуска
            using (ManagementObjectSearcher searcher_motherboard = new ManagementObjectSearcher(query_motherboard))
            {
                //исполняемый запрос
                foreach (ManagementObject process_motherboard in searcher_motherboard.Get())
                {
                    //инфо о motherboard (материнской плате)
                    //вывожу название матплаты в метку на форме
                    textBox10.AppendText(" || Model: " + process_motherboard["Manufacturer"].ToString()); // WMI запрос к локальному компьютеру - Материнская плата модель
                    textBox10.AppendText(" || Product: " + process_motherboard["Product"].ToString()); // WMI запрос к локальному компьютеру - Материнская плата продукт ID
                    textBox10.AppendText(" || SN.: " + process_motherboard["SerialNumber"].ToString()); // WMI запрос к локальному компьютеру - Материнская плата серийный номер
                }
            }

            // +
            // локальный опрос АРМ - BIOS
            // WMI запрос к локальному компьютеру - BIOS
            SelectQuery query_bios = new SelectQuery(@"Select * from Win32_BIOS");

            //инициализируем поисковый бот для запроса выбранного для запуска
            using (ManagementObjectSearcher searcher_bios = new ManagementObjectSearcher(query_bios))
            {
                //исполняемый запрос
                foreach (ManagementObject process_bios in searcher_bios.Get())
                {
                    //инфо о BIOS
                    textBox11.AppendText(" || BIOS Info name & ver.: " + process_bios["Name"].ToString()); // что ищем
                }
            }


            // +
            // локальный опрос АРМ - Процессор (CPU)
            // WMI запрос к локальному компьютеру - Процессор (CPU)
            //инициализируем выбранный запрос через команду
            SelectQuery query_cpu = new SelectQuery(@"Select * from Win32_Processor");

            //инициализируем поисковый бот для запроса выбранного для запуска
            using (ManagementObjectSearcher searcher_cpu = new ManagementObjectSearcher(query_cpu))
            {
                //исполняемый запрос
                foreach (ManagementObject process_cpu in searcher_cpu.Get())
                {
                    //инфо о Процессорах (CPU)
                    textBox12.AppendText(" || Info: " + process_cpu["Name"].ToString()); // что ищем
                }
            }


            // +
            // локальный опрос АРМ - Видео-адаптер
            // WMI запрос к локальному компьютеру - Видео-адаптер
            //инициализируем выбранный запрос через команду
            SelectQuery query_video = new SelectQuery(@"Select * from Win32_VideoController");

            //инициализируем поисковый бот для запроса выбранного для запуска
            using (ManagementObjectSearcher searcher_video = new ManagementObjectSearcher(query_video))
            {
                //исполняемый запрос
                foreach (ManagementObject process_video in searcher_video.Get())
                {
                    //инфо о Видео-адаптервх
                    textBox13.AppendText(" || Info: " + process_video["Description"].ToString()); // что ищем
                }
            }


            // +
            // локальный опрос АРМ - Монитор
            // WMI запрос к локальному компьютеру - Монитор имя
            //инициализируем выбранный запрос через команду
            SelectQuery query_monitor = new SelectQuery(@"Select * from Win32_DesktopMonitor");

            //инициализируем поисковый бот для запроса выбранного для запуска
            using (ManagementObjectSearcher searcher_monitor = new ManagementObjectSearcher(query_monitor))
            {
                //исполняемый запрос
                foreach (ManagementObject process_monitor in searcher_monitor.Get())
                {
                    //инфо о мониторах
                    textBox14.AppendText("Info: " + process_monitor["Name"].ToString()); // что ищем
                    textBox14.AppendText(" || DeviceID: " + process_monitor["DeviceID"].ToString()); // что ищем
                }
            }


            // +
            // локальный опрос АРМ - HDD
            // WMI запрос к локальному компьютеру - HDD
            ManagementObjectSearcher mosDisks = new ManagementObjectSearcher("SELECT * FROM Win32_DiskDrive"); //область поиска
            foreach (ManagementObject moDisk in mosDisks.Get())
            {

                textBox15.AppendText("Type: " + moDisk["MediaType"].ToString() + " || Model: " + moDisk["Model"].ToString() + " || SN: " + moDisk["SerialNumber"].ToString() + " || Interface: " + moDisk["InterfaceType"].ToString()); // что ищем

            }



            // локальный опрос АРМ - DVD-RW/R/ROM
            // WMI запрос к локальному компьютеру - 
            //ManagementObjectSearcher searcher_dvd = new ManagementObjectSearcher("SELECT Name FROM Win32_CDROMDrive");  //область поиска
            //foreach (ManagementObject moDisk in mosDisks.Get())
            //{
            //       textBox16.AppendText("Model: " + moDisk["Caption"].ToString()); // что ищем
            //}


            // +
            //инициализируем выбранный запрос через команду
            SelectQuery query_dvdinfo = new SelectQuery(@"Select * from Win32_CDROMDrive");

            //инициализируем поисковый бот для запроса выбранного для запуска
            using (ManagementObjectSearcher searcher_dvdinfo = new ManagementObjectSearcher(query_dvdinfo))
            {
                //исполняемый запрос
                foreach (ManagementObject process_dvdinfo in searcher_dvdinfo.Get())
                {
                    //инфо о CDROM/DVDROM
                    textBox16.AppendText("Model: " + process_dvdinfo["Name"].ToString()); // что ищем
                }
            }



        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            ManagementObjectSearcher mos = new ManagementObjectSearcher("select * from Win32_NetworkAdapter Where AdapterType='Ethernet 802.3'");  //область поиска и какой интерфейс ищем, соотв eth0

            foreach (ManagementObject mo in mos.Get())

            {

                comboBox1.Items.Add(mo["Name"].ToString());

            }

            // фокус на поле сканирования IP адресов
            this.ActiveControl = textBox_IP;
        }

        // сделаем коллбэк для данных пинга 8.8.8.8
        public static void Callback8888(object state)
        {
            var ping = new Ping();
            var ipAddress = IPAddress.Parse((String)state);

            var pingReply = ping.Send(ipAddress, 1000);

            (Application.OpenForms[0] as MainForm).Invoke((MethodInvoker)(delegate ()
            {
                (Application.OpenForms[0] as MainForm).toolStripStatusLabel2.Text = "[ " + DateTime.UtcNow.ToString() + " ] [ " + ipAddress + " ] [ " + pingReply.RoundtripTime + " ms ] [ " + pingReply.Status + " ]";
            }));
        }

        // сделаем коллбэк для данных пинга 1.1.1.1
        public static void Callback1111(object state)
        {
            var ping = new Ping();
            var ipAddress = IPAddress.Parse((String)state);

            var pingReply = ping.Send(ipAddress, 1000);

            (Application.OpenForms[0] as MainForm).Invoke((MethodInvoker)(delegate ()
            {
                (Application.OpenForms[0] as MainForm).toolStripStatusLabel4.Text = "[ " + DateTime.UtcNow.ToString() + " ] [ " + ipAddress + " ] [ " + pingReply.RoundtripTime + " ms ] [ " + pingReply.Status + " ]";
            }));
        }

        // сделаем коллбэк для данных пинга выбранной машины
        public static void CallbackARM(object state)
        {
            var ping = new Ping();
            var ipAddress = IPAddress.Parse((String)state);

            var pingReply = ping.Send(ipAddress, 1000);

            (Application.OpenForms[0] as MainForm).Invoke((MethodInvoker)(delegate ()
            {
                (Application.OpenForms[0] as MainForm).toolStripStatusLabel7.Text = "[ " + DateTime.UtcNow.ToString() + " ] [ " + ipAddress + " ] [ " + pingReply.RoundtripTime + " ms ] [ " + pingReply.Status + " ]";
            }));
        }

        private void MainForm_Shown(object sender, EventArgs e)
        {
            // показываем главный eth0 адаптер
            comboBox1.SelectedIndex = 0;

            // инициализируем date-time
            DateTime now = DateTime.Now;
            CultureInfo culture = new CultureInfo("ru-RU"); // "святая" Российская Федерация
            Thread.CurrentThread.CurrentCulture = culture;
            textBox2.Text = (now.ToString("dd/MM/yyyy"));

            //Console.WriteLine(now.ToString("yyyy-MM-ddTHH:mm:ss.fff"));

            ///////////////////////////////////////////////////////////
            // инициализируем систему подготовки пинга как внешней так и внутренней сети

            var period = 10 * 1000;
            // активизируем таймер пинга публичных гугл dns 8.8.8.8
            string ipAddress8888 = "8.8.8.8";
            System.Threading.Timer t = new System.Threading.Timer(Callback8888, ipAddress8888, 0, period);
            // активизируем таймер пинга публичных гугл dns 8.8.4.4
            string ipAddress1111 = "1.1.1.1";
            System.Threading.Timer t2 = new System.Threading.Timer(Callback1111, ipAddress1111, 0, period);


        }

        // очищаем список просканированных IP адресов
        private void button2_Click(object sender, EventArgs e)
        {
            listBox2.Items.Clear(); // чистим весь список
        }

        // как пример как тянуть через запрос в АД список АРМ в доменной сети
        //Список компьютеров в доменной сети, запрос к AD
        //try
        //{
        //    DirectoryEntry enTry = new DirectoryEntry("LDAP://OU=Computers,DC=varshots,DC=ru");
        //    DirectorySearcher mySearcher = new DirectorySearcher(enTry);
        //    int UF_ACCOUNTDISABLE = 0x0002; // Исключаем из поиска отключенный компьютеры
        //    String searchFilter = "(&(objectClass=computer)(!(userAccountControl:1.2.840.113556.1.4.803:=" + UF_ACCOUNTDISABLE.ToString() + ")))";
        //    mySearcher.Filter = (searchFilter);
        //    SearchResultCollection resEnt = mySearcher.FindAll();
        //    foreach (SearchResult srItem in resEnt)
        //    {
        //        listBox1.Items.Add(srItem.GetDirectoryEntry().Name.ToString().Substring(3).ToUpper()); //Добавляю список найденных компов в listbox 
        //    }
        //}
        //catch (Exception){
        //
        //}


        // сканирование напр IP "от" и "до", напр.: 192.168.1.1..255
        //кнопка запуска потока
        private void button3_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Control.CheckForIllegalCrossThreadCalls = false;
            //создание отдельного потока
            Thread potok1 = new Thread(start_thread_parse);
            // завершить поток при завершении основного потока (объявлять, если точно знаете, что вам это нужно,
            //иначе поток завершится не выполнив свою работу до конца)
            potok1.IsBackground = true;
            //запуск потока
            potok1.Start();

            //процедура для выполнения потока
            // отдельный процесс для  - сканирование напр IP "от" и "до", напр.: 192.168.1.1..255
            void start_thread_parse()
            {
                for (int i = Convert.ToInt32((this.textBox_IP4_n.Text)); i <= Convert.ToInt32((this.textBox_IP4_k.Text)); i++)
                {
                    // TODO: Здесь выполняется то, что нужно
                    Ping pingSender = new Ping();
                    PingOptions options = new PingOptions();
                    options.DontFragment = true;
                    // Буфер 32 байта
                    string data = "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa";
                    byte[] buffer = Encoding.ASCII.GetBytes(data);
                    int timeout = 25;
                    string ping_str = this.textBox_IP.Text + "." + this.textBox_IP3_n.Text + ".";
                    //PingReply reply = pingSender.Send(string.Format("192.168.1.{0}", i), timeout, buffer, options);
                    PingReply reply = pingSender.Send(ping_str + i, timeout, buffer, options);
                    this.listBox2.Items.Add("ping " + this.textBox_IP.Text + "." + this.textBox_IP3_n.Text + "." + i + " - " + reply.Status); //Добавляю список найденных компов в listbox 
                                                                                                                                              //Console.WriteLine("ping 192.168.0.{0} - {1};", i, reply.Status);
                }
            }
        }

        // при выделении адреса АРМ
        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            // раскладываем строку на части - первая часть
            //textBox7.Text = listBox2.SelectedItem.ToString();
            string firsWord = listBox2.SelectedItem.ToString().Substring(0, listBox2.SelectedItem.ToString().IndexOf(' '));
            // раскладываем строку на части - вторая часть
            string secondSentence = listBox2.SelectedItem.ToString().Substring(listBox2.SelectedItem.ToString().IndexOf(' ') + 1);
            // раскладываем строку на части - третья часть оно же IP
            string thirdSentence = secondSentence.Substring(0, secondSentence.IndexOf(' '));
            textBox7.Text = thirdSentence;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            ///////////////////////////////////////
            // активизируем таймер пинга IP машины
            var period = 10 * 1000;
            string ipAddressARM = textBox7.Text;
            System.Threading.Timer t2 = new System.Threading.Timer(CallbackARM, ipAddressARM, 0, period);



            

            // сетевой опрос АРМ - ОС
            // WMI запрос к локальному компьютеру - Пользователь
            // инициализируем выбранный запрос через команду
            SelectQuery query_os = new SelectQuery(@"Select * from Win32_OperatingSystem");

            //инициализируем поисковый бот для запроса выбранного для запуска
            using (ManagementObjectSearcher searcher_os = new ManagementObjectSearcher(query_os))
            {
                //исполняемый запрос
                foreach (ManagementObject process_os in searcher_os.Get())
                {
                    //инфо о Пользователе
                    textBox4.AppendText(" || OS info: " + process_os["Name"].ToString()); // что ищем
                }
            }

            // сетевой опрос АРМ - Имя АРМ
            // WMI запрос к локальному компьютеру - Пользователь
            // инициализируем выбранный запрос через команду
            SelectQuery query_name = new SelectQuery(@"Select * from Win32_ComputerSystem");

            //инициализируем поисковый бот для запроса выбранного для запуска
            using (ManagementObjectSearcher searcher_name = new ManagementObjectSearcher(query_name))
            {
                //исполняемый запрос
                foreach (ManagementObject process_name in searcher_name.Get())
                {
                    //инфо о Пользователе
                    textBox5.AppendText(" || PC info: " + process_name["Name"].ToString()); // что ищем
                }
            }


            // сетевой опрос АРМ - Пользователь
            // WMI запрос к локальному компьютеру - Пользователь
            // инициализируем выбранный запрос через команду
            SelectQuery query_user = new SelectQuery(@"Select * from Win32_ComputerSystem");

            //инициализируем поисковый бот для запроса выбранного для запуска
            using (ManagementObjectSearcher searcher_user = new ManagementObjectSearcher(query_user))
            {
                //исполняемый запрос
                foreach (ManagementObject process_user in searcher_user.Get())
                {
                    //инфо о Пользователе
                    textBox6.AppendText(" || User info: " + process_user["Username"].ToString()); // что ищем
                }
            }

            // сетевой опрос АРМ - Общее инфо
            // сетевой опрос АРМ - Имя АРМ + локальный опрос АРМ - Пользователь
            textBox3.AppendText(textBox5.Text + " " + textBox6.Text); // что ищем

            // сетевой опрос АРМ - МАС адрес
            //////////////////////////////////////////////////////
            // сетевой опрос АРМ - МАС адрес
            // WMI запрос к сетевому компьютеру - МАС адрес
            //////////////////////////////////////////////////////
            ManagementScope scope_mac = new ManagementScope("\\\\" + textBox7.Text + "\\root\\cimv2"); //область поиска. textBox7.Text - ТекстБокс с которого я возьму имя компьютера
            SelectQuery query_mac = new SelectQuery(@"Select * from Win32_NetworkAdapter where Name='" + comboBox1.SelectedItem.ToString() + "'"); // Запрос
            //инициализируем поисковый бот для запроса выбранного для запуска
            using (ManagementObjectSearcher searcher_mac = new ManagementObjectSearcher(scope_mac, query_mac))
            {
                //исполняемый запрос
                foreach (ManagementObject process_mac in searcher_mac.Get())
                {
                    //инфо о МАС адресе
                    //вывожу наименование мас адреса в метку на форме
                    textBox8.Text = (string)process_mac["MACAddress"]; // что ищем
                }
            }

            // сетевой опрос АРМ - IP
            //////////////////////////////////////////////////////
            // сетевой опрос АРМ - IP
            // WMI запрос к сетевому компьютеру - IP
            //////////////////////////////////////////////////////
            ManagementScope scope_ip = new ManagementScope("\\\\" + textBox7.Text + "\\root\\cimv2"); //область поиска. textBox7.Text - ТекстБокс с которого я возьму имя компьютера
            SelectQuery query_ip = new SelectQuery(@"Select * from Win32_NetworkAdapterConfiguration WHERE IPEnabled = 'TRUE'"); // Запрос
            //инициализируем поисковый бот для запроса выбранного для запуска
            using (ManagementObjectSearcher searcher_ip = new ManagementObjectSearcher(scope_ip, query_ip))
            {
                //исполняемый запрос
                foreach (ManagementObject process_ip in searcher_ip.Get())
                {
                    //инфо о IP
                    //вывожу значение IP в метку на форме
                    string[] arrIPAddress = (string[])(process_ip["IPAddress"]); // что ищем

                    string sIPAddress = arrIPAddress[0];

                    textBox9.AppendText(" || IP: " + sIPAddress.ToString()); // WMI запрос к локальному компьютеру - Материнская плата модель
                }
            }


            // сетевой опрос АРМ - Материнская плата
            //////////////////////////////////////////////////////
            // сетевой опрос АРМ - Материнская плата
            // WMI запрос к сетевому компьютеру - Материнская плата
            //////////////////////////////////////////////////////
            ManagementScope scope_motherboard = new ManagementScope("\\\\" + textBox7.Text + "\\root\\cimv2"); //область поиска. textBox7.Text - ТекстБокс с которого я возьму имя компьютера
            SelectQuery query_motherboard = new SelectQuery(@"Select * from Win32_BaseBoard"); // Запрос
            //инициализируем поисковый бот для запроса выбранного для запуска
            using (ManagementObjectSearcher searcher_motherboard = new ManagementObjectSearcher(scope_motherboard, query_motherboard))
            {
                //исполняемый запрос
                foreach (ManagementObject process_motherboard in searcher_motherboard.Get())
                {
                    //инфо о motherboard (материнской плате)
                    //вывожу название матплаты в метку на форме
                    textBox10.AppendText(" || Model: " + process_motherboard["Manufacturer"].ToString()); // WMI запрос к локальному компьютеру - Материнская плата модель
                    textBox10.AppendText(" || Product: " + process_motherboard["Product"].ToString()); // WMI запрос к локальному компьютеру - Материнская плата продукт ID
                    textBox10.AppendText(" || SN.: " + process_motherboard["SerialNumber"].ToString()); // WMI запрос к локальному компьютеру - Материнская плата серийный номер
                }
            }

            // сетевой опрос АРМ - BIOS
            //////////////////////////////////////////////////////
            // сетевой опрос АРМ - BIOS
            // WMI запрос к сетевому компьютеру - BIOS
            //////////////////////////////////////////////////////
            ManagementScope scope_bios = new ManagementScope("\\\\" + textBox7.Text + "\\root\\cimv2"); //область поиска. textBox7.Text - ТекстБокс с которого я возьму имя компьютера
            SelectQuery query_bios = new SelectQuery(@"Select * from Win32_Processor"); // Запрос
            //инициализируем поисковый бот для запроса выбранного для запуска
            using (ManagementObjectSearcher searcher_bios = new ManagementObjectSearcher(scope_bios, query_bios))
            {
                //исполняемый запрос
                foreach (ManagementObject process_bios in searcher_bios.Get())
                {
                    //инфо о Процессорах (CPU)
                    textBox11.AppendText(" || BIOS Info name & ver.: " + process_bios["Name"].ToString()); // что ищем
                }
            }

            // сетевой опрос АРМ - CPU
            //////////////////////////////////////////////////////
            // сетевой опрос АРМ - CPU
            // WMI запрос к сетевому компьютеру - CPU 
            //////////////////////////////////////////////////////
            ManagementScope scope_cpu = new ManagementScope("\\\\" + textBox7.Text + "\\root\\cimv2"); //область поиска. textBox7.Text - ТекстБокс с которого я возьму имя компьютера
            SelectQuery query_cpu = new SelectQuery(@"Select * from Win32_Processor"); // Запрос
            //инициализируем поисковый бот для запроса выбранного для запуска
            using (ManagementObjectSearcher searcher_cpu = new ManagementObjectSearcher(scope_cpu, query_cpu))
            {
                //исполняемый запрос
                foreach (ManagementObject process_cpu in searcher_cpu.Get())
                {
                    //инфо о Процессорах (CPU)
                    textBox12.AppendText(" || Info: " + process_cpu["Name"].ToString()); // что ищем
                }
            }


            // сетевой опрос АРМ - Video-adapter
            //////////////////////////////////////////////////////
            // сетевой опрос АРМ - Video-adapter
            // WMI запрос к сетевому компьютеру - Video-adapter
            //////////////////////////////////////////////////////
            ManagementScope scope_video = new ManagementScope("\\\\" + textBox7.Text + "\\root\\cimv2"); //область поиска. textBox7.Text - ТекстБокс с которого я возьму имя компьютера
            SelectQuery query_video = new SelectQuery(@"Select * from Win32_VideoController"); // Запрос
            //инициализируем поисковый бот для запроса выбранного для запуска
            using (ManagementObjectSearcher searcher_video = new ManagementObjectSearcher(scope_video, query_video))
            {
                //исполняемый запрос
                foreach (ManagementObject process_video in searcher_video.Get())
                {
                    //инфо о Видео-адаптервх
                    textBox13.AppendText(" || Info: " + process_video["Description"].ToString()); // что ищем
                }
            }

            // сетевой опрос АРМ - Monitor
            ManagementScope scope_monitor = new ManagementScope("\\\\" + textBox7.Text + "\\root\\cimv2"); //область поиска. textBox7.Text - ТекстБокс с которого я возьму имя компьютера
            SelectQuery query_monitor = new SelectQuery(@"Select * from Win32_DesktopMonitor"); // Запрос
            //инициализируем поисковый бот для запроса выбранного для запуска
            using (ManagementObjectSearcher searcher_monitor = new ManagementObjectSearcher(scope_monitor, query_monitor))
            {
                //исполняемый запрос
                foreach (ManagementObject process_monitor in searcher_monitor.Get())
                {
                    //инфо о мониторах
                    textBox14.AppendText("Info: " + process_monitor["Name"].ToString()); // что ищем
                    textBox14.AppendText(" || DeviceID: " + process_monitor["DeviceID"].ToString()); // что ищем
                }
            }

            // сетевой опрос АРМ - HDD
            ManagementScope scope_hdd = new ManagementScope("\\\\" + textBox7.Text + "\\root\\cimv2"); //область поиска. textBox7.Text - ТекстБокс с которого я возьму имя компьютера
            SelectQuery query_hdd = new SelectQuery(@"Select * from Win32_DiskDrive"); // Запрос
            //инициализируем поисковый бот для запроса выбранного для запуска
            using (ManagementObjectSearcher searcher_hdd = new ManagementObjectSearcher(scope_hdd, query_hdd))
            {
                //исполняемый запрос
                foreach (ManagementObject process_hdd in searcher_hdd.Get())
                {
                    //инфо о HDD
                    textBox15.AppendText("Type: " + process_hdd["MediaType"].ToString() + " || Model: " + process_hdd["Model"].ToString() + " || SN: " + process_hdd["SerialNumber"].ToString() + " || Interface: " + process_hdd["InterfaceType"].ToString());
                }
            }

            // сетевой опрос АРМ - DVD/CD
            ManagementScope scope_dvd = new ManagementScope("\\\\" + textBox7.Text + "\\root\\cimv2"); //область поиска. textBox7.Text - ТекстБокс с которого я возьму имя компьютера
            SelectQuery query_dvd = new SelectQuery(@"Select * from Win32_CDROMDrive"); // Запрос
            //инициализируем поисковый бот для запроса выбранного для запуска
            using (ManagementObjectSearcher searcher_dvd = new ManagementObjectSearcher(scope_dvd, query_dvd))
            {
                //исполняемый запрос
                foreach (ManagementObject process_dvd in searcher_dvd.Get())
                {
                    //инфо о DVD/CD
                    textBox16.AppendText(" || Model: " + process_dvd["Name"].ToString()); // что ищем
                }
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            // очистка значений опроса..

            // Общее инфо (очистка поля и обнуление поля)
            textBox3.Text = "";

            // ОС  (очистка поля и обнуление поля)
            textBox4.Text = "";

            // Имя АРМ (очистка поля и обнуление поля)
            textBox5.Text = "";

            // Пользователь (очистка поля и сброс к дефолтным значениям)
            textBox6.Text = "";

            // Сетевая карта (очистка поля и обнуление поля)
            ManagementObjectSearcher mos = new ManagementObjectSearcher("select * from Win32_NetworkAdapter Where AdapterType='Ethernet 802.3'");  //область поиска и какой интерфейс ищем, соотв eth0
            foreach (ManagementObject mo in mos.Get())
            {
                comboBox1.Items.Add(mo["Name"].ToString());
            }
            // фокус на поле сканирования IP адресов
            this.ActiveControl = textBox_IP;

            // МАС (очистка поля и обнуление поля)
            textBox8.Text = "";

            // IP (очистка поля и обнуление поля)
            textBox9.Text = "";

            // Материнская плата (очистка поля и обнуление поля)
            textBox10.Text = "";

            // BIOS (очистка поля и обнуление поля)
            textBox11.Text = "";

            // CPU (очистка поля и обнуление поля)
            textBox12.Text = "";

            // Видео-адаптер (очистка поля и обнуление поля)
            textBox13.Text = "";

            // Монитор (очистка поля и обнуление поля)
            textBox14.Text = "";

            // HDD (очистка поля и обнуление поля)
            textBox15.Text = "";

            // DVD-RW/R/ROM (очистка поля и обнуление поля)
            textBox16.Text = "";

            // Имя компьютера вручную(очистка поля и обнуление поля)
            tB_ComputerName.Text = "";

            // Выбранный АРМ из опрошенных(очистка поля и обнуление поля)
            textBox7.Text = "";
        }

        private void button6_Click(object sender, EventArgs e)
        {
            using (StreamWriter sw = new StreamWriter(@Directory.GetCurrentDirectory() + "\\Картотека\\" + textBox1.Text + ".txt"))
            {
                sw.Write(textBox2.Text + Environment.NewLine);
                sw.Write("======== Сведения о местоположении машины ========" + Environment.NewLine);
                sw.Write("Имя компьютера вручную = " + tB_ComputerName.Text + Environment.NewLine);
                sw.Write("IP из ранее опрошенных АРМ = " + textBox7.Text + Environment.NewLine);
                sw.Write("======== Общие сведения о АРМ и пользователе ========" + Environment.NewLine);
                sw.Write(textBox3.Text + Environment.NewLine);
                sw.Write(textBox4.Text + Environment.NewLine);
                sw.Write(textBox5.Text + Environment.NewLine);
                sw.Write(textBox6.Text + Environment.NewLine);
                sw.Write("======== Сетевые характеристики ========" + Environment.NewLine);
                sw.Write(comboBox1.Text + Environment.NewLine);
                sw.Write(textBox8.Text + Environment.NewLine);
                sw.Write(textBox9.Text + Environment.NewLine);
                sw.Write("======== Опрос АРМ на предмет программно-аппаратных характеристик ========" + Environment.NewLine);
                sw.Write(textBox10.Text + Environment.NewLine);
                sw.Write(textBox11.Text + Environment.NewLine);
                sw.Write(textBox12.Text + Environment.NewLine);
                sw.Write(textBox13.Text + Environment.NewLine);
                sw.Write(textBox14.Text + Environment.NewLine);
                sw.Write(textBox15.Text + Environment.NewLine);
                sw.Write(textBox16.Text + Environment.NewLine);
                sw.Write("================" + Environment.NewLine);

                MessageBox.Show("Файл {"+ textBox1.Text +"}, сохранен в картотеке Inventory IT по адресу - " + Directory.GetCurrentDirectory() + "\\Картотека\\", "Сохранение в картотеке..", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

               
            }

            // создаем рабочую книгу
            using (ExcelPackage excel = new ExcelPackage())
            {
                excel.Workbook.Worksheets.Add(textBox1.Text);
                excel.Workbook.Worksheets.Add("Worksheet2");
                excel.Workbook.Worksheets.Add("Worksheet3");

                // Добавить строку
                //List<string[]> headerRow = new List<string[]>()
                //{
                //  new string[] { "ID", "First Name", "Last Name", "DOB" }
                //};

                // определяем диапазон заголовка
                //string headerRange = "A1:" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + "1";

                // выбираем активный лист
                var worksheet = excel.Workbook.Worksheets[textBox1.Text];

                // строка данных данного заголовка
                //worksheet.Cells[headerRange].LoadFromArrays(headerRow);

                // определяем заголовки левого края для описаний
                worksheet.Cells["A1"].Value = "Текущая дата";
                worksheet.Cells["A2"].Value = "АРМ и инв.номер";
                worksheet.Cells["A3"].Value = "Имя компьютера вручную";
                worksheet.Cells["A4"].Value = "IP из ранее опрошенных АРМ";
                worksheet.Cells["A5"].Value = "Общее инфо";
                worksheet.Cells["A6"].Value = "ОС";
                worksheet.Cells["A7"].Value = "Имя АРМ";
                worksheet.Cells["A8"].Value = "Пользователь";
                worksheet.Cells["A9"].Value = "Сетевая карта";
                worksheet.Cells["A10"].Value = "MAC";
                worksheet.Cells["A11"].Value = "IP";
                worksheet.Cells["A12"].Value = "Материнская плата";
                worksheet.Cells["A13"].Value = "BIOS";
                worksheet.Cells["A14"].Value = "CPU";
                worksheet.Cells["A15"].Value = "Видео-адаптер";
                worksheet.Cells["A16"].Value = "Монитор";
                worksheet.Cells["A17"].Value = "HDD";
                worksheet.Cells["A18"].Value = "DVD-RW/R/ROM";

                // определяем стиль
                worksheet.Cells["A1"].Style.Font.Bold = true;
                worksheet.Cells["A1"].Style.Font.Size = 12;
                worksheet.Cells["A1"].Style.Font.Color.SetColor(System.Drawing.Color.Navy);

                // определяем стиль
                worksheet.Cells["A2"].Style.Font.Bold = true;
                worksheet.Cells["A2"].Style.Font.Size = 12;
                worksheet.Cells["A2"].Style.Font.Color.SetColor(System.Drawing.Color.Navy);

                // определяем стиль
                worksheet.Cells["A3"].Style.Font.Bold = true;
                worksheet.Cells["A3"].Style.Font.Size = 12;
                worksheet.Cells["A3"].Style.Font.Color.SetColor(System.Drawing.Color.Navy);

                // определяем стиль
                worksheet.Cells["A4"].Style.Font.Bold = true;
                worksheet.Cells["A4"].Style.Font.Size = 12;
                worksheet.Cells["A4"].Style.Font.Color.SetColor(System.Drawing.Color.Navy);

                // определяем стиль
                worksheet.Cells["A5"].Style.Font.Bold = true;
                worksheet.Cells["A5"].Style.Font.Size = 12;
                worksheet.Cells["A5"].Style.Font.Color.SetColor(System.Drawing.Color.Navy);

                // определяем стиль
                worksheet.Cells["A6"].Style.Font.Bold = true;
                worksheet.Cells["A6"].Style.Font.Size = 12;
                worksheet.Cells["A6"].Style.Font.Color.SetColor(System.Drawing.Color.Navy);

                // определяем стиль
                worksheet.Cells["A7"].Style.Font.Bold = true;
                worksheet.Cells["A7"].Style.Font.Size = 12;
                worksheet.Cells["A7"].Style.Font.Color.SetColor(System.Drawing.Color.Navy);

                // определяем стиль
                worksheet.Cells["A8"].Style.Font.Bold = true;
                worksheet.Cells["A8"].Style.Font.Size = 12;
                worksheet.Cells["A8"].Style.Font.Color.SetColor(System.Drawing.Color.Navy);

                // определяем стиль
                worksheet.Cells["A9"].Style.Font.Bold = true;
                worksheet.Cells["A9"].Style.Font.Size = 12;
                worksheet.Cells["A9"].Style.Font.Color.SetColor(System.Drawing.Color.Navy);

                // определяем стиль
                worksheet.Cells["A10"].Style.Font.Bold = true;
                worksheet.Cells["A10"].Style.Font.Size = 12;
                worksheet.Cells["A10"].Style.Font.Color.SetColor(System.Drawing.Color.Navy);

                // определяем стиль
                worksheet.Cells["A11"].Style.Font.Bold = true;
                worksheet.Cells["A11"].Style.Font.Size = 12;
                worksheet.Cells["A11"].Style.Font.Color.SetColor(System.Drawing.Color.Navy);

                // определяем стиль
                worksheet.Cells["A12"].Style.Font.Bold = true;
                worksheet.Cells["A12"].Style.Font.Size = 12;
                worksheet.Cells["A12"].Style.Font.Color.SetColor(System.Drawing.Color.Navy);

                // определяем стиль
                worksheet.Cells["A13"].Style.Font.Bold = true;
                worksheet.Cells["A13"].Style.Font.Size = 12;
                worksheet.Cells["A13"].Style.Font.Color.SetColor(System.Drawing.Color.Navy);

                // определяем стиль
                worksheet.Cells["A14"].Style.Font.Bold = true;
                worksheet.Cells["A14"].Style.Font.Size = 12;
                worksheet.Cells["A14"].Style.Font.Color.SetColor(System.Drawing.Color.Navy);

                // определяем стиль
                worksheet.Cells["A15"].Style.Font.Bold = true;
                worksheet.Cells["A15"].Style.Font.Size = 12;
                worksheet.Cells["A15"].Style.Font.Color.SetColor(System.Drawing.Color.Navy);

                // определяем стиль
                worksheet.Cells["A16"].Style.Font.Bold = true;
                worksheet.Cells["A16"].Style.Font.Size = 12;
                worksheet.Cells["A16"].Style.Font.Color.SetColor(System.Drawing.Color.Navy);

                // определяем стиль
                worksheet.Cells["A17"].Style.Font.Bold = true;
                worksheet.Cells["A17"].Style.Font.Size = 12;
                worksheet.Cells["A17"].Style.Font.Color.SetColor(System.Drawing.Color.Navy);

                // определяем стиль
                worksheet.Cells["A18"].Style.Font.Bold = true;
                worksheet.Cells["A18"].Style.Font.Size = 12;
                worksheet.Cells["A18"].Style.Font.Color.SetColor(System.Drawing.Color.Navy);

                //////////////////////////////////////////////
                // определяем заголовки левого края для данных
                worksheet.Cells["B1"].Value = textBox2.Text;
                worksheet.Cells["B2"].Value = textBox1.Text;
                //
                if (tB_ComputerName.Text == "")
                {
                    tB_ComputerName.Text = "Не определено";
                }
                worksheet.Cells["B3"].Value = tB_ComputerName.Text;
                //
                if (textBox7.Text == "")
                {
                    textBox7.Text = "Не определено";
                }
                worksheet.Cells["B4"].Value = textBox7.Text;
                //
                worksheet.Cells["B5"].Value = textBox3.Text;
                worksheet.Cells["B6"].Value = textBox4.Text;
                worksheet.Cells["B7"].Value = textBox5.Text;
                worksheet.Cells["B8"].Value = textBox6.Text;
                worksheet.Cells["B9"].Value = comboBox1.Text;
                worksheet.Cells["B10"].Value = textBox8.Text;
                worksheet.Cells["B11"].Value = textBox9.Text;
                worksheet.Cells["B12"].Value = textBox10.Text;
                worksheet.Cells["B13"].Value = textBox11.Text;
                worksheet.Cells["B14"].Value = textBox12.Text;
                worksheet.Cells["B15"].Value = textBox13.Text;
                worksheet.Cells["B16"].Value = textBox14.Text;
                worksheet.Cells["B17"].Value = textBox15.Text;
                //
                if (textBox16.Text == "") {
                    textBox16.Text = "Не обнаружено";    
                }
                worksheet.Cells["B18"].Value = textBox16.Text;
                //
                // определяем стиль
                worksheet.Cells["B1"].Style.Font.Size = 12;
                worksheet.Cells["B1"].Style.Font.Color.SetColor(System.Drawing.Color.DarkRed);

                // определяем стиль
                worksheet.Cells["B2"].Style.Font.Size = 12;
                worksheet.Cells["B2"].Style.Font.Color.SetColor(System.Drawing.Color.DarkRed);

                // определяем стиль
                worksheet.Cells["B3"].Style.Font.Size = 12;
                worksheet.Cells["B3"].Style.Font.Color.SetColor(System.Drawing.Color.DarkRed);

                // определяем стиль
                worksheet.Cells["B4"].Style.Font.Size = 12;
                worksheet.Cells["B4"].Style.Font.Color.SetColor(System.Drawing.Color.DarkRed);

                // определяем стиль
                worksheet.Cells["B5"].Style.Font.Size = 12;
                worksheet.Cells["B5"].Style.Font.Color.SetColor(System.Drawing.Color.DarkRed);

                // определяем стиль
                worksheet.Cells["B6"].Style.Font.Size = 12;
                worksheet.Cells["B6"].Style.Font.Color.SetColor(System.Drawing.Color.DarkRed);

                // определяем стиль
                worksheet.Cells["B7"].Style.Font.Size = 12;
                worksheet.Cells["B7"].Style.Font.Color.SetColor(System.Drawing.Color.DarkRed);

                // определяем стиль
                worksheet.Cells["B8"].Style.Font.Size = 12;
                worksheet.Cells["B8"].Style.Font.Color.SetColor(System.Drawing.Color.DarkRed);

                // определяем стиль
                worksheet.Cells["B9"].Style.Font.Size = 12;
                worksheet.Cells["B9"].Style.Font.Color.SetColor(System.Drawing.Color.DarkRed);

                // определяем стиль
                worksheet.Cells["B10"].Style.Font.Size = 12;
                worksheet.Cells["B10"].Style.Font.Color.SetColor(System.Drawing.Color.DarkRed);

                // определяем стиль
                worksheet.Cells["B11"].Style.Font.Size = 12;
                worksheet.Cells["B11"].Style.Font.Color.SetColor(System.Drawing.Color.DarkRed);

                // определяем стиль
                worksheet.Cells["B12"].Style.Font.Size = 12;
                worksheet.Cells["B12"].Style.Font.Color.SetColor(System.Drawing.Color.DarkRed);

                // определяем стиль
                worksheet.Cells["B13"].Style.Font.Size = 12;
                worksheet.Cells["B13"].Style.Font.Color.SetColor(System.Drawing.Color.DarkRed);

                // определяем стиль
                worksheet.Cells["B14"].Style.Font.Size = 12;
                worksheet.Cells["B14"].Style.Font.Color.SetColor(System.Drawing.Color.DarkRed);

                // определяем стиль
                worksheet.Cells["B15"].Style.Font.Size = 12;
                worksheet.Cells["B15"].Style.Font.Color.SetColor(System.Drawing.Color.DarkRed);

                // определяем стиль
                worksheet.Cells["B16"].Style.Font.Size = 12;
                worksheet.Cells["B16"].Style.Font.Color.SetColor(System.Drawing.Color.DarkRed);

                // определяем стиль
                worksheet.Cells["B17"].Style.Font.Size = 12;
                worksheet.Cells["B17"].Style.Font.Color.SetColor(System.Drawing.Color.DarkRed);

                // определяем стиль
                worksheet.Cells["B18"].Style.Font.Size = 12;
                worksheet.Cells["B18"].Style.Font.Color.SetColor(System.Drawing.Color.DarkRed);

                // опеределяем имя файла и путь
                FileInfo excelFile = new FileInfo(@Directory.GetCurrentDirectory() + "\\Картотека\\" + textBox1.Text + ".xlsx");

                // сохраняем
                excel.SaveAs(excelFile);
            }

        }
    }
}
