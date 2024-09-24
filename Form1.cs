using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Text;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace UpdateAD
{
    public partial class Form1 : Form
    {
        private PrincipalContext _context;
        private DirectoryEntry usersOU;

        public Form1()
        {
            InitializeComponent();
            InitializeDirectoryContext();
            LoadGroups();
            this.FormClosing += YourForm_FormClosing;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }
        private void LoadGroups()
        {
            string[] groups = {
            "_INET_CITADEL", "_Library", "_nextcloud", "1-я кафедра внутренних болезней", "1-я кафедра детских болезней",
            "2-я кафедра внутренних болезней", "2-я кафедра детских болезней", "3-я кафедра внутренних болезней", "All", "all_units",
            "anketa", "del SMBusiness", "disk I аттестация ИПК", "disk I кафедра гигиены", "disk I кафедра УЗИ ИПК",
            "disk I кафедра финансового менеджмента и информатизации ИПК", "disk I ОЗиЗ ИПК", "disk I САЦ", "disk I спорткомплекс ИПК", "disk I терапевтический факультет ИПК",
            "disk I факультет охраны материнства и детства  ИПК", "disk I хирургический факультет ИПК", "disk I ЦИТ", "disk I ЭТО ИПК", "disk Q IT-LAB",
            "disk Q LPO", "disk Q smk", "disk Q аудит", "disk Q библиотека", "disk Q бухгалтерия", "disk Q деканат военно медицинского факультета", "disk Q деканат лечебного факультета", "disk Q деканат МФИУ", "disk Q деканат ПДП",
            "disk Q деканат медпроф факультета", "disk Q деканат педиатрического факультета", "disk Q деканат стоматологического факультета", "disk Q деканат фармацевтического факультета", "disk Q диссертационный совет",
            "disk Q интернатура", "disk Q канцелярия", "disk Q кафедрa ортопедической стоматологии", "disk Q кафедра гигиены труда", "disk Q кафедра иностранных языков",
            "disk Q кафедра нормальной физиологии", "disk Q кафедра физического воспитания и спорта", "disk Q международный отдел", "disk Q НИЧ", "disk Q ОНПМ и ОМЗ", "disk Q ОНПМ и ОМЗ", "disk Q отдел воспитательной работы", "disk Q отдел договорных работ", "disk Q отдел идеологической работы",
            "disk Q отдел кадров", "disk Q отдел метод обеспечения доп образования взрослых", "disk Q отдел научных работников высшей категории", "disk Q отдел снабжения", "disk Q отдел труда и заработной платы",
            "disk Q профком", "disk Q приемная коммисия", "disk Q подготовительное отделение", "disk Q патентный отдел", "disk Q охрана труда",
            "disk Q учебный отдел", "disk Q учебно методический отдел", "disk Q студ клуб", "disk Q студ городок", "disk Q столовая", "disk Q спорт клуб", "disk Q РИО", "disk Q ПФО", "disk Q психологи",
            "disk Q юристы", "disk Q ЭТО", "disk Q ЦРИТ", "disk Q ЦППИКО", "disk Q учсовет",
            "Inet_new", "IE for EDGE", "ESX power and consle admin", "ESX admin", "disk Z LAB",
            "update_wsus", "Strong password", "SMBusiness ", "Simple password", "server_admin", "proxy_citadel", "moodle_course_creators", "Lock", "Kaspersky_managers",
            "Кафедра акушерства и гинекологии", "Доступ к сервисам мгновенных сообщенией", "Доступ к дискам КК", "Администраторы диска Q", "Wi-Fi",
            "Кафедра биоорганической химии disk Q", "Кафедра биологической химии disk Q", "Кафедра биологии", "Кафедра белорусского и русскго языка", "Кафедра анестезиологии и реаниматологии",
            "Кафедра детских инфекционных болезней", "Кафедра глазных болезней", "Кафедра гистологии", "Кафедра гигиены труда", "Кафедра гигиены детей и подростков", "Кафедра военно-полевой хирургии", "Кафедра военно-полевой терапии", "Кафедра военной эпидемиологии и военной гигиены", "Кафедра болезней уха горла носа",
            "Кафедра кардиологии и внутренних болезней", "Кафедра инфекционных болезней", "Кафедра иностранных языков diskQ", "Кафедра детской эндокринологии клинической генетики и иммунологии", "Кафедра детской хирургии",
            "Кафедра лучевой диагностики и лучевой терапии", "Кафедра латинского языка disk Q", "Кафедра консервативной стоматологии", "Кафедра кожных и венерических болезней", "Кафедра клинической фармакологии",
            "Кафедра нормальной физиологии для ограничений", "Кафедра нормальной физиологии  disk Q", "Кафедра нормальной анатомия", "Кафедра нервных и нейрохирургических болезней", "Кафедра морфологии человека", "Кафедра микробиологии вирусологии иммунологии", "Кафедра медицинской реабилитации и физиотерапии", "Кафедра медицинской и биологической физики disk Q", "Кафедра медицинская и биологической физики",
            "Кафедра организации медицинского обеспечения войск и экстремальн", "Кафедра оперативой хирургии и топографической анатомии", "Кафедра онкологии", "Кафедра общественного здоровья и здравоохранения", "Кафедра общей хирургии", "Кафедра общей химии", "Кафедра общей стоматологии disk Q", "Кафедра общей гигиены", "Кафедра общей врачебной практики",
            "Кафедра патологической анатомии", "Кафедра паталогической физиологии", "Кафедра ортопедической стоматологии", "Кафедра ортодонтии", "Кафедра организации фармации disk Q",
            "Кафедра психиатрии и медицинской психологии", "Кафедра пропедевтики детских болезней", "Кафедра пропедевтики внутренних болезней", "Кафедра поликлинической терапии", "Кафедра периодонтологии  disk Q",
            "Кафедра фармацевтической  химии disk Q", "Кафедра фармацевтических технологий и химий", "Кафедра фармакологии", "Кафедра урологии", "Кафедра травматологии и ортопедии", "Кафедра судебной медицины", "Кафедра стомотологической хирургии", "Кафедра стоматологии детского возраста  Q", "Кафедра радиоционной медицины и экологии  disk Q",
            "Кафедра эндокринологии", "Кафедра эндодонтии", "Кафедра челюстно-лицевой хирургии", "Кафедра хирургической стоматологии", "Кафедра хирургических болезней", "Кафедра хирургии и трансплантологии", "Кафедра фтизиопульмонологии", "Кафедра философии и политологии  disk Q", "Кафедра физического воспитания и спорта1",
            "Фотоархив  full control", "доступ к СКУД video", "Машины компьютерных классов", "Лаборатория практических навыков", "Кафедра эпидемиологии",
            "Фотоархив для просмотра"
        };

            checkedListBox1.Items.AddRange(groups);
        }
        private void InitializeDirectoryContext()
        {
            try
            {
                string domainName = "bsmu.by";
                string username = "AdAdd";
                string password = "!QAZ2wsx";

                _context = new PrincipalContext(ContextType.Domain, domainName, username, password);

            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при подключении к Active Directory: " + ex.Message);
            }
        }
   /*     public static string Transliterate(string input)
        {
            var transliterationMap = new Dictionary<char, string>
    {
        {'а', "a"}, {'б', "b"}, {'в', "v"}, {'г', "g"},
        {'д', "d"}, {'е', "e"}, {'ё', "yo"}, {'ж', "zh"},
        {'з', "z"}, {'и', "i"}, {'й', "y"}, {'к', "k"},
        {'л', "l"}, {'м', "m"}, {'н', "n"}, {'о', "o"},
        {'п', "p"}, {'р', "r"}, {'с', "s"}, {'т', "t"},
        {'у', "u"}, {'ф', "f"}, {'х', "kh"}, {'ц', "ts"},
        {'ч', "ch"}, {'ш', "sh"}, {'щ', "shch"}, {'ъ', ""},
        {'ы', "y"}, {'ь', ""}, {'э', "e"}, {'ю', "yu"}, {'я', "ya"},
        {'А', "A"}, {'Б', "B"}, {'В', "V"}, {'Г', "G"},
        {'Д', "D"}, {'Е', "E"}, {'Ё', "Yo"}, {'Ж', "Zh"},
        {'З', "Z"}, {'И', "I"}, {'Й', "Y"}, {'К', "K"},
        {'Л', "L"}, {'М', "M"}, {'Н', "N"}, {'О', "O"},
        {'П', "P"}, {'Р', "R"}, {'С', "S"}, {'Т', "T"},
        {'У', "U"}, {'Ф', "F"}, {'Х', "Kh"}, {'Ц', "Ts"},
        {'Ч', "Ch"}, {'Ш', "Sh"}, {'Щ', "Shch"}, {'Ъ', ""},
        {'Ы', "Y"}, {'Ь', ""}, {'Э', "E"}, {'Ю', "Yu"}, {'Я', "Ya"}
    };

            StringBuilder result = new StringBuilder();
            foreach (char c in input)
            {
                if (transliterationMap.ContainsKey(c))
                {
                    result.Append(transliterationMap[c]);
                }
                else
                {
                    result.Append(c); 
                }
            }
            return result.ToString();
        }*/

        private void AddUserToSelectedGroups(string userName)
        {
            try
            {
                using (var context = new PrincipalContext(ContextType.Domain, "bsmu.by"))
                {
                    foreach (object item in checkedListBox1.CheckedItems)
                    {
                        string groupName = item.ToString();

                        GroupPrincipal group = GroupPrincipal.FindByIdentity(context, groupName);

                        if (group == null)
                        {
                            MessageBox.Show($"Группа {groupName} не найдена. Пропуск добавления пользователя в эту группу.");
                            continue;
                        }

                        UserPrincipal user = UserPrincipal.FindByIdentity(context, userName);

                        if (user != null)
                        {
                            try
                            {
                                group.Members.Add(user);
                                group.Save();
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Ошибка при добавлении пользователя в группу " + groupName + ": " + ex.Message);
                            }
                        }
                        else
                        {
                            MessageBox.Show($"Пользователь {userName} не найден. Группа {groupName} пропущена.");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при добавлении пользователя в группы: " + ex.Message);
            }
        }


        private void AddUserToDefaultGroups(string userName)
        {
            string[] defaultGroups = { "my_bsmu_by", "support_bsmu_by", "_INET_USERS" };

            try
            {
                using (var context = new PrincipalContext(ContextType.Domain, "bsmu.by"))
                {
                    UserPrincipal user = UserPrincipal.FindByIdentity(context, userName);

                    if (user == null)
                    {
                        MessageBox.Show($"Пользователь {userName} не найден. Добавление в группы по умолчанию невозможно.");
                        return;
                    }

                    foreach (string groupName in defaultGroups)
                    {
                        GroupPrincipal group = GroupPrincipal.FindByIdentity(context, groupName);

                        if (group == null)
                        {
                            MessageBox.Show($"Группа {groupName} не найдена. Пропуск добавления пользователя в эту группу.");
                            continue;
                        }

                        try
                        {
                            group.Members.Add(user);
                            group.Save();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Ошибка при добавлении пользователя в группу по умолчанию " + groupName + ": " + ex.Message);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при добавлении пользователя в группы по умолчанию: " + ex.Message);
            }
        }

        private bool GroupExists(string groupName)
        {
            try
            {
                using (DirectoryEntry groupEntry = new DirectoryEntry($"LDAP://CN={groupName},OU=Groups,OU=University,DC=bsmu,DC=by"))
                {
                    return groupEntry.NativeObject != null;
                }
            }
            catch (DirectoryServicesCOMException)
            {
                return false; 
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                using (DirectoryEntry usersOU = new DirectoryEntry("LDAP://OU=Users,OU=University,DC=bsmu,DC=by"))
                {
                    if (usersOU != null)
                    {
                        string userName = $"{textBox3.Text} {textBox5.Text} {textBox4.Text}";
                        using (DirectoryEntry newUser = usersOU.Children.Add($"CN={userName}", "user"))
                        {
                            newUser.Properties["sn"].Value = textBox3.Text;
                            newUser.Properties["givenName"].Value = textBox5.Text;
                            newUser.Properties["initials"].Value = textBox4.Text.Substring(0, 1).ToUpper();
                            newUser.Properties["displayName"].Value = $"{textBox3.Text} {textBox5.Text} {textBox4.Text}";

                            //string lastNameLatin = Transliterate(textBox3.Text);
                            //string firstNameInitial = Transliterate(textBox5.Text.Substring(0, 1).ToUpper());
                            //string patronymicInitial = Transliterate(textBox4.Text.Substring(0, 1).ToUpper());

                            string samAccountName = $"{textBox1.Text}";
                            newUser.Properties["samAccountName"].Value = textBox1.Text;
                            newUser.Properties["userPrincipalName"].Value = $"{textBox1.Text}@bsmu.by";
                            newUser.Properties["description"].Value = textBox7.Text;
                            newUser.Properties["title"].Value = textBox8.Text;
                            newUser.Properties["telephoneNumber"].Value = textBox9.Text;
                            newUser.Properties["employeeID"].Value = textBox11.Text;

                            newUser.CommitChanges();

                            newUser.Invoke("SetPassword", "!QAZ2wsx");

                            newUser.Properties["userAccountControl"].Value = 512;
                            newUser.Properties["department"].Value = textBox6.Text;
                            newUser.Properties["pwdLastSet"].Value = 0;
                            newUser.CommitChanges();

                            AddUserToDefaultGroups(samAccountName);
                            AddUserToSelectedGroups(samAccountName);

                            MessageBox.Show("Пользователь успешно создан");
                        }
                    }
                    else
                    {
                        MessageBox.Show("OU Users не найдено.");
                    }
                }
            }
            catch (DirectoryServicesCOMException comEx)
            {
                MessageBox.Show("Ошибка при взаимодействии с AD: " + comEx.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message + "\n" + ex.InnerException?.Message);
            }
        }

        private void YourForm_FormClosing(object sender, FormClosingEventArgs e)
        { 
            if (usersOU != null)
            {
                usersOU.Dispose();
                usersOU = null; 
            }
        }
        private void label3_Click(object sender, EventArgs e) { }
        public void textBox3_TextChanged(object sender, EventArgs e) { }
        private void textBox5_TextChanged(object sender, EventArgs e) { }
        private void textBox4_TextChanged(object sender, EventArgs e) { }
        private void label2_Click(object sender, EventArgs e) { }
        private void textBox6_TextChanged(object sender, EventArgs e) { }
        private void textBox7_TextChanged(object sender, EventArgs e){ }
        private void label4_Click(object sender, EventArgs e){ }
        private void textBox9_TextChanged(object sender, EventArgs e){ }
        private void textBox8_TextChanged(object sender, EventArgs e){ }
        private void label7_Click(object sender, EventArgs e){ }
        private void textBox11_TextChanged(object sender, EventArgs e){ }
        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e){ }
        private void textBox2_TextChanged(object sender, EventArgs e){ }
        private void textBox1_TextChanged(object sender, EventArgs e) { }
        private void textBox1_TextChanged_1(object sender, EventArgs e){ }
    }
}
