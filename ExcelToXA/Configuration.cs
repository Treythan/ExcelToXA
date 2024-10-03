using Newtonsoft.Json;

namespace ExcelToXA
{
    public partial class Configuration : Form
    {
        public Configuration()
        {
            InitializeComponent();
            environmentBox.Items.Add(new Environment { DisplayText = "VV - XA9.1 American Emergency Vehicles", Index = 0 });
            environmentBox.Items.Add(new Environment { DisplayText = "77 - TEST ENVIRONMENT FOR VV", Index = 1 });

            vendorBox.Items.Add(new Vendor { DisplayText = "Kinequip", Index = 0, PNColumn = 2, QTYColumn = 3, HasHeaders = true });
            user.CharacterCasing = CharacterCasing.Upper;

            if (File.Exists(System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData) + @"\usersettings.json"))
            {
                UserSettings settings = JsonConvert.DeserializeObject<UserSettings>(
                    File.ReadAllText(System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData) + @"\usersettings.json"));
                user.Text = settings.Username;
                pass.Text = settings.Password;
                environmentBox.SelectedIndex = settings.Environment.Index;
                vendorBox.SelectedIndex = settings.Vendor.Index;
                checkBox1.Checked = settings.Remember;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Contact Me:\nTrey Noblett\ntreythan.noblett@revgroup.com\n(336) 846-8158\nTeams Message", "Help!", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void okBtn_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(user.Text) || String.IsNullOrEmpty(pass.Text))
            {
                MessageBox.Show("You must enter your XA Username and Password inside of the configuration before this tool can be used.",
                    "Error Logging In", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (environmentBox.SelectedItem == null)
            {
                MessageBox.Show("You must select the Environment to be used inside of the configuration before this tool can be used.",
                    "Error Logging In", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (vendorBox.SelectedItem == null)
            {
                MessageBox.Show("You must select the Vendor to be used inside of the configuration before this tool can be used.",
                    "Error Logging In", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            Program.xaLogin = new KeyValuePair<string, string>(user.Text, pass.Text);
            UserSettings userSettings = new UserSettings
            {
                Username = user.Text,
                Password = pass.Text,
                Vendor = vendorBox.SelectedItem as Vendor,
                Environment = environmentBox.SelectedItem as Environment,
                Remember = checkBox1.Checked
            };

            Program.main.userSettings = userSettings;

            if (checkBox1.Checked)
            {
                File.WriteAllText(System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData) + @"\usersettings.json",
                    JsonConvert.SerializeObject(userSettings, Formatting.Indented));
            }

            this.Close();
            Program.main.GetOrders();
        }
    }
}
