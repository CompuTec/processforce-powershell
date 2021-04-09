using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PFUIEventServerSample
{
    class BaseModel : INotifyPropertyChanged
    {
        public BaseModel()
        {
            Initialize();
        }

        #region Properties and Fields
        private CompuTec.Core.UI.PipeEvents.Server.EventServer server;

        //Pasword Required to send Messages Via NamedPipes server
        private const string Password = "UIEventServer";
        private SAPbouiCOM.Application SBO_Application;
        private SAPbobsCOM.Company oCompany;

        private bool _Connected;

        public event PropertyChangedEventHandler PropertyChanged;

        public bool Connected
        {
            get { return _Connected; }
            set
            {
                _Connected = value;
                OnValueChanged(nameof(Connected));
            }
        }
        private int _Spid;

        public ObservableCollection<CompuTec.Core.UI.PipeEvents.Message> Messages { get; set; }

        public int SAPPid
        {
            get { return _Spid; }
            set
            {
                _Spid = value;
                OnValueChanged(nameof(SAPPid));
            }
        }

        public BaseCommand Connect { get; private set; }

        #endregion
   
        #region private methods 
        private void Initialize()
        {
            Messages = new ObservableCollection<CompuTec.Core.UI.PipeEvents.Message>();
            Connect = new BaseCommand((x) =>
              {
                  return !Connected;
              }, (x) =>
             {
                  ConnectImpl();
              });
        }

        private void ConnectImpl()
        {
            try
            {
                //Connect to SAP UI 

                SAPbouiCOM.SboGuiApi SboGuiApi;
                string sConnectionString = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056";
                SboGuiApi = new SAPbouiCOM.SboGuiApi();
                // connect to a running SBO Application
                SboGuiApi.Connect(sConnectionString);
                // get an initialized application object
                SBO_Application = SboGuiApi.GetApplication(-1);
                oCompany = (SAPbobsCOM.Company)SBO_Application.Company.GetDICompany();
                Connected = true;
                //Get SAP Process Id as it is a part of the NamedPipeAddress
                SAPPid = GeSapPID();
                //Initialize NambedPipeServer to retrive the events
                InitializeEventServer();
            }catch(Exception ex)
            {
                System.Windows.MessageBox.Show($"ConnectionError:{ex.Message}");
            }

        }
     
        private int GeSapPID()
        {
            //this is how we are getting the sap process ID for both HANA and SQL instance
            
            SAPbobsCOM.Recordset recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string hanaQuery = @"select  Client_PID
from m_connections
where Connection_ID=Current_connection"; ;
            string sqlQuery=@"select host_process_id from sys.dm_exec_sessions where session_id=@@SPID"; ;
            recordset.DoQuery(oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB ? hanaQuery : sqlQuery);
            var pid = recordset.Fields.Item(0).Value;
            return pid;
        }

        private void InitializeEventServer()
        {
            //Initialize new Event Server
            server = new CompuTec.Core.UI.PipeEvents.Server.EventServer($"CompuTec.UI.EventServer_{SAPPid}");
            //Specify a Password for server
            server.ServerPassword = Password;
            //Specify number of possible instances this value indicates how many servers can use same Addres.
            //Each addon that is connected To Computec Event Sender consumes one instance please dont change
            //24 indicates that to the Single ProcessForce Application Can connect up to 24 other addons that receive those events
            server.MaxNumberOfServerInstances= 24;
            ///Assigne a handler to Event 
            server.MessageRecived += Server_MessageRecived;
            server.StartListening();
            
        }

        private void Server_MessageRecived(object obj, CompuTec.Core.UI.PipeEvents.Message e)
        {
            App.Current.Dispatcher.Invoke(() =>
            {
                Messages.Add(e);
            }            ); 
        }

        

        private void OnValueChanged(string name)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
        #endregion
    }
}
