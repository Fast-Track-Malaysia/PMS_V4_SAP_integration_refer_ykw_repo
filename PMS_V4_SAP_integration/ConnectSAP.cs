using SAPbobsCOM;
using Dapper;
using System.Text.Json;
using System.Configuration;
using Microsoft.Extensions.Configuration;
using System.Windows.Forms;
using PMS_V4_SAP_integration.Helper;

namespace PMS_V4_SAP_integration
{
    public class ConnectSAP
    {
        private int connectionResult;
        private int errorCode = 0;
        private string errorMessage = "";

        public SAPbobsCOM.Company oCompany = new SAPbobsCOM.Company(); //set public so that you can connect at other places



        private readonly IConfiguration _configuration;

        public ConnectSAP(IConfiguration configuration)
        {
            _configuration = configuration;

        }

        public int SAPConnect()
        {
            oCompany.Server = "ZHENXUAN\\CZX";
            oCompany.CompanyDB = "SBODemoUS";
            oCompany.UserName = "Manager";
            oCompany.Password = "1234";
            oCompany.DbServerType = BoDataServerTypes.dst_MSSQL2019;
            oCompany.LicenseServer = "localhost:30000";
            oCompany.DbUserName = "sa";
            oCompany.DbPassword = "sa";

             
            connectionResult = oCompany.Connect();

            if (connectionResult != 0)
            {
                oCompany.GetLastError(out errorCode, out errorMessage);
                return connectionResult;

            }

            
            return connectionResult;
        }

    }
}
