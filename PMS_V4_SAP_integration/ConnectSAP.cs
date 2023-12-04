using SAPbobsCOM;
using Dapper;
using System.Text.Json;
using System.Configuration;
using Microsoft.Extensions.Configuration;
using System.Windows.Forms;
using PMS_V4_SAP_integration.Helper;
using System;

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
            if (!string.IsNullOrEmpty(_configuration.GetSection("SAPBusinessOneConfig:Server").Value))
                oCompany.Server = _configuration.GetSection("SAPBusinessOneConfig:Server").Value;
            if (!string.IsNullOrEmpty(_configuration.GetSection("SAPBusinessOneConfig:CompanyDB").Value))
                oCompany.CompanyDB = _configuration.GetSection("SAPBusinessOneConfig:CompanyDB").Value; //need to set this to the PRD company DB 
            if (!string.IsNullOrEmpty(_configuration.GetSection("SAPBusinessOneConfig:UserName").Value))
                oCompany.UserName = _configuration.GetSection("SAPBusinessOneConfig:UserName").Value; //get the PRD db username and password
            if (!string.IsNullOrEmpty(_configuration.GetSection("SAPBusinessOneConfig:Password").Value))
                oCompany.Password = _configuration.GetSection("SAPBusinessOneConfig:Password").Value;
            if (!string.IsNullOrEmpty(_configuration.GetSection("SAPBusinessOneConfig:DbServerType").Value))
                oCompany.DbServerType = (BoDataServerTypes)Enum.Parse(typeof(BoDataServerTypes), _configuration.GetSection("SAPBusinessOneConfig:DbServerType").Value, true);
            if (!string.IsNullOrEmpty(_configuration.GetSection("SAPBusinessOneConfig:LicenseServer").Value))
                oCompany.LicenseServer = _configuration.GetSection("SAPBusinessOneConfig:LicenseServer").Value;
            if (!string.IsNullOrEmpty(_configuration.GetSection("SAPBusinessOneConfig:SLDServer").Value))
                oCompany.SLDServer = _configuration.GetSection("SAPBusinessOneConfig:SLDServer").Value;
            if (!string.IsNullOrEmpty(_configuration.GetSection("SAPBusinessOneConfig:DbUserName").Value))
                oCompany.DbUserName = _configuration.GetSection("SAPBusinessOneConfig:DbUserName").Value;
            if (!string.IsNullOrEmpty(_configuration.GetSection("SAPBusinessOneConfig:DbPassword").Value))
                oCompany.DbPassword = _configuration.GetSection("SAPBusinessOneConfig:DbPassword").Value;

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
