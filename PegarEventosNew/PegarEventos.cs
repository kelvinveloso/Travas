using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Windows.Forms;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Data.Odbc;
using System.Data.Common;
using System.Configuration;
using System.Data.SqlClient;
using System.Threading;


using SAPbobsCOM;
using System.Net.Mail;
using System.Net;
//using SendGrid.Helpers.Mail;
//using SendGrid;

namespace PegarEventosNew
{
    public class PegarEventos
    {


        public SAPbouiCOM.Company oCompany;
        //public SAPbouiCOM.Application oApplicationGlobal;
        public SAPbouiCOM.ProgressBar oProgBar;
        public SAPbouiCOM.Form oForm;

        public SAPbobsCOM.Company ooCompany;


        public string vBP_4 = "";

        public string vBPName_54 = "";

        public string vDocNum_8 = "";

        public string vTotal_29 = "";

        public string vEmail = "";

        public string vObs_16 = "";



        private SAPbouiCOM.Application oApplication;

        public void SetApplicationUi()
        {

            
            try
            {


                SAPbouiCOM.SboGuiApi oSboGuiApi = null;

                string sConnectionString = null;

                oSboGuiApi = new SAPbouiCOM.SboGuiApi();

                sConnectionString = System.Convert.ToString(Environment.GetCommandLineArgs().GetValue(1));

                oSboGuiApi.Connect(sConnectionString);

                UiApplication = oSboGuiApi.GetApplication(-1);
                //oApplicationGlobal = oSboGuiApi.GetApplication(-1);

                oApplication = oSboGuiApi.GetApplication(-1);

                DateTime current = DateTime.Now;






                UiApplication.StatusBar.SetText("Addon " + System.Windows.Forms.Application.ProductName + " UI Conectada...         " + current + " . ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                //oApplicationGlobal.StatusBar.SetText("Addon " + Application.ProductName + " UI Conectada...         " + current + " . ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);



            }
            catch (Exception ex)
            {

                MessageBox.Show("Erro conexao UI :  Execute o SAP como ADM... " + ex);
                Environment.Exit(0);

            }
        }

        private void SetApplication()
        {
            SAPbouiCOM.SboGuiApi oSboGuiApi = null;
            string sConnectionString = null;
            oSboGuiApi = new SAPbouiCOM.SboGuiApi();
            sConnectionString = System.Convert.ToString(Environment.GetCommandLineArgs().GetValue(1));
            oSboGuiApi.Connect(sConnectionString);
            oApplication = oSboGuiApi.GetApplication(-1);



        }


        static public SAPbobsCOM.Company DiCompany;

        static public SAPbouiCOM.Application UiApplication;
        

        public void SetApplicationDi()
        {


            try
            {
                int lErrCode;
                string sErrMsg = "", sCookie = "", sConnectionContext = "";
                DiCompany = new SAPbobsCOM.Company();

                //diCompany = (SAPbobsCOM.Company)uiApplication.Company.GetDICompany();
                sCookie = DiCompany.GetContextCookie();
                sConnectionContext = UiApplication.Company.GetConnectionContext(sCookie);

                if (DiCompany.Connected == true)
                {
                    DiCompany.Disconnect();

                }
                else
                {
                    DiCompany.SetSboLoginContext(sConnectionContext);
                    DiCompany.Connect();
                }
                DiCompany.GetLastError(out lErrCode, out sErrMsg);
                if (lErrCode != 0)
                {
                    throw new Exception(sErrMsg);

                }

                DateTime current = DateTime.Now;





                UiApplication.StatusBar.SetText("Addon " + System.Windows.Forms.Application.ProductName + " DI Conectada...     " + current + " . ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);

            }
            catch (Exception ex)
            {

                UiApplication.StatusBar.SetText(" Erro na conexao DI " + ex, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);

                Environment.Exit(0);

            }
        }


        public void CriaInfra()
        {
            //instancia variaveis globais
        

            //Cria Tabela Config
            string TabelaStr = System.Windows.Forms.Application.ProductName + "Config";

            TabelaStr = TabelaStr.Substring(0,10);

            if (!ExisteTB(TabelaStr))
            {

                UiApplication.StatusBar.SetText("Criando Tabela " + TabelaStr + " ...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                AddUserTable(TabelaStr, System.Windows.Forms.Application.ProductName + "Config", BoUTBTableType.bott_NoObject);
                UiApplication.StatusBar.SetText("Tabela " + TabelaStr + " criada ...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);


            }
            else
            {


                    UiApplication.StatusBar.SetText("Tabela " + System.Windows.Forms.Application.ProductName + " ja existe...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);


                
            }



            AddUserField(DiCompany.CompanyDB, "@" + TabelaStr, "SmtpClient", "SmtpClient", SAPbobsCOM.BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 254, null, null, null);

            AddUserField(DiCompany.CompanyDB, "@" + TabelaStr, "SmtpPorta", "SmtpPorta", SAPbobsCOM.BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 254, null, null, null);

            AddUserField(DiCompany.CompanyDB, "@" + TabelaStr, "LoginEmail", "LoginEmail", SAPbobsCOM.BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 254, null, null, null);

            AddUserField(DiCompany.CompanyDB, "@" + TabelaStr, "SenhaEmail", "SenhaEmail", SAPbobsCOM.BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 254, null, null, null);

            AddUserField(DiCompany.CompanyDB, "@" + TabelaStr, "EnableSSL", "EnableSSL", SAPbobsCOM.BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 254, null, null, null);

            AddUserField(DiCompany.CompanyDB, "@" + TabelaStr, "Remetente", "Remetente", SAPbobsCOM.BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 254, null, null, null);

            AddUserField(DiCompany.CompanyDB, "@" + TabelaStr, "Destinatario", "Destinatario", SAPbobsCOM.BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 254, null, null, null);

            AddUserField(DiCompany.CompanyDB, "@" + TabelaStr, "Assunto", "Assunto", SAPbobsCOM.BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 254, null, null, null);

            AddUserField(DiCompany.CompanyDB, "@" + TabelaStr, "Corpo", "Corpo", SAPbobsCOM.BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 254, null, null, null);

            AddUserField(DiCompany.CompanyDB, "@" + TabelaStr, "EmailFinanceiro", "EmailFinanceiro", SAPbobsCOM.BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 254, null, null, null);

            AddUserField(DiCompany.CompanyDB, "@" + TabelaStr, "LinkImagem", "LinkImagem", SAPbobsCOM.BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 254, null, null, null);


            //--==popula campos

            try
            {



                string Code = vProductName;
                Code = Code.ToUpper();
                string Code2 = "@" + Code;



                string sSquery = $@"select count(1) from ""{Code2}"" where ""Code"" = '{Code}'";



                int oResult = Convert.ToInt32(ExecuteSqlScalar(sSquery));



                if (oResult == 0)
                {

            

                        UiApplication.StatusBar.SetText("Inserindo Configurações...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);


                

                    var udt = DiCompany.UserTables.Item(Code);

                    udt.Code = Code;
                    udt.Name = Code;
                    udt.UserFields.Fields.Item("U_SmtpClient").Value = "smtp.gmail.com";

                    udt.UserFields.Fields.Item("U_SmtpPorta").Value = "587";
                    udt.UserFields.Fields.Item("U_LoginEmail").Value = "kelvin.veloso@gmail.com";
                    udt.UserFields.Fields.Item("U_SenhaEmail").Value = "jtermzlhcenpebyz";
                    udt.UserFields.Fields.Item("U_EnableSSL").Value = "Y";
                    udt.UserFields.Fields.Item("U_Remetente").Value = "kelvin.veloso@gmail.com";
                    udt.UserFields.Fields.Item("U_Destinatario").Value = "kelvin.veloso@gmail.com";
                    udt.UserFields.Fields.Item("U_Assunto").Value = "SAP BUSINESS ONE - Pagamento Aprovado";
                    udt.UserFields.Fields.Item("U_Corpo").Value = "Olá, o pagamento do Bem ou Serviço fornecido foi Aprovado !";
                    udt.UserFields.Fields.Item("U_EmailFinanceiro").Value = "kelvin.veloso@gmail.com";
                    udt.UserFields.Fields.Item("U_LinkImagem").Value = "https://upload.wikimedia.org/wikipedia/commons/thumb/5/59/SAP_2011_logo.svg/1200px-SAP_2011_logo.svg.png";




                    int iRet = udt.Add();
     

                        UiApplication.StatusBar.SetText("Configurações Inseridas...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);

              

                    if (iRet != 0)
                        throw new Exception(DiCompany.GetLastErrorDescription());
                }
            }
            catch (Exception ex)
            {

                UiApplication.StatusBar.SetText("Erro: " + ex, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);


            }


            //--==popula campos




        }

        string vProductName = System.Windows.Forms.Application.ProductName;
        string vTabela = System.Windows.Forms.Application.ProductName + "Config";

        public void EnvioEmailSimples(string pDestinatario, string pRemetente, string pAssunto, string pCorpo, string pAlias = "", bool pHtml = false)
        {

            try
            {

                string DataBase = UiApplication.Company.DatabaseName.ToString();

            



            string Code = vProductName;

            Code = Code.ToUpper();

            var udt = DiCompany.UserTables.Item(Code);

            udt.GetByKey(vProductName.ToUpper());



            string vSmtpClient = udt.UserFields.Fields.Item("U_SmtpClient").Value;
            string vSmtpPorta = udt.UserFields.Fields.Item("U_SmtpPorta").Value;

            string vLoginEmail = udt.UserFields.Fields.Item("U_LoginEmail").Value;


            string vSenhaEmail = udt.UserFields.Fields.Item("U_SenhaEmail").Value;

            string vEnableSSL = udt.UserFields.Fields.Item("U_EnableSSL").Value;


           

                var fromAddress = new MailAddress(pRemetente, pAlias);
                var toAddress = new MailAddress(pDestinatario, pDestinatario);
                string fromPassword = vSenhaEmail;
                string subject = pAssunto;
                string body = pCorpo;

                var smtp = new SmtpClient
                {
                    Host = vSmtpClient,
                    Port = Convert.ToInt32(vSmtpPorta),
                    EnableSsl = true,
                    UseDefaultCredentials = true,
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    Credentials = new NetworkCredential(fromAddress.Address, fromPassword),
                    Timeout = 20000

                };
                using (var message = new MailMessage(fromAddress, toAddress)
                {
                    Subject = subject
                    ,
                    Body = body
                    ,
                    IsBodyHtml = pHtml

                    //jtermzlhcenpebyz

                })
                {
                    smtp.Send(message);
                }
                UiApplication.SetStatusBarMessage("Email enviado para : " + pDestinatario, SAPbouiCOM.BoMessageTime.bmt_Short, false);
            }
            catch (Exception ex)
            {

                UiApplication.SetStatusBarMessage("Erro: " + ex, SAPbouiCOM.BoMessageTime.bmt_Short, false);


            }


        }


        public PegarEventos()
        {
            SetApplicationUi();
            SetApplicationDi();

            CriaInfra();


            //App Events
            oApplication.AppEvent += OApplication_AppEvent;
            //MenuEvents
            oApplication.MenuEvent += OApplication_MenuEvent;
            //Item Events
            oApplication.ItemEvent += OApplication_ItemEvent;

            oApplication.SetStatusBarMessage(string.Format("Addon "+vProductName+" Inicializado ..."), BoMessageTime.bmt_Short, false);

        }

        private void OApplication_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {

            

            if (pVal.FormType !=0)
            {

                    //------Bloco Especial para analise de Eventos

                    oApplication.SetStatusBarMessage(string.Format(

                       " BeforeAction : " + pVal.BeforeAction.ToString() //== "False"
                     + " ActionSuccess : " + pVal.ActionSuccess.ToString() //== "True"
                     + " ItemUID : " + pVal.ItemUID.ToString() //== "1"
                     + " FormTypeEx : " + pVal.FormTypeEx.ToString() //== "143"
                     + " EventType : " + pVal.EventType.ToString() //== "et_ITEM_PRESSED"
                     + " FormMode : " + pVal.FormMode.ToString() //== "3"


                        ), BoMessageTime.bmt_Short, false);

                    //Valida PN
                    if (

                        pVal.BeforeAction.ToString() == "True"
                        && pVal.ItemUID.ToString() == "1"
                        && pVal.FormTypeEx.ToString() == "134"
                        && pVal.EventType.ToString() == "et_ITEM_PRESSED"


                        //   pVal.BeforeAction.ToString() =="False"
                        //&& pVal.ActionSuccess.ToString() == "True"
                        //&& pVal.ItemUID.ToString() == "1"
                        //&& pVal.FormTypeEx.ToString() == "134"
                        //&& pVal.EventType.ToString() == "et_ITEM_PRESSED"
                        //&& pVal.FormMode.ToString() =="3"
                        )
                    {
                        SAPbouiCOM.Form UDFForm = oApplication.Forms.Item(pVal.FormUID);



                        string OCRD_CardName_7 = UDFForm.Items.Item("7").Specific.Value;

                        if (OCRD_CardName_7 == null || OCRD_CardName_7 == "")
                        {
                            oApplication.SetStatusBarMessage(string.Format("Para Adicionar um PN é obrigatório o Razão Social!"), BoMessageTime.bmt_Short, true);
                            BubbleEvent = false;
                        }

                        string OCRD_CardFName_2014 = UDFForm.Items.Item("2014").Specific.Value;

                        if (OCRD_CardFName_2014 == null || OCRD_CardFName_2014 == "")
                        {
                            oApplication.SetStatusBarMessage(string.Format("Para Adicionar um PN é obrigatório o Fantasia!"), BoMessageTime.bmt_Short, true);
                            BubbleEvent = false;
                        }


                    }


                        //If pVal.FormType = "139" And pVal.ItemUID = "1" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.BeforeAction = True Then
                        //Try
                        //Dim oedit As SAPbouiCOM.EditText
                        //oform = sbo_application.Forms.GetFormByTypeAndCount(pVal.FormType * -1, pVal.FormTypeCount)
                        //oedit = oform.Items.Item("U_PONo").Specific
                        //If oedit.Value = "" Then
                        //sbo_application.MessageBox("Cannot add while the user field is empty")
                        //BubbleEvent = False
                        //End If
                        //Catch ex As Exception
                        //sbo_application.MessageBox(ex.Message)
                        //End Try
                        //End If



                    BoEventTypes eventEnum = 0;
                eventEnum = pVal.EventType;

                #region CapturaVariaveis

                if (

                        //pVal.EventType.ToString() == "et_CLICK"


                        pVal.BeforeAction.ToString() == "True"
                        && pVal.ItemUID.ToString() == "1"
                        && pVal.FormTypeEx.ToString() == "143"
                        && pVal.EventType.ToString() == "et_ITEM_PRESSED"
                        



                    )
                {

                    //oApplication.SetStatusBarMessage(
                    //string.Format(
                    //"Evento: {0}, FormType: {1},FormId: {2}, Before: {3}, ItemUID: {4}   ",
                    //eventEnum.ToString(),
                    //pVal.FormType.ToString(),
                    //pVal.FormUID.ToString(),
                    //pVal.BeforeAction.ToString(),
                    //pVal.ItemUID.ToString()

                    //), BoMessageTime.bmt_Short, false);

                    SAPbouiCOM.Form UDFForm = oApplication.Forms.Item(pVal.FormUID);

                    vBP_4 = UDFForm.Items.Item("4").Specific.Value;
                    vDocNum_8 = UDFForm.Items.Item("8").Specific.Value;
                    vTotal_29 = UDFForm.Items.Item("29").Specific.Value;
                    vBPName_54 = UDFForm.Items.Item("54").Specific.Value;
                    vObs_16 = UDFForm.Items.Item("16").Specific.Value;

                        BusinessPartners bp = (BusinessPartners)DiCompany.GetBusinessObject(BoObjectTypes.oBusinessPartners);

                    bp.GetByKey(vBP_4);

                    vEmail =  bp.EmailAddress.ToString();


//                    oApplication.SetStatusBarMessage(string.Format(
//                        "Variaveis Capturadas "+" Email: "+ vEmail +" PN: " + vBP_4+ " - " + vBPName_54 + " Docto: " + vDocNum_8 + " Total "+vTotal_29
//                        ), BoMessageTime.bmt_Short, false);
                }
                else
                {
                    

                }
                #endregion


                #region Email


                if (
                    //regra
                    
                    
                    
                    pVal.BeforeAction.ToString()=="False" 
                    && pVal.ActionSuccess.ToString() == "True" 
                    && pVal.ItemUID.ToString() == "1"
                    && pVal.FormTypeEx.ToString() =="143"
                    && pVal.EventType.ToString() == "et_ITEM_PRESSED"
                    && ( pVal.FormMode.ToString() == "3" || pVal.FormMode.ToString() == "1")

                    )
                {
                    //faça



                    //var oForm = oApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);
                    //string cardcode = oForm.Items.Item("CardCode").Specific.Value;

                   // oApplication.SetStatusBarMessage(string.Format("Inicia Botao"), BoMessageTime.bmt_Short, false);

                        //string udfFormUID = oApplication.Forms.Item(FormUID).UDFFormUID;

                        //Form UDFForm = oApplication.Forms.Item(udfFormUID);

                        //((EditText)UDFForm.Items.Item("U_EndPedido").Specific).Value = "YOURVALUE";

                        //oApplication.SetStatusBarMessage(
                        //(
                        //  " BeforeAction: " + pVal.BeforeAction.ToString()
                        //+ " ActionSuccess: " + pVal.ActionSuccess.ToString()
                        //+ " ItemUID: " + pVal.ItemUID.ToString()
                        //+ " FormTypeEx: " + pVal.FormTypeEx.ToString()
                        //+ " EventType: " + pVal.EventType.ToString()
                        //+ " FormMode: " + pVal.FormMode.ToString()


                        //), BoMessageTime.bmt_Short, false);

                        try
                        {


                    string Code = vProductName;

                    Code = Code.ToUpper();

                    var udt = DiCompany.UserTables.Item(Code);

                    udt.GetByKey(vProductName.ToUpper());

                    if (vEmail == null || vEmail.Count() == 0)
                    {
                                oApplication.SetStatusBarMessage(string.Format("Fornecedor sem Email Cadastrado !"), BoMessageTime.bmt_Short, false);

                                goto Found;
                    }



                    string Mode = "";

                    if (pVal.FormMode.ToString() == "3")
                    {
                                Mode = "Registrada";        
                    }
                    if (pVal.FormMode.ToString() == "1")
                    {
                        Mode = "Atualizada";
                    }
                             string vImagem = udt.UserFields.Fields.Item("U_LinkImagem").Value;
                            string html = "";
                            string vCorpoHtml = udt.UserFields.Fields.Item("U_Corpo").Value;

                            html = "<html> <body>  <p> <center>  " + vCorpoHtml+ ",</p>  <p> <center> Sr.: <b>" + vBPName_54 + "</b>,</p> <center> <table border=\"1\">    <tr>         <td><b>Código</b></td>        <td><b>Nome do Fornecedor</b></td>        <td><b>Autorização Nº</b></td> <td><b>Valor Total</b></td>		 <td><b>Observações</b></td>    </tr>    <tr>        <td>" + vBP_4+ "</td>        <td>" + vBPName_54 + "</td>        <td>" + vDocNum_8 + "</td>		 <td>" + vTotal_29 + "</td>		 <td>" + vObs_16 + "</td>    </tr> </table> <br> <br> <center> <div> <img src = " + vImagem+ " alt=\"Image\" height=\"50\" width=\"120\"> </div> </body> </html>";



                            string vDestinatario = vEmail;
                    string vRemetente = udt.UserFields.Fields.Item("U_Remetente").Value;
                    string vAssunto= "Autorização de Pagto: " + vDocNum_8 + " " + Mode + " - " + udt.UserFields.Fields.Item("U_Assunto").Value ;
                            string vCorpo = html;
                                // udt.UserFields.Fields.Item("U_Corpo").Value+ " <br> "
                                //+html
                     //       + "Autorização de Pagto: " + vDocNum_8 + " " + Mode + "  " + " <br> "
                     
                     //+ "Dados: PN: " + vBP_4 + " <br> "
                     //+ "  " + vBPName_54 
                     //+ " Docto: " + vDocNum_8 + " <br> "
                     //+ " Total " + vTotal_29 + " <br> "
                     //+ " Obs. "+vObs_16 + " <br> "
                     //+ " . "
                    ;
                    string vEmailFinanceiro = udt.UserFields.Fields.Item("U_EmailFinanceiro").Value;
                    string vAlias = "Email Enviado por Sap Business One -  Addon EvokeMail";
                    bool vHtml = true;
                            #region Email Original


                            EnvioEmailSimples(vDestinatario, vRemetente, vAssunto, vCorpo, vAlias, vHtml);
                            EnvioEmailSimples(vEmailFinanceiro, vRemetente, vAssunto, vCorpo, vAlias, vHtml);
                            #endregion



                           // oApplication.SetStatusBarMessage(string.Format("Email Enviado"), BoMessageTime.bmt_Short, false);

                        }
                        catch (Exception ex)
                        {

                            oApplication.SetStatusBarMessage(string.Format("Erro"+ex), BoMessageTime.bmt_Short, false);

                        }


                    //oApplication.SetStatusBarMessage(string.Format
                    //(
                    //" Email enviado para "+vEmail+" para o PN: "+vBP_4+"-"+vBPName_54+" sobre o Documento "+vDocNum_8+" que foi Aprovado " +" no Valor total de "+vTotal_29
                    //), BoMessageTime.bmt_Short, false);



                    //oApplication.SetStatusBarMessage(string.Format("Finaliza Botao"), BoMessageTime.bmt_Short, false);
                    Found:

                        string vTeste = "";




                    }
                    else 
                {
                    
                        
                        // senão


                        //oApplication.SetStatusBarMessage(
                        //string.Format(
                        //"Evento: {0}, FormType: {1},FormId: {2}, Before: {3}, ItemUID: {4}   ",
                        //eventEnum.ToString(),
                        //pVal.FormType.ToString(),
                        //pVal.FormUID.ToString(),
                        //pVal.BeforeAction.ToString(),
                        //pVal.ItemUID.ToString()

                        //), BoMessageTime.bmt_Short, false);
                    }
                    #endregion


                }

            }
            catch (Exception ex)
            {

                oApplication.SetStatusBarMessage(string.Format("erro"+ex), BoMessageTime.bmt_Short, false);

            }

        }

        private void OApplication_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            if (pVal.BeforeAction)
            {
                //oApplication.SetStatusBarMessage(" Menu Item: " + pVal.MenuUID + " Antes ", SAPbouiCOM.BoMessageTime.bmt_Short, false);
            }
            else
            {
               // oApplication.SetStatusBarMessage(" Menu Item: " + pVal.MenuUID + " Depois ", SAPbouiCOM.BoMessageTime.bmt_Short, false);
            }



            //throw new NotImplementedException();
        }

        private void OApplication_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            //throw new NotImplementedException();

            switch (EventType)
            {

                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                oApplication.MessageBox("A empresa foi trocada");
                break;

                case SAPbouiCOM.BoAppEventTypes.aet_FontChanged:
                oApplication.MessageBox("A fonte foi Selecionada");
                break;

                case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                oApplication.MessageBox("A Liguagem foi trocada");
                break;

                case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    System.Windows.Forms.Application.Exit();
                    //oApplication.MessageBox("A empresa foi encerrada");

                    break;

                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    System.Windows.Forms.Application.Exit();
                    //oApplication.MessageBox("o evento ShutDown foi Chamado"
                    //+Environment.NewLine
                    //+"Fechando o Addon", 1);

                break;


            }

        }


        public bool ExisteTB(string TBName)
        {


            SAPbobsCOM.UserTablesMD oUserTable;
            oUserTable = (SAPbobsCOM.UserTablesMD)DiCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);
            bool ret = oUserTable.GetByKey(TBName);
            int errCode; string errMsg;
            DiCompany.GetLastError(out errCode, out errMsg);

            TBName = null;
            errMsg = null;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTable);
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            return (ret);
        }

        public void AddUserTable(string NomeTB, string Desc, SAPbobsCOM.BoUTBTableType oTableType)
        {
            NomeTB = NomeTB.Substring(0, 10);

            int lErrCode;
            UserTablesMD oUserTable;

            oUserTable = (UserTablesMD)DiCompany.GetBusinessObject(BoObjectTypes.oUserTables);

            try
            {
                oUserTable.TableName = NomeTB.Replace("@", "").Replace("[", "").Replace("]", "").Trim();
                oUserTable.TableDescription = Desc;
                oUserTable.TableType = oTableType;
                try
                {
                    oUserTable.Add();
                    string sErrMsg;
                    DiCompany.GetLastError(out lErrCode, out sErrMsg);
                    if (lErrCode != 0)
                    {
                        throw new Exception(sErrMsg);
                    }

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTable);
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
                catch (Exception ex)
                {
                    UiApplication.SetStatusBarMessage(" Erro "+ex, BoMessageTime.bmt_Short, false);
                }

            }
            catch (Exception ex)
            {
                UiApplication.SetStatusBarMessage(" Erro " + ex, BoMessageTime.bmt_Short, false);
            }
        }

        public object ExecuteSqlScalar(string query)
        {
            try
            {
                object obj = null;
                var oRs = (SAPbobsCOM.Recordset)DiCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRs.DoQuery(query);
                if (!oRs.EoF)
                {
                    obj = oRs.Fields.Item(0).Value;
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs);


                return obj;
            }
            catch (Exception ex)
            {
                UiApplication.SetStatusBarMessage("Erro: ", BoMessageTime.bmt_Short, false);
                throw;
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
        }

        public void AddUserField(String Banco, string NomeTabela, string NomeCampo, string DescCampo, SAPbobsCOM.BoFieldTypes Tipo, SAPbobsCOM.BoFldSubTypes SubTipo, Int16 Tamanho, string[,] valoresValidos, string valorDefault, string linkedTable, int linkedSystemObject = 0)
        {

            

            

            int lErrCode;
            string sErrMsg = "";

            NomeTabela = NomeTabela.ToUpper();

            NomeTabela = NomeTabela.Substring(0, 11);


            try
            {

                string sSquery = $@"Select ""FieldID"" From ""{Banco}"".""CUFD"" Where ""TableID"" = '{NomeTabela}' and ""AliasID"" = '{NomeCampo}'";

                object oResult = ExecuteSqlScalar(sSquery);

                if (oResult != null)


                if (oResult != null)

                    return;

                SAPbobsCOM.UserFieldsMD oUserField;
                oUserField = (SAPbobsCOM.UserFieldsMD)DiCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                oUserField.TableName = NomeTabela.Replace("@", "").Replace("[", "").Replace("]", "").Trim();
                oUserField.Name = NomeCampo;
                oUserField.Description = DescCampo;
                oUserField.Type = Tipo;
                oUserField.SubType = SubTipo;
                oUserField.DefaultValue = valorDefault;
                if (!string.IsNullOrEmpty(linkedTable)) oUserField.LinkedTable = linkedTable;
                if (linkedSystemObject != 0) oUserField.LinkedSystemObject = (UDFLinkedSystemObjectTypesEnum)linkedSystemObject;

                //adicionar valores válidos
                if (valoresValidos != null)
                {
                    Int32 qtd = valoresValidos.GetLength(0);
                    if (qtd > 0)
                    {
                        for (int i = 0; i < qtd; i++)
                        {
                            oUserField.ValidValues.Value = valoresValidos[i, 0];
                            oUserField.ValidValues.Description = valoresValidos[i, 1];
                            oUserField.ValidValues.Add();
                        }
                    }
                }

                if (Tamanho != 0)
                    oUserField.EditSize = Tamanho;

                try
                {
                    oUserField.Add();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserField);
                    DiCompany.GetLastError(out lErrCode, out sErrMsg);
                    if (lErrCode != 0)
                    {
                        throw new Exception(sErrMsg);
                    }
                    oUserField = null;


                    UiApplication.SetStatusBarMessage($@"Campo '{NomeCampo}' criado com sucesso...", SAPbouiCOM.BoMessageTime.bmt_Short, false);

                }
                catch (Exception ex)
                {

                 UiApplication.SetStatusBarMessage($@"Campo '{NomeCampo}', erro ..." + ex, SAPbouiCOM.BoMessageTime.bmt_Short, false);

                }
                oUserField = null;
            }
            catch (Exception ex)
            {

                    UiApplication.SetStatusBarMessage($@"Campo '{NomeCampo}', erro ..." + ex, SAPbouiCOM.BoMessageTime.bmt_Short, false);

                
            }
        }




        //-------------- fim

    }
}
