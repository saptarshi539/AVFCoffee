using CoffeeCore.Interfaces;
using FlexCel.Core;
using CoffeeCore.DTO;
using FlexCel.XlsAdapter;
using System;
using System.Collections.Generic;
using Microsoft.Extensions.Configuration;
using System.Data.SqlClient;
using Newtonsoft.Json;
using System.Net;
using System.Collections.Specialized;
using System.Net.Http;
using System.Threading.Tasks;

namespace CoffeeInfrastructure.Flexcel
{
    public class Flexcelsum : IFlexcelsum
    {
        private readonly IConfiguration _iconfiguration;

        public Flexcelsum(IConfiguration configuration)
        {
            _iconfiguration = configuration;
        }

        public ChartDataDTO getOutputFromExcel(double earlyHectares, double peakHectares, double oldHectares, bool conventional,
            bool organic, bool transition, double workerSalarySoles, double productionQuintales, double transportCostSoles, double costPriceSolesPerQuintal,
            double expSolesChem, double expSolesOrg)
        {
            //working in the develop branch
            //Excel sheet inputs from Juan
            Inputs inputs = new Inputs();
            AdvancedInputs advancedInputs = new AdvancedInputs();
            Budget_Equipo budget_Equipo = new Budget_Equipo();
            Budget_Establecimiento budget_Establecimiento = new Budget_Establecimiento();
            Budget_M_Obra budget_M_Obra = new Budget_M_Obra();
            Budget_Presupuesto budget_Presupuesto = new Budget_Presupuesto();
            Budget_Sostenemiento budget_Sostenemiento = new Budget_Sostenemiento();
            Budget_Valor_de_M_Obra budget_Valor_De_M_Obra = new Budget_Valor_de_M_Obra();
            Conversiones conversiones = new Conversiones();
            DatabaseSchema databaseSchema = new DatabaseSchema();
            InAdvanced advanced = new InAdvanced();
            Inputs_1_metric_currency inputs_1_Metric_Currency = new Inputs_1_metric_currency();
            Inputs_1_Ref inputs_1_Ref = new Inputs_1_Ref();
            OutcomeLAdjustment outcomeLAdjustment = new OutcomeLAdjustment();
            OutcomeTotalAdj outcomeTotalAdj = new OutcomeTotalAdj();
            OutcomeYAdjustment outcomeYAdjustment = new OutcomeYAdjustment();
            Output output = new Output();
            Output1_pre_metric_currency pre_Metric_Currency = new Output1_pre_metric_currency();
            Proportions proportions = new Proportions();
            Prporcion_de_productividad prporcion_De_Productividad = new Prporcion_de_productividad();

            XlsFile xls = new XlsFile(true);
            //xls.Open("file");
            TWorkspace workspace = new TWorkspace();
            workspace.Add(xls.ActiveFileName, xls);
            //actual calculation taking place in the excel sheet
            inputs.inputs(xls, earlyHectares, peakHectares, oldHectares, conventional, organic, transition, workerSalarySoles, productionQuintales, transportCostSoles, 
                costPriceSolesPerQuintal, expSolesChem, expSolesOrg);
            xls.Recalc();
            databaseSchema.Database_Schema(xls, workspace);
            xls.Recalc();
            conversiones.conversiones(xls);
            xls.Recalc();
            inputs_1_Ref.inputs1Ref(xls);
            xls.Recalc();
            inputs_1_Metric_Currency.Inputs1MetricCurrency(xls);
            xls.Recalc();
            advanced.Inputs_advanced(xls);
            xls.Recalc();
            proportions.proportions(xls);
            xls.Recalc();
            advancedInputs.Budget_Supuestos(xls);
            xls.Recalc();
            budget_Equipo.BudgetEquipo(xls);
            xls.Recalc();
            budget_M_Obra.BudgetMObra(xls);
            xls.Recalc();
            budget_Valor_De_M_Obra.Budget_Valor_M_De_Obra(xls);
            xls.Recalc();
            budget_Establecimiento.BudgetEstablecimiento(xls);
            
            xls.Recalc();
            budget_Sostenemiento.BudgetSostenemiento(xls);
            xls.Recalc();
            budget_Presupuesto.BudgetPresupuesto(xls);
            xls.Recalc();
            outcomeYAdjustment.Outcome_Y_Adjustment(xls);
            xls.Recalc();
            outcomeLAdjustment.Outcome_L_Adjustment(xls);
            //databaseSchema.Database_Schema(xls);
            xls.Recalc();
            outcomeTotalAdj.Outcome_TOTAL_Adj(xls);
            
            xls.Recalc();
            pre_Metric_Currency.Output1PreMetricCurrency(xls);
            
            xls.Recalc();
            prporcion_De_Productividad.ProporcionDeProductividad(xls);
            
            //var op = output.Outcome(xls, workspace);
            var op = databaseSchema.Database_Schema(xls, workspace);
            coopOutputDTO coopOutputDTO = new coopOutputDTO();
            coopOutputDTO.variableCostUSPound = 1.05;
            coopOutputDTO.fixedCostUSPound = 0.06;
            coopOutputDTO.totalCostAndDeprUSPound = 0.8;
            coopOutputDTO.totalCostUSPound = 1.91;
            coopOutputDTO.breakEvenCostUSPound = 1.34;
            Dictionary<String, object> outputDict = new Dictionary<String, object>();
            outputDict = op.Output;
            outputDict.Add("Coop", coopOutputDTO);
            //var futuresPrice = getFuturesPrice();
            
            //outputDict.Add("FuturesPrice", futuresPrice.Result);
            ChartDataDTO cdata = new ChartDataDTO();
            cdata.Output = outputDict;
            //Save the file as XLS
            //xls.Save(openFileDialog1.FileName);
            return cdata;
        }

        private async Task<string> getFuturesPrice()
        {
            var url = "https://aganalyticsdev.eastus2.cloudapp.azure.com/agriskmanagement/api/dataservice?sql=SELECT%20Top%201%20[SettlementPrice]/37500%20as%20p%20FROM%20[AgDB].[dbo].[CommodityFutures]%20where%20[Date]%20%3E%20%272017-12-31%27%20and%20Commodity%20=%20%27CoffeeC%27";
            var futuresPrice = "";
            HttpClient client = new HttpClient();
            HttpResponseMessage response = await client.GetAsync(url);
            if (response.IsSuccessStatusCode)
            {
                futuresPrice = await response.Content.ReadAsStringAsync();
            }
            string coffeeFuturesPrice = futuresPrice.Split("\n")[1];
            return coffeeFuturesPrice;
        }

        private void CreateFileForSheet2(XlsFile xls, Double peakHectares, Double oldHectares)
        {
            xls.NewFile(3, TExcelFileFormat.v2016);    //Create a new Excel file with 3 sheets.

            //Set the names of the sheets
            xls.ActiveSheet = 1;
            xls.SheetName = "Sheet1";
            xls.ActiveSheet = 2;
            xls.SheetName = "Sheet2";
            xls.ActiveSheet = 3;
            xls.SheetName = "Sheet3";

            xls.ActiveSheet = 2;    //Set the sheet we are working in.

            //Global Workbook Options
            xls.OptionsCheckCompatibility = false;

            //Sheet Options
            xls.SheetName = "Sheet2";

            //Printer Settings
            xls.PrintOptions = TPrintOptions.Orientation | TPrintOptions.NoPls;

            //Set the cell values
            xls.SetCellValue(9, 8, peakHectares);
            xls.SetCellValue(9, 9, oldHectares);

            //Cell selection and scroll position.
            xls.SelectCell(6, 8, false);

            //Standard Document Properties - Most are only for xlsx files. In xls files FlexCel will only change the Creation Date and Modified Date.
            xls.DocumentProperties.SetStandardProperty(TPropertyId.Author, "SAPTARSHI MALLICK");

            //You will normally not set LastSavedBy, since this is a new file.
            //If you don't set it, FlexCel will use the creator instead.
            //    xls.DocumentProperties.SetStandardProperty(TPropertyId.LastSavedBy, "SAPTARSHI MALLICK");


            //You will normally not set CreateDateTime, since this is a new file and FlexCel will automatically use the current datetime.
            //But if you are editing a file and want to preserve the original creation date, you need to either set PreserveCreationDate to true:
            //    xls.DocumentProperties.PreserveCreationDate = true;
            //Or you can hardcode a creating date by setting it in UTC time, ISO8601 format:
            //    xls.DocumentProperties.SetStandardProperty(TPropertyId.CreateTimeDate, "2017-11-09T15:42:06Z");

            xls.DocumentProperties.SetStandardProperty(TPropertyId.Company, "Cornell University");

            //xls.Save(System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Personal), "test1.xlsx"));
        }

        private void CreateFileForSheet1(XlsFile xls, Double earlyHectares)
        {
            //xls.NewFile(3, TExcelFileFormat.v2016);    //Create a new Excel file with 3 sheets.

            //Set the names of the sheets
            xls.ActiveSheet = 1;
            xls.SheetName = "Sheet1";
            xls.ActiveSheet = 2;
            xls.SheetName = "Sheet2";
            xls.ActiveSheet = 3;
            xls.SheetName = "Sheet3";

            xls.ActiveSheet = 1;    //Set the sheet we are working in.

            //Global Workbook Options
            xls.OptionsCheckCompatibility = false;

            //Printer Settings
            xls.PrintOptions = TPrintOptions.Orientation | TPrintOptions.NoPls;

            //Set the cell values
            xls.SetCellValue(9, 8, earlyHectares);

            //Cell selection and scroll position.
            xls.SelectCell(9, 8, false);

            //Standard Document Properties - Most are only for xlsx files. In xls files FlexCel will only change the Creation Date and Modified Date.
            xls.DocumentProperties.SetStandardProperty(TPropertyId.Author, "SAPTARSHI MALLICK");

            //You will normally not set LastSavedBy, since this is a new file.
            //If you don't set it, FlexCel will use the creator instead.
            //    xls.DocumentProperties.SetStandardProperty(TPropertyId.LastSavedBy, "SAPTARSHI MALLICK");


            //You will normally not set CreateDateTime, since this is a new file and FlexCel will automatically use the current datetime.
            //But if you are editing a file and want to preserve the original creation date, you need to either set PreserveCreationDate to true:
            //    xls.DocumentProperties.PreserveCreationDate = true;
            //Or you can hardcode a creating date by setting it in UTC time, ISO8601 format:
            //    xls.DocumentProperties.SetStandardProperty(TPropertyId.CreateTimeDate, "2017-11-09T15:42:06Z");

            xls.DocumentProperties.SetStandardProperty(TPropertyId.Company, "Cornell University");

            //xls.Save(System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Personal), "test1.xlsx"));
        }

        private ChartDataDTO CreateFileForSheet3(XlsFile xls)
        {
            //xls.NewFile(3, TExcelFileFormat.v2016);    //Create a new Excel file with 3 sheets.

            //Set the names of the sheets
            xls.ActiveSheet = 1;
            xls.SheetName = "Sheet1";
            xls.ActiveSheet = 2;
            xls.SheetName = "Sheet2";
            xls.ActiveSheet = 3;
            xls.SheetName = "Sheet3";

            xls.ActiveSheet = 3;    //Set the sheet we are working in.

            //Global Workbook Options
            xls.OptionsCheckCompatibility = false;

            //Sheet Options
            xls.SheetName = "Sheet3";

            //Set the cell values
            //xls.SetCellValue(9, 9, oldHectares);
            xls.SetCellValue(9, 8, new TFormula("=Sheet1!H9 + Sheet2!H9"));
            xls.SetCellValue(9, 10, new TFormula("=Sheet1!H9 + Sheet2!H9 +Sheet2!I9"));
            xls.SetCellValue(9, 11, new TFormula("=Sheet2!H9 +Sheet2!I9"));
            xls.SetCellValue(9, 12, new TFormula("=Sheet2!I9 - Sheet2!H9"));

            //Cell selection and scroll position.
            xls.SelectCell(9, 8, false);

            //Standard Document Properties - Most are only for xlsx files. In xls files FlexCel will only change the Creation Date and Modified Date.
            xls.DocumentProperties.SetStandardProperty(TPropertyId.Author, "SAPTARSHI MALLICK");

            xls.Recalc();
            //xls.Save(System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Personal), "test1.xlsx"));
            Double coopval1 = Convert.ToDouble(xls.GetCellValue(9, 11)) + 0.03;
            Double coopval2 = Convert.ToDouble(xls.GetCellValue(9, 12)) - 0.02;
            Double coopval3 = Convert.ToDouble(xls.GetCellValue(9, 10)) + 0.05;
            ChartDataDTO cd = new ChartDataDTO();
            List<Double> prod = new List<double>();
            List<Double> coop = new List<double>();
            coop.Add(Math.Round(coopval1, 2));
            coop.Add(Math.Round(coopval2, 2));
            coop.Add(Math.Round(coopval3, 2));
            //cd.cooperative = coop;
            prod.Add(Math.Round(Convert.ToDouble(xls.GetCellValue(9, 11)),2));
            prod.Add(Math.Round(Convert.ToDouble(xls.GetCellValue(9, 12)),2));
            prod.Add(Math.Round(Convert.ToDouble(xls.GetCellValue(9, 10)),2));
            //cd.producer = prod;
            

            //You will normally not set LastSavedBy, since this is a new file.
            //If you don't set it, FlexCel will use the creator instead.
            //    xls.DocumentProperties.SetStandardProperty(TPropertyId.LastSavedBy, "SAPTARSHI MALLICK");


            //You will normally not set CreateDateTime, since this is a new file and FlexCel will automatically use the current datetime.
            //But if you are editing a file and want to preserve the original creation date, you need to either set PreserveCreationDate to true:
            //    xls.DocumentProperties.PreserveCreationDate = true;
            //Or you can hardcode a creating date by setting it in UTC time, ISO8601 format:
            //    xls.DocumentProperties.SetStandardProperty(TPropertyId.CreateTimeDate, "2017-11-09T15:42:06Z");

            xls.DocumentProperties.SetStandardProperty(TPropertyId.Company, "Cornell University");

           

            return cd;//Convert.ToDouble(xls.GetCellValue(9, 12));
        }

        public void SaveUserInputs(string id, ChartInputDTO chartInputDTO)
        {
            String timeStamp = DateTime.Now.ToString();
            var conn = _iconfiguration.GetSection("ConnectionStrings").GetSection("CoffeeConnStr").Value;
            string sqlQuery = String.Format("Insert INTO [AVFCoffee].[dbo].[UserInput]" +
                   "(HectTreesEarly, HectTreesPeak, HectTreesOld, Conventional, Organic, Transition, WagePerDay, YieldPerHect, TransportCost, FinalPrice, UserID, ExpSolesChem, ExpSolesOrg, TimeStamp) VALUES" +
                   "(@HectEarly, @HectPeak, @HectOld, @Conv, @Org, @Trans, @Wpd, @YieldHect, @TransCost, @FinalPrice, @UserID, @ExpSolesChem, @ExpSolesOrg, @TimeStamp)");
            using (SqlConnection connect = new SqlConnection(conn))
            {
                connect.Open();
                SqlCommand command = new SqlCommand(sqlQuery);
                command.Parameters.AddWithValue("@HectEarly", chartInputDTO.earlyHectares);
                command.Parameters.AddWithValue("@HectPeak", chartInputDTO.peakHectares);
                command.Parameters.AddWithValue("@HectOld", chartInputDTO.oldHectares);
                command.Parameters.AddWithValue("@Conv", chartInputDTO.conventional);
                command.Parameters.AddWithValue("@Org", chartInputDTO.organic);
                command.Parameters.AddWithValue("@Trans", chartInputDTO.transition);
                command.Parameters.AddWithValue("@Wpd", chartInputDTO.workerSalarySoles);
                command.Parameters.AddWithValue("@YieldHect", chartInputDTO.productionQuintales);
                command.Parameters.AddWithValue("@TransCost", chartInputDTO.transportCostSoles);
                command.Parameters.AddWithValue("@FinalPrice", chartInputDTO.costPriceSolesPerQuintal);
                command.Parameters.AddWithValue("@UserID", id);
                command.Parameters.AddWithValue("@ExpSolesChem", chartInputDTO.expSolesChem);
                command.Parameters.AddWithValue("@ExpSolesOrg", chartInputDTO.expSolesOrg);
                command.Parameters.AddWithValue("@TimeStamp", timeStamp);
                command.Connection = connect;
                int result = command.ExecuteNonQuery();
                connect.Close();
                // Check Error
                if (result < 0)
                    Console.WriteLine("Error inserting data into Database!");
            }
        }

        

        public LoginInfoDTO GetUserInputs(String id)
        {
            try
            {
                UserInfoDTO uInfo = new UserInfoDTO();
                Dictionary<String, object> cOut = new Dictionary<String, object>();
                ChartDataDTO cData = new ChartDataDTO();
                ChartInputDTO chInput = new ChartInputDTO();
                ProducerOutputEnglishDTO pOutEnglishDTO = new ProducerOutputEnglishDTO();
                ProducerOutputSpanishDTO pOutSpanishDTO = new ProducerOutputSpanishDTO();
                LoginInfoDTO lInfo = new LoginInfoDTO();
                var conn = _iconfiguration.GetSection("ConnectionStrings").GetSection("CoffeeConnStr").Value;
                using (SqlConnection con = new SqlConnection(conn))
                {
                    con.Open();

                    SqlCommand comm = new SqlCommand("Select * from [AVFCoffee].[dbo].[UserInput] where UserID = @userid AND [TimeStamp] = (SELECT MAX(timestamp) FROM[AVFCoffee].[dbo].[UserInput] where UserID = @userid)", con);
                    comm.Parameters.AddWithValue("@userid", id);
                    // int result = command.ExecuteNonQuery();
                    using (SqlDataReader reader = comm.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            //var output = String.Format("{0}", reader["HectTreesEarly"]);
                            chInput.earlyHectares = Convert.ToDouble(reader["HectTreesEarly"].ToString());
                            chInput.peakHectares = Convert.ToDouble(reader["HectTreesPeak"].ToString());
                            chInput.oldHectares = Convert.ToDouble(reader["HectTreesOld"].ToString());
                            chInput.conventional = Convert.ToBoolean(reader["Conventional"].ToString());
                            chInput.organic = Convert.ToBoolean(reader["Organic"].ToString());
                            chInput.transition = Convert.ToBoolean(reader["Transition"].ToString());
                            chInput.workerSalarySoles = Convert.ToDouble(reader["WagePerDay"].ToString());
                            chInput.productionQuintales = Convert.ToDouble(reader["YieldPerHect"].ToString());
                            chInput.transportCostSoles = Convert.ToDouble(reader["TransportCost"].ToString());
                            chInput.costPriceSolesPerQuintal = Convert.ToDouble(reader["FinalPrice"].ToString());
                            chInput.expSolesOrg = Convert.ToDouble(reader["ExpSolesOrg"].ToString());
                            chInput.expSolesChem = Convert.ToDouble(reader["ExpSolesChem"].ToString());
                        }
                    }

                    con.Close();
                }

                using (SqlConnection con = new SqlConnection(conn))
                {
                    con.Open();

                    SqlCommand comm = new SqlCommand("Select * from [AVFCoffee].[dbo].[OutputProducer] where UserID = @userid AND [TimeStamp] = (SELECT MAX(timestamp) FROM[AVFCoffee].[dbo].[OutputProducer] where UserID = @userid)", con);
                    comm.Parameters.AddWithValue("@userid", id);
                    // int result = command.ExecuteNonQuery();
                    using (SqlDataReader reader = comm.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            //var output = String.Format("{0}", reader["HectTreesEarly"]);
                            pOutEnglishDTO.variableCostUSPound = Convert.ToDouble(reader["VariableCostUSPound"].ToString());
                            pOutEnglishDTO.fixedCostUSPound = Convert.ToDouble(reader["FixedCostUSPound"].ToString());
                            pOutEnglishDTO.totalCostAndDeprUSPound = Convert.ToDouble(reader["TotalCostAndDeprUSPound"].ToString());
                            pOutEnglishDTO.totalCostUSPound = Convert.ToDouble(reader["TotalCostUSPound"].ToString());
                            pOutEnglishDTO.breakEvenCostUSPound = Convert.ToDouble(reader["BreakEvenCostUSPound"].ToString());
                            pOutEnglishDTO.futuresPrice = Convert.ToDouble(reader["FuturesPrice"].ToString());
                            pOutEnglishDTO.status = true;
                            pOutSpanishDTO.variableCostSolesHect = Convert.ToDouble(reader["VariableCostSolesHect"].ToString());
                            pOutSpanishDTO.variableCostUSHect = Convert.ToDouble(reader["VariableCostUSHect"].ToString());
                            pOutSpanishDTO.totalCostUSHect = Convert.ToDouble(reader["TotalCostUSHect"].ToString());
                            pOutSpanishDTO.totalCostSolesHect = Convert.ToDouble(reader["TotalCostSolesHect"].ToString());
                            pOutSpanishDTO.breakEvenCostUSPound = Convert.ToDouble(reader["BreakEvenCostUSPound"].ToString());
                        } else
                        {
                            pOutEnglishDTO.status = false;
                        }
                    }

                    con.Close();
                }

                using (SqlConnection con = new SqlConnection(conn))
                {
                    con.Open();

                    SqlCommand comm = new SqlCommand("Select * from [AVFCoffee].[dbo].[User] where UserID = @userid", con);
                    comm.Parameters.AddWithValue("@userid", id);
                    // int result = command.ExecuteNonQuery();
                    using (SqlDataReader reader = comm.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            uInfo.Language = reader["Language"].ToString();
                            uInfo.UserName = reader["UserName"].ToString();
                            //var output = String.Format("{0}", reader["HectTreesEarly"]);
                            
                        }
                    }

                    con.Close();
                }
                if (uInfo.Language == "B2C_1_siupin_es")
                {
                    uInfo.Language = "ES";
                } else
                {
                    uInfo.Language = "EN";
                }
                coopOutputDTO coopOutputDTO = new coopOutputDTO();
                coopOutputDTO.variableCostUSPound = 1.05;
                coopOutputDTO.fixedCostUSPound = 0.06;
                coopOutputDTO.totalCostAndDeprUSPound = 0.8;
                coopOutputDTO.totalCostUSPound = 1.91;
                coopOutputDTO.breakEvenCostUSPound = 1.34;

                cOut.Add("ProducerOutputEnglish", pOutEnglishDTO);
                cOut.Add("ProducerOutputSpanish", pOutSpanishDTO);
                cOut.Add("Coop", coopOutputDTO);
                //var futuresPrice = getFuturesPrice();
                //cOut.Add("FuturesPrice", futuresPrice.Result);
                cData.Output = cOut;
                Dictionary<String, object> outDict = new Dictionary<String, object>();
                outDict.Add("Inputs", chInput);
                outDict.Add("Outputs", cOut);

                outDict.Add("User", uInfo);
                lInfo.loginfo = outDict;
                return lInfo;
            }
            catch (Exception e)
            {
                return null; 
            }
        }

        public ChartDataDTO SaveUserOutputs(string id, ChartDataDTO chartDataDTO)
        {
            String timeStamp = DateTime.Now.ToString();
            int resultProd, resultCoop;
            Dictionary<String, object> dict = new Dictionary<String, object>();
            dict = chartDataDTO.Output;
            Double variableCostUSpound, fixedCostUSPound, totalCostAndDeprUSPound, totalCostUSPound, variableCostUSHect, variableCostSolesHect, totalCostUSHect, totalCostSolesHect, breakEvenCostUSPound,
                coopVariableUSPound, coopFixedUSPound, coopTotalCostAndDeprUSPound, coopTotalCostUSPound, coopBreakEvenCostUSPound;
            String CoopId;
            var prod = dict["ProducerOutputEnglish"];
            ProducerOutputEnglishDTO producerEnglish = JsonConvert.DeserializeObject<ProducerOutputEnglishDTO>(prod.ToString());
            //JsonConvert.DeserializeObject<producerEnglish>;
            variableCostUSpound = producerEnglish.variableCostUSPound;
            fixedCostUSPound = producerEnglish.fixedCostUSPound;
            totalCostAndDeprUSPound = producerEnglish.totalCostAndDeprUSPound;
            totalCostUSPound = producerEnglish.totalCostUSPound;
            var prodSpan = dict["ProducerOutputSpanish"];
            ProducerOutputSpanishDTO producerSpanish = JsonConvert.DeserializeObject<ProducerOutputSpanishDTO>(prodSpan.ToString());
            variableCostUSHect = producerSpanish.variableCostUSHect;
            variableCostSolesHect = producerSpanish.variableCostSolesHect;
            totalCostUSHect = producerSpanish.totalCostUSHect;
            totalCostSolesHect = producerSpanish.totalCostSolesHect;
            breakEvenCostUSPound = producerSpanish.breakEvenCostUSPound;
            var coop1 = dict["Coop"];
            coopOutputDTO coop = JsonConvert.DeserializeObject<coopOutputDTO>(coop1.ToString());
            CoopId = Guid.NewGuid().ToString();
            coopVariableUSPound = coop.variableCostUSPound;
            coopFixedUSPound = coop.fixedCostUSPound;
            coopTotalCostAndDeprUSPound = coop.totalCostAndDeprUSPound;
            coopTotalCostUSPound = coop.totalCostUSPound;
            coopBreakEvenCostUSPound = coop.breakEvenCostUSPound;
            var conn = _iconfiguration.GetSection("ConnectionStrings").GetSection("CoffeeConnStr").Value;
            string sqlQuery = String.Format("Insert INTO [AVFCoffee].[dbo].[OutputProducer]" +
                   "(UserID, VariableCostUSPound, FixedCostUSPound, TotalCostAndDeprUSPound, TotalCostUSPound, VariableCostUSHect, VariableCostSolesHect, TotalCostUSHect, " +
                   "TotalCostSolesHect, BreakEvenCostUSPound, TimeStamp, FuturesPrice) VALUES" +
                   "(@id, @variableCostUSPound, @fixedCostUSPound, @totalCostAndDeprUSPound, @totalCostUSPound, @variableCostUSHect, @variableCostSolesHect, @totalCostUSHect, @totalCostSolesHect" +
                   ", @breakEvenCostUSPound, @TimeStamp, @FuturesPrice)");
            using (SqlConnection connect = new SqlConnection(conn))
            {
                connect.Open();
                SqlCommand command = new SqlCommand(sqlQuery);
                command.Parameters.AddWithValue("@id", id);
                command.Parameters.AddWithValue("@variableCostUSPound", variableCostUSpound);
                command.Parameters.AddWithValue("@fixedCostUSPound", fixedCostUSPound);
                command.Parameters.AddWithValue("@totalCostAndDeprUSPound", totalCostAndDeprUSPound);
                command.Parameters.AddWithValue("@totalCostUSPound", totalCostUSPound);
                command.Parameters.AddWithValue("@variableCostUSHect", variableCostUSHect);
                command.Parameters.AddWithValue("@variableCostSolesHect", variableCostSolesHect);
                command.Parameters.AddWithValue("@totalCostUSHect", totalCostUSHect);
                command.Parameters.AddWithValue("@totalCostSolesHect", totalCostSolesHect);
                command.Parameters.AddWithValue("@breakEvenCostUSPound", breakEvenCostUSPound);
                command.Parameters.AddWithValue("@TimeStamp", timeStamp);
                command.Parameters.AddWithValue("@FuturesPrice", getFuturesPrice().Result);
                command.Connection = connect;
                resultProd = command.ExecuteNonQuery();
                connect.Close();
            }

            string sqlQueryCoop = String.Format("Insert INTO [AVFCoffee].[dbo].[OutputCoop]" +
                   "(CoopID, VariableCostUSPound, FixedCostUSPound, TotalCostAndDeprUSPound, TotalCostUSPound, BreakEvenCostUSPound, TimeStamp) VALUES" +
                   "(@id, @variableCostUSPound, @fixedCostUSPound, @totalCostAndDeprUSPound, @totalCostUSPound, @breakEvenCostUSPound, @TimeStamp)");

            using (SqlConnection connect = new SqlConnection(conn))
            {
                connect.Open();
                SqlCommand command = new SqlCommand(sqlQueryCoop);
                command.Parameters.AddWithValue("@id", CoopId);
                command.Parameters.AddWithValue("@variableCostUSPound", coopVariableUSPound);
                command.Parameters.AddWithValue("@fixedCostUSPound", coopFixedUSPound);
                command.Parameters.AddWithValue("@totalCostAndDeprUSPound", coopTotalCostAndDeprUSPound);
                command.Parameters.AddWithValue("@totalCostUSPound", coopTotalCostUSPound);
                command.Parameters.AddWithValue("@breakEvenCostUSPound", coopBreakEvenCostUSPound);
                command.Parameters.AddWithValue("@TimeStamp", timeStamp);
                command.Connection = connect;
                resultCoop = command.ExecuteNonQuery();
                connect.Close();
                
            }

            if (resultCoop < 0 || resultProd < 0)
                Console.WriteLine("Error inserting data into Database!");
            return null;
            //throw new NotImplementedException();
        }

        public UserInfoDTO SaveUserInfo(UserInfoDTO userInfoDTO)
        {
            int resultUser;

            var conn = _iconfiguration.GetSection("ConnectionStrings").GetSection("CoffeeConnStr").Value;
            using (SqlConnection con = new SqlConnection(conn))
            {
                con.Open();

                SqlCommand comm = new SqlCommand("Select * from [AVFCoffee].[dbo].[User] where UserID = @userid", con);
                comm.Parameters.AddWithValue("@userid", userInfoDTO.UserID);
                //comm.Parameters.AddWithValue("@language", userInfoDTO.Language);
                // int result = command.ExecuteNonQuery();
                using (SqlDataReader reader = comm.ExecuteReader())
                {
                    if (!reader.Read())
                    {
                        string sqlQueryUser = String.Format("Insert INTO [AVFCoffee].[dbo].[User]" +
                    "(UserID, CoopID, UserName, Language) VALUES" +
                    "(@id, @CoopID, @UserName, @Language)");
                        //var conn = _iconfiguration.GetSection("ConnectionStrings").GetSection("CoffeeConnStr").Value;
                        using (SqlConnection connect = new SqlConnection(conn))
                        {
                            connect.Open();
                            SqlCommand command = new SqlCommand(sqlQueryUser);
                            command.Parameters.AddWithValue("@id", userInfoDTO.UserID);
                            command.Parameters.AddWithValue("@CoopID", 0);
                            command.Parameters.AddWithValue("@UserName", userInfoDTO.UserName);
                            command.Parameters.AddWithValue("@Language", userInfoDTO.Language);
                            command.Connection = connect;
                            resultUser = command.ExecuteNonQuery();
                            connect.Close();

                        }
                    } else
                    {
                        string sqlQueryUser = String.Format("Update [AVFCoffee].[dbo].[User]" +
                   "Set [Language] = @language Where UserID = @userid");
                        //var conn = _iconfiguration.GetSection("ConnectionStrings").GetSection("CoffeeConnStr").Value;
                        using (SqlConnection connect = new SqlConnection(conn))
                        {
                            connect.Open();
                            SqlCommand command = new SqlCommand(sqlQueryUser);
                            command.Parameters.AddWithValue("@language", userInfoDTO.Language);
                            command.Parameters.AddWithValue("@userid", userInfoDTO.UserID);
                            command.Connection = connect;
                            resultUser = command.ExecuteNonQuery();
                            connect.Close();

                        }
                    }
                }

                con.Close();
            }
            
            return userInfoDTO;
        }
    }
}
