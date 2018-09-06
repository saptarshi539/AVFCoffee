using CoffeeCore.Interfaces;
using FlexCel.Core;
using CoffeeCore.DTO;
using FlexCel.XlsAdapter;
using System;
using System.Collections.Generic;
using Microsoft.Extensions.Configuration;
using System.Data.SqlClient;
using Newtonsoft.Json;
using CoffeeInfrastructure.Helpers;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;

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
            Inputs_2 inputs_2 = new Inputs_2();
            InAdvanced inAdvanced = new InAdvanced();
            Language language = new Language();
            Metrics_English metrics_English = new Metrics_English();
            Metrics_Spanish metrics_Spanish = new Metrics_Spanish();
            InputsAdvanced2English inputsAdvanced2English = new InputsAdvanced2English();
            InputsAdvanced2Spanish inputsAdvanced2Spanish = new InputsAdvanced2Spanish();
            AdvancedInputs advancedInputs = new AdvancedInputs();
            Budget_Equipo budget_Equipo = new Budget_Equipo();
            Budget_Establecimiento budget_Establecimiento = new Budget_Establecimiento();
            Budget_M_Obra budget_M_Obra = new Budget_M_Obra();
            Budget_Presupuesto budget_Presupuesto = new Budget_Presupuesto();
            Budget_Sostenemiento budget_Sostenemiento = new Budget_Sostenemiento();
            Budget_Valor_de_M_Obra budget_Valor_De_M_Obra = new Budget_Valor_de_M_Obra();
            Proportions proportions = new Proportions();
            Gral_Conf gral_Conf = new Gral_Conf();
            Metrics metrics = new Metrics();
            Conversiones conversiones = new Conversiones();
            Prporcion_de_productividad prporcion = new Prporcion_de_productividad();
            Inputs_TOT inputs_TOT_advanced = new Inputs_TOT();
            Inputs_2_Conv inputs_2_Conv = new Inputs_2_Conv();
            Inputs_1 inputs_1 = new Inputs_1();
            DatabaseSchema databaseSchema = new DatabaseSchema();
            General_Conf_Summary_Spa conf_Summary_Spa = new General_Conf_Summary_Spa();
            OutcomeTotalAdj outcomeTotalAdj = new OutcomeTotalAdj();
            OutcomeYAdjustment outcomeYAdjustment = new OutcomeYAdjustment();
            Output1_pre_metric_currency output1_Pre_Metric_Currency = new Output1_pre_metric_currency();
            OutcomeLAdjustment outcomeLAdjustment = new OutcomeLAdjustment();
            Output output = new Output();
            InputsEnglish inputsEnglish = new InputsEnglish();
            InputsSpanish inputsSpanish = new InputsSpanish();
            Inputs_1_Ref inputs_1_Ref = new Inputs_1_Ref();
            Input_1 input_1 = new Input_1();
            Inputs inputs = new Inputs();
            XlsFile xls = new XlsFile(true);
            TWorkspace workspace = new TWorkspace();

            //get metrics
            MetricsDTO md = new MetricsDTO();
            md = GetMetrics();
            //xls.Open("file");

            workspace.Add(xls.ActiveFileName, xls);
            //actual calculation taking place in the excel sheet
            metrics.metrics(xls, md);
            inputs.inputs(xls, earlyHectares, peakHectares, oldHectares, conventional, organic, transition, workerSalarySoles, productionQuintales, transportCostSoles,
                costPriceSolesPerQuintal, expSolesChem, expSolesOrg);
            //databaseSchema.Database_Schema(xls);
            
            output.Outcome(xls);
            language.language(xls);
            inputsEnglish.InputEnglish(xls);
            inputsSpanish.InputSpanish(xls);
            input_1.Input_1_default(xls);
            conversiones.conversiones(xls);
            metrics_English.MetricsEnglish(xls);
            metrics_Spanish.MetricsSpanish(xls);
            inputs_2.Inputs_2_Default(xls);

            //databaseSchema.Database_Schema(xls, workspace);
            inputs_2_Conv.Inputs_2_Conv_inputs(xls);
            inputsAdvanced2English.InputAdvanced2English(xls);

            inputs_1.CreateFile(xls);
            advancedInputs.Budget_Supuestos(xls);
            proportions.proportions(xls);
            budget_Equipo.BudgetEquipo(xls);
            budget_Establecimiento.BudgetEstablecimiento(xls);
            budget_M_Obra.BudgetMObra(xls);
            budget_Sostenemiento.BudgetSostenemiento(xls);
            budget_Valor_De_M_Obra.Budget_Valor_M_De_Obra(xls);
            budget_Presupuesto.BudgetPresupuesto(xls);

            //if (Language == "EN")
            //{
            inputsAdvanced2Spanish.InputAdvancedSpanish(xls, "EN");

            prporcion.ProporcionDeProductividad(xls);
            inputs_1_Ref.inputs1Ref(xls);

            gral_Conf.Gral_Conf_Summary(xls);
            conf_Summary_Spa.GeneralConfSummarySpa(xls);



            inputs_TOT_advanced.CreateFile(xls);
            ////var advancedInputsDict = new Dictionary<string, object>();

            ////if (Language == "EN")
            ////{

            inAdvanced.Inputs_Advanced(xls);
            outcomeYAdjustment.Outcome_Y_Adjustment(xls);
            output1_Pre_Metric_Currency.Output1PreMetricCurrency(xls);
            outcomeLAdjustment.Outcome_L_Adjustment(xls);
            outcomeTotalAdj.Outcome_TOTAL_Adj(xls);
            //}
            //else
            //{
            //advancedInputsDict = inputsAdvanced2Spanish.InputAdvancedSpanish(xls, Language);
            //}



            //var op = output.Outcome(xls, workspace);
            var op = databaseSchema.Database_Schema(xls);
           
            
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

        private MetricsDTO GetMetrics()
        {
            MetricsInputDTO minput = new MetricsInputDTO();
            var conn = _iconfiguration.GetSection("ConnectionStrings").GetSection("CoffeeConnStr").Value;
            using (SqlConnection con = new SqlConnection(conn))
            {
                con.Open();
                var id = "1234";

                SqlCommand comm = new SqlCommand("Select * from [AVFCoffee].[dbo].[CoopGeneralConfig] where CoopID = @coopid", con);
                comm.Parameters.AddWithValue("@coopid", id);
                // int result = command.ExecuteNonQuery();
                using (SqlDataReader reader = comm.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        minput.coffeemeasurekilograms = Convert.ToBoolean(reader["CoffeeMeasureKilograms"].ToString());
                        minput.coffeemeasurepounds = Convert.ToBoolean(reader["CoffeeMeasurePounds"].ToString());
                        minput.coffeemeasurequintales = Convert.ToBoolean(reader["CoffeeMeasureQuintales"].ToString());
                        minput.coffeemeasurearrobas = Convert.ToBoolean(reader["CoffeeMeasureArrobas"].ToString());
                        minput.coffeemeasurecargas = Convert.ToBoolean(reader["CoffeeMeasureCargas"].ToString());
                        minput.lengthmeasuremeters = Convert.ToBoolean(reader["LengthMeasureMeters"].ToString());
                        minput.lengthmeasurefeet = Convert.ToBoolean(reader["LengthMeasureFeet"].ToString());
                        minput.farmareameasurehectares = Convert.ToBoolean(reader["FarmAreaMeasureHectares"].ToString());
                        minput.farmareameasuremanzanas = Convert.ToBoolean(reader["FarmAreaMeasureManzanas"].ToString());
                        minput.applicationmeasurekilograms = Convert.ToBoolean(reader["ApplicationMeasureKilograms"].ToString());
                        minput.applicationmeasurepounds = Convert.ToBoolean(reader["ApplicationMeasurePounds"].ToString());
                        minput.capacitymeasureliters = Convert.ToBoolean(reader["CapacityMeasureLiters"].ToString());
                        minput.capacitymeasuregallons = Convert.ToBoolean(reader["CapacityMeasureGallons"].ToString());
                        minput.currencyboliviaboliviano = Convert.ToBoolean(reader["CurrencyBoliviaBoliviano"].ToString());
                        minput.currencybrazilreal = Convert.ToBoolean(reader["CurrencyBrazilReal"].ToString());
                        minput.currencycolombiapeso = Convert.ToBoolean(reader["CurrencyColombiaPeso"].ToString());
                        minput.currencycostaricacolon = Convert.ToBoolean(reader["CurrencyCostaRicaColon"].ToString());
                        minput.currencycubapeso = Convert.ToBoolean(reader["CurrencyCubaPeso"].ToString());
                        minput.currencyguatemalaquetzal = Convert.ToBoolean(reader["CurrencyGuatemalaQuetzal"].ToString());
                        minput.currencyhaitigourde = Convert.ToBoolean(reader["CurrencyHaitiGourde"].ToString());
                        minput.currencyhonduraslempira = Convert.ToBoolean(reader["CurrencyHondurasLempira"].ToString());
                        minput.currencyjamaicadollar = Convert.ToBoolean(reader["CurrencyJamaicaDollar"].ToString());
                        minput.currencymexicopeso = Convert.ToBoolean(reader["CurrencyMexicoPeso"].ToString());
                        minput.currencynicaraguacordoba = Convert.ToBoolean(reader["CurrencyNicaraguaCordoba"].ToString());
                        minput.currencyperusol = Convert.ToBoolean(reader["CurrencyPeruSol"].ToString());
                        minput.currencyusdollar = Convert.ToBoolean(reader["CurrencyUSDollar"].ToString());
                        minput.currencyvenezuelabolivar = Convert.ToBoolean(reader["CurrencyVenezuelaBolivar"].ToString());
                    }
                    reader.Close();

                }
                con.Close();
            }
            MetricsDTO md = new MetricsDTO();
            if (minput.applicationmeasurekilograms)
            {
                md.applicationmeasurekilograms = 1;
            }
            else
            {
                md.applicationmeasurekilograms = 0;
            }
            if (minput.applicationmeasurepounds)
            {
                md.applicationmeasurepounds = 1;
            }
            else
            {
                md.applicationmeasurepounds = 0;
            }
            if (minput.capacitymeasuregallons)
            {
                md.capacitymeasuregallons = 1;
            }
            else
            {
                md.capacitymeasuregallons = 0;
            }
            if (minput.capacitymeasureliters)
            {
                md.capacitymeasureliters = 1;
            }
            else
            {
                md.capacitymeasureliters = 0;
            }

            if (minput.coffeemeasurearrobas)
            {
                md.coffeemeasurearrobas = 1;
            }
            else
            {
                md.coffeemeasurearrobas = 0;
            }
            if (minput.coffeemeasurecargas)
            {
                md.coffeemeasurecargas = 1;
            }
            else
            {
                md.coffeemeasurecargas = 0;
            }

            if (minput.coffeemeasurekilograms)
            {
                md.coffeemeasurekilograms = 1;
            }
            else
            {
                md.coffeemeasurekilograms = 0;
            }

            if (minput.coffeemeasurepounds)
            {
                md.coffeemeasurepounds = 1;
            }
            else
            {
                md.coffeemeasurepounds = 0;
            }
            if (minput.coffeemeasurequintales)
            {
                md.coffeemeasurequintales = 1;
            }
            else
            {
                md.coffeemeasurequintales = 0;
            }
            if (minput.currencyboliviaboliviano)
            {
                md.currencyboliviaboliviano = 1;
            }
            else
            {
                md.currencyboliviaboliviano = 0;
            }
            if (minput.currencybrazilreal)
            {
                md.currencybrazilreal = 1;
            }
            else
            {
                md.currencybrazilreal = 0;
            }
            if (minput.currencycolombiapeso)
            {
                md.currencycolombiapeso = 1;
            }
            else
            {
                md.currencycolombiapeso = 0;
            }
            if (minput.currencycostaricacolon)
            {
                md.currencycostaricacolon = 1;
            }
            else
            {
                md.currencycostaricacolon = 0;
            }
            if (minput.currencycubapeso)
            {
                md.currencycubapeso = 1;
            }
            else
            {
                md.currencycubapeso = 0;
            }

            if (minput.currencyguatemalaquetzal)
            {
                md.currencyguatemalaquetzal = 1;
            }
            else
            {
                md.currencyguatemalaquetzal = 0;
            }

            if (minput.currencyhaitigourde)
            {
                md.currencyhaitigourde = 1;
            }
            else
            {
                md.currencyhaitigourde = 0;
            }

            if (minput.currencyhonduraslempira)
            {
                md.currencyhonduraslempira = 1;
            }
            else
            {
                md.currencyhonduraslempira = 0;
            }

            if (minput.currencyjamaicadollar)
            {
                md.currencyjamaicadollar = 1;
            }
            else
            {
                md.currencyjamaicadollar = 0;
            }

            if (minput.currencymexicopeso)
            {
                md.currencymexicopeso = 1;
            }
            else
            {
                md.currencymexicopeso = 0;
            }

            if (minput.currencynicaraguacordoba)
            {
                md.currencynicaraguacordoba = 1;
            }
            else
            {
                md.currencynicaraguacordoba = 0;
            }

            if (minput.currencyperusol)
            {
                md.currencyperusol = 1;
            }
            else
            {
                md.currencyperusol = 0;
            }

            if (minput.currencyusdollar)
            {
                md.currencyusdollar = 1;
            }
            else
            {
                md.currencyusdollar = 0;
            }

            if (minput.currencyvenezuelabolivar)
            {
                md.currencyvenezuelabolivar = 1;
            }
            else
            {
                md.currencyvenezuelabolivar = 0;
            }
            if (minput.farmareameasurehectares)
            {
                md.farmareameasurehectares = 1;
            }
            else
            {
                md.farmareameasurehectares = 0;
            }
            if (minput.farmareameasuremanzanas)
            {
                md.farmareameasuremanzanas = 1;
            }
            else
            {
                md.farmareameasuremanzanas = 0;
            }

            if (minput.lengthmeasurefeet)
            {
                md.lengthmeasurefeet = 1;
            }
            else
            {
                md.lengthmeasurefeet = 0;
            }

            if (minput.lengthmeasuremeters)
            {
                md.lengthmeasuremeters = 1;
            }
            else
            {
                md.lengthmeasuremeters = 0;
            }
            return md;
        }

        private async Task<string> getFuturesPrice()
        {
            //TODO: change the expiry year every year and expiry month every month
            var url = "https://aganalyticsdev.eastus2.cloudapp.azure.com/agriskmanagement/api/dataservice?sql=SELECT%20Top%201%20[SettlementPrice]/100%20as%20p%20FROM%20[AgDB].[dbo].[CommodityFutures]%20where%20Commodity%20=%20%27CoffeeC%27%20and%20[ExpirationMonth]%20=%20%27May%27%20and%20[ExpirationYear]%20=%202018%20order%20by%20DaysToExpiry";
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
            prod.Add(Math.Round(Convert.ToDouble(xls.GetCellValue(9, 11)), 2));
            prod.Add(Math.Round(Convert.ToDouble(xls.GetCellValue(9, 12)), 2));
            prod.Add(Math.Round(Convert.ToDouble(xls.GetCellValue(9, 10)), 2));
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
            var userid = GetSmallHolderUserID(id);
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
                command.Parameters.AddWithValue("@UserID", userid);
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


        private String GetSmallHolderUserID(String number)
        {
            var conn = _iconfiguration.GetSection("ConnectionStrings").GetSection("CoffeeConnStr").Value;
            var id = "";
            using (SqlConnection con = new SqlConnection(conn))
            {
                con.Open();
                SqlCommand comm = new SqlCommand("Select * from [AVFCoffee].[dbo].[SmallHolder] where [PhoneNumber] = @phone", con);
                comm.Parameters.AddWithValue("@phone", number);
                using (SqlDataReader reader = comm.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        id = reader["FarmerID"].ToString();
                    }
                    reader.Close();
                }
                con.Close();
            }
            return id;
        }
        public LoginInfoDTO GetUserInputs(String number)
        {
            try
            {
                var conn = _iconfiguration.GetSection("ConnectionStrings").GetSection("CoffeeConnStr").Value;
                var id = GetSmallHolderUserID(number);
                UserInfoDTO uInfo = new UserInfoDTO();
                Dictionary<String, object> cOut = new Dictionary<String, object>();
                ChartDataDTO cData = new ChartDataDTO();
                ChartInputDTO chInput = new ChartInputDTO();
                ProducerOutputEnglishDTO pOutEnglishDTO = new ProducerOutputEnglishDTO();
                ProducerOutputSpanishDTO pOutSpanishDTO = new ProducerOutputSpanishDTO();
                LoginInfoDTO lInfo = new LoginInfoDTO();

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
                        }
                        else
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
                }
                else
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
            var userid = GetSmallHolderUserID(id);
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
                command.Parameters.AddWithValue("@id", userid);
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
                command.Parameters.AddWithValue("@FuturesPrice", 1.7);
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
                            command.Parameters.AddWithValue("@CoopID", userInfoDTO.CoopID);
                            command.Parameters.AddWithValue("@UserName", userInfoDTO.UserName);
                            command.Parameters.AddWithValue("@Language", userInfoDTO.Language);
                            command.Connection = connect;
                            resultUser = command.ExecuteNonQuery();
                            connect.Close();

                        }
                    }
                    else
                    {
                        string sqlQueryUser = String.Format("Update [AVFCoffee].[dbo].[User]" +
                   "Set [Language] = @language, [CoopID] = @coop Where UserID = @userid");
                        //var conn = _iconfiguration.GetSection("ConnectionStrings").GetSection("CoffeeConnStr").Value;
                        using (SqlConnection connect = new SqlConnection(conn))
                        {
                            connect.Open();
                            SqlCommand command = new SqlCommand(sqlQueryUser);
                            command.Parameters.AddWithValue("@language", userInfoDTO.Language);
                            command.Parameters.AddWithValue("@userid", userInfoDTO.UserID);
                            command.Parameters.AddWithValue("@coop", userInfoDTO.CoopID);
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

        public UserInfoDTO UpdateUserInfo(UserInfoDTO userInfoDTO)
        {
            throw new NotImplementedException();
        }
    }
}
