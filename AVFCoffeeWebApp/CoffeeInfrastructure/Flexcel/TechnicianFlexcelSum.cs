using CoffeeCore.DTO;
using CoffeeCore.Interfaces;
using FlexCel.Core;
using FlexCel.XlsAdapter;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data.SqlClient;
using CoffeeInfrastructure.Helpers;

namespace CoffeeInfrastructure.Flexcel
{
    public class TechnicianFlexcelSum : ITechnicianFlexcelSum
    {

        private readonly IConfiguration _iconfiguration;

        public TechnicianFlexcelSum(IConfiguration configuration)
        {
            _iconfiguration = configuration;
        }

        public TechnicianLoginInfoDTO GetUserMetrics()
        {
            
            TechnicianLoginInfoDTO tlInfo = new TechnicianLoginInfoDTO();
            MetricsInputDTO minput = new MetricsInputDTO();
            var conn = _iconfiguration.GetSection("ConnectionStrings").GetSection("CoffeeConnStr").Value;
            using (SqlConnection con = new SqlConnection(conn))
            {
                con.Open();

                SqlCommand comm = new SqlCommand("Select * from [AVFCoffee].[dbo].[CoopGeneralConfig] where CoopID = @coopid", con);
                comm.Parameters.AddWithValue("@coopid", "1234");
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
                }
                Dictionary<String, object> outDict = new Dictionary<String, object>();
                outDict.Add("Metrics", minput);
                tlInfo.technicianloginfo = outDict;
                con.Close();
                return tlInfo;
            }
        }


        public void saveUserMetrics(String[] data)
        {
            string language = data[0].ToString();
            //MetricsInputDTO metricsInputDTO = new MetricsInputDTO();
            var coffeemeasurekilograms = false;
            var coffeemeasurepounds = false;
            var coffeemeasurequintales = false;
            var coffeemeasurearrobas = false;
            var coffeemeasurecargas = false;
            var lengthmeasuremeters = false;
            var lengthmeasurefeet = false;
            var farmareameasurehectares = false;
            var farmareameasuremanzanas = false;
            var applicationmeasurekilograms = false;
            var applicationmeasurepounds = false;
            var capacitymeasureliters = false;
            var capacitymeasuregallons = false;
            var currencyboliviaboliviano = false;
            var currencybrazilreal = false;
            var currencycolombiapeso = false;
            var currencycostaricacolon = false;
            var currencycubapeso = false;
            var currencyguatemalaquetzal = false;
            var currencyhaitigourde = false;
            var currencyhonduraslempira = false;
            var currencyjamaicadollar = false;
            var currencymexicopeso = false;
            var currencynicaraguacordoba = false;
            var currencyperusol = false;
            var currencyusdollar = false;
            var currencyvenezuelabolivar = false;
            MetricsDTO md = new MetricsDTO();
            md.coffeemeasurekilograms = 0;
            md.coffeemeasurepounds = 0;
            md.coffeemeasurequintales = 0;
            md.currencyboliviaboliviano = 0;
            md.applicationmeasurekilograms = 0;
            md.applicationmeasurepounds = 0;
            md.capacitymeasuregallons = 0;
            md.capacitymeasureliters = 0;
            md.coffeemeasurearrobas = 0;
            md.coffeemeasurecargas = 0;
            md.currencybrazilreal = 0;
            md.currencycolombiapeso = 0;
            md.currencycostaricacolon = 0;
            md.currencycubapeso = 0;
            md.currencyguatemalaquetzal = 0;
            md.currencyhaitigourde = 0;
            md.currencyhonduraslempira = 0;
            md.currencyjamaicadollar = 0;
            md.currencymexicopeso = 0;
            md.currencynicaraguacordoba = 0;
            md.currencyperusol = 0;
            md.currencyusdollar = 0;
            md.currencyvenezuelabolivar = 0;
            md.farmareameasurehectares = 0;
            md.farmareameasuremanzanas = 0;
            md.lengthmeasurefeet = 0;
            md.lengthmeasuremeters = 0;
            if (data[1] == "Kilograms")
            {
                coffeemeasurekilograms = true;
                md.coffeemeasurekilograms = 1;
            } 
            else if (data[1] == "Pounds")
            {
                coffeemeasurepounds = true;
                md.coffeemeasurepounds = 1;
            }
            else if (data[1] == "Quintales")
            {
                coffeemeasurequintales = true;
                md.coffeemeasurequintales = 1;
            }
            else if (data[1] == "Arrobas")
            {
                coffeemeasurearrobas = true;
                md.coffeemeasurearrobas = 1;
            }
            else if (data[1] == "Cargas")
            {
                coffeemeasurecargas = true;
                md.coffeemeasurecargas = 1;
            }

            if (data[2] == "Meters")
            {
                lengthmeasuremeters = true;
                md.lengthmeasuremeters = 1;
            }
            else if (data[2] == "Feet")
            {
                lengthmeasurefeet = true;
                md.lengthmeasurefeet = 1;
            }

            if (data[3] == "Hectares")
            {
                farmareameasurehectares = true;
                md.farmareameasurehectares = 1;
            }
            else if (data[3] == "Manzanas")
            {
                farmareameasuremanzanas = true;
                md.farmareameasuremanzanas = 1;
            }

            if (data[4] == "Kilograms")
            {
                applicationmeasurekilograms = true;
                md.applicationmeasurekilograms = 1;
            }
            else if (data[4] == "Pounds")
            {
                applicationmeasurepounds = true;
                md.applicationmeasurepounds = 1;
            }

            if (data[5] == "Liters")
            {
                capacitymeasureliters = true;
                md.capacitymeasureliters = 1;
            }
            else if (data[5] == "Gallons")
            {
                capacitymeasuregallons = true;
                md.capacitymeasuregallons = 1;
            }


            if (data[6] == "Bolivian Boliviano")
            {
                currencyboliviaboliviano = true;
                md.currencyboliviaboliviano = 1;
            }
            else if (data[6] == "Brazilian Real")
            {
                currencybrazilreal = true;
                md.currencybrazilreal = 1;
            }
            else if (data[6] == "Colombian Peso")
            {
                currencycolombiapeso = true;
                md.currencycolombiapeso = 1;
            }
            else if (data[6] == "Costa Rican Colon")
            {
                currencycostaricacolon = true;
                md.currencycostaricacolon = 1;
            }
            else if (data[6] == "Cuban Peso")
            {
                currencycubapeso = true;
                md.currencycubapeso = 1;
            }
            else if (data[6] == "Guatemalan Quetzal")
            {
                currencyguatemalaquetzal = true;
                md.currencyguatemalaquetzal = 1;
            }
            else if (data[6] == "Jamaican Dollar")
            {
                currencyjamaicadollar = true;
                md.currencyjamaicadollar = 1;
            }
            else if (data[6] == "Honduran Lempira")
            {
                currencyhonduraslempira = true;
                md.currencyhonduraslempira = 1;
            }
            else if (data[6] == "Haitian Gourde")
            {
                currencyhaitigourde = true;
                md.currencyhaitigourde = 1;
            }
            else if (data[6] == "Mexican Peso")
            {
                currencymexicopeso = true;
                md.currencymexicopeso = 1;
            }
            else if (data[6] == "Nicaraguan Cordoba")
            {
                currencynicaraguacordoba = true;
                md.currencynicaraguacordoba = 1;
            }
            else if (data[6] == "Peruvian Sol")
            {
                currencyperusol = true;
                md.currencyperusol = 1;
            }
            else if (data[6] == "USD")
            {
                currencyusdollar = true;
                md.currencyusdollar = 1;
            }
            else if (data[6] == "Venezuelan Bolivar")
            {
                currencyvenezuelabolivar = true;
                md.currencyvenezuelabolivar = 1;
            }


            var conn = _iconfiguration.GetSection("ConnectionStrings").GetSection("CoffeeConnStr").Value;
            using (SqlConnection con = new SqlConnection(conn))
            {
                con.Open();

                SqlCommand comm = new SqlCommand("Update [AVFCoffee].[dbo].[CoopGeneralConfig] set [CoffeeMeasureKilograms] = @coffeekilo" +
      ",[CoffeeMeasurePounds] = @coffeepound,[CoffeeMeasureQuintales] = @coffeequintal,[CoffeeMeasureArrobas] = @coffeearrobas" +
      ",[CoffeeMeasureCargas] = @coffeecargas,[LengthMeasureMeters] = @lengthmeter ,[LengthMeasureFeet] = @lengthfeet ,[FarmAreaMeasureHectares] = @farmhectare" +
      ",[FarmAreaMeasureManzanas] = @farmmanzana ,[ApplicationMeasureKilograms] = @appkilo, [ApplicationMeasurePounds] = @apppound" +
      ",[CapacityMeasureLiters] = @capacityliter, [CapacityMeasureGallons] = @capacitygallon, [CurrencyBoliviaBoliviano] = @currbol" +
      ",[CurrencyBrazilReal] = @currbra, [CurrencyColombiaPeso] = @currcol, [CurrencyCostaRicaColon] = @currcr, [CurrencyCubaPeso] = @currcuba" +
      ",[CurrencyGuatemalaQuetzal] = @currguat, [CurrencyHaitiGourde] = @currhaiti, [CurrencyHondurasLempira] = @currhond, [CurrencyJamaicaDollar] = @currjam" +
      ",[CurrencyMexicoPeso] = @currmex, [CurrencyNicaraguaCordoba] = @currnica, [CurrencyPeruSol] = @currperu, [CurrencyUSDollar] = @currus" +
      ",[CurrencyVenezuelaBolivar] = @currvene where CoopID = @coopid", con);
                comm.Parameters.AddWithValue("@coopid", "1234");
                comm.Parameters.AddWithValue("@coffeekilo", coffeemeasurekilograms);
                comm.Parameters.AddWithValue("@coffeepound", coffeemeasurepounds);
                comm.Parameters.AddWithValue("@coffeequintal", coffeemeasurequintales);
                comm.Parameters.AddWithValue("@coffeecargas", coffeemeasurecargas);
                comm.Parameters.AddWithValue("@coffeearrobas", coffeemeasurearrobas);
                comm.Parameters.AddWithValue("@lengthmeter", lengthmeasuremeters);
                comm.Parameters.AddWithValue("@lengthfeet", lengthmeasurefeet);
                comm.Parameters.AddWithValue("@farmhectare", farmareameasurehectares);
                comm.Parameters.AddWithValue("@farmmanzana", farmareameasuremanzanas);
                comm.Parameters.AddWithValue("@appkilo", applicationmeasurekilograms);
                comm.Parameters.AddWithValue("@apppound", applicationmeasurepounds);
                comm.Parameters.AddWithValue("@capacityliter", capacitymeasureliters);
                comm.Parameters.AddWithValue("@capacitygallon", capacitymeasuregallons);
                comm.Parameters.AddWithValue("@currbol", currencyboliviaboliviano);
                comm.Parameters.AddWithValue("@currbra", currencybrazilreal);
                comm.Parameters.AddWithValue("@currcol", currencycolombiapeso);
                comm.Parameters.AddWithValue("@currcr", currencycostaricacolon);
                comm.Parameters.AddWithValue("@currcuba", currencycubapeso);
                comm.Parameters.AddWithValue("@currguat", currencyguatemalaquetzal);
                comm.Parameters.AddWithValue("@currhaiti", currencyhaitigourde);
                comm.Parameters.AddWithValue("@currhond", currencyhonduraslempira);
                comm.Parameters.AddWithValue("@currjam", currencyjamaicadollar);
                comm.Parameters.AddWithValue("@currmex", currencymexicopeso);
                comm.Parameters.AddWithValue("@currnica", currencynicaraguacordoba);
                comm.Parameters.AddWithValue("@currperu", currencyperusol);
                comm.Parameters.AddWithValue("@currus", currencyusdollar);
                comm.Parameters.AddWithValue("@currvene", currencyvenezuelabolivar);
                int result = comm.ExecuteNonQuery();
            }
            UpdateAdvancedInputs(md, language);
            //throw new NotImplementedException();
        }
        private void UpdateAdvancedInputs(MetricsDTO metricsDTO, string Language)
        {
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
            workspace.Add(xls.ActiveFileName, xls);

            //databaseSchema.Database_Schema(xls, workspace);
            ////xls.Recalc();
            //conf_Summary_Spa.GeneralConfSummarySpa(xls);
            ////xls.Recalc();
            //inputs_1.CreateFile(xls);
            ////xls.Recalc();
            //inputs_2_Conv.Inputs_2_Conv_inputs(xls);
            ////xls.Recalc();

            //inputs_TOT_advanced.CreateFile(xls);
            ////xls.Recalc();
            //prporcion.ProporcionDeProductividad(xls);
            ////xls.Recalc();
            //proportions.proportions(xls);
            ////xls.Recalc();
            //advancedInputs.Budget_Supuestos(xls);
            ////xls.Recalc();
            //budget_Equipo.BudgetEquipo(xls);
            ////xls.Recalc();
            //budget_Establecimiento.BudgetEstablecimiento(xls);
            ////xls.Recalc();
            //budget_M_Obra.BudgetMObra(xls);
            ////xls.Recalc();
            //budget_Presupuesto.BudgetPresupuesto(xls);
            ////xls.Recalc();
            //budget_Sostenemiento.BudgetSostenemiento(xls);
            ////xls.Recalc();
            //budget_Valor_De_M_Obra.Budget_Valor_M_De_Obra(xls);
            ////xls.Recalc();
            metrics.metrics(xls, metricsDTO);
            language.language(xls);
            inputs_1.CreateFile(xls);
            //xls.Recalc();
            metrics_English.MetricsEnglish(xls);
            //xls.Recalc();
            metrics_Spanish.MetricsSpanish(xls);
            inputs_2.Inputs_2_Default(xls);
            databaseSchema.Database_Schema(xls);
            inputs_2_Conv.Inputs_2_Conv_inputs(xls);
            inputsAdvanced2English.InputAdvanced2English(xls);
            //xls.Recalc();
            if (Language == "EN")
            {
                inputsAdvanced2Spanish.InputAdvancedSpanish(xls, Language);
            }
            inputsEnglish.InputEnglish(xls);
            inputsSpanish.InputSpanish(xls);
            input_1.Input_1_default(xls);
            //xls.Recalc();


            //xls.Recalc();
            //outcomeLAdjustment.Outcome_L_Adjustment(xls);
            //xls.Recalc();
            //outcomeYAdjustment.Outcome_Y_Adjustment(xls);
            //xls.Recalc();
            //output1_Pre_Metric_Currency.Output1PreMetricCurrency(xls);
            //xls.Recalc();
            //outcomeTotalAdj.Outcome_TOTAL_Adj(xls);
            //xls.Recalc();
            inputs_1_Ref.inputs1Ref(xls);
            //xls.Recalc();
            conversiones.conversiones(xls);
            //xls.Recalc(true);


            //xls.Recalc();

            //xls.Recalc();
            gral_Conf.Gral_Conf_Summary(xls);
            conf_Summary_Spa.GeneralConfSummarySpa(xls);
            //xls.Recalc();
            //return new Dictionary<string, object>();


            double earlyHectares = 0;
            double peakHectares = 0;
            double oldHectares = 0;
            bool conventional = false;
            bool organic = false;
            bool transition = false;
            double workerSalarySoles = 0;
            double productionQuintales = 0;
            double costPriceSolesPerQuintal = 0;
            double expSolesChem = 0;
            double expSolesOrg = 0;
            double transportCostSoles = 0;
            inputs.inputs(xls, earlyHectares, peakHectares, oldHectares, conventional, organic, transition, workerSalarySoles, productionQuintales, transportCostSoles,
                costPriceSolesPerQuintal, expSolesChem, expSolesOrg);
            var advancedInputsDict = new Dictionary<string, object>();

            if (Language == "EN")
            {

                advancedInputsDict = inAdvanced.Inputs_Advanced(xls);

            }
            else
            {
                advancedInputsDict = inputsAdvanced2Spanish.InputAdvancedSpanish(xls, Language);
            }

            var time = "";
            var conn1 = _iconfiguration.GetSection("ConnectionStrings").GetSection("CoffeeConnStr").Value;
            using (SqlConnection con = new SqlConnection(conn1))
            {
                con.Open();
                var id = "1234";

                SqlCommand comm = new SqlCommand("Select * from [AVFCoffee].[dbo].[UserInputsAdvanced] where UserID = @userid AND [TimeStamp] = (SELECT MAX(timestamp) FROM[AVFCoffee].[dbo].[UserInputsAdvanced] where UserID = @userid)", con);
                comm.Parameters.AddWithValue("@userid", id);
                // int result = command.ExecuteNonQuery();
                using (SqlDataReader reader = comm.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        time = reader["TimeStamp"].ToString();
                    }
                    reader.Close();
                }

                con.Close();
            }
            var conn2 = _iconfiguration.GetSection("ConnectionStrings").GetSection("CoffeeConnStr").Value;
            string sqlQuery = String.Format("Update [AVFCoffee].[dbo].[UserInputsAdvanced] " +
               "set [LGerminationSeedCollection] = @LGerminationSeedCollection, " +
                    "[LGerminationSeedSelection] = @LGerminationSeedSelection, " +
            "[LGerminationNurseryConstruction] = @LGerminationNurseryConstruction, " +
            "[LGerminationSeedingSupportIrrigation] = @LGerminationSeedingSupportIrrigation, " +
            "[LGerminationOthers] = @LGerminationOthers, " +
            "[LNurseryConstruction] = @LNurseryConstruction, " +
            "[LNurseryDrawnPulled] = @LNurseryDrawnPulled, " +
            "[LNurseryClean] = @LNurseryClean, " +
            "[LNurserySoilPreparationFertilizer] = @LNurserySoilPreparationFertilizer, " +
            "[LNurseryFilledLockedBags] = @LNurseryFilledLockedBags, " +
            "[LNurseryButterflySowing] = @LNurseryButterflySowing, " +
            "[LNurseryIrrigation] = @LNurseryIrrigation, " +
            "[LNurseryFoliarApplication] = @LNurseryFoliarApplication, " +
            "[LNurseryReseeding] = @LNurseryReseeding, " +
            "[LNurseryOthers] = @LNurseryOthers, " +
            "[LPPFieldCleaning] = @LPPFieldCleaning, " +
            "[LPPCuttingTrees] = @LPPCuttingTrees, " +
            "[LPPWoodCollection] = @LPPWoodCollection, " +
            "[LPPWoodChopping] = @LPPWoodChopping, " +
            "[LPPCoffeeLayout] = @LPPCoffeeLayout, " +
            "[LPPHoleDigging] = @LPPHoleDigging, " +
            "[LPPSeedlingTransportation] = @LPPSeedlingTransportation, " +
            "[LPPSeedlingTransplant] = @LPPSeedlingTransplant, " +
            "[LPPShadeAdjustment] = @LPPShadeAdjustment, " +
            "[LPPCompostMixing] = @LPPCompostMixing, " +
            "[LPPOthers] = @LPPOthers, " +
            "[LPPYWeeding] = @LPPYWeeding, " +
            "[LPPYOrganic] = @LPPYOrganic, " +
            "[LPPYChemical] = @LPPYChemical, " +
            "[LPPYFoliarSpraying] = @LPPYFoliarSpraying, " +
            "[LPPYOther] = @LPPYOther, " +
            "[LHPMYManualWeeding] = @LHPMYManualWeeding, " +
            "[LHPMYChemicalWeeding] = @LHPMYChemicalWeeding, " +
            "[LHPMYOrganicFertilizers] = @LHPMYOrganicFertilizers, " +
            "[LHPMYChemicalFertilizers] = @LHPMYChemicalFertilizers, " +
            "[LHPMYFoliarSpraying] = @LHPMYFoliarSpraying, " +
            "[LHPMYHedgerowsConstruction] = @LHPMYHedgerowsConstruction, " +
            "[LHPMYShadetreePruning] = @LHPMYShadetreePruning, " +
            "[LHPMYPestControl] = @LHPMYPestControl, " +
            "[LHPMYCoffeeGrowManagement] = @LHPMYCoffeeGrowManagement, " +
            "[LHPMYOthers] = @LHPMYOthers, " +
            "[LHPHYCoffeeCollecDays] = @LHPHYCoffeeCollecDays, " +
            "[LHPHYAdditionDays] = @LHPHYAdditionDays, " +
            "[LHPPYFermentation] = @LHPPYFermentation, " +
            "[LHPPYWashing] = @LHPPYWashing, " +
            "[LHPPYDrying] = @LHPPYDrying, " +
            "[LHPPYScreening] = @LHPPYScreening, " +
            "[LHPPYSelection] = @LHPPYSelection, " +
            "[LHPPYStorage] = @LHPPYStorage, " +
            "[LHPPYCoffeewastewater] = @LHPPYCoffeewastewater, " +
            "[LHPPYPulpManagement] = @LHPPYPulpManagement, " +
            "[LHPPYOthers] = @LHPPYOthers, " +
            "[LHPMMManualWeeding] = @LHPMMManualWeeding, " +
            "[LHPMMChemicalWeeding] = @LHPMMChemicalWeeding, " +
            "[LHPMMOrganicFertilizers] = @LHPMMOrganicFertilizers, " +
            "[LHPMMChemicalFertilizers] = @LHPMMChemicalFertilizers, " +
            "[LHPMMFoliarSpraying] = @LHPMMFoliarSpraying, " +
            "[LHPMMHedgerowsConstruction] = @LHPMMHedgerowsConstruction, " +
            "[LHPMMShadetreePruning] = @LHPMMShadetreePruning, " +
            "[LHPMMPestControl] = @LHPMMPestControl, " +
            "[LHPMMCoffeeGrowManagement] = @LHPMMCoffeeGrowManagement, " +
            "[LHPMMOthers] = @LHPMMOthers, " +
            "[LHPHMCoffeeCollecDays] = @LHPHMCoffeeCollecDays, " +
            "[LHPHMAdditionDays] = @LHPHMAdditionDays, " +
            "[LHPPMFermentation] = @LHPPMFermentation, " +
            "[LHPPMWashing] = @LHPPMWashing, " +
            "[LHPPMDrying] = @LHPPMDrying, " +
            "[LHPPMScreening] = @LHPPMScreening, " +
            "[LHPPMSelection] = @LHPPMSelection, " +
            "[LHPPMStorage] = @LHPPMStorage, " +
            "[LHPPMCoffeewastewater] = @LHPPMCoffeewastewater, " +
            "[LHPPMPulpManagement] = @LHPPMPulpManagement, " +
            "[LHPPMOthers] = @LHPPMOthers, " +
            "[LHPMOManualWeeding] = @LHPMOManualWeeding, " +
            "[LHPMOChemicalWeeding] = @LHPMOChemicalWeeding, " +
            "[LHPMOOrganicFertilizers] = @LHPMOOrganicFertilizers, " +
            "[LHPMOChemicalFertilizers] = @LHPMOChemicalFertilizers, " +
            "[LHPMOFoliarSpraying] = @LHPMOFoliarSpraying, " +
            "[LHPMOHedgerowsConstruction] = @LHPMOHedgerowsConstruction, " +
            "[LHPMOShadetreePruning] = @LHPMOShadetreePruning, " +
            "[LHPMOPestControl] = @LHPMOPestControl, " +
            "[LHPMOCoffeeGrowManagement] = @LHPMOCoffeeGrowManagement, " +
            "[LHPMOOthers] = @LHPMOOthers, " +
            "[LHPHOCoffeeCollecDays] = @LHPHOCoffeeCollecDays, " +
            "[LHPHOAdditionDays] = @LHPHOAdditionDays, " +
            "[LHPPOFermentation] = @LHPPOFermentation, " +
            "[LHPPOWashing] = @LHPPOWashing, " +
            "[LHPPODrying] = @LHPPODrying, " +
            "[LHPPOScreening] = @LHPPOScreening, " +
            "[LHPPOSelection] = @LHPPOSelection, " +
            "[LHPPOStorage] = @LHPPOStorage, " +
            "[LHPPOCoffeewastewater] = @LHPPOCoffeewastewater, " +
            "[LHPPOPulpManagement] = @LHPPOPulpManagement, " +
            "[LHPPOOthers] = @LHPPOOthers, " +
            "[IIFood] = @IIFood, " +
            "[IIAdditionalTransfers] = @IIAdditionalTransfers, " +
            "[IIDaysoftraining] = @IIDaysoftraining, " +
            "[ICCreditfromcooperative] = @ICCreditfromcooperative, " +
            "[ICCreditfromcooperativeTime] = @ICCreditfromcooperativeTime, " +
            "[ICCreditfromcooperativeInterest] = @ICCreditfromcooperativeInterest, " +
            "[ICCreditfromagent] = @ICCreditfromagent, " +
            "[ICCreditfromagentTime] = @ICCreditfromagentTime, " +
            "[ICCreditfromagentInterest] = @ICCreditfromagentInterest, " +
            //"[CostGerminator] = @CostGerminator, " +
            "[CostGerminatorSeeds] = @CostGerminatorSeeds, " +
            "[CostGerminatorSeedbed] = @CostGerminatorSeedbed, " +
            "[CostGerminatorSandSubstrate] = @CostGerminatorSandSubstrate, " +
            "[CostGerminatorCalciumSulfide] = @CostGerminatorCalciumSulfide, " +
            "[CostGerminatorLime] = @CostGerminatorLime, " +
            "[CostGerminatorPlastic] = @CostGerminatorPlastic, " +
            "[CostGerminatorOthers] = @CostGerminatorOthers, " +
            "[CostNurseryFertilizer] = @CostNurseryFertilizer, " +
            "[CostNurseryPlasticBags] = @CostNurseryPlasticBags, " +
            "[CostNurseryNetting] = @CostNurseryNetting, " +
            "[CostNurseryStuds] = @CostNurseryStuds, " +
            "[CostNurseryWire] = @CostNurseryWire, " +
            "[CostNurseryCiclonics] = @CostNurseryCiclonics, " +
            "[CostNurseryStaples] = @CostNurseryStaples, " +
            "[CostNurserySoil] = @CostNurserySoil, " +
            "[CostNurseryBioFert] = @CostNurseryBioFert, " +
            "[CostNurseryAgroChemicals] = @CostNurseryAgroChemicals, " +
            "[CostNurseryFungicide] = @CostNurseryFungicide, " +
            "[CostNurseryPhosphoricRock] = @CostNurseryPhosphoricRock, " +
            "[CostNurseryOthers] = @CostNurseryOthers, " +
            "[CostFLPPOrganicFert] = @CostFLPPOrganicFert, " +
            "[CostFLPPChemicalFert] = @CostFLPPChemicalFert, " +
            "[CostFVGOrganicFert] = @CostFVGOrganicFert, " +
            "[CostFVGChemicalFert] = @CostFVGChemicalFert, " +
            "[CostFMOtherFert] = @CostFMOtherFert, " +
            "[CostFMOrganicFoliar] = @CostFMOrganicFoliar, " +
            "[CostFMChemicalFoliar] = @CostFMChemicalFoliar, " +
            "[CostFMGasFuel] = @CostFMGasFuel, " +
            "[CostFMOthers] = @CostFMOthers, " +
            "[EGEManualSprayer] = @EGEManualSprayer, " +
            "[EGELifespam1] = @EGELifespam1, " +
            "[EGEMachetes] = @EGEMachetes, " +
            "[EGELifespam2] = @EGELifespam2, " +
            "[EGEShovel] = @EGEShovel, " +
            "[EGELifespam3] = @EGELifespam3, " +
            "[EGEHoe] = @EGEHoe, " +
            "[EGELifespam4] = @EGELifespam4, " +
            "[EGEWheelBarrow] = @EGEWheelBarrow, " +
            "[EGELifespam5] = @EGELifespam5, " +
            "[EGELime] = @EGELime, " +
            "[EGELifespam6] = @EGELifespam6, " +
            "[EGEAuger] = @EGEAuger, " +
            "[EGELifespam7] = @EGELifespam7, " +
            "[EGEMetalBar] = @EGEMetalBar, " +
            "[EGELifespam8] = @EGELifespam8, " +
            "[EGEHose] = @EGEHose, " +
            "[EGELifespam9] = @EGELifespam9, " +
            "[EGESprinklers] = @EGESprinklers, " +
            "[EGELifespam10] = @EGELifespam10, " +
            "[EGEChainSaw] = @EGEChainSaw, " +
            "[EGELifespam11] = @EGELifespam11, " +
            "[EGEHandSaw] = @EGEHandSaw, " +
            "[EGELifespam12] = @EGELifespam12, " +
            "[EGEMotorPump] = @EGEMotorPump, " +
            "[EGELifespam13] = @EGELifespam13, " +
            "[EGEPrunningScissors] = @EGEPrunningScissors, " +
            "[EGELifespam14] = @EGELifespam14, " +
            "[EGEAxe] = @EGEAxe, " +
            "[EGELifespam15] = @EGELifespam15, " +
            "[EEHScale] = @EEHScale, " +
            "[EEHLifespam1] = @EEHLifespam1, " +
            "[EEHVehicle] = @EEHVehicle, " +
            "[EEHLifespam2] = @EEHLifespam2, " +
            "[EEHWorkAnimal] = @EEHWorkAnimal, " +
            "[EEHLifespam3] = @EEHLifespam3, " +
            "[EEHMotorcycle] = @EEHMotorcycle, " +
            "[EEHLifespam4] = @EEHLifespam4, " +
            "[EEHBags] = @EEHBags, " +
            "[EEHLifespam5] = @EEHLifespam5, " +
            "[EEHSack] = @EEHSack, " +
            "[EEHLifespam6] = @EEHLifespam6, " +
            "[EEHStraw] = @EEHStraw, " +
            "[EEHLifespam7] = @EEHLifespam7, " +
            "[EEHBaskets] = @EEHBaskets, " +
            "[EEHLifespam8] = @EEHLifespam8, " +
            "[EEHBoxes] = @EEHBoxes, " +
            "[EEHLifespam9] = @EEHLifespam9, " +
            "[EEHOthers] = @EEHOthers, " +
            "[EEHLifespam10] = @EEHLifespam10, " +
            "[EEPPulperMachine] = @EEPPulperMachine, " +
            "[EEPLifespam1] = @EEPLifespam1, " +
            "[EEPTolca] = @EEPTolca, " +
            "[EEPLifespam2] = @EEPLifespam2, " +
            "[EEPEngine] = @EEPEngine, " +
            "[EEPLifespam3] = @EEPLifespam3, " +
            "[EEPTanks] = @EEPTanks, " +
            "[EEPLifespam4] = @EEPLifespam4, " +
            "[EEPWaterChannel] = @EEPWaterChannel, " +
            "[EEPLifespam5] = @EEPLifespam5, " +
            "[EEPPVCPipes] = @EEPPVCPipes, " +
            "[EEPLifespam6] = @EEPLifespam6, " +
            "[EEPFilteringSystem] = @EEPFilteringSystem, " +
            "[EEPLifespam7] = @EEPLifespam7, " +
            "[EEPScreeningMachine] = @EEPScreeningMachine, " +
            "[EEPLifespam8] = @EEPLifespam8, " +
            "[EEPDesmucilaginador] = @EEPDesmucilaginador, " +
            "[EEPLifespam9] = @EEPLifespam9, " +
            "[EEPMotorPump] = @EEPMotorPump, " +
            "[EEPLifespam10] = @EEPLifespam10, " +
            "[EEPOthersWetInput] = @EEPOthersWetInput, " +
            "[EEPLifespam11] = @EEPLifespam11, " +
            "[EEPConcrete] = @EEPConcrete, " +
            "[EEPLifespam12] = @EEPLifespam12, " +
            "[EEPPlastic] = @EEPPlastic, " +
            "[EEPLifespam13] = @EEPLifespam13, " +
            "[EEPRake] = @EEPRake, " +
            "[EEPLifespam14] = @EEPLifespam14, " +
            "[EEPBroom] = @EEPBroom, " +
            "[EEPLifespam15] = @EEPLifespam15, " +
            "[EEPStorageRoom] = @EEPStorageRoom, " +
            "[EEPLifespam16] = @EEPLifespam16, " +
            "[EEPOthersDryInput] = @EEPOthersDryInput, " +
            "[EEPLifespam17] = @EEPLifespam17, " +
            "[ACCApplicationFee] = @ACCApplicationFee, " +
            "[ACCAnnualMembership] = @ACCAnnualMembership, " +
            "[ACCLifeInsurance] = @ACCLifeInsurance, " +
            "[ACCFLOCertification] = @ACCFLOCertification, " +
            "[ACCOrganicCertification] = @ACCOrganicCertification, " +
            "[ACLLandValue] = @ACLLandValue, " +
            "[ACLPropertyTax] = @ACLPropertyTax, " +
            "[ACUSuperviseInvest] = @ACUSuperviseInvest, " +
            "[ACUAdministInvest] = @ACUAdministInvest, " +
            "[ACUTrainingInvest] = @ACUTrainingInvest, " +
            "[ACUExtraOrdInvest] = @ACUExtraOrdInvest, " +
            "[TGSeedPurchase] = @TGSeedPurchase, " +
            "[TGWoodTransportation] = @TGWoodTransportation, " +
            "[TGSandTransportation] = @TGSandTransportation, " +
            "[TGOthers] = @TGOthers, " +
            "[TNSoilTransportation] = @TNSoilTransportation, " +
            "[TNSacksMaterialShopping] = @TNSacksMaterialShopping, " +
            "[TNOthers] = @TNOthers, " +
            "[TLPWoodTransportation] = @TLPWoodTransportation, " +
            "[TLPCompostTransportation] = @TLPCompostTransportation, " +
            "[TLPPlantTransportation] = @TLPPlantTransportation, " +
            "[TLPOthers] = @TLPOthers, " +
            "[TOtherEquipment] = @TOtherEquipment, " +
            "[TOtherLaborTransportation] = @TOtherLaborTransportation, " +
            "[TOtherCoffeeTransportation] = @TOtherCoffeeTransportation, " +
            "[TOtherSupervisingActivities] = @TOtherSupervisingActivities, " +
            "[TOthers] = @TOthers WHERE [UserID] = @UserID and [TimeStamp] = @time");
            IList listInputs = (IList)advancedInputsDict["Inputs"];
            using (SqlConnection connect = new SqlConnection(conn2))
            {
                connect.Open();
                SqlCommand command = new SqlCommand(sqlQuery, connect);
                command.Parameters.AddWithValue("@UserID", "1234");
                command.Parameters.AddWithValue("@time", time);
                command.Parameters.AddWithValue("@LGerminationSeedCollection", listInputs[0]);
                command.Parameters.AddWithValue("@LGerminationSeedSelection", listInputs[1]);
                command.Parameters.AddWithValue("@LGerminationNurseryConstruction", listInputs[2]);
                command.Parameters.AddWithValue("@LGerminationSeedingSupportIrrigation", listInputs[3]);
                command.Parameters.AddWithValue("@LGerminationOthers", listInputs[4]);
                command.Parameters.AddWithValue("@LNurseryConstruction", listInputs[5]);
                command.Parameters.AddWithValue("@LNurseryDrawnPulled", listInputs[6]);
                command.Parameters.AddWithValue("@LNurseryClean", listInputs[7]);
                command.Parameters.AddWithValue("@LNurserySoilPreparationFertilizer", listInputs[8]);
                command.Parameters.AddWithValue("@LNurseryFilledLockedBags", listInputs[9]);
                command.Parameters.AddWithValue("@LNurseryButterflySowing", listInputs[10]);
                command.Parameters.AddWithValue("@LNurseryIrrigation", listInputs[11]);
                command.Parameters.AddWithValue("@LNurseryFoliarApplication", listInputs[12]);
                command.Parameters.AddWithValue("@LNurseryReseeding", listInputs[13]);
                command.Parameters.AddWithValue("@LNurseryOthers", listInputs[14]);
                command.Parameters.AddWithValue("@LPPFieldCleaning", listInputs[15]);
                command.Parameters.AddWithValue("@LPPCuttingTrees", listInputs[16]);
                command.Parameters.AddWithValue("@LPPWoodCollection", listInputs[17]);
                command.Parameters.AddWithValue("@LPPWoodChopping", listInputs[18]);
                command.Parameters.AddWithValue("@LPPCoffeeLayout", listInputs[19]);
                command.Parameters.AddWithValue("@LPPHoleDigging", listInputs[20]);
                command.Parameters.AddWithValue("@LPPSeedlingTransportation", listInputs[21]);
                command.Parameters.AddWithValue("@LPPSeedlingTransplant", listInputs[22]);
                command.Parameters.AddWithValue("@LPPShadeAdjustment", listInputs[23]);
                command.Parameters.AddWithValue("@LPPCompostMixing", listInputs[24]);
                command.Parameters.AddWithValue("@LPPOthers", listInputs[25]);
                command.Parameters.AddWithValue("@LPPYWeeding", listInputs[26]);
                command.Parameters.AddWithValue("@LPPYOrganic", listInputs[27]);
                command.Parameters.AddWithValue("@LPPYChemical", listInputs[28]);
                command.Parameters.AddWithValue("@LPPYFoliarSpraying", listInputs[29]);
                command.Parameters.AddWithValue("@LPPYOther", listInputs[30]);
                command.Parameters.AddWithValue("@LHPMYManualWeeding", listInputs[31]);
                command.Parameters.AddWithValue("@LHPMYChemicalWeeding", listInputs[32]);
                command.Parameters.AddWithValue("@LHPMYOrganicFertilizers", listInputs[33]);
                command.Parameters.AddWithValue("@LHPMYChemicalFertilizers", listInputs[34]);
                command.Parameters.AddWithValue("@LHPMYFoliarSpraying", listInputs[35]);
                command.Parameters.AddWithValue("@LHPMYHedgerowsConstruction", listInputs[36]);
                command.Parameters.AddWithValue("@LHPMYShadetreePruning", listInputs[37]);
                command.Parameters.AddWithValue("@LHPMYPestControl", listInputs[38]);
                command.Parameters.AddWithValue("@LHPMYCoffeeGrowManagement", listInputs[39]);
                command.Parameters.AddWithValue("@LHPMYOthers", listInputs[40]);
                command.Parameters.AddWithValue("@LHPHYCoffeeCollecDays", listInputs[41]);
                command.Parameters.AddWithValue("@LHPHYAdditionDays", listInputs[42]);
                command.Parameters.AddWithValue("@LHPPYFermentation", listInputs[43]);
                command.Parameters.AddWithValue("@LHPPYWashing", listInputs[44]);
                command.Parameters.AddWithValue("@LHPPYDrying", listInputs[45]);
                command.Parameters.AddWithValue("@LHPPYScreening", listInputs[46]);
                command.Parameters.AddWithValue("@LHPPYSelection", listInputs[47]);
                command.Parameters.AddWithValue("@LHPPYStorage", listInputs[48]);
                command.Parameters.AddWithValue("@LHPPYCoffeewastewater", listInputs[49]);
                command.Parameters.AddWithValue("@LHPPYPulpManagement", listInputs[50]);
                command.Parameters.AddWithValue("@LHPPYOthers", listInputs[51]);
                command.Parameters.AddWithValue("@LHPMMManualWeeding", listInputs[52]);
                command.Parameters.AddWithValue("@LHPMMChemicalWeeding", listInputs[53]);
                command.Parameters.AddWithValue("@LHPMMOrganicFertilizers", listInputs[54]);
                command.Parameters.AddWithValue("@LHPMMChemicalFertilizers", listInputs[55]);
                command.Parameters.AddWithValue("@LHPMMFoliarSpraying", listInputs[56]);
                command.Parameters.AddWithValue("@LHPMMHedgerowsConstruction", listInputs[57]);
                command.Parameters.AddWithValue("@LHPMMShadetreePruning", listInputs[58]);
                command.Parameters.AddWithValue("@LHPMMPestControl", listInputs[59]);
                command.Parameters.AddWithValue("@LHPMMCoffeeGrowManagement", listInputs[60]);
                command.Parameters.AddWithValue("@LHPMMOthers", listInputs[61]);
                command.Parameters.AddWithValue("@LHPHMCoffeeCollecDays", listInputs[62]);
                command.Parameters.AddWithValue("@LHPHMAdditionDays", listInputs[63]);
                command.Parameters.AddWithValue("@LHPPMFermentation", listInputs[64]);
                command.Parameters.AddWithValue("@LHPPMWashing", listInputs[65]);
                command.Parameters.AddWithValue("@LHPPMDrying", listInputs[66]);
                command.Parameters.AddWithValue("@LHPPMScreening", listInputs[67]);
                command.Parameters.AddWithValue("@LHPPMSelection", listInputs[68]);
                command.Parameters.AddWithValue("@LHPPMStorage", listInputs[69]);
                command.Parameters.AddWithValue("@LHPPMCoffeewastewater", listInputs[70]);
                command.Parameters.AddWithValue("@LHPPMPulpManagement", listInputs[71]);
                command.Parameters.AddWithValue("@LHPPMOthers", listInputs[72]);
                command.Parameters.AddWithValue("@LHPMOManualWeeding", listInputs[73]);
                command.Parameters.AddWithValue("@LHPMOChemicalWeeding", listInputs[74]);
                command.Parameters.AddWithValue("@LHPMOOrganicFertilizers", listInputs[75]);
                command.Parameters.AddWithValue("@LHPMOChemicalFertilizers", listInputs[76]);
                command.Parameters.AddWithValue("@LHPMOFoliarSpraying", listInputs[77]);
                command.Parameters.AddWithValue("@LHPMOHedgerowsConstruction", listInputs[78]);
                command.Parameters.AddWithValue("@LHPMOShadetreePruning", listInputs[79]);
                command.Parameters.AddWithValue("@LHPMOPestControl", listInputs[80]);
                command.Parameters.AddWithValue("@LHPMOCoffeeGrowManagement", listInputs[81]);
                command.Parameters.AddWithValue("@LHPMOOthers", listInputs[82]);
                command.Parameters.AddWithValue("@LHPHOCoffeeCollecDays", listInputs[83]);
                command.Parameters.AddWithValue("@LHPHOAdditionDays", listInputs[84]);
                command.Parameters.AddWithValue("@LHPPOFermentation", listInputs[85]);
                command.Parameters.AddWithValue("@LHPPOWashing", listInputs[86]);
                command.Parameters.AddWithValue("@LHPPODrying", listInputs[87]);
                command.Parameters.AddWithValue("@LHPPOScreening", listInputs[88]);
                command.Parameters.AddWithValue("@LHPPOSelection", listInputs[89]);
                command.Parameters.AddWithValue("@LHPPOStorage", listInputs[90]);
                command.Parameters.AddWithValue("@LHPPOCoffeewastewater", listInputs[91]);
                command.Parameters.AddWithValue("@LHPPOPulpManagement", listInputs[92]);
                command.Parameters.AddWithValue("@LHPPOOthers", listInputs[93]);
                command.Parameters.AddWithValue("@IIFood", listInputs[94]);
                command.Parameters.AddWithValue("@IIAdditionalTransfers", listInputs[95]);
                command.Parameters.AddWithValue("@IIDaysoftraining", listInputs[96]);
                command.Parameters.AddWithValue("@ICCreditfromcooperative", listInputs[97]);
                command.Parameters.AddWithValue("@ICCreditfromcooperativeTime", listInputs[98]);
                command.Parameters.AddWithValue("@ICCreditfromcooperativeInterest", listInputs[99]);
                command.Parameters.AddWithValue("@ICCreditfromagent", listInputs[100]);
                command.Parameters.AddWithValue("@ICCreditfromagentTime", listInputs[101]);
                command.Parameters.AddWithValue("@ICCreditfromagentInterest", listInputs[102]);
                //command.Parameters.AddWithValue.CostGerminator = Convert.ToDouble(reader["CostGerminator", listInputs[7]);
                command.Parameters.AddWithValue("@CostGerminatorSeeds", listInputs[103]);
                command.Parameters.AddWithValue("@CostGerminatorSeedbed", listInputs[104]);
                command.Parameters.AddWithValue("@CostGerminatorSandSubstrate", listInputs[105]);
                command.Parameters.AddWithValue("@CostGerminatorCalciumSulfide", listInputs[106]);
                command.Parameters.AddWithValue("@CostGerminatorLime", listInputs[107]);
                command.Parameters.AddWithValue("@CostGerminatorPlastic", listInputs[108]);
                command.Parameters.AddWithValue("@CostGerminatorOthers", listInputs[109]);
                command.Parameters.AddWithValue("@CostNurseryFertilizer", listInputs[110]);
                command.Parameters.AddWithValue("@CostNurseryPlasticBags", listInputs[111]);
                command.Parameters.AddWithValue("@CostNurseryNetting", listInputs[112]);
                command.Parameters.AddWithValue("@CostNurseryStuds", listInputs[113]);
                command.Parameters.AddWithValue("@CostNurseryWire", listInputs[114]);
                command.Parameters.AddWithValue("@CostNurseryCiclonics", listInputs[115]);
                command.Parameters.AddWithValue("@CostNurseryStaples", listInputs[116]);
                command.Parameters.AddWithValue("@CostNurserySoil", listInputs[117]);
                command.Parameters.AddWithValue("@CostNurseryBioFert", listInputs[118]);
                command.Parameters.AddWithValue("@CostNurseryAgroChemicals", listInputs[119]);
                command.Parameters.AddWithValue("@CostNurseryFungicide", listInputs[120]);
                command.Parameters.AddWithValue("@CostNurseryPhosphoricRock", listInputs[121]);
                command.Parameters.AddWithValue("@CostNurseryOthers", listInputs[122]);
                command.Parameters.AddWithValue("@CostFLPPOrganicFert", listInputs[123]);
                command.Parameters.AddWithValue("@CostFLPPChemicalFert", listInputs[124]);
                command.Parameters.AddWithValue("@CostFVGOrganicFert", listInputs[125]);
                command.Parameters.AddWithValue("@CostFVGChemicalFert", listInputs[126]);
                command.Parameters.AddWithValue("@CostFMOtherFert", listInputs[127]);
                command.Parameters.AddWithValue("@CostFMOrganicFoliar", listInputs[128]);
                command.Parameters.AddWithValue("@CostFMChemicalFoliar", listInputs[129]);
                command.Parameters.AddWithValue("@CostFMGasFuel", listInputs[130]);
                command.Parameters.AddWithValue("@CostFMOthers", listInputs[131]);
                command.Parameters.AddWithValue("@EGEManualSprayer", listInputs[132]);
                command.Parameters.AddWithValue("@EGELifespam1", listInputs[133]);
                command.Parameters.AddWithValue("@EGEMachetes", listInputs[134]);
                command.Parameters.AddWithValue("@EGELifespam2", listInputs[135]);
                command.Parameters.AddWithValue("@EGEShovel", listInputs[136]);
                command.Parameters.AddWithValue("@EGELifespam3", listInputs[137]);
                command.Parameters.AddWithValue("@EGEHoe", listInputs[138]);
                command.Parameters.AddWithValue("@EGELifespam4", listInputs[139]);
                command.Parameters.AddWithValue("@EGEWheelBarrow", listInputs[140]);
                command.Parameters.AddWithValue("@EGELifespam5", listInputs[141]);
                command.Parameters.AddWithValue("@EGELime", listInputs[142]);
                command.Parameters.AddWithValue("@EGELifespam6", listInputs[143]);
                command.Parameters.AddWithValue("@EGEAuger", listInputs[144]);
                command.Parameters.AddWithValue("@EGELifespam7", listInputs[145]);
                command.Parameters.AddWithValue("@EGEMetalBar", listInputs[146]);
                command.Parameters.AddWithValue("@EGELifespam8", listInputs[147]);
                command.Parameters.AddWithValue("@EGEHose", listInputs[148]);
                command.Parameters.AddWithValue("@EGELifespam9", listInputs[149]);
                command.Parameters.AddWithValue("@EGESprinklers", listInputs[150]);
                command.Parameters.AddWithValue("@EGELifespam10", listInputs[151]);
                command.Parameters.AddWithValue("@EGEChainSaw", listInputs[152]);
                command.Parameters.AddWithValue("@EGELifespam11", listInputs[153]);
                command.Parameters.AddWithValue("@EGEHandSaw", listInputs[154]);
                command.Parameters.AddWithValue("@EGELifespam12", listInputs[155]);
                command.Parameters.AddWithValue("@EGEMotorPump", listInputs[156]);
                command.Parameters.AddWithValue("@EGELifespam13", listInputs[157]);
                command.Parameters.AddWithValue("@EGEPrunningScissors", listInputs[158]);
                command.Parameters.AddWithValue("@EGELifespam14", listInputs[159]);
                command.Parameters.AddWithValue("@EGEAxe", listInputs[160]);
                command.Parameters.AddWithValue("@EGELifespam15", listInputs[161]);
                command.Parameters.AddWithValue("@EEHScale", listInputs[162]);
                command.Parameters.AddWithValue("@EEHLifespam1", listInputs[163]);
                command.Parameters.AddWithValue("@EEHVehicle", listInputs[164]);
                command.Parameters.AddWithValue("@EEHLifespam2", listInputs[165]);
                command.Parameters.AddWithValue("@EEHWorkAnimal", listInputs[166]);
                command.Parameters.AddWithValue("@EEHLifespam3", listInputs[167]);
                command.Parameters.AddWithValue("@EEHMotorcycle", listInputs[168]);
                command.Parameters.AddWithValue("@EEHLifespam4", listInputs[169]);
                command.Parameters.AddWithValue("@EEHBags", listInputs[170]);
                command.Parameters.AddWithValue("@EEHLifespam5", listInputs[171]);
                command.Parameters.AddWithValue("@EEHSack", listInputs[172]);
                command.Parameters.AddWithValue("@EEHLifespam6", listInputs[173]);
                command.Parameters.AddWithValue("@EEHStraw", listInputs[174]);
                command.Parameters.AddWithValue("@EEHLifespam7", listInputs[175]);
                command.Parameters.AddWithValue("@EEHBaskets", listInputs[176]);
                command.Parameters.AddWithValue("@EEHLifespam8", listInputs[177]);
                command.Parameters.AddWithValue("@EEHBoxes", listInputs[178]);
                command.Parameters.AddWithValue("@EEHLifespam9", listInputs[179]);
                command.Parameters.AddWithValue("@EEHOthers", listInputs[180]);
                command.Parameters.AddWithValue("@EEHLifespam10", listInputs[181]);
                command.Parameters.AddWithValue("@EEPPulperMachine", listInputs[182]);
                command.Parameters.AddWithValue("@EEPLifespam1", listInputs[183]);
                command.Parameters.AddWithValue("@EEPTolca", listInputs[184]);
                command.Parameters.AddWithValue("@EEPLifespam2", listInputs[185]);
                command.Parameters.AddWithValue("@EEPEngine", listInputs[186]);
                command.Parameters.AddWithValue("@EEPLifespam3", listInputs[187]);
                command.Parameters.AddWithValue("@EEPTanks", listInputs[188]);
                command.Parameters.AddWithValue("@EEPLifespam4", listInputs[189]);
                command.Parameters.AddWithValue("@EEPWaterChannel", listInputs[190]);
                command.Parameters.AddWithValue("@EEPLifespam5", listInputs[191]);
                command.Parameters.AddWithValue("@EEPPVCPipes", listInputs[192]);
                command.Parameters.AddWithValue("@EEPLifespam6", listInputs[193]);
                command.Parameters.AddWithValue("@EEPFilteringSystem", listInputs[194]);
                command.Parameters.AddWithValue("@EEPLifespam7", listInputs[195]);
                command.Parameters.AddWithValue("@EEPScreeningMachine", listInputs[196]);
                command.Parameters.AddWithValue("@EEPLifespam8", listInputs[197]);
                command.Parameters.AddWithValue("@EEPDesmucilaginador", listInputs[198]);
                command.Parameters.AddWithValue("@EEPLifespam9", listInputs[199]);
                command.Parameters.AddWithValue("@EEPMotorPump", listInputs[200]);
                command.Parameters.AddWithValue("@EEPLifespam10", listInputs[201]);
                command.Parameters.AddWithValue("@EEPOthersWetInput", listInputs[202]);
                command.Parameters.AddWithValue("@EEPLifespam11", listInputs[203]);
                command.Parameters.AddWithValue("@EEPConcrete", listInputs[204]);
                command.Parameters.AddWithValue("@EEPLifespam12", listInputs[205]);
                command.Parameters.AddWithValue("@EEPPlastic", listInputs[206]);
                command.Parameters.AddWithValue("@EEPLifespam13", listInputs[207]);
                command.Parameters.AddWithValue("@EEPRake", listInputs[208]);
                command.Parameters.AddWithValue("@EEPLifespam14", listInputs[209]);
                command.Parameters.AddWithValue("@EEPBroom", listInputs[210]);
                command.Parameters.AddWithValue("@EEPLifespam15", listInputs[211]);
                command.Parameters.AddWithValue("@EEPStorageRoom", listInputs[212]);
                command.Parameters.AddWithValue("@EEPLifespam16", listInputs[213]);
                command.Parameters.AddWithValue("@EEPOthersDryInput", listInputs[214]);
                command.Parameters.AddWithValue("@EEPLifespam17", listInputs[215]);
                command.Parameters.AddWithValue("@ACCApplicationFee", listInputs[216]);
                command.Parameters.AddWithValue("@ACCAnnualMembership", listInputs[217]);
                command.Parameters.AddWithValue("@ACCLifeInsurance", listInputs[218]);
                command.Parameters.AddWithValue("@ACCFLOCertification", listInputs[219]);
                command.Parameters.AddWithValue("@ACCOrganicCertification", listInputs[220]);
                command.Parameters.AddWithValue("@ACLLandValue", listInputs[221]);
                command.Parameters.AddWithValue("@ACLPropertyTax", listInputs[222]);
                command.Parameters.AddWithValue("@ACUSuperviseInvest", listInputs[223]);
                command.Parameters.AddWithValue("@ACUAdministInvest", listInputs[224]);
                command.Parameters.AddWithValue("@ACUTrainingInvest", listInputs[225]);
                command.Parameters.AddWithValue("@ACUExtraOrdInvest", listInputs[226]);
                command.Parameters.AddWithValue("@TGSeedPurchase", listInputs[227]);
                command.Parameters.AddWithValue("@TGWoodTransportation", listInputs[228]);
                command.Parameters.AddWithValue("@TGSandTransportation", listInputs[229]);
                command.Parameters.AddWithValue("@TGOthers", listInputs[230]);
                command.Parameters.AddWithValue("@TNSoilTransportation", listInputs[231]);
                command.Parameters.AddWithValue("@TNSacksMaterialShopping", listInputs[232]);
                command.Parameters.AddWithValue("@TNOthers", listInputs[233]);
                command.Parameters.AddWithValue("@TLPWoodTransportation", listInputs[234]);
                command.Parameters.AddWithValue("@TLPCompostTransportation", listInputs[235]);
                command.Parameters.AddWithValue("@TLPPlantTransportation", listInputs[236]);
                command.Parameters.AddWithValue("@TLPOthers", listInputs[237]);
                command.Parameters.AddWithValue("@TOtherEquipment", listInputs[238]);
                command.Parameters.AddWithValue("@TOtherLaborTransportation", listInputs[239]);
                command.Parameters.AddWithValue("@TOtherCoffeeTransportation", listInputs[240]);
                command.Parameters.AddWithValue("@TOtherSupervisingActivities", listInputs[241]);
                command.Parameters.AddWithValue("@TOthers", listInputs[242]);

                command.ExecuteNonQuery();
                connect.Close();
            }



        }
        public void saveUserAdvancedInputs(ChartInputAdvancedDTO inputAdvancedDTO)
        {
            //ChartInputAdvancedDTO inputAdvancedDTO = new ChartInputAdvancedDTO();
            //save user inputs
            String timeStamp = DateTime.Now.ToString();
            var conn = _iconfiguration.GetSection("ConnectionStrings").GetSection("CoffeeConnStr").Value;
            string sqlQuery = String.Format("Insert INTO [AVFCoffee].[dbo].[UserInputsAdvanced]" +
                   "([TimeStamp],[LGerminationSeedCollection],[LGerminationSeedSelection],[LGerminationNurseryConstruction]," +
                   "[LGerminationSeedingSupportIrrigation],[LGerminationOthers],[LNurseryConstruction],[LNurseryDrawnPulled]" +
                   ",[LNurseryClean],[LNurserySoilPreparationFertilizer],[LNurseryFilledLockedBags],[LNurseryButterflySowing]" +
                   ",[LNurseryIrrigation],[LNurseryFoliarApplication],[LNurseryReseeding],[LNurseryOthers],[LPPFieldCleaning]" +
                   ",[LPPCuttingTrees],[LPPWoodCollection],[LPPWoodChopping],[LPPCoffeeLayout],[LPPHoleDigging],[LPPSeedlingTransportation]" +
                   ",[LPPSeedlingTransplant],[LPPShadeAdjustment],[LPPCompostMixing],[LPPOthers],[LPPYWeeding],[LPPYOrganic],[LPPYChemical]" +
                   ",[LPPYFoliarSpraying],[LHPMYManualWeeding],[LHPMYChemicalWeeding],[LHPMYOrganicFertilizers],[LHPMYChemicalFertilizers]" +
                   ",[LHPMYFoliarSpraying],[LHPMYHedgerowsConstruction],[LHPMYShadetreePruning],[LHPMYPestControl],[LHPMYCoffeeGrowManagement]" +
                   ",[LHPMYOthers],[LHPHYCoffeeCollecDays],[LHPHYAdditionDays],[LHPPYFermentation],[LHPPYWashing],[LHPPYDrying],[LHPPYScreening]" +
                   ",[LHPPYSelection],[LHPPYStorage],[LHPPYCoffeewastewater],[LHPPYPulpManagement],[LHPPYOthers],[LHPMMManualWeeding]" +
                   ",[LHPMMChemicalWeeding],[LHPMMOrganicFertilizers],[LHPMMChemicalFertilizers],[LHPMMFoliarSpraying]," +
                   "[LHPMMHedgerowsConstruction],[LHPMMShadetreePruning],[LHPMMPestControl],[LHPMMCoffeeGrowManagement],[LHPMMOthers]" +
                   ",[LHPHMCoffeeCollecDays],[LHPHMAdditionDays],[LHPPMFermentation],[LHPPMWashing],[LHPPMDrying],[LHPPMScreening]" +
                   ",[LHPPMSelection],[LHPPMStorage],[LHPPMCoffeewastewater],[LHPPMPulpManagement],[LHPPMOthers],[LHPMOManualWeeding]" +
                   ",[LHPMOChemicalWeeding],[LHPMOOrganicFertilizers],[LHPMOChemicalFertilizers],[LHPMOFoliarSpraying]" +
                   ",[LHPMOHedgerowsConstruction],[LHPMOShadetreePruning],[LHPMOPestControl],[LHPMOCoffeeGrowManagement],[LHPMOOthers]" +
                   ",[LHPHOCoffeeCollecDays],[LHPHOAdditionDays],[LHPPOFermentation],[LHPPOWashing],[LHPPODrying],[LHPPOScreening]" +
                   ",[LHPPOSelection],[LHPPOStorage],[LHPPOCoffeewastewater],[LHPPOPulpManagement],[LHPPOOthers],[IIFood]" +
                   ",[IIAdditionalTransfers],[IIDaysoftraining],[ICCreditfromcooperative],[ICCreditfromcooperativeTime]" +
                   ",[ICCreditfromcooperativeInterest],[ICCreditfromagent],[ICCreditfromagentTime],[ICCreditfromagentInterest],[CostGerminator]" +
                   ",[CostGerminatorSeeds],[CostGerminatorSeedbed],[CostGerminatorSandSubstrate],[CostGerminatorCalciumSulfide]" +
                   ",[CostGerminatorLime],[CostGerminatorPlastic],[CostGerminatorOthers],[CostNurseryFertilizer],[CostNurseryPlasticBags]" +
                   ",[CostNurseryNetting],[CostNurseryStuds],[CostNurseryWire],[CostNurseryCiclonics],[CostNurseryStaples],[CostNurserySoil]" +
                   ",[CostNurseryBioFert],[CostNurseryAgroChemicals],[CostNurseryFungicide],[CostNurseryPhosphoricRock],[CostNurseryOthers]" +
                   ",[CostFLPPOrganicFert],[CostFLPPChemicalFert],[CostFVGOrganicFert],[CostFVGChemicalFert],[CostFMOtherFert]" +
                   ",[CostFMOrganicFoliar],[CostFMChemicalFoliar],[CostFMGasFuel],[CostFMOthers],[EGEManualSprayer],[EGEMachetes],[EGEShovel]" +
                   ",[EGEHoe],[EGEWheelBarrow],[EGELime],[EGEAuger],[EGEMetalBar],[EGEHose],[EGESprinklers],[EGEChainSaw],[EGEHandSaw]" +
                   ",[EGEMotorPump],[EGEPrunningScissors],[EGEAxe],[EEHScale],[EEHVehicle],[EEHWorkAnimal],[EEHMotorcycle],[EEHBags],[EEHSack]" +
                   ",[EEHStraw],[EEHBaskets],[EEHBoxes],[EEHOthers],[EEPPulperMachine],[EEPTolca],[EEPEngine],[EEPTanks],[EEPWaterChannel]" +
                   ",[EEPPVCPipes],[EEPFilteringSystem],[EEPScreeningMachine],[EEPDesmucilaginador],[EEPMotorPump],[EEPOthersWetInput]" +
                   ",[EEPConcrete],[EEPPlastic],[EEPRake],[EEPBroom],[EEPStorageRoom],[EEPOthersDryInput],[ACCApplicationFee]" +
                   ",[ACCAnnualMembership],[ACCLifeInsurance],[ACCFLOCertification],[ACCOrganicCertification],[ACLLandValue],[ACLPropertyTax]" +
                   ",[ACUSuperviseInvest],[ACUAdministInvest],[ACUTrainingInvest],[ACUExtraOrdInvest],[TGSeedPurchase],[TGWoodTransportation]" +
                   ",[TGSandTransportation],[TGOthers],[TNSoilTransportation],[TNSacksMaterialShopping],[TNOthers],[TLPWoodTransportation]" +
                   ",[TLPCompostTransportation],[TLPPlantTransportation],[TLPOthers],[TOtherEquipment],[TOtherLaborTransportation]" +
                   ",[TOtherCoffeeTransportation],[TOtherSupervisingActivities],[TOthers],[LPPYOther],[EGELifespam1],[EGELifespam2]" +
                   ",[EGELifespam3],[EGELifespam4],[EGELifespam5],[EGELifespam6],[EGELifespam7],[EGELifespam8],[EGELifespam9],[EGELifespam10]" +
                   ",[EGELifespam11],[EGELifespam12],[EGELifespam13],[EGELifespam14],[EGELifespam15],[EEHLifespam1],[EEHLifespam2],[EEHLifespam3]" +
                   ",[EEHLifespam4],[EEHLifespam5],[EEHLifespam6],[EEHLifespam7],[EEHLifespam8],[EEHLifespam9],[EEHLifespam10],[EEPLifespam1]" +
                   ",[EEPLifespam2],[EEPLifespam3],[EEPLifespam4],[EEPLifespam5],[EEPLifespam6],[EEPLifespam7],[EEPLifespam8],[EEPLifespam9],[CoopID]" +
                   ",[EEPLifespam10],[EEPLifespam11],[EEPLifespam12],[EEPLifespam13],[EEPLifespam14],[EEPLifespam15],[EEPLifespam16],[EEPLifespam17],[UserID]) VALUES" +
                   " (@TimeStamp,@LGerminationSeedCollection,@LGerminationSeedSelection,@LGerminationNurseryConstruction, " +
                   "@LGerminationSeedingSupportIrrigation,@LGerminationOthers,@LNurseryConstruction,@LNurseryDrawnPulled" +
                   ",@LNurseryClean,@LNurserySoilPreparationFertilizer,@LNurseryFilledLockedBags,@LNurseryButterflySowing" +
                   ",@LNurseryIrrigation,@LNurseryFoliarApplication,@LNurseryReseeding,@LNurseryOthers,@LPPFieldCleaning" +
                   ",@LPPCuttingTrees,@LPPWoodCollection,@LPPWoodChopping,@LPPCoffeeLayout,@LPPHoleDigging,@LPPSeedlingTransportation" +
                   ",@LPPSeedlingTransplant,@LPPShadeAdjustment,@LPPCompostMixing,@LPPOthers,@LPPYWeeding,@LPPYOrganic,@LPPYChemical" +
                   ",@LPPYFoliarSpraying,@LHPMYManualWeeding,@LHPMYChemicalWeeding,@LHPMYOrganicFertilizers,@LHPMYChemicalFertilizers" +
                   ",@LHPMYFoliarSpraying,@LHPMYHedgerowsConstruction,@LHPMYShadetreePruning,@LHPMYPestControl,@LHPMYCoffeeGrowManagement" +
                   ",@LHPMYOthers,@LHPHYCoffeeCollecDays,@LHPHYAdditionDays,@LHPPYFermentation,@LHPPYWashing,@LHPPYDrying,@LHPPYScreening" +
                   ",@LHPPYSelection,@LHPPYStorage,@LHPPYCoffeewastewater,@LHPPYPulpManagement,@LHPPYOthers,@LHPMMManualWeeding" +
                   ",@LHPMMChemicalWeeding,@LHPMMOrganicFertilizers,@LHPMMChemicalFertilizers,@LHPMMFoliarSpraying," +
                   "@LHPMMHedgerowsConstruction,@LHPMMShadetreePruning,@LHPMMPestControl,@LHPMMCoffeeGrowManagement,@LHPMMOthers" +
                   ",@LHPHMCoffeeCollecDays,@LHPHMAdditionDays,@LHPPMFermentation,@LHPPMWashing,@LHPPMDrying,@LHPPMScreening" +
                   ",@LHPPMSelection,@LHPPMStorage,@LHPPMCoffeewastewater,@LHPPMPulpManagement,@LHPPMOthers,@LHPMOManualWeeding" +
                   ",@LHPMOChemicalWeeding,@LHPMOOrganicFertilizers,@LHPMOChemicalFertilizers,@LHPMOFoliarSpraying" +
                   ",@LHPMOHedgerowsConstruction,@LHPMOShadetreePruning,@LHPMOPestControl,@LHPMOCoffeeGrowManagement,@LHPMOOthers" +
                   ",@LHPHOCoffeeCollecDays,@LHPHOAdditionDays,@LHPPOFermentation,@LHPPOWashing,@LHPPODrying,@LHPPOScreening" +
                   ",@LHPPOSelection,@LHPPOStorage,@LHPPOCoffeewastewater,@LHPPOPulpManagement,@LHPPOOthers,@IIFood" +
                   ",@IIAdditionalTransfers,@IIDaysoftraining,@ICCreditfromcooperative,@ICCreditfromcooperativeTime" +
                   ",@ICCreditfromcooperativeInterest,@ICCreditfromagent,@ICCreditfromagentTime,@ICCreditfromagentInterest,@CostGerminator" +
                   ",@CostGerminatorSeeds,@CostGerminatorSeedbed,@CostGerminatorSandSubstrate,@CostGerminatorCalciumSulfide" +
                   ",@CostGerminatorLime,@CostGerminatorPlastic,@CostGerminatorOthers,@CostNurseryFertilizer,@CostNurseryPlasticBags" +
                   ",@CostNurseryNetting,@CostNurseryStuds,@CostNurseryWire,@CostNurseryCiclonics,@CostNurseryStaples,@CostNurserySoil" +
                   ",@CostNurseryBioFert,@CostNurseryAgroChemicals,@CostNurseryFungicide,@CostNurseryPhosphoricRock,@CostNurseryOthers" +
                   ",@CostFLPPOrganicFert,@CostFLPPChemicalFert,@CostFVGOrganicFert,@CostFVGChemicalFert,@CostFMOtherFert" +
                   ",@CostFMOrganicFoliar,@CostFMChemicalFoliar,@CostFMGasFuel,@CostFMOthers,@EGEManualSprayer,@EGEMachetes,@EGEShovel" +
                   ",@EGEHoe,@EGEWheelBarrow,@EGELime,@EGEAuger,@EGEMetalBar,@EGEHose,@EGESprinklers,@EGEChainSaw,@EGEHandSaw" +
                   ",@EGEMotorPump,@EGEPrunningScissors,@EGEAxe,@EEHScale,@EEHVehicle,@EEHWorkAnimal,@EEHMotorcycle,@EEHBags,@EEHSack" +
                   ",@EEHStraw,@EEHBaskets,@EEHBoxes,@EEHOthers,@EEPPulperMachine,@EEPTolca,@EEPEngine,@EEPTanks,@EEPWaterChannel" +
                   ",@EEPPVCPipes,@EEPFilteringSystem,@EEPScreeningMachine,@EEPDesmucilaginador,@EEPMotorPump,@EEPOthersWetInput" +
                   ",@EEPConcrete,@EEPPlastic,@EEPRake,@EEPBroom,@EEPStorageRoom,@EEPOthersDryInput,@ACCApplicationFee" +
                   ",@ACCAnnualMembership,@ACCLifeInsurance,@ACCFLOCertification,@ACCOrganicCertification,@ACLLandValue,@ACLPropertyTax" +
                   ",@ACUSuperviseInvest,@ACUAdministInvest,@ACUTrainingInvest,@ACUExtraOrdInvest,@TGSeedPurchase,@TGWoodTransportation" +
                   ",@TGSandTransportation,@TGOthers,@TNSoilTransportation,@TNSacksMaterialShopping,@TNOthers,@TLPWoodTransportation" +
                   ",@TLPCompostTransportation,@TLPPlantTransportation,@TLPOthers,@TOtherEquipment,@TOtherLaborTransportation" +
                   ",@TOtherCoffeeTransportation,@TOtherSupervisingActivities,@TOthers,@LPPYOther,@EGELifespam1,@EGELifespam2" +
                   ",@EGELifespam3,@EGELifespam4,@EGELifespam5,@EGELifespam6,@EGELifespam7,@EGELifespam8,@EGELifespam9,@EGELifespam10" +
                   ",@EGELifespam11,@EGELifespam12,@EGELifespam13,@EGELifespam14,@EGELifespam15,@EEHLifespam1,@EEHLifespam2,@EEHLifespam3" +
                   ",@EEHLifespam4,@EEHLifespam5,@EEHLifespam6,@EEHLifespam7,@EEHLifespam8,@EEHLifespam9,@EEHLifespam10,@EEPLifespam1" +
                   ",@EEPLifespam2,@EEPLifespam3,@EEPLifespam4,@EEPLifespam5,@EEPLifespam6,@EEPLifespam7,@EEPLifespam8,@EEPLifespam9,@CoopID" +
                   ",@EEPLifespam10,@EEPLifespam11,@EEPLifespam12,@EEPLifespam13,@EEPLifespam14,@EEPLifespam15,@EEPLifespam16,@EEPLifespam17,@UserID)");
            using (SqlConnection connect = new SqlConnection(conn))
            {
                connect.Open();
                SqlCommand command = new SqlCommand(sqlQuery);
                command.Parameters.AddWithValue("@UserID", "1234");
                command.Parameters.AddWithValue("@CoopID", "1111");
                command.Parameters.AddWithValue("@LGerminationSeedCollection", inputAdvancedDTO.LGerminationSeedCollection);
                command.Parameters.AddWithValue("@LGerminationSeedSelection", inputAdvancedDTO.LGerminationSeedSelection);
                command.Parameters.AddWithValue("@LGerminationNurseryConstruction", inputAdvancedDTO.LGerminationNurseryConstruction);
                command.Parameters.AddWithValue("@LGerminationSeedingSupportIrrigation", inputAdvancedDTO.LGerminationSeedingSupportIrrigation);
                command.Parameters.AddWithValue("@LGerminationOthers", inputAdvancedDTO.LGerminationOthers);
                command.Parameters.AddWithValue("@LNurseryConstruction", inputAdvancedDTO.LNurseryConstruction);
                command.Parameters.AddWithValue("@LNurseryDrawnPulled", inputAdvancedDTO.LNurseryDrawnPulled);
                command.Parameters.AddWithValue("@LNurseryClean", inputAdvancedDTO.LNurseryClean);
                command.Parameters.AddWithValue("@LNurserySoilPreparationFertilizer", inputAdvancedDTO.LNurserySoilPreparationFertilizer);
                command.Parameters.AddWithValue("@LNurseryFilledLockedBags", inputAdvancedDTO.LNurseryFilledLockedBags);
                command.Parameters.AddWithValue("@LNurseryButterflySowing", inputAdvancedDTO.LNurseryButterflySowing);
                command.Parameters.AddWithValue("@LNurseryIrrigation", inputAdvancedDTO.LNurseryIrrigation);
                command.Parameters.AddWithValue("@LNurseryFoliarApplication", inputAdvancedDTO.LNurseryFoliarApplication);
                command.Parameters.AddWithValue("@LNurseryReseeding", inputAdvancedDTO.LNurseryReseeding);
                command.Parameters.AddWithValue("@LNurseryOthers", inputAdvancedDTO.LNurseryOthers);
                command.Parameters.AddWithValue("@LPPFieldCleaning", inputAdvancedDTO.LPPFieldCleaning);
                command.Parameters.AddWithValue("@LPPCuttingTrees", inputAdvancedDTO.LPPCuttingTrees);
                command.Parameters.AddWithValue("@LPPWoodCollection", inputAdvancedDTO.LPPWoodCollection);
                command.Parameters.AddWithValue("@LPPWoodChopping", inputAdvancedDTO.LPPWoodChopping);
                command.Parameters.AddWithValue("@LPPCoffeeLayout", inputAdvancedDTO.LPPCoffeeLayout);
                command.Parameters.AddWithValue("@LPPHoleDigging", inputAdvancedDTO.LPPHoleDigging);
                command.Parameters.AddWithValue("@LPPSeedlingTransportation", inputAdvancedDTO.LPPSeedlingTransportation);
                command.Parameters.AddWithValue("@LPPSeedlingTransplant", inputAdvancedDTO.LPPSeedlingTransplant);
                command.Parameters.AddWithValue("@LPPShadeAdjustment", inputAdvancedDTO.LPPShadeAdjustment);
                command.Parameters.AddWithValue("@LPPCompostMixing", inputAdvancedDTO.LPPCompostMixing);
                command.Parameters.AddWithValue("@LPPOthers", inputAdvancedDTO.LPPOthers);
                command.Parameters.AddWithValue("@LPPYWeeding", inputAdvancedDTO.LPPYWeeding);
                command.Parameters.AddWithValue("@LPPYOrganic", inputAdvancedDTO.LPPYOrganic);
                command.Parameters.AddWithValue("@LPPYChemical", inputAdvancedDTO.LPPYChemical);
                command.Parameters.AddWithValue("@LPPYFoliarSpraying", inputAdvancedDTO.LPPYFoliarSpraying);
                command.Parameters.AddWithValue("@LHPMYManualWeeding", inputAdvancedDTO.LHPMYManualWeeding);
                command.Parameters.AddWithValue("@LHPMYChemicalWeeding", inputAdvancedDTO.LHPMYChemicalWeeding);
                command.Parameters.AddWithValue("@LHPMYOrganicFertilizers", inputAdvancedDTO.LHPMYOrganicFertilizers);
                command.Parameters.AddWithValue("@LHPMYChemicalFertilizers", inputAdvancedDTO.LHPMYChemicalFertilizers);
                command.Parameters.AddWithValue("@LHPMYFoliarSpraying", inputAdvancedDTO.LHPMYFoliarSpraying);
                command.Parameters.AddWithValue("@LHPMYHedgerowsConstruction", inputAdvancedDTO.LHPMYHedgerowsConstruction);
                command.Parameters.AddWithValue("@LHPMYShadetreePruning", inputAdvancedDTO.LHPMYShadetreePruning);
                command.Parameters.AddWithValue("@LHPMYPestControl", inputAdvancedDTO.LHPMYPestControl);
                command.Parameters.AddWithValue("@LHPMYCoffeeGrowManagement", inputAdvancedDTO.LHPMYCoffeeGrowManagement);
                command.Parameters.AddWithValue("@LHPMYOthers", inputAdvancedDTO.LHPMYOthers);
                command.Parameters.AddWithValue("@LHPHYCoffeeCollecDays", inputAdvancedDTO.LHPHYCoffeeCollecDays);
                command.Parameters.AddWithValue("@LHPHYAdditionDays", inputAdvancedDTO.LHPHYAdditionDays);
                command.Parameters.AddWithValue("@LHPPYFermentation", inputAdvancedDTO.LHPPYFermentation);
                command.Parameters.AddWithValue("@LHPPYWashing", inputAdvancedDTO.LHPPYWashing);
                command.Parameters.AddWithValue("@LHPPYDrying", inputAdvancedDTO.LHPPYDrying);
                command.Parameters.AddWithValue("@LHPPYScreening", inputAdvancedDTO.LHPPYScreening);
                command.Parameters.AddWithValue("@LHPPYSelection", inputAdvancedDTO.LHPPYSelection);
                command.Parameters.AddWithValue("@LHPPYStorage", inputAdvancedDTO.LHPPYStorage);
                command.Parameters.AddWithValue("@LHPPYCoffeewastewater", inputAdvancedDTO.LHPPYCoffeewastewater);
                command.Parameters.AddWithValue("@LHPPYPulpManagement", inputAdvancedDTO.LHPPYPulpManagement);
                command.Parameters.AddWithValue("@LHPPYOthers", inputAdvancedDTO.LHPPYOthers);
                command.Parameters.AddWithValue("@LHPMMManualWeeding", inputAdvancedDTO.LHPMMManualWeeding);
                command.Parameters.AddWithValue("@LHPMMChemicalWeeding", inputAdvancedDTO.LHPMMChemicalWeeding);
                command.Parameters.AddWithValue("@LHPMMOrganicFertilizers", inputAdvancedDTO.LHPMMOrganicFertilizers);
                command.Parameters.AddWithValue("@LHPMMChemicalFertilizers", inputAdvancedDTO.LHPMMChemicalFertilizers);
                command.Parameters.AddWithValue("@LHPMMFoliarSpraying", inputAdvancedDTO.LHPMMFoliarSpraying);
                command.Parameters.AddWithValue("@LHPMMHedgerowsConstruction", inputAdvancedDTO.LHPMMHedgerowsConstruction);
                command.Parameters.AddWithValue("@LHPMMShadetreePruning", inputAdvancedDTO.LHPMMShadetreePruning);
                command.Parameters.AddWithValue("@LHPMMPestControl", inputAdvancedDTO.LHPMMPestControl);
                command.Parameters.AddWithValue("@LHPMMCoffeeGrowManagement", inputAdvancedDTO.LHPMMCoffeeGrowManagement);
                command.Parameters.AddWithValue("@LHPMMOthers", inputAdvancedDTO.LHPMMOthers);
                command.Parameters.AddWithValue("@LHPHMCoffeeCollecDays", inputAdvancedDTO.LHPHMCoffeeCollecDays);
                command.Parameters.AddWithValue("@LHPHMAdditionDays", inputAdvancedDTO.LHPHMAdditionDays);
                command.Parameters.AddWithValue("@LHPPMFermentation", inputAdvancedDTO.LHPPMFermentation);
                command.Parameters.AddWithValue("@LHPPMWashing", inputAdvancedDTO.LHPPMWashing);
                command.Parameters.AddWithValue("@LHPPMDrying", inputAdvancedDTO.LHPPMDrying);
                command.Parameters.AddWithValue("@LHPPMScreening", inputAdvancedDTO.LHPPMScreening);
                command.Parameters.AddWithValue("@LHPPMSelection", inputAdvancedDTO.LHPPMSelection);
                command.Parameters.AddWithValue("@LHPPMStorage", inputAdvancedDTO.LHPPMStorage);
                command.Parameters.AddWithValue("@LHPPMCoffeewastewater", inputAdvancedDTO.LHPPMCoffeewastewater);
                command.Parameters.AddWithValue("@LHPPMPulpManagement", inputAdvancedDTO.LHPPMPulpManagement);
                command.Parameters.AddWithValue("@LHPPMOthers", inputAdvancedDTO.LHPPMOthers);
                command.Parameters.AddWithValue("@LHPMOManualWeeding", inputAdvancedDTO.LHPMOManualWeeding);
                command.Parameters.AddWithValue("@LHPMOChemicalWeeding", inputAdvancedDTO.LHPMOChemicalWeeding);
                command.Parameters.AddWithValue("@LHPMOOrganicFertilizers", inputAdvancedDTO.LHPMOOrganicFertilizers);
                command.Parameters.AddWithValue("@LHPMOChemicalFertilizers", inputAdvancedDTO.LHPMOChemicalFertilizers);
                command.Parameters.AddWithValue("@LHPMOFoliarSpraying", inputAdvancedDTO.LHPMOFoliarSpraying);
                command.Parameters.AddWithValue("@LHPMOHedgerowsConstruction", inputAdvancedDTO.LHPMOHedgerowsConstruction);
                command.Parameters.AddWithValue("@LHPMOShadetreePruning", inputAdvancedDTO.LHPMOShadetreePruning);
                command.Parameters.AddWithValue("@LHPMOPestControl", inputAdvancedDTO.LHPMOPestControl);
                command.Parameters.AddWithValue("@LHPMOCoffeeGrowManagement", inputAdvancedDTO.LHPMOCoffeeGrowManagement);
                command.Parameters.AddWithValue("@LHPMOOthers", inputAdvancedDTO.LHPMOOthers);
                command.Parameters.AddWithValue("@LHPHOCoffeeCollecDays", inputAdvancedDTO.LHPHOCoffeeCollecDays);
                command.Parameters.AddWithValue("@LHPHOAdditionDays", inputAdvancedDTO.LHPHOAdditionDays);
                command.Parameters.AddWithValue("@LHPPOFermentation", inputAdvancedDTO.LHPPOFermentation);
                command.Parameters.AddWithValue("@LHPPOWashing", inputAdvancedDTO.LHPPOWashing);
                command.Parameters.AddWithValue("@LHPPODrying", inputAdvancedDTO.LHPPODrying);
                command.Parameters.AddWithValue("@LHPPOScreening", inputAdvancedDTO.LHPPOScreening);
                command.Parameters.AddWithValue("@LHPPOSelection", inputAdvancedDTO.LHPPOSelection);
                command.Parameters.AddWithValue("@LHPPOStorage", inputAdvancedDTO.LHPPOStorage);
                command.Parameters.AddWithValue("@LHPPOCoffeewastewater", inputAdvancedDTO.LHPPOCoffeewastewater);
                command.Parameters.AddWithValue("@LHPPOPulpManagement", inputAdvancedDTO.LHPPOPulpManagement);
                command.Parameters.AddWithValue("@LHPPOOthers", inputAdvancedDTO.LHPPOOthers);
                command.Parameters.AddWithValue("@IIFood", inputAdvancedDTO.IIFood);
                command.Parameters.AddWithValue("@IIAdditionalTransfers", inputAdvancedDTO.IIAdditionalTransfers);
                command.Parameters.AddWithValue("@IIDaysoftraining", inputAdvancedDTO.IIDaysoftraining);
                command.Parameters.AddWithValue("@ICCreditfromcooperative", inputAdvancedDTO.ICCreditfromcooperative);
                command.Parameters.AddWithValue("@ICCreditfromcooperativeTime", inputAdvancedDTO.ICCreditfromcooperativeTime);
                command.Parameters.AddWithValue("@ICCreditfromcooperativeInterest", inputAdvancedDTO.ICCreditfromcooperativeInterest);
                command.Parameters.AddWithValue("@ICCreditfromagent", inputAdvancedDTO.ICCreditfromagent);
                command.Parameters.AddWithValue("@ICCreditfromagentTime", inputAdvancedDTO.ICCreditfromagentTime);
                command.Parameters.AddWithValue("@ICCreditfromagentInterest", inputAdvancedDTO.ICCreditfromagentInterest);
                command.Parameters.AddWithValue("@CostGerminator", inputAdvancedDTO.CostGerminator);
                command.Parameters.AddWithValue("@CostGerminatorSeeds", inputAdvancedDTO.CostGerminatorSeeds);
                command.Parameters.AddWithValue("@CostGerminatorSeedbed", inputAdvancedDTO.CostGerminatorSeedbed);
                command.Parameters.AddWithValue("@CostGerminatorSandSubstrate", inputAdvancedDTO.CostGerminatorSandSubstrate);
                command.Parameters.AddWithValue("@CostGerminatorCalciumSulfide", inputAdvancedDTO.CostGerminatorCalciumSulfide);
                command.Parameters.AddWithValue("@CostGerminatorLime", inputAdvancedDTO.CostGerminatorLime);
                command.Parameters.AddWithValue("@CostGerminatorPlastic", inputAdvancedDTO.CostGerminatorPlastic);
                command.Parameters.AddWithValue("@CostGerminatorOthers", inputAdvancedDTO.CostGerminatorOthers);
                command.Parameters.AddWithValue("@CostNurseryFertilizer", inputAdvancedDTO.CostNurseryFertilizer);
                command.Parameters.AddWithValue("@CostNurseryPlasticBags", inputAdvancedDTO.CostNurseryPlasticBags);
                command.Parameters.AddWithValue("@CostNurseryNetting", inputAdvancedDTO.CostNurseryNetting);
                command.Parameters.AddWithValue("@CostNurseryStuds", inputAdvancedDTO.CostNurseryStuds);
                command.Parameters.AddWithValue("@CostNurseryWire", inputAdvancedDTO.CostNurseryWire);
                command.Parameters.AddWithValue("@CostNurseryCiclonics", inputAdvancedDTO.CostNurseryCiclonics);
                command.Parameters.AddWithValue("@CostNurseryStaples", inputAdvancedDTO.CostNurseryStaples);
                command.Parameters.AddWithValue("@CostNurserySoil", inputAdvancedDTO.CostNurserySoil);
                command.Parameters.AddWithValue("@CostNurseryBioFert", inputAdvancedDTO.CostNurseryBioFert);
                command.Parameters.AddWithValue("@CostNurseryAgroChemicals", inputAdvancedDTO.CostNurseryAgroChemicals);
                command.Parameters.AddWithValue("@CostNurseryFungicide", inputAdvancedDTO.CostNurseryFungicide);
                command.Parameters.AddWithValue("@CostNurseryPhosphoricRock", inputAdvancedDTO.CostNurseryPhosphoricRock);
                command.Parameters.AddWithValue("@CostNurseryOthers", inputAdvancedDTO.CostNurseryOthers);
                command.Parameters.AddWithValue("@CostFLPPOrganicFert", inputAdvancedDTO.CostFLPPOrganicFert);
                command.Parameters.AddWithValue("@CostFLPPChemicalFert", inputAdvancedDTO.CostFLPPChemicalFert);
                command.Parameters.AddWithValue("@CostFVGOrganicFert", inputAdvancedDTO.CostFVGOrganicFert);
                command.Parameters.AddWithValue("@CostFVGChemicalFert", inputAdvancedDTO.CostFVGChemicalFert);
                command.Parameters.AddWithValue("@CostFMOtherFert", inputAdvancedDTO.CostFMOtherFert);
                command.Parameters.AddWithValue("@CostFMOrganicFoliar", inputAdvancedDTO.CostFMOrganicFoliar);
                command.Parameters.AddWithValue("@CostFMChemicalFoliar", inputAdvancedDTO.CostFMChemicalFoliar);
                command.Parameters.AddWithValue("@CostFMGasFuel", inputAdvancedDTO.CostFMGasFuel);
                command.Parameters.AddWithValue("@CostFMOthers", inputAdvancedDTO.CostFMOthers);
                command.Parameters.AddWithValue("@EGEManualSprayer", inputAdvancedDTO.EGEManualSprayer);
                command.Parameters.AddWithValue("@EGEMachetes", inputAdvancedDTO.EGEMachetes);
                command.Parameters.AddWithValue("@EGEShovel", inputAdvancedDTO.EGEShovel);
                command.Parameters.AddWithValue("@EGEHoe", inputAdvancedDTO.EGEHoe);
                command.Parameters.AddWithValue("@EGEWheelBarrow", inputAdvancedDTO.EGEWheelBarrow);
                command.Parameters.AddWithValue("@EGELime", inputAdvancedDTO.EGELime);
                command.Parameters.AddWithValue("@EGEAuger", inputAdvancedDTO.EGEAuger);
                command.Parameters.AddWithValue("@EGEMetalBar", inputAdvancedDTO.EGEMetalBar);
                command.Parameters.AddWithValue("@EGEHose", inputAdvancedDTO.EGEHose);
                command.Parameters.AddWithValue("@EGESprinklers", inputAdvancedDTO.EGESprinklers);
                command.Parameters.AddWithValue("@EGEChainSaw", inputAdvancedDTO.EGEChainSaw);
                command.Parameters.AddWithValue("@EGEHandSaw", inputAdvancedDTO.EGEHandSaw);
                command.Parameters.AddWithValue("@EGEMotorPump", inputAdvancedDTO.EGEMotorPump);
                command.Parameters.AddWithValue("@EGEPrunningScissors", inputAdvancedDTO.EGEPrunningScissors);
                command.Parameters.AddWithValue("@EGEAxe", inputAdvancedDTO.EGEAxe);
                command.Parameters.AddWithValue("@EEHScale", inputAdvancedDTO.EEHScale);
                command.Parameters.AddWithValue("@EEHVehicle", inputAdvancedDTO.EEHVehicle);
                command.Parameters.AddWithValue("@EEHWorkAnimal", inputAdvancedDTO.EEHWorkAnimal);
                command.Parameters.AddWithValue("@EEHMotorcycle", inputAdvancedDTO.EEHMotorcycle);
                command.Parameters.AddWithValue("@EEHBags", inputAdvancedDTO.EEHBags);
                command.Parameters.AddWithValue("@EEHSack", inputAdvancedDTO.EEHSack);
                command.Parameters.AddWithValue("@EEHStraw", inputAdvancedDTO.EEHStraw);
                command.Parameters.AddWithValue("@EEHBaskets", inputAdvancedDTO.EEHBaskets);
                command.Parameters.AddWithValue("@EEHBoxes", inputAdvancedDTO.EEHBoxes);
                command.Parameters.AddWithValue("@EEHOthers", inputAdvancedDTO.EEHOthers);
                command.Parameters.AddWithValue("@EEPPulperMachine", inputAdvancedDTO.EEPPulperMachine);
                command.Parameters.AddWithValue("@EEPTolca", inputAdvancedDTO.EEPTolca);
                command.Parameters.AddWithValue("@EEPEngine", inputAdvancedDTO.EEPEngine);
                command.Parameters.AddWithValue("@EEPTanks", inputAdvancedDTO.EEPTanks);
                command.Parameters.AddWithValue("@EEPWaterChannel", inputAdvancedDTO.EEPWaterChannel);
                command.Parameters.AddWithValue("@EEPPVCPipes", inputAdvancedDTO.EEPPVCPipes);
                command.Parameters.AddWithValue("@EEPFilteringSystem", inputAdvancedDTO.EEPFilteringSystem);
                command.Parameters.AddWithValue("@EEPScreeningMachine", inputAdvancedDTO.EEPScreeningMachine);
                command.Parameters.AddWithValue("@EEPDesmucilaginador", inputAdvancedDTO.EEPDesmucilaginador);
                command.Parameters.AddWithValue("@EEPMotorPump", inputAdvancedDTO.EEPMotorPump);
                command.Parameters.AddWithValue("@EEPOthersWetInput", inputAdvancedDTO.EEPOthersWetInput);
                command.Parameters.AddWithValue("@EEPConcrete", inputAdvancedDTO.EEPConcrete);
                command.Parameters.AddWithValue("@EEPPlastic", inputAdvancedDTO.EEPPlastic);
                command.Parameters.AddWithValue("@EEPRake", inputAdvancedDTO.EEPRake);
                command.Parameters.AddWithValue("@EEPBroom", inputAdvancedDTO.EEPBroom);
                command.Parameters.AddWithValue("@EEPStorageRoom", inputAdvancedDTO.EEPStorageRoom);
                command.Parameters.AddWithValue("@EEPOthersDryInput", inputAdvancedDTO.EEPOthersDryInput);
                command.Parameters.AddWithValue("@ACCApplicationFee", inputAdvancedDTO.ACCApplicationFee);
                command.Parameters.AddWithValue("@ACCAnnualMembership", inputAdvancedDTO.ACCAnnualMembership);
                command.Parameters.AddWithValue("@ACCLifeInsurance", inputAdvancedDTO.ACCLifeInsurance);
                command.Parameters.AddWithValue("@ACCFLOCertification", inputAdvancedDTO.ACCFLOCertification);
                command.Parameters.AddWithValue("@ACCOrganicCertification", inputAdvancedDTO.ACCOrganicCertification);
                command.Parameters.AddWithValue("@ACLLandValue", inputAdvancedDTO.ACLLandValue);
                command.Parameters.AddWithValue("@ACLPropertyTax", inputAdvancedDTO.ACLPropertyTax);
                command.Parameters.AddWithValue("@ACUSuperviseInvest", inputAdvancedDTO.ACUSuperviseInvest);
                command.Parameters.AddWithValue("@ACUAdministInvest", inputAdvancedDTO.ACUAdministInvest);
                command.Parameters.AddWithValue("@ACUTrainingInvest", inputAdvancedDTO.ACUTrainingInvest);
                command.Parameters.AddWithValue("@ACUExtraOrdInvest", inputAdvancedDTO.ACUExtraOrdInvest);
                command.Parameters.AddWithValue("@TGSeedPurchase", inputAdvancedDTO.TGSeedPurchase);
                command.Parameters.AddWithValue("@TGWoodTransportation", inputAdvancedDTO.TGWoodTransportation);
                command.Parameters.AddWithValue("@TGSandTransportation", inputAdvancedDTO.TGSandTransportation);
                command.Parameters.AddWithValue("@TGOthers", inputAdvancedDTO.TGOthers);
                command.Parameters.AddWithValue("@TNSoilTransportation", inputAdvancedDTO.TNSoilTransportation);
                command.Parameters.AddWithValue("@TNSacksMaterialShopping", inputAdvancedDTO.TNSacksMaterialShopping);
                command.Parameters.AddWithValue("@TNOthers", inputAdvancedDTO.TNOthers);
                command.Parameters.AddWithValue("@TLPWoodTransportation", inputAdvancedDTO.TLPWoodTransportation);
                command.Parameters.AddWithValue("@TLPCompostTransportation", inputAdvancedDTO.TLPCompostTransportation);
                command.Parameters.AddWithValue("@TLPPlantTransportation", inputAdvancedDTO.TLPPlantTransportation);
                command.Parameters.AddWithValue("@TLPOthers", inputAdvancedDTO.TLPOthers);
                command.Parameters.AddWithValue("@TOtherEquipment", inputAdvancedDTO.TOtherEquipment);
                command.Parameters.AddWithValue("@TOtherLaborTransportation", inputAdvancedDTO.TOtherLaborTransportation);
                command.Parameters.AddWithValue("@TOtherCoffeeTransportation", inputAdvancedDTO.TOtherCoffeeTransportation);
                command.Parameters.AddWithValue("@TOtherSupervisingActivities", inputAdvancedDTO.TOtherSupervisingActivities);
                command.Parameters.AddWithValue("@TOthers", inputAdvancedDTO.TOthers);
                command.Parameters.AddWithValue("@LPPYOther", inputAdvancedDTO.LPPYOther);
                command.Parameters.AddWithValue("@EGELifespam1", inputAdvancedDTO.EGELifespam1);
                command.Parameters.AddWithValue("@EGELifespam2", inputAdvancedDTO.EGELifespam2);
                command.Parameters.AddWithValue("@EGELifespam3", inputAdvancedDTO.EGELifespam3);
                command.Parameters.AddWithValue("@EGELifespam4", inputAdvancedDTO.EGELifespam4);
                command.Parameters.AddWithValue("@EGELifespam5", inputAdvancedDTO.EGELifespam5);
                command.Parameters.AddWithValue("@EGELifespam6", inputAdvancedDTO.EGELifespam6);
                command.Parameters.AddWithValue("@EGELifespam7", inputAdvancedDTO.EGELifespam7);
                command.Parameters.AddWithValue("@EGELifespam8", inputAdvancedDTO.EGELifespam8);
                command.Parameters.AddWithValue("@EGELifespam9", inputAdvancedDTO.EGELifespam9);
                command.Parameters.AddWithValue("@EGELifespam10", inputAdvancedDTO.EGELifespam10);
                command.Parameters.AddWithValue("@EGELifespam11", inputAdvancedDTO.EGELifespam11);
                command.Parameters.AddWithValue("@EGELifespam12", inputAdvancedDTO.EGELifespam12);
                command.Parameters.AddWithValue("@EGELifespam13", inputAdvancedDTO.EGELifespam13);
                command.Parameters.AddWithValue("@EGELifespam14", inputAdvancedDTO.EGELifespam14);
                command.Parameters.AddWithValue("@EGELifespam15", inputAdvancedDTO.EGELifespam15);
                command.Parameters.AddWithValue("@EEHLifespam1", inputAdvancedDTO.EEHLifespam1);
                command.Parameters.AddWithValue("@EEHLifespam2", inputAdvancedDTO.EEHLifespam2);
                command.Parameters.AddWithValue("@EEHLifespam3", inputAdvancedDTO.EEHLifespam3);
                command.Parameters.AddWithValue("@EEHLifespam4", inputAdvancedDTO.EEHLifespam4);
                command.Parameters.AddWithValue("@EEHLifespam5", inputAdvancedDTO.EEHLifespam5);
                command.Parameters.AddWithValue("@EEHLifespam6", inputAdvancedDTO.EEHLifespam6);
                command.Parameters.AddWithValue("@EEHLifespam7", inputAdvancedDTO.EEHLifespam7);
                command.Parameters.AddWithValue("@EEHLifespam8", inputAdvancedDTO.EEHLifespam8);
                command.Parameters.AddWithValue("@EEHLifespam9", inputAdvancedDTO.EEHLifespam9);
                command.Parameters.AddWithValue("@EEHLifespam10", inputAdvancedDTO.EEHLifespam10);
                command.Parameters.AddWithValue("@EEPLifespam1", inputAdvancedDTO.EEPLifespam1);
                command.Parameters.AddWithValue("@EEPLifespam2", inputAdvancedDTO.EEPLifespam2);
                command.Parameters.AddWithValue("@EEPLifespam3", inputAdvancedDTO.EEPLifespam3);
                command.Parameters.AddWithValue("@EEPLifespam4", inputAdvancedDTO.EEPLifespam4);
                command.Parameters.AddWithValue("@EEPLifespam5", inputAdvancedDTO.EEPLifespam5);
                command.Parameters.AddWithValue("@EEPLifespam6", inputAdvancedDTO.EEPLifespam6);
                command.Parameters.AddWithValue("@EEPLifespam7", inputAdvancedDTO.EEPLifespam7);
                command.Parameters.AddWithValue("@EEPLifespam8", inputAdvancedDTO.EEPLifespam8);
                command.Parameters.AddWithValue("@EEPLifespam9", inputAdvancedDTO.EEPLifespam9);
                command.Parameters.AddWithValue("@EEPLifespam10", inputAdvancedDTO.EEPLifespam10);
                command.Parameters.AddWithValue("@EEPLifespam11", inputAdvancedDTO.EEPLifespam11);
                command.Parameters.AddWithValue("@EEPLifespam12", inputAdvancedDTO.EEPLifespam12);
                command.Parameters.AddWithValue("@EEPLifespam13", inputAdvancedDTO.EEPLifespam13);
                command.Parameters.AddWithValue("@EEPLifespam14", inputAdvancedDTO.EEPLifespam14);
                command.Parameters.AddWithValue("@EEPLifespam15", inputAdvancedDTO.EEPLifespam15);
                command.Parameters.AddWithValue("@EEPLifespam16", inputAdvancedDTO.EEPLifespam16);
                command.Parameters.AddWithValue("@EEPLifespam17", inputAdvancedDTO.EEPLifespam17);
                command.Parameters.AddWithValue("@TimeStamp", timeStamp);
                command.Connection = connect;
                int result = command.ExecuteNonQuery();
                connect.Close();
                // Check Error
                if (result < 0)
                    Console.WriteLine("Error inserting data into Database!");
            }
            //throw new NotImplementedException();
        }

        private void English(string lang)
        {
            var conn = _iconfiguration.GetSection("ConnectionStrings").GetSection("CoffeeConnStr").Value;
            var language = "";
            using (SqlConnection con = new SqlConnection(conn))
            {
                con.Open();

                SqlCommand comm = new SqlCommand("Select * from [AVFCoffee].[dbo].[User] where UserID = @UserID", con);
                //comm.Parameters.AddWithValue("@UserID", "0747ba8f-c8e3-42b6-9b48-6743583b7bb8");
                comm.Parameters.AddWithValue("@UserID", "85c6c0f5-1ae7-4664-9d9b-ce2caf94cf8e");
                //85c6c0f5 - 1ae7 - 4664 - 9d9b - ce2caf94cf8e
                // int result = command.ExecuteNonQuery();
                using (SqlDataReader reader = comm.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        language = reader["Language"].ToString();
                    }
                    reader.Close();
                }
                con.Close();
            }

            if (language != lang)
            {
                using (SqlConnection con = new SqlConnection(conn))
                {
                    con.Open();

                    SqlCommand comm = new SqlCommand("Update [AVFCoffee].[dbo].[User] set Language = @lange where UserID = @UserID", con);
                    comm.Parameters.AddWithValue("@UserID", "0747ba8f-c8e3-42b6-9b48-6743583b7bb8");
                    comm.Parameters.AddWithValue("@lange", lang);
                    // int result = command.ExecuteNonQuery();
                    comm.ExecuteNonQuery();
                    con.Close();
                }
            } 
        }

        public Dictionary<string, object> getInputs(string Language)
        {
            English(Language);
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
            workspace.Add(xls.ActiveFileName, xls);
            
            metrics.metrics(xls, md);
            language.language(xls);
            inputs_1.CreateFile(xls);
            metrics_English.MetricsEnglish(xls);
            metrics_Spanish.MetricsSpanish(xls);
            inputs_2.Inputs_2_Default(xls);
            //databaseSchema.Database_Schema(xls);
            inputs_2_Conv.Inputs_2_Conv_inputs(xls);
            inputsAdvanced2English.InputAdvanced2English(xls);
            if (Language == "EN")
            {
                inputsAdvanced2Spanish.InputAdvancedSpanish(xls, Language);
            }
            inputsEnglish.InputEnglish(xls);
            inputsSpanish.InputSpanish(xls);
            input_1.Input_1_default(xls);
            
            inputs_1_Ref.inputs1Ref(xls);
            //xls.Recalc();
            conversiones.conversiones(xls);
            gral_Conf.Gral_Conf_Summary(xls);
            conf_Summary_Spa.GeneralConfSummarySpa(xls);
            


            double earlyHectares = 0;
            double peakHectares = 0;
            double oldHectares = 0;
            bool conventional = false;
            bool organic = false;
            bool transition = false;
            double workerSalarySoles = 0;
            double productionQuintales = 0;
            double costPriceSolesPerQuintal = 0;
            double expSolesChem = 0;
            double expSolesOrg = 0;
            double transportCostSoles = 0;
            inputs.inputs(xls, earlyHectares, peakHectares, oldHectares, conventional, organic, transition, workerSalarySoles, productionQuintales, transportCostSoles,
                costPriceSolesPerQuintal, expSolesChem, expSolesOrg);
            var advancedInputsDict = new Dictionary<string, object>();
            
            if (Language == "EN")
            {

                advancedInputsDict = inAdvanced.Inputs_Advanced(xls);

            } else
            {
                advancedInputsDict = inputsAdvanced2Spanish.InputAdvancedSpanish(xls, Language); 
            }
            
            return advancedInputsDict;
        }

        public ChartInputAdvancedDTO getInputValues()
        {
            ChartInputAdvancedDTO inputsAdvanced = new ChartInputAdvancedDTO();
            var conn = _iconfiguration.GetSection("ConnectionStrings").GetSection("CoffeeConnStr").Value;
            using (SqlConnection con = new SqlConnection(conn))
            {
                con.Open();
                var id = "1234";

                SqlCommand comm = new SqlCommand("Select * from [AVFCoffee].[dbo].[UserInputsAdvanced] where UserID = @userid AND [TimeStamp] = (SELECT MAX(timestamp) FROM[AVFCoffee].[dbo].[UserInputsAdvanced] where UserID = @userid)", con);
                comm.Parameters.AddWithValue("@userid", id);
                // int result = command.ExecuteNonQuery();
                using (SqlDataReader reader = comm.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        inputsAdvanced.LGerminationSeedCollection = Convert.ToDouble(reader["LGerminationSeedCollection"].ToString());
                        inputsAdvanced.LGerminationSeedSelection = Convert.ToDouble(reader["LGerminationSeedSelection"].ToString());
                        inputsAdvanced.LGerminationNurseryConstruction = Convert.ToDouble(reader["LGerminationNurseryConstruction"].ToString());
                        inputsAdvanced.LGerminationSeedingSupportIrrigation = Convert.ToDouble(reader["LGerminationSeedingSupportIrrigation"].ToString());
                        inputsAdvanced.LGerminationOthers = Convert.ToDouble(reader["LGerminationOthers"].ToString());
                        inputsAdvanced.LNurseryConstruction = Convert.ToDouble(reader["LNurseryConstruction"].ToString());
                        inputsAdvanced.LNurseryDrawnPulled = Convert.ToDouble(reader["LNurseryDrawnPulled"].ToString());
                        inputsAdvanced.LNurseryClean = Convert.ToDouble(reader["LNurseryClean"].ToString());
                        inputsAdvanced.LNurserySoilPreparationFertilizer = Convert.ToDouble(reader["LNurserySoilPreparationFertilizer"].ToString());
                        inputsAdvanced.LNurseryFilledLockedBags = Convert.ToDouble(reader["LNurseryFilledLockedBags"].ToString());
                        inputsAdvanced.LNurseryButterflySowing = Convert.ToDouble(reader["LNurseryButterflySowing"].ToString());
                        inputsAdvanced.LNurseryIrrigation = Convert.ToDouble(reader["LNurseryIrrigation"].ToString());
                        inputsAdvanced.LNurseryFoliarApplication = Convert.ToDouble(reader["LNurseryFoliarApplication"].ToString());
                        inputsAdvanced.LNurseryReseeding = Convert.ToDouble(reader["LNurseryReseeding"].ToString());
                        inputsAdvanced.LNurseryOthers = Convert.ToDouble(reader["LNurseryOthers"].ToString());
                        inputsAdvanced.LPPFieldCleaning = Convert.ToDouble(reader["LPPFieldCleaning"].ToString());
                        inputsAdvanced.LPPCuttingTrees = Convert.ToDouble(reader["LPPCuttingTrees"].ToString());
                        inputsAdvanced.LPPWoodCollection = Convert.ToDouble(reader["LPPWoodCollection"].ToString());
                        inputsAdvanced.LPPWoodChopping = Convert.ToDouble(reader["LPPWoodChopping"].ToString());
                        inputsAdvanced.LPPCoffeeLayout = Convert.ToDouble(reader["LPPCoffeeLayout"].ToString());
                        inputsAdvanced.LPPHoleDigging = Convert.ToDouble(reader["LPPHoleDigging"].ToString());
                        inputsAdvanced.LPPSeedlingTransportation = Convert.ToDouble(reader["LPPSeedlingTransportation"].ToString());
                        inputsAdvanced.LPPSeedlingTransplant = Convert.ToDouble(reader["LPPSeedlingTransplant"].ToString());
                        inputsAdvanced.LPPShadeAdjustment = Convert.ToDouble(reader["LPPShadeAdjustment"].ToString());
                        inputsAdvanced.LPPCompostMixing = Convert.ToDouble(reader["LPPCompostMixing"].ToString());
                        inputsAdvanced.LPPOthers = Convert.ToDouble(reader["LPPOthers"].ToString());
                        inputsAdvanced.LPPYWeeding = Convert.ToDouble(reader["LPPYWeeding"].ToString());
                        inputsAdvanced.LPPYOrganic = Convert.ToDouble(reader["LPPYOrganic"].ToString());
                        inputsAdvanced.LPPYChemical = Convert.ToDouble(reader["LPPYChemical"].ToString());
                        inputsAdvanced.LPPYFoliarSpraying = Convert.ToDouble(reader["LPPYFoliarSpraying"].ToString());
                        inputsAdvanced.LPPYOther = Convert.ToDouble(reader["LPPYOther"].ToString());
                        inputsAdvanced.LHPMYManualWeeding = Convert.ToDouble(reader["LHPMYManualWeeding"].ToString());
                        inputsAdvanced.LHPMYChemicalWeeding = Convert.ToDouble(reader["LHPMYChemicalWeeding"].ToString());
                        inputsAdvanced.LHPMYOrganicFertilizers = Convert.ToDouble(reader["LHPMYOrganicFertilizers"].ToString());
                        inputsAdvanced.LHPMYChemicalFertilizers = Convert.ToDouble(reader["LHPMYChemicalFertilizers"].ToString());
                        inputsAdvanced.LHPMYFoliarSpraying = Convert.ToDouble(reader["LHPMYFoliarSpraying"].ToString());
                        inputsAdvanced.LHPMYHedgerowsConstruction = Convert.ToDouble(reader["LHPMYHedgerowsConstruction"].ToString());
                        inputsAdvanced.LHPMYShadetreePruning = Convert.ToDouble(reader["LHPMYShadetreePruning"].ToString());
                        inputsAdvanced.LHPMYPestControl = Convert.ToDouble(reader["LHPMYPestControl"].ToString());
                        inputsAdvanced.LHPMYCoffeeGrowManagement = Convert.ToDouble(reader["LHPMYCoffeeGrowManagement"].ToString());
                        inputsAdvanced.LHPMYOthers = Convert.ToDouble(reader["LHPMYOthers"].ToString());
                        inputsAdvanced.LHPHYCoffeeCollecDays = Convert.ToDouble(reader["LHPHYCoffeeCollecDays"].ToString());
                        inputsAdvanced.LHPHYAdditionDays = Convert.ToDouble(reader["LHPHYAdditionDays"].ToString());
                        inputsAdvanced.LHPPYFermentation = Convert.ToDouble(reader["LHPPYFermentation"].ToString());
                        inputsAdvanced.LHPPYWashing = Convert.ToDouble(reader["LHPPYWashing"].ToString());
                        inputsAdvanced.LHPPYDrying = Convert.ToDouble(reader["LHPPYDrying"].ToString());
                        inputsAdvanced.LHPPYScreening = Convert.ToDouble(reader["LHPPYScreening"].ToString());
                        inputsAdvanced.LHPPYSelection = Convert.ToDouble(reader["LHPPYSelection"].ToString());
                        inputsAdvanced.LHPPYStorage = Convert.ToDouble(reader["LHPPYStorage"].ToString());
                        inputsAdvanced.LHPPYCoffeewastewater = Convert.ToDouble(reader["LHPPYCoffeewastewater"].ToString());
                        inputsAdvanced.LHPPYPulpManagement = Convert.ToDouble(reader["LHPPYPulpManagement"].ToString());
                        inputsAdvanced.LHPPYOthers = Convert.ToDouble(reader["LHPPYOthers"].ToString());
                        inputsAdvanced.LHPMMManualWeeding = Convert.ToDouble(reader["LHPMMManualWeeding"].ToString());
                        inputsAdvanced.LHPMMChemicalWeeding = Convert.ToDouble(reader["LHPMMChemicalWeeding"].ToString());
                        inputsAdvanced.LHPMMOrganicFertilizers = Convert.ToDouble(reader["LHPMMOrganicFertilizers"].ToString());
                        inputsAdvanced.LHPMMChemicalFertilizers = Convert.ToDouble(reader["LHPMMChemicalFertilizers"].ToString());
                        inputsAdvanced.LHPMMFoliarSpraying = Convert.ToDouble(reader["LHPMMFoliarSpraying"].ToString());
                        inputsAdvanced.LHPMMHedgerowsConstruction = Convert.ToDouble(reader["LHPMMHedgerowsConstruction"].ToString());
                        inputsAdvanced.LHPMMShadetreePruning = Convert.ToDouble(reader["LHPMMShadetreePruning"].ToString());
                        inputsAdvanced.LHPMMPestControl = Convert.ToDouble(reader["LHPMMPestControl"].ToString());
                        inputsAdvanced.LHPMMCoffeeGrowManagement = Convert.ToDouble(reader["LHPMMCoffeeGrowManagement"].ToString());
                        inputsAdvanced.LHPMMOthers = Convert.ToDouble(reader["LHPMMOthers"].ToString());
                        inputsAdvanced.LHPHMCoffeeCollecDays = Convert.ToDouble(reader["LHPHMCoffeeCollecDays"].ToString());
                        inputsAdvanced.LHPHMAdditionDays = Convert.ToDouble(reader["LHPHMAdditionDays"].ToString());
                        inputsAdvanced.LHPPMFermentation = Convert.ToDouble(reader["LHPPMFermentation"].ToString());
                        inputsAdvanced.LHPPMWashing = Convert.ToDouble(reader["LHPPMWashing"].ToString());
                        inputsAdvanced.LHPPMDrying = Convert.ToDouble(reader["LHPPMDrying"].ToString());
                        inputsAdvanced.LHPPMScreening = Convert.ToDouble(reader["LHPPMScreening"].ToString());
                        inputsAdvanced.LHPPMSelection = Convert.ToDouble(reader["LHPPMSelection"].ToString());
                        inputsAdvanced.LHPPMStorage = Convert.ToDouble(reader["LHPPMStorage"].ToString());
                        inputsAdvanced.LHPPMCoffeewastewater = Convert.ToDouble(reader["LHPPMCoffeewastewater"].ToString());
                        inputsAdvanced.LHPPMPulpManagement = Convert.ToDouble(reader["LHPPMPulpManagement"].ToString());
                        inputsAdvanced.LHPPMOthers = Convert.ToDouble(reader["LHPPMOthers"].ToString());
                        inputsAdvanced.LHPMOManualWeeding = Convert.ToDouble(reader["LHPMOManualWeeding"].ToString());
                        inputsAdvanced.LHPMOChemicalWeeding = Convert.ToDouble(reader["LHPMOChemicalWeeding"].ToString());
                        inputsAdvanced.LHPMOOrganicFertilizers = Convert.ToDouble(reader["LHPMOOrganicFertilizers"].ToString());
                        inputsAdvanced.LHPMOChemicalFertilizers = Convert.ToDouble(reader["LHPMOChemicalFertilizers"].ToString());
                        inputsAdvanced.LHPMOFoliarSpraying = Convert.ToDouble(reader["LHPMOFoliarSpraying"].ToString());
                        inputsAdvanced.LHPMOHedgerowsConstruction = Convert.ToDouble(reader["LHPMOHedgerowsConstruction"].ToString());
                        inputsAdvanced.LHPMOShadetreePruning = Convert.ToDouble(reader["LHPMOShadetreePruning"].ToString());
                        inputsAdvanced.LHPMOPestControl = Convert.ToDouble(reader["LHPMOPestControl"].ToString());
                        inputsAdvanced.LHPMOCoffeeGrowManagement = Convert.ToDouble(reader["LHPMOCoffeeGrowManagement"].ToString());
                        inputsAdvanced.LHPMOOthers = Convert.ToDouble(reader["LHPMOOthers"].ToString());
                        inputsAdvanced.LHPHOCoffeeCollecDays = Convert.ToDouble(reader["LHPHOCoffeeCollecDays"].ToString());
                        inputsAdvanced.LHPHOAdditionDays = Convert.ToDouble(reader["LHPHOAdditionDays"].ToString());
                        inputsAdvanced.LHPPOFermentation = Convert.ToDouble(reader["LHPPOFermentation"].ToString());
                        inputsAdvanced.LHPPOWashing = Convert.ToDouble(reader["LHPPOWashing"].ToString());
                        inputsAdvanced.LHPPODrying = Convert.ToDouble(reader["LHPPODrying"].ToString());
                        inputsAdvanced.LHPPOScreening = Convert.ToDouble(reader["LHPPOScreening"].ToString());
                        inputsAdvanced.LHPPOSelection = Convert.ToDouble(reader["LHPPOSelection"].ToString());
                        inputsAdvanced.LHPPOStorage = Convert.ToDouble(reader["LHPPOStorage"].ToString());
                        inputsAdvanced.LHPPOCoffeewastewater = Convert.ToDouble(reader["LHPPOCoffeewastewater"].ToString());
                        inputsAdvanced.LHPPOPulpManagement = Convert.ToDouble(reader["LHPPOPulpManagement"].ToString());
                        inputsAdvanced.LHPPOOthers = Convert.ToDouble(reader["LHPPOOthers"].ToString());
                        inputsAdvanced.IIFood = Convert.ToDouble(reader["IIFood"].ToString());
                        inputsAdvanced.IIAdditionalTransfers = Convert.ToDouble(reader["IIAdditionalTransfers"].ToString());
                        inputsAdvanced.IIDaysoftraining = Convert.ToDouble(reader["IIDaysoftraining"].ToString());
                        inputsAdvanced.ICCreditfromcooperative = Convert.ToDouble(reader["ICCreditfromcooperative"].ToString());
                        inputsAdvanced.ICCreditfromcooperativeTime = Convert.ToDouble(reader["ICCreditfromcooperativeTime"].ToString());
                        inputsAdvanced.ICCreditfromcooperativeInterest = Convert.ToDouble(reader["ICCreditfromcooperativeInterest"].ToString());
                        inputsAdvanced.ICCreditfromagent = Convert.ToDouble(reader["ICCreditfromagent"].ToString());
                        inputsAdvanced.ICCreditfromagentTime = Convert.ToDouble(reader["ICCreditfromagentTime"].ToString());
                        inputsAdvanced.ICCreditfromagentInterest = Convert.ToDouble(reader["ICCreditfromagentInterest"].ToString());
                        //inputsAdvanced.CostGerminator = Convert.ToDouble(reader["CostGerminator"].ToString());
                        inputsAdvanced.CostGerminatorSeeds = Convert.ToDouble(reader["CostGerminatorSeeds"].ToString());
                        inputsAdvanced.CostGerminatorSeedbed = Convert.ToDouble(reader["CostGerminatorSeedbed"].ToString());
                        inputsAdvanced.CostGerminatorSandSubstrate = Convert.ToDouble(reader["CostGerminatorSandSubstrate"].ToString());
                        inputsAdvanced.CostGerminatorCalciumSulfide = Convert.ToDouble(reader["CostGerminatorCalciumSulfide"].ToString());
                        inputsAdvanced.CostGerminatorLime = Convert.ToDouble(reader["CostGerminatorLime"].ToString());
                        inputsAdvanced.CostGerminatorPlastic = Convert.ToDouble(reader["CostGerminatorPlastic"].ToString());
                        inputsAdvanced.CostGerminatorOthers = Convert.ToDouble(reader["CostGerminatorOthers"].ToString());
                        inputsAdvanced.CostNurseryFertilizer = Convert.ToDouble(reader["CostNurseryFertilizer"].ToString());
                        inputsAdvanced.CostNurseryPlasticBags = Convert.ToDouble(reader["CostNurseryPlasticBags"].ToString());
                        inputsAdvanced.CostNurseryNetting = Convert.ToDouble(reader["CostNurseryNetting"].ToString());
                        inputsAdvanced.CostNurseryStuds = Convert.ToDouble(reader["CostNurseryStuds"].ToString());
                        inputsAdvanced.CostNurseryWire = Convert.ToDouble(reader["CostNurseryWire"].ToString());
                        inputsAdvanced.CostNurseryCiclonics = Convert.ToDouble(reader["CostNurseryCiclonics"].ToString());
                        inputsAdvanced.CostNurseryStaples = Convert.ToDouble(reader["CostNurseryStaples"].ToString());
                        inputsAdvanced.CostNurserySoil = Convert.ToDouble(reader["CostNurserySoil"].ToString());
                        inputsAdvanced.CostNurseryBioFert = Convert.ToDouble(reader["CostNurseryBioFert"].ToString());
                        inputsAdvanced.CostNurseryAgroChemicals = Convert.ToDouble(reader["CostNurseryAgroChemicals"].ToString());
                        inputsAdvanced.CostNurseryFungicide = Convert.ToDouble(reader["CostNurseryFungicide"].ToString());
                        inputsAdvanced.CostNurseryPhosphoricRock = Convert.ToDouble(reader["CostNurseryPhosphoricRock"].ToString());
                        inputsAdvanced.CostNurseryOthers = Convert.ToDouble(reader["CostNurseryOthers"].ToString());
                        inputsAdvanced.CostFLPPOrganicFert = Convert.ToDouble(reader["CostFLPPOrganicFert"].ToString());
                        inputsAdvanced.CostFLPPChemicalFert = Convert.ToDouble(reader["CostFLPPChemicalFert"].ToString());
                        inputsAdvanced.CostFVGOrganicFert = Convert.ToDouble(reader["CostFVGOrganicFert"].ToString());
                        inputsAdvanced.CostFVGChemicalFert = Convert.ToDouble(reader["CostFVGChemicalFert"].ToString());
                        inputsAdvanced.CostFMOtherFert = Convert.ToDouble(reader["CostFMOtherFert"].ToString());
                        inputsAdvanced.CostFMOrganicFoliar = Convert.ToDouble(reader["CostFMOrganicFoliar"].ToString());
                        inputsAdvanced.CostFMChemicalFoliar = Convert.ToDouble(reader["CostFMChemicalFoliar"].ToString());
                        inputsAdvanced.CostFMGasFuel = Convert.ToDouble(reader["CostFMGasFuel"].ToString());
                        inputsAdvanced.CostFMOthers = Convert.ToDouble(reader["CostFMOthers"].ToString());
                        inputsAdvanced.EGEManualSprayer = Convert.ToDouble(reader["EGEManualSprayer"].ToString());
                        inputsAdvanced.EGELifespam1 = Convert.ToDouble(reader["EGELifespam1"].ToString());
                        inputsAdvanced.EGEMachetes = Convert.ToDouble(reader["EGEMachetes"].ToString());
                        inputsAdvanced.EGELifespam2 = Convert.ToDouble(reader["EGELifespam2"].ToString());
                        inputsAdvanced.EGEShovel = Convert.ToDouble(reader["EGEShovel"].ToString());
                        inputsAdvanced.EGELifespam3 = Convert.ToDouble(reader["EGELifespam3"].ToString());
                        inputsAdvanced.EGEHoe = Convert.ToDouble(reader["EGEHoe"].ToString());
                        inputsAdvanced.EGELifespam4 = Convert.ToDouble(reader["EGELifespam4"].ToString());
                        inputsAdvanced.EGEWheelBarrow = Convert.ToDouble(reader["EGEWheelBarrow"].ToString());
                        inputsAdvanced.EGELifespam5 = Convert.ToDouble(reader["EGELifespam5"].ToString());
                        inputsAdvanced.EGELime = Convert.ToDouble(reader["EGELime"].ToString());
                        inputsAdvanced.EGELifespam6 = Convert.ToDouble(reader["EGELifespam6"].ToString());
                        inputsAdvanced.EGEAuger = Convert.ToDouble(reader["EGEAuger"].ToString());
                        inputsAdvanced.EGELifespam7 = Convert.ToDouble(reader["EGELifespam7"].ToString());
                        inputsAdvanced.EGEMetalBar = Convert.ToDouble(reader["EGEMetalBar"].ToString());
                        inputsAdvanced.EGELifespam8 = Convert.ToDouble(reader["EGELifespam8"].ToString());
                        inputsAdvanced.EGEHose = Convert.ToDouble(reader["EGEHose"].ToString());
                        inputsAdvanced.EGELifespam9 = Convert.ToDouble(reader["EGELifespam9"].ToString());
                        inputsAdvanced.EGESprinklers = Convert.ToDouble(reader["EGESprinklers"].ToString());
                        inputsAdvanced.EGELifespam10 = Convert.ToDouble(reader["EGELifespam10"].ToString());
                        inputsAdvanced.EGEChainSaw = Convert.ToDouble(reader["EGEChainSaw"].ToString());
                        inputsAdvanced.EGELifespam11 = Convert.ToDouble(reader["EGELifespam11"].ToString());
                        inputsAdvanced.EGEHandSaw = Convert.ToDouble(reader["EGEHandSaw"].ToString());
                        inputsAdvanced.EGELifespam12 = Convert.ToDouble(reader["EGELifespam12"].ToString());
                        inputsAdvanced.EGEMotorPump = Convert.ToDouble(reader["EGEMotorPump"].ToString());
                        inputsAdvanced.EGELifespam13 = Convert.ToDouble(reader["EGELifespam13"].ToString());
                        inputsAdvanced.EGEPrunningScissors = Convert.ToDouble(reader["EGEPrunningScissors"].ToString());
                        inputsAdvanced.EGELifespam14 = Convert.ToDouble(reader["EGELifespam14"].ToString());
                        inputsAdvanced.EGEAxe = Convert.ToDouble(reader["EGEAxe"].ToString());
                        inputsAdvanced.EGELifespam15 = Convert.ToDouble(reader["EGELifespam15"].ToString());
                        inputsAdvanced.EEHScale = Convert.ToDouble(reader["EEHScale"].ToString());
                        inputsAdvanced.EEHLifespam1 = Convert.ToDouble(reader["EEHLifespam1"].ToString());
                        inputsAdvanced.EEHVehicle = Convert.ToDouble(reader["EEHVehicle"].ToString());
                        inputsAdvanced.EEHLifespam2 = Convert.ToDouble(reader["EEHLifespam2"].ToString());
                        inputsAdvanced.EEHWorkAnimal = Convert.ToDouble(reader["EEHWorkAnimal"].ToString());
                        inputsAdvanced.EEHLifespam3 = Convert.ToDouble(reader["EEHLifespam3"].ToString());
                        inputsAdvanced.EEHMotorcycle = Convert.ToDouble(reader["EEHMotorcycle"].ToString());
                        inputsAdvanced.EEHLifespam4 = Convert.ToDouble(reader["EEHLifespam4"].ToString());
                        inputsAdvanced.EEHBags = Convert.ToDouble(reader["EEHBags"].ToString());
                        inputsAdvanced.EEHLifespam5 = Convert.ToDouble(reader["EEHLifespam5"].ToString());
                        inputsAdvanced.EEHSack = Convert.ToDouble(reader["EEHSack"].ToString());
                        inputsAdvanced.EEHLifespam6 = Convert.ToDouble(reader["EEHLifespam6"].ToString());
                        inputsAdvanced.EEHStraw = Convert.ToDouble(reader["EEHStraw"].ToString());
                        inputsAdvanced.EEHLifespam7 = Convert.ToDouble(reader["EEHLifespam7"].ToString());
                        inputsAdvanced.EEHBaskets = Convert.ToDouble(reader["EEHBaskets"].ToString());
                        inputsAdvanced.EEHLifespam8 = Convert.ToDouble(reader["EEHLifespam8"].ToString());
                        inputsAdvanced.EEHBoxes = Convert.ToDouble(reader["EEHBoxes"].ToString());
                        inputsAdvanced.EEHLifespam9 = Convert.ToDouble(reader["EEHLifespam9"].ToString());
                        inputsAdvanced.EEHOthers = Convert.ToDouble(reader["EEHOthers"].ToString());
                        inputsAdvanced.EEHLifespam10 = Convert.ToDouble(reader["EEHLifespam10"].ToString());
                        inputsAdvanced.EEPPulperMachine = Convert.ToDouble(reader["EEPPulperMachine"].ToString());
                        inputsAdvanced.EEPLifespam1 = Convert.ToDouble(reader["EEPLifespam1"].ToString());
                        inputsAdvanced.EEPTolca = Convert.ToDouble(reader["EEPTolca"].ToString());
                        inputsAdvanced.EEPLifespam2 = Convert.ToDouble(reader["EEPLifespam2"].ToString());
                        inputsAdvanced.EEPEngine = Convert.ToDouble(reader["EEPEngine"].ToString());
                        inputsAdvanced.EEPLifespam3 = Convert.ToDouble(reader["EEPLifespam3"].ToString());
                        inputsAdvanced.EEPTanks = Convert.ToDouble(reader["EEPTanks"].ToString());
                        inputsAdvanced.EEPLifespam4 = Convert.ToDouble(reader["EEPLifespam4"].ToString());
                        inputsAdvanced.EEPWaterChannel = Convert.ToDouble(reader["EEPWaterChannel"].ToString());
                        inputsAdvanced.EEPLifespam5 = Convert.ToDouble(reader["EEPLifespam5"].ToString());
                        inputsAdvanced.EEPPVCPipes = Convert.ToDouble(reader["EEPPVCPipes"].ToString());
                        inputsAdvanced.EEPLifespam6 = Convert.ToDouble(reader["EEPLifespam6"].ToString());
                        inputsAdvanced.EEPFilteringSystem = Convert.ToDouble(reader["EEPFilteringSystem"].ToString());
                        inputsAdvanced.EEPLifespam7 = Convert.ToDouble(reader["EEPLifespam7"].ToString());
                        inputsAdvanced.EEPScreeningMachine = Convert.ToDouble(reader["EEPScreeningMachine"].ToString());
                        inputsAdvanced.EEPLifespam8 = Convert.ToDouble(reader["EEPLifespam8"].ToString());
                        inputsAdvanced.EEPDesmucilaginador = Convert.ToDouble(reader["EEPDesmucilaginador"].ToString());
                        inputsAdvanced.EEPLifespam9 = Convert.ToDouble(reader["EEPLifespam9"].ToString());
                        inputsAdvanced.EEPMotorPump = Convert.ToDouble(reader["EEPMotorPump"].ToString());
                        inputsAdvanced.EEPLifespam10 = Convert.ToDouble(reader["EEPLifespam10"].ToString());
                        inputsAdvanced.EEPOthersWetInput = Convert.ToDouble(reader["EEPOthersWetInput"].ToString());
                        inputsAdvanced.EEPLifespam11 = Convert.ToDouble(reader["EEPLifespam11"].ToString());
                        inputsAdvanced.EEPConcrete = Convert.ToDouble(reader["EEPConcrete"].ToString());
                        inputsAdvanced.EEPLifespam12 = Convert.ToDouble(reader["EEPLifespam12"].ToString());
                        inputsAdvanced.EEPPlastic = Convert.ToDouble(reader["EEPPlastic"].ToString());
                        inputsAdvanced.EEPLifespam13 = Convert.ToDouble(reader["EEPLifespam13"].ToString());
                        inputsAdvanced.EEPRake = Convert.ToDouble(reader["EEPRake"].ToString());
                        inputsAdvanced.EEPLifespam14 = Convert.ToDouble(reader["EEPLifespam14"].ToString());
                        inputsAdvanced.EEPBroom = Convert.ToDouble(reader["EEPBroom"].ToString());
                        inputsAdvanced.EEPLifespam15 = Convert.ToDouble(reader["EEPLifespam15"].ToString());
                        inputsAdvanced.EEPStorageRoom = Convert.ToDouble(reader["EEPStorageRoom"].ToString());
                        inputsAdvanced.EEPLifespam16 = Convert.ToDouble(reader["EEPLifespam16"].ToString());
                        inputsAdvanced.EEPOthersDryInput = Convert.ToDouble(reader["EEPOthersDryInput"].ToString());
                        inputsAdvanced.EEPLifespam17 = Convert.ToDouble(reader["EEPLifespam17"].ToString());
                        inputsAdvanced.ACCApplicationFee = Convert.ToDouble(reader["ACCApplicationFee"].ToString());
                        inputsAdvanced.ACCAnnualMembership = Convert.ToDouble(reader["ACCAnnualMembership"].ToString());
                        inputsAdvanced.ACCLifeInsurance = Convert.ToDouble(reader["ACCLifeInsurance"].ToString());
                        inputsAdvanced.ACCFLOCertification = Convert.ToDouble(reader["ACCFLOCertification"].ToString());
                        inputsAdvanced.ACCOrganicCertification = Convert.ToDouble(reader["ACCOrganicCertification"].ToString());
                        inputsAdvanced.ACLLandValue = Convert.ToDouble(reader["ACLLandValue"].ToString());
                        inputsAdvanced.ACLPropertyTax = Convert.ToDouble(reader["ACLPropertyTax"].ToString());
                        inputsAdvanced.ACUSuperviseInvest = Convert.ToDouble(reader["ACUSuperviseInvest"].ToString());
                        inputsAdvanced.ACUAdministInvest = Convert.ToDouble(reader["ACUAdministInvest"].ToString());
                        inputsAdvanced.ACUTrainingInvest = Convert.ToDouble(reader["ACUTrainingInvest"].ToString());
                        inputsAdvanced.ACUExtraOrdInvest = Convert.ToDouble(reader["ACUExtraOrdInvest"].ToString());
                        inputsAdvanced.TGSeedPurchase = Convert.ToDouble(reader["TGSeedPurchase"].ToString());
                        inputsAdvanced.TGWoodTransportation = Convert.ToDouble(reader["TGWoodTransportation"].ToString());
                        inputsAdvanced.TGSandTransportation = Convert.ToDouble(reader["TGSandTransportation"].ToString());
                        inputsAdvanced.TGOthers = Convert.ToDouble(reader["TGOthers"].ToString());
                        inputsAdvanced.TNSoilTransportation = Convert.ToDouble(reader["TNSoilTransportation"].ToString());
                        inputsAdvanced.TNSacksMaterialShopping = Convert.ToDouble(reader["TNSacksMaterialShopping"].ToString());
                        inputsAdvanced.TNOthers = Convert.ToDouble(reader["TNOthers"].ToString());
                        inputsAdvanced.TLPWoodTransportation = Convert.ToDouble(reader["TLPWoodTransportation"].ToString());
                        inputsAdvanced.TLPCompostTransportation = Convert.ToDouble(reader["TLPCompostTransportation"].ToString());
                        inputsAdvanced.TLPPlantTransportation = Convert.ToDouble(reader["TLPPlantTransportation"].ToString());
                        inputsAdvanced.TLPOthers = Convert.ToDouble(reader["TLPOthers"].ToString());
                        inputsAdvanced.TOtherEquipment = Convert.ToDouble(reader["TOtherEquipment"].ToString());
                        inputsAdvanced.TOtherLaborTransportation = Convert.ToDouble(reader["TOtherLaborTransportation"].ToString());
                        inputsAdvanced.TOtherCoffeeTransportation = Convert.ToDouble(reader["TOtherCoffeeTransportation"].ToString());
                        inputsAdvanced.TOtherSupervisingActivities = Convert.ToDouble(reader["TOtherSupervisingActivities"].ToString());
                        inputsAdvanced.TOthers = Convert.ToDouble(reader["TOthers"].ToString());
                    }
                }
                con.Close();
            }
            return inputsAdvanced;
        }

        public List<AnalysisDTO> GetAnalysis(string userID)
        {
            
            var analyses = new List<AnalysisDTO>();
            var conn = _iconfiguration.GetSection("ConnectionStrings").GetSection("CoffeeConnStr").Value;
            using (SqlConnection con = new SqlConnection(conn))
            {
                con.Open();
                SqlCommand comm = new SqlCommand("Select * from [AVFCoffee].[dbo].[Analysis] where [UserID] = @userid", con);
                comm.Parameters.AddWithValue("@userid", userID);
                // int result = command.ExecuteNonQuery();
                using (SqlDataReader reader = comm.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        AnalysisDTO analysis = new AnalysisDTO();
                        analysis.TimeStamp = reader["TimeStamp"].ToString();
                        analysis.Title = reader["Title"].ToString();
                        analyses.Add(analysis);
                    }
                }
                con.Close();
            }
            foreach (AnalysisDTO analysis in analyses)
            {
                var inputList = new List<string>();
                inputList.Add("farm1");
                inputList.Add("farm2");
                inputList.Add("farm3");
                analysis.Input = inputList;
            }
            return analyses;
        }

        public List<FarmInfoDTO> GetFarms(string coopid)
        {
            
            var farms = new List<FarmInfoDTO>();
            var conn = _iconfiguration.GetSection("ConnectionStrings").GetSection("CoffeeConnStr").Value;
            using (SqlConnection con = new SqlConnection(conn))
            {
                con.Open();
                SqlCommand comm = new SqlCommand("Select * from [AVFCoffee].[dbo].[SmallHolder] where CoopID = @coopid", con);
                comm.Parameters.AddWithValue("@coopid", "1234");
                // int result = command.ExecuteNonQuery();
                using (SqlDataReader reader = comm.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        FarmInfoDTO farmInfo = new FarmInfoDTO();
                        farmInfo.FarmName = reader["FarmName"].ToString();
                        farmInfo.Region = reader["Region"].ToString();
                        farmInfo.Elevation = Convert.ToInt32(reader["Elevation"].ToString());
                        farms.Add(farmInfo);
                    }
                }
                con.Close();
            }
            return farms;
        }
    }


}
