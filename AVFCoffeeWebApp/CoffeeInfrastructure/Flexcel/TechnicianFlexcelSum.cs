using CoffeeCore.DTO;
using CoffeeCore.Interfaces;
using FlexCel.Core;
using FlexCel.XlsAdapter;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;

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

            if (data[0] == "Kilograms")
            {
                coffeemeasurekilograms = true;
            }
            else if (data[0] == "Pounds")
            {
                coffeemeasurepounds = true;
            }
            else if (data[0] == "Quintales")
            {
                coffeemeasurequintales = true;
            }
            else if (data[0] == "Arrobas")
            {
                coffeemeasurearrobas = true;
            }
            else if (data[0] == "Cargas")
            {
                coffeemeasurecargas = true;
            }

            if (data[1] == "Meters")
            {
                lengthmeasuremeters = true;
            }
            else if (data[1] == "Feet")
            {
                lengthmeasurefeet = true;
            }

            if (data[2] == "Hectares")
            {
                farmareameasurehectares = true;
            }
            else if (data[2] == "Manzanas")
            {
                farmareameasuremanzanas = true;
            }

            if (data[3] == "Kilograms")
            {
                applicationmeasurekilograms = true;
            }
            else if (data[3] == "Pounds")
            {
                applicationmeasurepounds = true;
            }

            if (data[4] == "Liters")
            {
                capacitymeasureliters = true;
            }
            else if (data[4] == "Gallons")
            {
                capacitymeasuregallons = true;
            }


            if (data[5] == "Bolivian Boliviano")
            {
                currencyboliviaboliviano = true;
            }
            else if (data[5] == "Brazilian Real")
            {
                currencybrazilreal = true;
            }
            else if (data[5] == "Colombian Peso")
            {
                currencycolombiapeso = true;
            }
            else if (data[5] == "Costa Rican Colon")
            {
                currencycostaricacolon = true;
            }
            else if (data[5] == "Cuban Peso")
            {
                currencycubapeso = true;
            }
            else if (data[5] == "Guatemalan Quetzal")
            {
                currencyguatemalaquetzal = true;
            }
            else if (data[5] == "Jamaican Dollar")
            {
                currencyjamaicadollar = true;
            }
            else if (data[5] == "Honduran Lempira")
            {
                currencyhonduraslempira = true;
            }
            else if (data[5] == "Haitian Gourde")
            {
                currencyhaitigourde = true;
            }
            else if (data[5] == "Mexican Peso")
            {
                currencymexicopeso = true;
            }
            else if (data[5] == "Nicaraguan Cordoba")
            {
                currencynicaraguacordoba = true;
            }
            else if (data[5] == "Peruvian Sol")
            {
                currencyperusol = true;
            }
            else if (data[5] == "USD")
            {
                currencyusdollar = true;
            }
            else if (data[5] == "Venezuelan Bolivar")
            {
                currencyvenezuelabolivar = true;
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

            //throw new NotImplementedException();
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
                   ",[TOtherCoffeeTransportation],[TOtherSupervisingActivities],[TOthers],[UserID]) VALUES" +
                   "(@TimeStamp,@LGerminationSeedCollection,@LGerminationSeedSelection,@LGerminationNurseryConstruction, " +
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
                   ",@TOtherCoffeeTransportation,@TOtherSupervisingActivities,@TOthers,@UserID))");
            using (SqlConnection connect = new SqlConnection(conn))
            {
                connect.Open();
                SqlCommand command = new SqlCommand(sqlQuery);
                command.Parameters.AddWithValue("@UserID", "1234");
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
                command.Parameters.AddWithValue("@TimeStamp", timeStamp);
                command.Connection = connect;
                int result = command.ExecuteNonQuery();
                connect.Close();
                // Check Error
                if (result < 0)
                    Console.WriteLine("Error inserting data into Database!");
            }
            throw new NotImplementedException();
        }

        public Dictionary<string, object> getInputs()
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
                }
            }
            MetricsDTO md = new MetricsDTO();
            if (minput.applicationmeasurekilograms)
            {
                md.applicationmeasurekilograms = 1;
            } else
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
            XlsFile xls = new XlsFile(true);
            TWorkspace workspace = new TWorkspace();

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
            Inputs_1_Ref inputs_1_Ref = new Inputs_1_Ref();
            Inputs inputs = new Inputs();
            workspace.Add("Coffee Interactive Tool 2.0 08_10_18.xlsx", xls);
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
            //inputsAdvanced2English.InputAdvanced2English(xls);
            ////xls.Recalc();
            //inputsAdvanced2Spanish.InputAdvancedSpanish(xls);
            //xls.Recalc();
            language.language(xls);
            //xls.Recalc();
            metrics_English.MetricsEnglish(xls);
            //xls.Recalc();
            metrics_Spanish.MetricsSpanish(xls);
            //xls.Recalc();
            outcomeLAdjustment.Outcome_L_Adjustment(xls);
            //xls.Recalc();
            outcomeYAdjustment.Outcome_Y_Adjustment(xls);
            //xls.Recalc();
            output1_Pre_Metric_Currency.Output1PreMetricCurrency(xls);
            //xls.Recalc();
            outcomeTotalAdj.Outcome_TOTAL_Adj(xls);
            //xls.Recalc();
            inputs_1_Ref.inputs1Ref(xls);
            //xls.Recalc();
            conversiones.conversiones(xls);
            //xls.Recalc(true);
            metrics.metrics(xls, md);
            //xls.Recalc();

            //xls.Recalc();
            gral_Conf.Gral_Conf_Summary(xls);
            //xls.Recalc();



            //inputs.inputs(xls, earlyHectares, peakHectares, oldHectares, conventional, organic, transition, workerSalarySoles, productionQuintales, transportCostSoles,
            //    costPriceSolesPerQuintal, expSolesChem, expSolesOrg);
            var advancedInputsDict = inAdvanced.Inputs_Advanced(xls);


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

                SqlCommand comm = new SqlCommand("Select * from [AVFCoffee].[dbo].[UserInputsAdvanced] where UserID = @userid", con);
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
    }
}
