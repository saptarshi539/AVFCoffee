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
            } else if (data[0] == "Quintales")
            {
                coffeemeasurequintales = true;
            } else if (data[0] == "Arrobas")
            {
                coffeemeasurearrobas = true;
            } else if (data[0] == "Cargas")
            {
                coffeemeasurecargas = true;
            }

            if (data[1] == "Meters")
            {
                lengthmeasuremeters = true;
            } else if (data[1] == "Feet")
            {
                lengthmeasurefeet = true;
            }

            if (data[2] == "Hectares")
            {
                farmareameasurehectares = true;
            } else if (data[2] == "Manzanas")
            {
                farmareameasuremanzanas = true;
            }

            if (data[3] == "Kilograms")
            {
                applicationmeasurekilograms = true;
            } else if (data[3] == "Pounds")
            {
                applicationmeasurepounds = true;
            }

            if (data[4] == "Liters")
            {
                capacitymeasureliters = true;
            } else if (data[4] == "Gallons")
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
            } else if (data[5] == "USD")
            {
                currencyusdollar = true;
            } else if (data[5] == "Venezuelan Bolivar")
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

        public void saveUserAdvancedInputs()
        {
            ChartInputAdvancedDTO inputAdvancedDTO = new ChartInputAdvancedDTO();
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
                   ",[TOtherCoffeeTransportation],[TOtherSupervisingActivities],[TOthers]) VALUES" +
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
                   ",@TOtherCoffeeTransportation,@TOtherSupervisingActivities,@TOthers))");
            using (SqlConnection connect = new SqlConnection(conn))
            {
                connect.Open();
                SqlCommand command = new SqlCommand(sqlQuery);
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
    }
}
