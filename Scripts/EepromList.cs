using DocumentFormat.OpenXml.InkML;
using DocumentFormat.OpenXml.Office.CustomUI;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EepromReader.Scripts
{
    public static class EepromList
    {
        public static List<EepromMapping> mapping { get; private set; }
        public static List<EepromMapping> unfoldedMapping { get; private set; }

        static EepromList()
        {
            InitializeMapping();
            InitializeUnfoldedMapping();
        }

        private static void InitializeMapping()
        {
            mapping = new List<EepromMapping>()
                {
                    new EepromMapping { Address = 0, Name = "EE_Dummy", Length = 1, EepromDataType = typeof(byte) },

                    // RF Settings
                    new EepromMapping { Address = 1, Name = "EE_ZenderSterkte", Length = 1, EepromDataType = typeof(byte) },

                    new EepromMapping { Address = 2, Name = "EE_KanaalZender", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3, Name = "EE_ZenderAanUit", Length = 1, EepromDataType = typeof(byte) },

                    // Version number and reset counter
                    new EepromMapping { Address = 4, Name = "EE_Versie", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 6, Name = "EE_ResetTeller", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 8, Name = "EE_RitStatusMode", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 10, Name = "EE_DummyRitStatus", Length = 10, EepromDataType = typeof(byte[]) },

                    // Ingangen d1 en d2
                    new EepromMapping { Address = 19, Name = "EE_install_contact", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 20, Name = "EE_alarm_contact", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 21, Name = "EE_melding_contact", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 22, Name = "EE_status_contact", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 23, Name = "EE_status_rust_contact", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 24, Name = "EE_Mode_contact", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 25, Name = "EE_AfVertraging_contact", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 27, Name = "EE_OpVertraging_contact", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 29, Name = "EE_Dummy_Contact_Digital", Length = 10, EepromDataType = typeof(byte[]) },

                    // Accu Monitoring
                    new EepromMapping { Address = 39, Name = "EE_AccuSpanningMax", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 41, Name = "EE_AccuSpanningMin", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 43, Name = "EE_AccuAlarmVertraging", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 44, Name = "EE_BackupBatVoltChargeOn", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 46, Name = "EE_BackupBatVoltChargeOff", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 48, Name = "EE_BackupBatVoltAlarm", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 50, Name = "EE_DummyAccu", Length = 4, EepromDataType = typeof(byte[]) },

                    // kl Can Temp sensor instellingen
                    new EepromMapping { Address = 54, Name = "EE_KlTempSensorNummerExt", Length = 16, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 70, Name = "EE_KlTempSensorLogExt", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 72, Name = "EE_KlTempSensorAlarmExt", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 74, Name = "EE_KlTempSensorRegelingExt", Length = 2, EepromDataType = typeof(byte[]) },

                    // 2D Array of bytes (16x6 grid for sensor data)
                    new EepromMapping { Address = 76, Name = "EE_KlTempSensorIDExt", Length = 96, EepromDataType = typeof(byte[,]), RowCount = 16, ColumnCount = 6 },

                    // LTE Bands for Cat M1
                    new EepromMapping { Address = 172, Name = "EE_LteCatM1Band001to032", Length = 4, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 176, Name = "EE_LteCatM1Band033to064", Length = 4, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 180, Name = "EE_LteCatM1Band065to096", Length = 4, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 184, Name = "EE_LteCatM1Band097to128", Length = 4, EepromDataType = typeof(byte[]) },

                    // LTE Bands for NB IoT
                    new EepromMapping { Address = 188, Name = "EE_LteNbIotBand001to032", Length = 4, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 192, Name = "EE_LteNbIotBand033to064", Length = 4, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 196, Name = "EE_LteNbIotBand065to096", Length = 4, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 200, Name = "EE_LteNbIotBand097to128", Length = 4, EepromDataType = typeof(byte[]) },

                    // LTE Technology Support
                    new EepromMapping { Address = 204, Name = "EE_LteTechSup", Length = 1, EepromDataType = typeof(byte) },

                    // 2G Priority
                    new EepromMapping { Address = 205, Name = "EE_2G_Prio", Length = 1, EepromDataType = typeof(byte) },

                    // Telit Technology Enable
                    new EepromMapping { Address = 206, Name = "EE_TelitTechEnable", Length = 1, EepromDataType = typeof(byte) },

                    // Access Technology Registered Network
                    new EepromMapping { Address = 207, Name = "EE_AccessTechnologyRegisteredNetwork", Length = 1, EepromDataType = typeof(byte) },

                    // Reserved Temp (1D Array of bytes, 14 elements)
                    new EepromMapping { Address = 208, Name = "ReserveTemp", Length = 14, EepromDataType = typeof(byte[]) },

                    // CAN Switch Sensor Settings
                    new EepromMapping { Address = 222, Name = "EE_KlSwitchSensorNummerExt", Length = 16, EepromDataType = typeof(byte[]) },

                    // Status Sensor in Switch (1D Array, 2 bytes)
                    new EepromMapping { Address = 238, Name = "EE_StatusSensorInSwitchExt", Length = 2, EepromDataType = typeof(byte[]) },

                    // Rust Sensor in Switch (1D Array, 2 bytes)
                    new EepromMapping { Address = 240, Name = "EE_RustSensorInSwitchExt", Length = 2, EepromDataType = typeof(byte[]) },

                    // Alarm Sensor in Switch (1D Array, 2 bytes)
                    new EepromMapping { Address = 242, Name = "EE_AlarmSensorInSwitchExt", Length = 2, EepromDataType = typeof(byte[]) },

                    // Melding Sensor in Switch (1D Array, 2 bytes)
                    new EepromMapping { Address = 244, Name = "EE_MeldingSensorInSwitchExt", Length = 2, EepromDataType = typeof(byte[]) },

                    // 2D Array of Switch Sensor IDs (16x6 grid)
                    new EepromMapping { Address = 246, Name = "EE_KlSwitchSensorIDExt", Length = 96, EepromDataType = typeof(byte[,]), RowCount = 16, ColumnCount = 6 },

                    // GSM-GPRS Settings
                    new EepromMapping { Address = 342, Name = "EE_gsm_pin", Length = 6, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 348, Name = "EE_gsm_pdp_type", Length = 5, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 353, Name = "EE_gsm_pdp_addr", Length = 21, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 374, Name = "EE_GPRS_UserNaam", Length = 21, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 395, Name = "EE_GPRS_Password", Length = 21, EepromDataType = typeof(byte[]) },

                    // Server Port (as byte array)
                    new EepromMapping { Address = 416, Name = "EE_ServerPort", Length = 2, EepromDataType = typeof(byte[]) },

                    // Web Address
                    new EepromMapping { Address = 418, Name = "EE_WebAddr", Length = 21, EepromDataType = typeof(byte[]) }, // Staat als address 448 in de Eeprom, maar dat is fout

                    // GSM PIN Code OK status
                    new EepromMapping { Address = 439, Name = "EE_GSM_PINCODE_OKE", Length = 1, EepromDataType = typeof(byte) },

                    // GSM Enable Flags
                    new EepromMapping { Address = 440, Name = "EE_GSM_Enable", Length = 1, EepromDataType = typeof(byte) },

                    // GSM Phone Numbers and SMS Settings
                    new EepromMapping { Address = 441, Name = "EE_GSM_Telnr", Length = 13, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 454, Name = "EE_GSM_SMSnr", Length = 13, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 467, Name = "EE_GSM_SMS_Service", Length = 14, EepromDataType = typeof(byte[]) },

                    // GSM CCID
                    new EepromMapping { Address = 481, Name = "EE_GSM_CCID", Length = 26, EepromDataType = typeof(byte[]) },

                    // GSM Band
                    new EepromMapping { Address = 507, Name = "EE_GSM_BAND", Length = 1, EepromDataType = typeof(byte) },

                    // Source Server Port (as byte array)
                    new EepromMapping { Address = 508, Name = "EE_ServerPortSource", Length = 2, EepromDataType = typeof(byte[]) },

                    // Reserved GSM Data
                    new EepromMapping { Address = 509, Name = "EE_ResGSM2", Length = 20, EepromDataType = typeof(byte[]) },

                    // RTC Offset (as byte)
                    new EepromMapping { Address = 530, Name = "EE_RTCOffsetGMT", Length = 1, EepromDataType = typeof(byte) },

                    // Ventilation Control Settings
                    new EepromMapping { Address = 531, Name = "EE_RegelTijd", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 532, Name = "EE_VentielatieRegelTijd", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 533, Name = "EE_ToevoerVentielatieEnable", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 534, Name = "EE_ToevoerVentielatieMax", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 535, Name = "EE_ToevoerVentielatieMin", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 536, Name = "EE_ToevoerVentielatieBandbreedte", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 537, Name = "EE_ToevoerVentielatieBBCompensatie", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 538, Name = "EE_ToevoerVentielatieBBCompStart", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 540, Name = "EE_AfvoerVentielatieEnable", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 541, Name = "EE_AfvoerVentielatieMax", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 542, Name = "EE_AfvoerVentielatieMin", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 543, Name = "EE_AfvoerVentielatieBandbreedte", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 544, Name = "EE_AfvoerVentielatieBBCompensatie", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 545, Name = "EE_AfvoerVentielatieBBCompStart", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 547, Name = "EE_MaxVentielatieLoopTijd", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 548, Name = "EE_IngesteldRegelTemperatuur", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 550, Name = "EE_VentielatieMultiInOutPcbNr", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 551, Name = "EE_VentielatieMultiInOutEnable", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 552, Name = "EE_VerwarmingMode", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 553, Name = "EE_KoelingMode", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 554, Name = "EE_VerwarmingDiff", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 555, Name = "EE_KoelingDiff", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 556, Name = "EE_VerwarmingTemp", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 558, Name = "EE_KoelingTemp", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 560, Name = "EE_VerwarmingTempOffset", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 561, Name = "EE_KoelingTempOffset", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 562, Name = "EE_BuitenVoeler", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 563, Name = "EE_ToevoerVMax", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 564, Name = "EE_ToevoerVMin", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 565, Name = "EE_AfvoerVMax", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 566, Name = "EE_AfvoerVMin", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 567, Name = "EE_MultiInOutAnalogStepTime", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 568, Name = "EE_ResVent", Length = 2, EepromDataType = typeof(byte[]) },

                    // Generator Control Settings
                    new EepromMapping { Address = 570, Name = "EE_GeneratorAanAccuSpanning", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 572, Name = "EE_GeneratorAanVertraging", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 574, Name = "EE_GeneratorAanTijd", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 575, Name = "EE_GeneratorRegelingEnable", Length = 1, EepromDataType = typeof(byte) },

                    // Vent Settings
                    new EepromMapping { Address = 576, Name = "EE_VentUitTemp", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 578, Name = "EE_VentUitSchakelDiff", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 580, Name = "EE_Groot_Dummy", Length = 38, EepromDataType = typeof(byte[]) },

                    // GPS and Logging Settings 
                    new EepromMapping { Address = 618, Name = "EE_GpsFixDelay", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 620, Name = "EE_LogBySpeedDelay", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 622, Name = "EE_MaxGpsDopRead", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 624, Name = "EE_GpsMinDop", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 626, Name = "EE_NaarStandby", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 627, Name = "EE_LogTijdNormaal", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 629, Name = "EE_LogTijdAlarm", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 631, Name = "EE_AlarmVertragingTemp", Length = 1, EepromDataType = typeof(byte) },

                    // Thermostat Settings
                    new EepromMapping { Address = 632, Name = "EE_ThermostaatIngesteldeTemp", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 634, Name = "EE_TempSensorTermostaat", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 636, Name = "EE_RegelingTypeTermostaat", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 637, Name = "EE_SD_Termostaat", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 638, Name = "EE_NZ_Termostaat", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 639, Name = "EE_RelaisTermostaatMode", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 640, Name = "EE_TermostaatRegelingCondition", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 641, Name = "EE_ThermostaatMultiInOutNr", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 642, Name = "EE_ToevoerVentStartPerc", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 643, Name = "EE_ToevoerVentStarttijd", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 644, Name = "EE_AfvoerVentStartPerc", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 645, Name = "EE_AfvoerVentStarttijd", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 646, Name = "EE_Groot_Dummy2", Length = 36, EepromDataType = typeof(byte[]) },

                    // Alarm settings
                    new EepromMapping { Address = 682, Name = "EE_AlarmAbsoluutMin", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 684, Name = "EE_AlarmAbsoluutMax", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 686, Name = "EE_Groot_Dummy3", Length = 50, EepromDataType = typeof(byte[]) },

                    // RHI settings
                    new EepromMapping { Address = 736, Name = "EE_MU_AlarmTxtD1", Length = 26, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 762, Name = "EE_MU_AlarmTxtD2", Length = 26, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 788, Name = "EE_MU_AlarmTxtExt1", Length = 26, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 814, Name = "EE_MU_AlarmTxtExt2", Length = 26, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 840, Name = "EE_MU_AlarmTxtExt3", Length = 26, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 866, Name = "EE_MU_AlarmTxtExt4", Length = 26, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 892, Name = "EE_MU_AlarmTxtExt5", Length = 26, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 918, Name = "EE_MU_AlarmTxtExt6", Length = 26, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 944, Name = "EE_MU_AlarmTxtExt7", Length = 26, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 970, Name = "EE_MU_AlarmTxtExt8", Length = 26, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 996, Name = "EE_MU_AlarmTxtExt9", Length = 26, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 1022, Name = "EE_MU_AlarmTxtExt10", Length = 26, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 1048, Name = "EE_MU_AlarmTxtExt11", Length = 26, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 1074, Name = "EE_MU_AlarmTxtExt12", Length = 26, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 1100, Name = "EE_MU_AlarmTxtExt13", Length = 26, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 1126, Name = "EE_MU_AlarmTxtExt14", Length = 26, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 1152, Name = "EE_MU_AlarmTxtExt15", Length = 26, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 1178, Name = "EE_MU_AlarmTxtExt16", Length = 26, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 1204, Name = "EE_Groot_Dummy4", Length = 200, EepromDataType = typeof(byte[]) },

                    // Ingang settings (D1, D2, Ext1-Ext16)
                    new EepromMapping { Address = 1404, Name = "EE_MU_IngangD1Naam", Length = 15, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 1419, Name = "EE_MU_IngangD2Naam", Length = 15, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 1434, Name = "EE_MU_IngangExt1Naam", Length = 15, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 1449, Name = "EE_MU_IngangExt2Naam", Length = 15, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 1464, Name = "EE_MU_IngangExt3Naam", Length = 15, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 1479, Name = "EE_MU_IngangExt4Naam", Length = 15, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 1494, Name = "EE_MU_IngangExt5Naam", Length = 15, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 1509, Name = "EE_MU_IngangExt6Naam", Length = 15, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 1524, Name = "EE_MU_IngangExt7Naam", Length = 15, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 1539, Name = "EE_MU_IngangExt8Naam", Length = 15, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 1554, Name = "EE_MU_IngangExt9Naam", Length = 15, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 1569, Name = "EE_MU_IngangExt10Naam", Length = 15, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 1584, Name = "EE_MU_IngangExt11Naam", Length = 15, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 1599, Name = "EE_MU_IngangExt12Naam", Length = 15, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 1614, Name = "EE_MU_IngangExt13Naam", Length = 15, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 1629, Name = "EE_MU_IngangExt14Naam", Length = 15, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 1644, Name = "EE_MU_IngangExt15Naam", Length = 15, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 1659, Name = "EE_MU_IngangExt16Naam", Length = 15, EepromDataType = typeof(byte[]) },

                    // Temperature Sensor External Names (16 entries, each 15 bytes)
                    new EepromMapping { Address = 1674, Name = "EE_MU_IngangTempSensorExt", Length = 240, EepromDataType = typeof(byte[,]), RowCount = 16, ColumnCount = 15 },

                    // Additional settings and dummy arrays
                    new EepromMapping { Address = 1914, Name = "EE_Groot_Dummy5", Length = 200, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 2114, Name = "EE_MU_ToevoerVentNaam", Length = 15, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 2129, Name = "EE_MU_AfvoerVentNaam", Length = 15, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 2144, Name = "EE_MU_WaterNivauNaam", Length = 15, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 2159, Name = "EE_Groot_Dummy6", Length = 185, EepromDataType = typeof(byte[]) },

                    // Pictogram settings for various switches
                    new EepromMapping { Address = 2344, Name = "EE_MU_PictogramD1", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2345, Name = "EE_MU_PictogramD2", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2346, Name = "EE_MU_PictogramSwitchExt1", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2347, Name = "EE_MU_PictogramSwitchExt2", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2348, Name = "EE_MU_PictogramSwitchExt3", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2349, Name = "EE_MU_PictogramSwitchExt4", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2350, Name = "EE_MU_PictogramSwitchExt5", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2351, Name = "EE_MU_PictogramSwitchExt6", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2352, Name = "EE_MU_PictogramSwitchExt7", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2353, Name = "EE_MU_PictogramSwitchExt8", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2354, Name = "EE_MU_PictogramSwitchExt9", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2355, Name = "EE_MU_PictogramSwitchExt10", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2356, Name = "EE_MU_PictogramSwitchExt11", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2357, Name = "EE_MU_PictogramSwitchExt12", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2358, Name = "EE_MU_PictogramSwitchExt13", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2359, Name = "EE_MU_PictogramSwitchExt14", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2360, Name = "EE_MU_PictogramSwitchExt15", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2361, Name = "EE_MU_PictogramSwitchExt16", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2362, Name = "EE_MU_PictogramWaterNivau", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2363, Name = "EE_SmsSelectNew", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 2365, Name = "EE_SmsSendMaxProDayEnable", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2366, Name = "EE_SmsSendMaxProDay", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2367, Name = "EE_Groot_Dummy7", Length = 95, EepromDataType = typeof(byte[]) },

                    // Temperature Pictogram settings
                    new EepromMapping { Address = 2462, Name = "EE_MU_PictorgramTempExt1", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2463, Name = "EE_MU_PictorgramTempExt2", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2464, Name = "EE_MU_PictorgramTempExt3", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2465, Name = "EE_MU_PictorgramTempExt4", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2466, Name = "EE_MU_PictorgramTempExt5", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2467, Name = "EE_MU_PictorgramTempExt6", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2468, Name = "EE_MU_PictorgramTempExt7", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2469, Name = "EE_MU_PictorgramTempExt8", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2470, Name = "EE_MU_PictorgramTempExt9", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2471, Name = "EE_MU_PictorgramTempExt10", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2472, Name = "EE_MU_PictorgramTempExt11", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2473, Name = "EE_MU_PictorgramTempExt12", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2474, Name = "EE_MU_PictorgramTempExt13", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2475, Name = "EE_MU_PictorgramTempExt14", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2476, Name = "EE_MU_PictorgramTempExt15", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2477, Name = "EE_MU_PictorgramTempExt16", Length = 1, EepromDataType = typeof(byte) },

                    // General and water level settings
                    new EepromMapping { Address = 2478, Name = "EE_WachdagReset", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2479, Name = "EE_WaterNivauEnable", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2480, Name = "EE_WaterNivauMin", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 2482, Name = "EE_WaterNivauMax", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 2484, Name = "EE_WaterNivauAlarm", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2485, Name = "EE_waterNivauAlarmVertraging", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2486, Name = "EE_WaterTemperatuurSensor", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 2488, Name = "EE_WaterTemperatuurMax", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 2490, Name = "EE_WaterTemperatuurMin", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 2492, Name = "EE_WaterTemperatuurAlarmVertraging", Length = 1, EepromDataType = typeof(byte) },

                    // SMS and network settings
                    new EepromMapping { Address = 2493, Name = "EE_SmsSelect", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2494, Name = "EE_SmsNrsEnNaam", Length = 300, EepromDataType = typeof(byte[,,]), Depth = 2, RowCount = 10, ColumnCount = 15 },
                    new EepromMapping { Address = 2794, Name = "EE_BelSelect", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2795, Name = "EE_GsmNetworkName", Length = 20, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 2815, Name = "EE_GsmImei", Length = 18, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 2833, Name = "EE_GsmNetworkId", Length = 12, EepromDataType = typeof(byte[]) },

                    // Company information
                    new EepromMapping { Address = 2845, Name = "EE_CompanyName", Length = 26, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 2871, Name = "EE_CompanyInfo1", Length = 26, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 2897, Name = "EE_CompanyInfo2", Length = 26, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 2923, Name = "EE_CompanyInfo3", Length = 26, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 2949, Name = "EE_CompanyInfo4", Length = 26, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 2975, Name = "EE_CompanyInfo5", Length = 26, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3001, Name = "EE_CompanyInfo6", Length = 26, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3027, Name = "EE_CompanyInfo7", Length = 26, EepromDataType = typeof(byte[]) },

                    // Truck ID
                    new EepromMapping { Address = 3053, Name = "EE_TrukID", Length = 16, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3069, Name = "EE_MultiGprsMode", Length = 1, EepromDataType = typeof(byte) },

                    // Display List
                    new EepromMapping { Address = 3070, Name = "EE_weergaveLijst", Length = 40, EepromDataType = typeof(byte[,]), RowCount = 20, ColumnCount = 2},

                    // GPRS Send Variables
                    new EepromMapping { Address = 3110, Name = "EE_GprsSendVarMem", Length = 50, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3160, Name = "EE_GprsSendVarStartAddr", Length = 50, EepromDataType = typeof(byte[]) }, // Klopt niet met de werkhelijkheid van wat er op de Eeprom staat. Er zijn maar voor 50 bytes gedefineerd. Maar staat 200.
                    new EepromMapping { Address = 3360, Name = "EE_GprsSendVarLengte", Length = 50, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3410, Name = "EE_GprsSendVarAantal", Length = 50, EepromDataType = typeof(byte[]) },

                    // Power Save Relay
                    new EepromMapping { Address = 3460, Name = "EE_PowerSaveRelay", Length = 1, EepromDataType = typeof(byte) },

                    // Watchdog Variables
                    new EepromMapping { Address = 3461, Name = "EE_WD_WdrTeller", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3462, Name = "EE_WD_ProgrammaStatus", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3463, Name = "EE_WD_RitStatus", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3464, Name = "EE_WD_RustSetBy", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3465, Name = "EE_WD_TijdNaarStandby", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3467, Name = "EE_WD_InstallatieTimer", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3468, Name = "EE_WD_RitNummer", Length = 4, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3472, Name = "EE_WD_Alarm", Length = 9, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3481, Name = "EE_WD_Melding", Length = 6, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3487, Name = "EE_WD_SwitchWaardeObu", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3488, Name = "EE_WD_ToevoerVentielatieMomenteel", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3489, Name = "EE_WD_AfvoerVentielatieMomenteel", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3490, Name = "EE_WD_ToevoerVentielatieBereken", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3491, Name = "EE_WD_AfvoerVentielatieBereken", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3492, Name = "EE_WD_VentielatieRegelTijd", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3493, Name = "EE_WD_RegelTijdTimer", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3494, Name = "EE_WD_Verwarming", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3495, Name = "EE_WD_Koeling", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3496, Name = "EE_WD_LogTimer", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3498, Name = "EE_WD_Voeler_Temp", Length = 16, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3530, Name = "EE_WD_Voeler_Offset", Length = 16, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3546, Name = "EE_WD_SwitchWaardeExt", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3548, Name = "EE_WD_SwitchModeExt", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3550, Name = "EE_WD_SwitchVertragingOpExt", Length = 16, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3566, Name = "EE_WD_SwitchVertragingAfExt", Length = 16, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3582, Name = "EE_WD_SwitchConditionExt", Length = 16, EepromDataType = typeof(byte[]) },

                    // Relay and standby power settings
                    new EepromMapping { Address = 3598, Name = "EE_RelayMode", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3599, Name = "EE_StandbyPowerSave", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3600, Name = "EE_WD_GeneratorTimer", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3602, Name = "EE_GeneratorUrenTeller", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3604, Name = "EE_GeneratorOndehoudUrenTeller", Length = 2, EepromDataType = typeof(byte[]) },

                    // Compartments and sensor settings
                    new EepromMapping { Address = 3606, Name = "EE_CompartimentEnable", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3607, Name = "EE_CompTempOffset", Length = 8, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3615, Name = "EE_CompVentMaxOffset", Length = 8, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3623, Name = "EE_CompVentMinOffset", Length = 8, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3631, Name = "EE_CompBandbOffset", Length = 8, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3639, Name = "EE_CompBBCStartOffset", Length = 8, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3647, Name = "EE_CompBBCOffset", Length = 8, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3655, Name = "EE_CompSensorNr", Length = 8, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3663, Name = "EE_CompVentAbsMin", Length = 1, EepromDataType = typeof(byte) },

                    // Water level and alarm settings
                    new EepromMapping { Address = 3664, Name = "EE_WaterNivauIngangNr", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3665, Name = "EE_WaterNivauPcbNr", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3666, Name = "EE_WaterTankMeldingAanUit", Length = 1, EepromDataType = typeof(byte) },

                    // Additional settings
                    new EepromMapping { Address = 3667, Name = "EE_LogInstellingen", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3668, Name = "EE_TrackAndTraceEnableSpeed", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3669, Name = "EE_TrackAndTraceNoSpeedDelay", Length = 1, EepromDataType = typeof(byte) },

                    // Backup Battery and alarm settings
                    new EepromMapping { Address = 3670, Name = "EE_RelayBBackupBat", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3671, Name = "EE_TempBackupBat", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3672, Name = "EE_BackupBatType", Length = 1, EepromDataType = typeof(byte) },

                    // Output and alarm settings
                    new EepromMapping { Address = 3673, Name = "EE_AtrEnable", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3674, Name = "EE_Output50A", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3675, Name = "EE_Output100A", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3676, Name = "EE_Output200A", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3677, Name = "EE_Output300A", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3678, Name = "EE_OutputStart", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3679, Name = "EE_StroomStart", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3680, Name = "EE_AutoRitnummerEnable", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3681, Name = "EE_AlwaysLog", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3682, Name = "EE_GpsAlarmEnable", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3683, Name = "EE_GsmAlarmEnable", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3684, Name = "EE_GpsAlarmDelay", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3686, Name = "EE_GsmAlarmDelay", Length = 2, EepromDataType = typeof(byte[]) }
                };
        }

        // Names have a '0' at the end. This is representing a null or /0 character.
        // Keep in mind that when writing to the EEPROM, this character should be added again.
        //
        // Different data types from the EEPROM are:
        // - unsigned char
        // - unsigned int
        // - signed char
        // - signed int
        // - unsigned long int
        private static void InitializeUnfoldedMapping()
        {
            // DIT KLOPT NOG NIET HELEMAAL, MOET NOG WORDEN NAGEKEKEN EN AANGEPAST!!!
            unfoldedMapping = new List<EepromMapping>()
                {
                    new EepromMapping { Address = 0, Name = "EE_Dummy0", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 1, Name = "EE_ZenderSterkte", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2, Name = "EE_KanaalZender", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3, Name = "EE_ZenderAanUit", Length = 1, EepromDataType = typeof(byte) },

                    new EepromMapping { Address = 4, Name = "EE_Versie", Length = 2, EepromDataType = typeof(ushort) },
                    new EepromMapping { Address = 6, Name = "EE_ResetTeller", Length = 2, EepromDataType = typeof(ushort) },
                    new EepromMapping { Address = 8, Name = "EE_RitStatusMode", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 10, Name = "EE_DummyRitStatus", Length = 10, EepromDataType = typeof(byte[]) },

                    new EepromMapping { Address = 19, Name = "EE_install_contact", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 20, Name = "EE_alarm_contact", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 21, Name = "EE_melding_contact", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 22, Name = "EE_status_contact", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 23, Name = "EE_status_rust_contact", Length = 1, EepromDataType = typeof(byte) }, 
                    new EepromMapping { Address = 24, Name = "EE_Mode_contact", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 25, Name = "EE_AfVertraging_contact", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 27, Name = "EE_OpVertraging_contact", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 29, Name = "EE_Dummy_Contact_Digital", Length = 10, EepromDataType = typeof(byte[]) },

                    new EepromMapping { Address = 39, Name = "EE_AccuSpanningMax", Length = 2, EepromDataType = typeof(ushort) },
                    new EepromMapping { Address = 41, Name = "EE_AccuSpanningMin", Length = 2, EepromDataType = typeof(ushort) },
                    new EepromMapping { Address = 43, Name = "EE_AccuAlarmVertraging", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 44, Name = "EE_BackupBatVoltChargeOn", Length = 2, EepromDataType = typeof(ushort) },
                    new EepromMapping { Address = 46, Name = "EE_BackupBatVoltChargeOff", Length = 2, EepromDataType = typeof(ushort) },
                    new EepromMapping { Address = 48, Name = "EE_BackupBatVoltAlarm", Length = 2, EepromDataType = typeof(ushort) },
                    new EepromMapping { Address = 50, Name = "EE_DummyAccu", Length = 4, EepromDataType = typeof(byte[]) },

                    new EepromMapping { Address = 54, Name = "EE_KlTempSensorNummerExt", Length = 16, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 70, Name = "EE_KlTempSensorLogExt", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 72, Name = "EE_KlTempSensorAlarmExt", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 74, Name = "EE_KlTempSensorRegelingExt", Length = 2, EepromDataType = typeof(byte[]) },
                    
                    new EepromMapping { Address = 76, Name = "EE_KlTempSensorIDExt1", Length = 6, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 82, Name = "EE_KlTempSensorIDExt2", Length = 6, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 88, Name = "EE_KlTempSensorIDExt3", Length = 6, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 94, Name = "EE_KlTempSensorIDExt4", Length = 6, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 100, Name = "EE_KlTempSensorIDExt5", Length = 6, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 106, Name = "EE_KlTempSensorIDExt6", Length = 6, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 112, Name = "EE_KlTempSensorIDExt7", Length = 6, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 118, Name = "EE_KlTempSensorIDExt8", Length = 6, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 124, Name = "EE_KlTempSensorIDExt9", Length = 6, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 130, Name = "EE_KlTempSensorIDExt10", Length = 6, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 136, Name = "EE_KlTempSensorIDExt11", Length = 6, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 142, Name = "EE_KlTempSensorIDExt12", Length = 6, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 148, Name = "EE_KlTempSensorIDExt13", Length = 6, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 154, Name = "EE_KlTempSensorIDExt14", Length = 6, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 160, Name = "EE_KlTempSensorIDExt15", Length = 6, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 166, Name = "EE_KlTempSensorIDExt16", Length = 6, EepromDataType = typeof(byte[]) },

                    new EepromMapping { Address = 172, Name = "EE_LteCatM1Band001to032", Length = 4, EepromDataType = typeof(uint) },
                    new EepromMapping { Address = 176, Name = "EE_LteCatM1Band033to064", Length = 4, EepromDataType = typeof(uint) },
                    new EepromMapping { Address = 180, Name = "EE_LteCatM1Band065to096", Length = 4, EepromDataType = typeof(uint) },
                    new EepromMapping { Address = 184, Name = "EE_LteCatM1Band097to128", Length = 4, EepromDataType = typeof(uint) },

                    new EepromMapping { Address = 188, Name = "EE_LteNbIotBand001to032", Length = 4, EepromDataType = typeof(uint) },
                    new EepromMapping { Address = 192, Name = "EE_LteNbIotBand033to064", Length = 4, EepromDataType = typeof(uint) },
                    new EepromMapping { Address = 196, Name = "EE_LteNbIotBand065to096", Length = 4, EepromDataType = typeof(uint) },
                    new EepromMapping { Address = 200, Name = "EE_LteNbIotBand097to128", Length = 4, EepromDataType = typeof(uint) },

                    new EepromMapping { Address = 204, Name = "EE_LteTechSup", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 205, Name = "EE_2G_Prio", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 206, Name = "EE_TelitTechEnable", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 207, Name = "EE_AccessTechnologyRegisteredNetwork", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 208, Name = "ReserveTemp", Length = 14, EepromDataType = typeof(byte[]) },

                    new EepromMapping { Address = 222, Name = "EE_KlSwitchSensorNummerExt", Length = 16, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 238, Name = "EE_StatusSensorInSwitchExt", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 240, Name = "EE_RustSensorInSwitchExt", Length = 2, EepromDataType = typeof(byte[]) },    
                    new EepromMapping { Address = 242, Name = "EE_AlarmSensorInSwitchExt", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 244, Name = "EE_MeldingSensorInSwitchExt", Length = 2, EepromDataType = typeof(byte[]) },

                    new EepromMapping { Address = 246, Name = "EE_KlSwitchSensorIDExt1", Length = 6, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 252, Name = "EE_KlSwitchSensorIDExt2", Length = 6, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 258, Name = "EE_KlSwitchSensorIDExt3", Length = 6, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 264, Name = "EE_KlSwitchSensorIDExt4", Length = 6, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 270, Name = "EE_KlSwitchSensorIDExt5", Length = 6, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 276, Name = "EE_KlSwitchSensorIDExt6", Length = 6, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 282, Name = "EE_KlSwitchSensorIDExt7", Length = 6, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 288, Name = "EE_KlSwitchSensorIDExt8", Length = 6, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 294, Name = "EE_KlSwitchSensorIDExt9", Length = 6, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 300, Name = "EE_KlSwitchSensorIDExt10", Length = 6, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 306, Name = "EE_KlSwitchSensorIDExt11", Length = 6, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 312, Name = "EE_KlSwitchSensorIDExt12", Length = 6, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 318, Name = "EE_KlSwitchSensorIDExt13", Length = 6, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 324, Name = "EE_KlSwitchSensorIDExt14", Length = 6, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 330, Name = "EE_KlSwitchSensorIDExt15", Length = 6, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 336, Name = "EE_KlSwitchSensorIDExt16", Length = 6, EepromDataType = typeof(byte[]) },

                    new EepromMapping { Address = 342, Name = "EE_gsm_pin", Length = 6, EepromDataType = typeof(string)},
                    new EepromMapping { Address = 348, Name = "EE_gsm_pdp_type", Length = 5, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 353, Name = "EE_gsm_pdp_addr", Length = 21, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 374, Name = "EE_GPRS_UserNaam", Length = 21, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 395, Name = "EE_GPRS_Password", Length = 21, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 416, Name = "EE_ServerPort", Length = 2, EepromDataType = typeof(ushort) },
                    new EepromMapping { Address = 418, Name = "EE_WebAddr", Length = 21, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 439, Name = "EE_GSM_PINCODE_OKE", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 440, Name = "EE_GSM_Enable", Length = 1, EepromDataType = typeof(byte) },

                    new EepromMapping { Address = 441, Name = "EE_GSM_Telnr", Length = 13, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 454, Name = "EE_GSM_SMSnr", Length = 13, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 467, Name = "EE_GSM_SMS_Service", Length = 14, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 481, Name = "EE_GSM_CCID", Length = 26, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 507, Name = "EE_GSM_BAND", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 508, Name = "EE_ServerPortSource", Length = 2, EepromDataType = typeof(ushort) },
                    new EepromMapping { Address = 510, Name = "EE_ResGSM2", Length = 20, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 530, Name = "EE_RTCOffsetGMT", Length = 1, EepromDataType = typeof(sbyte) },

                    new EepromMapping { Address = 531, Name = "EE_RegelTijd", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 532, Name = "EE_VentielatieRegelTijd", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 533, Name = "EE_ToevoerVentielatieEnable", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 534, Name = "EE_ToevoerVentielatieMax", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 535, Name = "EE_ToevoerVentielatieMin", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 536, Name = "EE_ToevoerVentielatieBandbreedte", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 537, Name = "EE_ToevoerVentielatieBBCompensatie", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 538, Name = "EE_ToevoerVentielatieBBCompStart", Length = 2, EepromDataType = typeof(short) },
                    new EepromMapping { Address = 540, Name = "EE_AfvoerVentielatieEnable", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 541, Name = "EE_AfvoerVentielatieMax", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 542, Name = "EE_AfvoerVentielatieMin", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 543, Name = "EE_AfvoerVentielatieBandbreedte", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 544, Name = "EE_AfvoerVentielatieBBCompensatie", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 545, Name = "EE_AfvoerVentielatieBBCompStart", Length = 2, EepromDataType = typeof(short) },
                    new EepromMapping { Address = 547, Name = "EE_MaxVentielatieLoopTijd", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 548, Name = "EE_IngesteldRegelTemperatuur", Length = 2, EepromDataType = typeof(short) },
                    new EepromMapping { Address = 550, Name = "EE_VentielatieMultiInOutPcbNr", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 551, Name = "EE_VentielatieMultiInOutEnable", Length = 1, EepromDataType = typeof(byte) },

                    new EepromMapping { Address = 552, Name = "EE_VerwarmingMode", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 553, Name = "EE_KoelingMode", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 554, Name = "EE_VerwarmingDiff", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 555, Name = "EE_KoelingDiff", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 556, Name = "EE_VerwarmingTemp", Length = 2, EepromDataType = typeof(short) },
                    new EepromMapping { Address = 558, Name = "EE_KoelingTemp", Length = 2, EepromDataType = typeof(short) },
                    new EepromMapping { Address = 560, Name = "EE_VerwarmingTempOffset", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 561, Name = "EE_KoelingTempOffset", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 562, Name = "EE_BuitenVoeler", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 563, Name = "EE_ToevoerVMax", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 564, Name = "EE_ToevoerVMin", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 565, Name = "EE_AfvoerVMax", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 566, Name = "EE_AfvoerVMin", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 567, Name = "EE_MultiInOutAnalogStepTime", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 568, Name = "EE_ResVent", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 570, Name = "EE_GeneratorAanAccuSpanning", Length = 2, EepromDataType = typeof(ushort) },
                    new EepromMapping { Address = 572, Name = "EE_GeneratorAanVertraging", Length = 2, EepromDataType = typeof(ushort) },
                    new EepromMapping { Address = 574, Name = "EE_GeneratorAanTijd", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 575, Name = "EE_GeneratorRegelingEnable", Length = 1, EepromDataType = typeof(byte) },

                    new EepromMapping { Address = 576, Name = "EE_VentUitTemp", Length = 2, EepromDataType = typeof(short) },
                    new EepromMapping { Address = 578, Name = "EE_VentUitSchakelDiff", Length = 2, EepromDataType = typeof(short) },

                    new EepromMapping { Address = 580, Name = "EE_Groot_Dummy", Length = 38, EepromDataType = typeof(byte[]) },

                    new EepromMapping { Address = 618, Name = "EE_GpsFixDelay", Length = 2, EepromDataType = typeof(ushort) },
                    new EepromMapping { Address = 620, Name = "EE_LogBySpeedDelay", Length = 2, EepromDataType = typeof(ushort) },
                    new EepromMapping { Address = 622, Name = "EE_MaxGpsDopRead", Length = 2, EepromDataType = typeof(ushort) },
                    new EepromMapping { Address = 624, Name = "EE_GpsMinDop", Length = 2, EepromDataType = typeof(ushort) },

                    new EepromMapping { Address = 626, Name = "EE_NaarStandby", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 627, Name = "EE_LogTijdNormaal", Length = 2, EepromDataType = typeof(ushort) },
                    new EepromMapping { Address = 629, Name = "EE_LogTijdAlarm", Length = 2, EepromDataType = typeof(ushort) },
                    new EepromMapping { Address = 631, Name = "EE_AlarmVertragingTemp", Length = 1, EepromDataType = typeof(byte) },

                    new EepromMapping { Address = 632, Name = "EE_ThermostaatIngesteldeTemp", Length = 2, EepromDataType = typeof(short) },
                    new EepromMapping { Address = 634, Name = "EE_TempSensorTermostaat", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 636, Name = "EE_RegelingTypeTermostaat", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 637, Name = "EE_SD_Termostaat", Length = 1, EepromDataType = typeof(sbyte) },
                    new EepromMapping { Address = 638, Name = "EE_NZ_Termostaat", Length = 1, EepromDataType = typeof(sbyte) },
                    new EepromMapping { Address = 639, Name = "EE_RelaisTermostaatMode", Length = 1, EepromDataType = typeof(sbyte) },
                    new EepromMapping { Address = 640, Name = "EE_TermostaatRegelingCondition", Length = 1, EepromDataType = typeof(sbyte) },
                    new EepromMapping { Address = 641, Name = "EE_ThermostaatMultiInOutNr", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 642, Name = "EE_ToevoerVentStartPerc", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 643, Name = "EE_ToevoerVentStarttijd", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 644, Name = "EE_AfvoerVentStartPerc", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 645, Name = "EE_AfvoerVentStarttijd", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 646, Name = "EE_Groot_Dummy2", Length = 36, EepromDataType = typeof(byte[]) },

                    new EepromMapping { Address = 682, Name = "EE_AlarmAbsoluutMin", Length = 2, EepromDataType = typeof(short) },
                    new EepromMapping { Address = 684, Name = "EE_AlarmAbsoluutMax", Length = 2, EepromDataType = typeof(short) },
                    new EepromMapping { Address = 686, Name = "EE_Groot_Dummy3", Length = 50, EepromDataType = typeof(byte[]) },

                    new EepromMapping { Address = 736, Name = "EE_MU_AlarmTxtD1", Length = 26, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 762, Name = "EE_MU_AlarmTxtD2", Length = 26, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 788, Name = "EE_MU_AlarmTxtExt1", Length = 26, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 814, Name = "EE_MU_AlarmTxtExt2", Length = 26, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 840, Name = "EE_MU_AlarmTxtExt3", Length = 26, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 866, Name = "EE_MU_AlarmTxtExt4", Length = 26, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 892, Name = "EE_MU_AlarmTxtExt5", Length = 26, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 918, Name = "EE_MU_AlarmTxtExt6", Length = 26, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 944, Name = "EE_MU_AlarmTxtExt7", Length = 26, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 970, Name = "EE_MU_AlarmTxtExt8", Length = 26, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 996, Name = "EE_MU_AlarmTxtExt9", Length = 26, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 1022, Name = "EE_MU_AlarmTxtExt10", Length = 26, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 1048, Name = "EE_MU_AlarmTxtExt11", Length = 26, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 1074, Name = "EE_MU_AlarmTxtExt12", Length = 26, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 1100, Name = "EE_MU_AlarmTxtExt13", Length = 26, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 1126, Name = "EE_MU_AlarmTxtExt14", Length = 26, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 1152, Name = "EE_MU_AlarmTxtExt15", Length = 26, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 1178, Name = "EE_MU_AlarmTxtExt16", Length = 26, EepromDataType = typeof(string) },

                    new EepromMapping { Address = 1204, Name = "EE_Groot_Dummy4", Length = 200, EepromDataType = typeof(byte[]) },

                    new EepromMapping { Address = 1404, Name = "EE_MU_IngangD1Naam", Length = 15, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 1419, Name = "EE_MU_IngangD2Naam", Length = 15, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 1434, Name = "EE_MU_IngangExt1Naam", Length = 15, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 1449, Name = "EE_MU_IngangExt2Naam", Length = 15, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 1464, Name = "EE_MU_IngangExt3Naam", Length = 15, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 1479, Name = "EE_MU_IngangExt4Naam", Length = 15, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 1494, Name = "EE_MU_IngangExt5Naam", Length = 15, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 1509, Name = "EE_MU_IngangExt6Naam", Length = 15, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 1524, Name = "EE_MU_IngangExt7Naam", Length = 15, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 1539, Name = "EE_MU_IngangExt8Naam", Length = 15, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 1554, Name = "EE_MU_IngangExt9Naam", Length = 15, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 1569, Name = "EE_MU_IngangExt10Naam", Length = 15, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 1584, Name = "EE_MU_IngangExt11Naam", Length = 15, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 1599, Name = "EE_MU_IngangExt12Naam", Length = 15, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 1614, Name = "EE_MU_IngangExt13Naam", Length = 15, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 1629, Name = "EE_MU_IngangExt14Naam", Length = 15, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 1644, Name = "EE_MU_IngangExt15Naam", Length = 15, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 1659, Name = "EE_MU_IngangExt16Naam", Length = 15, EepromDataType = typeof(string) },

                    new EepromMapping { Address = 1674, Name = "EE_MU_IngangTempSensorExt1", Length = 15, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 1689, Name = "EE_MU_IngangTempSensorExt2", Length = 15, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 1704, Name = "EE_MU_IngangTempSensorExt3", Length = 15, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 1719, Name = "EE_MU_IngangTempSensorExt4", Length = 15, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 1734, Name = "EE_MU_IngangTempSensorExt5", Length = 15, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 1749, Name = "EE_MU_IngangTempSensorExt6", Length = 15, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 1764, Name = "EE_MU_IngangTempSensorExt7", Length = 15, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 1779, Name = "EE_MU_IngangTempSensorExt8", Length = 15, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 1794, Name = "EE_MU_IngangTempSensorExt9", Length = 15, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 1809, Name = "EE_MU_IngangTempSensorExt10", Length = 15, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 1824, Name = "EE_MU_IngangTempSensorExt11", Length = 15, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 1839, Name = "EE_MU_IngangTempSensorExt12", Length = 15, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 1854, Name = "EE_MU_IngangTempSensorExt13", Length = 15, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 1869, Name = "EE_MU_IngangTempSensorExt14", Length = 15, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 1884, Name = "EE_MU_IngangTempSensorExt15", Length = 15, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 1899, Name = "EE_MU_IngangTempSensorExt16", Length = 15, EepromDataType = typeof(string) },

                    new EepromMapping { Address = 1914, Name = "EE_Groot_Dummy5", Length = 200, EepromDataType = typeof(byte[]) },

                    new EepromMapping { Address = 2114, Name = "EE_MU_ToevoerVentNaam", Length = 15, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 2129, Name = "EE_MU_AfvoerVentNaam", Length = 15, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 2144, Name = "EE_MU_WaterNivauNaam", Length = 15, EepromDataType = typeof(string) },

                    new EepromMapping { Address = 2159, Name = "EE_Groot_Dummy6", Length = 185, EepromDataType = typeof(byte[])},

                    new EepromMapping { Address = 2344, Name = "EE_MU_PictogramD1", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2345, Name = "EE_MU_PictogramD2", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2346, Name = "EE_MU_PictogramSwitchExt1", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2347, Name = "EE_MU_PictogramSwitchExt2", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2348, Name = "EE_MU_PictogramSwitchExt3", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2349, Name = "EE_MU_PictogramSwitchExt4", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2350, Name = "EE_MU_PictogramSwitchExt5", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2351, Name = "EE_MU_PictogramSwitchExt6", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2352, Name = "EE_MU_PictogramSwitchExt7", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2353, Name = "EE_MU_PictogramSwitchExt8", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2354, Name = "EE_MU_PictogramSwitchExt9", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2355, Name = "EE_MU_PictogramSwitchExt10", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2356, Name = "EE_MU_PictogramSwitchExt11", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2357, Name = "EE_MU_PictogramSwitchExt12", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2358, Name = "EE_MU_PictogramSwitchExt13", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2359, Name = "EE_MU_PictogramSwitchExt14", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2360, Name = "EE_MU_PictogramSwitchExt15", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2361, Name = "EE_MU_PictogramSwitchExt16", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2362, Name = "EE_MU_PictogramWaterNivau", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2363, Name = "EE_SmsSelectNew", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 2365, Name = "EE_SmsSendMaxProDayEnable", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2366, Name = "EE_SmsSendMaxProDay", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2367, Name = "EE_Groot_Dummy7", Length = 95, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 2462, Name = "EE_MU_PictorgramTempExt1", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2463, Name = "EE_MU_PictorgramTempExt2", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2464, Name = "EE_MU_PictorgramTempExt3", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2465, Name = "EE_MU_PictorgramTempExt4", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2466, Name = "EE_MU_PictorgramTempExt5", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2467, Name = "EE_MU_PictorgramTempExt6", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2468, Name = "EE_MU_PictorgramTempExt7", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2469, Name = "EE_MU_PictorgramTempExt8", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2470, Name = "EE_MU_PictorgramTempExt9", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2471, Name = "EE_MU_PictorgramTempExt10", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2472, Name = "EE_MU_PictorgramTempExt11", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2473, Name = "EE_MU_PictorgramTempExt12", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2474, Name = "EE_MU_PictorgramTempExt13", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2475, Name = "EE_MU_PictorgramTempExt14", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2476, Name = "EE_MU_PictorgramTempExt15", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2477, Name = "EE_MU_PictorgramTempExt16", Length = 1, EepromDataType = typeof(byte) },
                    
                    new EepromMapping { Address = 2478, Name = "EE_WachdagReset", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2479, Name = "EE_WaterNivauEnable", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2480, Name = "EE_WaterNivauMin", Length = 2, EepromDataType = typeof(ushort) },
                    new EepromMapping { Address = 2482, Name = "EE_WaterNivauMax", Length = 2, EepromDataType = typeof(ushort) },
                    new EepromMapping { Address = 2484, Name = "EE_WaterNivauAlarm", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2485, Name = "EE_waterNivauAlarmVertraging", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2486, Name = "EE_WaterTemperatuurSensor", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 2488, Name = "EE_WaterTemperatuurMax", Length = 2, EepromDataType = typeof(short) },
                    new EepromMapping { Address = 2490, Name = "EE_WaterTemperatuurMin", Length = 2, EepromDataType = typeof(short) },
                    new EepromMapping { Address = 2492, Name = "EE_WaterTemperatuurAlarmVertraging", Length = 1, EepromDataType = typeof(byte) },
                    
                    new EepromMapping { Address = 2493, Name = "EE_SmsSelect", Length = 1, EepromDataType = typeof(byte) },

                    new EepromMapping { Address = 2494, Name = "EE_SmsNrsEnNaam1", Length = 15, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 2509, Name = "EE_SmsNrsEnNaam2", Length = 15, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 2524, Name = "EE_SmsNrsEnNaam3", Length = 15, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 2539, Name = "EE_SmsNrsEnNaam4", Length = 15, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 2554, Name = "EE_SmsNrsEnNaam5", Length = 15, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 2569, Name = "EE_SmsNrsEnNaam6", Length = 15, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 2584, Name = "EE_SmsNrsEnNaam7", Length = 15, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 2599, Name = "EE_SmsNrsEnNaam8", Length = 15, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 2614, Name = "EE_SmsNrsEnNaam9", Length = 15, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 2629, Name = "EE_SmsNrsEnNaam10", Length = 15, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 2644, Name = "EE_SmsNrsEnNaam11", Length = 15, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 2659, Name = "EE_SmsNrsEnNaam12", Length = 15, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 2674, Name = "EE_SmsNrsEnNaam13", Length = 15, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 2689, Name = "EE_SmsNrsEnNaam14", Length = 15, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 2704, Name = "EE_SmsNrsEnNaam15", Length = 15, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 2719, Name = "EE_SmsNrsEnNaam16", Length = 15, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 2734, Name = "EE_SmsNrsEnNaam17", Length = 15, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 2749, Name = "EE_SmsNrsEnNaam18", Length = 15, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 2764, Name = "EE_SmsNrsEnNaam19", Length = 15, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 2779, Name = "EE_SmsNrsEnNaam20", Length = 15, EepromDataType = typeof(string) },

                    new EepromMapping { Address = 2794, Name = "EE_BelSelect", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 2795, Name = "EE_GsmNetworkName", Length = 20, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 2815, Name = "EE_GsmImei", Length = 18, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 2833, Name = "EE_GsmNetworkId", Length = 12, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 2845, Name = "EE_CompanyName", Length = 26, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 2871, Name = "EE_CompanyInfo1", Length = 26, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 2897, Name = "EE_CompanyInfo2", Length = 26, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 2923, Name = "EE_CompanyInfo3", Length = 26, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 2949, Name = "EE_CompanyInfo4", Length = 26, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 2975, Name = "EE_CompanyInfo5", Length = 26, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 3001, Name = "EE_CompanyInfo6", Length = 26, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 3027, Name = "EE_CompanyInfo7", Length = 26, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 3053, Name = "EE_TrukID", Length = 16, EepromDataType = typeof(string) },
                    new EepromMapping { Address = 3069, Name = "EE_MultiGprsMode", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3070, Name = "EE_weergaveLijst1", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3072, Name = "EE_weergaveLijst2", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3074, Name = "EE_weergaveLijst3", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3076, Name = "EE_weergaveLijst4", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3078, Name = "EE_weergaveLijst5", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3080, Name = "EE_weergaveLijst6", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3082, Name = "EE_weergaveLijst7", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3084, Name = "EE_weergaveLijst8", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3086, Name = "EE_weergaveLijst9", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3088, Name = "EE_weergaveLijst10", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3090, Name = "EE_weergaveLijst11", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3092, Name = "EE_weergaveLijst12", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3094, Name = "EE_weergaveLijst13", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3096, Name = "EE_weergaveLijst14", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3098, Name = "EE_weergaveLijst15", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3100, Name = "EE_weergaveLijst16", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3102, Name = "EE_weergaveLijst17", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3104, Name = "EE_weergaveLijst18", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3106, Name = "EE_weergaveLijst19", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3108, Name = "EE_weergaveLijst20", Length = 2, EepromDataType = typeof(byte[]) },
                    
                    new EepromMapping { Address = 3110, Name = "EE_GprsSendVarMem", Length = 50, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3160, Name = "EE_GprsSendVarStartAddr", Length = 50, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3360, Name = "EE_GprsSendVarLengte", Length = 50, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3410, Name = "EE_GprsSendVarAantal", Length = 50, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3460, Name = "EE_PowerSaveRelay", Length = 1, EepromDataType = typeof(byte) },

                    // wd reset var bewaar
                    new EepromMapping { Address = 3461, Name = "EE_WD_WdrTeller", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3462, Name = "EE_WD_ProgrammaStatus", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3463, Name = "EE_WD_RitStatus", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3464, Name = "EE_WD_RustSetBy", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3465, Name = "EE_WD_TijdNaarStandby", Length = 2, EepromDataType = typeof(ushort) },
                    new EepromMapping { Address = 3467, Name = "EE_WD_InstallatieTimer", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3468, Name = "EE_WD_RitNummer", Length = 4, EepromDataType = typeof(uint) },
                    new EepromMapping { Address = 3472, Name = "EE_WD_Alarm", Length = 9, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3481, Name = "EE_WD_Melding", Length = 6, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3487, Name = "EE_WD_SwitchWaardeObu", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3488, Name = "EE_WD_ToevoerVentielatieMomenteel", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3489, Name = "EE_WD_AfvoerVentielatieMomenteel", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3490, Name = "EE_WD_ToevoerVentielatieBereken", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3491, Name = "EE_WD_AfvoerVentielatieBereken", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3492, Name = "EE_WD_VentielatieRegelTijd", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3493, Name = "EE_WD_RegelTijdTimer", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3494, Name = "EE_WD_Verwarming", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3495, Name = "EE_WD_Koeling", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3496, Name = "EE_WD_LogTimer", Length = 2, EepromDataType = typeof(ushort) },
                    new EepromMapping { Address = 3498, Name = "EE_WD_Voeler_Temp", Length = 16, EepromDataType = typeof(byte[]) }, // Gedefineerd met 16 bytes, comments geven 32 bytes aan.
                    new EepromMapping { Address = 3530, Name = "EE_WD_Voeler_Offset", Length = 16, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3546, Name = "EE_WD_SwitchWaardeExt", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3548, Name = "EE_WD_SwitchModeExt", Length = 2, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3550, Name = "EE_WD_SwitchVertragingOpExt", Length = 16, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3566, Name = "EE_WD_SwitchVertragingAfExt", Length = 16, EepromDataType = typeof(byte[]) },
                    new EepromMapping { Address = 3582, Name = "EE_WD_SwitchConditionExt", Length = 16, EepromDataType = typeof(byte[]) },

                    new EepromMapping { Address = 3598, Name = "EE_RelayMode", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3599, Name = "EE_StandbyPowerSave", Length = 1, EepromDataType = typeof(byte) },

                    new EepromMapping { Address = 3600, Name = "EE_WD_GeneratorTimer", Length = 2, EepromDataType = typeof(ushort) },
                    new EepromMapping { Address = 3602, Name = "EE_GeneratorUrenTeller", Length = 2, EepromDataType = typeof(ushort) },
                    new EepromMapping { Address = 3604, Name = "EE_GeneratorOndehoudUrenTeller", Length = 2, EepromDataType = typeof(ushort) },

                    new EepromMapping { Address = 3606, Name = "EE_CompartimentEnable", Length = 1, EepromDataType = typeof(byte) },

                    new EepromMapping { Address = 3607, Name = "EE_CompTempOffset", Length = 8, EepromDataType = typeof(sbyte[]) },
                    new EepromMapping { Address = 3615, Name = "EE_CompVentMaxOffset", Length = 8, EepromDataType = typeof(sbyte[]) },
                    new EepromMapping { Address = 3623, Name = "EE_CompVentMinOffset", Length = 8, EepromDataType = typeof(sbyte[]) },
                    new EepromMapping { Address = 3631, Name = "EE_CompBandbOffset", Length = 8, EepromDataType = typeof(sbyte[]) },
                    new EepromMapping { Address = 3639, Name = "EE_CompBBCStartOffset", Length = 8, EepromDataType = typeof(sbyte[]) },
                    new EepromMapping { Address = 3647, Name = "EE_CompBBCOffset", Length = 8, EepromDataType = typeof(sbyte[]) },
                    new EepromMapping { Address = 3655, Name = "EE_CompSensorNr", Length = 8, EepromDataType = typeof(byte[]) },

                    new EepromMapping { Address = 3663, Name = "EE_CompVentAbsMin", Length = 1, EepromDataType = typeof(byte) },

                    new EepromMapping { Address = 3664, Name = "EE_WaterNivauIngangNr", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3665, Name = "EE_WaterNivauPcbNr", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3666, Name = "EE_WaterTankMeldingAanUit", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3667, Name = "EE_LogInstellingen", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3668, Name = "EE_TrackAndTraceEnableSpeed", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3669, Name = "EE_TrackAndTraceNoSpeedDelay", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3670, Name = "EE_RelayBBackupBat", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3671, Name = "EE_TempBackupBat", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3672, Name = "EE_BackupBatType", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3673, Name = "EE_AtrEnable", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3674, Name = "EE_Output50A", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3675, Name = "EE_Output100A", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3676, Name = "EE_Output200A", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3677, Name = "EE_Output300A", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3678, Name = "EE_OutputStart", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3679, Name = "EE_StroomStart", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3680, Name = "EE_AutoRitnummerEnable", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3681, Name = "EE_AlwaysLog", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3682, Name = "EE_GpsAlarmEnable", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3683, Name = "EE_GsmAlarmEnable", Length = 1, EepromDataType = typeof(byte) },
                    new EepromMapping { Address = 3684, Name = "EE_GpsAlarmDelay", Length = 2, EepromDataType = typeof(ushort) },
                    new EepromMapping { Address = 3686, Name = "EE_GsmAlarmDelay", Length = 2, EepromDataType = typeof(ushort) },

                };
        }

        public static void UpdateUnfoldedByteValue(int address, byte value)
        {
            unfoldedMapping.FirstOrDefault(item => item.Address ==  address).Value = value;
        }

        public static void UpdateUnfoldedSByteValue(int address, sbyte value)
        {
            unfoldedMapping.FirstOrDefault(item => item.Address == address).Value = value;
        }

        public static void UpdateUnfoldedByteArrayValue(int address, byte[] value)
        {
            var temp_mapping = unfoldedMapping.FirstOrDefault(item => item.Address == address);
            if (temp_mapping != null && temp_mapping.EepromDataType == typeof(byte[]))
            {
                temp_mapping.Value = value;
            }
            else
            {
                throw new InvalidOperationException("Address not found or data type mismatch.");
            }
        }

        public static void UpdateUnfoldedSbyteArrayValue(int address, sbyte[] value)
        {
            var temp_mapping = unfoldedMapping.FirstOrDefault(item => item.Address == address);
            if (temp_mapping != null && temp_mapping.EepromDataType == typeof(sbyte[]))
            {
                temp_mapping.Value = value;
            }
            else
            {
                throw new InvalidOperationException("Address not found or data type mismatch.");
            }
        }

        public static void UpdateUnfoldedCharValue(int address, char value)
        {
            unfoldedMapping.FirstOrDefault(item => item.Address == address).Value = value;
        }

        public static void UpdateUnfoldedUIntValue(int address, uint value)
        {
            unfoldedMapping.FirstOrDefault(item => item.Address == address).Value = value;
        }

        public static void UpdateUnfoldedIntValue(int address, int value)
        {
            unfoldedMapping.FirstOrDefault(item => item.Address == address).Value = value;
        }

        public static void UpdateUnfoldedStringValue(int address, string value)
        {
            unfoldedMapping.FirstOrDefault(item => item.Address == address).Value = value;
        }

        public static EepromMapping GetMappingByAddress(int address)
        {
            return mapping.FirstOrDefault(item => item.Address == address);
        }

        public static EepromMapping GetUnfoldedMappingByAddress(int address)
        {
            return unfoldedMapping.FirstOrDefault(item => item.Address == address);
        }

        // Check if the address exists in the mapping
        // Returns true if the address exists in the mapping
        public static bool doesUnmappingExist(int address)
        {
            return unfoldedMapping.Any(item => item.Address == address);
        }

        // Print the mapping from address
        public static void PrintMapping(int address)
        {
            var mappingTemp = GetMappingByAddress(address);

            Console.WriteLine($"Address: {mappingTemp.Address}");
            Console.WriteLine($"Name: {mappingTemp.Name}");
            Console.WriteLine($"Length: {mappingTemp.Length}");
            Console.WriteLine($"EepromDataType: {mappingTemp.EepromDataType}");
        }

        // Print the unfolded mapping from address
        public static void PrintUnfoldedMapping(int address)
        {
            var uMapping = unfoldedMapping.FirstOrDefault(item => item.Address == address);

            Console.WriteLine($"Address: {uMapping.Address}");
            Console.WriteLine($"Name: {uMapping.Name}");
            Console.WriteLine($"Length: {uMapping.Length}");
            Console.WriteLine($"Value: {uMapping.Value}");
            Console.WriteLine($"EepromDataType: {uMapping.EepromDataType}");
        }
    }
}
