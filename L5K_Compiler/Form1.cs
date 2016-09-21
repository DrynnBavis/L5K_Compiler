using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;

namespace L5K_Compiler
{
    public partial class Form1 : Form
    {
        string outputPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\";
        string[] m1756L71S = @"	MODULE Local (Parent := ""Local"",
	              ParentModPortId := 1,
	              CatalogNumber := ""1756-L71S"",
	              Vendor := 1,
	              ProductType := 14,
	              ProductCode := 158,
	              Major := 27,
	              Minor := 11,
	              PortLabel := ""RxBACKPLANE"",
	              ChassisSize := 13,
	              Slot := 11,
	              Mode := 2#0000_0000_0000_0001,
	              CompatibleModule := 0,
	              KeyMask := 2#0000_0000_0001_1111,
	              SafetyNetwork := 16#0000_3acc_033e_6fa0)
	END_MODULE".Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
        string [] m1756EN2T = @"	MODULE Drives (Description := ""DPxxxx & DPxxxx & DPxxxx"",
	               Parent := ""Local"",
	               ParentModPortId := 1,
	               CatalogNumber := ""1756-EN2T"",
	               Vendor := 1,
	               ProductType := 12,
	               ProductCode := 166,
	               Major := 10,
	               Minor := 1,
	               PortLabel := ""RxBACKPLANE"",
	               Slot := 5,
	               NodeAddress := ""192.168.0.1"",
	               Mode := 2#0000_0000_0000_0000,
	               CompatibleModule := 1,
	               KeyMask := 2#0000_0000_0001_1111)
			ExtendedProp := [[[___<public><ConfigID>131178</ConfigID></public>___]]]
			ConfigData := [20,0,393217,33619969,256,0];
			CONNECTION Input2(Rate := 500000,
                               EventID := 0)

            END_CONNECTION

    END_MODULE".Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
        string[] mPowerFlex753ENETR = @"	MODULE _xxxx_01M02 (Description := ""VFD Speed Control"",
	                    Parent := ""Drives"",
	                    ParentModPortId := 2,
	                    CatalogNumber := ""PowerFlex 753-ENETR"",
	                    Vendor := 1,
	                    ProductType := 142,
	                    ProductCode := 1168,
	                    Major := 11,
	                    Minor := 2,
	                    PortLabel := ""ENet"",
	                    Slot := 0,
	                    NodeAddress := ""192.168.0.2"",
	                    CommMethod := 536870913,
	                    Mode := 2#0000_0000_0000_0000,
	                    CompatibleModule := 0,
	                    KeyMask := 2#0000_0000_0000_0000,
	                    DrivesADCMode := 1,
	                    DrivesADCEnabled := 0)
			ExtendedProp := [[[___<public><AOPVersion>40040100</AOPVersion><DriveConfigCode>4</DriveConfigCode><DriveRatingOptions>0</DriveRatingOptions><PromptImport>0</PromptImport><CommModulePort>5</CommModulePort><IO_XML_INPUT>&lt;DataTypes&gt;&lt;DataType Name=$QAB:PowerFlex753_R_7E7342AA:I:0$Q Class=$QIO$Q&gt;&lt;Members&gt;&lt;Member Name=$Qpad$Q DataType=$QDINT$Q Hidden=$Q1$Q/&gt;&lt;Member Name=$QDriveStatus$Q DataType=$QDINT$Q Radix=$QBinary$Q/&gt;&lt;Member Name=$QDriveStatus_Ready$Q DataType=$QBIT$Q BitNumber=$Q0$Q Target=$QDriveStatus$Q&gt;&lt;/Member&gt;&lt;Member Name=$QDriveStatus_Active$Q DataType=$QBIT$Q BitNumber=$Q1$Q Target=$QDriveStatus$Q&gt;&lt;/Member&gt;&lt;Member Name=$QDriveStatus_CommandDir$Q DataType=$QBIT$Q BitNumber=$Q2$Q Target=$QDriveStatus$Q&gt;&lt;/Member&gt;&lt;Member Name=$QDriveStatus_ActualDir$Q DataType=$QBIT$Q BitNumber=$Q3$Q Target=$QDriveStatus$Q&gt;&lt;/Member&gt;&lt;Member Name=$QDriveStatus_Accelerating$Q DataType=$QBIT$Q BitNumber=$Q4$Q Target=$QDriveStatus$Q&gt;&lt;/Member&gt;&lt;Member Name=$QDriveStatus_Decelerating$Q DataType=$QBIT$Q BitNumber=$Q5$Q Target=$QDriveStatus$Q&gt;&lt;/Member&gt;&lt;Member Name=$QDriveStatus_Alarm$Q DataType=$QBIT$Q BitNumber=$Q6$Q Target=$QDriveStatus$Q&gt;&lt;/Member&gt;&lt;Member Name=$QDriveStatus_Faulted$Q DataType=$QBIT$Q BitNumber=$Q7$Q Target=$QDriveStatus$Q&gt;&lt;/Member&gt;&lt;Member Name=$QDriveStatus_AtSpeed$Q DataType=$QBIT$Q BitNumber=$Q8$Q Target=$QDriveStatus$Q&gt;&lt;/Member&gt;&lt;Member Name=$QDriveStatus_Manual$Q DataType=$QBIT$Q BitNumber=$Q9$Q Target=$QDriveStatus$Q&gt;&lt;/Member&gt;&lt;Member Name=$QDriveStatus_SpdRefBit0$Q DataType=$QBIT$Q BitNumber=$Q10$Q Target=$QDriveStatus$Q&gt;&lt;/Member&gt;&lt;Member Name=$QDriveStatus_SpdRefBit1$Q DataType=$QBIT$Q BitNumber=$Q11$Q Target=$QDriveStatus$Q&gt;&lt;/Member&gt;&lt;Member Name=$QDriveStatus_SpdRefBit2$Q DataType=$QBIT$Q BitNumber=$Q12$Q Target=$QDriveStatus$Q&gt;&lt;/Member&gt;&lt;Member Name=$QDriveStatus_SpdRefBit3$Q DataType=$QBIT$Q BitNumber=$Q13$Q Target=$QDriveStatus$Q&gt;&lt;/Member&gt;&lt;Member Name=$QDriveStatus_SpdRefBit4$Q DataType=$QBIT$Q BitNumber=$Q14$Q Target=$QDriveStatus$Q&gt;&lt;/Member&gt;&lt;Member Name=$QDriveStatus_Running$Q DataType=$QBIT$Q BitNumber=$Q16$Q Target=$QDriveStatus$Q&gt;&lt;/Member&gt;&lt;Member Name=$QDriveStatus_Jogging$Q DataType=$QBIT$Q BitNumber=$Q17$Q Target=$QDriveStatus$Q&gt;&lt;/Member&gt;&lt;Member Name=$QDriveStatus_Stopping$Q DataType=$QBIT$Q BitNumber=$Q18$Q Target=$QDriveStatus$Q&gt;&lt;/Member&gt;&lt;Member Name=$QDriveStatus_DCBraking$Q DataType=$QBIT$Q BitNumber=$Q19$Q Target=$QDriveStatus$Q&gt;&lt;/Member&gt;&lt;Member Name=$QDriveStatus_DBActive$Q DataType=$QBIT$Q BitNumber=$Q20$Q Target=$QDriveStatus$Q&gt;&lt;/Member&gt;&lt;Member Name=$QDriveStatus_SpeedMode$Q DataType=$QBIT$Q BitNumber=$Q21$Q Target=$QDriveStatus$Q&gt;&lt;/Member&gt;&lt;Member Name=$QDriveStatus_PositionMode$Q DataType=$QBIT$Q BitNumber=$Q22$Q Target=$QDriveStatus$Q&gt;&lt;/Member&gt;&lt;Member Name=$QDriveStatus_TorqueMode$Q DataType=$QBIT$Q BitNumber=$Q23$Q Target=$QDriveStatus$Q&gt;&lt;/Member&gt;&lt;Member Name=$QDriveStatus_AtZeroSpeed$Q DataType=$QBIT$Q BitNumber=$Q24$Q Target=$QDriveStatus$Q&gt;&lt;/Member&gt;&lt;Member Name=$QDriveStatus_AtHome$Q DataType=$QBIT$Q BitNumber=$Q25$Q Target=$QDriveStatus$Q&gt;&lt;/Member&gt;&lt;Member Name=$QDriveStatus_AtLimit$Q DataType=$QBIT$Q BitNumber=$Q26$Q Target=$QDriveStatus$Q&gt;&lt;/Member&gt;&lt;Member Name=$QDriveStatus_CurrLimit$Q DataType=$QBIT$Q BitNumber=$Q27$Q Target=$QDriveStatus$Q&gt;&lt;/Member&gt;&lt;Member Name=$QDriveStatus_BusFrqReg$Q DataType=$QBIT$Q BitNumber=$Q28$Q Target=$QDriveStatus$Q&gt;&lt;/Member&gt;&lt;Member Name=$QDriveStatus_EnableOn$Q DataType=$QBIT$Q BitNumber=$Q29$Q Target=$QDriveStatus$Q&gt;&lt;/Member&gt;&lt;Member Name=$QDriveStatus_MotorOL$Q DataType=$QBIT$Q BitNumber=$Q30$Q Target=$QDriveStatus$Q&gt;&lt;/Member&gt;&lt;Member Name=$QDriveStatus_Regen$Q DataType=$QBIT$Q BitNumber=$Q31$Q Target=$QDriveStatus$Q&gt;&lt;/Member&gt;&lt;Member Name=$QFeedback$Q DataType=$QREAL$Q Radix=$QDecimal$Q/&gt;&lt;Member Name=$QOutputCurrent$Q DataType=$QREAL$Q Radix=$QDecimal$Q/&gt;&lt;Member Name=$QLastFaultCode$Q DataType=$QDINT$Q Radix=$QDecimal$Q/&gt;&lt;Member Name=$QMotorNPAmps$Q DataType=$QREAL$Q Radix=$QDecimal$Q/&gt;&lt;Member Name=$QMaxFwdSpeed$Q DataType=$QREAL$Q Radix=$QDecimal$Q/&gt;&lt;Member Name=$QDCBusVolts$Q DataType=$QREAL$Q Radix=$QDecimal$Q/&gt;&lt;Member Name=$QMtrOLCounts$Q DataType=$QREAL$Q Radix=$QDecimal$Q/&gt;&lt;Member Name=$QPTCSts$Q DataType=$QDINT$Q Radix=$QBinary$Q/&gt;&lt;Member Name=$QPTCSts_PTCOk$Q DataType=$QBIT$Q BitNumber=$Q0$Q Target=$QPTCSts$Q&gt;&lt;/Member&gt;&lt;Member Name=$QPTCSts_Reserved$Q DataType=$QBIT$Q BitNumber=$Q1$Q Target=$QPTCSts$Q&gt;&lt;/Member&gt;&lt;Member Name=$QPTCSts_OverTemp$Q DataType=$QBIT$Q BitNumber=$Q2$Q Target=$QPTCSts$Q&gt;&lt;/Member&gt;&lt;Member Name=$QStartInhibits$Q DataType=$QDINT$Q Radix=$QBinary$Q/&gt;&lt;Member Name=$QStartInhibits_Faulted$Q DataType=$QBIT$Q BitNumber=$Q0$Q Target=$QStartInhibits$Q&gt;&lt;/Member&gt;&lt;Member Name=$QStartInhibits_Alarm$Q DataType=$QBIT$Q BitNumber=$Q1$Q Target=$QStartInhibits$Q&gt;&lt;/Member&gt;&lt;Member Name=$QStartInhibits_Enable$Q DataType=$QBIT$Q BitNumber=$Q2$Q Target=$QStartInhibits$Q&gt;&lt;/Member&gt;&lt;Member Name=$QStartInhibits_Precharge$Q DataType=$QBIT$Q BitNumber=$Q3$Q Target=$QStartInhibits$Q&gt;&lt;/Member&gt;&lt;Member Name=$QStartInhibits_Stop$Q DataType=$QBIT$Q BitNumber=$Q4$Q Target=$QStartInhibits$Q&gt;&lt;/Member&gt;&lt;Member Name=$QStartInhibits_Database$Q DataType=$QBIT$Q BitNumber=$Q5$Q Target=$QStartInhibits$Q&gt;&lt;/Member&gt;&lt;Member Name=$QStartInhibits_Startup$Q DataType=$QBIT$Q BitNumber=$Q6$Q Target=$QStartInhibits$Q&gt;&lt;/Member&gt;&lt;Member Name=$QStartInhibits_Safety$Q DataType=$QBIT$Q BitNumber=$Q7$Q Target=$QStartInhibits$Q&gt;&lt;/Member&gt;&lt;Member Name=$QStartInhibits_Sleep$Q DataType=$QBIT$Q BitNumber=$Q8$Q Target=$QStartInhibits$Q&gt;&lt;/Member&gt;&lt;Member Name=$QStartInhibits_Profiler$Q DataType=$QBIT$Q BitNumber=$Q9$Q Target=$QStartInhibits$Q&gt;&lt;/Member&gt;&lt;/Members&gt;&lt;/DataType&gt;&lt;/DataTypes&gt;</IO_XML_INPUT><IO_XML_OUTPUT>&lt;DataTypes&gt;&lt;DataType Name=$QAB:PowerFlex753_R_B34DFDA2:O:0$Q Class=$QIO$Q&gt;&lt;Members&gt;&lt;Member Name=$QLogicCommand$Q DataType=$QDINT$Q Radix=$QBinary$Q/&gt;&lt;Member Name=$QLogicCommand_Stop$Q DataType=$QBIT$Q BitNumber=$Q0$Q Target=$QLogicCommand$Q&gt;&lt;/Member&gt;&lt;Member Name=$QLogicCommand_Start$Q DataType=$QBIT$Q BitNumber=$Q1$Q Target=$QLogicCommand$Q&gt;&lt;/Member&gt;&lt;Member Name=$QLogicCommand_Jog1$Q DataType=$QBIT$Q BitNumber=$Q2$Q Target=$QLogicCommand$Q&gt;&lt;/Member&gt;&lt;Member Name=$QLogicCommand_ClearFaults$Q DataType=$QBIT$Q BitNumber=$Q3$Q Target=$QLogicCommand$Q&gt;&lt;/Member&gt;&lt;Member Name=$QLogicCommand_Forward$Q DataType=$QBIT$Q BitNumber=$Q4$Q Target=$QLogicCommand$Q&gt;&lt;/Member&gt;&lt;Member Name=$QLogicCommand_Reverse$Q DataType=$QBIT$Q BitNumber=$Q5$Q Target=$QLogicCommand$Q&gt;&lt;/Member&gt;&lt;Member Name=$QLogicCommand_Manual$Q DataType=$QBIT$Q BitNumber=$Q6$Q Target=$QLogicCommand$Q&gt;&lt;/Member&gt;&lt;Member Name=$QLogicCommand_AccelTime1$Q DataType=$QBIT$Q BitNumber=$Q8$Q Target=$QLogicCommand$Q&gt;&lt;/Member&gt;&lt;Member Name=$QLogicCommand_AccelTime2$Q DataType=$QBIT$Q BitNumber=$Q9$Q Target=$QLogicCommand$Q&gt;&lt;/Member&gt;&lt;Member Name=$QLogicCommand_DecelTime1$Q DataType=$QBIT$Q BitNumber=$Q10$Q Target=$QLogicCommand$Q&gt;&lt;/Member&gt;&lt;Member Name=$QLogicCommand_DecelTime2$Q DataType=$QBIT$Q BitNumber=$Q11$Q Target=$QLogicCommand$Q&gt;&lt;/Member&gt;&lt;Member Name=$QLogicCommand_SpdRefSel0$Q DataType=$QBIT$Q BitNumber=$Q12$Q Target=$QLogicCommand$Q&gt;&lt;/Member&gt;&lt;Member Name=$QLogicCommand_SpdRefSel1$Q DataType=$QBIT$Q BitNumber=$Q13$Q Target=$QLogicCommand$Q&gt;&lt;/Member&gt;&lt;Member Name=$QLogicCommand_SpdRefSel2$Q DataType=$QBIT$Q BitNumber=$Q14$Q Target=$QLogicCommand$Q&gt;&lt;/Member&gt;&lt;Member Name=$QLogicCommand_CoastStop$Q DataType=$QBIT$Q BitNumber=$Q16$Q Target=$QLogicCommand$Q&gt;&lt;/Member&gt;&lt;Member Name=$QLogicCommand_CLimitStop$Q DataType=$QBIT$Q BitNumber=$Q17$Q Target=$QLogicCommand$Q&gt;&lt;/Member&gt;&lt;Member Name=$QLogicCommand_Run$Q DataType=$QBIT$Q BitNumber=$Q18$Q Target=$QLogicCommand$Q&gt;&lt;/Member&gt;&lt;Member Name=$QLogicCommand_Jog2$Q DataType=$QBIT$Q BitNumber=$Q19$Q Target=$QLogicCommand$Q&gt;&lt;/Member&gt;&lt;Member Name=$QReference$Q DataType=$QREAL$Q Radix=$QDecimal$Q/&gt;&lt;Member Name=$QAccelTime1$Q DataType=$QREAL$Q Radix=$QDecimal$Q/&gt;&lt;Member Name=$QDecelTime1$Q DataType=$QREAL$Q Radix=$QDecimal$Q/&gt;&lt;/Members&gt;&lt;/DataType&gt;&lt;/DataTypes&gt;</IO_XML_OUTPUT><LgxVersion>27</LgxVersion><CommModuleMajorRev>1</CommModuleMajorRev><CommModuleMinorRev>1</CommModuleMinorRev><Port0CCV>0</Port0CCV><Port0CCVInfo>1,0,0,0,0,0,0,0,0,</Port0CCVInfo><Port0HLP_Size>0</Port0HLP_Size><Port0Type>PowerFlex 753</Port0Type><DriveRatingCode>1107296306</DriveRatingCode><ConfigID>115</ConfigID><UsingNAT>0</UsingNAT><Port0DeviceDefinition>0002000000900000000400010000000000010001020B    0002020B010100020001000000004200003200000000</Port0DeviceDefinition></public>___]]]
			ConfigData := [360,0,6,0,1,0,50,16896,0,0,0,0,0,0,0,0,535,0,537,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,7,0,951,0,26,0,520,0,11,0,418,0,251
		,0,933,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0
		,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0];
			ConfigScript (Size := 2488) := [-76,9,0,0,4,0,0,0,0,0,0,0,0,0,0,0,25,0,0,0,8,-106,0,0,0,1,0,0,0,1,0,0,0,8,0,0,0,75,2,32,-110,36,0,-1,-1,0,0,0,1,9,0,0,8,10,0,0,0,1,0,0,0,1,0,0,0,74
		,0,0,0,16,3,32,-98,36,1,48,5,2,0,3,0,0,2,1,3,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0
		,0,0,1,0,0,0,74,0,0,0,16,3,32,-98,36,1,48,5,3,0,3,44,1,9,4,0,0,0,0,4,0,0,0,0,4,0,0,0,0,0,3,49,1,9,4,0,0,0,0,4,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0
		,0,0,0,0,0,0,0,0,0,0,0,0,1,0,0,0,74,0,0,0,16,3,32,-98,36,1,48,5,4,0,47,0,0,5,1,0,0,3,0,0,1,2,0,0,0,1,0,0,2,1,0,0,2,1,0,1,1,2,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0
		,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,0,0,0,74,0,0,0,16,3,32,-98,36,1,48,5,5,0,3,0,0,2,1,4,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0
		,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,0,0,0,74,0,0,0,16,3,32,-98,36,1,48,5,6,0,1,0,0,31,36,2,0,16,0,80,0,111,0,119,0,101,0
		,114,0,70,0,108,0,101,0,120,0,32,0,55,0,53,0,51,0,32,0,32,0,32,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,0,0,0,74,0,0,0,16,3,32,-98,36
		,1,48,5,7,0,6,1,0,6,24,7,0,100,0,2,0,8,0,65,0,109,0,112,0,115,0,32,0,32,0,32,0,32,0,0,6,2,0,6,24,11,0,100,0,2,0,8,0,66,0,117,0,115,0,32,0
		,86,0,68,0,67,0,32,0,0,0,0,0,0,1,0,0,0,74,0,0,0,16,3,32,-98,36,1,48,5,8,0,6,3,0,6,24,8,0,100,0,2,0,8,0,79,0,117,0,116,0,32,0,86,0,108,0,116
		,0,115,0,0,6,4,0,6,24,9,0,100,0,2,0,8,0,79,0,117,0,116,0,32,0,80,0,119,0,114,0,32,0,0,0,0,0,0,1,0,0,0,74,0,0,0,16,3,32,-98,36,1,48,5,9,0,6
		,5,0,6,24,14,0,100,0,2,0,8,0,69,0,108,0,112,0,32,0,107,0,87,0,72,0,114,0,0,6,6,0,6,24,5,0,100,0,2,0,8,0,84,0,114,0,113,0,32,0,67,0,117
		,0,114,0,32,0,0,0,0,0,0,1,0,0,0,74,0,0,0,16,3,32,-98,36,1,48,5,10,0,3,70,0,9,4,1,0,0,0,0,3,25,0,9,4,0,0,-56,67,4,51,51,-109,64,4,0,0,72,66
		,4,0,-32,-75,68,4,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,0,0,0,74,0,0,0,16,3,32,-98,36,1,48,5,11,0,3,30,0,9,4,0,0,16,64,0,3,71
		,0,9,4,0,0,72,66,0,3,73,0,9,4,92,-113,66,65,4,0,0,-56,66,4,0,0,-32,63,4,0,0,0,64,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,0,0,0,74,0,0,0,16,3,32
		,-98,36,1,48,5,12,0,3,77,0,9,4,0,0,0,0,4,0,0,0,0,4,0,0,96,64,0,3,110,0,9,4,0,0,0,0,4,0,0,0,0,0,3,109,2,9,4,0,0,52,66,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0
		,0,0,0,0,1,0,0,0,74,0,0,0,16,3,32,-98,36,1,48,5,13,0,3,36,0,9,4,0,0,-56,67,4,0,0,2,67,4,0,0,-128,64,0,3,43,0,9,4,1,0,0,0,4,0,0,0,0,0,3,60,0,9,4
		,92,-113,66,65,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,0,0,0,74,0,0,0,16,3,32,-98,36,1,48,5,14,0,3,61,0,9,4,92,-113,66,65,4,0,0,-56,66,4
		,0,0,72,65,0,3,8,2,9,4,0,0,112,66,4,0,0,112,-62,0,3,5,1,9,4,0,0,32,65,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,0,0,0,74,0,0,0,16,3,32,-98,36,1,48
		,5,15,0,3,6,1,9,4,0,0,0,0,0,3,14,1,9,2,0,0,0,3,24,1,9,4,0,0,32,65,4,0,0,0,0,0,3,73,1,9,4,0,0,112,66,0,3,119,1,9,4,-35,-124,59,68,0,0,0,0,0,0,0,0
		,0,0,0,0,1,0,0,0,74,0,0,0,16,3,32,-98,36,1,48,5,16,0,3,127,1,9,4,0,0,-8,65,0,3,-118,1,9,4,0,0,-96,64,0,3,-90,1,9,4,0,0,16,65,4,0,0,16,65,0
		,3,-76,1,9,4,0,0,-96,64,0,3,-73,1,9,4,0,0,-96,64,0,0,0,0,0,0,0,0,0,0,1,0,0,0,74,0,0,0,16,3,32,-98,36,1,48,5,17,0,3,-61,1,9,4,0,0,52,67,0,3,-58
		,1,9,4,0,0,52,67,0,3,-51,1,9,4,0,0,122,67,0,3,-45,1,9,4,0,0,-128,64,0,3,12,2,9,4,0,0,32,65,4,-113,-62,117,61,0,0,0,0,0,0,0,0,0,0,1,0,0,0,74
		,0,0,0,16,3,32,-98,36,1,48,5,18,0,3,35,2,9,4,0,0,112,66,0,3,40,2,9,4,0,0,112,66,0,3,44,2,9,4,0,0,32,65,4,0,0,32,65,0,3,52,2,9,4,0,0,112,66
		,0,3,59,2,9,4,0,0,-96,64,0,0,0,0,0,0,0,0,0,0,1,0,0,0,74,0,0,0,16,3,32,-98,36,1,48,5,19,0,3,60,2,9,4,0,0,32,65,4,0,0,-96,65,4,0,0,-16,65,4,0,0
		,32,66,4,0,0,72,66,4,0,0,72,66,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,0,0,0,74,0,0,0,16,3,32,-98,36,1,48,5,20,0,3,90,2,9,4,0
		,0,112,66,0,3,94,2,9,4,0,0,112,66,0,3,125,0,9,4,-119,0,0,0,4,3,0,0,0,0,3,124,2,9,4,0,0,32,65,0,3,-128,0,9,4,-119,0,0,0,0,0,0,0,0,0,0,0,0,0,1
		,0,0,0,74,0,0,0,16,3,32,-98,36,1,48,5,21,0,3,-127,0,9,4,3,0,0,0,0,3,-120,2,9,4,0,0,32,65,0,3,-33,2,9,4,0,0,-96,64,0,3,17,3,9,4,0,0,-56,65
		,4,0,0,-56,-63,0,3,56,3,9,4,0,0,-128,62,0,0,0,0,0,0,0,0,0,0,1,0,0,0,74,0,0,0,16,3,32,-98,36,1,48,5,22,0,3,3,6,9,4,0,0,0,64,4,0,0,32,65,0,3,7,6
		,9,4,-51,-52,76,63,4,0,0,0,63,4,0,0,32,64,0,3,93,6,9,4,0,0,112,66,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,0,0,0,74,0,0,0,16,3,32,-98,36,1,48
		,5,23,0,3,105,6,9,4,0,0,112,66,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,0,0,0,74,0,0
		,0,16,3,32,-98,36,1,48,5,24,0,3,60,0,9,4,92,-113,66,65,4,92,-113,66,65,0,3,73,0,9,4,92,-113,66,65,0,3,73,1,9,4,0,0,112,66,0,3,33,2
		,9,4,107,3,0,0,0,3,35,2,9,4,0,0,112,66,0,0,0,0,0,0,0,0,0,0,1,0,0,0,74,0,0,0,16,3,32,-98,36,1,48,5,25,0,3,40,2,9,4,0,0,112,66,0,3,52,2,9,4,0,0,112
		,66,0,3,65,2,9,4,0,0,72,66,0,3,90,2,9,4,0,0,112,66,0,3,94,2,9,4,0,0,112,66,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,0,0,0,74,0,0,0,16,3,32,-98,36,1,48,5
		,26,0,1,0,0,31,36,2,0,16,0,86,0,70,0,68,0,95,0,83,0,112,0,101,0,101,0,100,0,95,0,67,0,111,0,110,0,116,0,114,0,111,0,0,0,0,0,0,0,0,0,0
		,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,0,0,0,74,0,0,0,16,3,32,-98,36,1,48,5,27,0,6,1,0,6,24,7,0,100,0,2,0,8,0,65,0,109,0,112,0,115,0,32,0,32,0,32,0,32
		,0,0,6,2,0,6,24,11,0,100,0,2,0,8,0,66,0,117,0,115,0,32,0,86,0,68,0,67,0,32,0,0,0,0,0,0,1,0,0,0,74,0,0,0,16,3,32,-98,36,1,48,5,28,0,6,3,0,6,24
		,8,0,100,0,2,0,8,0,79,0,117,0,116,0,32,0,86,0,108,0,116,0,115,0,0,6,4,0,6,24,9,0,100,0,2,0,8,0,79,0,117,0,116,0,32,0,80,0,119,0,114,0
		,32,0,0,0,0,0,0,1,0,0,0,74,0,0,0,16,3,32,-98,36,1,48,5,29,0,6,5,0,6,24,14,0,100,0,2,0,8,0,69,0,108,0,112,0,32,0,107,0,87,0,72,0,114,0,0,6
		,6,0,6,24,5,0,100,0,2,0,8,0,84,0,114,0,113,0,32,0,67,0,117,0,114,0,32,0,0,0,0,0,0,0,0,0,45,0,0,0,8,101,0,0,0,1,0,0,0,6,0,0,0,24,0,0,0,16,3,32
		,-110,36,0,48,38,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,64,0,0,0,0,0,0,36,0,0,0,8,61,0,0,0,1,0,0,0,1,0,0,0,11,0,0,0,16,4,32,-105,37,0,0,0,48,3,3,3,0,0,0
		,-56,-81,0,0,29,0,0,0,0,100,0,0,0,2,0,0,0,-128,0,0,0,-81,96,99,-38,51,-23,17,-26,-121,-33,0,12,41,-25,-114,-55,0,0,0];
			CONNECTION Output (Rate := 20000,
			                   InputCxnPoint := 1,
			                   InputSize := 44,
			                   OutputCxnPoint := 2,
			                   OutputSize := 16,
			                   EventID := 0,
			                   Unicast := Yes)
					InputData (Class := Standard) := [0,0,0.00000000e+000,0.00000000e+000,0,0.00000000e+000,0.00000000e+000,0.00000000e+000,0.00000000e+000
		,0,0];
					OutputData (Class := Standard) := [0,0.00000000e+000,0.00000000e+000,0.00000000e+000];
			END_CONNECTION

	END_MODULE".Split(new string[] { Environment.NewLine }, StringSplitOptions.None);

        public Form1()
        {
            InitializeComponent();
        }

        private void ChoosePath()
        {
            FolderBrowserDialog folderBrowser = new FolderBrowserDialog();
            folderBrowser.Reset();
            folderBrowser.Description = "Please select a folder containing test data for analysis.";
            folderBrowser.ShowNewFolderButton = false;
            folderBrowser.RootFolder = Environment.SpecialFolder.MyComputer;
            folderBrowser.SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\";
            if (folderBrowser.ShowDialog() == System.Windows.Forms.DialogResult.Cancel)
            {
                return;
            }
            outputPath = folderBrowser.SelectedPath + "\\";
        }

        private void CompileL5K()
        {
            string fileName = Microsoft.VisualBasic.Interaction.InputBox("Please enter a name for the L5K file:", 
                "Enter File Name", "NewFile");
            File.WriteAllLines(outputPath + fileName + ".l5k", m1756EN2T);
        }

        private void changePathBtn_Click(object sender, EventArgs e)
        {
            ChoosePath();
        }

        private void compileBtn_Click(object sender, EventArgs e)
        {
            CompileL5K();
            //adding a dif comment
            //Now this is my first change!
        }
    }
}
