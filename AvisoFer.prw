#Include 'Protheus.ch'
#Include 'FWMVCDEF.ch'
#Include 'topconn.ch'
#Include "TbiConn.ch"

// +--------------------------+---------------------------------+-------------+
// | Programa   : AvisoFer.prw | Autor : Luciana Rosa		    | 01.01.2020. |
// +--------------------------+---------------------------------+-------------+
// | Descrição  : Aviso de Férias.                          				  |
// +--------------------------------------------------------------------------+


User Function AvFer()

Local cQrySQB	:= ''
Local cQrySQB2	:= ''
Local nCount	:= 0
Local _cHtml    := ''
Local lZebrado 	:= .F.
Local dIniPA	:= '' //Inicio do Período Aquisitivo
Local dFimPA	:= '' //Fim do Período Aquisitivo
Local dIniFe	:= Nil //Inicio de Período de Gozo
Local nDiaFer 	:= 0 // Dias de Férias
Local dFimFe	:= Nil //Fim de Período de Gozo
Local nDiaDir 	:= 0 //Inicio de Período de Gozo
Local dRettrb	:= Nil //Retorno ao trabalho


cQrySQB:="SELECT QB_DEPTO AS DEPARTAMENTO,RA_MAT AS MATRICULA,RA_NOME AS NOME,QB_MATRESP AS GESTOR,*  FROM SQB100 QB"
cQrySQB+=" INNER JOIN SRA100 RA ON QB.QB_MATRESP=RA.RA_MAT AND RA.D_E_L_E_T_=''"
cQrySQB+=" WHERE QB.D_E_L_E_T_=''"
//cQrySQB+=" AND QB_DEPTO = '000015'"
//cQrySQB+=" AND QB_MATRESP = '003005'"


If Select("SQBM")>0 
    dbSelectArea("SQBM")
    dbCloseArea("SQBM")
EndIf
TCquery cQrySQB New Alias "SQBM"

While !SQBM->(Eof())
    nCount ++//poe um ponto de parada aqui
    cQrySQB2 := " SELECT C.QB_DEPSUP AS DEPTO_SUPERIOR,B.RA_DEPTO AS COD_DEPARTAMENTO,QB_DESCRIC AS DEPARTAMENTO,C.QB_MATRESP AS COD_GESTOR, "
    cQrySQB2 += " RF_MAT AS MATRICULA ,B.RA_NOME AS NOME,RF_DATABAS AS DATA_BASE,RF_DIASDIR AS DIAS_DIREITO,RF_DATAINI AS FERIAS_01,RF_DATAFIM AS FIM_FERIAS ,RF_PERC13S AS DECIMO,RF_DFEPRO1 AS DIA_FERIAS01,RF_DATINI2 AS FERIAS_02,RF_DFEPRO2 AS DIA_FERIAS02,RF_DATINI3  AS FERIAS_03,RF_DFEPRO3 AS DIA_FERIAS02,RF_PD,RF_STATUS "
    cQrySQB2 += " FROM SRF100 A "
    cQrySQB2 += " JOIN SRA100 B ON  A.RF_MAT=B.RA_MAT AND A.RF_FILIAL= B.RA_FILIAL AND B.D_E_L_E_T_ = '' "
    cQrySQB2 += " JOIN SQB100 C ON  B.RA_DEPTO= C.QB_DEPTO "
    cQrySQB2 += " WHERE A.D_E_L_E_T_='' "
    cQrySQB2 += " AND RA_MSBLQL='2' "
	cQrySQB2 += " AND RA_SITFOLH <>'D' "
    cQrySQB2 += " AND RF_STATUS='1' "
    cQrySQB2 += " AND RF_PD in ('200', '237') "
    cQrySQB2 += " AND QB_DEPTO='" + SQBM->DEPARTAMENTO + "' "
    //cQrySQB2 += " and QB_MATRESP ='" + SQBM->MATRICULA + "' "
	cQrySQB2 += " and QB_MATRESP ='002021' "// fixei registro
    cQrySQB2 += " AND RA_DEMISSA='' "
	cQrySQB2 += " AND RF_DIASDIR<>''"
    cQrySQB2 += "AND RF_DATAINI between '20200201' and '20200229' "

        If Select("SRFT")>0 
        dbSelectArea("SRFT")
        dbCloseArea("SRFT")
    EndIf
    TCquery cQrySQB2 New Alias "SRFT"

	
	If !SRFT->(Eof())
		 
	    _cHtml :="<html>                          "	                                       
		_cHtml +="	<body text='#000000' bgcolor='#FFFFFF' link='#FFFF00' vlink='#FFFF00' alink='#FFFF99'> "
		_cHtml +=" 		<center><b><font face='Arial,Helvetica' color='#FF0000' size=+1>Aviso de Férias - " + SRFT->DEPARTAMENTO + "</font></center>"
		_cHtml +="		<hr size=5 align='center'>"
		_cHtml +=" 		<table border='0' bgcolor='lightgray'>"
		_cHtml +=" 			<tr><td valign='top' bgcolor='lightgray'>"
		_cHtml +=" 					<table border='0' width='100%'>" 
		_cHtml +=" 						<tr bgcolor='#0B0B61'><td width='5%'><font color='#FFFFFF'><b>Matricula</b></font></td>
		_cHtml +=" 											  <td width='25%'><font color='#FFFFFF'><b>Nome</b></font></td>
		_cHtml +=" 											  <td width='25%'><font color='#FFFFFF'><b>Inicio do Período Aquisitivo</b></font></td>
	    _cHtml +=" 											  <td width='25%'><font color='#FFFFFF'><b>Fim do Período Aquisitivo</b></font></td>
	    _cHtml +=" 											  <td width='25%'><font color='#FFFFFF'><b>Inicio de Ferias</b></font></td>
	    _cHtml +=" 											  <td width='25%'><font color='#FFFFFF'><b>Fim de Ferias</b></font></td>
		_cHtml +=" 											  <td width='25%'><font color='#FFFFFF'><b>Retorno de Trabalho</b></font></td>
		_cHtml +=" 														</tr>"
		lZebrado := .F.
	        While !SRFT->(Eof())  

				dIniPA	:= Dtoc(Stod(SRFT->DATA_BASE) ) //Inicio do Período Aquisitivo
				dFimPA	:= Dtoc(Stod(SRFT->FIM_FERIAS)) //Fim do Período Aquisitivo
				dIniFe	:= Stod(SRFT->FERIAS_01) //Inicio de Período de Gozo
				nDiaFer := SRFT->DIA_FERIAS01 // Dias de Férias
				dFimFe	:= dIniFe + nDiaFer //Fim de Período de Gozo
				nDiaDir := str(SRFT->DIAS_DIREITO) //Dias de Direito
				dRettrb	:= dFimFe + 1	//Retorno ao trabalho

				_cHtml +="	  					<tr bgcolor=" + Iif( !lZebrado, 'white', '#D3DFEE' ) + ">
				_cHtml +="											<td> "+ SRFT->MATRICULA +" </td>
	            _cHtml +="											<td> "+ SRFT->NOME +" </td>
	            _cHtml +="											<td> "+ dIniPA +" </td>
	            _cHtml +="											<td> "+ dFimPA +" </td>
				_cHtml +="											<td> "+ cValtoChar(dIniFe) +" </td>
				_cHtml +="											<td> "+ cValtoChar(dFimFe) +" </td>
				_cHtml +="											<td> "+ cValtoChar(dRettrb)  +" </td>
				_cHtml +="											</tr>"
				lZebrado := !lZebrado
			    SRFT->(DBSKIP())
	        End
		_cHtml +=" 					</table>"
		_cHtml +=" 		</table>"
		_cHtml	+=" <p><span style=color: midnightblue;font-family: calibri;'>E-mail automático. Favor não responder a esse e-mail!</span> </p>"
		_cHtml +=" 		<br> </br>"
		_cHtml +=" 	</body>"
		_cHtml +=" </html>"
	
		U_ENVMAIL("workflow@sinaf.com.br","lrosa@sinaf.com.br","Aviso de Ferias",_cHTML,NIL,.T.)
		
    Endif

	
    //Alert("Estou no registro "+cValToChar(nCount)+ "Campo: "+ SQBM->DEPARTAMENTO+" "+SQBM->MATRICULA+ " "+ SQBM->NOME)
    SQBM->(DBSKIP())	
End
Alert("FIM")

Return

//MonthSum(Data,nMes)
//AnoMes(MonthSub(Stod(_cAnoMesIni + "28"),2))
//_cMesAno	:= MonthSub(LastDay(_cMesAno),2) //2 meses
//MonthSum(Data,nMes)//Soma mes(es) a uma Data
//MonthSum(Dtoc(Stod(SRFT->FERIAS_01)),1 )
//DaySum(Data,nDias) // Soma dia(s) a uma Data
//DaySum(MonthSum(Dtoc(Stod(SRFT->FERIAS_01)),1 ),1)
//DaySum(Dtoc(Stod(SRFT->FERIAS_01),SRFT->DIA_FERIAS01) // Soma dia(s) a uma Data

    
